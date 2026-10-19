import asyncio
import json
import subprocess
import sys
import threading
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest
from tests.test_export_build_fallbacks import FakeAccess, _addin, _build_env, _unwrap
from msaccess_vcs_mcp import tools, build_owner
from msaccess_vcs_mcp.access_gate import AccessGate, _retention
from msaccess_vcs_mcp.operation_manager import OperationManager
from msaccess_vcs_mcp.access_com.instance_registry import OwnedInstance


@pytest.fixture
def fake_host(monkeypatch, tmp_path):
    process = MagicMock()
    gone = threading.Event()
    process.exited.side_effect = gone.is_set
    process.force.side_effect = lambda: gone.set() or True
    monkeypatch.setenv('ACCESS_VCS_OWNED_INSTANCES_PATH', str(tmp_path / 'owned.json'))
    monkeypatch.setattr(build_owner, 'pid_from_access_app', lambda app: 771)
    monkeypatch.setattr(build_owner.registry, 'process_create_time', lambda pid: 4321)
    monkeypatch.setattr(build_owner, 'open_exact_process', lambda pid, stamp: process)
    monkeypatch.setenv('ACCESS_VCS_BUSY_WAIT_SEC', '0.02')
    return process, gone


@pytest.mark.parametrize('start', [RuntimeError('reply lost'), {'unknown': True}, {'sync': True, 'result': None}])
def test_uncertain_build_returns_while_host_and_gate_remain_owned(tmp_path, fake_host, start):
    manager = OperationManager()
    app = MagicMock()
    threads = []
    app.CloseCurrentDatabase.side_effect = lambda: threads.append(threading.get_ident())
    app.Quit.side_effect = lambda: (threads.append(threading.get_ident()), fake_host[1].set())
    fake = FakeAccess(async_start=json.dumps(start) if isinstance(start, dict) else start)
    addin = _addin(tmp_path, fake)
    gate = AccessGate()
    with _build_env(tmp_path, addin, manager, True, app) as src:
        async def invoke():
            return await gate.run_exclusive('build', 'output', _unwrap(tools.vcs_rebuild_database), True, str(src), str(tmp_path / 'out.accdb'))
        result = asyncio.run(invoke())
        try:
            assert result['completion_unconfirmed'] is True
            assert fake.dispatched == [('APIAsync', 'BuildAs')]
            app.CloseCurrentDatabase.assert_not_called()
            app.Quit.assert_not_called()
            assert gate.current_in_flight() is not None
            busy = asyncio.run(gate.run_exclusive('next', None, lambda: None, False))
            assert busy['error_pattern'] == 'server_busy'
            operation_id = result['operation_id']
            assert manager.request_cancel(operation_id)
            assert manager.route_callback(operation_id, {'type': 'complete'})
        finally:
            fake_host[1].set()
            gate._executor.shutdown(wait=True)
        assert gate.current_in_flight() is None
        assert manager.get_operation(result['operation_id']) is None
        assert len(set(threads)) <= 1


def test_retention_publishes_wrapped_result_after_runner_and_keeps_slot(monkeypatch):
    from msaccess_vcs_mcp.access_gate import _retain_build_host
    gate = AccessGate()
    entered = threading.Event()
    release = threading.Event()
    seen = []
    def retain():
        with pytest.raises(RuntimeError):
            asyncio.get_running_loop()
        seen.append(threading.get_ident())
        entered.set()
        release.wait(2)
    async def body():
        seen.append(threading.get_ident())
        _retain_build_host(retain)
        return {'success': False, 'logged': True}
    try:
        result = asyncio.run(gate.run_exclusive('build', None, body, True))
        assert result == {'success': False, 'logged': True}
        assert entered.wait(1)
        assert gate.current_in_flight() is not None
    finally:
        release.set()
        gate._executor.shutdown(wait=True)
    assert len(set(seen)) == 1


def test_owner_terminal_survives_detach_duplicate_and_closed_loop():
    manager = OperationManager()
    async def register():
        operation_id, _ = manager.register_operation(database_path='out.accdb')
        operation = manager.attach_build_owner(operation_id, {'state': 'pending'})
        return operation_id, operation
    operation_id, operation = asyncio.run(register())
    manager.detach_waiter(operation_id)
    assert manager.route_callback(operation_id, {'type': 'complete'})
    assert manager.route_callback(operation_id, {'type': 'error'})
    assert operation.terminal_callback == {'type': 'complete'}
    assert operation.settlement.is_set()
    assert manager.is_database_busy('out.accdb')
    manager.update_build_cleanup(operation_id, {'state': 'complete'}, release=True)
    assert not manager.is_database_busy('out.accdb')


@pytest.mark.parametrize("reply", [(None, None), RuntimeError("sync reply lost")])
def test_callback_free_build_retains_original_host_until_exact_exit(tmp_path, fake_host, reply):
    manager = OperationManager()
    app = MagicMock()
    fake = FakeAccess(api=reply)
    addin = _addin(tmp_path, fake)
    gate = AccessGate()
    with _build_env(tmp_path, addin, manager, False, app) as src:
        result = asyncio.run(gate.run_exclusive('build', None, _unwrap(tools.vcs_rebuild_database), True, str(src), str(tmp_path / 'out.accdb')))
        try:
            assert result['completion_unconfirmed']
            assert result['cleanup']['state'] == 'pending'
            assert manager.get_operation(result['operation_id']) is not None
            assert fake.dispatched == [('API', 'BuildAs')]
            app.Quit.assert_not_called()
        finally:
            fake_host[1].set()
            gate._executor.shutdown(wait=True)
        app.Quit.assert_not_called()
        assert manager.pending_count() == 0


def test_terminal_build_cleanup_failure_does_not_change_success(tmp_path, fake_host, monkeypatch):
    manager = OperationManager()
    app = MagicMock()
    app.Quit.side_effect = fake_host[1].set
    def complete(info):
        manager.route_callback(json.loads(info)['operation_id'], {'type': 'complete', 'output_path': 'built.accdb'})
    fake = FakeAccess(on_async=complete)
    addin = _addin(tmp_path, fake)
    gate = AccessGate()
    monkeypatch.setattr(build_owner.shutil, 'rmtree', lambda *args: (_ for _ in ()).throw(OSError('locked')))
    monkeypatch.setattr(build_owner.time, 'sleep', lambda delay: None)
    with _build_env(tmp_path, addin, manager, True, app) as src:
        result = asyncio.run(gate.run_exclusive('build', None, _unwrap(tools.vcs_rebuild_database), True, str(src), str(tmp_path / 'out.accdb')))
        gate._executor.shutdown(wait=True)
    assert result['success'] is True
    assert result['cleanup']['state'] == 'cleanup_failed'
    assert manager.get_operation(result['operation_id']) is not None


def test_cancelled_caller_keeps_host_until_terminal(tmp_path, fake_host):
    manager = OperationManager()
    app = MagicMock()
    app.Quit.side_effect = fake_host[1].set
    fake = FakeAccess()
    addin = _addin(tmp_path, fake)
    gate = AccessGate()
    with _build_env(tmp_path, addin, manager, True, app) as src:
        async def run():
            task = asyncio.create_task(gate.run_exclusive('build', None, _unwrap(tools.vcs_rebuild_database), True, str(src), str(tmp_path / 'out.accdb')))
            for _ in range(100):
                if fake.dispatched:
                    break
                await asyncio.sleep(0.005)
            assert fake.dispatched
            task.cancel()
            with pytest.raises(asyncio.CancelledError):
                await task
        asyncio.run(run())
        try:
            app.Quit.assert_not_called()
            assert gate.current_in_flight() is not None
            operation_id = next(iter(manager._operations))
            assert manager.request_cancel(operation_id)
            assert manager.route_callback(operation_id, {'type': 'cancelled'})
        finally:
            fake_host[1].set()
            gate._executor.shutdown(wait=True)
        assert manager.pending_count() == 0


def test_recycler_does_not_close_active_build_owner(monkeypatch):
    from msaccess_vcs_mcp.access_com import connection
    record = OwnedInstance(771, 'out.accdb', 4321, build_operation_id='build')
    resolve = MagicMock()
    monkeypatch.setattr(connection, '_resolve_instance_for_record', resolve)
    assert connection._close_owned_record(record) is None
    resolve.assert_not_called()


def test_terminal_callback_ack_does_not_wait_for_cleanup(tmp_path, fake_host):
    manager = OperationManager()
    app = MagicMock()
    closing = threading.Event()
    release = threading.Event()
    def close():
        closing.set()
        assert release.wait(2)
    app.CloseCurrentDatabase.side_effect = close
    app.Quit.side_effect = fake_host[1].set
    fake = FakeAccess(async_start=RuntimeError('lost reply'))
    gate = AccessGate()
    with _build_env(tmp_path, _addin(tmp_path, fake), manager, True, app) as src:
        result = asyncio.run(gate.run_exclusive('build', None, _unwrap(tools.vcs_rebuild_database), True, str(src), str(tmp_path / 'out.accdb')))
        try:
            assert manager.route_callback(result['operation_id'], {'type': 'complete'})
            assert closing.wait(1)
            assert gate.current_in_flight() is not None
            assert manager.get_operation(result['operation_id']) is not None
            assert manager.route_callback(result['operation_id'], {'type': 'complete'})
        finally:
            release.set()
            gate._executor.shutdown(wait=True)


def test_settled_hung_close_is_bounded_using_pinned_process(tmp_path, fake_host, monkeypatch):
    manager = OperationManager()
    app = MagicMock()
    threads = []
    def close():
        threads.append(threading.get_ident())
        assert fake_host[1].wait(2)
    app.CloseCurrentDatabase.side_effect = close
    app.Quit.side_effect = lambda: threads.append(threading.get_ident())
    monkeypatch.setattr(build_owner, 'get_close_timeout', lambda: 0.02)
    fake = FakeAccess(on_async=lambda info: manager.route_callback(json.loads(info)['operation_id'], {'type': 'complete'}))
    gate = AccessGate()
    with _build_env(tmp_path, _addin(tmp_path, fake), manager, True, app) as src:
        result = asyncio.run(gate.run_exclusive('build', None, _unwrap(tools.vcs_rebuild_database), True, str(src), str(tmp_path / 'out.accdb')))
        gate._executor.shutdown(wait=True)
    assert result['success'] is True
    assert result['cleanup']['state'] == 'complete'
    fake_host[0].force.assert_called_once()
    assert len(set(threads)) == 1


def test_unknown_exact_identity_refuses_before_dispatch(tmp_path, fake_host, monkeypatch):
    manager = OperationManager()
    app = MagicMock()
    fake = FakeAccess()
    gate = AccessGate()
    monkeypatch.setattr(build_owner, 'open_exact_process', lambda pid, stamp: None)
    with _build_env(tmp_path, _addin(tmp_path, fake), manager, True, app) as src:
        result = asyncio.run(gate.run_exclusive('build', None, _unwrap(tools.vcs_rebuild_database), True, str(src), str(tmp_path / 'out.accdb')))
        gate._executor.shutdown(wait=True)
    assert result['success'] is False
    assert result['cleanup']['state'] == 'recovery_required'
    assert fake.dispatched == []
    # This host never opened the target, so the reservation is released.
    app.Quit.assert_called_once()
    assert manager.get_operation(result['operation_id']) is None
    manager.register_operation(database_path=str(tmp_path / 'out.accdb'))


def test_public_wrappers_complete_before_retained_reply(tmp_path, fake_host, monkeypatch):
    from unittest.mock import AsyncMock
    manager = OperationManager()
    app = MagicMock()
    fake = FakeAccess(async_start=RuntimeError('lost reply'))
    gate = AccessGate()
    monkeypatch.setattr(tools, 'get_access_gate', lambda: gate)
    monkeypatch.setattr(tools, '_ensure_env_loaded', AsyncMock())
    monkeypatch.setattr(tools, 'load_config', lambda: {})
    completed = []
    def normalize(call_id, result):
        if result is not None:
            completed.append(call_id)
            return {**result, 'interruption_adjusted': True}
    monkeypatch.setattr(tools, 'finish_gated_call', normalize)
    monkeypatch.setenv('ACCESS_VCS_ENABLE_LOGGING', 'false')
    with _build_env(tmp_path, _addin(tmp_path, fake), manager, True, app) as src:
        result = asyncio.run(tools.vcs_rebuild_database(str(src), str(tmp_path / 'out.accdb')))
        try:
            assert result['interruption_adjusted'] is True
            assert completed
            assert result['cleanup']['state'] == 'pending'
            app.Quit.assert_not_called()
        finally:
            fake_host[1].set()
            gate._executor.shutdown(wait=True)


def _terminal_build(tmp_path, manager, app):
    fake = FakeAccess(on_async=lambda info: manager.route_callback(json.loads(info)['operation_id'], {'type': 'complete'}))
    gate = AccessGate()
    with _build_env(tmp_path, _addin(tmp_path, fake), manager, True, app) as src:
        result = asyncio.run(gate.run_exclusive('build', None, _unwrap(tools.vcs_rebuild_database), True, str(src), str(tmp_path / 'out.accdb')))
        gate._executor.shutdown(wait=True)
    return result


def _wait_released(manager, operation_id):
    pause = threading.Event()
    for _ in range(100):
        if manager.get_operation(operation_id) is None:
            return True
        pause.wait(0.05)
    return False


@pytest.mark.parametrize('damage', ['{not json', '{"instances": 5}', '[{"pid": "x"}]', 'save'])
def test_cleanup_keeps_claim_until_persisted_removal_is_confirmed(tmp_path, fake_host, monkeypatch, damage):
    manager = OperationManager()
    app = MagicMock()
    app.Quit.side_effect = fake_host[1].set
    location = tmp_path / 'owned.json'
    entry = {}
    removing = threading.Event()
    real_rmtree = build_owner.shutil.rmtree
    def remove_then_damage(folder):
        real_rmtree(folder)
        removing.set()
        entry['bytes'] = location.read_bytes()
        if damage != 'save':
            location.write_text(damage, encoding='utf-8')
    monkeypatch.setattr(build_owner.shutil, 'rmtree', remove_then_damage)
    writable = threading.Event()
    real_save = build_owner.registry._save_strict
    def save(records, target):
        if damage == 'save' and removing.is_set() and not writable.is_set():
            raise OSError('disk full')
        real_save(records, target)
    monkeypatch.setattr(build_owner.registry, '_save_strict', save)
    result = _terminal_build(tmp_path, manager, app)
    operation_id = result['operation_id']
    assert result['success'] is True
    assert result['cleanup']['state'] == 'cleanup_failed'
    assert manager.get_operation(operation_id) is not None
    assert location.read_bytes() == (entry['bytes'] if damage == 'save' else damage.encode())
    if damage == 'save':
        writable.set()
    else:
        location.write_bytes(entry['bytes'])
    assert _wait_released(manager, operation_id)
    assert build_owner.registry._load_strict(location) == []


def test_stopped_observer_keeps_claim_and_reported_state(tmp_path, fake_host, monkeypatch):
    manager = OperationManager()
    app = MagicMock()
    app.Quit.side_effect = fake_host[1].set
    monkeypatch.setattr(build_owner.shutil, 'rmtree', lambda folder: (_ for _ in ()).throw(OSError('locked')))
    result = _terminal_build(tmp_path, manager, app)
    operation_id = result['operation_id']
    assert result['cleanup']['state'] == 'cleanup_failed'
    assert build_owner.stop_observers(5.0) == []
    threading.Event().wait(0.3)
    operation = manager.get_operation(operation_id)
    assert operation is not None
    assert [item.pid for item in build_owner.registry._load_strict(tmp_path / 'owned.json')] == [771]


def test_observer_removes_claim_from_registry_pinned_at_start(tmp_path, fake_host, monkeypatch):
    manager = OperationManager()
    app = MagicMock()
    app.Quit.side_effect = fake_host[1].set
    unlocked = threading.Event()
    real_rmtree = build_owner.shutil.rmtree
    def remove(folder):
        if not unlocked.is_set():
            raise OSError('locked')
        real_rmtree(folder)
    monkeypatch.setattr(build_owner.shutil, 'rmtree', remove)
    result = _terminal_build(tmp_path, manager, app)
    assert result['cleanup']['state'] == 'cleanup_failed'
    monkeypatch.setenv('ACCESS_VCS_OWNED_INSTANCES_PATH', str(tmp_path / 'other.json'))
    unlocked.set()
    assert _wait_released(manager, result['operation_id'])
    assert build_owner.registry._load_strict(tmp_path / 'owned.json') == []
    assert not (tmp_path / 'other.json').exists()


def _owner(manager):
    operation_id, _ = manager.register_operation(database_path='out.accdb')
    token = _retention.set([])
    try:
        return build_owner._BuildOwner(manager, operation_id)
    finally:
        _retention.reset(token)


def test_interpreter_exit_is_not_held_by_a_pending_build(tmp_path):
    marker = tmp_path / 'marker.txt'
    location = tmp_path / 'owned.json'
    child = subprocess.run(
        [sys.executable, str(Path(__file__).with_name('build_exit_child.py')), str(marker), str(location)],
        cwd=Path(__file__).parents[1], capture_output=True, text=True, timeout=30,
    )
    assert child.returncode == 0, child.stdout + child.stderr
    assert marker.read_text(encoding='utf-8').splitlines() == ['returned pending', 'atexit-ran']
    # The next server still sees the host as owned and not cleaned up.
    assert [item.pid for item in build_owner.registry._load_strict(location)] == [771]


@pytest.mark.parametrize('damage', ['{not json', '{"instances": 5}', 'transient'])
def test_capture_refuses_untrusted_registry_and_keeps_other_claims(tmp_path, fake_host, monkeypatch, damage):
    location = tmp_path / 'owned.json'
    if damage == 'transient':
        others = [OwnedInstance(100, 'user.accdb', 1), OwnedInstance(200, 'other.accdb', 2, build_operation_id='other-build')]
        build_owner.registry._save_strict(others, location)
        real_read = Path.read_text
        failed = []
        def read_text(self, *args, **kwargs):
            if self == location and not failed:
                failed.append(self)
                raise PermissionError(13, 'sharing violation')
            return real_read(self, *args, **kwargs)
        monkeypatch.setattr(Path, 'read_text', read_text)
    else:
        location.write_text(damage, encoding='utf-8')
    before = location.read_bytes()
    manager = OperationManager()
    owner = _owner(manager)
    with pytest.raises(RuntimeError, match='could not be persisted'):
        owner.capture(MagicMock(), 'out.accdb')
    fake_host[1].set()
    owner.retire()
    assert manager.get_operation(owner.operation_id) is None
    assert location.read_bytes() == before


_HOLD_LOCK = (
    "import msvcrt, sys\n"
    "handle = open(sys.argv[1], 'a+b')\n"
    "msvcrt.locking(handle.fileno(), msvcrt.LK_NBLCK, 1)\n"
    "print('locked', flush=True)\n"
    "sys.stdin.readline()\n"
)


def test_strict_registry_writes_wait_for_another_process(tmp_path, monkeypatch):
    location = tmp_path / 'owned.json'
    mine = OwnedInstance(771, 'out.accdb', 4321, build_operation_id='mine')
    build_owner.registry._save_strict([mine], location)
    monkeypatch.setattr(build_owner.registry, 'LOCK_WAIT_SEC', 0.3)
    holder = subprocess.Popen(
        [sys.executable, '-c', _HOLD_LOCK, str(location) + '.lock'],
        stdin=subprocess.PIPE, stdout=subprocess.PIPE, text=True,
    )
    try:
        assert holder.stdout.readline().strip() == 'locked'
        with pytest.raises(build_owner.registry.RegistryStateUnknown, match='lock'):
            build_owner.registry.remove_build_owned(771, 4321, location)
        assert build_owner.registry._load_strict(location) == [mine]
    finally:
        holder.communicate('\n', timeout=10)
    build_owner.registry.remove_build_owned(771, 4321, location)
    assert build_owner.registry._load_strict(location) == []


_WRITER = (
    "import sys\n"
    "from pathlib import Path\n"
    "from msaccess_vcs_mcp.access_com import instance_registry as registry\n"
    "location, first = Path(sys.argv[1]), int(sys.argv[2])\n"
    "print('ready', flush=True)\n"
    "sys.stdin.readline()\n"
    "for pid in range(first, first + 40):\n"
    "    registry.register_build_owned(pid, 'out.accdb', 1, f'build-{pid}', location)\n"
)


def test_concurrent_processes_keep_every_build_claim(tmp_path):
    location = tmp_path / 'owned.json'
    writers = [
        subprocess.Popen([sys.executable, '-c', _WRITER, str(location), str(first)],
                         stdin=subprocess.PIPE, stdout=subprocess.PIPE, text=True)
        for first in (1000, 2000)
    ]
    try:
        for writer in writers:
            assert writer.stdout.readline().strip() == 'ready'
        for writer in writers:
            writer.stdin.write('\n')
            writer.stdin.flush()
        for writer in writers:
            assert writer.wait(30) == 0
    finally:
        for writer in writers:
            if writer.poll() is None:
                writer.kill()
                writer.wait(5)
    pids = sorted(item.pid for item in build_owner.registry._load_strict(location))
    assert pids == [*range(1000, 1040), *range(2000, 2040)]


def test_result_returned_after_finish_is_not_rewritten(tmp_path, fake_host, monkeypatch):
    manager = OperationManager()
    app = MagicMock()
    app.Quit.side_effect = fake_host[1].set
    gate = AccessGate()
    def broken(*args, **kwargs):
        raise RuntimeError('formatting failed')
    monkeypatch.setattr(tools, '_build_start_outcome', broken)
    fake = FakeAccess(async_start=RuntimeError('reply lost'))
    with _build_env(tmp_path, _addin(tmp_path, fake), manager, True, app) as src:
        result = asyncio.run(gate.run_exclusive('build', None, _unwrap(tools.vcs_rebuild_database), True, str(src), str(tmp_path / 'out.accdb')))
        try:
            assert result['cleanup']['state'] == 'pending'
            assert manager.route_callback(result['operation_id'], {'type': 'complete'})
            assert _wait_released(manager, result['operation_id'])
        finally:
            fake_host[1].set()
            gate._executor.shutdown(wait=True)
    assert result['cleanup']['state'] == 'pending'


@pytest.mark.parametrize('damage', ['{not json', '{"instances": 5}', 'transient'])
def test_ordinary_registration_never_overwrites_an_unread_registry(tmp_path, monkeypatch, damage):
    location = tmp_path / 'owned.json'
    monkeypatch.setenv('ACCESS_VCS_OWNED_INSTANCES_PATH', str(location))
    monkeypatch.setattr(build_owner.registry, 'list_access_pids_or_none', lambda: None)
    if damage == 'transient':
        others = [OwnedInstance(100, 'user.accdb', 1), OwnedInstance(200, 'out.accdb', 2, build_operation_id='retained')]
        build_owner.registry._save_strict(others, location)
        real_read = Path.read_text
        failed = []
        def read_text(self, *args, **kwargs):
            if self == location and not failed:
                failed.append(self)
                raise PermissionError(13, 'sharing violation')
            return real_read(self, *args, **kwargs)
        monkeypatch.setattr(Path, 'read_text', read_text)
    else:
        location.write_text(damage, encoding='utf-8')
    before = location.read_bytes()
    assert build_owner.registry.register_owned(555, 'user.accdb', create_time=5) is None
    assert location.read_bytes() == before


def test_refused_capture_takes_back_only_its_own_claim(tmp_path, fake_host, monkeypatch):
    location = tmp_path / 'owned.json'
    others = [OwnedInstance(100, 'user.accdb', 1), OwnedInstance(200, 'other.accdb', 2, build_operation_id='other-build')]
    build_owner.registry._save_strict(others, location)
    real_load = build_owner.registry._load_strict
    loads = []
    def load(path):
        loads.append(path)
        if len(loads) == 2:
            raise build_owner.registry.RegistryStateUnknown('confirmation read failed')
        return real_load(path)
    monkeypatch.setattr(build_owner.registry, '_load_strict', load)
    manager = OperationManager()
    owner = _owner(manager)
    with pytest.raises(RuntimeError, match='could not be persisted'):
        owner.capture(MagicMock(), 'out.accdb')
    assert real_load(location) == others
    fake_host[1].set()
    owner.retire()
    assert manager.get_operation(owner.operation_id) is None
