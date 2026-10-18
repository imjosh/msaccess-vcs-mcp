"""Caller recovery requires process exit and all worker lifetimes to settle."""

from dataclasses import replace
import os
from types import SimpleNamespace
from unittest.mock import MagicMock

import pytest

from msaccess_vcs_mcp.access_com import connection as conn
from msaccess_vcs_mcp.access_com import instance_registry as registry
from msaccess_vcs_mcp.access_com import process_qos as qos
from msaccess_vcs_mcp.vba_worker_manager import VBAWorkerManager
from msaccess_vcs_mcp import vba_worker_manager as workers
from msaccess_vcs_mcp.com_recovery import COMRecoveryManager


@pytest.fixture
def recovery(monkeypatch):
    now = [20.0]
    records = [registry.OwnedInstance(123, 'target.accdb', 456)]
    calls = []

    class Process:
        exit_at = 62.0
        unknown = False
        def exited(self):
            return None if self.unknown else now[0] >= self.exit_at
        def force(self):
            calls.append(('force', now[0]))
            self.exit_at = now[0]
            return True
        def close(self):
            calls.append(('handle_closed', now[0]))

    process = Process()

    class Attempt:
        released_at = 50.0
        evidence = conn._CloseEvidence(quit_returned_at=3.0)
        def __init__(self, app):
            calls.append(('marshal', now[0]))
        def _update(self, **fields):
            self.evidence = replace(self.evidence, **fields)
        def start(self):
            pass
        def join(self, timeout):
            settled = now[0] >= self.released_at
            self.evidence = replace(self.evidence, apartment_released=settled, worker_completed=settled)
            return settled

    class Replacement:
        def __init__(self, path):
            calls.append(('replacement', path))
        def __enter__(self):
            return self
        def __exit__(self, *args):
            return False

    monkeypatch.setattr(conn, '_pending_closes', {})
    monkeypatch.setattr(conn, '_recycle_closed', {})
    monkeypatch.setattr(conn, '_QuitAttempt', Attempt)
    monkeypatch.setattr(conn, '_resolve_instance_for_record', lambda record: object())
    monkeypatch.setattr(conn, 'AccessConnection', Replacement)
    monkeypatch.setattr(conn.time, 'monotonic', lambda: now[0])
    monkeypatch.setattr(conn.time, 'sleep', lambda interval: now.__setitem__(0, now[0] + interval))
    monkeypatch.setattr(conn, '_log_instance_event', lambda *a, **kw: None)
    monkeypatch.setattr(qos, 'open_exact_process', lambda pid, stamp: process)
    monkeypatch.setattr(registry, 'owned_records_for_paths', lambda paths: list(records))
    monkeypatch.setattr(registry, 'unregister_owned', lambda *args: None)
    return now, records, calls, process, Attempt


def test_actual_recycle_allows_late_natural_exit_without_force(recovery):
    now, _, calls, _, _ = recovery
    assert conn.recycle_owned_instance('target.accdb') is True
    assert now[0] >= 62
    assert not any(call[0] == 'force' for call in calls)
    assert calls[-1] == ('replacement', 'target.accdb')


def test_actual_recycle_force_clock_starts_at_quit_return(recovery):
    now, _, calls, process, _ = recovery
    process.exit_at = 100
    assert conn.recycle_owned_instance('target.accdb') is True
    forced_at = next(call[1] for call in calls if call[0] == 'force')
    assert 63 <= forced_at < 63.1


def test_live_recovery_worker_blocks_retry_after_registry_pruning(recovery):
    _, records, calls, process, attempt = recovery
    process.exit_at = 0
    attempt.released_at = 100
    key = (os.path.normcase(os.path.abspath('target.accdb')), 123, 456)
    conn._pending_closes[key] = conn._PendingClose(
        'target.accdb', records[0], process, attempt(object()))
    assert conn.recycle_owned_instance('target.accdb') is False
    records.clear()
    assert conn.recycle_owned_instance('target.accdb') is False
    assert not any(call[0] == 'replacement' for call in calls)
    assert len([call for call in calls if call[0] == 'marshal']) == 1
    attempt.released_at = 0
    assert conn.recycle_owned_instance('target.accdb') is True


def test_live_predecessor_blocks_before_replacement_and_retry_can_recover(recovery):
    _, records, calls, process, attempt = recovery
    process.exit_at = 0
    attempt.released_at = 0
    assert conn.recycle_owned_instance('target.accdb', settle_predecessor=lambda: False) is False
    records.clear()
    assert not any(call[0] == 'replacement' for call in calls)
    assert conn.recycle_owned_instance('target.accdb', settle_predecessor=lambda: True) is True


def test_unknown_process_observation_blocks_force_and_replacement(recovery):
    _, _, calls, process, _ = recovery
    process.unknown = True
    assert conn.recycle_owned_instance('target.accdb') is False
    assert not any(call[0] in {'force', 'replacement'} for call in calls)


def test_force_acceptance_without_exit_refuses_replacement(recovery):
    _, _, calls, process, _ = recovery
    process.exit_at = 100
    def accepted():
        calls.append(('force', 63))
        return True
    process.force = accepted
    assert conn.recycle_owned_instance('target.accdb') is False
    assert not any(call[0] == 'replacement' for call in calls)


def test_unavailable_identity_never_resolves_com_or_forces(recovery, monkeypatch):
    _, _, calls, _, _ = recovery
    monkeypatch.setattr(qos, 'open_exact_process', lambda *args: None)
    assert conn.recycle_owned_instance('target.accdb') is False
    assert calls == []


def test_one_blocked_record_refuses_replacement(recovery, monkeypatch):
    _, records, calls, process, attempt = recovery
    records.append(registry.OwnedInstance(124, 'target.accdb', 789))
    process.exit_at = 0
    attempt.released_at = 0
    monkeypatch.setattr(qos, 'open_exact_process', lambda pid, stamp: process if pid == 123 else None)
    assert conn.recycle_owned_instance('target.accdb') is False
    assert not any(call[0] == 'replacement' for call in calls)


def test_replacement_admission_refusal_preserves_close_evidence(recovery, monkeypatch):
    _, _, _, process, attempt = recovery
    process.exit_at = 0
    attempt.released_at = 0
    def refused(path):
        raise RuntimeError('unsupported provider')
    monkeypatch.setattr(conn, 'AccessConnection', refused)
    assert conn.recycle_owned_instance('target.accdb') is False
    closed = next(iter(conn._recycle_closed.values()))
    assert closed[0]['evidence']['process_exited'] is True


@pytest.mark.parametrize('retry_success', [False, True])
def test_actual_manager_recycle_probes_and_retains_failed_retry(recovery, monkeypatch, retry_success):
    _, _, calls, process, attempt = recovery
    process.exit_at = 0
    attempt.released_at = 0
    manager = VBAWorkerManager()
    manager.recovery = MagicMock()
    manager.recovery.should_probe.return_value = True
    probes = iter([{'success': False, 'error': 'blocked'}, {'success': retry_success, 'error': 'retry failed'}])
    monkeypatch.setattr(manager, 'probe', lambda *args: next(probes))
    worker = MagicMock()
    worker.is_alive.return_value = False
    monkeypatch.setattr(VBAWorkerManager, '_active_worker', worker)
    result = manager._probe_if_needed('target.accdb', None)
    assert any(call[0] == 'replacement' for call in calls)
    assert (result is None) is retry_success
    assert manager.recovery.mark_healthy.called is retry_success
    if not retry_success:
        assert manager.recovery.mark_failure.call_count == 2


def test_actual_manager_live_predecessor_refuses_before_admission(recovery, monkeypatch):
    _, _, calls, process, attempt = recovery
    process.exit_at = 0
    attempt.released_at = 0
    manager = VBAWorkerManager()
    manager.recovery = MagicMock()
    manager.recovery.should_probe.return_value = True
    monkeypatch.setattr(manager, 'probe', lambda *args: {'success': False})
    worker = MagicMock()
    worker.is_alive.return_value = True
    monkeypatch.setattr(VBAWorkerManager, '_active_worker', worker)
    assert manager._probe_if_needed('target.accdb', None) is not None
    assert not any(call[0] == 'replacement' for call in calls)


def test_exact_handle_rejects_reused_pid_before_force(monkeypatch):
    import win32process
    kernel = MagicMock()
    kernel.OpenProcess.return_value = 99
    monkeypatch.setattr(qos, '_kernel32', lambda: kernel)
    monkeypatch.setattr(win32process, 'GetProcessTimes', lambda handle: {'CreationTime': 789})
    assert qos.open_exact_process(123, 456) is None
    kernel.CloseHandle.assert_called_once_with(99)
    kernel.TerminateProcess.assert_not_called()


def test_exact_handle_keeps_force_acceptance_separate_from_exit(monkeypatch):
    import win32process
    kernel = MagicMock()
    kernel.OpenProcess.return_value = 99
    kernel.WaitForSingleObject.return_value = 258
    kernel.TerminateProcess.return_value = True
    monkeypatch.setattr(qos, '_kernel32', lambda: kernel)
    monkeypatch.setattr(win32process, 'GetProcessTimes', lambda handle: {'CreationTime': 456})
    process = qos.open_exact_process(123, 456)
    assert process.exited() is False
    assert process.force() is True
    assert process.exited() is False
    kernel.WaitForSingleObject.return_value = 0
    assert process.exited() is True
    process.close()
    kernel.TerminateProcess.assert_called_once_with(99, 1)
    kernel.CloseHandle.assert_called_once_with(99)


@pytest.mark.parametrize('retained', [False, True])
def test_exact_exit_before_proxy_resolution_allows_replacement(recovery, monkeypatch, retained):
    _, records, calls, process, _ = recovery
    process.exit_at = 0
    record = records[0]
    if retained:
        key = (os.path.normcase(os.path.abspath('target.accdb')), 123, 456)
        conn._pending_closes[key] = conn._PendingClose('target.accdb', record, process)
        records.clear()
    resolved = MagicMock(side_effect=AssertionError('Exited host must not require COM'))
    monkeypatch.setattr(conn, '_resolve_instance_for_record', resolved)
    closed = []
    monkeypatch.setattr(conn, '_log_instance_event', lambda event, **fields: closed.append(fields))
    assert conn.recycle_owned_instance('target.accdb') is True
    resolved.assert_not_called()
    assert not conn.owned_recovery_pending('target.accdb')
    assert calls == [('handle_closed', 20.0), ('replacement', 'target.accdb')]
    assert closed[0]['quit'] is False
    assert closed[0]['terminated'] is False
    assert closed[0]['evidence'] == conn.asdict(conn._CloseEvidence(process_exited=True))


@pytest.fixture
def real_probe(monkeypatch):
    monkeypatch.setattr(VBAWorkerManager, '_active_worker', None)
    monkeypatch.setattr(workers.pythoncom, 'CoInitialize', lambda: None)
    monkeypatch.setattr(workers.pythoncom, 'CoUninitialize', lambda: None)
    rot = MagicMock(return_value=SimpleNamespace())
    monkeypatch.setattr(workers.VCSAddinIntegration, '_find_access_in_rot', rot)
    monkeypatch.setattr(workers.VCSAddinIntegration, '__init__', lambda self, *args: None)
    monkeypatch.setattr(workers.VCSAddinIntegration, 'load_addin', lambda self, *args, **kwargs: None)
    monkeypatch.setattr(workers, 'ensure_access_visible', lambda app: None)
    for name in ('log_diagnostic_event', 'log_vba_worker_event', 'log_com_recovery_event'):
        monkeypatch.setattr(workers, name, lambda *args, **kwargs: None)
    manager = VBAWorkerManager()
    manager.recovery = COMRecoveryManager()
    return manager, rot


@pytest.mark.parametrize('failed_before', [False, True])
def test_real_initial_probe_settles_live_recovery_before_rot(recovery, real_probe, failed_before):
    _, records, calls, process, attempt = recovery
    manager, rot = real_probe
    process.exit_at = 0
    attempt.released_at = 100
    key = (os.path.normcase(os.path.abspath('target.accdb')), 123, 456)
    conn._pending_closes[key] = conn._PendingClose(
        'target.accdb', records[0], process, attempt(object()))
    records.clear()
    if failed_before:
        manager.recovery.mark_failure('target.accdb', 'previous close pending', 'access_unresponsive')
    assert manager._probe_if_needed('target.accdb', None) is not None
    assert manager.recovery.get_state('target.accdb').status != 'recovered'
    rot.assert_not_called()
    assert not any(call[0] == 'replacement' for call in calls)
    attempt.released_at = 0
    assert manager._probe_if_needed('target.accdb', None) is None
    assert manager.recovery.get_state('target.accdb').status == 'recovered'
    rot.assert_called_once_with('target.accdb')
    assert len([call for call in calls if call[0] == 'marshal']) == 1


@pytest.mark.parametrize('barrier', ['attempt', 'closed', 'admission'])
@pytest.mark.parametrize('operation', ['probe', 'run_vba'])
def test_real_worker_refuses_rot_during_pending_recovery(recovery, real_probe, barrier, operation):
    _, records, _, process, _ = recovery
    manager, rot = real_probe
    if barrier == 'closed':
        conn._recycle_closed[os.path.normcase(os.path.abspath('target.accdb'))] = [{'pid': 123}]
    else:
        key = (os.path.normcase(os.path.abspath('target.accdb')), 123, 456)
        conn._pending_closes[key] = conn._PendingClose(
            'target.accdb', records[0], process,
            admission_worker=SimpleNamespace(is_alive=lambda: True) if barrier == 'admission' else None)
    result = manager._run_worker(operation=operation, database_path='target.accdb',
                                 addin_path=None, code='Debug.Print 1', timeout_seconds=1)
    assert result['success'] is False
    rot.assert_not_called()


def test_real_initial_probe_waits_for_predecessor_after_retained_closure(recovery, real_probe, monkeypatch):
    _, records, calls, _, _ = recovery
    manager, rot = real_probe
    records.clear()
    conn._recycle_closed[os.path.normcase(os.path.abspath('target.accdb'))] = [{'pid': 123}]
    worker = MagicMock()
    worker.is_alive.return_value = True
    monkeypatch.setattr(VBAWorkerManager, '_active_worker', worker)
    assert manager._probe_if_needed('target.accdb', None) is not None
    rot.assert_not_called()
    assert not any(call[0] == 'replacement' for call in calls)
    worker.is_alive.return_value = False
    assert manager._probe_if_needed('target.accdb', None) is None
    assert manager.recovery.get_state('target.accdb').status == 'recovered'
    assert calls[-1] == ('replacement', 'target.accdb')


def test_real_initial_probe_without_pending_recovery_preserves_responsive_host(recovery, real_probe):
    manager, rot = real_probe
    manager.recovery.mark_failure('target.accdb', 'previous failure', 'access_unresponsive')
    assert manager._probe_if_needed('target.accdb', None) is None
    assert manager.recovery.get_state('target.accdb').status == 'recovered'
    rot.assert_called_once_with('target.accdb')
    assert recovery[2] == []


def test_live_admission_worker_blocks_even_after_exact_exit(recovery):
    _, records, calls, process, _ = recovery
    process.exit_at = 0
    alive = [True]
    key = (os.path.normcase(os.path.abspath('target.accdb')), 123, 456)
    conn._pending_closes[key] = conn._PendingClose(
        'target.accdb', records[0], process,
        admission_worker=SimpleNamespace(is_alive=lambda: alive[0]))
    records.clear()
    assert conn.recycle_owned_instance('target.accdb') is False
    assert calls == []
    alive[0] = False
    assert conn.recycle_owned_instance('target.accdb') is True
    assert not conn.owned_recovery_pending('target.accdb')
