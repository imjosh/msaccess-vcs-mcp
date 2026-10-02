"""Dispatch regressions through the registered public dialog tool wrappers."""

from __future__ import annotations

import asyncio
import json
import threading
import time
from unittest.mock import AsyncMock, MagicMock

import pytest

from msaccess_vcs_mcp import tools
from msaccess_vcs_mcp.access_gate import get_access_gate, reset_access_gate
from msaccess_vcs_mcp.dialog_recovery import reset_interruptions
from msaccess_vcs_mcp.operation_manager import OperationManager

from .test_dialog_recovery import FakeBackend, _button, _win

DEADLINE = 1.0
FAILSAFE = 5.0
DIALOG_TOOLS = (
    tools.vcs_list_dialogs,
    tools.vcs_dismiss_dialog,
    tools.vcs_recover_dialogs,
    tools.vcs_automation_status,
)
CLICK_TOOLS = (tools.vcs_dismiss_dialog, tools.vcs_recover_dialogs)


class Blocker:
    """Coordinate a fake with the loop, even if broken dispatch blocks that loop."""

    def __init__(self):
        self.entered = threading.Event()
        self.release = threading.Event()
        self.expired = False

    def wait(self):
        self.entered.set()
        # asyncio deadlines cannot fire if a regressed wrapper runs inline.
        # This fallback frees that loop so the test can fail and clean up.
        self.expired = not self.release.wait(FAILSAFE)


class BlockingBackend(FakeBackend):
    def __init__(self, block_at="listing", delivered=True):
        super().__init__([
            _win(hwnd=1, title="Northwind : Database", class_name="OMain"),
            _win(hwnd=2, title="VCS Probe", texts=("Hello",), buttons=(_button(21, "OK"),)),
        ])
        self.blocker = Blocker()
        self.block_at = block_at
        self.delivered = delivered

    def list_windows(self):
        if self.block_at == "listing" and not self.blocker.entered.is_set():
            self.blocker.wait()
        return super().list_windows()

    def click(self, button, *, expected_pid=None, timeout_ms=5000):
        if self.block_at == "click":
            self.blocker.wait()
        if not self.delivered:
            self.clicked.append(button.hwnd)
            return False
        return super().click(button, expected_pid=expected_pid, timeout_ms=timeout_ms)


@pytest.fixture
def public_tools(monkeypatch, tmp_path):
    """Keep configuration and COM external to the public wrapper under test."""
    db = tmp_path / "Northwind.accdb"
    db.write_bytes(b"")
    reset_access_gate()
    reset_interruptions()
    monkeypatch.setattr(tools, "_ensure_env_loaded", AsyncMock())
    monkeypatch.setattr(tools, "load_config", lambda: {})
    monkeypatch.setattr(tools, "get_config", lambda: {})
    monkeypatch.setenv("ACCESS_VCS_ENABLE_LOGGING", "false")
    monkeypatch.setenv("ACCESS_VCS_DISABLE_DIAGNOSTIC_LOG", "true")
    monkeypatch.setenv("ACCESS_VCS_BUSY_WAIT_SEC", "0.05")
    connection = MagicMock()
    connection.return_value.__enter__.return_value.connect.return_value = (object(), object())
    monkeypatch.setattr(tools, "AccessConnection", connection)
    addin = MagicMock()
    addin.call_sync.return_value = "fake option value"
    monkeypatch.setattr(tools, "VCSAddinIntegration", lambda *_args: addin)
    try:
        yield str(db), addin
    finally:
        reset_interruptions()
        reset_access_gate()


def _install_backend(monkeypatch, backend):
    monkeypatch.setattr("msaccess_vcs_mcp.dialog_recovery._default_backend", lambda: backend)


async def _entered(blocker):
    entered = await asyncio.wait_for(asyncio.to_thread(blocker.entered.wait, DEADLINE), DEADLINE)
    assert entered, "The public tool did not reach its fake backend"
    assert not blocker.expired, "The backend blocked the event loop until its fallback deadline"


async def _dialog_call(tool, db, timeout=1.0):
    kwargs = {"pid": 10, "create_time": 1000, "timeout_seconds": timeout}
    if tool is tools.vcs_dismiss_dialog:
        kwargs.update(dialog_id="hwnd:2", button="OK")
    elif tool is tools.vcs_recover_dialogs:
        kwargs["policy"] = "safe"
    return await tool(db, **kwargs)


async def _loop_progress(blocker, task):
    async def probe():
        await asyncio.sleep(0)
        assert not blocker.release.is_set()
        assert not blocker.expired
        assert not task.done(), "The backend must still be blocked while this coroutine progresses"
        return "progress"

    assert await asyncio.wait_for(asyncio.create_task(probe()), DEADLINE) == "progress"


def _assert_finished(tool, backend, result):
    assert result["success"] is True
    assert result["pid"] == 10
    assert result["create_time"] == 1000
    if tool in CLICK_TOOLS:
        assert backend.clicked == [21]
        assert result["dialogs"] == []
        assert [item["dialog_id"] for item in result["closed"]] == ["hwnd:2"]
        if tool is tools.vcs_dismiss_dialog:
            assert result["dismissed"] is True
    else:
        assert backend.clicked == []
        assert [item["dialog_id"] for item in result["dialogs"]] == ["hwnd:2"]
        if tool is tools.vcs_automation_status:
            assert result["ready"] is False
    assert backend.closed == []


@pytest.mark.parametrize("tool", DIALOG_TOOLS, ids=lambda tool: tool.__name__)
@pytest.mark.parametrize("gate_busy", [False, True], ids=["gate_free", "gate_busy"])
def test_blocked_public_dialog_keeps_loop_and_gated_tool_responsive(
    public_tools, monkeypatch, tool, gate_busy,
):
    db, addin = public_tools
    backend = BlockingBackend()
    _install_backend(monkeypatch, backend)
    gate_blocker = Blocker()

    def get_option(*args):
        gate_blocker.wait()
        return "fake option value"

    async def scenario():
        tasks = []
        try:
            if gate_busy:
                addin.call_sync.side_effect = get_option
                tasks.append(asyncio.create_task(tools.vcs_get_option(db, "ShowDebug")))
                await _entered(gate_blocker)
            dialog_task = asyncio.create_task(_dialog_call(tool, db))
            tasks.append(dialog_task)
            await _entered(backend.blocker)
            await _loop_progress(backend.blocker, dialog_task)
            gated = await asyncio.wait_for(tools.vcs_get_option(db, "ShowDebug"), DEADLINE)
            if gate_busy:
                assert gated["success"] is False
                assert gated["error_pattern"] == "server_busy"
                assert gated["recoverable"] is True
                assert gated["busy_with"]["tool"] == "vcs_get_option"
                assert gated["busy_with"]["database"] == db
            else:
                assert gated == {
                    "success": True, "option": "ShowDebug", "value": "fake option value",
                }
            # The gated result must arrive before the dialog backend is released.
            await _loop_progress(backend.blocker, dialog_task)
            backend.blocker.release.set()
            result = await asyncio.wait_for(dialog_task, DEADLINE)
            _assert_finished(tool, backend, result)
            if gate_busy:
                gate_blocker.release.set()
                assert (await asyncio.wait_for(tasks[0], DEADLINE))["success"] is True
                assert not gate_blocker.expired
        finally:
            backend.blocker.release.set()
            gate_blocker.release.set()
            await asyncio.gather(*tasks, return_exceptions=True)

    asyncio.run(scenario())


@pytest.mark.parametrize("tool", DIALOG_TOOLS, ids=lambda tool: tool.__name__)
def test_public_dialog_worker_ceiling_returns_recoverable_timeout(
    public_tools, monkeypatch, tool,
):
    db, _addin = public_tools
    backend = BlockingBackend()
    _install_backend(monkeypatch, backend)
    # Shorten only the extra worker margin; the real wrapper enforces the ceiling.
    monkeypatch.setattr(tools, "EXEMPT_WORKER_MARGIN_SEC", 0.0)

    async def scenario():
        task = asyncio.create_task(_dialog_call(tool, db, timeout=0.2))
        try:
            await _entered(backend.blocker)
            result = await asyncio.wait_for(asyncio.shield(task), DEADLINE)
            assert not backend.blocker.release.is_set()
            assert not backend.blocker.expired
            assert result["success"] is False
            assert result["error_pattern"] == "tool_timeout"
            assert result["recoverable"] is True
            assert tool.__name__ in result["error"]
            assert backend.clicked == []
            assert backend.closed == []
        finally:
            backend.blocker.release.set()
            await asyncio.gather(task, return_exceptions=True)

    # asyncio.run drains released worker threads before the fixture restores patches.
    asyncio.run(scenario())


@pytest.mark.parametrize("tool", CLICK_TOOLS, ids=lambda tool: tool.__name__)
@pytest.mark.parametrize("delivered", [True, False], ids=["delivered", "undelivered"])
def test_blocked_public_click_returns_dismissal_outcome(
    public_tools, monkeypatch, tool, delivered,
):
    db, _addin = public_tools
    backend = BlockingBackend(block_at="click", delivered=delivered)
    _install_backend(monkeypatch, backend)

    async def scenario():
        task = asyncio.create_task(_dialog_call(tool, db))
        try:
            await _entered(backend.blocker)
            await _loop_progress(backend.blocker, task)
            backend.blocker.release.set()
            result = await asyncio.wait_for(task, DEADLINE)
            if delivered:
                _assert_finished(tool, backend, result)
            else:
                assert result["success"] is False
                assert result["error_pattern"] == "dismiss_uncertain"
                assert result["uncertain"] is True
                assert [item["dialog_id"] for item in result["dialogs"]] == ["hwnd:2"]
                assert backend.clicked == [21]
                assert backend.closed == []
        finally:
            backend.blocker.release.set()
            await asyncio.gather(task, return_exceptions=True)

    asyncio.run(scenario())


# How long the gated call stays inside its blocking step before the other
# tools are called. Under FAILSAFE, so the step is still blocked when they answer.
HELD = 2.0
GATED_ASYNC_CALLS = {
    "vcs_export_database": lambda db, src, out: tools.vcs_export_database(db),
    "vcs_list_objects": lambda db, src, out: tools.vcs_list_objects(db),
    "vcs_import_objects": lambda db, src, out: tools.vcs_import_objects(db, src),
    "vcs_run_tests": lambda db, src, out: tools.vcs_run_tests(db),
    "vcs_rebuild_database": lambda db, src, out: tools.vcs_rebuild_database(src, out),
}


@pytest.mark.parametrize("name", GATED_ASYNC_CALLS)
def test_exempt_tools_answer_while_gated_async_tool_blocks_in_com(
    public_tools, monkeypatch, tmp_path, name,
):
    """A gated async tool stuck in its first COM step (connect, or the build host's
    Access startup) must not hold the loop the dialog, status and cancel tools need."""
    db, _addin = public_tools
    src = tmp_path / "src"
    src.mkdir()
    out = str(tmp_path / "Built.accdb")
    _install_backend(monkeypatch, BlockingBackend(block_at=None))
    blocker = Blocker()

    def blocking_com_step(*_args, **_kwargs):
        blocker.wait()
        raise RuntimeError("Access stopped answering")

    connection = MagicMock()
    connection.return_value.__enter__.return_value.connect.side_effect = blocking_com_step
    monkeypatch.setattr(tools, "AccessConnection", connection)
    monkeypatch.setattr(tools, "create_isolated_access_app", blocking_com_step)

    async def scenario():
        gated = asyncio.create_task(GATED_ASYNC_CALLS[name](db, str(src), out))
        try:
            await _entered(blocker)
            await asyncio.sleep(HELD)
            await _loop_progress(blocker, gated)
            identity = {"pid": 10, "create_time": 1000, "timeout_seconds": 1.0}
            listed = await asyncio.wait_for(tools.vcs_list_dialogs(db, **identity), DEADLINE)
            assert [item["dialog_id"] for item in listed["dialogs"]] == ["hwnd:2"]
            status = await asyncio.wait_for(tools.vcs_automation_status(db, **identity), DEADLINE)
            assert status["success"] is True
            assert status["ready"] is False
            cancel = await asyncio.wait_for(tools.vcs_cancel_operation("not-running"), DEADLINE)
            assert cancel["success"] is False
            assert not gated.done(), "The gated call must still be blocked while the others answer"
            blocker.release.set()
            result = await asyncio.wait_for(gated, DEADLINE)
            assert result["success"] is False
            assert "Access stopped answering" in str(result)
            assert not blocker.expired
        finally:
            blocker.release.set()
            await asyncio.gather(gated, return_exceptions=True)

    asyncio.run(scenario())


def test_cancelled_public_async_worker_cleans_up_its_original_operation(
    public_tools, monkeypatch,
):
    """Forward cancellation at the next await, without retiring identity in COM."""
    db, addin = public_tools
    work = Blocker()
    cleanup = Blocker()
    manager = OperationManager()
    monkeypatch.setattr(tools, "_get_operation_manager", lambda: manager)
    monkeypatch.setattr(tools, "get_callback_url", lambda: "http://fake-callback")
    addin.call_sync.return_value = '{"success":true}'
    operation_ids = []

    def launch(callback_info, *_args):
        operation_ids.append(json.loads(callback_info)["operation_id"])
        work.wait()
        return {"async": True}

    addin.call_async.side_effect = launch
    connection = MagicMock()
    connection.return_value.__enter__.return_value.connect.return_value = (object(), object())
    connection.return_value.__exit__.side_effect = lambda *_args: cleanup.wait()
    monkeypatch.setattr(tools, "AccessConnection", connection)
    backend = FakeBackend([
        _win(hwnd=1, title="Northwind : Database", class_name="OMain"),
        _win(hwnd=2, title="Microsoft Visual Basic", texts=("Run-time error '13':",),
             buttons=(_button(21, "End"), _button(22, "Debug"))),
    ])
    _install_backend(monkeypatch, backend)
    identity = {"pid": 10, "create_time": 1000, "timeout_seconds": 1.0}
    gate = get_access_gate()
    gate._slot = slot = CountedSlot()
    runner_cleanup = threading.Event()
    release_runner = threading.Event()
    original_wait = manager.wait_for_completion

    async def child():
        try:
            await asyncio.sleep(5)
        finally:
            runner_cleanup.set()
            while not release_runner.is_set():
                await asyncio.sleep(0.005)

    async def wait_for_completion(*args, **kwargs):
        asyncio.create_task(child())
        return await original_wait(*args, **kwargs)

    monkeypatch.setattr(manager, "wait_for_completion", wait_for_completion)

    async def scenario():
        task = asyncio.create_task(tools.vcs_run_tests(db))
        try:
            await _entered(work)
            owner = gate.current_in_flight()
            operation_id, = operation_ids
            task.cancel()
            with pytest.raises(asyncio.CancelledError):
                await task
            assert gate.current_in_flight() is owner
            assert manager.pending_count() == 1
            requested = await asyncio.wait_for(tools.vcs_cancel_operation(operation_id), DEADLINE)
            assert requested["success"] is True
            assert requested["operation_id"] == operation_id
            assert requested["cancel_requested"] is True
            assert (await tools.vcs_cancel_operation("refused-caller"))["success"] is False
            # An interruption arriving AFTER caller cancellation still names the worker.
            dismissed = await asyncio.wait_for(
                tools.vcs_dismiss_dialog(db, "hwnd:2", button="End", **identity), DEADLINE,
            )
            assert dismissed["last_interruption"]["busy_with"]["tool"] == "vcs_run_tests"
            work.release.set()
            await _entered(cleanup)
            assert manager.pending_count() == 0  # Cancellation reached the apartment's await.
            assert gate.current_in_flight() is owner
            assert slot.releases == 0
            busy = await asyncio.wait_for(tools.vcs_get_option(db, "ShowDebug"), DEADLINE)
            assert busy["error_pattern"] == "server_busy"
            cleanup.release.set()
            async def until(event):
                while not event.is_set():
                    await asyncio.sleep(0.005)
            await asyncio.wait_for(until(runner_cleanup), DEADLINE)
            # Runner is cleaning up a child after the tool body has unwound.
            # A new dialog action must still belong to the original gate owner.
            assert gate.current_in_flight() is owner
            assert slot.releases == 0
            backend.windows.append(_win(
                hwnd=2, title="Microsoft Visual Basic", texts=("Run-time error '13':",),
                buttons=(_button(21, "End"), _button(22, "Debug")),
            ))
            dismissed = await asyncio.wait_for(
                tools.vcs_dismiss_dialog(db, "hwnd:2", button="End", **identity), DEADLINE,
            )
            assert dismissed["last_interruption"]["busy_with"]["tool"] == "vcs_run_tests"
            release_runner.set()
            async def released():
                while gate.current_in_flight() is not None:
                    await asyncio.sleep(0.005)
            await asyncio.wait_for(released(), DEADLINE)
            assert slot.releases == 1
            report = await tools.vcs_automation_status(db, **identity)
            assert report["ready"] is True
            assert report["last_interruption"] is None
            assert (await tools.vcs_cancel_operation(operation_id))["success"] is False
        finally:
            work.release.set()
            cleanup.release.set()
            release_runner.set()
            await asyncio.gather(task, return_exceptions=True)
            gate._executor.shutdown(wait=True)
            assert not work.expired
            assert not cleanup.expired

    asyncio.run(scenario())


class CountedSlot:
    """Observe releases while retaining the real nonblocking lock behavior."""

    def __init__(self):
        self.lock = threading.Lock()
        self.releases = 0

    def acquire(self, *, blocking):
        return self.lock.acquire(blocking=blocking)

    def release(self):
        self.releases += 1
        self.lock.release()


@pytest.mark.parametrize("tool", [tools.vcs_get_option, tools.vcs_run_tests], ids=["sync", "async"])
@pytest.mark.parametrize("stop", ["cancel", "timeout"])
@pytest.mark.parametrize("fail", [False, True], ids=["returns", "fails"])
def test_cancelled_public_caller_retains_worker_gate_and_interruptions(
    public_tools, monkeypatch, tool, stop, fail,
):
    """Caller completion is independent of blocked COM and connection cleanup."""
    db, addin = public_tools
    work = Blocker()
    cleanup = Blocker()
    gate = get_access_gate()
    gate._slot = slot = CountedSlot()
    backend = FakeBackend([
        _win(hwnd=1, title="Northwind : Database", class_name="OMain"),
        _win(hwnd=2, title="Microsoft Visual Basic", texts=("Run-time error '13':",),
             buttons=(_button(21, "End"), _button(22, "Debug"))),
    ])
    _install_backend(monkeypatch, backend)
    logs = []
    monkeypatch.setattr("msaccess_vcs_mcp.usage_logging._initialize_logging", lambda: True)
    monkeypatch.setattr("msaccess_vcs_mcp.usage_logging.log_tool_call", lambda **entry: logs.append(entry))
    monkeypatch.setattr(tools, "get_callback_url", lambda: None)
    monkeypatch.setattr(tools, "_get_operation_manager", lambda: None)

    def connect():
        work.wait()
        if fail:
            raise RuntimeError("fake COM failure")
        return object(), object()

    connection = MagicMock()
    conn = connection.return_value.__enter__.return_value
    conn.connect.side_effect = connect
    connection.return_value.__exit__.side_effect = lambda *_args: cleanup.wait()
    monkeypatch.setattr(tools, "AccessConnection", connection)
    addin.call_sync.side_effect = lambda name, *_args: (
        '{"success":true}' if name == "SetOption" else
        '{"allPassed":true,"summary":{"passed":1}}' if name == "RunFilteredTests" else
        "fake option value"
    )
    identity = {"pid": 10, "create_time": 1000, "timeout_seconds": 1.0}

    async def status():
        return await asyncio.wait_for(tools.vcs_automation_status(db, **identity), DEADLINE)

    async def assert_held(owner):
        started = time.monotonic()
        busy = await asyncio.wait_for(tools.vcs_get_option(db, "ShowDebug"), 0.5)
        assert busy["error_pattern"] == "server_busy"
        assert busy["busy_with"]["tool"] == tool.__name__
        assert busy["busy_with"]["database"] == db
        assert 0.04 <= time.monotonic() - started < 0.5
        assert gate.current_in_flight() is owner
        assert slot.releases == 0
        report = await status()
        assert report["gate_busy"] is True
        assert report["ready"] is False
        assert report["operation"]["tool"] == tool.__name__
        assert report["last_interruption"]["busy_with"]["tool"] == tool.__name__

    async def scenario():
        call = tool(db, "ShowDebug") if tool is tools.vcs_get_option else tool(db)
        task = asyncio.create_task(call)
        try:
            await _entered(work)
            owner = gate.current_in_flight()
            assert owner is not None
            # Reserve through the public dialog wrapper against the original call.
            dismissed = await asyncio.wait_for(
                tools.vcs_dismiss_dialog(db, "hwnd:2", button="End", **identity), DEADLINE,
            )
            assert dismissed["execution_interrupted"] is True
            if stop == "cancel":
                task.cancel()
                with pytest.raises(asyncio.CancelledError):
                    await task
            else:
                with pytest.raises(asyncio.TimeoutError):
                    await asyncio.wait_for(task, 0.01)
            await assert_held(owner)
            listed = await asyncio.wait_for(tools.vcs_list_dialogs(db, **identity), DEADLINE)
            assert listed["success"] is True
            recovered = await asyncio.wait_for(
                tools.vcs_recover_dialogs(db, policy="report", **identity), DEADLINE,
            )
            assert recovered["success"] is True
            cancel = await asyncio.wait_for(tools.vcs_cancel_operation("not-running"), DEADLINE)
            assert cancel["success"] is False
            # A second cancellation of the finished caller cannot release the worker.
            task.cancel()
            work.release.set()
            await _entered(cleanup)
            await assert_held(owner)
            cleanup.release.set()
            async def released():
                while gate.current_in_flight() is not None:
                    await asyncio.sleep(0.005)
            await asyncio.wait_for(released(), DEADLINE)
            assert slot.releases == 1
            report = await status()
            assert report["gate_busy"] is False
            assert report["ready"] is True
            assert report["last_interruption"] is None
            original_logs = [entry for entry in logs if entry["tool_name"] == tool.__name__]
            assert len(original_logs) == 1
            if tool is tools.vcs_get_option:
                assert original_logs[0]["result"]["execution_interrupted"] is True
            # Use a fresh connection so later work has no fake blockers.
            later_connection = MagicMock()
            later_connection.return_value.__enter__.return_value.connect.return_value = (object(), object())
            monkeypatch.setattr(tools, "AccessConnection", later_connection)
            later = await asyncio.wait_for(tools.vcs_get_option(db, "ShowDebug"), DEADLINE)
            assert later["success"] is True, later
            assert "execution_interrupted" not in later
            assert slot.releases == 2
            assert gate.current_in_flight() is None
        finally:
            work.release.set()
            cleanup.release.set()
            await asyncio.gather(task, return_exceptions=True)
            # Drain the actual apartment work before restoring monkeypatches.
            gate._executor.shutdown(wait=True)
            assert not work.expired
            assert not cleanup.expired

    asyncio.run(scenario())
