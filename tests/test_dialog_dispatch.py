"""Dispatch regressions through the registered public dialog tool wrappers."""

from __future__ import annotations

import asyncio
import threading
from unittest.mock import AsyncMock, MagicMock

import pytest

from msaccess_vcs_mcp import tools
from msaccess_vcs_mcp.access_gate import reset_access_gate
from msaccess_vcs_mcp.dialog_recovery import reset_interruptions

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
