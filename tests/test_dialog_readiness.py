"""``vcs_automation_status`` readiness, at the tool boundary.

``ready: true`` requires every condition to be positively confirmed. Tests call
the tool and assert on the returned dictionary.
"""

from __future__ import annotations

import asyncio
import threading
from types import SimpleNamespace
from unittest.mock import patch

import pytest

from msaccess_vcs_mcp import tools
from msaccess_vcs_mcp.access_gate import get_access_gate, reset_access_gate
from msaccess_vcs_mcp.dialog_recovery import ProcessIdentity, WindowInfo, reset_interruptions

from .test_dialog_identity import ScriptedBackend, _main, _msgbox

DB = r"C:\data\Northwind.accdb"
OTHER_DB = r"C:\data\Other.accdb"


class ProbeBackend(ScriptedBackend):
    def __init__(self, windows, identities=None, responsive=True):
        super().__init__(windows, identities)
        self._responsive = responsive

    def responsive(self, hwnd, timeout_ms):
        return self._responsive


@pytest.fixture(autouse=True)
def _clean():
    reset_interruptions()
    reset_access_gate()
    with (
        patch("msaccess_vcs_mcp.tools.validate_database_path", return_value=DB),
        patch("msaccess_vcs_mcp.dialog_recovery.list_owned", return_value=[]),
    ):
        yield
    reset_interruptions()
    reset_access_gate()


def _status(backend, **kwargs):
    with patch("msaccess_vcs_mcp.dialog_recovery._default_backend", return_value=backend):
        return asyncio.run(tools.vcs_automation_status(DB, **kwargs))


def test_healthy_instance_is_ready():
    result = _status(ProbeBackend([_main()]), pid=10)
    assert result["ready"] is True
    assert "error_pattern" not in result


def test_dead_process_is_not_ready_and_says_not_running():
    backend = ProbeBackend([], {10: ProcessIdentity("MSACCESS.EXE", 1000, False)})
    result = _status(backend, pid=10)
    assert result["ready"] is False
    assert result["error_pattern"] == "access_not_running"


def test_dead_owned_instance_found_by_database_says_not_running():
    record = SimpleNamespace(pid=10, database_path=DB)
    backend = ProbeBackend([], {10: ProcessIdentity(None, None, False)})
    with patch("msaccess_vcs_mcp.dialog_recovery.list_owned", return_value=[record]):
        result = _status(backend)
    assert result["ready"] is False
    assert result["error_pattern"] == "access_not_running"


def test_live_process_without_windows_is_not_ready_and_says_no_windows():
    result = _status(ProbeBackend([_main(pid=99)]), pid=10)
    assert result["ready"] is False
    assert result["error_pattern"] == "no_windows_to_probe"


def test_unknown_liveness_is_not_ready():
    backend = ProbeBackend([_main()], {10: ProcessIdentity("MSACCESS.EXE", 1000, None)})
    result = _status(backend, pid=10)
    assert result["ready"] is False


@pytest.mark.parametrize(
    "identity",
    [ProcessIdentity("MSACCESS.EXE", None, True), ProcessIdentity(None, 1000, True)],
)
def test_unconfirmed_identity_is_not_ready(identity):
    result = _status(ProbeBackend([_main()], {10: identity}), pid=10)
    assert result["ready"] is False
    assert result["error_pattern"] == "identity_unconfirmed"


def test_unresponsive_is_not_ready():
    result = _status(ProbeBackend([_main()], responsive=False), pid=10)
    assert result["ready"] is False
    assert result["error_pattern"] == "access_unresponsive"


def test_break_mode_is_not_ready():
    paused = WindowInfo(hwnd=3, pid=10, title="Module1 (Code) [break]", class_name="wndclass_desked_gsk")
    result = _status(ProbeBackend([_main(), paused]), pid=10)
    assert result["ready"] is False
    assert result["error_pattern"] == "vba_break"


def test_blocking_dialog_is_not_ready():
    result = _status(ProbeBackend([_main(), _msgbox()]), pid=10)
    assert result["ready"] is False
    assert result["error_pattern"] == "blocking_dialog"


def _status_while_gate_held(backend, held_database):
    gate = get_access_gate()
    entered = threading.Event()

    async def hold():
        entered.set()
        await asyncio.sleep(0.3)

    async def main():
        task = asyncio.create_task(gate.run_exclusive("vcs_run_tests", held_database, hold, True))
        assert await asyncio.to_thread(entered.wait, 5)
        with patch("msaccess_vcs_mcp.dialog_recovery._default_backend", return_value=backend):
            result = await tools.vcs_automation_status(DB, pid=10)
        await task
        return result

    return asyncio.run(main())


def test_gate_busy_with_this_database_is_not_ready():
    result = _status_while_gate_held(ProbeBackend([_main()]), DB)
    assert result["ready"] is False
    assert result["error_pattern"] == "server_busy"


def test_gate_busy_with_another_database_does_not_block_readiness():
    result = _status_while_gate_held(ProbeBackend([_main()]), OTHER_DB)
    assert result["ready"] is True
