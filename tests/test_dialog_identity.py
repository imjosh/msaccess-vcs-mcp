"""Process identity, target resolution and pre-click re-verification, at the tool boundary.

Tests call the ``vcs_*`` dialog tools and assert on the returned dict and on what
the fake window backend recorded (clicks, closes).
"""

from __future__ import annotations

import asyncio
from unittest.mock import patch

import pytest

from msaccess_vcs_mcp import tools
from msaccess_vcs_mcp.access_gate import reset_access_gate
from msaccess_vcs_mcp.dialog_recovery import (
    ButtonInfo,
    ProcessIdentity,
    WindowInfo,
    reset_interruptions,
)

DB = r"C:\data\Northwind.accdb"
ACCESS = ProcessIdentity("MSACCESS.EXE", 1000, True)


class ScriptedBackend:
    """Window backend whose window list and process identities can change between calls."""

    def __init__(self, windows, identities=None):
        self._window_script = [list(windows)]
        self._identity_script = {pid: [ident] for pid, ident in (identities or {}).items()}
        self.clicked: list[int] = []
        self.closed: list[int] = []

    def script_windows(self, *snapshots):
        """List calls return these in order; the last one repeats."""
        self._window_script = [list(snapshot) for snapshot in snapshots]

    def script_identity(self, pid, *identities):
        self._identity_script[pid] = list(identities)

    def list_windows(self):
        if len(self._window_script) > 1:
            return self._window_script.pop(0)
        return list(self._window_script[0])

    def process_identity(self, pid):
        script = self._identity_script.get(pid, [ACCESS])
        if len(script) > 1:
            return script.pop(0)
        return script[0]

    def click(self, hwnd, *, expected_pid=None, timeout_ms=5000):
        self.clicked.append(hwnd)
        return True

    def close(self, hwnd):
        self.closed.append(hwnd)

    def responsive(self, hwnd, timeout_ms):
        return True


def _main(hwnd=1, pid=10, title="Northwind : Database"):
    return WindowInfo(hwnd=hwnd, pid=pid, title=title, class_name="OMain")


def _msgbox(hwnd=2, pid=10, button_hwnd=21, title="VCS Probe"):
    return WindowInfo(
        hwnd=hwnd,
        pid=pid,
        title=title,
        class_name="#32770",
        texts=("Hello",),
        buttons=(ButtonInfo(button_hwnd, "OK"),),
    )


@pytest.fixture(autouse=True)
def _clean(tmp_path):
    reset_interruptions()
    reset_access_gate()
    with (
        patch("msaccess_vcs_mcp.tools.validate_database_path", return_value=DB),
        patch("msaccess_vcs_mcp.dialog_recovery.list_owned", return_value=[]),
    ):
        yield
    reset_interruptions()
    reset_access_gate()


def _call(tool, backend, *args, **kwargs):
    with patch("msaccess_vcs_mcp.dialog_recovery._default_backend", return_value=backend):
        return asyncio.run(tool(*args, **kwargs))


def _dismiss(backend, **kwargs):
    kwargs.setdefault("button", "OK")
    return _call(tools.vcs_dismiss_dialog, backend, DB, "hwnd:2", **kwargs)


def test_unreadable_creation_time_refuses_click():
    backend = ScriptedBackend([_main(), _msgbox()], {10: ProcessIdentity("MSACCESS.EXE", None, True)})
    result = _dismiss(backend, pid=10)
    assert result["success"] is False
    assert result["error_pattern"] == "identity_unconfirmed"
    assert backend.clicked == []


def test_unknown_process_name_refuses_click_with_its_own_pattern():
    backend = ScriptedBackend([_main(), _msgbox()], {10: ProcessIdentity(None, 1000, True)})
    result = _dismiss(backend, pid=10)
    assert result["error_pattern"] == "identity_unconfirmed"
    assert result["error_pattern"] not in {"not_access_process", "process_identity_mismatch"}
    assert backend.clicked == []


def test_unconfirmed_identity_refuses_close_too():
    addin = WindowInfo(hwnd=2, pid=10, title="MSAccessVCS", class_name="OForm")
    backend = ScriptedBackend([_main(), addin], {10: ProcessIdentity(None, 1000, True)})
    result = _dismiss(backend, pid=10, button=None, action="close")
    assert result["error_pattern"] == "identity_unconfirmed"
    assert backend.closed == []


def test_unconfirmed_identity_still_inspects():
    backend = ScriptedBackend([_main(), _msgbox()], {10: ProcessIdentity(None, None, None)})
    result = _call(tools.vcs_list_dialogs, backend, DB, pid=10)
    assert result["success"] is True
    assert result["identity_confirmed"] is False
    assert [item["dialog_id"] for item in result["dialogs"]] == ["hwnd:2"]
    assert backend.clicked == []


def test_recover_with_unconfirmed_identity_clicks_nothing():
    backend = ScriptedBackend([_main(), _msgbox()], {10: ProcessIdentity("MSACCESS.EXE", None, True)})
    result = _call(tools.vcs_recover_dialogs, backend, DB, policy="safe", pid=10)
    assert result["error_pattern"] == "identity_unconfirmed"
    assert backend.clicked == []


def test_not_access_and_mismatch_keep_their_own_patterns():
    backend = ScriptedBackend([_main(), _msgbox()], {10: ProcessIdentity("Code.exe", 1000, True)})
    assert _dismiss(backend, pid=10)["error_pattern"] == "not_access_process"

    backend = ScriptedBackend([_main(), _msgbox()])
    assert _dismiss(backend, pid=10, create_time=999)["error_pattern"] == "process_identity_mismatch"
    assert backend.clicked == []


def test_caller_create_time_that_cannot_be_read_back_is_unconfirmed():
    backend = ScriptedBackend([_main(), _msgbox()], {10: ProcessIdentity("MSACCESS.EXE", None, True)})
    result = _dismiss(backend, pid=10, create_time=1000)
    assert result["error_pattern"] == "identity_unconfirmed"
    assert backend.clicked == []


def test_matching_identity_clicks_and_reports_observed_create_time():
    backend = ScriptedBackend([_main(), _msgbox()])
    backend.script_windows([_main(), _msgbox()], [_main(), _msgbox()], [_main()])
    result = _dismiss(backend, pid=10)
    assert result["success"] is True
    assert result["create_time"] == 1000
    assert backend.clicked == [21]


def test_creation_time_changing_before_click_is_dialog_changed():
    # No create_time passed: the value observed at resolution is fixed and rechecked.
    backend = ScriptedBackend([_main(), _msgbox()])
    backend.script_identity(10, ACCESS, ACCESS, ProcessIdentity("MSACCESS.EXE", 2000, True))
    result = _dismiss(backend, pid=10)
    assert result["success"] is False
    assert result["error_pattern"] == "dialog_changed"
    assert backend.clicked == []


def test_handle_reused_by_another_process_is_dialog_changed():
    backend = ScriptedBackend([_main(), _msgbox()])
    backend.script_windows(
        [_main(), _msgbox()],
        [_main(), _msgbox(pid=99)],
    )
    result = _dismiss(backend, pid=10)
    assert result["error_pattern"] == "dialog_changed"
    assert backend.clicked == []


def test_dialog_closed_before_click_is_dialog_changed():
    backend = ScriptedBackend([_main(), _msgbox()])
    backend.script_windows([_main(), _msgbox()], [_main()])
    result = _dismiss(backend, pid=10)
    assert result["error_pattern"] == "dialog_changed"
    assert backend.clicked == []


def test_button_that_left_the_dialog_is_dialog_changed():
    backend = ScriptedBackend([_main(), _msgbox()])
    backend.script_windows(
        [_main(), _msgbox()],
        [_main(), _msgbox(button_hwnd=77)],
    )
    result = _dismiss(backend, pid=10)
    assert result["error_pattern"] == "dialog_changed"
    assert backend.clicked == []


def test_close_reverifies_before_closing():
    addin = WindowInfo(hwnd=2, pid=10, title="MSAccessVCS", class_name="OForm")
    backend = ScriptedBackend([_main(), addin])
    backend.script_windows([_main(), addin], [_main()])
    result = _dismiss(backend, pid=10, button=None, action="close")
    assert result["error_pattern"] == "dialog_changed"
    assert backend.closed == []


def test_recover_skips_a_dialog_that_changed_and_clicks_nothing():
    backend = ScriptedBackend([_main(), _msgbox(title="Microsoft Access")])
    backend.script_windows([_main(), _msgbox(title="Microsoft Access")], [_main()])
    result = _call(tools.vcs_recover_dialogs, backend, DB, policy="safe", pid=10)
    assert backend.clicked == []
    assert [item["reason"] for item in result["skipped"]] == ["dialog_changed"]


def test_editor_window_with_database_name_does_not_make_target_ambiguous():
    editor = WindowInfo(hwnd=5, pid=30, title="Northwind.accdb - Notepad", class_name="Notepad")
    backend = ScriptedBackend(
        [_main(), _msgbox(), editor],
        {30: ProcessIdentity("notepad.exe", 5000, True)},
    )
    result = _call(tools.vcs_list_dialogs, backend, DB)
    assert result["success"] is True
    assert result["pid"] == 10


def test_editor_with_unreadable_name_does_not_hide_the_access_instance():
    editor = WindowInfo(hwnd=5, pid=30, title="Northwind.accdb - Editor", class_name="Editor")
    backend = ScriptedBackend([_main(), editor], {30: ProcessIdentity(None, None, None)})
    result = _call(tools.vcs_list_dialogs, backend, DB)
    assert result["success"] is True
    assert result["pid"] == 10


def test_two_access_instances_with_same_database_name_are_ambiguous():
    other = _main(hwnd=3, pid=20)
    backend = ScriptedBackend(
        [_main(), _msgbox(), other],
        {20: ProcessIdentity("MSACCESS.EXE", 2000, True)},
    )
    result = _dismiss(backend)
    assert result["error_pattern"] == "ambiguous_instance"
    assert {item["pid"] for item in result["candidates"]} == {10, 20}
    assert backend.clicked == []


def test_only_match_has_unreadable_process_name_is_unconfirmed_not_missing():
    backend = ScriptedBackend([_main()], {10: ProcessIdentity(None, None, None)})
    result = _dismiss(backend)
    assert result["error_pattern"] == "identity_unconfirmed"
    assert backend.clicked == []
