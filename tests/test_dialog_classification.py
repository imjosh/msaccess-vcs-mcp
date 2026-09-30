"""Dialog classification and the wait after an action, at the tool boundary.

Tests call ``vcs_list_dialogs``, ``vcs_dismiss_dialog`` and ``vcs_recover_dialogs``
and assert on the returned dictionary and on what the fake backend recorded.
"""

from __future__ import annotations

import asyncio
import time
from unittest.mock import patch

import pytest

from msaccess_vcs_mcp import tools
from msaccess_vcs_mcp.access_gate import reset_access_gate
from msaccess_vcs_mcp.dialog_recovery import (
    CLICK_DELIVERED,
    ButtonInfo,
    ProcessIdentity,
    WindowInfo,
    reset_interruptions,
)

DB = r"C:\data\Northwind.accdb"
ACCESS = ProcessIdentity("MSACCESS.EXE", 1000, True)


class SlowCloseBackend:
    """A click is delivered, but the dialog closes only after ``close_after`` more listings.

    ``close_after=None`` means it never closes.
    """

    def __init__(self, windows, close_after=0):
        self.windows = list(windows)
        self.close_after = close_after
        self.clicked: list[int] = []
        self.pressed: list[ButtonInfo] = []
        self.closed: list[int] = []
        self._pending_close: dict[int, int] = {}

    def process_identity(self, pid):
        return ACCESS

    def list_windows(self):
        for hwnd in list(self._pending_close):
            if self._pending_close[hwnd] <= 0:
                self.windows = [w for w in self.windows if w.hwnd != hwnd]
                del self._pending_close[hwnd]
            else:
                self._pending_close[hwnd] -= 1
        return list(self.windows)

    def click(self, button, *, expected_pid=None, timeout_ms=5000):
        self.clicked.append(button.hwnd)
        self.pressed.append(button)
        owner = next((w for w in self.windows if button in w.buttons), None)
        if owner is not None and self.close_after is not None:
            self._pending_close[owner.hwnd] = self.close_after
        return CLICK_DELIVERED

    def close(self, hwnd):
        self.closed.append(hwnd)

    def responsive(self, hwnd, timeout_ms):
        return True


def _main():
    return WindowInfo(hwnd=1, pid=10, title="Northwind : Database", class_name="OMain")


def _dialog(hwnd, title, text="Something happened.", buttons=("OK",)):
    return WindowInfo(
        hwnd=hwnd,
        pid=10,
        title=title,
        class_name="#32770",
        texts=(text,),
        buttons=tuple(ButtonInfo(hwnd * 10 + i, name) for i, name in enumerate(buttons)),
    )


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


def _call(tool, backend, *args, **kwargs):
    with patch("msaccess_vcs_mcp.dialog_recovery._default_backend", return_value=backend):
        return asyncio.run(tool(DB, *args, **kwargs))


def test_unrecognised_dialog_is_reported_as_unknown_with_title_text_and_buttons():
    odd = _dialog(2, "Contoso Tool", "Pick one", ("Alpha", "Beta"))
    result = _call(tools.vcs_list_dialogs, SlowCloseBackend([_main(), odd]))
    item = next(d for d in result["dialogs"] if d["dialog_id"] == "hwnd:2")
    assert item["kind"] == "unknown"
    assert item["title"] == "Contoso Tool"
    assert item["message"] == "Pick one"
    assert item["buttons"] == ["Alpha", "Beta"]
    assert result["blocking_dialog"] is True


def test_custom_caption_ok_only_msgbox_is_vba_msgbox_and_reported_with_title_and_text():
    box = _dialog(2, "Contoso Tool", "Hello")
    result = _call(tools.vcs_list_dialogs, SlowCloseBackend([_main(), box]))
    item = next(d for d in result["dialogs"] if d["dialog_id"] == "hwnd:2")
    assert item["kind"] == "vba_msgbox"
    assert item["title"] == "Contoso Tool"
    assert item["message"] == "Hello"
    assert item["buttons"] == ["OK"]
    assert result["blocking_dialog"] is True


def test_safe_policy_clicks_a_custom_caption_ok_only_msgbox_and_reports_it():
    box = _dialog(2, "Contoso Tool", "Hello")
    backend = SlowCloseBackend([_main(), box])
    result = _call(tools.vcs_recover_dialogs, backend, policy="safe", pid=10)
    assert backend.clicked == [20]
    assert result["success"] is True
    closed = result["closed"][0]
    assert closed["kind"] == "vba_msgbox"
    assert closed["title"] == "Contoso Tool"
    assert closed["message"] == "Hello"


def test_safe_policy_ignores_help_button_on_single_button_msgbox():
    box = _dialog(2, "Contoso Tool", "Hello", ("OK", "Help"))
    backend = SlowCloseBackend([_main(), box])
    _call(tools.vcs_recover_dialogs, backend, policy="safe", pid=10)
    assert backend.clicked == [20]


@pytest.mark.parametrize(
    "buttons", [("Yes", "No"), ("OK", "Cancel"), ("Alpha", "Beta"), ("Retry",), ("Save",)]
)
def test_custom_caption_msgbox_with_other_buttons_stays_unknown_and_is_not_clicked(buttons):
    box = _dialog(2, "Contoso Tool", "Pick one", buttons)
    backend = SlowCloseBackend([_main(), box])
    listed = _call(tools.vcs_list_dialogs, backend)
    assert next(d for d in listed["dialogs"] if d["dialog_id"] == "hwnd:2")["kind"] == "unknown"
    result = _call(tools.vcs_recover_dialogs, backend, policy="safe", pid=10)
    assert backend.clicked == []
    assert [item["kind"] for item in result["skipped"]] == ["unknown"]


@pytest.mark.parametrize("policy", ["safe", "end_runtime_error"])
@pytest.mark.parametrize(
    "text",
    ["Do you want to save changes?", "This will delete all rows.", "Discard your edits", "Overwrite it"],
)
def test_destructive_text_in_a_single_button_msgbox_stops_the_click(policy, text):
    box = _dialog(2, "Contoso Tool", text)
    backend = SlowCloseBackend([_main(), box])
    result = _call(tools.vcs_recover_dialogs, backend, policy=policy, pid=10)
    assert backend.clicked == []
    assert [item["kind"] for item in result["skipped"]] == ["vba_msgbox"]


def test_report_policy_never_clicks_a_single_button_msgbox():
    box = _dialog(2, "Contoso Tool", "Hello")
    backend = SlowCloseBackend([_main(), box])
    _call(tools.vcs_recover_dialogs, backend, policy="report", pid=10)
    assert backend.clicked == []


@pytest.mark.parametrize(
    ("kind", "title", "text"),
    [
        ("access_dialog", "Microsoft Access", "Something went wrong."),
        ("vba_compile_error", "Microsoft Visual Basic for Applications", "Compile error: Syntax error"),
    ],
)
def test_safe_policy_does_not_click_an_ok_only_access_dialog_or_compile_error(kind, title, text):
    known = _dialog(2, title, text)
    backend = SlowCloseBackend([_main(), known])
    listed = _call(tools.vcs_list_dialogs, backend)
    assert next(d for d in listed["dialogs"] if d["dialog_id"] == "hwnd:2")["kind"] == kind
    result = _call(tools.vcs_recover_dialogs, backend, policy="safe", pid=10)
    assert backend.clicked == []
    assert [item["kind"] for item in result["skipped"]] == [kind]


def test_end_runtime_error_policy_clicks_ok_on_a_compile_error_but_not_an_access_dialog():
    compile_err = _dialog(
        2, "Microsoft Visual Basic for Applications", "Compile error: Syntax error"
    )
    access = _dialog(3, "Microsoft Access", "Something went wrong.")
    backend = SlowCloseBackend([_main(), compile_err, access])
    _call(tools.vcs_recover_dialogs, backend, policy="end_runtime_error", pid=10)
    assert backend.clicked == [20]


def test_runtime_error_with_destructive_text_gets_end_under_end_runtime_error_and_nothing_under_safe():
    runtime = _dialog(
        2,
        "Microsoft Visual Basic",
        "Run-time error '3021': the record was deleted; discard the edit",
        ("End", "Debug", "Help"),
    )
    safe_backend = SlowCloseBackend([_main(), runtime])
    _call(tools.vcs_recover_dialogs, safe_backend, policy="safe", pid=10)
    assert safe_backend.clicked == []

    end_backend = SlowCloseBackend([_main(), runtime])
    result = _call(tools.vcs_recover_dialogs, end_backend, policy="end_runtime_error", pid=10)
    assert end_backend.clicked == [20]  # End, never Debug
    assert result["interrupted"] is True


def test_safe_policy_skips_a_recognised_dialog_with_destructive_text():
    known = _dialog(2, "Microsoft Access", "Do you want to save changes?")
    backend = SlowCloseBackend([_main(), known])
    _call(tools.vcs_recover_dialogs, backend, policy="safe", pid=10)
    assert backend.clicked == []


def test_explicit_dismiss_can_still_click_an_unknown_dialog():
    odd = _dialog(2, "Contoso Tool", "Hello", ("Alpha", "Beta"))
    backend = SlowCloseBackend([_main(), odd])
    result = _call(tools.vcs_dismiss_dialog, backend, "hwnd:2", button="Alpha", pid=10)
    assert backend.clicked == [20]
    assert result["success"] is True


def test_debug_is_never_clicked_by_either_tool():
    runtime = _dialog(
        2, "Microsoft Visual Basic", "Run-time error '11':", ("End", "Debug", "Help")
    )
    backend = SlowCloseBackend([_main(), runtime])
    refused = _call(tools.vcs_dismiss_dialog, backend, "hwnd:2", button="Debug", pid=10)
    assert refused["error_pattern"] == "debug_refused"
    _call(tools.vcs_recover_dialogs, backend, policy="end_runtime_error", pid=10)
    assert backend.clicked == [20]  # End only


@pytest.mark.parametrize("tool_name", ["dismiss", "recover"])
def test_slow_close_is_waited_for_not_reported_as_still_open(tool_name):
    known = _dialog(2, "Contoso Tool", "Something went wrong.")
    backend = SlowCloseBackend([_main(), known], close_after=4)
    if tool_name == "dismiss":
        result = _call(tools.vcs_dismiss_dialog, backend, "hwnd:2", button="OK", pid=10)
    else:
        result = _call(tools.vcs_recover_dialogs, backend, policy="safe", pid=10)
    assert result["success"] is True
    assert result.get("error_pattern") != "dismiss_uncertain"
    assert result["dialogs"] == []


@pytest.mark.parametrize("tool_name", ["dismiss", "recover"])
def test_dialog_still_open_at_the_deadline_is_dismiss_uncertain(tool_name):
    known = _dialog(2, "Contoso Tool", "Something went wrong.")
    backend = SlowCloseBackend([_main(), known], close_after=None)
    started = time.monotonic()
    if tool_name == "dismiss":
        result = _call(
            tools.vcs_dismiss_dialog, backend, "hwnd:2", button="OK", pid=10, timeout_seconds=0.3
        )
    else:
        result = _call(
            tools.vcs_recover_dialogs, backend, policy="safe", pid=10, timeout_seconds=0.3
        )
    assert result["success"] is False
    assert result["error_pattern"] == "dismiss_uncertain"
    assert time.monotonic() - started < 3
    assert backend.clicked == [20]  # not retried


def test_recover_stops_early_when_only_uncovered_dialogs_remain():
    known = _dialog(2, "Contoso Tool", "Something went wrong.")
    odd = _dialog(3, "Contoso Tool", "Hello", ("Alpha", "Beta"))
    backend = SlowCloseBackend([_main(), known, odd], close_after=0)
    started = time.monotonic()
    result = _call(tools.vcs_recover_dialogs, backend, policy="safe", pid=10, timeout_seconds=10)
    assert time.monotonic() - started < 3
    assert result.get("error_pattern") != "dismiss_uncertain"
    assert [d["dialog_id"] for d in result["dialogs"]] == ["hwnd:3"]
    assert [d["kind"] for d in result["dialogs"]] == ["unknown"]


def test_dialog_ids_from_listing_are_accepted_and_malformed_ids_are_not_found():
    known = _dialog(2, "Microsoft Access", "Something went wrong.")
    backend = SlowCloseBackend([_main(), known])
    listed = _call(tools.vcs_list_dialogs, backend)
    dialog_id = next(d["dialog_id"] for d in listed["dialogs"] if d["kind"] == "access_dialog")
    bad = _call(tools.vcs_dismiss_dialog, backend, "bogus", button="OK", pid=10)
    assert bad["error_pattern"] == "dialog_not_found"
    assert backend.clicked == []
    ok = _call(tools.vcs_dismiss_dialog, backend, dialog_id, button="OK", pid=10)
    assert ok["success"] is True


# Recorded from Microsoft 365 Access (2026-09-30). ``MsgBox`` with the ``@`` bold
# form (the add-in's ``MsgBox2``) and Access's own error dialogs are a top-level
# ``NUIDialog`` with one ``NetUIHWND`` child and no Win32 buttons. MSAA lists the
# caption as the first static text, then the message lines, then the visible push
# buttons; each button is addressed by its host window and child-index path.
NETUI_HOST = 77


def _netui(hwnd, title, texts, buttons, first_path=9):
    return WindowInfo(
        hwnd=hwnd,
        pid=10,
        title=title,
        class_name="NUIDialog",
        texts=(title, *texts),
        buttons=tuple(
            ButtonInfo(NETUI_HOST, name, (0, first_path + index))
            for index, name in enumerate(buttons)
        ),
    )


def _msgbox2_ok(first_path=9):
    return _netui(
        2,
        "Version Control Add-in",
        ("Bold part", "line one here\n\nline two"),
        ("OK",),
        first_path,
    )


def test_netui_ok_only_msgbox_is_listed_as_a_blocking_vba_msgbox():
    result = _call(tools.vcs_list_dialogs, SlowCloseBackend([_main(), _msgbox2_ok()]))
    item = next(d for d in result["dialogs"] if d["dialog_id"] == "hwnd:2")
    assert item["kind"] == "vba_msgbox"
    assert item["class_name"] == "NUIDialog"
    assert item["title"] == "Version Control Add-in"
    assert item["message"] == "Bold part\nline one here\n\nline two"
    assert item["buttons"] == ["OK"]
    assert result["blocking_dialog"] is True
    assert result["ready"] is False


def test_safe_policy_presses_ok_on_a_netui_ok_only_msgbox():
    backend = SlowCloseBackend([_main(), _msgbox2_ok()])
    result = _call(tools.vcs_recover_dialogs, backend, policy="safe", pid=10)
    assert backend.pressed == [ButtonInfo(NETUI_HOST, "OK", (0, 9))]
    assert result["success"] is True
    assert [item["kind"] for item in result["closed"]] == ["vba_msgbox"]
    assert result["dialogs"] == []


def test_dismiss_presses_ok_on_a_netui_msgbox():
    backend = SlowCloseBackend([_main(), _msgbox2_ok()])
    result = _call(tools.vcs_dismiss_dialog, backend, "hwnd:2", button="OK", pid=10)
    assert result["success"] is True
    assert result["dismissed"] is True
    assert backend.pressed == [ButtonInfo(NETUI_HOST, "OK", (0, 9))]


def test_netui_yes_no_box_is_unknown_not_clicked_by_safe_and_can_be_answered_explicitly():
    box = _netui(2, "Version Control Add-in", ("Keep it?",), ("Yes", "No"))
    backend = SlowCloseBackend([_main(), box])
    listed = _call(tools.vcs_list_dialogs, backend)
    assert next(d for d in listed["dialogs"] if d["dialog_id"] == "hwnd:2")["kind"] == "unknown"
    safe = _call(tools.vcs_recover_dialogs, backend, policy="safe", pid=10)
    assert backend.pressed == []
    assert [item["kind"] for item in safe["skipped"]] == ["unknown"]
    answered = _call(tools.vcs_dismiss_dialog, backend, "hwnd:2", button="No", pid=10)
    assert answered["success"] is True
    assert backend.pressed == [ButtonInfo(NETUI_HOST, "No", (0, 10))]


def test_netui_access_error_is_an_access_dialog_and_safe_does_not_click_it():
    error = _netui(
        2,
        "Microsoft Access",
        ("The form name 'NoSuchForm' is misspelled or refers to a form that doesn't exist.",),
        ("OK",),
    )
    backend = SlowCloseBackend([_main(), error])
    listed = _call(tools.vcs_list_dialogs, backend)
    assert next(d for d in listed["dialogs"] if d["dialog_id"] == "hwnd:2")["kind"] == "access_dialog"
    _call(tools.vcs_recover_dialogs, backend, policy="safe", pid=10)
    assert backend.pressed == []


def test_netui_ok_only_msgbox_with_destructive_text_is_not_clicked_by_safe():
    box = _netui(2, "Version Control Add-in", ("This will delete all rows.",), ("OK",))
    backend = SlowCloseBackend([_main(), box])
    _call(tools.vcs_recover_dialogs, backend, policy="safe", pid=10)
    assert backend.pressed == []


class MovedButtonBackend(SlowCloseBackend):
    """From the second listing on, the OK button is at a different MSAA path."""

    def __init__(self, windows):
        super().__init__(windows)
        self.listings = 0

    def list_windows(self):
        self.listings += 1
        if self.listings == 2:
            self.windows = [_msgbox2_ok(first_path=12) if w.hwnd == 2 else w for w in self.windows]
        return super().list_windows()


def test_netui_press_is_refused_when_the_button_moved_since_inspection():
    backend = MovedButtonBackend([_main(), _msgbox2_ok()])
    result = _call(tools.vcs_dismiss_dialog, backend, "hwnd:2", button="OK", pid=10)
    assert result["error_pattern"] == "dialog_changed"
    assert backend.pressed == []
