"""Dialog inspection and recovery that does not require Access."""

from __future__ import annotations

import asyncio
from unittest.mock import patch

import pytest

from msaccess_vcs_mcp.dialog_recovery import (
    ButtonInfo,
    ProcessIdentity,
    WindowInfo,
    auto_button,
    classify_window,
    dismiss_dialog,
    dismiss_one,
    inspect_windows,
    recover_windows,
    reset_interruptions,
    resolve_target,
)
from msaccess_vcs_mcp.access_gate import get_access_gate, reset_access_gate


class FakeBackend:
    def __init__(self, windows, identities=None):
        self.windows = list(windows)
        self.identities = dict(identities or {})
        self.clicked: list[int] = []
        self.closed: list[int] = []

    def process_identity(self, pid: int) -> ProcessIdentity:
        return self.identities.get(pid, ProcessIdentity("MSACCESS.EXE", 1000, True))

    def list_windows(self):
        return list(self.windows)

    def click(self, hwnd: int, *, expected_pid=None, timeout_ms=5000) -> bool:
        self.clicked.append(hwnd)
        self.windows = [window for window in self.windows if window.hwnd != hwnd and not any(
            button.hwnd == hwnd for button in window.buttons
        )]
        # Drop the dialog whose button was clicked.
        self.windows = [
            window
            for window in self.windows
            if not any(button.hwnd == hwnd for button in window.buttons)
        ]
        return True

    def close(self, hwnd: int) -> None:
        self.closed.append(hwnd)
        self.windows = [window for window in self.windows if window.hwnd != hwnd]

    def responsive(self, hwnd: int, timeout_ms: int) -> bool:
        return True


def _button(hwnd: int, text: str) -> ButtonInfo:
    return ButtonInfo(hwnd, text)


def _win(**kwargs) -> WindowInfo:
    defaults = dict(
        hwnd=1,
        pid=10,
        title="",
        class_name="#32770",
        texts=(),
        buttons=(),
    )
    defaults.update(kwargs)
    return WindowInfo(**defaults)


@pytest.fixture(autouse=True)
def _clear_interruptions():
    reset_interruptions()
    reset_access_gate()
    yield
    reset_interruptions()
    reset_access_gate()


def test_classify_runtime_compile_break_and_msgbox():
    runtime = _win(
        title="Microsoft Visual Basic",
        texts=("Run-time error '13':", "Type mismatch"),
        buttons=(_button(2, "End"), _button(3, "Debug"), _button(4, "Help")),
    )
    compile_error = _win(
        hwnd=11,
        title="Microsoft Visual Basic",
        texts=("Compile error:", "Expected: end of statement"),
        buttons=(_button(12, "OK"), _button(13, "Help")),
    )
    paused = _win(
        hwnd=21,
        title="Northwind - Module1 (Code) [break]",
        class_name="wndclass_desked_gsk",
        texts=(),
        buttons=(),
    )
    msg = _win(hwnd=31, title="VCS Probe", texts=("Hello",), buttons=(_button(32, "OK"),))
    addin = _win(hwnd=41, title="MSAccessVCS", class_name="OForm", texts=("Merge Complete",), buttons=())

    assert classify_window(runtime) == "vba_runtime_error"
    assert classify_window(compile_error) == "vba_compile_error"
    assert classify_window(paused) == "vba_break"
    assert classify_window(msg) == "vba_msgbox"
    assert classify_window(addin) == "addin_window"


def test_break_mode_is_not_a_dialog_and_blocks_readiness():
    windows = [
        _win(hwnd=1, title="Northwind : Database", class_name="OMain", texts=()),
        _win(hwnd=2, title="Module1 (Code) [break]", class_name="wndclass_desked_gsk"),
    ]
    report = inspect_windows(windows, r"C:\data\Northwind.accdb", backend=FakeBackend(windows), responsive=True)
    assert report["success"] is True
    assert report["break_mode"] is True
    assert report["ready"] is False
    kinds = {item["kind"] for item in report["dialogs"]}
    assert "vba_break" in kinds
    assert report["dialogs"][0]["is_dialog"] is False or any(
        item["kind"] == "vba_break" and item["is_dialog"] is False for item in report["dialogs"]
    )


def test_two_instances_only_the_named_pid_is_touched():
    windows = [
        _win(hwnd=1, pid=10, title="VCS Probe", texts=("one",), buttons=(_button(11, "OK"),)),
        _win(
            hwnd=2,
            pid=20,
            title="Other",
            texts=("two",),
            buttons=(_button(21, "OK"),),
        ),
    ]
    backend = FakeBackend(windows)
    result = dismiss_one(
        windows,
        r"C:\data\Northwind.accdb",
        "hwnd:1",
        button="OK",
        pid=10,
        create_time=1000,
        responsive=True,
        backend=backend,
    )
    assert result["success"] is True
    assert backend.clicked == [11]
    assert 21 not in backend.clicked
    assert result["pid"] == 10


def test_ambiguous_database_is_not_clicked():
    windows = [
        _win(hwnd=1, pid=10, title="Northwind : Database", class_name="OMain"),
        _win(hwnd=2, pid=20, title="Northwind : Database", class_name="OMain"),
    ]
    resolved = resolve_target(windows, r"C:\data\Northwind.accdb", FakeBackend(windows))
    assert resolved["success"] is False
    assert resolved["error_pattern"] == "ambiguous_instance"
    assert len(resolved["candidates"]) == 2


def test_create_time_mismatch_refuses_the_pid():
    windows = [_win(hwnd=1, pid=10, title="Northwind", class_name="OMain")]
    resolved = resolve_target(windows, r"C:\data\Northwind.accdb", FakeBackend(windows), pid=10, create_time=999)
    assert resolved["error_pattern"] == "process_identity_mismatch"


def test_safe_policy_acks_ok_only_and_skips_unknown_and_destructive():
    ok_dialog = _win(hwnd=1, title="Microsoft Access", texts=("Something went wrong.",), buttons=(_button(11, "OK"),))
    yes_no = _win(
        hwnd=2,
        title="Confirm",
        texts=("Delete these records?",),
        buttons=(_button(21, "Yes"), _button(22, "No")),
    )
    save = _win(
        hwnd=3,
        title="Microsoft Access",
        texts=("Do you want to save changes?",),
        buttons=(_button(31, "Yes"), _button(32, "No")),
    )
    runtime = _win(
        hwnd=4,
        title="Microsoft Visual Basic",
        texts=("Run-time error '11':", "Division by zero"),
        buttons=(_button(41, "End"), _button(42, "Debug"), _button(43, "Help")),
    )
    windows = [ok_dialog, yes_no, save, runtime]
    backend = FakeBackend(windows)
    result = recover_windows(
        windows,
        r"C:\data\Northwind.accdb",
        policy="safe",
        pid=10,
        responsive=True,
        backend=backend,
    )
    assert backend.clicked == [11]
    assert 42 not in backend.clicked
    assert 21 not in backend.clicked
    skipped_kinds = {item["kind"] for item in result["skipped"]}
    assert "vba_runtime_error" in skipped_kinds
    assert "unknown" in skipped_kinds
    assert result["failure_dialog_dismissed"] is False


def test_end_runtime_error_clicks_end_and_records_interruption():
    runtime = _win(
        hwnd=4,
        title="Microsoft Visual Basic",
        texts=("Run-time error '11':", "Division by zero"),
        buttons=(_button(41, "End"), _button(42, "Debug"), _button(43, "Help")),
    )
    backend = FakeBackend([runtime])
    result = recover_windows(
        [runtime],
        r"C:\data\Northwind.accdb",
        policy="end_runtime_error",
        pid=10,
        responsive=True,
        backend=backend,
    )
    assert backend.clicked == [41]
    assert result["interrupted"] is True
    assert result["failure_dialog_dismissed"] is True
    follow = inspect_windows([runtime], r"C:\data\Northwind.accdb", backend=backend, pid=10, responsive=True)
    assert follow["execution_interrupted"] is True
    assert follow["last_interruption"]["button"] == "End"
    assert follow["last_interruption"]["message"].startswith("Run-time error")


def test_accelerator_captions_match_end_no_and_debug():
    runtime = _win(
        hwnd=4,
        title="Microsoft Visual Basic",
        texts=(),
        buttons=(
            _button(41, "&Continue"),
            _button(42, "&End"),
            _button(43, "&Debug"),
            _button(44, "&Help"),
        ),
    )
    assert classify_window(runtime) == "vba_runtime_error"
    backend = FakeBackend([runtime])
    result = recover_windows(
        [runtime],
        r"C:\data\Northwind.accdb",
        policy="end_runtime_error",
        pid=10,
        responsive=True,
        backend=backend,
    )
    assert backend.clicked == [42]
    assert result["interrupted"] is True
    refused = dismiss_one(
        [runtime],
        r"C:\data\Northwind.accdb",
        "hwnd:4",
        button="&Debug",
        pid=10,
        responsive=True,
        backend=FakeBackend([runtime]),
    )
    assert refused["error_pattern"] == "debug_refused"
    yes_no = _win(
        hwnd=5,
        title="Confirm",
        texts=("Keep this record?",),
        buttons=(_button(51, "&Yes"), _button(52, "&No")),
    )
    chooser = FakeBackend([yes_no])
    dismissed = dismiss_one(
        [yes_no],
        r"C:\data\Northwind.accdb",
        "hwnd:5",
        button="No",
        pid=10,
        responsive=True,
        backend=chooser,
    )
    assert dismissed["success"] is True
    assert chooser.clicked == [52]


def test_explicit_debug_is_refused():
    runtime = _win(
        hwnd=4,
        title="Microsoft Visual Basic",
        texts=("Run-time error '11':", "Division by zero"),
        buttons=(_button(41, "End"), _button(42, "Debug")),
    )
    backend = FakeBackend([runtime])
    result = dismiss_one(
        [runtime],
        r"C:\data\Northwind.accdb",
        "hwnd:4",
        button="Debug",
        pid=10,
        responsive=True,
        backend=backend,
    )
    assert result["error_pattern"] == "debug_refused"
    assert backend.clicked == []


def test_close_finished_addin_window_does_not_cancel_while_gate_busy_on_same_database():
    addin = _win(hwnd=7, title="MSAccessVCS", class_name="OForm", texts=("Running",))
    backend = FakeBackend([addin])
    with patch("msaccess_vcs_mcp.dialog_recovery._gate_snapshot", return_value={"gate_busy": True, "operation": {"tool": "vcs_run_tests", "same_database": True}}):
        blocked = dismiss_one(
            [addin],
            r"C:\data\Northwind.accdb",
            "hwnd:7",
            action="close",
            pid=10,
            responsive=True,
            backend=backend,
        )
    assert blocked["error_pattern"] == "operation_in_progress"
    assert backend.closed == []

    other = {"gate_busy": True, "operation": {"tool": "vcs_run_tests", "same_database": False}}
    with patch("msaccess_vcs_mcp.dialog_recovery._gate_snapshot", return_value=other):
        closed = dismiss_one(
            [addin],
            r"C:\data\Northwind.accdb",
            "hwnd:7",
            action="close",
            pid=10,
            responsive=True,
            backend=backend,
        )
    assert closed["success"] is True
    assert closed["interrupted"] is False
    assert backend.closed == [7]


def test_cancel_addin_window_is_an_interruption():
    addin = _win(hwnd=7, title="MSAccessVCS", class_name="OForm")
    backend = FakeBackend([addin])
    result = dismiss_one(
        [addin],
        r"C:\data\Northwind.accdb",
        "hwnd:7",
        action="cancel",
        pid=10,
        responsive=True,
        backend=backend,
    )
    assert result["interrupted"] is True
    assert backend.closed == [7]


def test_unresponsive_process_is_not_clicked():
    dialog = _win(hwnd=1, title="VCS Probe", texts=("Hi",), buttons=(_button(11, "OK"),))
    backend = FakeBackend([dialog])
    result = dismiss_one(
        [dialog],
        r"C:\data\Northwind.accdb",
        "hwnd:1",
        button="OK",
        pid=10,
        responsive=False,
        backend=backend,
    )
    assert result["error_pattern"] == "access_unresponsive"
    assert backend.clicked == []


def test_auto_button_never_returns_debug_or_yes():
    runtime = _win(
        title="Microsoft Visual Basic",
        texts=("Run-time error '5':",),
        buttons=(_button(1, "End"), _button(2, "Debug")),
    )
    assert auto_button(runtime, "vba_runtime_error", "safe") is None
    assert auto_button(runtime, "vba_runtime_error", "end_runtime_error") == "End"
    confirm = _win(texts=("Overwrite?",), buttons=(_button(3, "Yes"), _button(4, "No")))
    assert auto_button(confirm, "unknown", "safe") is None


def test_dialog_tools_run_while_the_access_gate_is_held(tmp_path):
    """Recovery must not wait behind the COM apartment lock."""
    db = tmp_path / "Northwind.accdb"
    db.write_bytes(b"")
    gate = get_access_gate()
    entered = asyncio.Event()

    async def hold():
        entered.set()
        await asyncio.sleep(0.3)
        return "held"

    async def main():
        task = asyncio.create_task(
            gate.run_exclusive("vcs_run_tests", str(db), hold, True)
        )
        await entered.wait()
        with (
            patch("msaccess_vcs_mcp.access_gate._read_busy_wait_sec", return_value=0.05),
            patch(
                "msaccess_vcs_mcp.tools.list_dialogs",
                return_value={"success": True, "exempt_probe": True, "dialogs": []},
            ),
            patch("msaccess_vcs_mcp.tools.validate_database_path", return_value=db),
        ):
            from msaccess_vcs_mcp.tools import vcs_list_dialogs

            result = await vcs_list_dialogs(str(db))
        await task
        return result

    result = asyncio.run(main())
    assert result.get("error_pattern") != "server_busy"
    assert result.get("exempt_probe") is True


class HungClickBackend(FakeBackend):
    """Click is never delivered (target thread hung); the dialog stays open."""

    def click(self, hwnd: int, *, expected_pid=None, timeout_ms=5000) -> bool:
        self.clicked.append(hwnd)
        return False


def test_undelivered_click_is_dismiss_uncertain_and_not_retried():
    import time

    windows = [
        _win(hwnd=1, pid=10, title="VCS Probe", texts=("one",), buttons=(_button(11, "OK"),)),
    ]
    backend = HungClickBackend(windows)
    started = time.monotonic()
    result = dismiss_dialog(
        r"C:\data\Northwind.accdb",
        "hwnd:1",
        button="OK",
        pid=10,
        create_time=1000,
        timeout_seconds=1.0,
        backend=backend,
    )
    assert time.monotonic() - started < 1.0 + 5.0
    assert result["success"] is False
    assert result["error_pattern"] == "dismiss_uncertain"
    assert backend.clicked == [11]


def test_exempt_tool_runs_off_event_loop(monkeypatch):
    import asyncio
    import threading
    import time

    from msaccess_vcs_mcp import tools

    main_thread = threading.get_ident()
    seen: dict[str, int] = {}

    def blocking(**_kwargs):
        seen["thread"] = threading.get_ident()
        time.sleep(0.3)
        return {"success": True}

    async def scenario():
        ticks = 0

        async def ticker():
            nonlocal ticks
            while True:
                await asyncio.sleep(0.02)
                ticks += 1

        task = asyncio.create_task(ticker())
        result = await tools._run_exempt_in_worker("vcs_list_dialogs", blocking, (), {})
        task.cancel()
        return result, ticks

    result, ticks = asyncio.run(scenario())
    assert result == {"success": True}
    assert seen["thread"] != main_thread
    assert ticks >= 5


def test_exempt_tool_ceiling_returns_recoverable_timeout(monkeypatch):
    import asyncio
    import time

    from msaccess_vcs_mcp import tools

    monkeypatch.setattr(tools, "EXEMPT_WORKER_MARGIN_SEC", 0.0)
    monkeypatch.setenv("ACCESS_VCS_DIALOG_TIMEOUT_SEC", "0.2")

    def hung(**_kwargs):
        time.sleep(1.0)
        return {"success": True}

    result = asyncio.run(tools._run_exempt_in_worker("vcs_list_dialogs", hung, (), {}))
    assert result["success"] is False
    assert result["error_pattern"] == "tool_timeout"
    assert result["recoverable"] is True
