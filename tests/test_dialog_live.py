"""Live dialog recovery against disposable Access databases.

These tests launch Microsoft Access on the interactive desktop. They are
marked integration so the default suite does not start Access.
"""

from __future__ import annotations

import json
import os
import queue
import subprocess
import sys
import tempfile
import threading
import time
import uuid
from contextlib import contextmanager
from dataclasses import dataclass, field
from pathlib import Path

import pytest

from msaccess_vcs_mcp.dialog_recovery import (
    _button_label,
    automation_status,
    dismiss_dialog,
    list_dialogs,
    recover_dialogs,
    reset_interruptions,
)


def _has_button(buttons: list, name: str) -> bool:
    return any(_button_label(str(item)) == name for item in buttons)


@dataclass(frozen=True)
class _AccessIdentity:
    pid: int
    create_time: int


@dataclass
class _Child:
    proc: subprocess.Popen[str]
    lines: queue.Queue[str] = field(default_factory=queue.Queue)
    identity: _AccessIdentity | None = None


def _creation_time(handle) -> int:
    import win32process

    creation = win32process.GetProcessTimes(handle)["CreationTime"]
    return int(creation.timestamp() * 10_000_000)


@contextmanager
def _confirmed_process(identity: _AccessIdentity | None, *, terminate: bool = False):
    """Keep the verified handle open so PID reuse cannot redirect termination."""
    handle = None
    try:
        import win32api
        import win32con

        if identity is not None and identity.pid > 0 and identity.create_time > 0:
            access = win32con.PROCESS_QUERY_LIMITED_INFORMATION
            if terminate:
                access |= win32con.PROCESS_TERMINATE | win32con.SYNCHRONIZE
            handle = win32api.OpenProcess(access, False, identity.pid)
            matched = _creation_time(handle) == identity.create_time
        else:
            matched = False
    except Exception:
        matched = False
    try:
        yield handle if matched else None
    finally:
        if handle is not None:
            win32api.CloseHandle(handle)


def _matches(identity: _AccessIdentity | None) -> bool:
    with _confirmed_process(identity) as handle:
        return handle is not None


def _kill(identity: _AccessIdentity | None) -> None:
    try:
        with _confirmed_process(identity, terminate=True) as handle:
            if handle is not None:
                import win32api
                import win32event

                win32api.TerminateProcess(handle, 1)
                win32event.WaitForSingleObject(handle, 5000)
    except Exception:
        pass


def _read_json(child: _Child, timeout: float) -> dict | None:
    try:
        line = child.lines.get(timeout=max(0, timeout))
    except queue.Empty:
        return None
    if not line:
        return None
    try:
        payload = json.loads(line)
    except ValueError:
        return {"event": "error", "error": "invalid child JSON"}
    if not isinstance(payload, dict):
        return {"event": "error", "error": "invalid child payload"}
    if payload.get("event") == "launched" and child.identity is None:
        pid, stamp = payload.get("pid"), payload.get("create_time")
        if type(pid) is int and pid > 0 and type(stamp) is int and stamp > 0:
            child.identity = _AccessIdentity(pid, stamp)
    return payload


def _start_child(db_path: Path, scenario: str) -> _Child:
    proc = subprocess.Popen(
        [sys.executable, str(Path(__file__).resolve()), "--child", str(db_path), scenario],
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )
    child = _Child(proc)

    def _reader() -> None:
        if proc.stdout is not None:
            for line in proc.stdout:
                child.lines.put(line)
        child.lines.put("")

    threading.Thread(target=_reader, daemon=True).start()
    return child


def _wait_ready(child: _Child, timeout: float = 45) -> dict:
    deadline = time.monotonic() + timeout
    payload = None
    while time.monotonic() < deadline:
        payload = _read_json(child, deadline - time.monotonic())
        if payload and payload.get("event") == "launched":
            continue
        if (
            payload
            and payload.get("event") == "ready"
            and child.identity is not None
            and payload.get("pid") == child.identity.pid
            and _matches(child.identity)
        ):
            return payload
        break
    raise AssertionError(f"Access child did not become ready ({payload!r}).")


def _wait_dialog(db_path: Path, pid: int, title: str, timeout: float = 20) -> dict:
    deadline = time.monotonic() + timeout
    last: dict = {}
    while time.monotonic() < deadline:
        last = list_dialogs(str(db_path), pid=pid, timeout_seconds=2)
        for item in last.get("dialogs") or []:
            if title.lower() in (item.get("title") or "").lower():
                return last
            if title.lower() in (item.get("message") or "").lower():
                return last
        time.sleep(0.25)
    return last


def _stop(child: _Child) -> None:
    # A sibling may have launched but never reached _wait_ready after another
    # child's failure. Consume its launch handshake before cleaning it up.
    if child.identity is None:
        _read_json(child, 2)
    if _matches(child.identity) and child.proc.stdin is not None:
        try:
            child.proc.stdin.write("quit\n")
            child.proc.stdin.flush()
        except Exception:
            pass
    try:
        child.proc.wait(timeout=8)
    except subprocess.TimeoutExpired:
        pass
    # Even an exited Python helper may have left its Access process behind.
    _kill(child.identity)
    if child.proc.poll() is None:
        child.proc.kill()
        child.proc.wait(timeout=8)


def _app_pid(app) -> tuple[int, int]:
    import win32process

    hwnd_value = app.hWndAccessApp
    hwnd = int(hwnd_value() if callable(hwnd_value) else hwnd_value)
    _thread, pid = win32process.GetWindowThreadProcessId(hwnd)
    return int(pid), hwnd


def _quit_owned(app, identity: _AccessIdentity | None) -> None:
    try:
        pid, _hwnd = _app_pid(app)
        if identity is not None and pid == identity.pid and _matches(identity):
            app.Quit(2)
    except Exception:
        pass


def _child_main(db_path: str, scenario: str) -> int:
    import pythoncom
    import win32com.client
    import win32api
    import win32con

    pythoncom.CoInitialize()
    app = win32com.client.DispatchEx("Access.Application")
    identity = None
    try:
        # Claim this DispatchEx instance before database creation or any other
        # COM work that can fail or block before readiness.
        pid, hwnd = _app_pid(app)
        handle = win32api.OpenProcess(win32con.PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
        try:
            identity = _AccessIdentity(pid, _creation_time(handle))
        finally:
            win32api.CloseHandle(handle)
        print(json.dumps({
            "event": "launched", "pid": pid, "create_time": identity.create_time,
        }), flush=True)
        app.Visible = True
        if os.path.exists(db_path):
            os.remove(db_path)
        app.NewCurrentDatabase(db_path)
        print(json.dumps({"event": "ready", "pid": int(pid), "hwnd": hwnd}), flush=True)
        if scenario == "idle":
            sys.stdin.readline()
        elif scenario == "msgbox_ok":
            app.Eval('MsgBox("dialog recovery probe",0,"ProbeTitle")')
            print(json.dumps({"event": "done"}), flush=True)
        elif scenario == "msgbox2_ok":
            # The exact Eval string the add-in's MsgBox2 builds; Access draws it as a NUIDialog.
            app.Eval("MsgBox('Probe bold@netui probe line@@',0,'NetUIProbe','',0)")
            print(json.dumps({"event": "done"}), flush=True)
        elif scenario == "access_error":
            # Without UserControl Access returns the error to COM instead of showing it.
            app.UserControl = True
            try:
                app.DoCmd.OpenForm("NoSuchProbeForm")
            except Exception as exc:
                print(json.dumps({"event": "done", "error": str(exc)}), flush=True)
            else:
                print(json.dumps({"event": "done"}), flush=True)
        elif scenario == "msgbox_yesno":
            result = app.Eval('MsgBox("Keep this record?",4,"ConfirmKeep")')
            print(json.dumps({"event": "done", "result": int(result)}), flush=True)
        elif scenario == "runtime":
            _install_proc(
                app,
                "Public Sub Boom()\r\nErr.Raise 11, \"Probe\", \"division by zero probe\"\r\nEnd Sub\r\n",
            )
            app.SetOption("Error Trapping", 0)
            try:
                app.Run("Boom")
            except Exception as exc:
                print(json.dumps({"event": "done", "error": str(exc)}), flush=True)
            else:
                print(json.dumps({"event": "done"}), flush=True)
        elif scenario == "compile":
            _install_proc(app, "Public Sub Bad()\r\nDim x As\r\nEnd Sub\r\n")
            try:
                app.Run("Bad")
            except Exception as exc:
                print(json.dumps({"event": "done", "error": str(exc)}), flush=True)
            else:
                print(json.dumps({"event": "done"}), flush=True)
        elif scenario == "break":
            _install_proc(app, "Public Sub PauseHere()\r\nStop\r\nEnd Sub\r\n")
            app.Run("PauseHere")
            print(json.dumps({"event": "done"}), flush=True)
        else:
            print(json.dumps({"event": "error", "error": f"unknown scenario {scenario}"}), flush=True)
            return 2
        return 0
    except Exception as exc:
        print(json.dumps({"event": "error", "error": str(exc)}), flush=True)
        return 1
    finally:
        _quit_owned(app, identity)
        pythoncom.CoUninitialize()


def _install_proc(app, code: str) -> None:
    component = app.VBE.ActiveVBProject.VBComponents.Add(1)
    component.Name = "modProbe"
    component.CodeModule.AddFromString(code)


@pytest.mark.integration
def test_live_access_dialogs_can_be_inspected_while_blocked():
    """MsgBox, Yes/No, and a second Access instance, using Win32 recovery."""
    reset_interruptions()
    started: list[_Child] = []
    with tempfile.TemporaryDirectory(prefix="vcs-dialog-") as folder:
        db_a = Path(folder) / f"ProbeA-{uuid.uuid4().hex[:8]}.accdb"
        db_b = Path(folder) / f"ProbeB-{uuid.uuid4().hex[:8]}.accdb"
        try:
            child_a = _start_child(db_a, "msgbox_ok")
            started.append(child_a)
            child_b = _start_child(db_b, "idle")
            started.append(child_b)
            ready_a = _wait_ready(child_a)
            ready_b = _wait_ready(child_b)
            pid_a = int(ready_a["pid"])
            pid_b = int(ready_b["pid"])

            listed = _wait_dialog(db_a, pid_a, "ProbeTitle")
            matches = [
                item
                for item in listed.get("dialogs") or []
                if "probetitle" in (item.get("title") or "").lower()
                or "dialog recovery probe" in (item.get("message") or "").lower()
            ]
            assert matches, listed
            dialog = matches[0]
            assert dialog["kind"] == "vba_msgbox"  # custom caption, single OK button
            assert _has_button(dialog["buttons"], "ok")
            assert listed["ready"] is False
            assert listed["pid"] == pid_a

            other = list_dialogs(str(db_b), pid=pid_b, timeout_seconds=2)
            assert all(item["dialog_id"] != dialog["dialog_id"] for item in other.get("dialogs") or [])
            missed = dismiss_dialog(
                str(db_b),
                dialog["dialog_id"],
                button="OK",
                pid=pid_b,
                timeout_seconds=2,
            )
            assert missed["success"] is False
            assert missed["error_pattern"] == "dialog_not_found"

            wrong_stamp = dismiss_dialog(
                str(db_a),
                dialog["dialog_id"],
                button="OK",
                pid=pid_a,
                create_time=1,
                timeout_seconds=2,
            )
            assert wrong_stamp["error_pattern"] == "process_identity_mismatch"

            closed = dismiss_dialog(
                str(db_a),
                dialog["dialog_id"],
                button="OK",
                pid=pid_a,
                create_time=listed.get("create_time"),
                timeout_seconds=5,
            )
            assert closed["success"] is True, closed
            assert closed["dismissed"] is True
            done = _read_json(child_a, 15)
            assert done and done.get("event") == "done", done
            status = automation_status(str(db_a), pid=pid_a, timeout_seconds=2)
            assert status["break_mode"] is False
        finally:
            for child in started:
                _stop(child)


@pytest.mark.integration
def test_live_yes_no_is_not_auto_approved():
    reset_interruptions()
    with tempfile.TemporaryDirectory(prefix="vcs-dialog-") as folder:
        db_path = Path(folder) / f"ProbeYes-{uuid.uuid4().hex[:8]}.accdb"
        proc = _start_child(db_path, "msgbox_yesno")
        pid = 0
        try:
            ready = _wait_ready(proc)
            pid = int(ready["pid"])
            listed = _wait_dialog(db_path, pid, "ConfirmKeep")
            matches = [
                item
                for item in listed.get("dialogs") or []
                if "confirmkeep" in (item.get("title") or "").lower()
            ]
            assert matches, listed
            dialog = matches[0]
            safe = recover_dialogs(str(db_path), policy="safe", pid=pid, timeout_seconds=2)
            assert any(
                item.get("dialog_id") == dialog["dialog_id"]
                for item in (safe.get("skipped") or [])
            ), safe
            assert not any(
                item.get("dialog_id") == dialog["dialog_id"]
                for item in (safe.get("closed") or [])
            ), safe
            closed = dismiss_dialog(
                str(db_path),
                dialog["dialog_id"],
                button="No",
                pid=pid,
                timeout_seconds=5,
            )
            assert closed["success"] is True, closed
            done = _read_json(proc, 15)
            assert done and done.get("event") == "done", done
            assert int(done["result"]) == 7
        finally:
            _stop(proc)


@pytest.mark.integration
def test_live_netui_msgbox2_is_listed_and_dismissed():
    """An @-form MsgBox (as MsgBox2 shows) is a NUIDialog with no Win32 buttons."""
    reset_interruptions()
    with tempfile.TemporaryDirectory(prefix="vcs-dialog-") as folder:
        db_path = Path(folder) / f"ProbeNui-{uuid.uuid4().hex[:8]}.accdb"
        proc = _start_child(db_path, "msgbox2_ok")
        pid = 0
        try:
            ready = _wait_ready(proc)
            pid = int(ready["pid"])
            listed = _wait_dialog(db_path, pid, "NetUIProbe")
            matches = [
                item
                for item in listed.get("dialogs") or []
                if (item.get("title") or "") == "NetUIProbe"
            ]
            assert matches, listed
            dialog = matches[0]
            assert dialog["class_name"] == "NUIDialog"
            assert dialog["kind"] == "vba_msgbox"
            assert "Probe bold" in dialog["message"]
            assert "netui probe line" in dialog["message"]
            assert [_button_label(b) for b in dialog["buttons"]] == ["ok"]
            assert listed["blocking_dialog"] is True
            assert listed["ready"] is False

            closed = dismiss_dialog(
                str(db_path), dialog["dialog_id"], button="OK", pid=pid, timeout_seconds=5
            )
            assert closed["success"] is True, closed
            assert closed["dismissed"] is True
            done = _read_json(proc, 15)
            assert done and done.get("event") == "done", done
        finally:
            _stop(proc)


@pytest.mark.integration
def test_live_netui_access_error_is_reported_not_clicked_by_safe():
    reset_interruptions()
    with tempfile.TemporaryDirectory(prefix="vcs-dialog-") as folder:
        db_path = Path(folder) / f"ProbeNuiErr-{uuid.uuid4().hex[:8]}.accdb"
        proc = _start_child(db_path, "access_error")
        pid = 0
        try:
            ready = _wait_ready(proc)
            pid = int(ready["pid"])
            listed = _wait_dialog(db_path, pid, "NoSuchProbeForm")
            matches = [
                item
                for item in listed.get("dialogs") or []
                if "nosuchprobeform" in (item.get("message") or "").lower()
            ]
            assert matches, listed
            dialog = matches[0]
            assert dialog["class_name"] == "NUIDialog"
            assert dialog["kind"] == "access_dialog"
            safe = recover_dialogs(str(db_path), policy="safe", pid=pid, timeout_seconds=2)
            assert not any(
                item.get("dialog_id") == dialog["dialog_id"] for item in safe.get("closed") or []
            ), safe
            closed = dismiss_dialog(
                str(db_path), dialog["dialog_id"], button="OK", pid=pid, timeout_seconds=5
            )
            assert closed["success"] is True, closed
            done = _read_json(proc, 15)
            assert done and done.get("event") == "done", done
        finally:
            _stop(proc)


@pytest.mark.integration
def test_live_runtime_compile_and_break():
    """End/Debug, compile error, and break mode. Each phase records what Access showed."""
    reset_interruptions()
    notes: list[str] = []
    with tempfile.TemporaryDirectory(prefix="vcs-dialog-") as folder:
        _exercise_runtime(Path(folder), notes)
        _exercise_compile(Path(folder), notes)
        _exercise_break(Path(folder), notes)
    missing = [note for note in notes if "not shown" in note]
    assert not missing, notes


def _exercise_runtime(folder: Path, notes: list[str]) -> None:
    db_path = folder / f"ProbeRun-{uuid.uuid4().hex[:8]}.accdb"
    proc = _start_child(db_path, "runtime")
    pid = 0
    try:
        ready = _wait_ready(proc)
        pid = int(ready["pid"])
        listed = _wait_dialog(db_path, pid, "Microsoft Visual Basic", timeout=15)
        dialogs = listed.get("dialogs") or []
        runtime = [item for item in dialogs if item.get("kind") == "vba_runtime_error"]
        if not runtime:
            early = _read_json(proc, 1)
            notes.append(f"runtime dialog not shown; child={early!r} listed={listed!r}")
            return
        dialog = runtime[0]
        assert _has_button(dialog["buttons"], "end")
        assert _has_button(dialog["buttons"], "debug")
        refused = dismiss_dialog(
            str(db_path), dialog["dialog_id"], button="Debug", pid=pid, timeout_seconds=2
        )
        assert refused["error_pattern"] == "debug_refused"
        safe = recover_dialogs(str(db_path), policy="safe", pid=pid, timeout_seconds=2)
        assert not any(_button_label(item.get("button") or "") == "end" for item in safe.get("closed") or [])
        ended = recover_dialogs(
            str(db_path), policy="end_runtime_error", pid=pid, timeout_seconds=5
        )
        assert any(
            _button_label(item.get("button") or "") == "end" for item in ended.get("closed") or []
        ), ended
        status = automation_status(str(db_path), pid=pid, timeout_seconds=2)
        assert status.get("execution_interrupted") is True
        assert status.get("failure_dialog_dismissed") or status.get("last_interruption")
        notes.append("runtime End/Debug dialog ended without Debug")
    finally:
        _stop(proc)


def _exercise_compile(folder: Path, notes: list[str]) -> None:
    db_path = folder / f"ProbeCompile-{uuid.uuid4().hex[:8]}.accdb"
    proc = _start_child(db_path, "compile")
    pid = 0
    try:
        ready = _wait_ready(proc)
        pid = int(ready["pid"])
        listed = _wait_dialog(db_path, pid, "compile", timeout=15)
        compile_dialogs = [
            item
            for item in listed.get("dialogs") or []
            if item.get("kind") == "vba_compile_error"
            or "compile error" in (item.get("message") or "").lower()
        ]
        if not compile_dialogs:
            early = _read_json(proc, 1)
            notes.append(f"compile dialog not shown; child={early!r} listed={listed!r}")
            return
        dialog = compile_dialogs[0]
        closed = dismiss_dialog(
            str(db_path), dialog["dialog_id"], button="OK", pid=pid, timeout_seconds=5
        )
        assert closed.get("failure_dialog_dismissed") is True or closed.get("interrupted") is True
        notes.append("compile error dialog dismissed and recorded as a failure")
    finally:
        _stop(proc)


def _exercise_break(folder: Path, notes: list[str]) -> None:
    db_path = folder / f"ProbeBreak-{uuid.uuid4().hex[:8]}.accdb"
    proc = _start_child(db_path, "break")
    pid = 0
    try:
        ready = _wait_ready(proc)
        pid = int(ready["pid"])
        deadline = time.monotonic() + 15
        listed: dict = {}
        while time.monotonic() < deadline:
            listed = list_dialogs(str(db_path), pid=pid, timeout_seconds=2)
            if listed.get("break_mode"):
                break
            time.sleep(0.25)
        if not listed.get("break_mode"):
            early = _read_json(proc, 1)
            notes.append(f"break mode not shown; child={early!r} listed={listed!r}")
            return
        assert any(item.get("kind") == "vba_break" for item in listed.get("dialogs") or [])
        assert all(item.get("is_dialog") is not True or item.get("kind") != "vba_break" for item in listed["dialogs"])
        refused = dismiss_dialog(
            str(db_path),
            next(item["dialog_id"] for item in listed["dialogs"] if item["kind"] == "vba_break"),
            action="close",
            pid=pid,
            timeout_seconds=2,
        )
        assert refused["error_pattern"] == "vba_break"
        notes.append("break mode reported and not closed as a dialog")
    finally:
        _stop(proc)


if __name__ == "__main__" and "--child" in sys.argv:
    index = sys.argv.index("--child")
    raise SystemExit(_child_main(sys.argv[index + 1], sys.argv[index + 2]))
