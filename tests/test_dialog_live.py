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


def _access_pids() -> set[int]:
    import win32api
    import win32con
    import win32process

    found: set[int] = set()
    for pid in win32process.EnumProcesses():
        if not pid:
            continue
        try:
            handle = win32api.OpenProcess(
                win32con.PROCESS_QUERY_LIMITED_INFORMATION, False, pid
            )
        except Exception:
            continue
        try:
            path = win32process.GetModuleFileNameEx(handle, 0)
        except Exception:
            path = ""
        finally:
            win32api.CloseHandle(handle)
        if os.path.basename(path).lower() == "msaccess.exe":
            found.add(int(pid))
    return found


def _kill(pid: int) -> None:
    if pid <= 0:
        return
    subprocess.run(
        ["taskkill", "/PID", str(pid), "/F"],
        capture_output=True,
        check=False,
    )


def _read_json(proc: subprocess.Popen[str], timeout: float) -> dict | None:
    holder: queue.Queue[str] = queue.Queue()

    def _reader() -> None:
        line = proc.stdout.readline() if proc.stdout is not None else ""
        holder.put(line)

    thread = threading.Thread(target=_reader, daemon=True)
    thread.start()
    thread.join(timeout)
    if thread.is_alive() or holder.empty():
        return None
    line = holder.get()
    if not line:
        return None
    return json.loads(line)


def _start_child(db_path: Path, scenario: str) -> subprocess.Popen[str]:
    return subprocess.Popen(
        [sys.executable, str(Path(__file__).resolve()), "--child", str(db_path), scenario],
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )


def _wait_ready(proc: subprocess.Popen[str], before: set[int], timeout: float = 45) -> dict:
    payload = _read_json(proc, timeout)
    if payload and payload.get("event") == "ready":
        return payload
    born = _access_pids() - before
    for pid in born:
        _kill(pid)
    proc.kill()
    err = ""
    if proc.stderr is not None:
        err = proc.stderr.read()
    raise AssertionError(
        f"Access child did not become ready ({payload!r}). stderr={err[-2000:]}"
    )


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


def _stop(proc: subprocess.Popen[str], pid: int) -> None:
    if proc.stdin is not None:
        try:
            proc.stdin.write("quit\n")
            proc.stdin.flush()
        except Exception:
            pass
    try:
        proc.wait(timeout=8)
    except subprocess.TimeoutExpired:
        _kill(pid)
        proc.kill()


def _child_main(db_path: str, scenario: str) -> int:
    import pythoncom
    import win32com.client
    import win32process

    pythoncom.CoInitialize()
    app = win32com.client.DispatchEx("Access.Application")
    try:
        app.Visible = True
        if os.path.exists(db_path):
            os.remove(db_path)
        app.NewCurrentDatabase(db_path)
        hwnd_value = app.hWndAccessApp
        if callable(hwnd_value):
            hwnd_value = hwnd_value()
        hwnd = int(hwnd_value)
        _thread, pid = win32process.GetWindowThreadProcessId(hwnd)
        print(json.dumps({"event": "ready", "pid": int(pid), "hwnd": hwnd}), flush=True)
        if scenario == "idle":
            sys.stdin.readline()
        elif scenario == "msgbox_ok":
            app.Eval('MsgBox("dialog recovery probe",0,"ProbeTitle")')
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
        try:
            app.Quit(2)
        except Exception:
            pass


def _install_proc(app, code: str) -> None:
    component = app.VBE.ActiveVBProject.VBComponents.Add(1)
    component.Name = "modProbe"
    component.CodeModule.AddFromString(code)


@pytest.mark.integration
def test_live_access_dialogs_can_be_inspected_while_blocked():
    """MsgBox, Yes/No, and a second Access instance, using Win32 recovery."""
    reset_interruptions()
    before = _access_pids()
    started: list[tuple[subprocess.Popen[str], int]] = []
    with tempfile.TemporaryDirectory(prefix="vcs-dialog-") as folder:
        db_a = Path(folder) / f"ProbeA-{uuid.uuid4().hex[:8]}.accdb"
        db_b = Path(folder) / f"ProbeB-{uuid.uuid4().hex[:8]}.accdb"
        child_a = _start_child(db_a, "msgbox_ok")
        child_b = _start_child(db_b, "idle")
        try:
            ready_a = _wait_ready(child_a, before)
            ready_b = _wait_ready(child_b, before | {int(ready_a["pid"])})
            pid_a = int(ready_a["pid"])
            pid_b = int(ready_b["pid"])
            started.extend([(child_a, pid_a), (child_b, pid_b)])

            listed = _wait_dialog(db_a, pid_a, "ProbeTitle")
            matches = [
                item
                for item in listed.get("dialogs") or []
                if "probetitle" in (item.get("title") or "").lower()
                or "dialog recovery probe" in (item.get("message") or "").lower()
            ]
            assert matches, listed
            dialog = matches[0]
            assert dialog["kind"] == "vba_msgbox"
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
            for proc, pid in started:
                _stop(proc, pid)
            leaked = _access_pids() - before
            for pid in leaked:
                _kill(pid)


@pytest.mark.integration
def test_live_yes_no_is_not_auto_approved():
    reset_interruptions()
    before = _access_pids()
    with tempfile.TemporaryDirectory(prefix="vcs-dialog-") as folder:
        db_path = Path(folder) / f"ProbeYes-{uuid.uuid4().hex[:8]}.accdb"
        proc = _start_child(db_path, "msgbox_yesno")
        pid = 0
        try:
            ready = _wait_ready(proc, before)
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
            _stop(proc, pid)
            for extra in _access_pids() - before:
                _kill(extra)


@pytest.mark.integration
def test_live_runtime_compile_and_break():
    """End/Debug, compile error, and break mode. Each phase records what Access showed."""
    reset_interruptions()
    before = _access_pids()
    notes: list[str] = []
    with tempfile.TemporaryDirectory(prefix="vcs-dialog-") as folder:
        _exercise_runtime(Path(folder), before, notes)
        _exercise_compile(Path(folder), before, notes)
        _exercise_break(Path(folder), before, notes)
    missing = [note for note in notes if "not shown" in note]
    assert not missing, notes


def _exercise_runtime(folder: Path, before: set[int], notes: list[str]) -> None:
    db_path = folder / f"ProbeRun-{uuid.uuid4().hex[:8]}.accdb"
    proc = _start_child(db_path, "runtime")
    pid = 0
    try:
        ready = _wait_ready(proc, before)
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
        _stop(proc, pid)
        for extra in _access_pids() - before:
            _kill(extra)


def _exercise_compile(folder: Path, before: set[int], notes: list[str]) -> None:
    db_path = folder / f"ProbeCompile-{uuid.uuid4().hex[:8]}.accdb"
    proc = _start_child(db_path, "compile")
    pid = 0
    try:
        ready = _wait_ready(proc, before)
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
        _stop(proc, pid)
        for extra in _access_pids() - before:
            _kill(extra)


def _exercise_break(folder: Path, before: set[int], notes: list[str]) -> None:
    db_path = folder / f"ProbeBreak-{uuid.uuid4().hex[:8]}.accdb"
    proc = _start_child(db_path, "break")
    pid = 0
    try:
        ready = _wait_ready(proc, before)
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
        _kill(pid)
        proc.kill()
        for extra in _access_pids() - before:
            _kill(extra)


if __name__ == "__main__" and "--child" in sys.argv:
    index = sys.argv.index("--child")
    raise SystemExit(_child_main(sys.argv[index + 1], sys.argv[index + 2]))
