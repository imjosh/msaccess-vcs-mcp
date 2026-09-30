"""Exercise live-test cleanup with fake processes; never start real Access."""

from __future__ import annotations

import io
import json
import subprocess
import sys
from types import SimpleNamespace

import pytest

from tests import test_dialog_live as live


STAMP = 10_000_000


class ProcessTable:
    def __init__(self, monkeypatch):
        self.processes = {10: STAMP, 20: STAMP, 99: STAMP}
        self.actions = []
        self.handles = []
        self.waited = []
        self.query_failed = False
        self.stamp_failed = False
        monkeypatch.setitem(sys.modules, "win32con", SimpleNamespace(
            PROCESS_QUERY_LIMITED_INFORMATION=0x1000, PROCESS_TERMINATE=1, SYNCHRONIZE=0x100000,
        ))
        monkeypatch.setitem(sys.modules, "win32api", SimpleNamespace(
            OpenProcess=self.open, CloseHandle=self.close, TerminateProcess=self.terminate,
        ))
        monkeypatch.setitem(sys.modules, "win32process", SimpleNamespace(
            GetProcessTimes=self.times, GetWindowThreadProcessId=lambda hwnd: (1, hwnd),
        ))
        monkeypatch.setitem(sys.modules, "win32event", SimpleNamespace(
            WaitForSingleObject=self.wait,
        ))

    def open(self, access, inherit, pid):
        if self.query_failed or pid not in self.processes:
            raise OSError("process query failed")
        handle = SimpleNamespace(pid=pid, stamp=self.processes[pid], access=access)
        self.handles.append(handle)
        return handle

    def times(self, handle):
        if self.stamp_failed or handle.stamp is None:
            raise OSError("creation time unreadable")
        return {"CreationTime": SimpleNamespace(timestamp=lambda: handle.stamp / 10_000_000)}

    def close(self, handle):
        self.handles.remove(handle)

    def terminate(self, handle, exit_code):
        assert handle.access & 1
        self.actions.append(("terminate", handle.pid, handle.stamp))
        if self.processes.get(handle.pid) == handle.stamp:
            self.processes.pop(handle.pid, None)

    def wait(self, handle, timeout):
        assert handle in self.handles
        assert handle.access & 0x100000
        self.waited.append((handle.pid, handle.stamp))
        return 0


class ChildProcess:
    def __init__(self, table, pid, *, hung=False, on_wait=None):
        self.table = table
        self.pid = pid
        self.stdin = io.StringIO()
        self.hung = hung
        self.exited = False
        self.on_wait = on_wait

    def wait(self, timeout):
        if self.exited:
            return 0
        if self.on_wait is not None:
            self.on_wait()
            self.on_wait = None
        if self.hung and not self.exited:
            raise subprocess.TimeoutExpired("fake child", timeout)
        self.exited = True
        if self.stdin is not None and self.stdin.getvalue() == "quit\n":
            self.table.actions.append(("quit", self.pid))
            self.table.processes.pop(self.pid, None)
        return 0

    def poll(self):
        return 0 if self.exited else None

    def kill(self):
        self.table.actions.append(("kill_helper", self.pid))
        self.exited = True


def child(table, pid=10, *, claimed=True, hung=False, on_wait=None, events=()):
    result = live._Child(ChildProcess(table, pid, hung=hung, on_wait=on_wait))
    if claimed:
        result.identity = live._AccessIdentity(pid, STAMP)
    for event in events:
        result.lines.put(json.dumps(event))
    # EOF makes missing-handshake tests deterministic without a real wait.
    result.lines.put("")
    return result


def launch(pid=10, stamp=STAMP):
    return {"event": "launched", "pid": pid, "create_time": stamp}


@pytest.mark.parametrize("hung", [False, True])
def test_teardown_leaves_concurrently_launched_unclaimed_access(monkeypatch, hung):
    table = ProcessTable(monkeypatch)
    owned = child(table, hung=hung)
    live._stop(owned)
    assert table.actions == (
        [("terminate", 10, STAMP), ("kill_helper", 10)] if hung else [("quit", 10)]
    )
    assert table.processes == {20: STAMP, 99: STAMP}
    assert not table.handles


@pytest.mark.parametrize("reason", ["reused", "unreadable", "query_failed", "unclaimed"])
def test_unknown_or_mismatched_identity_never_closes_or_terminates(monkeypatch, reason):
    table = ProcessTable(monkeypatch)
    owned = child(table, claimed=reason != "unclaimed", hung=True)
    if reason == "reused":
        table.processes[10] = STAMP * 2
    elif reason == "unreadable":
        table.stamp_failed = True
    elif reason == "query_failed":
        table.query_failed = True
    before = dict(table.processes)
    live._stop(owned)
    assert owned.proc.stdin.getvalue() == ""
    assert table.actions == [("kill_helper", 10)]
    assert table.processes == before
    assert not table.handles


def test_force_termination_rechecks_identity_after_graceful_wait(monkeypatch):
    table = ProcessTable(monkeypatch)
    owned = child(table, hung=True, on_wait=lambda: table.processes.update({10: STAMP * 2}))
    live._stop(owned)
    assert owned.proc.stdin.getvalue() == "quit\n"
    assert table.actions == [("kill_helper", 10)]
    assert table.processes[10] == STAMP * 2


def test_force_termination_uses_the_same_verified_handle(monkeypatch):
    table = ProcessTable(monkeypatch)
    original = table.times

    def reuse_pid_after_query(handle):
        result = original(handle)
        table.processes[10] = STAMP * 2
        return result

    monkeypatch.setattr(sys.modules["win32process"], "GetProcessTimes", reuse_pid_after_query)
    live._kill(live._AccessIdentity(10, STAMP))
    assert table.actions == [("terminate", 10, STAMP)]
    assert table.processes[10] == STAMP * 2
    assert table.waited == [(10, STAMP)]
    assert not table.handles


@pytest.mark.parametrize("failed_child", [10, 20])
def test_partial_startup_cleans_all_confirmed_children_without_a_pid_sweep(monkeypatch, failed_child):
    table = ProcessTable(monkeypatch)
    children = [child(table, pid, claimed=False, hung=True, events=[
        launch(pid),
        {"event": "error", "error": "startup failed"} if pid == failed_child
        else {"event": "ready", "pid": pid},
    ]) for pid in (10, 20)]
    starters = iter(children)
    monkeypatch.setattr(live, "_start_child", lambda *args: next(starters))
    with pytest.raises(AssertionError, match="did not become ready"):
        live.test_live_access_dialogs_can_be_inspected_while_blocked()
    assert table.actions == [
        ("terminate", 10, STAMP), ("kill_helper", 10),
        ("terminate", 20, STAMP), ("kill_helper", 20),
    ]
    assert table.processes == {99: STAMP}


def test_second_launch_failure_still_cleans_first_child(monkeypatch):
    table = ProcessTable(monkeypatch)
    first = child(table, claimed=False, hung=True, events=[launch()])
    calls = 0

    def start(*args):
        nonlocal calls
        calls += 1
        if calls == 2:
            raise OSError("Popen failed")
        return first

    monkeypatch.setattr(live, "_start_child", start)
    with pytest.raises(OSError, match="Popen failed"):
        live.test_live_access_dialogs_can_be_inspected_while_blocked()
    assert table.actions == [("terminate", 10, STAMP), ("kill_helper", 10)]
    assert table.processes == {20: STAMP, 99: STAMP}


@pytest.mark.parametrize("event", [
    {"event": "ready", "pid": 10}, launch(stamp=None), launch(stamp=0),
    launch(stamp=True), launch(pid=0), launch(pid=True),
])
def test_missing_or_invalid_launch_claim_is_never_cleanup_authority(monkeypatch, event):
    table = ProcessTable(monkeypatch)
    owned = child(table, claimed=False, hung=True, events=[event])
    live._stop(owned)
    assert owned.identity is None
    assert owned.proc.stdin.getvalue() == ""
    assert table.actions == [("kill_helper", 10)]
    assert table.processes[10] == STAMP


@pytest.mark.parametrize("stamp,app_pid", [(STAMP, 10), (STAMP * 2, 10), (STAMP, 99), (None, 10)])
def test_child_finally_rechecks_its_com_process_before_quit(monkeypatch, stamp, app_pid):
    table = ProcessTable(monkeypatch)
    table.processes[10] = stamp
    app = SimpleNamespace(hWndAccessApp=app_pid, Quit=lambda mode: table.actions.append(("quit", app_pid)))
    live._quit_owned(app, live._AccessIdentity(10, STAMP))
    assert table.actions == ([("quit", 10)] if stamp == STAMP and app_pid == 10 else [])


def test_child_reports_ownership_before_database_startup_failure(monkeypatch, capsys, tmp_path):
    table = ProcessTable(monkeypatch)
    app = SimpleNamespace(hWndAccessApp=10, Quit=lambda mode: table.actions.append(("quit", 10)))

    def fail_open(path):
        raise OSError("database startup failed")

    app.NewCurrentDatabase = fail_open
    monkeypatch.setitem(sys.modules, "pythoncom", SimpleNamespace(
        CoInitialize=lambda: None, CoUninitialize=lambda: None,
    ))
    client = SimpleNamespace(DispatchEx=lambda name: app)
    monkeypatch.setitem(sys.modules, "win32com", SimpleNamespace(client=client))
    monkeypatch.setitem(sys.modules, "win32com.client", client)
    assert live._child_main(str(tmp_path / "probe.accdb"), "idle") == 1
    events = [json.loads(line) for line in capsys.readouterr().out.splitlines()]
    assert events[0] == launch()
    assert events[1] == {"event": "error", "error": "database startup failed"}
    assert table.actions == [("quit", 10)]


def test_exited_helper_does_not_hide_a_confirmed_access_leak(monkeypatch):
    table = ProcessTable(monkeypatch)
    owned = child(table)
    owned.proc.stdin = None
    live._stop(owned)
    assert table.actions == [("terminate", 10, STAMP)]
    assert table.processes == {20: STAMP, 99: STAMP}


def test_stdout_reader_records_launch_before_readiness(monkeypatch, tmp_path):
    table = ProcessTable(monkeypatch)
    proc = ChildProcess(table, 10)
    ready = {"event": "ready", "pid": 10, "hwnd": 10}
    proc.stdout = io.StringIO(json.dumps(launch()) + "\n" + json.dumps(ready) + "\n")
    monkeypatch.setattr(live.subprocess, "Popen", lambda *args, **kwargs: proc)
    owned = live._start_child(tmp_path / "probe.accdb", "idle")
    assert live._wait_ready(owned, timeout=1) == ready
    assert owned.identity == live._AccessIdentity(10, STAMP)
    live._stop(owned)
    assert table.actions == [("quit", 10)]


def test_timeout_does_not_lose_a_later_launch_handshake(monkeypatch):
    table = ProcessTable(monkeypatch)
    owned = live._Child(ChildProcess(table, 10, hung=True))
    with pytest.raises(AssertionError, match="did not become ready"):
        live._wait_ready(owned, timeout=0)
    owned.lines.put(json.dumps(launch()))
    live._stop(owned)
    assert table.actions == [("terminate", 10, STAMP), ("kill_helper", 10)]
