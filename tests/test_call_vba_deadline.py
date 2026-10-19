"""vcs_call_vba's one deadline, through the public tool and the real gate (MCP6).

No COM: the connection, add-in load, session check and target are fakes that
block on events, so each test can hold the call at a chosen step.
"""

from __future__ import annotations

import asyncio
import random
import subprocess
import sys
import threading
import time
from pathlib import Path
from unittest.mock import AsyncMock

import pytest

from msaccess_vcs_mcp import tools
from msaccess_vcs_mcp.access_gate import (
    Attempt,
    Phase,
    get_access_gate,
    reset_access_gate,
)
from msaccess_vcs_mcp.dialog_recovery import _gate_snapshot

TAG = "msaccess-vcs.call-vba-result/1"
ADDIN = r"C:\library\Version Control.accda"
SETTLE = 5.0


class Step:
    def __init__(self):
        self.entered = threading.Event()
        self.release = threading.Event()
        self.release.set()
        self.threads: list[int] = []

    def hold(self):
        self.release.clear()
        return self

    def __call__(self, *_args, **_kwargs):
        self.threads.append(threading.get_ident())
        self.entered.set()
        assert self.release.wait(SETTLE), "test never released this step"


class FakeApp:
    def __init__(self, step):
        self.step = step
        self.runs: list[tuple] = []

    def Run(self, procedure, *args):
        self.runs.append((procedure, args))
        self.step()
        return (TAG, "returned", "ok")


class Boundary:
    def __init__(self, monkeypatch, db):
        self.db = db
        self.steps = {name: Step() for name in ("connect", "load_addin", "ensure_session", "run")}
        self.app = FakeApp(self.steps["run"])
        boundary = self

        class Connection:
            def __init__(self, *_args, **_kwargs):
                pass

            def __enter__(self):
                return self

            def __exit__(self, *_exc):
                return False

            def connect(self):
                boundary.steps["connect"]()
                return boundary.app, object()

        class Addin:
            addin_path = ADDIN

            def __init__(self, *_args):
                pass

            def load_addin(self, *_args, **_kwargs):
                boundary.steps["load_addin"]()

        def session(*_args):
            boundary.steps["ensure_session"]()
            return "accepted-session"

        monkeypatch.setattr(tools, "AccessConnection", Connection)
        monkeypatch.setattr(tools, "VCSAddinIntegration", Addin)
        monkeypatch.setattr(tools, "ensure_session", session)

    def call(self, timeout=0.2, function="Mod.Echo", args=("hi",)):
        return tools.vcs_call_vba(self.db, function, list(args), timeout_seconds=timeout)

    def release_all(self):
        for step in self.steps.values():
            step.release.set()


@pytest.fixture
def boundary(monkeypatch, tmp_path):
    db = tmp_path / "user.accdb"
    db.write_bytes(b"")
    reset_access_gate()
    monkeypatch.setattr(tools, "_ensure_env_loaded", AsyncMock())
    monkeypatch.setattr(tools, "load_config", lambda: {})
    monkeypatch.setattr(tools, "get_config", lambda: {"ACCESS_VCS_ADDIN_PATH": ADDIN})
    monkeypatch.setattr(tools, "_check_database_busy", lambda _path: None)
    monkeypatch.setenv("ACCESS_VCS_ENABLE_LOGGING", "false")
    monkeypatch.setenv("ACCESS_VCS_DISABLE_DIAGNOSTIC_LOG", "true")
    monkeypatch.setenv("ACCESS_VCS_BUSY_WAIT_SEC", "0.05")
    fake = Boundary(monkeypatch, str(db))
    try:
        yield fake
    finally:
        fake.release_all()
        gate = get_access_gate()
        gate._executor.shutdown(wait=True)
        reset_access_gate()


async def _gate_idle(limit=SETTLE):
    gate = get_access_gate()
    deadline = time.monotonic() + limit
    while gate.current_in_flight() is not None:
        assert time.monotonic() < deadline, "the gate never settled"
        await asyncio.sleep(0.005)


async def _entered(step, limit=SETTLE):
    deadline = time.monotonic() + limit
    while not step.entered.is_set():
        assert time.monotonic() < deadline, "the call never reached this step"
        await asyncio.sleep(0.005)


def _other_gated_call(db):
    return tools.vcs_call_vba(db, "Mod.Other", [], timeout_seconds=1.0)


@pytest.mark.parametrize("blocked", ["connect", "load_addin", "ensure_session"])
def test_blocked_validation_answers_by_the_deadline_and_never_dispatches(boundary, blocked):
    """D1 and D2: every pre-dispatch step counts, and late validation never reaches the target."""
    boundary.steps[blocked].hold()

    async def scenario():
        started = time.monotonic()
        reply = await asyncio.wait_for(boundary.call(timeout=0.5), 2.0)
        assert time.monotonic() - started < 1.5
        assert boundary.steps[blocked].entered.is_set()
        assert reply["error_pattern"] == "timeout"
        assert reply["timed_out"] is True
        assert reply["phase"] == "validating"
        assert reply["dispatch"] == "not_dispatched"
        assert reply["retry_safe"] is True
        assert reply["gate_retained"] is True
        assert reply["function"] == "Mod.Echo"

        busy = await _other_gated_call(boundary.db)
        assert busy["error_pattern"] == "server_busy"
        assert busy["busy_with"]["call_id"] == reply["call_id"]
        assert busy["busy_with"]["phase"] == "validating"
        assert busy["busy_with"]["caller"] == "expired"
        assert busy["busy_with"]["dispatch"] == "not_dispatched"
        status = _gate_snapshot(boundary.db)["operation"]
        assert status["call_id"] == reply["call_id"]
        assert status["phase"] == "validating"

        boundary.steps[blocked].release.set()
        await _gate_idle()
        assert boundary.app.runs == []
        if blocked != "ensure_session":
            # The session handshake acts on Access; an expired call never starts it.
            assert boundary.steps["ensure_session"].threads == []

        follow = await boundary.call(timeout=2.0)
        assert follow["success"] is True
        assert len(boundary.app.runs) == 1

    asyncio.run(scenario())


def test_target_running_past_the_deadline_holds_the_gate(boundary):
    """D3 and D4: the target runs to settlement on the apartment and the slot stays held."""

    async def scenario():
        assert (await boundary.call(timeout=2.0))["success"] is True  # warm the apartment
        apartment = set(boundary.steps["run"].threads)
        boundary.steps["run"].hold()
        boundary.steps["run"].entered.clear()
        reply = await asyncio.wait_for(boundary.call(timeout=0.5), 2.0)
        assert boundary.steps["run"].entered.is_set()
        assert reply["phase"] == "target_possibly_running"
        assert reply["dispatch"] == "possibly_dispatched"
        assert reply["retry_safe"] is False
        assert reply["gate_retained"] is True
        # Run is still blocked on the same apartment thread, not a fresh one.
        assert set(boundary.steps["run"].threads) == apartment

        busy = await _other_gated_call(boundary.db)
        assert busy["error_pattern"] == "server_busy"
        assert busy["busy_with"]["call_id"] == reply["call_id"]
        assert busy["busy_with"]["dispatch"] == "possibly_dispatched"
        assert "Do not re-issue" in busy["error"]
        assert busy["busy_with"]["overdue_ms"] >= 0

        boundary.steps["run"].release.set()
        await _gate_idle()
        assert len(boundary.app.runs) == 2

    asyncio.run(scenario())


def test_target_runs_on_the_thread_that_validated_the_session(boundary):
    async def scenario():
        assert (await boundary.call(timeout=2.0))["success"] is True

    asyncio.run(scenario())
    connect, run = boundary.steps["connect"].threads, boundary.steps["run"].threads
    assert connect == run and len(run) == 1
    assert connect[0] != threading.get_ident()


def test_slot_wait_counts_toward_the_deadline(boundary, monkeypatch):
    monkeypatch.setenv("ACCESS_VCS_BUSY_WAIT_SEC", "5")
    gate = get_access_gate()
    hold = threading.Event()

    async def scenario():
        occupant = asyncio.create_task(gate.run_exclusive("vcs_occupant", None, hold.wait, False, 5))
        while gate.current_in_flight() is None:
            await asyncio.sleep(0.005)
        started = time.monotonic()
        reply = await boundary.call(timeout=0.2)
        assert time.monotonic() - started < 1.0
        assert reply["phase"] == "queued"
        assert reply["dispatch"] == "not_dispatched"
        assert reply["gate_retained"] is False
        assert reply["call_id"] is None
        hold.set()
        await occupant
        assert boundary.app.runs == []

    asyncio.run(scenario())


def test_result_at_the_boundary_is_returned(boundary):
    async def scenario():
        reply = await boundary.call(timeout=2.0)
        assert reply["success"] is True
        assert reply["result"] == "ok"
        assert "timed_out" not in reply

    asyncio.run(scenario())


def test_cancelled_caller_never_dispatches_and_keeps_the_slot(boundary):
    boundary.steps["ensure_session"].hold()

    async def scenario():
        task = asyncio.create_task(boundary.call(timeout=5.0))
        await _entered(boundary.steps["ensure_session"])
        task.cancel()
        with pytest.raises(asyncio.CancelledError):
            await task
        assert get_access_gate().current_in_flight() is not None
        boundary.steps["ensure_session"].release.set()
        await _gate_idle()
        assert boundary.app.runs == []

    asyncio.run(scenario())


def test_expiry_after_the_permit_cannot_revoke_dispatch(boundary):
    """The rebuild launch path: the callback runs after the permit, then the target once."""
    permitted = Step().hold()

    def on_admitted(call_args):
        permitted()
        return call_args

    async def scenario():
        gate = get_access_gate()
        reply = await gate.run_exclusive(
            "vcs_rebuild_addin", boundary.db,
            lambda: tools._execute_call_vba(boundary.db, "VCS.API", ["GetVCSVersion"], on_admitted=on_admitted),
            False, deadline_seconds=0.2,
        )
        assert reply["phase"] == "dispatch_permitted"
        assert reply["dispatch"] == "possibly_dispatched"
        assert reply["retry_safe"] is False
        permitted.release.set()
        await _gate_idle()
        assert len(boundary.app.runs) == 1
        assert boundary.app.runs[0][0].endswith(".APIExecute")

    asyncio.run(scenario())


def test_permit_and_expiry_have_one_winner():
    """Whichever of permit and caller expiry takes the lock first decides; replies agree."""
    now = [0.0]
    refused = Attempt(1, "vcs_call_vba", None, 10.0, 0.0, clock=lambda: now[0])
    refused.begin_validation()
    assert refused.caller_expired("deadline")["dispatch"] == "not_dispatched"
    assert refused.permit_dispatch() is None

    granted = Attempt(2, "vcs_call_vba", None, 10.0, 0.0, clock=lambda: now[0])
    granted.begin_validation()
    assert granted.permit_dispatch() is not None
    view = granted.caller_expired("deadline")
    assert view["phase"] == "dispatch_permitted"
    assert view["dispatch"] == "possibly_dispatched"

    late = Attempt(3, "vcs_call_vba", None, 10.0, 0.0, clock=lambda: now[0])
    late.begin_validation()
    now[0] = 10.0  # The loop has not noticed yet; the clock alone refuses.
    assert late.permit_dispatch() is None
    assert late.snapshot()["caller"] == "expired"

    rng = random.Random(6)
    for index in range(300):
        attempt = Attempt(index, "vcs_call_vba", None, 10.0, 0.0)
        attempt.begin_validation()
        start = threading.Barrier(2)
        outcome = {}

        def permit():
            start.wait()
            time.sleep(rng.random() / 10000)
            outcome["permit"] = attempt.permit_dispatch()

        def expire():
            start.wait()
            time.sleep(rng.random() / 10000)
            outcome["view"] = attempt.caller_expired("deadline")

        workers = [threading.Thread(target=permit), threading.Thread(target=expire)]
        for worker in workers:
            worker.start()
        for worker in workers:
            worker.join()
        expected = "possibly_dispatched" if outcome["permit"] is not None else "not_dispatched"
        assert outcome["view"]["dispatch"] == expected


def test_phase_only_moves_forward():
    attempt = Attempt(1, "vcs_call_vba", None)
    attempt.begin_validation()
    with pytest.raises(RuntimeError):
        attempt.begin_validation()
    permit = attempt.permit_dispatch()
    with pytest.raises(RuntimeError):
        attempt.permit_dispatch()
    app = FakeApp(Step())
    permit.run(app, "P")
    with pytest.raises(RuntimeError):
        permit.run(app, "P")
    attempt.settle()
    assert attempt.snapshot()["phase"] == "settled"
    assert attempt.snapshot()["dispatch"] == "dispatched"
    assert Phase.SETTLED > Phase.TARGET_POSSIBLY_RUNNING


def test_calls_without_the_permit_path_report_no_dispatch():
    """A retained build never claims it was not dispatched."""
    build = Attempt(1, "vcs_rebuild_database", None)
    assert "dispatch" not in build.snapshot()
    build.settle()
    assert "dispatch" not in build.snapshot()


CHILD = Path(__file__).with_name("call_vba_exit_child.py")


@pytest.mark.parametrize("mode", ["deadline", "no-deadline"])
def test_interpreter_exit_waits_only_as_long_as_the_caller_would(tmp_path, mode):
    marker = tmp_path / "marker.txt"
    started = time.monotonic()
    completed = subprocess.run(
        [sys.executable, str(CHILD), str(marker), mode], capture_output=True, text=True, timeout=30,
    )
    elapsed = time.monotonic() - started
    notes = marker.read_text(encoding="utf-8").splitlines()
    assert completed.returncode == 0, completed.stderr
    if mode == "deadline":
        # A target that never returns: exit waits for its deadline and grace, then records it.
        assert notes[0].startswith("returned timeout")
        assert any(line.startswith("abandoned ") and "possibly_dispatched" in line for line in notes)
        assert "body-finished" not in notes
        assert elapsed < 20
    else:
        # No deadline: exit lets the body finish, as the executor join did.
        assert notes == ["caller-left", "body-finished", "atexit-ran"]


def test_the_target_has_one_run_site():
    """Application.Run on the target is reachable only through DispatchPermit.run."""
    import inspect

    from msaccess_vcs_mcp import access_gate

    for function in (tools._execute_call_vba, tools._run_target):
        assert ".Run(" not in inspect.getsource(function)
    assert inspect.getsource(access_gate.DispatchPermit.run).count(".Run(") == 1


def test_rebuild_status_survives_a_launch_timeout(boundary, tmp_path):
    source = tmp_path / "Version Control.accda.src"
    source.mkdir()
    boundary.steps["connect"].hold()

    async def scenario():
        reply = await tools.vcs_call_vba(
            boundary.db, "VCS.API", ["RebuildAddIn", str(source)], timeout_seconds=0.3,
        )
        assert reply["dispatch"] == "not_dispatched"
        assert reply["rebuild_status_file"].endswith("rebuild-status.json")
        boundary.steps["connect"].release.set()
        await _gate_idle()

    asyncio.run(scenario())


def test_rebuild_launch_timeout_releases_a_late_operation(boundary, monkeypatch, tmp_path):
    """C1-3: the caller left before the callback registered; the body releases it."""
    source = tmp_path / "Version Control.accda.src"
    source.mkdir()
    (tmp_path / "Version Control.accda").write_bytes(b"x")
    registering = Step().hold()
    events: list[str] = []

    class Manager:
        def set_event_loop(self, _loop):
            pass

        def register_operation(self, **_kwargs):
            registering()
            events.append("registered")
            return "operation-1", asyncio.Queue()

        def create_callback_info(self, operation_id, _url, _client):
            return operation_id

        def unregister_operation(self, operation_id):
            events.append("unregistered " + operation_id)

    watched = AsyncMock(side_effect=AssertionError("a timed-out launch is never watched"))
    monkeypatch.setenv("ACCESS_VCS_CALL_VBA_TIMEOUT_SEC", "0.5")
    monkeypatch.setattr(tools, "check_write_permission", lambda _config: None)
    monkeypatch.setattr(tools, "get_callback_url", lambda: "http://127.0.0.1:1/callback")
    monkeypatch.setattr(tools, "_get_operation_manager", Manager)
    monkeypatch.setattr(tools, "_is_installed_addin_path", lambda _path: False)
    monkeypatch.setattr(tools, "close_owned_instances_holding", lambda *_a: [{"pid": 1}])
    monkeypatch.setattr(tools, "list_access_pids", lambda: [])
    monkeypatch.setattr(tools, "wait_for_rebuild_status", watched)
    rebuild = tools.vcs_rebuild_addin
    while hasattr(rebuild, "__wrapped__"):
        rebuild = rebuild.__wrapped__

    async def scenario():
        reply = await asyncio.wait_for(rebuild(str(source)), 3.0)
        assert reply["error_pattern"] == "timeout"
        assert reply["dispatch"] == "possibly_dispatched"
        assert reply["owned_instances_closed"] == 1
        assert events == []
        registering.release.set()
        await _gate_idle()
        assert events == ["registered", "unregistered operation-1"]
        assert len(boundary.app.runs) == 1
        watched.assert_not_awaited()

    asyncio.run(scenario())


def test_busy_reply_for_a_retained_build_claims_no_dispatch(boundary):
    gate = get_access_gate()
    hold = threading.Event()

    async def scenario():
        occupant = asyncio.create_task(gate.run_exclusive("vcs_rebuild_database", boundary.db, hold.wait, False, 5))
        while gate.current_in_flight() is None:
            await asyncio.sleep(0.005)
        busy = await _other_gated_call(boundary.db)
        assert busy["error_pattern"] == "server_busy"
        assert busy["busy_with"]["phase"] == "admitted"
        assert "dispatch" not in busy["busy_with"]
        assert "Do not re-issue" not in busy["error"]
        hold.set()
        await occupant

    asyncio.run(scenario())


def test_build_exit_hook_runs_before_the_gate_exit_wait():
    """Exit hooks run newest first: build_owner marks exit, then the gate waits."""
    from msaccess_vcs_mcp import access_gate, build_owner

    hooks = [getattr(hook, "func", hook) for hook in threading._threading_atexits]
    assert hooks.index(access_gate._gate_interpreter_exit) < hooks.index(build_owner._interpreter_exit)
