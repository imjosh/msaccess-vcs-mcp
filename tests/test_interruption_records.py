"""Interruption records: keyed by identity, carried by the gated call, used up once."""

from __future__ import annotations

import asyncio
import json
import threading
from unittest.mock import patch

import pytest

from msaccess_vcs_mcp import tools as tools_module
from msaccess_vcs_mcp.access_gate import reset_access_gate
from msaccess_vcs_mcp.dialog_recovery import (
    CLICK_DELIVERED,
    CLICK_NOT_SENT,
    CLICK_UNCERTAIN,
    ButtonInfo,
    ProcessIdentity,
    WindowInfo,
    dismiss_dialog_in_windows,
    inspect_windows,
    recover_windows,
    reset_interruptions,
)

DB = r"C:\data\Northwind.accdb"
TEST_TOOL = "vcs_test_interruptible"


class FakeBackend:
    def __init__(self, windows, create_time=1000):
        self.windows = list(windows)
        self.create_time = create_time

    def process_identity(self, pid):
        return ProcessIdentity("MSACCESS.EXE", self.create_time, True)

    def list_windows(self):
        return list(self.windows)

    def click(self, button, *, expected_pid=None, timeout_ms=5000):
        return CLICK_DELIVERED

    def close(self, hwnd):
        pass

    def responsive(self, hwnd, timeout_ms):
        return True


def _runtime_error(pid=10):
    return WindowInfo(
        hwnd=4,
        pid=pid,
        title="Microsoft Visual Basic",
        class_name="#32770",
        texts=("Run-time error '11':", "Division by zero"),
        buttons=(ButtonInfo(41, "End"), ButtonInfo(42, "Debug")),
    )


def _end_runtime_error(backend):
    return dismiss_dialog_in_windows(
        backend.list_windows(), DB, "hwnd:4", button="End", pid=10, responsive=True, backend=backend
    )


@pytest.fixture(autouse=True)
def _clean(monkeypatch):
    reset_interruptions()
    reset_access_gate()

    async def _noop(*_a, **_k):
        return None

    monkeypatch.setattr(tools_module, "_ensure_env_loaded", _noop)
    monkeypatch.setattr(tools_module, "load_config", lambda: {})
    yield
    reset_interruptions()
    reset_access_gate()


def _register(body):
    return tools_module.vcs_tool(TEST_TOOL)(body)


def _run_held(body_result, dismiss):
    """Hold one gated call open, run ``dismiss`` off-loop, then let the call finish."""

    async def scenario():
        started = asyncio.Event()
        release = asyncio.Event()

        async def body(database_path: str):
            started.set()
            await release.wait()
            return body_result

        task = asyncio.create_task(_register(body)(DB))
        await started.wait()
        await asyncio.to_thread(dismiss)
        release.set()
        return await task

    return asyncio.run(scenario())


def test_held_call_with_ended_runtime_error_returns_execution_interrupted():
    backend = FakeBackend([_runtime_error()])
    result = _run_held({"success": True, "value": 1}, lambda: _end_runtime_error(backend))
    assert result["success"] is False
    assert result["execution_interrupted"] is True
    assert result["error_pattern"] == "execution_interrupted"
    assert result["error"]
    assert result["value"] == 1


def test_original_error_text_is_kept():
    backend = FakeBackend([_runtime_error()])
    result = _run_held(
        {"success": False, "error": "COM says boom", "error_pattern": "com_error"},
        lambda: _end_runtime_error(backend),
    )
    assert result["error"] == "COM says boom"
    assert result["error_pattern"] == "execution_interrupted"
    assert result["execution_interrupted"] is True


@pytest.mark.parametrize(
    "marker",
    [{"decision_required": True}, {"error_pattern": "decision_required"}],
    ids=["decision_required_flag", "decision_required_pattern"],
)
def test_decision_required_stays_primary_and_gains_the_interruption_flag(marker):
    backend = FakeBackend([_runtime_error()])
    result = _run_held(
        {
            "success": False,
            **marker,
            "decisions": [{"id": 1}],
            "runtime_error": "Run-time error '11': Division by zero",
            "error": "A merge conflict needs a decision",
        },
        lambda: _end_runtime_error(backend),
    )
    assert result["success"] is False
    assert result["execution_interrupted"] is True
    assert result.get("error_pattern") == marker.get("error_pattern")
    assert result.get("decision_required") == marker.get("decision_required")
    assert result["decisions"] == [{"id": 1}]
    assert result["runtime_error"] == "Run-time error '11': Division by zero"
    assert result["error"] == "A merge conflict needs a decision"
    # The record was still used up.
    report = inspect_windows([], DB, backend=FakeBackend([]), pid=10, responsive=True)
    assert report["execution_interrupted"] is False


def test_record_is_used_up_by_the_call_it_interrupted():
    backend = FakeBackend([_runtime_error()])

    async def scenario():
        started, release = asyncio.Event(), asyncio.Event()
        first = True

        async def body(database_path: str):
            nonlocal first
            if first:
                first = False
                started.set()
                await release.wait()
            return {"success": True}

        tool = _register(body)
        task = asyncio.create_task(tool(DB))
        await started.wait()
        await asyncio.to_thread(_end_runtime_error, backend)
        release.set()
        return await task, await tool(DB)

    interrupted, healthy = asyncio.run(scenario())
    assert interrupted["execution_interrupted"] is True
    assert healthy == {"success": True}


def test_compile_error_dismissed_by_recovery_also_records():
    compile_err = WindowInfo(
        hwnd=5,
        pid=10,
        title="Microsoft Visual Basic for Applications",
        class_name="#32770",
        texts=("Compile error:", "Syntax error"),
        buttons=(ButtonInfo(51, "OK"),),
    )
    backend = FakeBackend([compile_err])

    def recover():
        recover_windows(
            [compile_err], DB, policy="end_runtime_error", pid=10, responsive=True, backend=backend
        )

    result = _run_held({"success": True}, recover)
    assert result["execution_interrupted"] is True
    assert result["success"] is False


def test_new_identity_with_same_pid_shows_no_interruption():
    backend = FakeBackend([_runtime_error()], create_time=1000)
    _end_runtime_error(backend)
    same = inspect_windows([], DB, backend=backend, pid=10, responsive=True)
    assert same["execution_interrupted"] is True

    reused = FakeBackend([], create_time=2000)
    fresh = inspect_windows([], DB, backend=reused, pid=10, responsive=True)
    assert fresh["execution_interrupted"] is False
    assert fresh["last_interruption"] is None


def test_free_record_shows_until_next_gated_call_on_that_database():
    backend = FakeBackend([_runtime_error()])
    _end_runtime_error(backend)  # gate is free: nothing to attach to
    shown = inspect_windows([], DB, backend=backend, pid=10, responsive=True)
    assert shown["last_interruption"]["busy_with"] is None

    async def body(database_path: str):
        return {"success": True}

    result = asyncio.run(_register(body)(DB))
    assert result == {"success": True}  # a free record never taints the next call
    gone = inspect_windows([], DB, backend=backend, pid=10, responsive=True)
    assert gone["execution_interrupted"] is False


def test_free_record_survives_gated_call_on_another_database():
    backend = FakeBackend([_runtime_error()])
    _end_runtime_error(backend)

    async def body(database_path: str):
        return {"success": True}

    asyncio.run(_register(body)(r"C:\other.accdb"))
    kept = inspect_windows([], DB, backend=backend, pid=10, responsive=True)
    assert kept["execution_interrupted"] is True


def _compile_error():
    return WindowInfo(
        hwnd=5,
        pid=10,
        title="Microsoft Visual Basic for Applications",
        class_name="#32770",
        texts=("Compile error:", "Syntax error"),
        buttons=(ButtonInfo(51, "OK"),),
    )


def _access_dialog_mentioning_error():
    return WindowInfo(
        hwnd=6,
        pid=10,
        title="Microsoft Access",
        class_name="#32770",
        texts=("The record could not be saved because of an error.",),
        buttons=(ButtonInfo(61, "OK"),),
    )


def _dismiss_explicitly(window, button):
    backend = FakeBackend([window])
    return dismiss_dialog_in_windows(
        [window], DB, f"hwnd:{window.hwnd}", button=button, pid=10, responsive=True, backend=backend
    ), backend


def _dismiss_by_recovery(window, _button):
    backend = FakeBackend([window])
    return recover_windows(
        [window], DB, policy="end_runtime_error", pid=10, responsive=True, backend=backend
    ), backend


@pytest.mark.parametrize("dismiss", [_dismiss_explicitly, _dismiss_by_recovery])
@pytest.mark.parametrize(
    "window, button, is_failure",
    [
        (_runtime_error(), "End", True),
        (_compile_error(), "OK", True),
        # Wording is not a failure signature: only runtime or compile errors and End are.
        (_access_dialog_mentioning_error(), "OK", False),
    ],
    ids=["runtime_error", "compile_error", "access_dialog_saying_error"],
)
def test_both_dismissal_paths_share_one_failure_rule(dismiss, window, button, is_failure):
    result, backend = dismiss(window, button)
    assert result["success"] is True
    assert result["failure_dialog_dismissed"] is is_failure
    assert result["interrupted"] is is_failure
    follow = inspect_windows([], DB, backend=backend, pid=10, responsive=True)
    assert follow["execution_interrupted"] is is_failure


# The final outcome through the public tools: the client result, the usage entry
# and recent-call history agree.


class ClosingBackend(FakeBackend):
    """A click closes the dialog that owns the button, so the dismissal is confirmed."""

    def click(self, button, *, expected_pid=None, timeout_ms=5000):
        self.windows = [w for w in self.windows if button not in w.buttons]
        return CLICK_DELIVERED


@pytest.fixture
def usage_log(tmp_path, monkeypatch):
    from msaccess_vcs_mcp import usage_logging

    log_dir = tmp_path / "usage"
    monkeypatch.setenv("ACCESS_VCS_ENABLE_LOGGING", "true")
    monkeypatch.setenv("ACCESS_VCS_LOG_DIR", str(log_dir))
    monkeypatch.setenv("ACCESS_VCS_DISABLE_DIAGNOSTIC_LOG", "true")
    usage_logging.reset_logging()
    yield log_dir / "vcs-mcp-usage.jsonl"
    usage_logging.reset_logging()


def _tool_entries(log_file):
    entries = [json.loads(line) for line in log_file.read_text(encoding="utf-8").splitlines()]
    return [e for e in entries if e.get("event") == "tool_call" and e.get("tool") == TEST_TOOL]


def _run_through_public_dismiss(body_result, *, is_async):
    """Hold a gated call, End its runtime error through ``vcs_dismiss_dialog``, then finish it.

    Returns the dismissal, the held call's result, and ``vcs_get_recent_calls``.
    """
    started, release = threading.Event(), threading.Event()
    if is_async:
        async def body(database_path: str):
            started.set()
            await asyncio.to_thread(release.wait, 5)
            return dict(body_result)
    else:
        def body(database_path: str):
            started.set()
            release.wait(5)
            return dict(body_result)

    async def scenario():
        task = asyncio.create_task(_register(body)(DB))
        assert await asyncio.to_thread(started.wait, 5)
        dismissed = await tools_module.vcs_dismiss_dialog(DB, "hwnd:4", button="End", pid=10)
        release.set()
        result = await task
        recent = await tools_module.vcs_get_recent_calls(limit=10)
        return dismissed, result, recent

    backend = ClosingBackend([_runtime_error()])
    with (
        patch("msaccess_vcs_mcp.tools.validate_database_path", return_value=DB),
        patch("msaccess_vcs_mcp.dialog_recovery.list_owned", return_value=[]),
        patch("msaccess_vcs_mcp.dialog_recovery._default_backend", return_value=backend),
    ):
        return asyncio.run(scenario())


@pytest.mark.parametrize("is_async", [False, True], ids=["sync", "async"])
def test_interrupted_call_logs_the_final_failure_not_the_handler_success(usage_log, is_async):
    dismissed, result, recent = _run_through_public_dismiss(
        {"success": True, "value": 1}, is_async=is_async
    )
    assert dismissed["success"] is True
    assert dismissed["interrupted"] is True

    assert result["success"] is False
    assert result["execution_interrupted"] is True
    assert result["error_pattern"] == "execution_interrupted"
    assert result["value"] == 1

    entries = _tool_entries(usage_log)
    assert len(entries) == 1
    entry = entries[0]
    assert entry["success"] is False
    assert entry["error_pattern"] == result["error_pattern"]
    assert entry["error"] == result["error"]
    assert entry["execution_interrupted"] is True
    assert isinstance(entry["execution_time_ms"], (int, float))
    assert entry["execution_time_ms"] >= 0

    history = [e for e in recent["entries"] if e.get("tool") == TEST_TOOL]
    assert len(history) == 1
    assert history[0]["success"] is False
    assert history[0]["error_pattern"] == "execution_interrupted"
    assert history[0]["execution_interrupted"] is True


def test_original_error_and_policy_cleanup_survive_normalization_and_logging(usage_log):
    _dismissed, result, _recent = _run_through_public_dismiss(
        {
            "success": False,
            "error": "COM says boom",
            "error_pattern": "com_error",
            "policy_cleanup_error": "ClearOperationPolicy refused",
        },
        is_async=False,
    )
    assert result["error"] == "COM says boom"
    assert result["policy_cleanup_error"] == "ClearOperationPolicy refused"
    assert result["error_pattern"] == "execution_interrupted"

    [entry] = _tool_entries(usage_log)
    assert entry["error"] == "COM says boom"
    assert entry["policy_cleanup_error"] == "ClearOperationPolicy refused"
    assert entry["error_pattern"] == "execution_interrupted"


def test_interrupted_decision_required_is_logged_as_decision_required(usage_log):
    _dismissed, result, _recent = _run_through_public_dismiss(
        {
            "success": False,
            "decision_required": True,
            "error_pattern": "decision_required",
            "decisions": [{"id": 1}],
        },
        is_async=True,
    )
    assert result["error_pattern"] == "decision_required"
    assert result["execution_interrupted"] is True

    [entry] = _tool_entries(usage_log)
    assert entry["success"] is False
    assert entry["error_pattern"] == "decision_required"
    assert entry["execution_interrupted"] is True


def test_healthy_call_after_an_interrupted_one_logs_success(usage_log):
    _run_through_public_dismiss({"success": True}, is_async=False)

    async def body(database_path: str):
        return {"success": True}

    healthy = asyncio.run(_register(body)(DB))
    assert healthy == {"success": True}

    first, second = _tool_entries(usage_log)
    assert first["execution_interrupted"] is True
    assert second["success"] is True
    assert "execution_interrupted" not in second
    assert "error_pattern" not in second


def test_call_that_raises_uses_up_its_record(usage_log):
    backend = FakeBackend([_runtime_error()])

    async def scenario():
        started, release = asyncio.Event(), asyncio.Event()

        async def body(database_path: str):
            started.set()
            await release.wait()
            raise RuntimeError("handler broke")

        task = asyncio.create_task(_register(body)(DB))
        await started.wait()
        await asyncio.to_thread(_end_runtime_error, backend)
        release.set()
        with pytest.raises(RuntimeError, match="handler broke"):
            await task

    asyncio.run(scenario())
    [entry] = _tool_entries(usage_log)
    assert entry["success"] is False
    report = inspect_windows([], DB, backend=FakeBackend([]), pid=10, responsive=True)
    assert report["execution_interrupted"] is False


# The dismissal releases the blocked call from inside the click: Access resumes
# and the call returns before the click does. The interruption is reserved before
# the click, so the call still sees it.


class ReleasingBackend(FakeBackend):
    """A click that lets the held call finish before the click returns.

    ``finishing`` is set once the handler has returned and its result is being
    finalized. The click then gives that finalization time to run to completion,
    as it would if nothing waited on the click, and returns ``outcome``. With
    ``hang`` the click returns only once the call has finished.
    """

    def __init__(self, windows, *, release, finishing, finished, outcome, hang):
        super().__init__(windows)
        self.release, self.finishing, self.finished = release, finishing, finished
        self.outcome = outcome
        self.hang = hang

    def click(self, button, *, expected_pid=None, timeout_ms=5000):
        self.release.set()
        assert self.finishing.wait(5), "the held call never reached its finalization"
        self.finished.wait(5 if self.hang else 0.3)
        if self.outcome == CLICK_DELIVERED:
            self.windows = [w for w in self.windows if button not in w.buttons]
        return self.outcome


def _run_released_inside_click(
    monkeypatch,
    body_result,
    *,
    is_async=False,
    recover=False,
    outcome=CLICK_DELIVERED,
    hang=False,
    timeout_seconds=None,
):
    """End the held call's runtime error through a public tool whose click releases the call.

    Returns the dialog tool's result, the held call's result, and an inspection
    taken after both finished.
    """
    started, release = threading.Event(), threading.Event()
    finishing, finished = threading.Event(), threading.Event()
    real_finish = tools_module.finish_gated_call

    def observed_finish(call_id, result):
        if result is not None:
            finishing.set()
        try:
            return real_finish(call_id, result)
        finally:
            if result is not None:
                finished.set()

    monkeypatch.setattr(tools_module, "finish_gated_call", observed_finish)

    if is_async:
        async def body(database_path: str):
            started.set()
            await asyncio.to_thread(release.wait, 5)
            return dict(body_result)
    else:
        def body(database_path: str):
            started.set()
            release.wait(5)
            return dict(body_result)

    backend = ReleasingBackend(
        [_runtime_error()],
        release=release,
        finishing=finishing,
        finished=finished,
        outcome=outcome,
        hang=hang,
    )

    async def scenario():
        task = asyncio.create_task(_register(body)(DB))
        assert await asyncio.to_thread(started.wait, 5)
        if recover:
            dialog = await tools_module.vcs_recover_dialogs(
                DB, policy="end_runtime_error", pid=10, timeout_seconds=timeout_seconds
            )
        else:
            dialog = await tools_module.vcs_dismiss_dialog(
                DB, "hwnd:4", button="End", pid=10, timeout_seconds=timeout_seconds
            )
        return dialog, await task

    with (
        patch("msaccess_vcs_mcp.tools.validate_database_path", return_value=DB),
        patch("msaccess_vcs_mcp.dialog_recovery.list_owned", return_value=[]),
        patch("msaccess_vcs_mcp.dialog_recovery._default_backend", return_value=backend),
    ):
        dialog, result = asyncio.run(scenario())
    after = inspect_windows([], DB, backend=FakeBackend([]), pid=10, responsive=True)
    return dialog, result, after


def _healthy_call():
    async def body(database_path: str):
        return {"success": True}

    return asyncio.run(_register(body)(DB))


@pytest.mark.parametrize("is_async", [False, True], ids=["sync", "async"])
@pytest.mark.parametrize("recover", [False, True], ids=["dismiss", "recover"])
def test_call_released_inside_the_click_is_interrupted(monkeypatch, recover, is_async):
    dialog, result, after = _run_released_inside_click(
        monkeypatch, {"success": True, "value": 1}, recover=recover, is_async=is_async
    )
    assert dialog["success"] is True
    assert dialog["interrupted"] is True
    assert result["success"] is False
    assert result["execution_interrupted"] is True
    assert result["error_pattern"] == "execution_interrupted"
    assert result["value"] == 1
    assert "interruption_uncertain" not in result
    # Used up by the call it interrupted: nothing is left to show or inherit.
    assert after["execution_interrupted"] is False
    assert _healthy_call() == {"success": True}


def test_call_released_inside_the_click_logs_the_interruption(monkeypatch, usage_log):
    _dialog, result, _after = _run_released_inside_click(monkeypatch, {"success": True})
    [entry] = _tool_entries(usage_log)
    assert entry["success"] is False
    assert entry["error_pattern"] == "execution_interrupted"
    assert entry["execution_interrupted"] is True
    assert entry["error"] == result["error"]


@pytest.mark.parametrize("recover", [False, True], ids=["dismiss", "recover"])
def test_click_that_was_not_sent_leaves_the_call_alone(monkeypatch, recover):
    dialog, result, after = _run_released_inside_click(
        monkeypatch, {"success": True, "value": 1}, recover=recover, outcome=CLICK_NOT_SENT
    )
    assert dialog["error_pattern"] == "dismiss_uncertain"
    assert result == {"success": True, "value": 1}
    assert after["execution_interrupted"] is False
    assert after["last_interruption"] is None


@pytest.mark.parametrize("recover", [False, True], ids=["dismiss", "recover"])
def test_uncertain_click_never_leaves_a_success(monkeypatch, usage_log, recover):
    dialog, result, after = _run_released_inside_click(
        monkeypatch, {"success": True, "value": 1}, recover=recover, outcome=CLICK_UNCERTAIN
    )
    assert dialog["error_pattern"] == "dismiss_uncertain"
    assert result["success"] is False
    assert result["interruption_uncertain"] is True
    assert result["error_pattern"] == "interruption_uncertain"
    assert result["error"]
    assert "execution_interrupted" not in result
    assert result["value"] == 1
    # Not a confirmed interruption, and not carried to a later call.
    assert after["execution_interrupted"] is False
    assert _healthy_call() == {"success": True}
    first, second = _tool_entries(usage_log)
    assert first["success"] is False
    assert first["interruption_uncertain"] is True
    assert first["error_pattern"] == "interruption_uncertain"
    assert second["success"] is True
    assert "interruption_uncertain" not in second


def test_uncertain_click_keeps_the_call_failure_primary(monkeypatch):
    _dialog, result, _after = _run_released_inside_click(
        monkeypatch,
        {"success": False, "error": "COM says boom", "error_pattern": "com_error"},
        outcome=CLICK_UNCERTAIN,
    )
    assert result["success"] is False
    assert result["error"] == "COM says boom"
    assert result["error_pattern"] == "com_error"
    assert result["interruption_uncertain"] is True


def test_uncertain_click_keeps_decision_required_primary(monkeypatch):
    _dialog, result, _after = _run_released_inside_click(
        monkeypatch,
        {
            "success": False,
            "decision_required": True,
            "error_pattern": "decision_required",
            "decisions": [{"id": 1}],
        },
        outcome=CLICK_UNCERTAIN,
    )
    assert result["success"] is False
    assert result["interruption_uncertain"] is True
    assert result["error_pattern"] == "decision_required"
    assert result["decisions"] == [{"id": 1}]


def test_click_that_outlives_its_deadline_is_uncertain_and_leaves_nothing_behind(monkeypatch):
    # The call waits for the click only until its deadline (twice the timeout
    # plus a margin), then finishes while the click is still running.
    dialog, result, after = _run_released_inside_click(
        monkeypatch, {"success": True}, hang=True, timeout_seconds=0.1
    )
    assert dialog["success"] is True
    assert result["success"] is False
    assert result["interruption_uncertain"] is True
    assert "execution_interrupted" not in result
    # The click then reports delivery, but its call already used the reservation up.
    assert after["execution_interrupted"] is False
    assert after["last_interruption"] is None
    assert _healthy_call() == {"success": True}


@pytest.mark.parametrize("outcome", [CLICK_NOT_SENT, CLICK_UNCERTAIN])
def test_undelivered_click_outside_a_call_records_nothing(outcome):
    class Undelivered(FakeBackend):
        def click(self, button, *, expected_pid=None, timeout_ms=5000):
            return outcome

    backend = Undelivered([_runtime_error()])
    result = _end_runtime_error(backend)
    assert result["error_pattern"] == "dismiss_uncertain"
    after = inspect_windows([], DB, backend=backend, pid=10, responsive=True)
    assert after["execution_interrupted"] is False


def test_reservation_for_a_call_already_released_is_dropped_by_the_next_call():
    # The dismissal read the in-flight call just before the gate released it, and
    # the reservation landed after that call had used up its records.
    from msaccess_vcs_mcp.access_gate import InFlight, get_access_gate

    released = InFlight(tool=TEST_TOOL, database=DB, started_at=0.0, call_id=987654)
    backend = FakeBackend([_runtime_error()])
    with patch.object(get_access_gate(), "current_in_flight", return_value=released):
        _end_runtime_error(backend)
    orphan = inspect_windows([], DB, backend=backend, pid=10, responsive=True)
    assert orphan["last_interruption"]["busy_with"]["tool"] == TEST_TOOL

    assert _healthy_call() == {"success": True}
    after = inspect_windows([], DB, backend=backend, pid=10, responsive=True)
    assert after["execution_interrupted"] is False
