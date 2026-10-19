"""Serialize Access COM work across concurrent MCP tool calls.

One MCP server process is shared by every Cursor window. Sync tool bodies
previously ran on the asyncio event loop and blocked all other requests.
This module fronts Access-touching tools with a single COM apartment thread
and a one-at-a-time slot so callers get a fast ``server_busy`` answer
instead of queueing into a client-side timeout. Waiting for the slot uses no
worker thread, so a thread pool filled by hung workers cannot delay
``server_busy`` past its deadline.

Async tool bodies run on that apartment thread too, in an event loop of their
own. A blocking COM step inside one (connect, add-in probe, a synchronous API
call) then holds up only that thread, never the server loop that the dialog,
status and cancel tools answer on, and every COM object the body creates is
used and released on the thread that created it.

Caller cancellation ends the wait, not ownership. The worker keeps its slot
and call identity until the body and apartment-loop cleanup actually finish.

A tool registered with a deadline gets one budget for the whole call: the slot
wait, connecting, session validation and the target itself. Every gated call
carries an ``Attempt``. A deadline body asks it for a ``DispatchPermit``
before reaching its target, and the permit is refused once the caller has
stopped waiting or the deadline has passed. Caller expiry never releases the
slot; the call settles only when its body returns on the apartment.
"""

from __future__ import annotations

import asyncio
import contextvars
import itertools
import os
import queue
import threading
import time
from concurrent.futures import Future
from dataclasses import dataclass, field
from enum import IntEnum
from typing import Any, Callable

try:
    import pythoncom

    COM_AVAILABLE = True
except ImportError:
    COM_AVAILABLE = False

DEFAULT_BUSY_WAIT_SEC = 15.0
SLOT_POLL_SEC = 0.02
# How long interpreter exit waits past an in-flight call's own deadline.
EXIT_GRACE_SEC = 5.0
_call_ids = itertools.count(1)
_retention = contextvars.ContextVar("build_retention", default=None)
_attempt_var = contextvars.ContextVar("access_attempt", default=None)


def _retain_build_host(finish: Callable[[], None]) -> None:
    lifetime = _retention.get()
    if lifetime is None:
        raise RuntimeError("Build host requires the Access apartment gate")
    if lifetime:
        raise RuntimeError("The current call already owns a build host")
    lifetime.append(finish)

# Tools that can run independently of the Access gate. Dialog/status/cancel
# tools must stay responsive during a long operation. The version probe uses
# its own COM-initialized worker; rebuild acquires the gate for its launch.
EXEMPT_TOOLS = frozenset({
    "vcs_get_version_info",
    "vcs_cancel_operation",
    "vcs_get_recent_calls",
    # Launch holds the gate itself; the subsequent status-file wait must not.
    "vcs_rebuild_addin",
    # Dialog recovery uses Win32/UI messages, not Access COM, so it must stay
    # callable while a modal dialog or VBA break holds the Access gate.
    "vcs_list_dialogs",
    "vcs_dismiss_dialog",
    "vcs_recover_dialogs",
    "vcs_automation_status",
})


def _read_busy_wait_sec() -> float:
    try:
        value = float(os.environ.get("ACCESS_VCS_BUSY_WAIT_SEC", str(DEFAULT_BUSY_WAIT_SEC)))
    except ValueError:
        return DEFAULT_BUSY_WAIT_SEC
    return value if value > 0 else DEFAULT_BUSY_WAIT_SEC


def _init_com_apartment() -> None:
    if COM_AVAILABLE:
        pythoncom.CoInitialize()


class Phase(IntEnum):
    """Where a gated call is. It only moves forward."""

    ADMITTED = 1
    VALIDATING = 2
    DISPATCH_PERMITTED = 3
    TARGET_POSSIBLY_RUNNING = 4
    SETTLED = 5


WAITING, ANSWERED, EXPIRED = "waiting", "answered", "expired"


def derive_dispatch(phase: Phase, permit_path: bool, reached_target: bool) -> str | None:
    """Whether the target was reached, as every reply reports it.

    Only calls that ask for a dispatch permit can say. Any other call, such
    as a retained build, has no dispatch value rather than a false one.
    """
    if not permit_path:
        return None
    if phase <= Phase.VALIDATING:
        return "not_dispatched"
    if phase is Phase.SETTLED:
        return "dispatched" if reached_target else "not_dispatched"
    return "possibly_dispatched"


class DispatchPermit:
    """The single right to reach the target, granted by ``Attempt.permit_dispatch``."""

    def __init__(self, attempt: "Attempt") -> None:
        self._attempt = attempt
        self._used = False

    def run(self, app: Any, procedure: str, *args: Any) -> Any:
        """Call ``Application.Run`` once on the apartment. A caller expiry from here on cannot stop it."""
        if self._used:
            raise RuntimeError("A dispatch permit is used once")
        self._used = True
        self._attempt._reach_target()
        return app.Run(procedure, *args)


class Attempt:
    """One gated call: its phase, its caller, and the deadline they share.

    ``permit_dispatch`` and ``caller_expired`` take the same lock, so exactly
    one of them decides whether a target is reached after the deadline.
    """

    def __init__(
        self,
        call_id: int,
        tool: str,
        database: str | None,
        deadline_seconds: float | None = None,
        received_at: float | None = None,
        clock: Callable[[], float] = time.monotonic,
    ) -> None:
        self.call_id = call_id
        self.tool = tool
        self.database = database
        self.deadline_seconds = deadline_seconds
        self._clock = clock
        self.received_at = clock() if received_at is None else received_at
        self.deadline_at = None if deadline_seconds is None else self.received_at + deadline_seconds
        self._lock = threading.Lock()
        self._phase = Phase.ADMITTED
        self._caller = WAITING
        self._cause: str | None = None
        self._permit_path = False
        self._reached_target = False
        self._annotations: dict[str, Any] = {}

    def remaining(self) -> float | None:
        if self.deadline_at is None:
            return None
        return self.deadline_at - self._clock()

    def annotate(self, **fields: Any) -> None:
        """Fields the timeout reply carries, such as the function being called."""
        with self._lock:
            self._annotations.update(fields)

    def annotations(self) -> dict[str, Any]:
        with self._lock:
            return dict(self._annotations)

    def _advance(self, phase: Phase) -> None:
        if phase <= self._phase:
            raise RuntimeError(f"Attempt {self.call_id} cannot move from {self._phase.name} to {phase.name}")
        self._phase = phase

    def begin_validation(self) -> None:
        with self._lock:
            self._permit_path = True
            self._advance(Phase.VALIDATING)

    def _still_wanted(self) -> bool:
        # Called with the lock held. The clock decides even before the loop
        # has noticed the deadline.
        remaining = None if self.deadline_at is None else self.deadline_at - self._clock()
        if self._caller == WAITING and remaining is not None and remaining <= 0:
            self._caller, self._cause = EXPIRED, "deadline"
        return self._caller == WAITING

    def checkpoint(self) -> bool:
        """Whether validation should continue: False once the caller stopped waiting.

        Validation steps that act on Access (the session handshake) run only
        after a passing checkpoint, so an expired call stops at the next step
        instead of changing add-in state nobody will read.
        """
        with self._lock:
            return self._still_wanted()

    def permit_dispatch(self) -> DispatchPermit | None:
        """Grant the target dispatch, or None once the caller is no longer waiting."""
        with self._lock:
            if not self._still_wanted():
                return None
            self._advance(Phase.DISPATCH_PERMITTED)
            return DispatchPermit(self)

    def _reach_target(self) -> None:
        with self._lock:
            self._reached_target = True
            self._advance(Phase.TARGET_POSSIBLY_RUNNING)

    def caller_expired(self, cause: str) -> dict[str, Any]:
        """Record that the caller stopped waiting. It never settles the attempt."""
        with self._lock:
            if self._phase is Phase.SETTLED:
                self._caller = ANSWERED
            elif self._caller == WAITING:
                self._caller, self._cause = EXPIRED, cause
            return self._snapshot()

    def caller_answered(self) -> None:
        with self._lock:
            if self._caller == WAITING:
                self._caller = ANSWERED

    def settle(self) -> None:
        with self._lock:
            if self._phase is not Phase.SETTLED:
                self._phase = Phase.SETTLED

    def snapshot(self) -> dict[str, Any]:
        with self._lock:
            return self._snapshot()

    def _snapshot(self) -> dict[str, Any]:
        now = self._clock()
        view: dict[str, Any] = {
            "call_id": self.call_id,
            "phase": self._phase.name.lower(),
            "caller": self._caller,
            "elapsed_ms": round((now - self.received_at) * 1000, 2),
        }
        dispatch = derive_dispatch(self._phase, self._permit_path, self._reached_target)
        if dispatch is not None:
            view["dispatch"] = dispatch
        if self.deadline_seconds is not None:
            view["deadline_seconds"] = self.deadline_seconds
            if self._phase is not Phase.SETTLED and now > self.deadline_at:
                view["overdue_ms"] = round((now - self.deadline_at) * 1000, 2)
        if self._cause is not None:
            view["expired_by"] = self._cause
        return view


def current_attempt() -> Attempt | None:
    """The gated call running in this context, or None outside the gate."""
    return _attempt_var.get()


def _timeout_next_step(dispatch: str, retained: bool) -> str:
    if dispatch == "possibly_dispatched":
        return (
            "The target may still be running, and the server holds the Access slot until it returns. "
            "Do not re-issue it blindly: vcs_automation_status() shows this call_id until the target returns, "
            "and the server's diagnostic log records the late outcome as attempt_settled_after_expiry. "
            "If Access is waiting on a dialog, vcs_list_dialogs() and vcs_dismiss_dialog() can end it."
        )
    if retained:
        return (
            "The target was never reached and will not run for this request. The server still holds the "
            "Access slot until its current step returns; vcs_automation_status() shows it. Retry after that."
        )
    return "Nothing ran for this request. Retry when the server is free."


def timeout_reply(
    tool: str,
    deadline_seconds: float,
    view: dict[str, Any],
    *,
    gate_retained: bool,
    annotations: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """The one reply for a call whose deadline passed, wherever it was noticed."""
    dispatch = view.get("dispatch", "not_dispatched")
    phase = view["phase"]
    reply: dict[str, Any] = {
        "success": False,
        "error": f"{tool} did not finish within its {deadline_seconds:g} second deadline (phase: {phase}).",
        "error_pattern": "timeout",
        "recoverable": True,
        "timed_out": True,
        "call_id": view.get("call_id"),
        "phase": phase,
        "dispatch": dispatch,
        "retry_safe": dispatch == "not_dispatched",
        "gate_retained": gate_retained,
        "deadline_seconds": deadline_seconds,
        "duration_ms": view.get("elapsed_ms"),
        "next_step": _timeout_next_step(dispatch, gate_retained),
    }
    for name, value in (annotations or {}).items():
        reply.setdefault(name, value)
    return reply


def attempt_timeout_reply(attempt: Attempt, *, gate_retained: bool) -> dict[str, Any]:
    return timeout_reply(
        attempt.tool, attempt.deadline_seconds or 0.0, attempt.snapshot(),
        gate_retained=gate_retained, annotations=attempt.annotations(),
    )


@dataclass(frozen=True)
class InFlight:
    tool: str
    database: str | None
    started_at: float
    # Identifies this call among all others. Interruption records use it to
    # find the call they belong to, so a later call never inherits one.
    call_id: int = 0
    attempt: Attempt | None = field(default=None, compare=False)


class ServerLoopContext:
    """The MCP ``Context`` handed to a tool body that runs on the apartment loop.

    Progress goes out through the client session, which belongs to the server
    loop, so each report is sent there and awaited from the body's loop.
    """

    def __init__(self, ctx: Any, loop: asyncio.AbstractEventLoop) -> None:
        self._ctx = ctx
        self._loop = loop

    async def report_progress(
        self, progress: float, total: float | None = None, message: str | None = None
    ) -> None:
        sent = asyncio.run_coroutine_threadsafe(
            self._ctx.report_progress(progress=progress, total=total, message=message),
            self._loop,
        )
        await asyncio.wrap_future(sent)


def _busy_error(in_flight: InFlight) -> dict[str, Any]:
    elapsed_ms = round((time.perf_counter() - in_flight.started_at) * 1000, 2)
    db_hint = f" on {in_flight.database}" if in_flight.database else ""
    busy_with: dict[str, Any] = {
        "tool": in_flight.tool,
        "database": in_flight.database,
        "elapsed_ms": elapsed_ms,
    }
    detail = ""
    if in_flight.attempt is not None:
        view = in_flight.attempt.snapshot()
        busy_with.update({key: view[key] for key in ("call_id", "phase", "caller", "dispatch", "overdue_ms") if key in view})
        if view.get("caller") == EXPIRED and view.get("dispatch") == "possibly_dispatched":
            detail = (
                f" Call {view['call_id']} passed its deadline after it was dispatched; it is still running "
                "and holds the slot. Do not re-issue that call."
            )
    return {
        "success": False,
        "error": (
            f"Another Access operation is in progress ({in_flight.tool}{db_hint}). "
            "The MCP server runs one Access operation at a time across all Cursor "
            "windows sharing this server process. Retry after the in-flight call "
            "completes, or call vcs_get_recent_calls() to see what is running." + detail
        ),
        "error_pattern": "server_busy",
        "recoverable": True,
        "busy_with": busy_with,
        "retry_after_seconds": 5,
    }


class _Apartment:
    """The gate's single COM apartment thread, run as a daemon.

    Interpreter exit therefore never joins a thread stuck inside a COM call.
    ``_gate_interpreter_exit`` decides how long exit waits for the call in flight.
    """

    def __init__(self, initializer: Callable[[], None], name: str = "vcs-access-apartment") -> None:
        self._name = name
        self._initializer = initializer
        self._queue: queue.SimpleQueue = queue.SimpleQueue()
        self._lock = threading.Lock()
        self._thread: threading.Thread | None = None
        self._closed = False

    def submit(self, fn: Callable[..., Any], /, *args: Any, **kwargs: Any) -> Future:
        future: Future = Future()
        with self._lock:
            if self._closed:
                raise RuntimeError("cannot schedule new futures after shutdown")
            self._queue.put((future, fn, args, kwargs))
            if self._thread is None:
                self._thread = threading.Thread(target=self._work, name=f"{self._name}_0", daemon=True)
                self._thread.start()
        return future

    def _work(self) -> None:
        broken: BaseException | None = None
        try:
            self._initializer()
        except BaseException as error:
            broken = error
        while True:
            item = self._queue.get()
            if item is None:
                return
            future, fn, args, kwargs = item
            del item
            if not future.set_running_or_notify_cancel():
                continue
            if broken is not None:
                future.set_exception(broken)
                continue
            try:
                result = fn(*args, **kwargs)
            except BaseException as error:
                future.set_exception(error)
            else:
                future.set_result(result)
            del future, fn, args, kwargs

    def shutdown(self, wait: bool = True, *, cancel_futures: bool = False) -> None:
        with self._lock:
            self._closed = True
            if cancel_futures:
                while True:
                    try:
                        item = self._queue.get_nowait()
                    except queue.Empty:
                        break
                    if item is not None:
                        item[0].cancel()
            self._queue.put(None)
            thread = self._thread
        if wait and thread is not None:
            thread.join()


class AccessGate:
    """One COM apartment thread and one in-flight Access operation at a time."""

    def __init__(self) -> None:
        self._state_lock = threading.Lock()
        self._in_flight: InFlight | None = None
        self._slot = threading.Lock()
        self._executor = _Apartment(_init_com_apartment)
        self._com_initialized = False

    @property
    def com_initialized(self) -> bool:
        return self._com_initialized

    def _mark_com_initialized(self) -> None:
        self._com_initialized = True

    def _note_apartment_used(self) -> None:
        if COM_AVAILABLE and not self._com_initialized:
            # Executor initializer runs once per worker thread; record it
            # for tests that assert COM was initialized.
            self._mark_com_initialized()

    def _run_on_apartment_loop(
        self, fn: Callable[..., Any], args: tuple, kwargs: dict,
        context: contextvars.Context, body_task: dict[str, Any], task_lock: threading.Lock,
    ) -> Any:
        """Run an async body to completion in a fresh event loop on the apartment thread.

        The body keeps the caller's context variables. Cancelling the caller
        cancels the body at its next await; a blocking step it is inside runs on.
        """
        async def body() -> Any:
            with task_lock:
                body_task["loop"] = asyncio.get_running_loop()
                body_task["task"] = asyncio.current_task()
                if body_task.get("cancelled"):
                    raise asyncio.CancelledError
            return await fn(*args, **kwargs)

        with asyncio.Runner() as runner:
            return runner.run(body(), context=context)

    def _release_slot(self) -> None:
        with self._state_lock:
            self._in_flight = None
            self._slot.release()

    def current_in_flight(self) -> InFlight | None:
        with self._state_lock:
            return self._in_flight

    async def _acquire_slot(self, wait_sec: float) -> bool:
        """Take the slot within ``wait_sec``, polling on the event loop.

        Not ``asyncio.to_thread``: abandoned workers in the default executor
        would leave this wait queued before its timeout even started.
        """
        deadline = time.monotonic() + wait_sec
        while not self._slot.acquire(blocking=False):
            remaining = deadline - time.monotonic()
            if remaining <= 0:
                return False
            await asyncio.sleep(min(SLOT_POLL_SEC, remaining))
        return True

    async def run_exclusive(
        self,
        tool: str,
        database: str | None,
        fn: Callable[..., Any],
        is_async: bool,
        /,
        *args: Any,
        on_worker_finished: Callable[[], None] | None = None,
        deadline_seconds: float | None = None,
        received_at: float | None = None,
        **kwargs: Any,
    ) -> Any:
        """Hold the slot through worker completion, independently of the caller.

        ``on_worker_finished`` runs on the apartment after body/loop cleanup,
        before releasing ownership. It is skipped if no worker starts.

        With ``deadline_seconds`` the caller waits at most that long from
        ``received_at`` (default: now), slot wait included, and then gets the
        timeout reply. The body keeps the slot until it returns.
        """
        entered = time.monotonic() if received_at is None else received_at
        busy_wait = _read_busy_wait_sec()
        wait = busy_wait
        if deadline_seconds is not None:
            budget = entered + deadline_seconds - time.monotonic()
            wait = min(busy_wait, max(0.0, budget))
        acquired = wait > 0 and await self._acquire_slot(wait)
        if not acquired and deadline_seconds is not None and wait < busy_wait:
            # The caller's budget ran out before the slot was free; nothing ran.
            queued = {"phase": "queued", "dispatch": "not_dispatched",
                      "elapsed_ms": round((time.monotonic() - entered) * 1000, 2)}
            return timeout_reply(tool, deadline_seconds, queued, gate_retained=False)
        if not acquired:
            current = self.current_in_flight()
            if current is not None:
                return _busy_error(current)
            # Slot may have freed between timeout and the read — one short retry.
            if not await self._acquire_slot(0.1):
                return _busy_error(
                    InFlight(tool="unknown", database=None, started_at=time.perf_counter())
                )

        attempt = Attempt(next(_call_ids), tool, database, deadline_seconds, entered)
        with self._state_lock:
            self._in_flight = InFlight(
                tool=tool,
                database=database,
                started_at=time.perf_counter(),
                call_id=attempt.call_id,
                attempt=attempt,
            )

        context = contextvars.copy_context()
        body_task: dict[str, Any] = {}
        task_lock = threading.Lock()

        retention: list[Callable[[], None]] = []
        context.run(_retention.set, retention)
        context.run(_attempt_var.set, attempt)
        reply = Future()

        def run() -> Any:
            result = None
            failure = None
            try:
                self._note_apartment_used()
                if is_async:
                    result = self._run_on_apartment_loop(fn, args, kwargs, context, body_task, task_lock)
                else:
                    result = context.run(fn, *args, **kwargs)
                return result
            except BaseException as error:
                failure = error
                raise
            finally:
                try:
                    if retention:
                        if failure is None:
                            reply.set_result(result)
                        else:
                            reply.set_exception(failure)
                        retention.pop()()
                finally:
                    try:
                        if on_worker_finished is not None:
                            on_worker_finished()
                    finally:
                        attempt.settle()
                        if attempt.snapshot()["caller"] == EXPIRED:
                            _record_late_settlement(attempt)

        try:
            worker = self._executor.submit(run)
        except BaseException:
            self._release_slot()  # Submission failed; no worker owns the slot.
            raise
        # The concurrent future finishes only after the body AND Runner cleanup.
        # Its callback needs no caller loop, which may already have closed. It
        # also releases a submitted future cancelled before it starts.
        worker.add_done_callback(lambda _worker: self._release_slot())
        def finish_reply(done):
            if not reply.done():
                try:
                    reply.set_result(done.result())
                except BaseException as error:
                    reply.set_exception(error)
            else:
                done.exception()

        worker.add_done_callback(finish_reply)
        pending = asyncio.wrap_future(reply)
        # A detached worker can fail after the caller leaves. Retrieve its
        # exception even then so the loop never reports an unobserved future.
        pending.add_done_callback(lambda done: None if done.cancelled() else done.exception())
        try:
            if attempt.deadline_at is None:
                result = await asyncio.shield(pending)
            else:
                try:
                    result = await asyncio.wait_for(asyncio.shield(pending), max(0.0, attempt.remaining()))
                except TimeoutError:
                    view = attempt.caller_expired("deadline")
                    if view["phase"] != "settled":
                        return timeout_reply(
                            tool, deadline_seconds, view,
                            gate_retained=True, annotations=attempt.annotations(),
                        )
                    # The body returned at the boundary; its reply follows.
                    result = await asyncio.shield(pending)
            attempt.caller_answered()
            return result
        except asyncio.CancelledError:
            attempt.caller_expired("cancelled")
            if is_async:
                with task_lock:
                    body_task["cancelled"] = True
                    if "task" in body_task:
                        try:
                            body_task["loop"].call_soon_threadsafe(body_task["task"].cancel)
                        except RuntimeError:
                            pass  # The body finished and its loop is closed.
            raise


_gate: AccessGate | None = None
_gate_lock = threading.Lock()


def get_access_gate() -> AccessGate:
    global _gate
    with _gate_lock:
        if _gate is None:
            _gate = AccessGate()
        return _gate


def reset_access_gate() -> None:
    """Reset the module singleton (tests only)."""
    global _gate
    with _gate_lock:
        if _gate is not None:
            _gate._executor.shutdown(wait=False, cancel_futures=True)
        _gate = None


def _record_late_settlement(attempt: Attempt) -> None:
    """Durable record of how a call ended after its caller stopped waiting."""
    try:
        from .usage_logging import log_diagnostic_event

        log_diagnostic_event(
            "attempt_settled_after_expiry",
            tool=attempt.tool, database=attempt.database, **attempt.snapshot(),
        )
    except Exception:
        pass


def _record_abandoned(attempt: Attempt) -> None:
    try:
        from .usage_logging import log_diagnostic_event

        log_diagnostic_event(
            "attempt_abandoned_at_exit",
            tool=attempt.tool, database=attempt.database, **attempt.snapshot(),
        )
    except Exception:
        pass


def _gate_interpreter_exit() -> None:
    """Let the call in flight finish before the interpreter finalizes.

    A call without a deadline (builds, exports, imports) is waited for as
    the old executor join did. A call with one is waited for until its own
    deadline plus ``EXIT_GRACE_SEC``, as the old timeout thread bounded it;
    past that it is recorded as abandoned and nothing is replayed. Hooks run
    newest first, so ``build_owner``'s exit hook has already run.
    """
    gate = _gate
    if gate is None:
        return
    while True:
        current = gate.current_in_flight()
        if current is None:
            return
        attempt = current.attempt
        if (attempt is not None and attempt.deadline_at is not None
                and time.monotonic() > attempt.deadline_at + EXIT_GRACE_SEC):
            _record_abandoned(attempt)
            return
        time.sleep(SLOT_POLL_SEC)


threading._register_atexit(_gate_interpreter_exit)
