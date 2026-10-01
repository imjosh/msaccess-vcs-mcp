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
"""

from __future__ import annotations

import asyncio
import contextvars
import itertools
import os
import threading
import time
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass
from typing import Any, Callable

try:
    import pythoncom

    COM_AVAILABLE = True
except ImportError:
    COM_AVAILABLE = False

DEFAULT_BUSY_WAIT_SEC = 15.0
SLOT_POLL_SEC = 0.02
_call_ids = itertools.count(1)

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


@dataclass(frozen=True)
class InFlight:
    tool: str
    database: str | None
    started_at: float
    # Identifies this call among all others. Interruption records use it to
    # find the call they belong to, so a later call never inherits one.
    call_id: int = 0


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
    return {
        "success": False,
        "error": (
            f"Another Access operation is in progress ({in_flight.tool}{db_hint}). "
            "The MCP server runs one Access operation at a time across all Cursor "
            "windows sharing this server process. Retry after the in-flight call "
            "completes, or call vcs_get_recent_calls() to see what is running."
        ),
        "error_pattern": "server_busy",
        "recoverable": True,
        "busy_with": {
            "tool": in_flight.tool,
            "database": in_flight.database,
            "elapsed_ms": elapsed_ms,
        },
        "retry_after_seconds": 5,
    }


class AccessGate:
    """One COM apartment thread and one in-flight Access operation at a time."""

    def __init__(self) -> None:
        self._state_lock = threading.Lock()
        self._in_flight: InFlight | None = None
        self._slot = threading.Lock()
        self._executor = ThreadPoolExecutor(
            max_workers=1,
            thread_name_prefix="vcs-access-apartment",
            initializer=_init_com_apartment,
        )
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

    async def _run_on_apartment_loop(self, fn: Callable[..., Any], args: tuple, kwargs: dict) -> Any:
        """Run an async body to completion in a fresh event loop on the apartment thread.

        The body keeps the caller's context variables. Cancelling the caller
        cancels the body at its next await; a blocking step it is inside runs on.
        """
        context = contextvars.copy_context()
        body_task: dict[str, Any] = {}

        async def body() -> Any:
            body_task["loop"] = asyncio.get_running_loop()
            body_task["task"] = asyncio.current_task()
            if body_task.get("cancelled"):
                raise asyncio.CancelledError
            return await fn(*args, **kwargs)

        def run() -> Any:
            self._note_apartment_used()
            with asyncio.Runner() as runner:
                return runner.run(body(), context=context)

        try:
            return await asyncio.get_running_loop().run_in_executor(self._executor, run)
        except asyncio.CancelledError:
            body_task["cancelled"] = True
            if "task" in body_task:
                try:
                    body_task["loop"].call_soon_threadsafe(body_task["task"].cancel)
                except RuntimeError:
                    pass  # The body finished and its loop is closed.
            raise

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
        **kwargs: Any,
    ) -> Any:
        if not await self._acquire_slot(_read_busy_wait_sec()):
            current = self.current_in_flight()
            if current is not None:
                return _busy_error(current)
            # Slot may have freed between timeout and the read — one short retry.
            if not await self._acquire_slot(0.1):
                return _busy_error(
                    InFlight(tool="unknown", database=None, started_at=time.perf_counter())
                )

        with self._state_lock:
            self._in_flight = InFlight(
                tool=tool,
                database=database,
                started_at=time.perf_counter(),
                call_id=next(_call_ids),
            )

        try:
            if is_async:
                return await self._run_on_apartment_loop(fn, args, kwargs)

            def _run_sync() -> Any:
                self._note_apartment_used()
                return fn(*args, **kwargs)

            loop = asyncio.get_running_loop()
            return await loop.run_in_executor(self._executor, _run_sync)
        finally:
            with self._state_lock:
                self._in_flight = None
            self._slot.release()


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
