"""Bounded worker threads for sync gate-exempt tools.

A gate-exempt tool (the dialog tools, ``vcs_get_recent_calls``,
``vcs_cancel_operation``) must answer while Access is blocked, so its body runs
off the event loop and off the COM apartment thread. Cancelling an await does
not stop a thread, and a Win32 or MSAA call into a hung Access UI thread can
outlive the response that started it.

So these workers stay out of the asyncio default executor, which other waits
share, and each tool has a fixed budget: at most ``MAX_WORKERS_PER_TOOL``
threads alive, abandoned or not, and at most ``MAX_WAITING_PER_TOOL`` calls
waiting for one of them. Repeated hangs therefore cannot grow the thread count
or a backlog, and hung dialog inspections never use up another tool's budget.
One deadline covers the wait for a worker and the run: past it, a call that
never started raises ``WorkerCapacityUnavailable`` and one that started raises
``asyncio.TimeoutError``. A full waiting line refuses at once.
"""

from __future__ import annotations

import asyncio
import threading
import time
from typing import Any, Callable

MAX_WORKERS_PER_TOOL = 2
MAX_WAITING_PER_TOOL = 2
ADMISSION_POLL_SEC = 0.02


class WorkerCapacityUnavailable(Exception):
    """No worker of this tool became free before the deadline, or the waiting line was full."""

    def __init__(self, running: int) -> None:
        super().__init__(f"{running} workers still running")
        self.running = running


def _deliver(future: asyncio.Future, value: Any, error: BaseException | None) -> None:
    if future.done():
        # The call was abandoned at its deadline.
        return
    if error is not None:
        future.set_exception(error)
    else:
        future.set_result(value)


class ExemptWorkers:
    """The worker budget of one gate-exempt tool."""

    def __init__(
        self,
        name: str,
        max_workers: int = MAX_WORKERS_PER_TOOL,
        max_waiting: int = MAX_WAITING_PER_TOOL,
    ) -> None:
        self._name = name
        self._max_workers = max_workers
        self._max_waiting = max_waiting
        self._lock = threading.Lock()
        self._running = 0
        self._waiting = 0

    @property
    def running(self) -> int:
        """Threads started and not yet returned, including abandoned ones."""
        with self._lock:
            return self._running

    async def run(
        self, fn: Callable[..., Any], args: tuple, kwargs: dict, deadline: float
    ) -> Any:
        """``fn(*args, **kwargs)`` in a worker thread, answered by ``deadline`` (monotonic).

        An abandoned thread keeps its place in the budget until it returns.
        """
        await self._admit(deadline)
        loop = asyncio.get_running_loop()
        future = loop.create_future()
        thread = threading.Thread(
            target=self._work,
            args=(loop, future, fn, args, kwargs),
            name=f"vcs-exempt-{self._name}",
            daemon=True,
        )
        try:
            thread.start()
        except BaseException:
            self._release()
            raise
        return await asyncio.wait_for(future, max(deadline - time.monotonic(), 0.0))

    def _try_admit(self, queued: bool) -> bool:
        """Caller holds ``_lock``. A new call never overtakes one already waiting."""
        if self._running >= self._max_workers or (not queued and self._waiting):
            return False
        self._running += 1
        return True

    async def _admit(self, deadline: float) -> None:
        with self._lock:
            if self._try_admit(queued=False):
                return
            if self._waiting >= self._max_waiting:
                raise WorkerCapacityUnavailable(self._running)
            self._waiting += 1
        try:
            while True:
                remaining = deadline - time.monotonic()
                with self._lock:
                    if self._try_admit(queued=True):
                        return
                    if remaining <= 0:
                        raise WorkerCapacityUnavailable(self._running)
                # Polling needs no thread; a thread pool here could itself be full.
                await asyncio.sleep(min(ADMISSION_POLL_SEC, remaining))
        finally:
            with self._lock:
                self._waiting -= 1

    def _release(self) -> None:
        with self._lock:
            self._running -= 1

    def _work(self, loop, future, fn, args, kwargs) -> None:
        try:
            value, error = fn(*args, **kwargs), None
        except BaseException as exc:
            value, error = None, exc
        finally:
            self._release()
        try:
            loop.call_soon_threadsafe(_deliver, future, value, error)
        except RuntimeError:
            # The loop closed after the call was abandoned.
            pass


_budgets: dict[str, ExemptWorkers] = {}
_budgets_lock = threading.Lock()


def workers_for(name: str) -> ExemptWorkers:
    with _budgets_lock:
        if name not in _budgets:
            _budgets[name] = ExemptWorkers(name)
        return _budgets[name]


def wait_idle(timeout: float) -> bool:
    """Wait until no worker of any tool is running (tests only). False on timeout."""
    deadline = time.monotonic() + timeout
    while True:
        with _budgets_lock:
            budgets = list(_budgets.values())
        if all(budget.running == 0 for budget in budgets):
            return True
        if time.monotonic() >= deadline:
            return False
        time.sleep(ADMISSION_POLL_SEC)
