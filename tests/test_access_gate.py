"""Tests for the Access operation gate."""

from __future__ import annotations

import asyncio
import contextvars
import threading
import time
from unittest.mock import patch

import pytest

from msaccess_vcs_mcp.access_gate import (
    AccessGate,
    EXEMPT_TOOLS,
    InFlight,
    ServerLoopContext,
    _busy_error,
    reset_access_gate,
)


@pytest.fixture(autouse=True)
def _reset_gate():
    reset_access_gate()
    yield
    reset_access_gate()


def test_run_exclusive_executes_sync_fn_in_apartment():
    gate = AccessGate()
    seen: list[str] = []

    def work():
        seen.append("done")
        return 42

    result = asyncio.run(
        gate.run_exclusive("vcs_test", r"C:\db.accdb", work, False)
    )
    assert result == 42
    assert seen == ["done"]


async def _until(event: threading.Event) -> None:
    """Wait for an event set on another thread without using a worker thread."""
    while not event.is_set():
        await asyncio.sleep(0.005)


def test_run_exclusive_executes_async_fn_in_apartment():
    """An async body runs in its own loop on the apartment thread, never on the caller's."""
    gate = AccessGate()
    seen: dict[str, object] = {}

    async def work():
        await asyncio.sleep(0)
        seen["thread"] = threading.current_thread().name
        seen["loop"] = asyncio.get_running_loop()
        return "async"

    async def runner():
        result = await gate.run_exclusive("vcs_test", None, work, True)
        return result, asyncio.get_running_loop()

    result, caller_loop = asyncio.run(runner())
    assert result == "async"
    assert str(seen["thread"]).startswith("vcs-access-apartment")
    assert seen["loop"] is not caller_loop


def test_async_body_blocking_step_leaves_caller_loop_free():
    gate = AccessGate()
    entered = threading.Event()
    release = threading.Event()

    async def work():
        entered.set()
        release.wait(5)  # A blocking COM step, such as connect or the add-in probe.
        return "done"

    async def runner():
        task = asyncio.create_task(gate.run_exclusive("vcs_test", None, work, True))
        await asyncio.wait_for(_until(entered), 1.0)
        ticks = 0
        started = time.monotonic()
        while time.monotonic() - started < 0.2:
            await asyncio.sleep(0.01)
            ticks += 1
        release.set()
        return ticks, await asyncio.wait_for(task, 1.0)

    ticks, result = asyncio.run(runner())
    assert result == "done"
    assert ticks >= 5


def test_async_body_keeps_caller_context_variables():
    marker = contextvars.ContextVar("marker", default="unset")
    gate = AccessGate()

    async def work():
        return marker.get()

    async def runner():
        marker.set("caller")
        return await gate.run_exclusive("vcs_test", None, work, True)

    assert asyncio.run(runner()) == "caller"


def test_cancelling_the_caller_cancels_the_async_body():
    gate = AccessGate()
    entered = threading.Event()
    body_cancelled = threading.Event()

    async def work():
        entered.set()
        try:
            await asyncio.sleep(5)
        except asyncio.CancelledError:
            body_cancelled.set()
            raise

    async def runner():
        task = asyncio.create_task(gate.run_exclusive("vcs_test", None, work, True))
        await asyncio.wait_for(_until(entered), 1.0)
        task.cancel()
        with pytest.raises(asyncio.CancelledError):
            await task
        await asyncio.wait_for(_until(body_cancelled), 1.0)

    asyncio.run(runner())


@pytest.mark.parametrize("is_async", [False, True], ids=["sync", "async"])
def test_cancelled_worker_releases_after_caller_loop_closes(is_async):
    """Ownership and exception handling outlive the loop that submitted work."""
    gate = AccessGate()
    entered = threading.Event()
    release = threading.Event()

    def blocking_step():
        entered.set()
        assert release.wait(5)
        raise RuntimeError("detached worker failed")

    async def async_work():
        blocking_step()

    async def caller():
        work = async_work if is_async else blocking_step
        task = asyncio.create_task(gate.run_exclusive("original", None, work, is_async))
        await asyncio.wait_for(_until(entered), 1.0)
        owner = gate.current_in_flight()
        task.cancel()
        with pytest.raises(asyncio.CancelledError):
            await task
        assert gate.current_in_flight() is owner
        return owner

    try:
        owner = asyncio.run(caller())  # This loop closes while the worker is blocked.
        assert gate.current_in_flight() is owner
    finally:
        release.set()
        gate._executor.shutdown(wait=True)
    assert gate.current_in_flight() is None
    assert gate._slot.acquire(blocking=False)
    gate._slot.release()


def test_cancelled_async_worker_keeps_slot_through_runner_cleanup():
    """Runner shutdown must finish a child's async cancellation cleanup first."""
    gate = AccessGate()
    entered = threading.Event()
    cleanup = threading.Event()
    release = threading.Event()

    async def child():
        try:
            entered.set()
            await asyncio.sleep(5)
        finally:
            cleanup.set()
            await _until(release)

    async def work():
        asyncio.create_task(child())
        await asyncio.sleep(5)

    async def scenario():
        task = asyncio.create_task(gate.run_exclusive("original", None, work, True))
        try:
            await asyncio.wait_for(_until(entered), 1.0)
            owner = gate.current_in_flight()
            task.cancel()
            with pytest.raises(asyncio.CancelledError):
                await task
            await asyncio.wait_for(_until(cleanup), 1.0)
            assert gate.current_in_flight() is owner
            with patch("msaccess_vcs_mcp.access_gate._read_busy_wait_sec", return_value=0.05):
                busy = await asyncio.wait_for(
                    gate.run_exclusive("later", None, lambda: "later", False), 0.5,
                )
            assert busy["error_pattern"] == "server_busy"
            assert busy["busy_with"]["tool"] == "original"
            release.set()
            assert await gate.run_exclusive("later", None, lambda: "later", False) == "later"
            assert gate.current_in_flight() is None
        finally:
            release.set()
            await asyncio.gather(task, return_exceptions=True)
            gate._executor.shutdown(wait=True)

    asyncio.run(scenario())


def test_submission_failure_releases_slot():
    gate = AccessGate()
    try:
        with patch.object(gate._executor, "submit", side_effect=RuntimeError("cannot submit")):
            with pytest.raises(RuntimeError, match="cannot submit"):
                asyncio.run(gate.run_exclusive("original", None, lambda: None, False))
        assert gate.current_in_flight() is None
        assert asyncio.run(gate.run_exclusive("later", None, lambda: "later", False)) == "later"
    finally:
        gate._executor.shutdown(wait=True)


def test_server_loop_context_reports_progress_on_the_server_loop():
    gate = AccessGate()
    reports: list[tuple[object, dict]] = []

    class FakeContext:
        async def report_progress(self, **kwargs):
            reports.append((asyncio.get_running_loop(), kwargs))

    async def work(ctx):
        await ctx.report_progress(progress=1.0, message="Exporting")
        return "done"

    async def runner():
        loop = asyncio.get_running_loop()
        ctx = ServerLoopContext(FakeContext(), loop)
        return await gate.run_exclusive("vcs_test", None, work, True, ctx), loop

    result, server_loop = asyncio.run(runner())
    assert result == "done"
    assert reports == [(server_loop, {"progress": 1.0, "total": None, "message": "Exporting"})]


def test_run_exclusive_serializes_concurrent_calls():
    gate = AccessGate()
    order: list[int] = []

    async def runner():
        async def first():
            order.append(1)
            await asyncio.sleep(0.05)
            order.append(2)
            return 1

        async def second():
            order.append(3)
            return 2

        with patch("msaccess_vcs_mcp.access_gate._read_busy_wait_sec", return_value=2.0):
            t1 = asyncio.create_task(
                gate.run_exclusive("first", None, first, True)
            )
            await asyncio.sleep(0.01)
            t2 = asyncio.create_task(
                gate.run_exclusive("second", None, second, True)
            )
            return await asyncio.gather(t1, t2)

    r1, r2 = asyncio.run(runner())
    assert r1 == 1
    assert r2 == 2
    assert order == [1, 2, 3]


def test_run_exclusive_returns_busy_when_slot_unavailable():
    gate = AccessGate()
    entered = threading.Event()

    async def runner():
        async def slow():
            entered.set()
            await asyncio.sleep(0.3)
            return "slow"

        with patch("msaccess_vcs_mcp.access_gate._read_busy_wait_sec", return_value=0.05):
            slow_task = asyncio.create_task(
                gate.run_exclusive("vcs_run_tests", r"C:\big.accdb", slow, True)
            )
            await _until(entered)
            busy = await gate.run_exclusive(
                "vcs_call_vba", r"C:\other.accdb", lambda: None, False
            )
            await slow_task
            return busy

    busy = asyncio.run(runner())
    assert busy["success"] is False
    assert busy["error_pattern"] == "server_busy"
    assert busy["recoverable"] is True
    assert busy["busy_with"]["tool"] == "vcs_run_tests"
    assert busy["busy_with"]["database"] == r"C:\big.accdb"


def test_run_exclusive_releases_slot_after_exception():
    gate = AccessGate()

    def boom():
        raise RuntimeError("fail")

    with pytest.raises(RuntimeError, match="fail"):
        asyncio.run(gate.run_exclusive("vcs_test", None, boom, False))

    result = asyncio.run(
        gate.run_exclusive("vcs_test", None, lambda: "ok", False)
    )
    assert result == "ok"


@pytest.mark.parametrize("slot_free", [True, False], ids=["free", "busy"])
def test_slot_wait_needs_no_worker_thread(slot_free):
    """A default executor full of hung workers cannot delay the slot or server_busy."""
    from concurrent.futures import ThreadPoolExecutor

    gate = AccessGate()
    hung = threading.Event()

    async def runner():
        loop = asyncio.get_running_loop()
        loop.set_default_executor(ThreadPoolExecutor(max_workers=1))
        holder = None
        try:
            with patch("msaccess_vcs_mcp.access_gate._read_busy_wait_sec", return_value=0.1):
                if not slot_free:
                    entered = threading.Event()

                    async def hold():
                        entered.set()
                        await asyncio.sleep(1.0)

                    holder = asyncio.create_task(gate.run_exclusive("vcs_run_tests", None, hold, True))
                    await _until(entered)
                loop.run_in_executor(None, hung.wait, 5)
                started = time.monotonic()
                result = await asyncio.wait_for(
                    gate.run_exclusive("vcs_test", None, lambda: "ran", False), 1.0
                )
                return result, time.monotonic() - started
        finally:
            hung.set()
            if holder is not None:
                holder.cancel()

    result, elapsed = asyncio.run(runner())
    if slot_free:
        assert result == "ran"
    else:
        assert result["error_pattern"] == "server_busy"
        assert result["busy_with"]["tool"] == "vcs_run_tests"
    assert elapsed < 0.6


def test_busy_error_shape():
    err = _busy_error(InFlight("vcs_export_database", r"C:\db.accdb", time.perf_counter()))
    assert err["error_pattern"] == "server_busy"
    assert err["busy_with"]["tool"] == "vcs_export_database"
    assert "retry_after_seconds" in err


def test_exempt_tools_include_status_queries():
    assert "vcs_get_version_info" in EXEMPT_TOOLS
    assert "vcs_get_recent_calls" in EXEMPT_TOOLS
    assert "vcs_cancel_operation" in EXEMPT_TOOLS
    assert "vcs_rebuild_addin" in EXEMPT_TOOLS
    assert "vcs_list_dialogs" in EXEMPT_TOOLS
    assert "vcs_dismiss_dialog" in EXEMPT_TOOLS
    assert "vcs_recover_dialogs" in EXEMPT_TOOLS
    assert "vcs_automation_status" in EXEMPT_TOOLS


def test_com_initializer_runs_for_sync_work():
    gate = AccessGate()
    asyncio.run(gate.run_exclusive("vcs_test", None, lambda: None, False))
    assert gate.com_initialized is True
