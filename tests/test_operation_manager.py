"""Tests for monotonic MCP progress reporting."""

from __future__ import annotations

import asyncio
import logging
import threading
from unittest.mock import AsyncMock, MagicMock

import pytest

from msaccess_vcs_mcp.operation_manager import (
    MonotonicProgressReporter,
    OperationManager,
)


@pytest.fixture(autouse=True)
def _reset_manager():
    OperationManager._instance = None
    yield
    OperationManager._instance = None


def test_format_message_keeps_vba_counts_in_text():
    assert MonotonicProgressReporter.format_message("queries", 28, 30) == (
        "queries (28/30)"
    )
    assert MonotonicProgressReporter.format_message("", 4, 10) == "(4/10)"
    assert MonotonicProgressReporter.format_message("starting") == "starting"
    assert MonotonicProgressReporter.format_message("log", -1, -1) == "log"
    assert MonotonicProgressReporter.format_message("queries", "28", "30") == (
        "queries (28/30)"
    )


def test_reporter_is_strictly_increasing():
    ctx = MagicMock()
    ctx.report_progress = AsyncMock()
    reporter = MonotonicProgressReporter()

    async def _run():
        await reporter.emit(ctx, message="queries", vba_progress=28, vba_total=30)
        await reporter.emit(ctx, message="modules", vba_progress=1, vba_total=50)
        await reporter.emit(ctx, message="exported forms")

    asyncio.run(_run())

    progresses = [call.kwargs["progress"] for call in ctx.report_progress.await_args_list]
    assert progresses == [1.0, 2.0, 3.0]
    assert all(call.kwargs["total"] is None for call in ctx.report_progress.await_args_list)
    assert ctx.report_progress.await_args_list[0].kwargs["message"] == "queries (28/30)"
    assert ctx.report_progress.await_args_list[1].kwargs["message"] == "modules (1/50)"
    assert ctx.report_progress.await_args_list[2].kwargs["message"] == "exported forms"


def test_wait_for_completion_forwards_ctx_and_stays_monotonic():
    manager = OperationManager.get_instance()
    operation_id, _queue = manager.register_operation(timeout_ms=5000)
    ctx = MagicMock()
    ctx.report_progress = AsyncMock()

    async def _run():
        async def _feed():
            await asyncio.sleep(0)
            manager.route_callback(operation_id, {
                "type": "progress",
                "progress": 28,
                "total": 30,
                "message": "queries",
            })
            manager.route_callback(operation_id, {
                "type": "log",
                "message": "Skipping unchanged object",
            })
            manager.route_callback(operation_id, {
                "type": "progress",
                "progress": 1,
                "total": 50,
                "message": "modules",
            })
            manager.route_callback(operation_id, {
                "type": "complete",
                "message": "done",
                "log_path": r"C:\src\logs\Export_1.log",
            })

        feeder = asyncio.create_task(_feed())
        result = await manager.wait_for_completion(operation_id, ctx=ctx, timeout_seconds=2)
        await feeder
        return result

    result = asyncio.run(_run())
    assert result["success"] is True
    assert result["log_messages"] == ["Skipping unchanged object"]
    progresses = [call.kwargs["progress"] for call in ctx.report_progress.await_args_list]
    assert progresses == [1.0, 2.0, 3.0]
    assert progresses == sorted(progresses)
    assert all(call.kwargs["total"] is None for call in ctx.report_progress.await_args_list)


def test_wait_for_completion_without_ctx_still_collects_logs():
    manager = OperationManager.get_instance()
    operation_id, _queue = manager.register_operation()

    async def _run():
        manager.route_callback(operation_id, {"type": "log", "message": "hello"})
        manager.route_callback(operation_id, {"type": "complete", "message": "ok"})
        return await manager.wait_for_completion(operation_id, ctx=None, timeout_seconds=2)

    result = asyncio.run(_run())
    assert result["success"] is True
    assert result["log_messages"] == ["hello"]


def test_wait_for_completion_forwards_results_path_on_error():
    """Failed test runs post type error but still attach results_path."""
    manager = OperationManager.get_instance()
    operation_id, _queue = manager.register_operation(timeout_ms=5000)

    async def _run():
        async def _feed():
            await asyncio.sleep(0)
            manager.route_callback(operation_id, {
                "type": "error",
                "message": "Operation failed",
                "log_path": r"C:\src\logs\TestRun_1.log",
                "results_path": r"C:\src\logs\TestResults_1.json",
            })

        feeder = asyncio.create_task(_feed())
        result = await manager.wait_for_completion(operation_id, timeout_seconds=2)
        await feeder
        return result

    result = asyncio.run(_run())
    assert result["success"] is False
    assert result["error"] == "Operation failed"
    assert result["log_path"] == r"C:\src\logs\TestRun_1.log"
    assert result["results_path"] == r"C:\src\logs\TestResults_1.json"


@pytest.mark.parametrize("msg_type", ["error", "complete", "cancelled"])
def test_wait_for_completion_forwards_runtime_error(msg_type):
    """The add-in's error text reaches the async caller only on the callback."""
    manager = OperationManager.get_instance()
    operation_id, _queue = manager.register_operation(timeout_ms=5000)

    async def _run():
        async def _feed():
            await asyncio.sleep(0)
            manager.route_callback(operation_id, {
                "type": msg_type,
                "message": "Division by zero",
                "runtime_error": "Division by zero",
                "errorNumber": 11,
            })

        feeder = asyncio.create_task(_feed())
        result = await manager.wait_for_completion(operation_id, timeout_seconds=2)
        await feeder
        return result

    result = asyncio.run(_run())
    assert result["runtime_error"] == "Division by zero"
    assert result["errorNumber"] == 11


def test_wait_for_completion_omits_runtime_error_when_absent():
    manager = OperationManager.get_instance()
    operation_id, _queue = manager.register_operation(timeout_ms=5000)

    async def _run():
        async def _feed():
            await asyncio.sleep(0)
            manager.route_callback(operation_id, {"type": "error", "message": "Operation failed"})

        feeder = asyncio.create_task(_feed())
        result = await manager.wait_for_completion(operation_id, timeout_seconds=2)
        await feeder
        return result

    result = asyncio.run(_run())
    assert "runtime_error" not in result
    assert "errorNumber" not in result


@pytest.mark.parametrize("terminal_type", ["complete", "error", "cancelled"])
def test_progress_delivery_failure_preserves_operation_outcome(terminal_type, caplog):
    """Notification failures stay diagnostic and do not stop callback processing."""
    manager = OperationManager.get_instance()
    operation_id, _queue = manager.register_operation(timeout_ms=5000)
    ctx = MagicMock()
    ctx.report_progress = AsyncMock(side_effect=[RuntimeError("client disconnected"), None])
    caplog.set_level(logging.DEBUG, logger="msaccess_vcs_mcp.operation_manager")

    async def _run():
        manager.route_callback(operation_id, {
            "type": "progress", "progress": 28, "total": 30, "message": "queries",
        })
        manager.route_callback(operation_id, {"type": "log", "message": "exporting modules"})
        manager.route_callback(operation_id, {
            "type": terminal_type,
            "message": "finished",
            "log_path": r"C:\src\logs\Export_1.log",
        })
        return await manager.wait_for_completion(operation_id, ctx=ctx, timeout_seconds=2)

    result = asyncio.run(_run())

    assert result["success"] is (terminal_type == "complete")
    assert result["log_path"] == r"C:\src\logs\Export_1.log"
    if terminal_type == "complete":
        assert result["log_messages"] == ["exporting modules"]
    elif terminal_type == "error":
        assert result["error"] == "finished"
    else:
        assert result["cancelled"] is True
    assert manager.pending_count() == 0
    assert [call.kwargs["progress"] for call in ctx.report_progress.await_args_list] == [1.0, 2.0]
    assert all(call.kwargs["total"] is None for call in ctx.report_progress.await_args_list)
    failure = next(record for record in caplog.records if "Failed to report progress" in record.message)
    assert failure.levelno == logging.WARNING
    assert failure.exc_info is not None
    assert "Progress reported: 2.0 - exporting modules" in caplog.text


def test_callbacks_reach_the_loop_that_registered_the_operation():
    """A gated tool waits on the Access apartment's loop while the manager keeps the server's."""
    manager = OperationManager.get_instance()
    server_loop = asyncio.new_event_loop()
    manager.set_event_loop(server_loop)  # As at server startup.
    registered = threading.Event()
    box: dict = {}

    async def gated_body():
        operation_id, _queue = manager.register_operation(command="Export")
        box["operation_id"] = operation_id
        registered.set()
        return await manager.wait_for_completion(operation_id, timeout_seconds=2)

    def apartment():
        box["result"] = asyncio.run(gated_body())

    thread = threading.Thread(target=apartment)
    try:
        thread.start()
        assert registered.wait(2)
        # The HTTP handler thread routes the terminal callback.
        assert manager.route_callback(box["operation_id"], {"type": "complete", "message": "done"})
        thread.join(3)
        assert box["result"]["success"] is True
        assert box["result"]["message"] == "done"
    finally:
        thread.join(3)
        server_loop.close()


def test_callback_after_its_loop_closed_is_not_routed():
    manager = OperationManager.get_instance()

    async def register():
        return manager.register_operation(command="Export")[0]

    operation_id = asyncio.run(register())
    assert manager.route_callback(operation_id, {"type": "progress", "message": "late"}) is False
