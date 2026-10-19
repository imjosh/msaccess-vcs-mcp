"""One mutation retains its callback identity until terminal evidence."""

import asyncio
import json
import os
import time
from contextlib import contextmanager
from unittest.mock import AsyncMock, MagicMock, patch

import pytest

from msaccess_vcs_mcp import tools
from msaccess_vcs_mcp.operation_manager import OperationManager
from tests.test_export_build_fallbacks import FakeAccess, _addin, _build_env, _export_env, _unwrap
from tests.test_scoped_sync import _patch_import_tool


@contextmanager
def route(tmp_path, name, manager, start):
    if name in {"merge", "tests"}:
        with _patch_import_tool(tmp_path, async_result=start if isinstance(start, dict) else None) as (db, src, addin, _):
            addin.call_sync.return_value = json.dumps({"success": True})
            if isinstance(start, Exception):
                addin.call_async.side_effect = start
            with patch.object(tools, "_get_operation_manager", return_value=manager):
                call = (lambda: _unwrap(tools.vcs_import_objects)(db, str(src))) if name == "merge" else (lambda: _unwrap(tools.vcs_run_tests)(db))
                yield call, db, addin
        return
    fake = FakeAccess(async_start=json.dumps(start) if isinstance(start, dict) else start)
    fake.export_folder = str(tmp_path / "export") + os.sep
    addin = _addin(tmp_path, fake)
    if name == "export":
        with _export_env(tmp_path, addin, manager, True) as (db, out):
            yield lambda: _unwrap(tools.vcs_export_database)(str(db), str(out)), str(db), addin
    else:
        from tests.build_fakes import owned_build
        app = MagicMock()
        with owned_build(app) as (gate, gone), _build_env(tmp_path, addin, manager, True, app) as src:
            out = str(tmp_path / "output.accdb")
            yield lambda: gate.run_exclusive("build", out, _unwrap(tools.vcs_rebuild_database), True, str(src), out), out, addin


@pytest.mark.parametrize("name", ["merge", "export", "build", "tests"])
@pytest.mark.parametrize("failure", ["wait", "lost_reply", "unknown_marker"])
def test_uncertain_operation_keeps_cancel_and_blocks_admission_until_late_terminal(tmp_path, monkeypatch, name, failure):
    manager = OperationManager()
    start = RuntimeError("reply lost after dispatch") if failure == "lost_reply" else {"unexpected": True} if failure == "unknown_marker" else {"async": True}
    with route(tmp_path, name, manager, start) as (invoke, database, addin):
        if failure == "wait":
            manager.wait_for_completion = AsyncMock(side_effect=RuntimeError("wait failed"))
        result = asyncio.run(invoke())
        assert result["success"] is False
        assert result["completion_unconfirmed"] is True
        assert result.get("started") is (True if failure == "wait" else None)
        operation_id = result["operation_id"]
        assert manager.is_database_busy(database)
        assert manager.request_cancel(operation_id)
        assert manager.is_cancelled(operation_id)
        if name in {"merge", "tests"}:
            addin.call_async.assert_called_once()
            addin.merge_build.assert_not_called()
            assert not any(c.args[0] == "RunFilteredTests" for c in addin.call_sync.call_args_list)
        else:
            assert addin._app.Run.call_count == (2 if name == "export" else 1)

        monkeypatch.setattr(tools, "_get_operation_manager", lambda: manager)
        assert manager.get_busy_status(database)["operation_id"] == operation_id
        assert manager.route_callback(operation_id, {"type": "cancelled", "message": "late terminal"})
        deadline = time.monotonic() + 2
        while manager.is_database_busy(database) and time.monotonic() < deadline:
            time.sleep(0.01)
        assert not manager.is_database_busy(database)
        assert not manager.is_cancelled(operation_id)
        assert manager.get_busy_status(database) is None
        next_id, _ = manager.register_operation(database_path=database)
        manager.unregister_operation(next_id)


def test_timeout_and_cancelled_wait_keep_tracking_after_loop_closes():
    manager = OperationManager()

    async def wait():
        operation_id, _ = manager.register_operation(database_path="timeout.accdb")
        result = await manager.wait_for_completion(operation_id, timeout_seconds=0.001)
        return operation_id, result

    operation_id, result = asyncio.run(wait())
    assert result["operation_id"] == operation_id
    assert manager.is_database_busy("timeout.accdb")
    assert manager.route_callback(operation_id, {"type": "complete"})
    assert not manager.is_database_busy("timeout.accdb")

    async def cancelled_wait():
        operation_id, _ = manager.register_operation(database_path="cancelled.accdb")
        task = asyncio.create_task(manager.wait_for_completion(operation_id))
        await asyncio.sleep(0)
        task.cancel()
        with pytest.raises(asyncio.CancelledError):
            await task
        return operation_id

    operation_id = asyncio.run(cancelled_wait())
    assert manager.request_cancel(operation_id)
    assert manager.route_callback(operation_id, {"type": "cancelled"})
    assert not manager.is_database_busy("cancelled.accdb")


def test_known_predispatch_failure_releases_reservation_without_dispatch(tmp_path, monkeypatch):
    from msaccess_vcs_mcp import addin_integration
    from msaccess_vcs_mcp.operation_manager import dispatch_once

    fake = FakeAccess()
    addin = _addin(tmp_path, fake)
    monkeypatch.setattr(addin_integration, "ensure_session", lambda *args: (_ for _ in ()).throw(RuntimeError("session setup failed")))
    manager = OperationManager()
    operation_id, _ = manager.register_operation(database_path="fixture.accdb")
    outcome = asyncio.run(dispatch_once(addin, manager, operation_id, "{}", "Export"))
    assert outcome.kind == "refused"
    assert outcome.payload["started"] is False
    assert outcome.payload["error_pattern"] == "async_not_started"
    assert fake.dispatched == []
    assert not manager.is_database_busy("fixture.accdb")


def test_duplicate_registration_cannot_replace_original_reservation():
    manager = OperationManager()
    operation_id, _ = manager.register_operation(database_path="fixture.accdb")
    with pytest.raises(RuntimeError, match="unresolved"):
        manager.register_operation(database_path="FIXTURE.accdb")
    assert manager.get_busy_status("fixture.accdb")["operation_id"] == operation_id


def test_terminal_callback_reconciles_registration_after_loop_closed():
    manager = OperationManager()

    async def register():
        return manager.register_operation(database_path="fixture.accdb")[0]

    operation_id = asyncio.run(register())
    assert manager.route_callback(operation_id, {"type": "complete", "message": "done"})
    assert not manager.is_database_busy("fixture.accdb")


@pytest.mark.parametrize("tool, args", [
    (tools.vcs_export_object, ("fixture.accdb", "Module", "Example")),
    (tools.vcs_import_object, ("fixture.accdb", "Module", "Example", "example.bas")),
    (tools.vcs_execute_sql, ("fixture.accdb", "DELETE FROM Example")),
    (tools.vcs_call_vba, ("fixture.accdb", "Example")),
])
def test_public_gated_tools_refuse_retained_database_before_connecting(monkeypatch, tool, args):
    manager = OperationManager()
    operation_id, _ = manager.register_operation(database_path="fixture.accdb")
    monkeypatch.setattr(tools, "_get_operation_manager", lambda: manager)
    monkeypatch.setattr(tools, "_ensure_env_loaded", AsyncMock())
    monkeypatch.setattr(tools, "load_config", lambda: {})
    with patch.object(tools, "AccessConnection") as connection:
        result = asyncio.run(tool(*args))
    assert result["active_operation_id"] == operation_id
    connection.assert_not_called()
