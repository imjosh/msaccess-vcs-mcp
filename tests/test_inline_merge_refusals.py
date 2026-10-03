"""M52: inline full-merge refusals never replace an existing owner.

Real decorated wrappers, gate, integration parser and callback manager; fake COM.
No busy live Access call is used. Independent expected wires come from M48's
shared fixture. The existing owner's queue is distinct from the refused caller.
"""
import asyncio
import json
import os
import time
from unittest.mock import AsyncMock, Mock, patch

import pytest

from msaccess_vcs_mcp import tools
from msaccess_vcs_mcp.access_gate import AccessGate
from msaccess_vcs_mcp.operation_manager import OperationManager
from .test_export_build_fallbacks import FakeAccess, _addin, _export_env


@pytest.fixture
def boundary(monkeypatch):
    gate = AccessGate()
    monkeypatch.setattr(tools, "get_access_gate", lambda: gate)
    monkeypatch.setattr(tools, "_ensure_env_loaded", AsyncMock())
    monkeypatch.setattr(tools, "load_config", lambda: {})
    monkeypatch.setenv("ACCESS_VCS_ENABLE_LOGGING", "false")
    monkeypatch.setenv("ACCESS_VCS_DISABLE_DIAGNOSTIC_LOG", "true")
    yield
    gate._executor.shutdown(wait=True)


@pytest.mark.parametrize("nested_dict", [False, True])
def test_refusal_keeps_existing_owner_and_never_borrows_its_log(
    boundary, tmp_path, addin_start_refusal, nested_dict,
):
    expected = addin_start_refusal["expected"]
    raw = addin_start_refusal["raw"]
    if expected.get("decision_required"):
        raw = json.dumps({**json.loads(raw), "started": False})
    manager = OperationManager()
    owner_id, owner_queue = manager.register_operation(database_path=str(tmp_path / "owner.accdb"), command="Export")
    owner = manager.get_operation(owner_id)
    own_log = tmp_path / "original-owner.log"
    own_log.write_text("Original operation only", encoding="utf-8")
    owner_journal = [{"kind": "acknowledged", "title": "original owner journal", "resolution": "acknowledged"}]
    manager.unregister_operation = Mock(wraps=manager.unregister_operation)
    manager.wait_for_completion = AsyncMock(wraps=manager.wait_for_completion)
    refused_ids = []

    def post(info):
        request_id = json.loads(info)["operation_id"]
        refused_ids.append(request_id)

    if nested_dict and raw.startswith("{"):
        raw = json.loads(raw)
    start = json.dumps({"sync": True, "result": raw})
    fake = FakeAccess(api=raw, async_start=start, on_async=post)
    fake.export_folder = str(tmp_path / "export") + os.sep
    addin = _addin(tmp_path, fake)

    def old_logs(src):
        for name in ("Export.log", "Build.log", "Merge.log"):
            log = src / name
            log.write_text("Old unrelated legacy run", encoding="utf-8")
            os.utime(log, (1, 1))

    with _export_env(tmp_path, addin, manager, True) as (db, src):
        old_logs(src)
        with patch.object(tools, "validate_source_directory", return_value=src):
            result = asyncio.run(tools.vcs_import_objects(str(db), str(src)))
    mismatches = []
    for key, value in expected.items():
        if key != "api_refused":  # The public import shape exposes error/pattern.
            if result.get(key) != value:
                mismatches.append({"field": key, "expected": value, "actual": result.get(key)})
    assert fake.dispatched == [("APIAsync", "MergeBuild")]
    assert result["log_path"] is None and "log_excerpt" not in result
    assert not result.get("completion_unconfirmed")
    assert manager.get_operation(owner_id) is owner
    assert not owner.finished and not manager.is_cancelled(owner_id)
    assert owner_queue.empty()
    assert manager.pending_count() == 1
    request_id, = refused_ids
    assert request_id != owner_id
    assert not manager.request_cancel(request_id)
    assert not manager.is_cancelled(owner_id)
    late = {**expected, "type": "error", "message": expected["error"]}
    assert not manager.route_callback(request_id, late)
    assert not manager.route_callback(request_id, late)
    manager.unregister_operation.assert_called_once_with(request_id)
    manager.wait_for_completion.assert_not_awaited()
    assert manager.route_callback(owner_id, {"type": "complete", "success": True, "operation_id": owner_id,
        "decisions": owner_journal, "log_path": str(own_log)})
    original = asyncio.run(manager.wait_for_completion(owner_id))
    assert original["success"] is True and original["decisions"] == owner_journal
    assert original["log_path"] == str(own_log)
    assert manager.pending_count() == 0
    assert not mismatches, mismatches


@pytest.mark.parametrize("error,pattern", [
    ("Another API command is still running.", "operation_already_running"),
    ("The call arrived back in the project that sent it.", "api_self_dispatch"),
])
@pytest.mark.parametrize("wrapped", [False, True])
def test_inline_merge_refusal_cannot_borrow_a_current_owner_log(boundary, tmp_path, error, pattern, wrapped):
    marked = "VCS_API_REFUSED: " + error
    raw = json.dumps({"success": False, "error": marked}) if wrapped else marked
    fake = FakeAccess(async_start=json.dumps({"sync": True, "result": raw}))
    fake.export_folder = str(tmp_path / "export") + os.sep
    manager = OperationManager()
    manager.wait_for_completion = AsyncMock(wraps=manager.wait_for_completion)
    with _export_env(tmp_path, _addin(tmp_path, fake), manager, True) as (db, src):
        legacy = src / "Merge.log"
        legacy.write_text("Old legacy run", encoding="utf-8")
        os.utime(legacy, (1, 1))
        own_log = src / "logs/Merge_20261002_120000_000.log"
        own_log.parent.mkdir()
        own_log.write_text("Another operation's current log", encoding="utf-8")
        # Deterministic >= call-start candidate; it must still be rejected for a refusal.
        os.utime(own_log, (time.time() + 10, time.time() + 10))
        with patch.object(tools, "validate_source_directory", return_value=src):
            result = asyncio.run(tools.vcs_import_objects(str(db), str(src)))
    assert fake.dispatched == [("APIAsync", "MergeBuild")]
    assert manager.pending_count() == 0
    manager.wait_for_completion.assert_not_awaited()
    expected = {"success": False, "error": error, "error_pattern": pattern, "log_path": None}
    mismatches = [{"field": k, "expected": v, "actual": result.get(k)} for k, v in expected.items() if result.get(k) != v]
    if "log_excerpt" in result:
        mismatches.append({"field": "log_excerpt", "expected": "absent", "actual": result["log_excerpt"]})
    assert not mismatches, mismatches
