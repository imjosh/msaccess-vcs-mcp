"""M49: synchronous full merges retain dispatcher refusals at the public boundary."""

import asyncio
import json
import os
from unittest.mock import AsyncMock, Mock, patch

import pytest

from msaccess_vcs_mcp import tools
from msaccess_vcs_mcp.access_gate import AccessGate
from msaccess_vcs_mcp.operation_manager import OperationManager

from .test_export_build_fallbacks import FakeAccess, _addin, _export_env


@pytest.fixture
def public_merge(tmp_path, monkeypatch):
    gate = AccessGate()
    monkeypatch.setattr(tools, "get_access_gate", lambda: gate)
    monkeypatch.setattr(tools, "_ensure_env_loaded", AsyncMock())
    monkeypatch.setattr(tools, "load_config", lambda: {})
    monkeypatch.setenv("ACCESS_VCS_ENABLE_LOGGING", "false")
    monkeypatch.setenv("ACCESS_VCS_DISABLE_DIAGNOSTIC_LOG", "true")

    def run(raw, route, tuple_return):
        fake = FakeAccess(
            api=(raw, None) if tuple_return else raw,
            async_start=RuntimeError("async unavailable") if route == "exception" else '{}',
        )
        manager = OperationManager()
        manager.unregister_operation = Mock(wraps=manager.unregister_operation)
        manager.wait_for_completion = AsyncMock(wraps=manager.wait_for_completion)
        callback = route != "no-callback"
        fake.export_folder = str(tmp_path / "export") + os.sep
        with (
            _export_env(tmp_path, _addin(tmp_path, fake), manager, callback) as (db, src),
            patch.object(tools, "validate_source_directory", return_value=src),
        ):
            # A previous operation's logs must never be attributed to this refusal.
            for name in ("Build.log", "Merge.log"):
                log = src / name
                log.write_text("Unrelated earlier operation", encoding="utf-8")
                os.utime(log, (1, 1))
            result = asyncio.run(tools.vcs_import_objects(str(db), str(src)))
        assert manager.pending_count() == 0
        manager.wait_for_completion.assert_not_awaited()
        if callback:
            manager.unregister_operation.assert_called_once()
        assert fake.dispatched == (
            ([] if not callback else [("APIAsync", "MergeBuild")]) + [("API", "MergeBuild")]
        )
        assert fake.arguments[-1] == ("block",)
        return result

    yield run
    gate._executor.shutdown(wait=True)


@pytest.mark.parametrize("route", ["no-callback", "unknown", "exception"])
@pytest.mark.parametrize("tuple_return", [False, True])
def test_public_merge_preserves_refusal(public_merge, addin_start_refusal, route, tuple_return):
    result = public_merge(addin_start_refusal["raw"], route, tuple_return)
    for key, value in addin_start_refusal["expected"].items():
        if key != "api_refused":  # The public import shape exposes the pattern and text.
            assert result[key] == value
    assert result["imported_count"] == 0
    assert result["log_path"] is None
    assert "log_excerpt" not in result
    assert "completion_unconfirmed" not in result
    assert "started" not in result


@pytest.mark.parametrize("route", ["no-callback", "unknown", "exception"])
@pytest.mark.parametrize("tuple_return", [False, True])
@pytest.mark.parametrize("decision", [False, True])
def test_public_merge_start_and_decision_precedence(public_merge, route, tuple_return, decision):
    payload = {"success": True, "started": True}
    journal = [{"kind": "conflict", "object": "Form1", "resolution": "blocked"}]
    if decision:
        payload.update(decision_required=True, decisions=journal, error_pattern="other")
    result = public_merge(json.dumps(payload), route, tuple_return)
    assert result["success"] is False
    assert result["log_path"] is None
    assert "log_excerpt" not in result
    if decision:
        assert result["error_pattern"] == "decision_required"
        assert result["decisions"] == journal
    else:
        assert result["started"] is True
        assert result["completion_unconfirmed"] is True
        assert "error_pattern" not in result


@pytest.mark.parametrize("route", ["no-callback", "unknown", "exception"])
@pytest.mark.parametrize("log_key", ["log_path", "logPath"])
def test_public_merge_preserves_explicit_refusal_log(public_merge, tmp_path, route, log_key):
    own_log = tmp_path / "own.log"
    own_log.write_text("This operation's explicit diagnostic", encoding="utf-8")
    raw = json.dumps({
        "success": False, "error": "VCS_API_REFUSED: Another API command is still running.",
        log_key: str(own_log),
    })
    result = public_merge(raw, route, True)
    assert result["error"] == "Another API command is still running."
    assert result["error_pattern"] == "operation_already_running"
    assert result["log_path"] == str(own_log)
    assert result["log_excerpt"] == "This operation's explicit diagnostic"
