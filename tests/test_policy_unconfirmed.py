"""Request policy forwarding, terminal preservation, and admission refusals."""
import asyncio
import json
from unittest.mock import AsyncMock, Mock

import pytest

from msaccess_vcs_mcp import addin_integration
from msaccess_vcs_mcp.addin_integration import VCSAddinIntegration, start_only_result
from msaccess_vcs_mcp.decision_policy import (
    call_under_policy, normalize_terminal_result, run_import_merge,
)
from msaccess_vcs_mcp.operation_manager import OperationManager
from msaccess_vcs_mcp.tools import _addin_json_result


@pytest.mark.parametrize("command,args", [
    ("ImportObject", ("query", "q|name")),
    ("ExportObject", ("module", "modExample")),
    ("ImportByType", (["queries", "forms"], False)),
    ("ExportByType", ("forms", True)),
])
@pytest.mark.parametrize("policy", ["block", "decline", "skip", "prefer_source", "prefer_database", None])
def test_policy_is_an_argument_of_only_this_request(command, args, policy):
    addin = Mock()
    payload = {"success": True, "skipped": True, "transferred": False,
               "decisions": [{"resolution": "future_resolution", "detail": {"revision": 2}}]}
    addin.call_sync.return_value = json.dumps(payload)
    result, state = call_under_policy(addin, policy, command, *args, parse_result=_addin_json_result)
    expected = (*args, policy) if policy else args
    addin.call_sync.assert_called_once_with(command, *expected)
    assert state == "completed"
    assert result == payload


@pytest.mark.parametrize("raw", [None, "", "not json", "null", "[]", True])
def test_incompatible_result_is_an_explicit_failure(raw):
    addin = Mock()
    addin.call_sync.return_value = raw
    result, _state = call_under_policy(addin, "block", "ImportObject", "query", "q", parse_result=_addin_json_result)
    assert result["success"] is False
    assert result["error_pattern"] == "invalid_addin_response"
    addin.call_sync.assert_called_once_with("ImportObject", "query", "q", "block")


@pytest.mark.parametrize("policy", ["block", None])
def test_refused_request_has_no_setup_cleanup_or_retry(policy):
    addin = Mock()
    refusal = {"success": False, "error_pattern": "operation_already_running", "error": "busy", "request_detail": "kept"}
    addin.call_sync.return_value = json.dumps(refusal)
    result, state = call_under_policy(addin, policy, "ExportObject", "query", "q", parse_result=_addin_json_result)
    assert result == refusal
    assert state == "refused"
    assert addin.call_sync.call_count == 1


@pytest.mark.parametrize("callback_type", ["complete", "error", "cancelled"])
def test_callback_and_terminal_normalization_preserve_new_fields(callback_type):
    manager = OperationManager()
    operation_id, _queue = manager.register_operation(timeout_ms=1000)
    journal = [{"resolution": "keep_source_future", "future": {"old_hash": "abc"}}]
    payload = {"operation_id": operation_id, "type": callback_type, "message": "outcome",
               "decisions": journal, "skipped": True, "transferred": False,
               "future_outcome": {"version": 2}, "initiating_source": "mcp", "policy": "skip"}
    manager.route_callback(operation_id, payload)
    completed = asyncio.run(manager.wait_for_completion(operation_id, timeout_seconds=1))
    result = normalize_terminal_result(completed, "failed")
    for key in ("operation_id", "decisions", "skipped", "transferred", "future_outcome", "initiating_source", "policy"):
        assert result[key] == payload[key]
    assert result["success"] is (callback_type == "complete")


def test_async_third_argument_keeps_boolean_empty_and_pipe(monkeypatch, tmp_path):
    addin_path = tmp_path / "Version Control.accda"
    addin_path.touch()
    addin = VCSAddinIntegration(str(addin_path))
    addin._app = Mock()
    addin._app.Run.return_value = json.dumps({"success": True, "async": True, "operation_id": "client-id"})
    if hasattr(addin_integration, "ensure_session"):
        monkeypatch.setattr(addin_integration, "ensure_session", lambda *_: "accepted-session-envelope")
    callback = json.dumps({"operation_id": "client-id", "callback_url": "http://127.0.0.1/callback"})
    started = addin.call_async(callback, "ImportByType", "queries|forms", False, "")
    wire = addin._app.Run.call_args.args
    prefix = ("accepted-session-envelope", callback, "ImportByType") if wire[0].endswith("APIExecuteAsync") else (callback, "ImportByType")
    assert wire[1:] == (*prefix, "queries|forms", False, "")
    assert type(wire[-2]) is bool
    assert started["operation_id"] == "client-id"


@pytest.mark.parametrize("pattern", ["operation_already_running", "timer_unavailable", "invalid_decision_policy"])
def test_synchronous_async_refusal_never_waits_or_falls_back(pattern):
    addin = Mock()
    refusal = {"success": False, "error_pattern": pattern, "error": "refused", "request_field": "kept"}
    addin.call_async.return_value = refusal
    manager = Mock()
    manager.wait_for_completion = AsyncMock()
    result = asyncio.run(run_import_merge(addin, "user.accdb", "user.accdb.src", "block", callback_info="{}", op_manager=manager, operation_id="client-id"))
    assert result == refusal
    addin.call_async.assert_called_once_with("{}", "MergeBuild", "block")
    manager.wait_for_completion.assert_not_awaited()
    manager.unregister_operation.assert_called_once_with("client-id")
    addin.merge_build.assert_not_called()


@pytest.mark.parametrize("command", ["export", "VBA export"])
def test_structured_export_completion_is_final(command):
    payload = {"success": True, "skipped": True, "transferred": False,
               "operation_id": "direct-id", "decisions": [{"resolution": "unfamiliar"}]}
    assert start_only_result(json.dumps(payload), command, "Export") == payload


@pytest.mark.parametrize("policy", ["block", "skip", None])
def test_export_vba_fallback_forwards_policy_and_keeps_final_outcome(tmp_path, policy):
    library = tmp_path / "Version Control.accda"
    library.touch()
    addin = VCSAddinIntegration(str(library))
    payload = {"success": True, "skipped": True, "transferred": False, "operation_id": "stable"}
    addin._call_addin_function = Mock(return_value=json.dumps(payload))
    result = addin.export_vba("user.accdb", decision_policy=policy)
    addin._call_addin_function.assert_called_once_with("ExportVBA", *((policy,) if policy else ()))
    for key, value in payload.items():
        assert result[key] == value


def test_sync_dispatch_exception_retains_id_and_never_replays(monkeypatch, tmp_path):
    import uuid
    path = tmp_path / "Version Control.accda"
    path.touch()
    addin = VCSAddinIntegration(str(path))
    addin._addin_loaded = True
    addin._app = Mock()
    addin._app.Run.side_effect = RuntimeError("disconnected after possible execution")
    if hasattr(addin_integration, "request_policy_contract_confirmed"):
        monkeypatch.setattr(addin_integration, "request_policy_contract_confirmed", lambda _: True)
    result = json.loads(addin.call_sync("ImportObject", "query", "q", "block"))
    assert result["success"] is False
    assert result["error_pattern"] == "operation_outcome_unconfirmed"
    assert result["dispatch_uncertain"] is True
    uuid.UUID(result["operation_id"])
    addin._app.Run.assert_called_once()
