"""Request policies through public MCP wrappers."""
import asyncio
import json
from unittest.mock import AsyncMock
import pytest
from msaccess_vcs_mcp import tools
from msaccess_vcs_mcp.access_gate import reset_access_gate
from msaccess_vcs_mcp.usage_logging import reset_logging
from tests.test_scoped_sync import _patch_import_tool, _patch_export_tool

@pytest.fixture(autouse=True)
def runtime(monkeypatch, tmp_path):
    monkeypatch.setattr(tools, "_ensure_env_loaded", AsyncMock())
    monkeypatch.setattr(tools, "load_config", lambda: None)
    monkeypatch.setenv("ACCESS_VCS_LOG_DIR", str(tmp_path / "usage"))
    monkeypatch.setenv("ACCESS_VCS_DIAGNOSTIC_LOG_DIR", str(tmp_path / "diag"))
    reset_access_gate()
    reset_logging()
    yield
    reset_access_gate()
    reset_logging()

@pytest.mark.parametrize("name,command", [("vcs_import_object", "ImportObject"), ("vcs_export_object", "ExportObject"), ("vcs_import_objects", "ImportByType"), ("vcs_export_database", "ExportByType")])
@pytest.mark.parametrize("policy", ["block", "prefer_source", "prefer_database", "skip", "decline", None])
def test_public_tool_forwards_policy_once(tmp_path, name, command, policy):
    patcher = _patch_export_tool if command == "ExportByType" else _patch_import_tool
    with patcher(tmp_path) as (db, src, addin, ops):
        payload = {"success": True, "skipped": True, "transferred": False,
                   "decisions": [{"resolution": "future", "extra": [1]}]}
        addin.call_sync.return_value = json.dumps(payload)
        kwargs = {"noninteractive": policy is not None, "decision_policy": policy or "block"}
        if command in {"ImportByType", "ExportByType"}:
            kwargs["object_types"] = ["queries"]
            result = asyncio.run(getattr(tools, name)(db, str(src), **kwargs))
            args = ("queries", False)
        else:
            result = asyncio.run(getattr(tools, name)(db, "query", "q", **kwargs))
            args = ("query", "q")
        addin.call_sync.assert_called_once_with(command, *args, *((policy,) if policy else ()))
        for key in ("success", "skipped", "transferred", "decisions"):
            assert result[key] == payload[key]

@pytest.mark.parametrize("name", ["vcs_import_object", "vcs_export_object", "vcs_import_objects", "vcs_export_database"])
def test_invalid_policy_is_rejected_before_access(tmp_path, name):
    with _patch_import_tool(tmp_path) as (db, src, addin, ops):
        positional = (db, str(src)) if name.endswith(("objects", "database")) else (db, "query", "q")
        result = asyncio.run(getattr(tools, name)(*positional, decision_policy="unknown"))
        assert result["success"] is False
        assert result["error_pattern"] == "invalid_decision_policy"
        addin.load_addin.assert_not_called()
        addin.call_sync.assert_not_called()


@pytest.mark.parametrize("full_export", [False, True])
@pytest.mark.parametrize("transport", ["callback", "inline", "synchronous"])
@pytest.mark.parametrize("policy", ["block", "prefer_source", "prefer_database", "skip", "decline", None])
def test_full_export_forwards_request_and_keeps_transfer_fields(tmp_path, monkeypatch, full_export, transport, policy):
    with _patch_export_tool(tmp_path) as (db, src, addin, ops):
        payload = {"success": True, "operation_id": "public-id", "skipped": True,
                   "transferred": False, "future_transfer": {"reason": "preserved"},
                   "decisions": [{"resolution": "future_source_preserved", "extra": 9}]}
        addin.export_source.return_value = payload
        addin.call_async.return_value = {"async": True} if transport == "callback" else {"sync": True, "result": json.dumps(payload)}
        ops.wait_for_completion = AsyncMock(return_value=payload)
        if transport == "synchronous":
            monkeypatch.setattr(tools, "get_callback_url", lambda: None)
        result = asyncio.run(tools.vcs_export_database(db, str(src), full_export=full_export, noninteractive=policy is not None, decision_policy=policy or "block"))
        for key, value in payload.items():
            assert result[key] == value
        if transport == "synchronous":
            addin.export_source.assert_called_once_with(db, str(src), full_export=full_export, decision_policy=policy)
            addin.call_async.assert_not_called()
        else:
            command = "FullExport" if full_export else "Export"
            addin.call_async.assert_called_once_with("{}", command, *((policy,) if policy else ()))
            addin.export_source.assert_not_called()


def test_full_export_queue_refusal_does_not_wait_or_retry(tmp_path):
    with _patch_export_tool(tmp_path) as (db, src, addin, ops):
        refusal = {"success": False, "error_pattern": "timer_unavailable", "error": "queue full", "request_detail": {"owner": "other"}}
        addin.call_async.return_value = refusal
        ops.wait_for_completion = AsyncMock()
        result = asyncio.run(tools.vcs_export_database(db, str(src)))
        for key, value in refusal.items():
            assert result[key] == value
        ops.wait_for_completion.assert_not_awaited()
        addin.export_source.assert_not_called()
        ops.unregister_operation.assert_called_once_with("op-1")


@pytest.mark.parametrize("name,command", [("vcs_import_object", "ImportObject"), ("vcs_export_object", "ExportObject")])
@pytest.mark.parametrize("noninteractive", [True, False])
def test_public_single_tools_preserve_mcp_id_and_source(tmp_path, monkeypatch, name, command, noninteractive):
    import uuid
    from unittest.mock import Mock
    from msaccess_vcs_mcp import addin_integration
    from msaccess_vcs_mcp.addin_integration import VCSAddinIntegration
    library = tmp_path / "Version Control.accda"
    library.touch()
    integration = VCSAddinIntegration(str(library))
    integration._addin_loaded = True
    app = Mock()
    integration._app = app
    integration.load_addin = Mock()
    wire = []
    if hasattr(addin_integration, "request_policy_contract_confirmed"):
        monkeypatch.setattr(addin_integration, "request_policy_contract_confirmed", lambda _: True)
    def dispatch(endpoint, *arguments):
        production = endpoint.endswith(".APIExecuteMCP")
        assert production or endpoint.endswith(".APIMCP")
        index = 1 if production else 0
        if production:
            assert arguments[0] == "{}"
        operation_id, method, *forwarded = arguments[index:]
        uuid.UUID(operation_id)
        assert method == command
        assert forwarded == ["query", "q", *(["block"] if noninteractive else [])]
        wire.append((endpoint, operation_id))
        return json.dumps({"success": True, "operation_id": operation_id,
                           "initiating_source": "mcp", "skipped": True, "transferred": False})
    app.Run.side_effect = dispatch
    with _patch_import_tool(tmp_path) as (db, src, ignored, ops):
        monkeypatch.setattr(tools, "VCSAddinIntegration", lambda *_: integration)
        first = asyncio.run(getattr(tools, name)(db, "query", "q", noninteractive=noninteractive))
        second = asyncio.run(getattr(tools, name)(db, "query", "q", noninteractive=noninteractive))
    assert len(wire) == 2
    assert first["operation_id"] == wire[0][1]
    assert second["operation_id"] == wire[1][1]
    assert first["operation_id"] != second["operation_id"]
    assert first["initiating_source"] == second["initiating_source"] == "mcp"
    assert first["skipped"] is True and first["transferred"] is False
