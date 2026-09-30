"""Policy scopes through public tools, including their gate and logging wrappers."""

import asyncio
import json
from unittest.mock import AsyncMock

import pytest

from msaccess_vcs_mcp import tools
from msaccess_vcs_mcp.access_gate import reset_access_gate
from msaccess_vcs_mcp.usage_logging import reset_logging
from tests.test_scoped_sync import _patch_import_tool


@pytest.fixture(autouse=True)
def tool_runtime(monkeypatch, tmp_path):
    monkeypatch.setattr(tools, "_ensure_env_loaded", AsyncMock())
    monkeypatch.setattr(tools, "load_config", lambda: None)
    monkeypatch.setenv("ACCESS_VCS_LOG_DIR", str(tmp_path / "usage"))
    monkeypatch.setenv("ACCESS_VCS_DIAGNOSTIC_LOG_DIR", str(tmp_path / "diag"))
    reset_access_gate()
    reset_logging()
    yield
    reset_access_gate()
    reset_logging()


@pytest.fixture(params=[
    ("vcs_import_objects", "ImportByType"),
    ("vcs_import_object", "ImportObject"),
    ("vcs_export_object", "ExportObject"),
])
def public_tool(request):
    tool_name, command = request.param

    def run(db, src, **kwargs):
        if command == "ImportByType":
            call = tools.vcs_import_objects(db, str(src), object_types=["forms"], **kwargs)
        else:
            call = getattr(tools, tool_name)(db, "form", "frmExample", **kwargs)
        return asyncio.run(call)

    return run, command


@pytest.mark.parametrize("operation_result", [
    {"success": True, "decisions": []},
    {"success": False, "error": "operation refused", "decisions": [{"kind": "conflict"}]},
    RuntimeError("operation raised"),
])
def test_cleanup_failure_keeps_operation_outcome(tmp_path, public_tool, operation_result):
    run, command = public_tool
    events = []
    log_path = str(tmp_path / "operation.log")
    (tmp_path / "operation.log").write_text("operation details\n", encoding="utf-8")
    with _patch_import_tool(tmp_path) as (db, src, addin, ops):
        def call(name, *args):
            events.append((name, *args))
            if name == command:
                if isinstance(operation_result, Exception):
                    raise operation_result
                return json.dumps({**operation_result, "logPath": log_path})
            if name == "ClearOperationPolicy":
                raise RuntimeError("cleanup raised")
            return json.dumps({"success": True})

        addin.call_sync.side_effect = call
        result = run(db, src, decision_policy="decline")

    assert events[0] == ("SetOperationPolicy", "decline")
    assert events[1][0] == command
    assert events[2:] == [("ClearOperationPolicy",)]
    assert result["policy_cleanup_error"] == "cleanup raised"
    if isinstance(operation_result, Exception):
        assert result["success"] is False
        assert result["error"] == "operation raised"
    else:
        assert result["success"] == operation_result["success"]
        assert result["decisions"] == operation_result["decisions"]
        assert result["log_path"] == result["logPath"] == log_path
        if "error" in operation_result:
            assert result["error"] == operation_result["error"]


@pytest.mark.parametrize("noninteractive", [True, False])
def test_setup_refusal_prevents_operation_and_cleanup(tmp_path, public_tool, noninteractive):
    run, command = public_tool
    refusal = {"success": False, "error": "scope already open", "decisions": []}
    with _patch_import_tool(tmp_path) as (db, src, addin, ops):
        addin.call_sync.return_value = json.dumps(refusal)
        result = run(db, src, noninteractive=noninteractive)

    assert result == refusal
    if noninteractive:
        addin.call_sync.assert_called_once_with("SetOperationPolicy", "block")
    else:
        addin.call_sync.assert_called_once_with("SetInteractionMode", 0)


def test_invalid_policy_never_reaches_access(tmp_path, public_tool):
    run, command = public_tool
    with _patch_import_tool(tmp_path) as (db, src, addin, ops):
        result = run(db, src, decision_policy="unknown")

    assert result["success"] is False
    assert result["error_pattern"] == "invalid_decision_policy"
    addin.load_addin.assert_not_called()
    addin.call_sync.assert_not_called()
