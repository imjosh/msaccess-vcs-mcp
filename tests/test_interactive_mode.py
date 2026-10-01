"""Interactive operations require confirmation, including with older add-ins."""

import asyncio
import json
from unittest.mock import AsyncMock, Mock, patch

import pytest

from msaccess_vcs_mcp.decision_policy import select_interactive_mode
from msaccess_vcs_mcp import tools
from msaccess_vcs_mcp.access_gate import reset_access_gate
from msaccess_vcs_mcp.usage_logging import reset_logging
from tests.interaction_mode_contract import INTERACTIVE_CONFIRMED, INTERACTIVE_REFUSED
from tests.test_scoped_sync import _patch_import_tool


@pytest.mark.parametrize("response", [
    None, "", "not JSON", "null", "[]", True, 0,
    {}, {"success": True},
    {"success": True, "effective_mode": 2},
    {"success": True, "effective_mode": "0"},
    {"success": True, "effective_mode": False},
    {"success": 1, "effective_mode": 0},
])
def test_unconfirmed_mode_prevents_start(response):
    addin = Mock()
    addin.call_sync.return_value = response

    refusal = select_interactive_mode(addin, None)

    assert refusal is not None
    assert refusal["success"] is False
    assert refusal["error_pattern"] == "interaction_mode_unconfirmed"
    assert "A24" in refusal["error"]
    addin.call_sync.assert_called_once_with("SetInteractionMode", 0)


@pytest.fixture
def public_runtime(monkeypatch, tmp_path):
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
    "scoped_import", "import_object", "export_object",
    "merge_async", "merge_sync",
])
def interactive_tool(request, public_runtime):
    path = request.param
    command = {
        "scoped_import": "ImportByType", "import_object": "ImportObject",
        "export_object": "ExportObject",
    }.get(path, "MergeBuild")

    def run(db, src):
        if path == "scoped_import":
            call = tools.vcs_import_objects(db, str(src), object_types=["forms"], noninteractive=False)
        elif path in {"import_object", "export_object"}:
            call = getattr(tools, "vcs_" + path)(db, "form", "frmExample", noninteractive=False)
        else:
            call = tools.vcs_import_objects(db, str(src), noninteractive=False)
        with patch.object(tools, "get_callback_url", return_value=None if path.endswith("sync") else "http://localhost:1/cb"):
            return asyncio.run(call)

    return run, command


@pytest.mark.parametrize("selection", [None, "", {"success": True}, INTERACTIVE_REFUSED])
def test_public_tool_never_starts_or_clears_an_unconfirmed_mode(tmp_path, interactive_tool, selection):
    run, command = interactive_tool
    with _patch_import_tool(tmp_path) as (db, src, addin, ops):
        addin.call_sync.return_value = json.dumps(selection) if isinstance(selection, dict) else selection
        result = run(db, src)

    assert result["success"] is False
    if selection == INTERACTIVE_REFUSED:
        assert result == INTERACTIVE_REFUSED
    else:
        assert result["error_pattern"] == "interaction_mode_unconfirmed"
    addin.call_sync.assert_called_once_with("SetInteractionMode", 0)
    addin.call_async.assert_not_called()
    addin.merge_build.assert_not_called()
    ops.register_operation.assert_not_called()


def test_public_tool_starts_after_confirmed_interactive_mode(tmp_path, interactive_tool):
    run, command = interactive_tool
    events = []
    with _patch_import_tool(tmp_path) as (db, src, addin, ops):
        def sync(name, *args):
            events.append(name)
            return json.dumps(INTERACTIVE_CONFIRMED if name == "SetInteractionMode" else {"success": True})

        def async_call(*args):
            events.append(args[1])
            return {"sync": True, "result": json.dumps({"success": True})}

        def merge(*args):
            events.append("MergeBuild")
            return {"success": True}

        addin.call_sync.side_effect = sync
        addin.call_async.side_effect = async_call
        addin.merge_build.side_effect = merge
        result = run(db, src)

    assert result["success"] is True
    assert events[0] == "SetInteractionMode"
    assert command in events[1:]
    assert "SetOperationPolicy" not in events
    assert "ClearOperationPolicy" not in events
