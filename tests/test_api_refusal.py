"""M39: a dispatcher refusal is a failure on every hard-coded API call.

The real ``VCSAddinIntegration.call_sync`` / ``call_async`` run against a fake
``Application.Run`` that returns the add-in's marked refusal for one method and
a plain success for every other, so each row proves the refusal is recognised
at the integration boundary and the tool never reports success.
"""

import asyncio
import json
from contextlib import contextmanager
from unittest.mock import MagicMock, patch

import pytest

from msaccess_vcs_mcp.addin_integration import (
    API_REFUSED_PATTERN,
    API_REFUSED_PREFIX,
    VCSAddinIntegration,
)

REFUSAL = API_REFUSED_PREFIX + (
    "VCS API refused a call to 'X': another API command is still running."
)
PATTERN = "operation_already_running"


def _unwrap(tool_fn):
    fn = tool_fn
    while hasattr(fn, "__wrapped__"):
        fn = fn.__wrapped__
    return fn


@pytest.fixture(autouse=True)
def _diag_dir(monkeypatch, tmp_path_factory):
    monkeypatch.setenv("ACCESS_VCS_DIAGNOSTIC_LOG_DIR", str(tmp_path_factory.mktemp("diag")))


class FakeApp:
    """Application.Run stand-in: refuses ``refused`` and logs every dispatch."""

    def __init__(self, refused, tuple_return=False):
        self.refused = refused
        self.tuple_return = tuple_return
        self.calls = []

    def CurrentDb(self):
        return MagicMock()

    def Run(self, name, *args):
        if name.endswith(".APIAsync"):
            self.calls.append(args[1])
            if args[1] == self.refused:
                return json.dumps({"success": False, "error": REFUSAL})
            return json.dumps({"sync": True, "result": json.dumps({"success": True})})
        method = args[0]
        self.calls.append(method)
        if method == self.refused:
            result = REFUSAL
        elif method in ("IsVBACompiled", "CompileVBA"):
            result = True
        elif method == "SetInteractionMode":
            result = json.dumps({"success": True, "effective_mode": 0})
        else:
            result = json.dumps({"success": True})
        return (result, None) if self.tuple_return else result


def _integration(app):
    addin = VCSAddinIntegration.__new__(VCSAddinIntegration)
    addin.addin_path = r"C:\x\Version Control.accda"
    addin._addin_loaded = True
    addin._app = app
    return addin


@contextmanager
def _tool_env(tmp_path, app):
    db = tmp_path / "t.accdb"
    db.touch()
    src = tmp_path / "src"
    src.mkdir(exist_ok=True)
    conn = MagicMock()
    conn.__enter__ = MagicMock(return_value=conn)
    conn.__exit__ = MagicMock(return_value=False)
    conn.connect.return_value = (MagicMock(), MagicMock())
    addin = _integration(app)
    with (
        patch("msaccess_vcs_mcp.tools.AccessConnection", return_value=conn),
        patch("msaccess_vcs_mcp.tools.VCSAddinIntegration", return_value=addin),
        patch("msaccess_vcs_mcp.tools.validate_database_path", return_value=db),
        patch("msaccess_vcs_mcp.tools.validate_source_directory", return_value=src),
        patch("msaccess_vcs_mcp.tools.validate_export_directory", return_value=src),
        patch("msaccess_vcs_mcp.tools.check_write_permission", return_value=None),
        patch("msaccess_vcs_mcp.tools._check_database_busy", return_value=None),
        patch("msaccess_vcs_mcp.tools.get_callback_url", return_value=None),
        patch("msaccess_vcs_mcp.tools._get_operation_manager", return_value=None),
        patch("msaccess_vcs_mcp.tools.get_session_id", return_value="s1"),
        patch(
            "msaccess_vcs_mcp.tools.get_config",
            return_value={"ACCESS_VCS_ADDIN_PATH": str(tmp_path / "Version Control.accda")},
        ),
    ):
        yield str(db), str(src)


def _call(tool_name, db, src, **extra):
    from msaccess_vcs_mcp import tools

    fn = _unwrap(getattr(tools, tool_name))
    kwargs = {
        "vcs_export_object": dict(object_type="modules", object_name="m"),
        "vcs_import_object": dict(object_type="modules", object_name="m"),
        "vcs_export_database": dict(output_dir=src, object_types=["modules"]),
        "vcs_import_objects": dict(source_dir=src, object_types=["modules"]),
        "vcs_check_vba_compiled": {},
        "vcs_compile_vba": {},
        "vcs_execute_sql": dict(sql="SELECT 1"),
        "vcs_set_option": dict(option_name="ShowDebug", value=True),
        "vcs_get_option": dict(option_name="ShowDebug"),
        "vcs_get_log": {},
    }[tool_name]
    result = fn(db, **kwargs, **extra)
    return asyncio.run(result) if asyncio.iscoroutine(result) else result


# (tool, refused add-in method, extra tool kwargs). Covers the review's
# section 2 inventory for every call the server makes through call_sync other
# than vcs_call_vba and the RunVBA worker (tested separately below).
ROWS = [
    ("vcs_export_object", "ExportObject", {}),
    ("vcs_export_object", "SetOperationPolicy", {}),
    ("vcs_import_object", "ImportObject", {}),
    ("vcs_import_object", "SetOperationPolicy", {}),
    ("vcs_export_database", "ExportByType", {}),
    ("vcs_import_objects", "ImportByType", {}),
    ("vcs_import_objects", "SetOperationPolicy", {}),
    ("vcs_import_objects", "SetInteractionMode", {"noninteractive": False}),
    ("vcs_check_vba_compiled", "IsVBACompiled", {}),
    ("vcs_compile_vba", "CompileVBA", {}),
    ("vcs_execute_sql", "ExecuteSQL", {}),
    ("vcs_set_option", "SetOption", {}),
    ("vcs_get_option", "GetOption", {}),
    ("vcs_get_log", "GetLogContent", {}),
]


@pytest.mark.parametrize("tuple_return", [False, True])
@pytest.mark.parametrize("tool_name,refused,extra", ROWS)
def test_refusal_is_a_failure_with_pattern(tmp_path, tool_name, refused, extra, tuple_return):
    app = FakeApp(refused, tuple_return)
    with _tool_env(tmp_path, app) as (db, src):
        result = _call(tool_name, db, src, **extra)
    assert result["success"] is False
    assert result["error_pattern"] == PATTERN
    assert "another API command is still running" in result["error"]
    assert not result["error"].startswith(API_REFUSED_PREFIX)
    assert result.get("compiled") in (None, False)
    if refused in ("SetOperationPolicy", "SetInteractionMode"):
        assert app.calls == [refused]  # the object call never dispatched


@pytest.mark.parametrize("tool_name", ["vcs_export_object", "vcs_import_object"])
def test_clear_policy_refusal_becomes_policy_cleanup_error(tmp_path, tool_name):
    app = FakeApp("ClearOperationPolicy")
    with _tool_env(tmp_path, app) as (db, src):
        result = _call(tool_name, db, src)
    assert "another API command is still running" in result["policy_cleanup_error"]


def test_run_vba_refusal_is_a_failure(tmp_path):
    from msaccess_vcs_mcp import tools

    app = FakeApp("RunVBA")
    raw = _integration(app).call_sync("RunVBA", "x = 1")
    with (
        _tool_env(tmp_path, app) as (db, _src),
        patch(
            "msaccess_vcs_mcp.tools.run_vba_resilient",
            return_value={"success": True, "result": raw},
        ),
        patch("msaccess_vcs_mcp.tools.log_code_execution"),
    ):
        result = _unwrap(tools.vcs_run_vba)(db, code="x = 1")
    assert result["success"] is False
    assert result["error_pattern"] == PATTERN


def test_call_sync_returns_non_refusals_unchanged():
    addin = _integration(FakeApp("Nothing"))
    assert addin.call_sync("IsVBACompiled") is True
    assert addin.call_sync("GetOption", "x") == json.dumps({"success": True})


def test_call_sync_refusal_shape():
    parsed = json.loads(_integration(FakeApp("GetOption")).call_sync("GetOption", "x"))
    assert parsed["success"] is False
    assert parsed["error_pattern"] == API_REFUSED_PATTERN == PATTERN
    assert parsed["api_refused"] is True


@pytest.mark.parametrize("command", ["RunFilteredTests", "MergeBuild", "Export"])
def test_inline_async_envelope_refusal_gets_the_pattern(command):
    from msaccess_vcs_mcp.decision_policy import is_start_refusal

    result = _integration(FakeApp(command)).call_async("{}", command)
    assert result["success"] is False
    assert result["error_pattern"] == PATTERN
    assert not result["error"].startswith(API_REFUSED_PREFIX)
    assert is_start_refusal(result)


def test_call_async_bare_refusal_string():
    app = MagicMock()
    app.Run.return_value = REFUSAL
    result = _integration(app).call_async("{}", "Export")
    assert result["success"] is False and result["error_pattern"] == PATTERN


def test_call_async_normal_results_unchanged():
    result = _integration(FakeApp("Nothing")).call_async("{}", "Export")
    assert result == {"sync": True, "result": json.dumps({"success": True})}


@pytest.mark.parametrize("tool_name,method", [
    ("vcs_execute_sql", "ExecuteSQL"),
    ("vcs_set_option", "SetOption"),
    ("vcs_export_object", "ExportObject"),
    ("vcs_import_objects", "ImportByType"),
    ("vcs_export_database", "ExportByType"),
])
def test_non_json_from_a_json_contract_method_is_a_failure(tmp_path, tool_name, method):
    class Garbled(FakeApp):
        def Run(self, name, *args):
            if args and args[0] == method:
                return "not json"
            return super().Run(name, *args)

    with _tool_env(tmp_path, Garbled(None)) as (db, src):
        result = _call(tool_name, db, src)
    assert result["success"] is False
    assert result["error_pattern"] == "invalid_addin_response"


def test_raw_contract_methods_keep_wrapping(tmp_path):
    class Raw(FakeApp):
        def Run(self, name, *args):
            if args and args[0] in ("GetOption", "GetLogContent"):
                return "plain value"
            return super().Run(name, *args)

    with _tool_env(tmp_path, Raw(None)) as (db, src):
        assert _call("vcs_get_option", db, src) == {
            "success": True, "option": "ShowDebug", "value": "plain value",
        }
        assert _call("vcs_get_log", db, src) == {"success": True, "content": "plain value"}
