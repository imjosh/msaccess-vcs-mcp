"""M41: setup must acknowledge success before any dependent dispatch."""

import asyncio
import json
from unittest.mock import MagicMock, patch

import pytest

from tests.test_api_refusal import FakeApp, REFUSAL, _tool_env, _unwrap
from msaccess_vcs_mcp import tools


class SetupApp(FakeApp):
    def __init__(self, method, response):
        super().__init__(None)
        self.method = method
        self.response = response
        self.setup_args = []

    def Run(self, name, *args):
        if name.endswith(".APIExecuteAsync"):
            name, args = name[:-15] + "APIAsync", args[1:]
        elif name.endswith(".APIExecute"):
            name, args = name[:-10] + "API", args[1:]
        if args[0] == self.method:
            self.calls.append(args[0])
            self.setup_args.append(args[1:])
            return self.response
        if args[0] == "RunFilteredTests":
            self.calls.append(args[0])
            return json.dumps({"allPassed": True, "summary": {"passed": 1}})
        return super().Run(name, *args)


FAILURES = [
    json.dumps({"success": False, "error": "setup failed"}),
    REFUSAL,
    None,
    "",
    "not json",
    "{}",
    json.dumps({"success": "true"}),
    json.dumps({"success": 1}),
]


@pytest.mark.parametrize("response", FAILURES)
@pytest.mark.parametrize("callback", [False, True])
@pytest.mark.parametrize("filter_value", [None, "SQL"])
def test_filter_failure_never_dispatches_tests(tmp_path, response, callback, filter_value):
    app = SetupApp("SetOption", response)
    manager = MagicMock()
    with (
        _tool_env(tmp_path, app) as (db, _),
        patch("msaccess_vcs_mcp.tools.get_callback_url", return_value="http://callback" if callback else None),
        patch("msaccess_vcs_mcp.tools._get_operation_manager", return_value=manager),
    ):
        result = asyncio.run(_unwrap(tools.vcs_run_tests)(db, filter=filter_value))
    assert result["success"] is False
    assert app.calls == ["SetOption"]
    assert app.setup_args == [("DefaultTestFilter", filter_value or "")]
    manager.register_operation.assert_not_called()
    if response == FAILURES[0]:
        assert result["error"] == "setup failed"
    if response == REFUSAL:
        assert result["error_pattern"] == "operation_already_running"


@pytest.mark.parametrize("response", FAILURES)
def test_registration_failure_never_writes_option(tmp_path, response):
    app = SetupApp("RegisterSession", response)
    with _tool_env(tmp_path, app) as (db, _):
        result = _unwrap(tools.vcs_set_option)(db, "ShowDebug", True)
    assert result["success"] is False
    assert app.calls == ["RegisterSession"]
    assert app.setup_args == [("s1",)]
    if response == FAILURES[0]:
        assert result["error"] == "setup failed"
    if response == REFUSAL:
        assert result["error_pattern"] == "operation_already_running"


@pytest.mark.parametrize("filter_value", [None, "SQL"])
def test_successful_filter_setup_runs_tests(tmp_path, filter_value):
    app = SetupApp("SetOption", '{"success": true}')
    with _tool_env(tmp_path, app) as (db, _):
        result = asyncio.run(_unwrap(tools.vcs_run_tests)(db, filter=filter_value))
    assert result["success"] is True
    assert app.calls == ["SetOption", "RunFilteredTests"]
    assert app.setup_args == [("DefaultTestFilter", filter_value or "")]


@pytest.mark.parametrize("session_id", ["s1", None])
def test_successful_registration_or_no_session_writes_option(tmp_path, session_id):
    app = SetupApp("RegisterSession", '{"success": true}')
    with (
        _tool_env(tmp_path, app) as (db, _),
        patch("msaccess_vcs_mcp.tools.get_session_id", return_value=session_id),
    ):
        result = _unwrap(tools.vcs_set_option)(db, "ShowDebug", True)
    assert result["success"] is True
    assert app.calls == (["RegisterSession", "SetOption"] if session_id else ["SetOption"])
