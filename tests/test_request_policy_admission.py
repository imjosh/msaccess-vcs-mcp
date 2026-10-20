"""Production contract admission, using the existing identity/session seam."""
import json
from unittest.mock import Mock
import pytest
from msaccess_vcs_mcp import compatibility_session as s
from msaccess_vcs_mcp.addin_integration import VCSAddinIntegration
from tests.test_compatibility_session import AddinTransport, transport

pytestmark = pytest.mark.compatibility_session

@pytest.mark.parametrize("asynchronous", [False, True])
@pytest.mark.parametrize("marker", [None, "older-contract", s.REQUEST_POLICY_CONTRACT])
def test_contract_marker_is_checked_before_policy_dispatch(transport, marker, asynchronous):
    app, path = transport
    original = app.Run
    def run(procedure, *arguments):
        if procedure.endswith(".APIExecuteMCP"):
            app.calls["APIExecuteMCP"] += 1
            app.payloads.append(("APIExecuteMCP", arguments))
            return "true"
        response = original(procedure, *arguments)
        if procedure.endswith(".APIIdentity") and marker is not None:
            return json.dumps({**json.loads(response), "request_policy_contract": marker})
        if procedure.endswith(".APIExecuteAsync"):
            return json.dumps({"success": True, "async": True, "operation_id": "client-id"})
        return response
    app.Run = run
    app.CurrentDb = lambda: type("Database", (), {"Name": "disposable"})()
    integration = VCSAddinIntegration(path)
    integration._app = app
    integration._addin_loaded = True
    if asynchronous:
        result = integration.call_async('{"operation_id":"client-id"}', "ImportObject", "query", "q|x", "block")
        endpoint = "APIExecuteAsync"
        expected_tail = ('{"operation_id":"client-id"}', "ImportObject", "query", "q|x", "block")
    else:
        raw = integration.call_sync("ImportObject", "query", "q|x", "block")
        result = json.loads(raw)
        endpoint = "APIExecuteMCP"
        expected_tail = ("ImportObject", "query", "q|x", "block")
    assert app.calls["APIIdentity"] == app.calls["APIHandshake"] == 1
    if marker != s.REQUEST_POLICY_CONTRACT:
        assert result["error_pattern"] == "request_policy_unsupported"
        assert result["started"] is False
        assert app.calls[endpoint] == 0
    else:
        assert app.calls[endpoint] == 1
        wire = [args for method, args in app.payloads if method == endpoint][-1]
        assert "session_id" in json.loads(wire[0])
        if asynchronous:
            assert wire[1:] == expected_tail
        else:
            import uuid
            uuid.UUID(wire[1])
            assert wire[2:] == expected_tail

def test_no_marker_keeps_omitted_policy_legacy_call_usable(transport):
    app, path = transport
    app.CurrentDb = lambda: type("Database", (), {"Name": "disposable"})()
    integration = VCSAddinIntegration(path)
    integration._app = app
    integration._addin_loaded = True
    result = integration.call_sync("ExportObject", "query", "q")
    assert result == "true"
    assert app.calls["APIExecute"] == 1

def test_failed_existing_session_cannot_be_admitted_by_contract_marker(transport):
    app, path = transport
    app.reject = True
    integration = VCSAddinIntegration(path)
    integration._app = app
    integration._addin_loaded = True
    app.CurrentDb = lambda: type("Database", (), {"Name": "disposable"})()
    with pytest.raises(s.AdmissionError):
        integration.call_sync("ExportObject", "query", "q", "block")
    assert app.calls["APIExecute"] == 0


@pytest.mark.parametrize("command,arguments", [
    ("exportobject", ("query", "q", "block")),
    ("ImportByType", ("queries", False, "")),
    ("EXPORT", ("block",)),
    ("Build", ("source", "block")),
    ("MergeBuild", ("skip",)),
    ("RunFilteredTests", ("block",)),
])
def test_policy_slot_is_checked_case_insensitively_before_dispatch(command, arguments):
    result = s.request_policy_refusal("unadvertised-envelope", command, arguments)
    assert result["error_pattern"] == "request_policy_unsupported"
    assert result["started"] is False
