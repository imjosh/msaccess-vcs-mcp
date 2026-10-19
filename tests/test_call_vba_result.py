"""Provider outcomes through the actual public-result consumer, without COM."""

import json
from datetime import datetime
from decimal import Decimal
from pathlib import Path
from unittest.mock import MagicMock

import pytest

from msaccess_vcs_mcp import tools
from msaccess_vcs_mcp.compatibility_session import AdmissionError, session_failure


TAG = "msaccess-vcs.call-vba-result/1"
FAILURE_TEXT = '{"success":false,"error":"Example failure","errorNumber":5}'
COMPATIBILITY_TEXT = '{"success":false,"error_pattern":"version_incompatible"}'


def frame(value):
    return (value, *range(30))


@pytest.fixture
def boundary(monkeypatch):
    connection = MagicMock()
    connection.__enter__.return_value = connection
    connection.connect.return_value = (MagicMock(), MagicMock())
    monkeypatch.setattr(tools, "AccessConnection", lambda path: connection)
    monkeypatch.setattr(tools, "validate_database_path", Path)
    monkeypatch.setattr(tools, "get_config", lambda: {"ACCESS_VCS_ADDIN_PATH": r"C:\library\Version Control.accda"})
    monkeypatch.setattr(tools.VCSAddinIntegration, "load_addin", lambda *args, **kwargs: True)
    monkeypatch.setattr(tools, "ensure_session", lambda *args: "accepted-session")
    monkeypatch.setattr(tools, "log_code_execution", lambda *args, **kwargs: None)
    dispatch = MagicMock()
    monkeypatch.setattr(tools, "_run_target", lambda _permit, app, procedure, args: dispatch(app, procedure, args))

    def call(raw, args=None, function="ExampleFunction", transport=None):
        dispatch.return_value = transport if transport is not None else {"success": True, "result": raw}
        return tools._execute_call_vba(r"C:\fixtures\user.accdb", function, args)

    return call, dispatch


@pytest.mark.parametrize("wrap", [lambda x: x, frame], ids=["direct", "framed"])
def test_provider_error_and_identical_user_failure_bytes_have_distinct_origins(boundary, wrap):
    call, dispatch = boundary
    failed = call(wrap([TAG, "failed", 5, "Example failure"]))
    assert failed["success"] is False
    assert failed["error_number"] == 5
    assert failed["error"] == "Example failure"
    assert failed["error_pattern"] == "vba_execution_failed"
    returned = call(wrap([TAG, "returned", FAILURE_TEXT]))
    assert returned["success"] is True
    assert returned["result"] == FAILURE_TEXT
    assert dispatch.call_count == 2


@pytest.mark.parametrize("value", [
    None, "", 0, 42, 1.25, True, False, Decimal("123.40"), datetime(2026, 1, 2, 3, 4),
    FAILURE_TEXT, "VCS_API_REFUSED: legitimate target text", COMPATIBILITY_TEXT,
    [TAG, "failed", 5, "target array"], (TAG, "returned", "target tuple"),
    tuple(range(31)), [[TAG, "returned", "nested"]],
])
@pytest.mark.parametrize("wrap", [lambda x: x, frame], ids=["direct", "framed"])
def test_target_values_stay_opaque_until_existing_public_conversion(boundary, value, wrap):
    call, dispatch = boundary
    result = call(wrap((TAG, "returned", value)))
    assert result["success"] is True
    assert result["result"] == (str(value) if value is not None else None)
    dispatch.assert_called_once()


@pytest.mark.parametrize("raw", [
    None, "legacy value", FAILURE_TEXT, "{bad json", "{}", "[]", (), [],
    '{"error_pattern":[]}', '{"error_pattern":{}}',
    frame('{"error_pattern":[]}'), frame('{"error_pattern":{}}'),
    ["old-contract", "returned", 1], [TAG], [TAG, "returned"],
    [TAG, "returned", 1, 2], [TAG, "unknown", 1],
    [TAG, "failed", True, "boolean is not a VBA error number"],
    [TAG, "failed", 5.0, "float"], [TAG, "failed", "5", "string"],
    [TAG, "failed", 5, None], [TAG, "failed", 5],
    ([TAG, "returned", "singleton must not unwrap"],),
    frame(frame([TAG, "returned", "nested frame must not unwrap"])),
    tuple([TAG, *range(30)]),
    frame([TAG, "failed", False, "malformed framed"]),
])
def test_malformed_or_old_provider_reply_fails_once_without_fallback(boundary, raw):
    call, dispatch = boundary
    result = call(raw)
    assert result["success"] is False
    assert result["error_pattern"] == "vba_result_contract_invalid"
    assert result["provider_response"] == repr(raw)
    assert "started" not in result
    dispatch.assert_called_once()
    assert dispatch.call_args.args[2][1] == "CallVBAResult"


@pytest.mark.parametrize("raw", ["VCS_API_REFUSED: busy", COMPATIBILITY_TEXT])
@pytest.mark.parametrize("wrap", [lambda x: x, frame], ids=["direct", "framed"])
def test_unwrapped_admission_refusal_remains_failure(boundary, raw, wrap):
    call, dispatch = boundary
    result = call(wrap(raw))
    assert result["success"] is False
    if raw == COMPATIBILITY_TEXT:
        assert result["error_pattern"] == "version_incompatible"
    else:
        assert result["error"] == "busy"
    dispatch.assert_called_once()


@pytest.mark.parametrize("success", [True, None])
@pytest.mark.parametrize("pattern", ["compatibility_session_invalid", "version_incompatible", "version_unconfirmed"])
def test_known_unwrapped_refusal_cannot_claim_success(boundary, success, pattern):
    call, dispatch = boundary
    diagnostic = {"error_pattern": pattern, "error": "admission refused"}
    if success is not None:
        diagnostic["success"] = success
    result = call(json.dumps(diagnostic))
    assert result["success"] is False
    assert result["error_pattern"] == pattern
    assert result["error"] == "admission refused"
    dispatch.assert_called_once()


@pytest.mark.parametrize("count", range(4))
def test_zero_to_three_arguments_dispatch_once_with_session(boundary, count):
    call, dispatch = boundary
    args = [f"argument-{i}" for i in range(count)]
    assert call([TAG, "returned", "ok"], args)["success"] is True
    dispatch.assert_called_once()
    actual = dispatch.call_args.args
    assert actual[1].endswith(".APIExecute")
    assert actual[2] == ["accepted-session", "CallVBAResult", "ExampleFunction", json.dumps(args)]


def test_argument_overflow_never_dispatches(boundary):
    call, dispatch = boundary
    assert call(None, ["a", "b", "c", "d"])["success"] is False
    dispatch.assert_not_called()


def test_session_refusal_never_dispatches(boundary, monkeypatch):
    call, dispatch = boundary
    def refuse(*args):
        raise AdmissionError(session_failure("stale_instance"))
    monkeypatch.setattr(tools, "ensure_session", refuse)
    with pytest.raises(AdmissionError):
        call(None)
    dispatch.assert_not_called()


def test_target_failure_is_reported_dispatched_and_never_retried(boundary):
    call, dispatch = boundary
    result = call(None, transport={"success": False, "error": RuntimeError("target failure")})
    assert result["success"] is False
    assert "target failure" in result["error"]
    assert result["dispatch"] == "dispatched"
    dispatch.assert_called_once()


@pytest.mark.parametrize("raw,expected", [
    (FAILURE_TEXT, FAILURE_TEXT), (frame("legacy raw API value"), "legacy raw API value"),
])
def test_raw_api_keeps_its_arguments_and_legacy_projection(boundary, raw, expected):
    call, dispatch = boundary
    result = call(raw, ["GetVCSVersion"], "VCS.API")
    assert result["success"] is True
    assert result["result"] == expected
    dispatch.assert_called_once()
    assert dispatch.call_args.args[2] == ["accepted-session", "GetVCSVersion"]
