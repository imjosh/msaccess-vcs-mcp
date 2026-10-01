"""Noninteractive setup and cleanup need an affirmative acknowledgment from the add-in."""

import json
from unittest.mock import Mock, patch

import pytest

from msaccess_vcs_mcp.decision_policy import call_under_policy, clear_operation_policy

_CONFIRMED = {"success": True, "policy": "block"}


def _run(addin, policy="block"):
    return call_under_policy(addin, policy, "ExportObject", "query", "q", parse_result=json.loads)


@pytest.mark.parametrize("response", [
    None, "", "not JSON", "null", "[]", True, 0,
    {}, {"success": True},
    {"success": True, "policy": None},
    {"success": True, "policy": "skip"},
    {"success": True, "policy": 1},
    {"success": 1, "policy": "block"},
])
def test_unconfirmed_setup_prevents_dispatch_and_cleanup(response):
    addin = Mock()
    addin.call_sync.return_value = json.dumps(response) if isinstance(response, dict) else response

    result, state = _run(addin)

    assert state == "refused"
    assert result["success"] is False
    assert result["error_pattern"] == "policy_unconfirmed"
    addin.call_sync.assert_called_once_with("SetOperationPolicy", "block")


def test_confirmation_is_normalised_and_run_proceeds():
    addin = Mock()
    replies = {
        "SetOperationPolicy": json.dumps({"success": True, "policy": " Block "}),
        "ExportObject": json.dumps({"success": True}),
        "ClearOperationPolicy": json.dumps({"success": True}),
    }
    addin.call_sync.side_effect = lambda c, *a: replies[c]

    result, state = _run(addin)

    assert state == "completed"
    assert result == {"success": True}
    assert [c.args[0] for c in addin.call_sync.call_args_list] == [
        "SetOperationPolicy", "ExportObject", "ClearOperationPolicy",
    ]


def test_add_in_refusal_passes_through_unchanged():
    refusal = {"success": False, "error_pattern": "operation_already_running", "error": "busy"}
    addin = Mock()
    addin.call_sync.return_value = json.dumps(refusal)

    result, state = _run(addin)

    assert (result, state) == (refusal, "refused")
    addin.call_sync.assert_called_once()


@pytest.mark.parametrize("response", [None, "", "not JSON", {}, {"success": 1}])
def test_unconfirmed_cleanup_is_a_secondary_error(response):
    addin = Mock()
    replies = {
        "SetOperationPolicy": json.dumps(_CONFIRMED),
        "ExportObject": json.dumps({"success": True, "logPath": "x"}),
    }
    addin.call_sync.side_effect = lambda c, *a: (
        (json.dumps(response) if isinstance(response, dict) else response)
        if c == "ClearOperationPolicy" else replies[c]
    )

    with patch("msaccess_vcs_mcp.decision_policy.log_policy_cleanup_failed") as log:
        result, state = _run(addin)

    assert state == "completed"
    assert result["success"] is True and result["logPath"] == "x"
    assert result["policy_cleanup_error"]
    log.assert_called_once()


def test_clear_requires_success_true():
    addin = Mock()
    addin.call_sync.return_value = json.dumps({"success": True})
    assert clear_operation_policy(addin) is None


def test_older_busy_add_in_never_reports_success():
    """The older add-in returns VBA Empty (None) on every call while busy."""
    addin = Mock()
    addin.call_sync.return_value = None

    result, state = _run(addin)

    assert result["success"] is False
    assert result["error_pattern"] == "policy_unconfirmed"
    assert state == "refused"
    # Cleanup alone, were it reached, must not read as cleared either.
    with patch("msaccess_vcs_mcp.decision_policy.log_policy_cleanup_failed"):
        assert clear_operation_policy(addin)


def test_public_tool_with_older_add_in_returning_empty_everywhere(tmp_path):
    from tests.test_import_decisions import _patch_import_tool, _run

    with _patch_import_tool(tmp_path) as (db, src, addin, ops):
        addin.call_sync.return_value = None  # set, operation and clear would all be Empty
        result = _run((db, src), object_types=["forms"])

    assert result["success"] is False
    assert result["error_pattern"] == "policy_unconfirmed"
    assert [c.args[0] for c in addin.call_sync.call_args_list] == ["SetOperationPolicy"]
    addin.call_async.assert_not_called()
    addin.merge_build.assert_not_called()
