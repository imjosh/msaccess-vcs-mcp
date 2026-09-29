"""Every import path returns decision_required through the shared normaliser."""

import asyncio
import json

import pytest

from msaccess_vcs_mcp.tools import vcs_import_objects
from tests.test_scoped_sync import _patch_import_tool, _unwrap_sync

DECISIONS = [{"object": "Form1", "type": "conflict"}]
BLOCKED = {
    "success": False,
    "error_pattern": "decision_required",
    "decision_required": True,
    "decisions": DECISIONS,
    "error": "Conflict needs a decision",
}


def _run(tmp_path_args, **kwargs):
    db_path, src_path = tmp_path_args
    return asyncio.run(_unwrap_sync(vcs_import_objects)(db_path, str(src_path), **kwargs))


def _assert_decision(result):
    assert result["success"] is False
    assert result["decision_required"] is True
    assert result["error_pattern"] == "decision_required"
    assert result["decisions"] == DECISIONS


def _wait_returning(payload):
    async def _wait(*_a, **_k):
        return payload
    return _wait


def test_async_completion_decision_required(tmp_path):
    with _patch_import_tool(tmp_path, async_result={"async": True, "timeout_ms": 1000}) as (
        db, src, addin, ops,
    ):
        ops.wait_for_completion = _wait_returning(dict(BLOCKED))
        _assert_decision(_run((db, src)))


def test_async_completion_success_flag_cannot_hide_decision(tmp_path):
    payload = {**BLOCKED, "success": True}
    with _patch_import_tool(tmp_path, async_result={"async": True, "timeout_ms": 1000}) as (
        db, src, addin, ops,
    ):
        ops.wait_for_completion = _wait_returning(payload)
        _assert_decision(_run((db, src)))


def test_inline_sync_marker_decision_required(tmp_path):
    with _patch_import_tool(
        tmp_path, async_result={"sync": True, "result": json.dumps(BLOCKED)}
    ) as (db, src, addin, ops):
        _assert_decision(_run((db, src)))
        addin.merge_build.assert_not_called()


def test_inline_sync_marker_without_result_is_not_success(tmp_path):
    with _patch_import_tool(tmp_path, async_result={"sync": True}) as (db, src, addin, ops):
        assert _run((db, src))["success"] is False


def test_sync_fallback_decision_required(tmp_path):
    with _patch_import_tool(tmp_path, async_result={"unexpected": True}) as (db, src, addin, ops):
        addin.merge_build.return_value = dict(BLOCKED)
        _assert_decision(_run((db, src)))


def test_async_exception_fallback_decision_required(tmp_path):
    with _patch_import_tool(tmp_path) as (db, src, addin, ops):
        addin.call_async.side_effect = RuntimeError("boom")
        addin.merge_build.return_value = dict(BLOCKED)
        _assert_decision(_run((db, src)))


def test_no_callback_decision_required(tmp_path):
    from unittest.mock import patch

    with _patch_import_tool(tmp_path) as (db, src, addin, ops):
        addin.merge_build.return_value = dict(BLOCKED)
        with patch("msaccess_vcs_mcp.tools.get_callback_url", return_value=None):
            _assert_decision(_run((db, src)))
        addin.call_async.assert_not_called()


def test_category_scoped_decision_required(tmp_path):
    with _patch_import_tool(tmp_path, call_sync_result=json.dumps(BLOCKED)) as (
        db, src, addin, ops,
    ):
        _assert_decision(_run((db, src), object_types=["forms"]))


def test_started_marker_is_not_a_final_success_with_callback(tmp_path):
    """A started start-result plus a callback refusal yields the callback's answer."""
    with _patch_import_tool(tmp_path, async_result={"async": True, "timeout_ms": 1000}) as (
        db, src, addin, ops,
    ):
        ops.wait_for_completion = _wait_returning(dict(BLOCKED))
        result = _run((db, src))
    assert result["success"] is False


@pytest.mark.parametrize(
    "pattern",
    ["invalid_decision_policy", "merge_not_available", "operation_already_running"],
)
def test_new_error_patterns_pass_through(tmp_path, pattern):
    refusal = {"success": False, "error_pattern": pattern, "error": "refused"}
    with _patch_import_tool(tmp_path, async_result=refusal) as (db, src, addin, ops):
        result = _run((db, src))
    assert result["success"] is False
    assert result["error_pattern"] == pattern
    assert result["error"] == "refused"
    addin.merge_build.assert_not_called()


@pytest.mark.parametrize("object_types", [None, ["forms"]])
def test_invalid_policy_refused_with_pattern_before_access(tmp_path, object_types):
    with _patch_import_tool(tmp_path) as (db, src, addin, ops):
        result = _run((db, src), decision_policy="bogus", object_types=object_types)
    assert result["success"] is False
    assert result["error_pattern"] == "invalid_decision_policy"
    assert "prefer_source" in result["error"]
    addin.merge_build.assert_not_called()
    addin.call_sync.assert_not_called()


def test_run_tests_invalid_policy_refused_with_pattern():
    from msaccess_vcs_mcp.tools import vcs_run_tests

    fn = vcs_run_tests
    while hasattr(fn, "__wrapped__"):
        fn = fn.__wrapped__
    result = asyncio.run(fn(r"C:\x\y.accdb", decision_policy="bogus"))
    assert result["success"] is False
    assert result["error_pattern"] == "invalid_decision_policy"


def test_duplicate_refusal_returned_once(tmp_path):
    """A refusal delivered synchronously and again by callback is one result."""
    refusal = {
        "success": False, "error_pattern": "operation_already_running", "error": "busy",
    }
    with _patch_import_tool(tmp_path, async_result=refusal) as (db, src, addin, ops):
        result = _run((db, src))
        # The operation is unregistered so a late callback is dropped, and the
        # merge is not retried or waited on.
        ops.unregister_operation.assert_called_once_with("op-1")
        ops.wait_for_completion = None
    assert result["error_pattern"] == "operation_already_running"
    addin.merge_build.assert_not_called()
    addin.call_async.assert_called_once()


def _scripted_sync(import_result, clear_error=None):
    """call_sync stand-in: policy set succeeds, import and clear are scripted."""
    def _call(command, *_args):
        if command == "SetOperationPolicy":
            return json.dumps({"success": True})
        if command == "ImportByType":
            if isinstance(import_result, Exception):
                raise import_result
            return json.dumps(import_result)
        if command == "ClearOperationPolicy":
            if clear_error:
                raise clear_error
            return json.dumps({"success": True})
        return None
    return _call


def test_operation_and_cleanup_both_raise_returns_original_error(tmp_path):
    with _patch_import_tool(tmp_path) as (db, src, addin, ops):
        addin.call_sync.side_effect = _scripted_sync(
            RuntimeError("import exploded"), clear_error=RuntimeError("clear exploded")
        )
        result = _run((db, src), object_types=["forms"])
    assert result["success"] is False
    assert result["error"] == "import exploded"
    assert result["policy_cleanup_error"] == "clear exploded"


def test_cleanup_failure_is_attached_and_logged_on_success(tmp_path):
    from unittest.mock import patch

    with _patch_import_tool(tmp_path) as (db, src, addin, ops):
        addin.call_sync.side_effect = _scripted_sync(
            {"success": True}, clear_error=RuntimeError("clear exploded")
        )
        with patch("msaccess_vcs_mcp.tools.log_diagnostic_event") as log:
            result = _run((db, src), object_types=["forms"])
    assert result["success"] is True
    assert result["policy_cleanup_error"] == "clear exploded"
    log.assert_called_once_with("policy_cleanup_failed", error="clear exploded")


def test_policy_set_refusal_is_a_normal_result(tmp_path):
    refusal = {"success": False, "error_pattern": "operation_already_running", "error": "busy"}
    with _patch_import_tool(tmp_path, call_sync_result=json.dumps(refusal)) as (db, src, addin, ops):
        result = _run((db, src), object_types=["forms"])
    assert result["error_pattern"] == "operation_already_running"
    commands = [c.args[0] for c in addin.call_sync.call_args_list]
    assert "ImportByType" not in commands


@pytest.mark.parametrize("scoped", [False, True])
def test_interactive_sends_explicit_mode_before_operation(tmp_path, scoped):
    kwargs = {"noninteractive": False}
    if scoped:
        kwargs["object_types"] = ["forms"]
    with _patch_import_tool(tmp_path, call_sync_result=json.dumps({"success": True})) as (
        db, src, addin, ops,
    ):
        events = []
        addin.call_sync.side_effect = lambda c, *a: events.append(c) or json.dumps({"success": True})
        addin.call_async.side_effect = lambda *a: events.append("async") or {"async": True}
        addin.merge_build.side_effect = lambda *a: events.append("merge") or {"success": True}
        _run((db, src), **kwargs)
    assert events[0] == "SetInteractionMode"
    assert "SetOperationPolicy" not in events
    addin.call_sync.assert_any_call("SetInteractionMode", 0)


def test_noninteractive_never_sends_interactive_mode(tmp_path):
    with _patch_import_tool(tmp_path) as (db, src, addin, ops):
        addin.call_sync.side_effect = _scripted_sync({"success": True})
        _run((db, src), object_types=["forms"])
    commands = [c.args[0] for c in addin.call_sync.call_args_list]
    assert "SetInteractionMode" not in commands
