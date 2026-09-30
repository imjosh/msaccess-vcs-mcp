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
        with patch("msaccess_vcs_mcp.decision_policy.log_diagnostic_event") as log:
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


def test_cleanup_failure_reaches_usage_log_and_keeps_operation_error(tmp_path, monkeypatch):
    import msaccess_vcs_mcp.usage_logging as usage_logging

    log_dir = tmp_path / "usage"
    monkeypatch.setenv("ACCESS_VCS_ENABLE_LOGGING", "true")
    monkeypatch.setenv("ACCESS_VCS_LOG_DIR", str(log_dir))
    monkeypatch.setenv("ACCESS_VCS_DIAGNOSTIC_LOG_DIR", str(tmp_path / "diag"))
    usage_logging.reset_logging()
    try:
        with _patch_import_tool(tmp_path) as (db, src, addin, ops):
            addin.call_sync.side_effect = _scripted_sync(
                {"success": False, "error": "merge failed"},
                clear_error=RuntimeError("clear exploded"),
            )
            result = _run((db, src), object_types=["forms"])
    finally:
        usage_logging.reset_logging()

    assert result["success"] is False
    assert result["error"] == "merge failed"
    assert result["policy_cleanup_error"] == "clear exploded"
    entries = [
        json.loads(line)
        for line in (log_dir / "vcs-mcp-usage.jsonl").read_text(encoding="utf-8").splitlines()
    ]
    cleanup = [e for e in entries if e.get("event") == "policy_cleanup_failed"]
    assert len(cleanup) == 1
    assert cleanup[0]["error"] == "clear exploded"


INTERACTION_REFUSAL = {"success": False, "error": "A noninteractive scope is open"}


@pytest.mark.parametrize("scoped", [False, True])
def test_refused_interactive_mode_is_the_result_and_nothing_starts(tmp_path, scoped):
    kwargs = {"noninteractive": False}
    if scoped:
        kwargs["object_types"] = ["forms"]
    with _patch_import_tool(tmp_path) as (db, src, addin, ops):
        addin.call_sync.side_effect = lambda c, *a: (
            json.dumps(INTERACTION_REFUSAL) if c == "SetInteractionMode" else json.dumps({"success": True})
        )
        result = _run((db, src), **kwargs)
    assert result["success"] is False
    assert result["error"] == "A noninteractive scope is open"
    commands = [c.args[0] for c in addin.call_sync.call_args_list]
    assert commands == ["SetInteractionMode"]
    addin.call_async.assert_not_called()
    addin.merge_build.assert_not_called()


STARTED = {"success": True, "started": True, "operation_id": "op-9"}


def _assert_unconfirmed(result):
    assert result["success"] is not True
    assert result["started"] is True
    assert result["completion_unconfirmed"] is True
    assert "vcs_get_recent_calls" in result["error"]


def test_started_marker_on_sync_fallback_is_not_success(tmp_path):
    with _patch_import_tool(tmp_path, async_result={"unexpected": True}) as (db, src, addin, ops):
        addin.merge_build.return_value = dict(STARTED)
        _assert_unconfirmed(_run((db, src)))


def test_started_marker_on_async_exception_fallback_is_not_success(tmp_path):
    with _patch_import_tool(tmp_path) as (db, src, addin, ops):
        addin.call_async.side_effect = RuntimeError("boom")
        addin.merge_build.return_value = dict(STARTED)
        _assert_unconfirmed(_run((db, src)))


def test_started_marker_without_callback_is_not_success(tmp_path):
    from unittest.mock import patch

    with _patch_import_tool(tmp_path) as (db, src, addin, ops):
        addin.merge_build.return_value = dict(STARTED)
        with patch("msaccess_vcs_mcp.tools.get_callback_url", return_value=None):
            _assert_unconfirmed(_run((db, src)))
        addin.call_async.assert_not_called()


def test_inline_sync_success_still_normalises_as_success(tmp_path):
    with _patch_import_tool(
        tmp_path, async_result={"sync": True, "result": json.dumps({"success": True})}
    ) as (db, src, addin, ops):
        result = _run((db, src))
    assert result["success"] is True
    assert result["imported_count"] == "See log for details"
    assert "completion_unconfirmed" not in result
    assert "error" not in result


def test_callback_success_still_normalises_as_success(tmp_path):
    with _patch_import_tool(tmp_path, async_result={"async": True, "timeout_ms": 1000}) as (
        db, src, addin, ops,
    ):
        ops.wait_for_completion = _wait_returning({"success": True, "log_path": None})
        result = _run((db, src))
    assert result["success"] is True
    assert result["imported_count"] == "See log for details"
    assert "completion_unconfirmed" not in result
    assert "error" not in result


# vcs_import_object and vcs_export_object run under the same policy set and clear as
# the category-scoped merge (X04).

SINGLE_OBJECT_TOOLS = [
    ("vcs_import_object", "ImportObject"),
    ("vcs_export_object", "ExportObject"),
]

MERGE_REFUSED = {
    "success": False,
    "error": "Import completed with errors: ERROR: Merging not supported for add-in forms.",
}


def _single_object_tool(tool_name):
    import msaccess_vcs_mcp.tools as tools

    return _unwrap_sync(getattr(tools, tool_name))


def _single_object_sync(command, result, *, set_result=None, clear_error=None, events=None):
    """call_sync stand-in for a single-object call; records each command in ``events``."""
    def _call(name, *args):
        if events is not None:
            events.append((name, *args))
        if name == "SetOperationPolicy":
            return json.dumps(set_result or {"success": True})
        if name == command:
            if isinstance(result, Exception):
                raise result
            return json.dumps(result)
        if name == "ClearOperationPolicy":
            if clear_error:
                raise clear_error
            return json.dumps({"success": True})
        return None
    return _call


@pytest.mark.parametrize("tool_name,command", SINGLE_OBJECT_TOOLS)
def test_single_object_sets_block_policy_and_clears_it(tmp_path, tool_name, command):
    events = []
    with _patch_import_tool(tmp_path) as (db, src, addin, ops):
        addin.call_sync.side_effect = _single_object_sync(command, {"success": True}, events=events)
        result = _single_object_tool(tool_name)(db, "form", "frmVCSMain")
    assert result["success"] is True
    assert events == [
        ("SetOperationPolicy", "block"),
        (command, "form", "frmVCSMain"),
        ("ClearOperationPolicy",),
    ]


@pytest.mark.parametrize("tool_name,command", SINGLE_OBJECT_TOOLS)
def test_single_object_passes_decision_policy(tmp_path, tool_name, command):
    events = []
    with _patch_import_tool(tmp_path) as (db, src, addin, ops):
        addin.call_sync.side_effect = _single_object_sync(command, {"success": True}, events=events)
        _single_object_tool(tool_name)(db, "query", "qryFoo", decision_policy="prefer_source")
    assert events[0] == ("SetOperationPolicy", "prefer_source")


@pytest.mark.parametrize("tool_name,command", SINGLE_OBJECT_TOOLS)
def test_single_object_failure_is_returned_and_policy_cleared(tmp_path, tool_name, command):
    events = []
    with _patch_import_tool(tmp_path) as (db, src, addin, ops):
        addin.call_sync.side_effect = _single_object_sync(command, MERGE_REFUSED, events=events)
        result = _single_object_tool(tool_name)(db, "form", "frmVCSMain")
    assert result["success"] is False
    assert "Merging not supported for add-in forms" in result["error"]
    assert events[-1] == ("ClearOperationPolicy",)


@pytest.mark.parametrize("tool_name,command", SINGLE_OBJECT_TOOLS)
def test_single_object_clears_policy_when_the_call_raises(tmp_path, tool_name, command):
    events = []
    with _patch_import_tool(tmp_path) as (db, src, addin, ops):
        addin.call_sync.side_effect = _single_object_sync(
            command, RuntimeError("call exploded"), events=events
        )
        result = _single_object_tool(tool_name)(db, "form", "frmVCSMain")
    assert result["success"] is False
    assert result["error"] == "call exploded"
    assert events[-1] == ("ClearOperationPolicy",)


@pytest.mark.parametrize("tool_name,command", SINGLE_OBJECT_TOOLS)
def test_single_object_call_and_cleanup_both_raise_returns_original_error(
    tmp_path, tool_name, command,
):
    with _patch_import_tool(tmp_path) as (db, src, addin, ops):
        addin.call_sync.side_effect = _single_object_sync(
            command, RuntimeError("call exploded"), clear_error=RuntimeError("clear exploded")
        )
        result = _single_object_tool(tool_name)(db, "form", "frmVCSMain")
    assert result["success"] is False
    assert result["error"] == "call exploded"
    assert result["policy_cleanup_error"] == "clear exploded"


@pytest.mark.parametrize("tool_name,command", SINGLE_OBJECT_TOOLS)
def test_single_object_policy_set_refusal_is_the_result(tmp_path, tool_name, command):
    events = []
    refusal = {"success": False, "error_pattern": "operation_already_running", "error": "busy"}
    with _patch_import_tool(tmp_path) as (db, src, addin, ops):
        addin.call_sync.side_effect = _single_object_sync(
            command, {"success": True}, set_result=refusal, events=events
        )
        result = _single_object_tool(tool_name)(db, "form", "frmVCSMain")
    assert result["error_pattern"] == "operation_already_running"
    assert command not in [e[0] for e in events]


@pytest.mark.parametrize("tool_name,command", SINGLE_OBJECT_TOOLS)
def test_single_object_decision_required_is_surfaced(tmp_path, tool_name, command):
    with _patch_import_tool(tmp_path) as (db, src, addin, ops):
        addin.call_sync.side_effect = _single_object_sync(command, {**BLOCKED, "success": True})
        result = _single_object_tool(tool_name)(db, "form", "frmVCSMain")
    _assert_decision(result)


# The shape ImportObject and ExportObject return for a blocked prompt (X05), with the
# error from a raised call alongside it.
ADDIN_SINGLE_OBJECT_BLOCKED = {
    "success": False,
    "error_pattern": "decision_required",
    "decision_required": True,
    "error": "A required decision was not covered by the decision policy.",
    "decisions": [{
        "kind": "decision_required",
        "title": "Confirm",
        "message": "Overwrite?",
        "resolution": "decision_required",
    }],
    "runtime_error": "Object variable not set",
    "errorNumber": 91,
    "logPath": r"C:\src\logs\Merge_1.log",
}


@pytest.mark.parametrize("tool_name,command", SINGLE_OBJECT_TOOLS)
def test_single_object_addin_decision_result_is_kept_whole(tmp_path, tool_name, command):
    events = []
    with _patch_import_tool(tmp_path) as (db, src, addin, ops):
        addin.call_sync.side_effect = _single_object_sync(
            command, ADDIN_SINGLE_OBJECT_BLOCKED, events=events
        )
        result = _single_object_tool(tool_name)(db, "module", "modFoo")
    assert result["success"] is False
    assert result["decision_required"] is True
    assert result["error_pattern"] == "decision_required"
    assert result["decisions"] == ADDIN_SINGLE_OBJECT_BLOCKED["decisions"]
    assert result["runtime_error"] == "Object variable not set"
    assert result["errorNumber"] == 91
    assert result["log_path"] == r"C:\src\logs\Merge_1.log"
    assert events[-1] == ("ClearOperationPolicy",)


@pytest.mark.parametrize("tool_name,command", SINGLE_OBJECT_TOOLS)
def test_single_object_interactive_sends_explicit_mode_and_no_policy(tmp_path, tool_name, command):
    events = []
    with _patch_import_tool(tmp_path) as (db, src, addin, ops):
        addin.call_sync.side_effect = _single_object_sync(command, {"success": True}, events=events)
        _single_object_tool(tool_name)(db, "form", "frmVCSMain", noninteractive=False)
    assert events == [("SetInteractionMode", 0), (command, "form", "frmVCSMain")]


@pytest.mark.parametrize("tool_name,command", SINGLE_OBJECT_TOOLS)
def test_single_object_refused_interactive_mode_is_the_result(tmp_path, tool_name, command):
    events = []
    with _patch_import_tool(tmp_path) as (db, src, addin, ops):
        def _call(name, *args):
            events.append(name)
            return json.dumps(INTERACTION_REFUSAL if name == "SetInteractionMode" else {"success": True})
        addin.call_sync.side_effect = _call
        result = _single_object_tool(tool_name)(db, "form", "frmVCSMain", noninteractive=False)
    assert result["error"] == "A noninteractive scope is open"
    assert events == ["SetInteractionMode"]


@pytest.mark.parametrize("tool_name,command", SINGLE_OBJECT_TOOLS)
def test_single_object_invalid_policy_refused_before_access(tmp_path, tool_name, command):
    with _patch_import_tool(tmp_path) as (db, src, addin, ops):
        result = _single_object_tool(tool_name)(db, "form", "frmVCSMain", decision_policy="bogus")
    assert result["success"] is False
    assert result["error_pattern"] == "invalid_decision_policy"
    addin.call_sync.assert_not_called()


@pytest.mark.parametrize("path", ["async", "inline", "fallback", "exception", "no_callback"])
@pytest.mark.parametrize("outcome", ["success", "decision", "refusal", "started"])
def test_public_import_result_contract(tmp_path, monkeypatch, path, outcome):
    """Exercise the gate/config/logging wrapper as well as all merge result paths."""
    from unittest.mock import AsyncMock

    from msaccess_vcs_mcp import tools
    from msaccess_vcs_mcp.access_gate import reset_access_gate
    from msaccess_vcs_mcp.dialog_recovery import reset_interruptions

    reset_access_gate()
    reset_interruptions()
    monkeypatch.setattr(tools, "_ensure_env_loaded", AsyncMock())
    monkeypatch.setattr(tools, "load_config", lambda: {})
    monkeypatch.setenv("ACCESS_VCS_ENABLE_LOGGING", "false")
    monkeypatch.setenv("ACCESS_VCS_DIAGNOSTIC_LOG_DIR", str(tmp_path / "diag"))
    log_path = tmp_path / "Merge.log"
    log_path.write_text("merge details\n", encoding="utf-8")
    payload = {
        "success": True,
        "decisions": json.dumps(DECISIONS),
        "log_path": str(log_path),
    }
    if outcome == "decision":
        payload.update(BLOCKED, success=True, decisions=json.dumps(DECISIONS))
        payload["runtime_error"] = "Original error"
    elif outcome == "refusal":
        payload.update(success=False, error_pattern="merge_not_available", error="refused")
    elif outcome == "started":
        payload.update(STARTED)

    start = {"async": True, "timeout_ms": 1000}
    if path == "inline":
        start = {"sync": True, "result": json.dumps(payload)}
    elif path == "fallback":
        start = {"unexpected": True}
    # A refusal at the async start must never be retried.
    if path == "async" and outcome == "refusal":
        start = dict(payload)
    try:
        with _patch_import_tool(tmp_path, async_result=start) as (db, src, addin, ops):
            ops.wait_for_completion = _wait_returning(dict(payload))
            addin.merge_build.return_value = dict(payload)
            if path == "exception":
                addin.call_async.side_effect = RuntimeError("async unavailable")
            elif path == "no_callback":
                monkeypatch.setattr(tools, "get_callback_url", lambda: None)
            result = asyncio.run(tools.vcs_import_objects(db, str(src)))
            unconfirmed = outcome == "started" and path in {"fallback", "exception", "no_callback"}
            expected = {
                "success": outcome not in {"decision", "refusal"} and not unconfirmed,
                "imported_count": 0 if outcome == "refusal" else "See log for details",
                "database_path": db,
                "source_dir": str(src),
                "decisions": DECISIONS,
                "log_path": str(log_path),
                "log_path": str(log_path),
            }
            if outcome == "decision":
                expected.update(
                    decision_required=True, error_pattern="decision_required",
                    error=BLOCKED["error"], runtime_error="Original error",
                    log_excerpt="merge details",
                )
            elif outcome == "refusal":
                expected.update(error_pattern="merge_not_available", error="refused", log_excerpt="merge details")
            elif unconfirmed:
                _assert_unconfirmed(result)
                expected.update(
                    started=True, completion_unconfirmed=True,
                    error=result["error"], log_excerpt="merge details",
                )
            assert result == expected
            if path in {"async", "inline"}:
                addin.merge_build.assert_not_called()
            else:
                addin.merge_build.assert_called_once_with(db, str(src), "block")
            if path == "async" and outcome != "refusal":
                ops.unregister_operation.assert_not_called()
            elif path != "no_callback":
                ops.unregister_operation.assert_called_once_with("op-1")
    finally:
        reset_interruptions()
        reset_access_gate()
