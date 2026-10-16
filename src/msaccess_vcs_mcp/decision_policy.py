"""Request policy arguments and extensible operation outcomes.

Each noninteractive operation receives its own explicit policy. The add-in owns
that policy through completion and releases it without caller cleanup.
"""

from __future__ import annotations

import json
from typing import TYPE_CHECKING, Any, Callable

if TYPE_CHECKING:
    from .addin_integration import VCSAddinIntegration


DECISION_POLICIES = frozenset({
    "block",
    "prefer_source",
    "prefer_database",
    "skip",
    "decline",
})


SYNCHRONOUS_OPERATION_METHODS = frozenset(name.casefold() for name in (
    "Export", "FullExport", "ExportVBA", "Build", "MergeBuild", "ImportObject",
    "ExportObject", "ImportByType", "ExportByType", "RunFilteredTests", "RunTests",
    "RunTestsHeadless", "RunRoundtripTests", "BuildHeadless", "MergeHeadless",
    "BuildAs", "RebuildAddIn", "LoadSelected", "MergeAllSource", "ExportSelected",
    "ShowOptions", "RepairColors", "SplitFiles", "LocalizeLibraryReferences",
    "ActivateHook", "StandardizeLetterCasing", "OpenTestRunner",
))


def is_operation_method(command: str) -> bool:
    return str(command).casefold() in SYNCHRONOUS_OPERATION_METHODS


class InvalidDecisionPolicy(ValueError):
    """An unknown ``decision_policy``; refused before Access is called."""
    # An exception rather than a return value: it is raised from ``noninteractive_policy``
    # at the top of two tool bodies and turned into a result by their ``except``.

    error_pattern = "invalid_decision_policy"


def invalid_policy_result(exc: InvalidDecisionPolicy, **extra: Any) -> dict[str, Any]:
    return {"success": False, "error": str(exc), "error_pattern": exc.error_pattern, **extra}


def interactive_tests_unsupported_result() -> dict[str, Any]:
    """The refusal for ``vcs_run_tests(noninteractive=False)``; returned before any add-in call.

    The add-in marks every API call as automation and ``ExecuteTests`` then runs
    headless whatever mode was selected, so an interactive run cannot be had
    through the tool (X11). Refusing is better than a flag that does nothing,
    or worse, shows a prompt that holds the Access gate.
    """
    return {
        "success": False,
        "error_pattern": "interactive_tests_unsupported",
        "error": (
            "vcs_run_tests is always headless: an automation test run shows no console and "
            "answers prompts unattended, so noninteractive=False is refused and nothing was "
            "started. Run with noninteractive=True and a decision_policy, or run the tests "
            "from the add-in's ribbon for the interactive console."
        ),
    }


def noninteractive_policy(noninteractive: bool, decision_policy: str | None) -> str | None:
    """Return the request policy name, or None to retain legacy behavior.

    Raises InvalidDecisionPolicy for an unknown policy. The tool wrappers turn
    that into ``success: false`` with ``error_pattern: invalid_decision_policy``
    without calling Access.
    """
    if not noninteractive:
        return None
    policy = (decision_policy or "block").strip().lower()
    if policy not in DECISION_POLICIES:
        raise InvalidDecisionPolicy(
            "Unknown decision_policy. Use block, prefer_source, prefer_database, skip, or decline."
        )
    return policy


def select_interactive_mode(addin: Any, policy: str | None) -> None:
    """Compatibility helper. An omitted request policy requires no mode mutation."""
    return None


def policy_args(policy: str | None) -> tuple[str, ...]:
    """The trailing argument for an operation that accepts a request policy."""
    return (policy,) if policy else ()


def clear_operation_policy(addin: Any) -> None:
    """Compatibility helper. Request ownership releases policy inside the add-in."""
    return None


def call_under_policy(
    addin: VCSAddinIntegration, policy: str | None, command: str, *args: Any,
    parse_result: Callable[[Any], dict[str, Any]],
) -> tuple[dict[str, Any], str]:
    """Dispatch one request with its trailing policy and preserve its result.

    No setup or cleanup call can mutate another request's mode or journal.
    An omitted policy retains the add-in's documented legacy behavior.
    """
    try:
        result = parse_result(addin.call_sync(command, *args, *policy_args(policy)))
    except Exception as exc:
        return {"success": False, "error": str(exc)}, "raised"
    return result, "refused" if is_start_refusal(result) else "completed"


def coerce_decisions(value: Any) -> Any:
    if isinstance(value, str):
        try:
            return json.loads(value)
        except json.JSONDecodeError:
            return value
    return value


def is_decision_required(payload: dict[str, Any]) -> bool:
    """The add-in reported a prompt the decision policy did not cover."""
    return bool(payload.get("decision_required") or payload.get("error_pattern") == "decision_required")


def apply_decision_result(result: dict[str, Any], completion: dict[str, Any] | None) -> dict[str, Any]:
    """Surface an unresolved add-in prompt. Never leave success true in that case."""
    if not completion:
        return result
    decisions = completion.get("decisions")
    if decisions is not None:
        result["decisions"] = coerce_decisions(decisions)
    if is_decision_required(completion):
        result["success"] = False
        result["decision_required"] = True
        result["error_pattern"] = "decision_required"
        result["error"] = completion.get("error") or completion.get("message") or (
            "A required decision was not covered by the decision policy."
        )
    return result


def surface_own_decision(result: dict[str, Any]) -> dict[str, Any]:
    """Apply the decision fields ``result`` already carries to itself.

    For a sync add-in call, where the result and the completion are one dict.
    """
    return apply_decision_result(result, result)


def parse_addin_payload(value: Any) -> dict[str, Any]:
    """Parse an add-in result (JSON string or dict) into a dict, never raising."""
    if isinstance(value, str):
        try:
            value = json.loads(value)
        except json.JSONDecodeError:
            return {"success": False, "error": f"Unparseable add-in result: {value[:200]}"}
    if isinstance(value, dict):
        return value
    return {"success": False, "error": "Add-in returned no result"}


def is_start_refusal(payload: dict[str, Any]) -> bool:
    """A sync start result that refused (or declined) the operation."""
    return payload.get("success") is False and not payload.get("async") and not payload.get("sync")


_MERGE_COMPLETION_UNCONFIRMED_ERROR = (
    "The merge started, but no completion callback exists on this path to "
    "report its outcome. It may have succeeded or failed. Read log_path, or "
    "call vcs_get_recent_calls() and vcs_get_log(log_type=\"Merge\")."
)


def normalize_terminal_result(payload: dict[str, Any], default_error: str) -> dict[str, Any]:
    """Interpret an add-in completion or refusal, leaving response formatting to the tool.

    Keeps ``error_pattern``, ``decisions``, ``decision_required``, ``cancelled``,
    ``runtime_error`` and ``errorNumber`` from the payload. Decisions override
    success. An unconfirmed start retains its metadata and guidance, and never
    reports success. ``default_error`` names the failure when the payload has no
    text. Log aliases stay with the handler.
    """
    unconfirmed = bool(payload.get("completion_unconfirmed"))
    success = payload.get("success") is True and not unconfirmed
    # Keep extensible journals, transfer outcomes, identity, and future fields.
    result: dict[str, Any] = {**payload, "success": success}
    for key in ("decisions", "error_pattern", "decision_required"):
        if result.get(key) is None:
            result.pop(key, None)
    if payload.get("cancelled"):
        result["cancelled"] = True
    if unconfirmed:
        result["started"] = True
        result["completion_unconfirmed"] = True
        result["error"] = _MERGE_COMPLETION_UNCONFIRMED_ERROR
    elif not success:
        result["error"] = payload.get("error") or payload.get("message") or default_error
    if payload.get("error_pattern"):
        result["error_pattern"] = payload["error_pattern"]
    if payload.get("runtime_error"):
        result["runtime_error"] = payload["runtime_error"]
    if payload.get("errorNumber") is not None:
        result["errorNumber"] = payload["errorNumber"]
    return apply_decision_result(result, payload)


def normalize_import_result(payload: dict[str, Any]) -> dict[str, Any]:
    """Interpret a merge completion or refusal."""
    return normalize_terminal_result(payload, "Import failed")


async def run_import_merge(
    addin: VCSAddinIntegration, database_path: str, source_dir: str, policy: str | None,
    *, callback_info: str | None = None, op_manager: Any = None,
    operation_id: str | None = None, ctx: Any = None,
) -> dict[str, Any]:
    """Resolve every full-merge start path to its last add-in payload.

    The handler registers the callback operation. Here we follow async starts,
    parse inline completions, and return refusals without retrying. Missing
    markers or an async exception fall back to sync, whose start is unconfirmed.
    """
    def merge_sync() -> dict[str, Any]:
        merged = addin.merge_build(database_path, source_dir, policy)
        if merged.get("started") and merged.get("success"):
            merged = {**merged, "completion_unconfirmed": True}
        return merged

    if op_manager is None:
        return merge_sync()
    try:
        started = addin.call_async(callback_info, "MergeBuild", *policy_args(policy))
        if started.get("async"):
            return await op_manager.wait_for_completion(
                operation_id, ctx=ctx,
                timeout_seconds=started.get("timeout_ms", 300000) / 1000,
            )
        op_manager.unregister_operation(operation_id)
        if started.get("sync"):
            # Already completed inline: never merge twice.
            return parse_addin_payload(started.get("result"))
        if is_start_refusal(started):
            # Drop the duplicate refusal delivered by callback.
            return started
        return merge_sync()
    except Exception:
        op_manager.unregister_operation(operation_id)
        return merge_sync()
