"""Decision policy and add-in decision results for the tools that run a noninteractive operation.

``vcs_import_objects``, ``vcs_import_object``, ``vcs_export_object`` and
``vcs_run_tests`` pick a ``decision_policy``, hand it to the add-in, and turn an
uncovered prompt into ``error_pattern: decision_required``.
Everything about that lives here: validating the policy, the interactive-mode and
policy-clear calls, normalising the ``decisions`` an add-in result carries, and
reading the start result of an async call. ``tools.py`` keeps the tool handlers and
calls in. ``call_under_policy`` owns session-policy setup, execution and cleanup;
handlers supply their result parser and keep tool-specific response formatting.
"""

from __future__ import annotations

import json
from typing import TYPE_CHECKING, Any, Callable

if TYPE_CHECKING:
    from .addin_integration import VCSAddinIntegration

from .usage_logging import log_diagnostic_event, log_policy_cleanup_failed

DECISION_POLICIES = frozenset({
    "block",
    "prefer_source",
    "prefer_database",
    "skip",
    "decline",
})


class InvalidDecisionPolicy(ValueError):
    """An unknown ``decision_policy``; refused before Access is called."""
    # An exception rather than a return value: it is raised from ``noninteractive_policy``
    # at the top of two tool bodies and turned into a result by their ``except``.

    error_pattern = "invalid_decision_policy"


def invalid_policy_result(exc: InvalidDecisionPolicy, **extra: Any) -> dict[str, Any]:
    return {"success": False, "error": str(exc), "error_pattern": exc.error_pattern, **extra}


def noninteractive_policy(noninteractive: bool, decision_policy: str | None) -> str | None:
    """Return the add-in policy name, or None to keep interactive behavior.

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


_INTERACTION_MODE_NORMAL = 0  # add-in eInteractionMode.eimNormal


def select_interactive_mode(addin: Any, policy: str | None) -> dict[str, Any] | None:
    """Send an explicit interactive mode, so the run never inherits a stale one.

    Does nothing when ``policy`` is set: a noninteractive run is scoped by the
    add-in itself. Call before the operation starts. Returns the add-in's
    refusal (for example from an enclosing noninteractive scope) as the tool
    result, or None on confirmed acceptance. A24 or later must report success
    and effective mode 0; older Empty returns cannot confirm interactive mode.
    Never clear a caller-owned policy to make this selection succeed.
    """
    if policy:
        return None
    selected = _as_dict(addin.call_sync("SetInteractionMode", _INTERACTION_MODE_NORMAL))
    if (
        selected.get("success") is True
        and type(selected.get("effective_mode")) is int
        and selected["effective_mode"] == _INTERACTION_MODE_NORMAL
    ):
        return None
    if selected.get("success") is not False:
        return {
            "success": False,
            "error_pattern": "interaction_mode_unconfirmed",
            "error": (
                "SetInteractionMode did not confirm interactive mode; no operation started. "
                "An A24 or later add-in build is required. Upgrade the add-in if its response "
                "is empty or unsupported; an enclosing policy must be cleared by its owner."
            ),
        }
    return {
        **selected,
        "error": str(selected.get("error") or selected.get("message") or "SetInteractionMode refused"),
    }


def policy_args(policy: str | None) -> tuple[str, ...]:
    """The trailing add-in argument for a call that takes an optional policy."""
    return (policy,) if policy else ()


def _as_dict(raw: Any) -> dict[str, Any]:
    """An add-in return as a dict; anything that is not a JSON object becomes empty."""
    if isinstance(raw, str):
        try:
            raw = json.loads(raw)
        except json.JSONDecodeError:
            return {}
    return raw if isinstance(raw, dict) else {}


def _policy_unconfirmed(policy: str) -> dict[str, Any]:
    return {
        "success": False,
        "error_pattern": "policy_unconfirmed",
        "error": (
            f"SetOperationPolicy did not confirm the {policy} policy; no operation started. "
            "An add-in build whose SetOperationPolicy returns the policy it set is required. "
            "Upgrade the add-in if its response is empty or unsupported."
        ),
    }


def _policy_confirmed(raw: Any, policy: str) -> bool:
    """True only for ``{success: true, policy: <requested policy, normalised>}``."""
    reply = _as_dict(raw)
    echoed = reply.get("policy")
    return (
        reply.get("success") is True
        and isinstance(echoed, str)
        and echoed.strip().lower() == policy
    )


def clear_operation_policy(addin: Any) -> str | None:
    """Clear the add-in session policy. Never raises; returns a failure message or None.

    A session policy set through ``SetOperationPolicy`` is caller-owned: the
    add-in's ``Finish`` does not close it, so this call is what restores the
    interaction mode after the operation. ``ClearOperationPolicy`` is
    idempotent (a second clear, or a clear when none is set, is a no-op), so a
    caller can always call it from ``finally``. A failure is written to the
    usage log (``policy_cleanup_failed``) and the diagnostic log, and the
    caller attaches it to the tool result as ``policy_cleanup_error``; it is
    never allowed to replace the operation's own result. Only
    ``{success: true}`` counts as cleared; Empty or malformed returns do not.
    """
    try:
        cleared = _as_dict(addin.call_sync("ClearOperationPolicy"))
        if cleared.get("success") is True:
            return None
        message = str(
            cleared.get("error") or cleared.get("message")
            or "ClearOperationPolicy did not confirm the policy was cleared"
        )
    except Exception as e:
        message = str(e) or type(e).__name__
    log_diagnostic_event("policy_cleanup_failed", error=message)
    log_policy_cleanup_failed(message)
    return message


def call_under_policy(
    addin: VCSAddinIntegration, policy: str | None, command: str, *args: Any,
    parse_result: Callable[[Any], dict[str, Any]],
) -> tuple[dict[str, Any], str]:
    """
    Run one sync add-in call under ``policy``, or in explicit interactive mode.

    ``parse_result`` keeps tool-specific result formatting with the caller.
    Sets the policy through ``SetOperationPolicy`` first, proceeding only on a
    confirmed ``{success: true, policy}`` reply (else ``policy_unconfirmed``), and clears it in
    ``finally``, so the call is covered whether or not the add-in keeps a
    session policy past ``Finish``. A cleanup failure is attached as
    ``policy_cleanup_error`` and never replaces the call's result.

    Returns the result and how the call went: ``"refused"`` when the policy
    set or the interactive-mode request was refused (the refusal is the
    result and the call never started), ``"raised"`` when the call raised
    (the result is a plain failure with the exception text), otherwise
    ``"completed"`` with the add-in's parsed result.
    """
    if policy:
        # A refusal here (for example operation_already_running) is a normal
        # result. Anything but a confirmed set fails closed with no dispatch
        # and no cleanup call: nothing was set, so there is nothing to clear.
        raw = addin.call_sync("SetOperationPolicy", policy)
        if not _policy_confirmed(raw, policy):
            if _as_dict(raw).get("success") is False:
                return parse_result(raw), "refused"
            return _policy_unconfirmed(policy), "refused"
    mode_refusal = select_interactive_mode(addin, policy)
    if mode_refusal:
        return mode_refusal, "refused"
    state = "completed"
    try:
        result = parse_result(addin.call_sync(command, *args))
    except Exception as e:
        state = "raised"
        result = {"success": False, "error": str(e)}
    finally:
        cleanup_error = clear_operation_policy(addin) if policy else None
    if cleanup_error:
        # Secondary information: never replaces the operation's result.
        result["policy_cleanup_error"] = cleanup_error
    return result, state


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
    result: dict[str, Any] = {"success": success}
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
