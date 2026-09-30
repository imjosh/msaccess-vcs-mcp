"""Decision policy and add-in decision results for the tools that run a noninteractive operation.

``vcs_import_objects`` and ``vcs_run_tests`` pick a ``decision_policy``, hand it to
the add-in, and turn an uncovered prompt into ``error_pattern: decision_required``.
Everything about that lives here: validating the policy, the interactive-mode and
policy-clear calls, normalising the ``decisions`` an add-in result carries, and
reading the start result of an async call. ``tools.py`` keeps the tool handlers and
calls in. Setting the policy stays in ``vcs_import_objects``, which returns the
add-in's refusal as its own result.
"""

from __future__ import annotations

import json
from typing import Any

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
    result, or None; on a refusal the operation must not start.
    """
    if policy:
        return None
    selected = _as_dict(addin.call_sync("SetInteractionMode", _INTERACTION_MODE_NORMAL))
    if selected.get("success") is not False:
        return None
    return {
        **selected,
        "error": str(selected.get("error") or selected.get("message") or "SetInteractionMode refused"),
    }


def policy_args(policy: str | None) -> tuple[str, ...]:
    """The trailing add-in argument for a call that takes an optional policy."""
    return (policy,) if policy else ()


def _as_dict(raw: Any) -> dict[str, Any]:
    """An add-in return as a dict; anything that is not a JSON object counts as no refusal."""
    if isinstance(raw, str):
        try:
            raw = json.loads(raw)
        except json.JSONDecodeError:
            return {}
    return raw if isinstance(raw, dict) else {}


def clear_operation_policy(addin: Any) -> str | None:
    """Clear the add-in operation policy. Never raises; returns a failure message.

    ``ClearOperationPolicy`` is idempotent, so this is safe after the add-in's
    ``Finish`` has already restored the mode. A failure is written to the
    usage and diagnostic logs so a policy left set in the add-in is visible.
    """
    try:
        cleared = _as_dict(addin.call_sync("ClearOperationPolicy"))
        if cleared.get("success") is False:
            message = str(cleared.get("error") or cleared.get("message") or "ClearOperationPolicy refused")
        else:
            return None
    except Exception as e:
        message = str(e) or type(e).__name__
    log_diagnostic_event("policy_cleanup_failed", error=message)
    log_policy_cleanup_failed(message)
    return message


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
