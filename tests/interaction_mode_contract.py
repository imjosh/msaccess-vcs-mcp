"""A24 SetInteractionMode responses; shared by tool regression fixtures."""

INTERACTIVE_CONFIRMED = {"success": True, "requested_mode": 0, "effective_mode": 0}
INTERACTIVE_REFUSED = {
    "success": False,
    "requested_mode": 0,
    "effective_mode": 2,
    "error_pattern": "interaction_mode_refused",
    "error": "The requested interaction mode could not take effect. An enclosing noninteractive scope or an active operation must be released by its owner first.",
}
