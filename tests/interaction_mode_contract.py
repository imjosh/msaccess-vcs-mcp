"""A24 SetInteractionMode responses; shared by tool regression fixtures."""

INTERACTIVE_CONFIRMED = {"success": True, "requested_mode": 0, "effective_mode": 0}
INTERACTIVE_REFUSED = {
    "success": False,
    "requested_mode": 0,
    "effective_mode": 2,
    "error_pattern": "interaction_mode_refused",
    "error": "The requested interaction mode could not take effect. An enclosing noninteractive scope or an active operation must be released by its owner first.",
}


def policy_aware(reply):
    """call_sync stand-in: SetOperationPolicy echoes its policy (the real contract), else ``reply``."""
    import json

    def call(command, *args):
        if command == "SetOperationPolicy":
            return json.dumps({"success": True, "policy": args[0]})
        return reply
    return call
