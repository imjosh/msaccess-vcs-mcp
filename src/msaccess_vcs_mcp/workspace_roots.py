"""Validate roots with a narrow adapter for native Windows drive paths."""

from pathlib import PureWindowsPath
import re
from typing import Any

from mcp import types
from pydantic import model_validator

from .usage_logging import log_diagnostic_event


class WorkspaceRootsResult(types.ListRootsResult):
    """Normalize drive-absolute paths before the SDK validates file URIs.

    This result type is local to our roots request. SDK models and validation
    for every other response remain unchanged.
    """

    @model_validator(mode="before")
    @classmethod
    def normalize_windows_roots(cls, value: Any) -> Any:
        if not isinstance(value, dict) or not isinstance(value.get("roots"), list):
            return value
        roots = []
        converted = 0
        for root in value["roots"]:
            if isinstance(root, dict):
                uri = root.get("uri")
                if isinstance(uri, str) and re.match(r"^[A-Za-z]:[\\/]", uri):
                    # Reject device paths, controls and characters illegal in a
                    # Windows path. Relative paths and UNC paths are not adapted.
                    tail = uri[2:]
                    if not any(ord(char) < 32 or char in '<>:"|?*' for char in tail):
                        root = {**root, "uri": PureWindowsPath(uri).as_uri()}
                        converted += 1
            roots.append(root)
        if converted:
            log_diagnostic_event("workspace_roots_normalized", count=converted,
                                 compatibility="windows_drive_absolute")
        return {**value, "roots": roots}


async def list_workspace_roots(session) -> types.ListRootsResult:
    """Use the SDK's public request API with our validating result subtype."""
    return await session.send_request(
        types.ServerRequest(types.ListRootsRequest()), WorkspaceRootsResult,
    )
