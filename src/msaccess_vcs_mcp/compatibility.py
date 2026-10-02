"""Consumer-owned SemVer requirements and installed-library admission (X16)."""

from __future__ import annotations

import json
import os
import re
import subprocess
import sys
from dataclasses import dataclass
from functools import total_ordering
from importlib.resources import files
from typing import Any

from . import __version__
from .config import get_config, get_default_addin_path

# Josh's assigned release requirement: no published release is claimed here.
# The contract includes policy/mode acknowledgments, BuildAs(source, output),
# terminal result journals and dispatcher refusals; v5.0.1 does not deliver it.
REQUIREMENT_STATUS = "assigned_unpublished"
DISCOVERY_TIMEOUT_SEC = 10.0

# These tools call the add-in or may register callbacks/mutate a target before
# doing so. CLI dispatch uses the same registered wrappers over stdio.
ADDIN_DEPENDENT_TOOLS = frozenset({
    "vcs_export_database", "vcs_import_objects", "vcs_rebuild_database",
    "vcs_rebuild_addin", "vcs_check_vba_compiled", "vcs_compile_vba",
    "vcs_export_object", "vcs_import_object", "vcs_execute_sql", "vcs_call_vba",
    "vcs_run_vba", "vcs_set_option", "vcs_get_option", "vcs_get_log",
    "vcs_run_tests", "vcs_end_session",
})

_SEMVER = re.compile(
    r"(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)"
    r"(?:-([0-9A-Za-z-]+(?:\.[0-9A-Za-z-]+)*))?"
    r"(?:\+([0-9A-Za-z-]+(?:\.[0-9A-Za-z-]+)*))?"
)


@total_ordering
@dataclass(frozen=True, eq=False)
class Version:
    core: tuple[int, int, int]
    prerelease: tuple[str, ...] = ()

    @classmethod
    def parse(cls, value: Any) -> Version:
        match = _SEMVER.fullmatch(value) if isinstance(value, str) else None
        if match is None:
            raise ValueError("Expected strict SemVer MAJOR.MINOR.PATCH")
        pre = tuple(match[4].split(".")) if match[4] else ()
        if any(p.isdigit() and len(p) > 1 and p.startswith("0") for p in pre):
            raise ValueError("Numeric prerelease identifiers cannot have leading zeroes")
        return cls(tuple(int(match[i]) for i in (1, 2, 3)), pre)

    def __eq__(self, other: object) -> bool:
        if not isinstance(other, Version):
            return NotImplemented
        return (self.core, self.prerelease) == (other.core, other.prerelease)

    def __lt__(self, other: Version) -> bool:
        if not isinstance(other, Version):
            return NotImplemented
        if self.core != other.core:
            return self.core < other.core
        if not self.prerelease or not other.prerelease:
            return bool(self.prerelease) and not other.prerelease
        for left, right in zip(self.prerelease, other.prerelease):
            if left == right:
                continue
            if left.isdigit() and right.isdigit():
                return int(left) < int(right)
            if left.isdigit() != right.isdigit():
                return left.isdigit()
            return left < right
        return len(self.prerelease) < len(other.prerelease)


@dataclass(frozen=True)
class Requirement:
    """Half-open stable range plus explicitly enumerated prerelease identities.

    Syntax: ``>=MIN <MAX; prerelease=V1,V2``. No wildcards, OR, normalization,
    or implicit prerelease admission. Major zero is bounded to one minor.
    Build metadata never contributes to comparison or prerelease identity.
    """

    minimum: str
    maximum: str
    prereleases: tuple[str, ...] = ()

    def __post_init__(self) -> None:
        low, high = Version.parse(self.minimum), Version.parse(self.maximum)
        boundary = (low.core[0] + 1, 0, 0) if low.core[0] else (0, low.core[1] + 1, 0)
        if low.prerelease or high.prerelease or high.core != boundary or not low < high:
            raise ValueError("Range must end at the next stable major or major-zero minor")
        for item in self.prereleases:
            pre = Version.parse(item)
            if not pre.prerelease or not low.core <= pre.core < high.core:
                raise ValueError("Admitted prerelease must have a core inside the range")

    @property
    def text(self) -> str:
        stable = f">={self.minimum} <{self.maximum}"
        return stable + ("; prerelease=" + ",".join(self.prereleases) if self.prereleases else "")

    def reason(self, installed: Any) -> str | None:
        if installed is None or installed == "":
            return "unknown_version"
        try:
            value = Version.parse(installed)
        except ValueError:
            return "invalid_version"
        if value.core >= Version.parse(self.maximum).core:
            return "unsupported_boundary"
        if value.prerelease:
            return None if any(value == Version.parse(p) for p in self.prereleases) else "prerelease_not_admitted"
        return "below_minimum" if value < Version.parse(self.minimum) else None


ADDIN_REQUIREMENT = Requirement("5.2.0", "6.0.0", ("5.2.0-dev.16",))


def compatibility_result(
    installed: Any, requirement: Requirement = ADDIN_REQUIREMENT,
    *, component: str = "addin", detail: str | None = None,
) -> dict[str, Any]:
    reason = requirement.reason(installed)
    result = {
        "success": reason is None, "component": component,
        "installed_version": installed if isinstance(installed, str) else None,
        "required_range": requirement.text, "minimum_version": requirement.minimum,
        "requirement_status": REQUIREMENT_STATUS, "compatibility_reason": reason,
    }
    if reason is None:
        return result
    label = "MSAccess VCS add-in" if component == "addin" else "MCP server"
    if reason == "below_minimum":
        action = f"Update the {label} to a build satisfying {requirement.text}, then reconnect the MCP server."
    elif reason == "unsupported_boundary":
        action = (
            f"Use a {label} satisfying {requirement.text}, or a consumer release that explicitly "
            "supports the installed version. A newer compatible server/workflow may be needed; "
            f"upgrading this {label} again is not the remedy."
        )
    else:
        action = (
            f"Confirm the configured {label} installation and its strict SemVer identity; "
            f"use a build satisfying {requirement.text}, then reconnect."
        )
    result.update(
        error_pattern="version_incompatible" if reason in {"below_minimum", "unsupported_boundary", "prerelease_not_admitted"} else "version_unconfirmed",
        error=f"{label} installed version {result['installed_version'] or 'unknown'}; required {requirement.text} (minimum {requirement.minimum}). {action}",
        recovery_action=action,
    )
    if detail:
        result["discovery_error"] = detail
    return result


def server_metadata() -> dict[str, Any]:
    return {
        "mcp_version": __version__, "supported_addin_range": ADDIN_REQUIREMENT.text,
        "addin_requirement_status": REQUIREMENT_STATUS,
    }


def workflow_requirement() -> Requirement:
    """Separate consumer-owned workflow declaration, bundled for current clients."""
    declaration = json.loads(files("msaccess_vcs_mcp").joinpath("workflow_requirement.json").read_text(encoding="utf-8"))
    return Requirement(declaration["minimum"], declaration["maximum"], tuple(declaration["prereleases"]))


def workflow_preflight_instructions() -> str:
    requirement = workflow_requirement()
    return (
        "**Required compatibility preflight (workflow-owned):**\n"
        f"This distributed workflow requires MCP server {requirement.text} "
        f"(minimum {requirement.minimum}; assigned, unpublished release). "
        "Before database work and after reconnecting to a changed server, call the "
        "read-only vcs_get_version_info() and compare mcp_version with this workflow "
        "requirement using strict SemVer 2.0.0; build metadata has no precedence. "
        "Stable versions must lie in the range; only explicitly listed prereleases "
        "are admitted. Major-zero compatibility is bounded to one minor. "
        "Older servers need only their existing mcp_version field. If the tool is "
        "absent from the inventory or its version is missing/malformed, stop dependent "
        "calls and report compatibility unconfirmed; do not probe unsupported APIs. "
        "For an older server, notify the user of installed version, required range, "
        "minimum, and the action: update to a compatible build, restart the server "
        "process, and reconnect. For a newer unsupported major/minor, explain the "
        "supported combination: use a supporting workflow or reconnect to a supported "
        "server; upgrading the same server again is not the remedy. Check again if "
        "connection identity changes; if change detection is unavailable, recheck at "
        "each new database workflow. Compatible checks permit normal calls without "
        "feature probing. If addin_compatibility reports failure, or an operation "
        "returns version_incompatible/version_unconfirmed, stop dependent work and "
        "relay error, component, installed_version, required_range, minimum_version "
        "and recovery_action. A failed check authorizes notification only, not an "
        "installation, rebuild, upgrade or downgrade. Per-call policy/mode "
        "acknowledgments still confirm requested state.\n\n"
    )


def configured_addin_path() -> str:
    return get_config().get("ACCESS_VCS_ADDIN_PATH") or get_default_addin_path()


def _installation_identity(path: str) -> tuple:
    stat = os.stat(path)
    return (os.path.normcase(os.path.realpath(path)), stat.st_dev, stat.st_ino,
            stat.st_size, stat.st_mtime_ns, stat.st_ctime_ns)


def _read_installed_version(path: str) -> Any:
    """Only this short-lived DAO process is terminated on timeout, never Access."""
    output = subprocess.run(
        [sys.executable, "-m", "msaccess_vcs_mcp.version_probe", path],
        stdin=subprocess.DEVNULL, capture_output=True, text=True, encoding="utf-8", timeout=DISCOVERY_TIMEOUT_SEC,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0), check=True,
    )
    reply = json.loads(output.stdout)
    if not isinstance(reply, dict):
        raise ValueError("Version discovery did not return an object")
    if reply.get("error"):
        raise RuntimeError(reply["error"])
    return reply.get("version")


def inspect_installed_addin(path: str | None = None) -> dict[str, Any]:
    """No admission cache: re-read the library on every call, even at the same path.

    A file changed during discovery is refused. This inspection invokes no Access
    application or VBA APIs, so it cannot dispatch into a running operation.
    """
    installed = None
    path = path or configured_addin_path()
    try:
        before = _installation_identity(path)
        installed = _read_installed_version(path)
        if _installation_identity(path) != before:
            raise RuntimeError("Installed library changed during version discovery; retry")
        result = compatibility_result(installed)
    except (OSError, ValueError, RuntimeError, subprocess.SubprocessError) as exc:
        result = compatibility_result(None, detail=str(exc))
    result["addin_path"] = path
    return result
