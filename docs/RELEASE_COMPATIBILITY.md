# Release compatibility

The workflow owns its server requirement in the shipped
[`workflow_requirement.json`](../src/msaccess_vcs_mcp/workflow_requirement.json);
[AGENT_WORKFLOWS.md](AGENT_WORKFLOWS.md#required-compatibility-preflight) and the
MCP initialization instructions distribute its preflight.
The server owns its add-in requirement in `compatibility.ADDIN_REQUIREMENT`:
stable **>=5.2.0 <6.0.0**, with explicit development admission
**5.2.0-dev.16**. Josh assigned these ranges on 2026-10-02. Publication and
release verification remain human-owned; these assignments do not claim the
releases are available. The server source identifies as **0.2.0-dev.16**.
Setuptools derives its PEP 440 package identity (`0.2.0.dev16`) from that source;
the public API identity is the strict SemVer `mcp_version`.

## Contract and provenance

The add-in minimum covers structured mode/policy acknowledgments, named-output
`BuildAs(source, output)`, terminal journals, cancellation requests, and API
dispatcher refusals delivered on the dialog-hardening branch. Local tag v5.0.1
has a Sub `SetInteractionMode`, lacks `APICapabilities` and the policy contract;
it cannot supply this API. The existing development source's stable 5.2.0
number does not establish that these edits shipped. X16 changes the source
`AppVersion` to 5.2.0-dev.16 for explicit test admission after a rebuild. A
rebuild preserves that identity; it does not publish 5.2.0. At release cut,
review the complete required API and set the assigned stable version before
publishing. Never publish changed contents under a previously released version.
M52 and verification gate 1 disposition remain open independently of X16.

The workflow minimum covers the read-only metadata, shared pre-dispatch add-in
gate and structured compatibility refusals implemented here. Earlier servers
remain diagnosable through their existing version tool, but are not admitted
for database work by this workflow. X13 remains Josh's human guidance review;
X14 packaging is a future consumer.

## Range semantics

Use [SemVer 2.0.0](https://semver.org/spec/v2.0.0.html), not lexical sorting or
PEP 440 comparison of API versions. Project syntax is `>=MIN <MAX`, optionally
followed by `; prerelease=V1,V2`. Both bounds are strict three-part stable
SemVer. The lower bound is inclusive; the upper bound is exclusive. For major
1+, MAX is the next major at `.0.0`. For major zero it is the next minor at
`.0`. Support for multiple majors requires a deliberate contract revision;
there is no wildcard/OR syntax. Prereleases are rejected except exact enumerated
identities whose core lies inside the stable bounds. An explicitly admitted
prerelease of the minimum core is allowed despite preceding the stable minimum.
Build metadata is ignored for precedence and explicit prerelease identity.
Numeric prerelease fields compare numerically, numbers sort before text,
text compares in ASCII order, and stable follows prerelease of the same core.
Missing/invalid versions are refused. There is no input normalization: leading
`v`, whitespace, two/four-part versions and leading numeric zeroes are invalid.

## Enforcement and diagnostics

The `vcs_tool` wrapper admits every add-in-dependent MCP call before the tool
body. CLI commands reach the same wrappers through stdio. Gated calls inspect
on the Access apartment worker; the exempt rebuild checks before its body can
register callbacks or acquire its own launch gate. Refusal precedes target
opening/mutation, policy setup, callbacks, and operation dispatch. Existing
installed-target, path, write-permission and per-call acknowledgment checks
still apply. Named-output builds rely on this versioned contract; their former
mandatory capability probe is removed from the public path.

Installed-library discovery reads DAO `AppVersion`, the same property used by
`GetVCSVersion`. A hidden subprocess using the server interpreter opens only
that library with `DAO.DBEngine.120`, shared and read-only. It does not start
Access, execute AutoExec or VBA, attach to a target, or change interaction
state. Its 10-second deadline terminates only the discovery Python process.
Its stdin is disconnected from the MCP protocol pipe.
Access processes are never terminated by discovery. DAO.DBEngine.120 must be
available to the server interpreter; a missing/mismatched ACE DAO runtime is
a discovery failure, not evidence of an add-in version mismatch. There is no compatibility
cache: each dependent call reads again and compares file identity before and
after discovery. Changes during discovery are refused. Installer registry
settings and configured paths are re-read, including compiled `.accde` changes.
As with existing path preconditions, replacement after inspection is a possible
external race; installation must not run concurrently with an operation.

`vcs_get_version_info` returns `mcp_version` and `supported_addin_range` even
without Access, a library, or an available worker. `addin_requirement_status`
states `assigned_unpublished`. When obtainable, `vcs_version` and
`addin_compatibility` describe the installed version and its admission verdict.
Existing `access_version` and `bitness` keys remain, with null values because
this metadata check does not start/attach to Access. Paths, logs and callback
fields remain. A successful metadata retrieval does not mean the add-in was
accepted: check `addin_compatibility.success`.

These exceptions need no compatible add-in API: version metadata, recent-call
inspection, `vcs_list_dialogs`, `vcs_dismiss_dialog`, `vcs_recover_dialogs`,
`vcs_automation_status`, and `vcs_cancel_operation` (records an HTTP cancel
request). `vcs_list_objects` and `vcs_diff_database` use Access/DAO rather than
the add-in and retain their existing guards. `vcs_get_log`, option reads,
compile checks and `vcs_end_session` call add-in APIs and are gated. Rebuild is
gated too; it is not an automatic compatibility recovery exception. If an
incompatible install prevents rebuild, use the documented manual **Build From
Source**/installer procedure after user authorization.

## Stable error contract and recovery

Refusals return `success:false`, `error_pattern`, `component:"addin"`,
`installed_version` (string or null), `required_range`, `minimum_version`,
`requirement_status`, `compatibility_reason`, `error`, `recovery_action` and
`addin_path`. A discovery failure also carries `discovery_error`.

| Pattern | Reason | User action |
|---|---|---|
| `version_incompatible` | `below_minimum` | Update the add-in to a compatible assigned release or explicitly admitted development build; reconnect. |
| `version_incompatible` | `unsupported_boundary` | Use a supported add-in/server combination, or a newer server explicitly supporting the installed major. Updating that add-in again is not the remedy. |
| `version_incompatible` | `prerelease_not_admitted` | Use a supported stable build or obtain an explicit consumer requirement admitting this prerelease. |
| `version_unconfirmed` | `unknown_version`, `invalid_version` | Verify path/install/version, resolve discovery failure, reconnect, and repeat metadata. |

For server updates, use the installation method in [README.md](../README.md),
selecting a version that satisfies the workflow rather than assuming latest is
supported. A development install uses the project venv and source checkout;
restart the process to load its new code. For add-in development builds, edit
source `dbs-properties.json` `AppVersion`, then use the supported
[rebuild workflow](../../msaccess-vcs-addin/docs/agentic-rebuild.md) and
[development test-host workflow](../../msaccess-vcs-addin/docs/agent-test-runs.md).
Never patch the installed library in place. The legacy `Deploy` version
incrementer accepts numeric stable versions, not prerelease identities: use
source rebuild for development; release owners set the stable identity before
Deploy/publication. Neither failed gate grants permission to update components.
