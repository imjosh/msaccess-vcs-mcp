# MCP release and compatibility

These skills belong to the MCP release that supplies them. Use that release's
server and its supported add-in combination. Preserve this entire guidance tree;
individual skill directories and mixed-version installation are unsupported.
Client integration packages are the preferred future delivery route (separate X14).

## Preflight before database work

1. In a release artifact, read the generated `release.json` beside this policy.
   It identifies `package_version`, the corresponding SemVer `mcp_version`, and
   hashes of both skills/shared references. Builds derive it from package metadata
   and the same source version used by the server, without a separate release constant.
   Source checkouts use `src/msaccess_vcs_mcp/__init__.py` and the authoritative
   `workflow_requirement.json` beside it; a copied directory without provenance
   is unsupported. Packaged guidance includes a byte-identical copy of that JSON.
2. Before database operations, call the read-only `vcs_get_version_info()`.
   Read `mcp_version` even when installation diagnostics make `success` false.
   Compare the exact release identity including its build suffix with provenance.
   Also consume the workflow declaration: strict SemVer stable versions must fall
   between its inclusive minimum and exclusive maximum; prereleases must match an
   explicitly listed identity, ignoring build metadata for that range check.
   The exact release comparison is a distribution check, separate from the existing
   workflow requirement. It does not replace server/add-in admission.
3. Repeat after reconnection, server replacement or a changed connection identity.
   When change detection is unavailable, repeat at each new database workflow.
   Missing provenance, absent version tool, missing/malformed version, a failed
   requirement or a different release stops dependent work. Report the component,
   observed identity, required release/range/minimum and recovery action. Older
   servers need only valid `mcp_version`; establish an absent tool from inventory.
   Use guidance from the responding server's release or reconnect to the intended
   release. A too-new provider needs a supported combination; another provider
   upgrade does not necessarily fix it. This check authorizes no installation,
   rebuild, upgrade or downgrade.
4. Relay a failed `addin_compatibility` or operational `version_incompatible`,
   `version_unconfirmed` or `compatibility_session_invalid`, including `error`,
   `component`, `installed_version`, `required_range`, `minimum_version` and
   `recovery_action`. The server and add-in reuse their automatic mutual session
   admission and instance-bound validation; metadata does not establish a session
   or prove VBA connectivity. Proceed without capability probes solely for release
   compatibility. Per-call policy/mode acknowledgments remain required.

## Inspection and recovery after refusal

Retain server-permitted metadata/recent calls, Win32 dialog/status/recovery,
HTTP cancellation and raw DAO inventory/diff under their existing guards.
Add-in options/logs/compile, end-session and rebuild require admission. Do not
bypass a refusal through VBA or replay an uncertain dispatched call. Retry database
changes only after compatibility and the previous outcome are known. Recovery
does not authorize closing user-owned windows or expanding operation permissions.
