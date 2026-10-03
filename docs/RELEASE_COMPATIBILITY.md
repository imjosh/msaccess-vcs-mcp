# Release compatibility

X17 uses automatic mutual admission through named VBA procedures. Each side
checks the other's version/protocol once. Later commands carry an accepted
session ID and bound caller identities, without sending versions again. COM
is the existing procedure-call transport; compatibility does not drive the UI.

## Consumer declarations and identities

| Consumer | Authoritative declaration | Supported provider |
|---|---|---|
| Server | compatibility.ADDIN_REQUIREMENT | Add-in >=6.0.0 <7.0.0; explicit 6.0.0-dev.17 |
| Add-in | modCompatibilitySession.SERVER_* | Server >=0.3.0 <0.4.0; explicit 0.3.0-dev.17 |
| Workflow | [workflow_requirement.json](../src/msaccess_vcs_mcp/workflow_requirement.json) | Server >=0.3.0 <0.4.0; explicit 0.3.0-dev.17 |

The add-in requirement comes from its command contract, independently of the
workflow requirement. Protocol: **msaccess-vcs.session/1**. Development pair:
add-in **6.0.0-dev.17**, server **0.3.0-dev.17**; package metadata uses PEP 440
**0.3.0.dev17**. These X17 implementation assignments are **unpublished**.
Legacy refusal breaks the add-in API (major 6) and the major-zero server API
(minor 0.3). Historical X16 receipts qualify only the previous contract.
Release owners must verify and separately authorize publication of stable identities.

## Version policy

Use [SemVer 2.0.0](https://semver.org/spec/v2.0.0.html). Project range syntax is
`>=MIN <MAX; prerelease=V1,V2`, inclusive minimum/exclusive maximum. Stable
majors end at the next major; major-zero ranges end at the next minor. There
is no wildcard/OR syntax. Only enumerated prereleases are admitted, including
an enumerated minimum-core prerelease. Numeric fields compare numerically;
numeric prerelease fields sort before text, text uses ASCII order, and stable
follows prerelease. Ignore build metadata. Refuse missing/invalid versions,
leading v/whitespace, two-part versions and leading-zero numeric fields.

## Automatic handshake and execution

Initial discovery reads installed AppVersion with a shared, read-only DAO
subprocess in the server interpreter. It starts no Access, executes no VBA,
sets no policy, acquires no root and registers no callback. Its ten-second
limit terminates only the probe; stdin is DEVNULL. Successful discovery is
cached by path/file/requirement identity; registry paths are reread each call.

On the existing Access connection, **APIIdentity** returns the loaded version,
path, random incarnation, installation generation and protocol. The server
checks that loaded version separately from installed metadata. **APIHandshake**
receives the server version/protocol and expected incarnations; the add-in
checks its own requirement. Only mutual acceptance enters the server cache.
A missing named procedure fails via COM, never as an unknown legacy API method
that could open a modal error. No agent preflight activates this enforcement.

**APIExecute/APIExecuteAsync** validate the cached session before source/mode/
policy changes, dispatch or callback registration. Timers capture the envelope
and validate again before starting. Quick fallbacks use the same dispatcher.
Root continuations that already began retain their existing lease and callback.
Raw target calls relay through the admitted **CallVBA** command; RunVBA uses the
same boundary. The VCS.API spelling in vcs_call_vba is a server alias for
APIExecute. Rebuild callback registration follows admission; its disconnected
worker establishes a fresh session before the build.

**APIValidateSession** is the server's cheap pre-dispatch check. A failed check
may re-handshake before sending a command. No dispatched call, COM exception or
uncertain result is replayed for compatibility. Per-call policy/mode acknowledgments,
path/permission guards, results, cancellation and log attribution remain required.
Sessions do not replace operation IDs, callbacks, override-file IDs or policy tokens.
Foreign callers cannot clear another caller's session policy.

## Cache and invalidation

Server and add-in incarnations are random UUIDs; VBA reset/reload loses the
add-in session dictionary. Connections get separate UUIDs with retained object
references. The server cache includes Access PID **and creation time**, file,
requirement and protocol identities. Missing/failed identity checks refuse reuse.
The add-in checks volume/file ID/creation stamp; mutable translation-table writes
do not replace loaded code. The server additionally checks file timestamps/size.
Changed instance, connection, configuration, installation, requirement or protocol
invalidates approval. An old loaded library does not adopt a replacement's version.
Closing/reopening a target can leave the same library loaded; that alone is not
an add-in reload or MCP connection replacement.

Caches are bounded to 128 entries. Add-in sessions expire after 30 idle minutes;
a policy-owner entry is retained until its owner clears the policy or disconnects.
Expiry never clears another caller's policy. Automatic idle-policy disposal remains
an explicit lifecycle qualification gap. APIDisconnectSession refuses a policy owner's active
disconnect; otherwise it removes that session and cleans its own idle policy.
Invalidation never cancels an already executing root or borrows its result/log.
An invalidated queued start posts one refusal to its own callback.

## Refusal, diagnostics and migration

Patterns: version_incompatible, version_unconfirmed, compatibility_session_invalid.
Failures include success:false, component, observed version-or-unknown, required
range/minimum, reason, requirement status, error and recovery action. Session
refusals identify started:false. Too-new versions advise a supported combination,
not another provider upgrade. A failure never authorizes install/rebuild/upgrade.

Read-only vcs_get_version_info exposes server version/protocol/incarnation and
installed metadata without operational admission; metadata_establishes_session:false
makes that distinction explicit. Recent calls, Win32 dialogs/status/recovery,
HTTP cancellation and raw DAO inventory/diff retain their exceptions and guards.
Add-in options/logs/compile, end-session and rebuild require admission.

An old library lacks the new API and is refused. An explicitly authorized
migration uses the previous admitted X16 pair's [source rebuild](../../msaccess-vcs-addin/docs/agentic-rebuild.md)
once, or manual Build From Source/installer; then restart/reconnect the new
server. Never disable the gate or patch a loaded add-in. Source rebuild keeps
the prerelease; the release owner sets a numeric stable identity before legacy
Deploy. Merge add-in first and coordinate server shipping; publication is separate.

Evidence and unfinished criteria: [X17 report](../../verification/X17/README.md).
Live scope is Windows 64-bit Access 16/DAO.DBEngine.120 and Python 3.12 in the
project venv. Manual/ribbon/direct-VCS routes remain unchanged following the
user's simplification direction. They cannot universally distinguish human
VBA from arbitrary external VBA. X17's broader legacy-bypass/bootstrap coverage
remains explicitly unfinished; the named-API evidence is not a VBA sandbox or
whole-ticket completion. M52, full VERIFY-1/VERIFY-2 acceptance, A36, X13, X14,
other environments/desktop packaging and publication remain separate assignments.
