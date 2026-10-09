# AGENTS.md — msaccess-vcs-mcp

This repository is the Python MCP bridge to the MSAccess VCS add-in. Its tools
operate on users' Access databases; contributor work here changes the bridge.

## Read for the task

| Working on | Read |
| --- | --- |
| Editing/syncing a user's project, fresh build, verification, or project tests | [docs/AGENT_WORKFLOWS.md](docs/AGENT_WORKFLOWS.md); optional [access-vcs-sync](skills/access-vcs-sync/SKILL.md) |
| Busy server, timeouts, dialogs, cancellation, interruption, or cleanup | [docs/DIALOGS.md](docs/DIALOGS.md); optional [access-vcs-recover](skills/access-vcs-recover/SKILL.md) |
| Installing/distributing these workflow skills | [docs/AGENT_WORKFLOWS.md](docs/AGENT_WORKFLOWS.md#optional-skill-installation-and-distribution) |
| Compatibility preflight, refused admission, or changed connection | [docs/AGENT_WORKFLOWS.md](docs/AGENT_WORKFLOWS.md#required-compatibility-preflight), [docs/RELEASE_COMPATIBILITY.md](docs/RELEASE_COMPATIBILITY.md) |
| Python setup/tests, architecture, COM ownership/threading, registration, logging | [docs/CONTRIBUTING.md](docs/CONTRIBUTING.md) |
| Rebuilding the add-in or running its own tests | [docs/ADDIN_DEVELOPMENT.md](docs/ADDIN_DEVELOPMENT.md) |
| Policy, callback, interruption, or Win32 dialog implementation | [docs/DIALOGS.md](docs/DIALOGS.md), [docs/SETUP_CALLS.md](docs/SETUP_CALLS.md) |
| VBA bridge/callback implementation | [docs/VBA_INTEGRATION.md](docs/VBA_INTEGRATION.md), [docs/VBA_CALLBACK_API.md](docs/VBA_CALLBACK_API.md) |
| Server integration-test procedures | [docs/TESTING.md](docs/TESTING.md) |

## Essential contributor invariants

- Use the project `venv` for Python and pytest; install editable development
  dependencies there. Setup and test commands live in the contributor guide.
- Register tools with `@vcs_tool` in `tools.py`; preserve config reload, logging,
  installed-target refusal, and the Access gate. Return structured outcomes.
- Keep gated COM objects on their creating apartment thread. Dialog/status/cancel
  inspection must remain responsive independently of a blocked COM call.
- Leave Access visible. Use `open_current_database` and show it after opening;
  setting visibility earlier can change `UserControl` and trigger AutoExec UI.
- Close/recycle only a server-owned process confirmed by PID **and creation time**.
  Unknown identity is not ownership. Preserve user-owned windows and unsaved work.
- The installed add-in is a library, never a tool target. Do not patch a loaded
  add-in project; source changes use the add-in rebuild workflow.
- Preserve the distinction between start acknowledgment and terminal outcome,
  between cancel request and confirmed cancellation, and between operation result
  and cleanup failure. These are server-enforced contracts, not retry heuristics.
- For user-project tests, host the run on the user's database. The development
  add-in is a host only for the add-in's own suite; use its maintained disposable-host lifecycle.

## Active cross-repo work

`feat/noninteractive-dialogs` uses the parent-folder tracker shared with
`msaccess-vcs-addin`: `../issues/INDEX.md` and
`../specs/dialog-handling-hardening.md`. For work in that tracker, pick a `todo`
issue with completed blockers and update both its file and index.
