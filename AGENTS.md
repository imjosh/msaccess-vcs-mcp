# MSAccess VCS MCP contributor guide

This repository bridges MCP tools to the Access VCS add-in. This historical-derived
UP-004 consumer uses request-owned policies and requires qualification with the
matching repaired add-in. It does not claim the production compatibility-session
contract or published release support.

| Task | Read |
| --- | --- |
| User database edits, imports, exports, and tests | [docs/AGENT_WORKFLOWS.md](docs/AGENT_WORKFLOWS.md) |
| Request policies, blocked conflicts, skipped transfers, or deletion review | [docs/DIALOGS.md](docs/DIALOGS.md) |
| Python setup, tests, COM ownership, or logging | [docs/CONTRIBUTING.md](docs/CONTRIBUTING.md) |
| Setup acknowledgments and dispatch | [docs/SETUP_CALLS.md](docs/SETUP_CALLS.md) |

Use the project virtual environment. For isolated checkouts, use the primary venv
with PYTHONPATH pointed at the selected checkout's src. Keep the shared editable
installation unchanged. Run the selected tests with python -m pytest.

Register tools with @vcs_tool in tools.py and preserve configuration reload,
logging, installed-target refusal, and the Access gate. Keep COM objects on their
creating apartment thread. Dialog, status, and cancel inspection must remain
responsive independently of blocked COM work. Leave Access visible after opening
the database through open_current_database.

Close or recycle only a server-owned process confirmed by PID and creation time.
Unknown identity is not ownership. Preserve user-owned windows and unsaved work.
The installed add-in is a library, never a tool target. Edit add-in source and use
the supported source rebuild. Preserve admission versus terminal completion and
cancel request versus confirmed cancellation. Keep extensible journals, transfer
outcomes, original errors, and cleanup diagnostics in structured results.

For the add-in's own suite, use a fresh disposable development host with matching
source context and its maintained lifecycle. User-project tests run in the user's
database. An all-EMPTY run does not prove a passing suite.
