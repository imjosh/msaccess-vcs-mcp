# MCP contributor guidance

Use the project virtual environment for Python and pytest. For a standalone
checkout, create and activate `venv`, then install development dependencies there.

```powershell
python -m venv venv
.\venv\Scripts\Activate.ps1
python -m pip install -e ".[dev]"
python -m pytest -m "not integration"
```

For isolated checkouts that share an existing environment, use that environment's
Python executable and set `PYTHONPATH` to the selected checkout's `src`. Keep the
shared editable installation unchanged. Record the selected source path and Git
identity with check results. Windows asyncio requires local socket access even
when tests mock every Access boundary.

Register tools with `@vcs_tool` in `tools.py`. Preserve configuration reload,
usage logging, installed-target refusal, and the Access gate. Keep COM objects
on their creating apartment thread. Gate-exempt recovery must remain responsive
when another call is blocked. Open databases with `open_current_database` and
show Access afterward. Confirm server process ownership with PID and creation
time before closing it. Unknown identity is not ownership.

The installed add-in is a library and cannot be a tool target. Add-in changes
use exported source and the supported rebuild workflow. Keep native tests
separate from mocked unit tests and serialize changes to the installed library.

Use [DIALOGS.md](DIALOGS.md) for request policies, conflict protection, transfer
results, and refusal recovery. Use [VBA_CALLBACK_API.md](VBA_CALLBACK_API.md) for
argument forwarding and callbacks. Setup acknowledgments remain documented in
[SETUP_CALLS.md](SETUP_CALLS.md). User-project workflows are maintained in
[AGENT_WORKFLOWS.md](AGENT_WORKFLOWS.md).
