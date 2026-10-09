# MCP server contributor guidance

For user-database operations, use [AGENT_WORKFLOWS.md](AGENT_WORKFLOWS.md).
For add-in rebuilds and its own test host, read [ADDIN_DEVELOPMENT.md](ADDIN_DEVELOPMENT.md).
For policy/callback and Win32 dialog contracts, read [DIALOGS.md](DIALOGS.md).

## Development Environment Setup

**Always use the project virtual environment.** The package is installed in editable mode inside `venv/`. Do not install globally or suggest `pip install` outside the venv.

```powershell
# Activate the virtual environment (Windows PowerShell)
cd C:\path\to\msaccess-vcs-mcp
.\venv\Scripts\Activate.ps1

# If the venv doesn't exist yet, create it first:
python -m venv venv
.\venv\Scripts\Activate.ps1
pip install -e ".[dev]"
```

### Running Tests

Tests must run inside the activated virtual environment:

```powershell
.\venv\Scripts\Activate.ps1

# Run all unit tests
pytest

# Run a specific test file
pytest tests/test_usage_logging.py -v

# Skip integration tests (require Access installed)
pytest -m "not integration"

# Run with coverage
pytest --cov=msaccess_vcs_mcp --cov-report=html
```

## Architecture

Register tools with `@vcs_tool("name")` in `tools.py`. Its async wrapper discovers
workspace environment files lazily, refreshes configuration, and refuses installed
add-in targets before entering the Access gate. Gated bodies execute on the COM
apartment; exempt tools use independent bounded workers. Interruption attribution
is applied before usage logging so the client and audit stream see the same
outcome. `mcp.tool()` exposes the wrapped function to MCP clients.

```
AI Agent ──► MCP Server (Python) ──► VCS Add-in (VBA) ──► Access Database
                 │
           Path validation
           Permission checks
           Usage logging
           Async progress tracking
```

### Access window visibility

Any Access instance holding a database the server works with is left **visible**, whether the server created it or attached to one already running. A hidden instance strands the user: an error dialog, a VBA break, or a trust prompt blocks every later call with nothing on screen to explain why, and nobody can dismiss what they cannot see. `COM automation` normally starts Access hidden, so this is deliberate, not incidental.

These rules follow from that, both enforced in `access_com/connection.py`:

- Open databases through `open_current_database(app, path)`, never a bare `app.OpenCurrentDatabase(...)`. It lowers `Application.UserControl` across the open — `OpenCurrentDatabase` runs the target's AutoExec, and the add-in's own `AutoRun` opens its installer form when that flag says a person is watching, stranding the instance the server is about to drive — and it shows the window afterwards.
- Show the window *after* the database opens, via `ensure_access_visible(app)`. Making a window visible can itself set `UserControl`, which is why the order is not interchangeable.
- After that, instances the server created also get `UserControl = True` so the window is a normal interactive Access app. `Visible` alone is not enough on a COM-created process.

The one deliberate exception is `validate_access_installation()` in `config.py`: it opens no database and quits immediately, so a window would only flash on screen with nothing to act on.

### Full-power cores

Access is single-threaded. MCP-launched Access processes disable EcoQoS and run at Above Normal so Windows prefers a performance core. User-owned Access the server attaches to is left alone. Do not pin CPU affinity.

## Configuration

All settings come from environment variables (loaded from `.env` / `.env.local` in the project root). See [`.env.example`](../.env.example) for the full list.

Read [`.env.example`](../.env.example) for settings and `config.py` for defaults;
keep configuration facts there rather than caching another variable catalog.

### Which Access windows the server may close

Persistence only pays off if the server can still get a file back when it needs to replace one, so it has to know which windows are its own. That is recorded on disk in the ownership registry, keyed by PID *and* process creation time — a PID alone is reused by Windows, and a reused PID would authorize closing a stranger's process.

The rules, all enforced in `access_com/instance_registry.py` and `access_com/connection.py`:

- **Only a confirmed match is owned.** If the creation stamp cannot be read, or the record does not carry one, the answer is "not ours". Every ambiguity resolves toward leaving the window alone.
- **A failed process query is not an empty one.** `list_access_pids_or_none()` returns `None` when `tasklist` itself fails, and the registry keeps its records rather than forgetting every window it created.
- **Closing re-verifies identity.** Both ways of resolving an instance from a path can land on a different process than the one recorded, and a moniker bind can even start a new one. `pid` and creation time are re-checked against the record before anything is closed.
- **A hung owned instance is terminated.** The graceful `CloseCurrentDatabase` + `Quit` runs in a worker thread with `ACCESS_VCS_CLOSE_TIMEOUT_SEC`; a blocking call here would hang the Access gate. If the process survives, it is force-terminated. This is reachable only for a PID proven to be server-created, where unsaved state is acceptable loss.
- **A live process keeps its claim.** A record is dropped only once the process is confirmed gone. Forgetting a live owned instance would make it permanently uncloseable.
- **A loaded add-in counts as holding a file.** Every tool call loads the add-in as a library, which locks it no matter which database that instance has open. `load_addin` records this, and a rebuild closes those instances too — matching only on the open database would make rebuilds refuse almost every time.

One MCP server process is shared across all Cursor windows. Gated tools run in a single COM apartment thread with one Access operation at a time; an async tool body runs there in an event loop of its own, so a blocking COM step never stalls the server loop. A long call in one window causes others to get `error_pattern: server_busy` with `busy_with` naming the in-flight tool. User-workflow recovery is maintained in [access-vcs-recover](../skills/access-vcs-recover/SKILL.md).

`vcs_run_vba` executes Access COM work in a daemon worker thread with a hard timeout. If a snippet hangs because Access is in break mode, blocked on a modal dialog, or otherwise unresponsive, the MCP server abandons that thread and returns a recoverable timeout. It does **not** kill `MSACCESS.EXE` or close user-owned Access windows; after Access becomes responsive, the next call runs an automatic probe and resumes normal operation.

### The installed add-in is never a target

No tool accepts the installed add-in as `database_path`, `output_path`, or `template_path`. That file exists to be loaded as a library: opening it as a database, or writing into it, resets a VBA project while it is executing. `_refuse_installed_addin_target` runs inside the `vcs_tool` wrapper, ahead of the gate and of any COM work, so the rule holds for every tool rather than the handful that grew their own guards — export, import, rebuild, `run_vba`, `run_tests`, `call_vba`, and the rest. Refusals carry `error_pattern: installed_addin_refused`.

`vcs_get_version_info()` reports the installed add-in's version without opening it, which is the one thing you might legitimately want from that file.

The comparison ignores the file extension, mirroring the add-in's `modInstall.PathsMatchIgnoringExtension`: a compiled install is a `.accde` built from the same `.accda`, and only one of the two is ever named in `ACCESS_VCS_ADDIN_PATH`. Folder parameters (`source_dir`, `output_dir`) are not checked — an export folder beside the install is not the install.

**Resolving the install path.** With `ACCESS_VCS_ADDIN_PATH` unset, `get_default_addin_path()` reads `HKCU\Software\VB and VBA Program Settings\MSAccessVCS\Install`, which is the only place to read it from — the add-in's own `GetInstalledAddInFileName` is built from exactly these two values. `Install Folder` is present only for a folder the user chose (the installer deletes it when the folder is the default, so absence means `%AppData%\MSAccessVCS`, not "not installed"), and `Compile accde` decides the extension. Do not reconstruct the path from `%AppData%` alone or assume `.accda`. Registry settings are reread on every lookup, so installer changes take effect
without restarting. `reset_addin_path_cache()` remains a compatibility no-op;
`ACCESS_VCS_ADDIN_PATH` still overrides registry discovery.

## Logging

The server writes two parallel JSON Lines streams. Both filenames use the `vcs-mcp-` prefix so they don't collide with other tools that share the same logs directory.

### VCS operation logs (written by the add-in, not the server)

Separate from the two streams below, the add-in writes a per-operation log to `{source_dir}/logs/<Base>_<yyyymmdd_hhnnss_fff>.log`, where `<Base>` is `Export`, `Merge`, `Build`, `TestRun`, or `Other`. Note the base name tracks the *operation*, not the tool: `vcs_import_objects` and `vcs_import_object` both produce `Merge_*.log`.

Agent log-retrieval procedures live in [access-vcs-recover](../skills/access-vcs-recover/SKILL.md).

The add-in's sync API returns this as camelCase `logPath`; `_addin_json_result` in `tools.py` normalizes it to `log_path` at the boundary and keeps the original key as an alias. Async completion callbacks already use `log_path`. Keep the normalizer even if the add-in changes: a newer server may run against an older add-in build.

### Diagnostic stream (`vcs-mcp-diagnostic-<instance>.jsonl`) — always on

Captures server lifecycle events: `server_start`, `startup_env_load`, `lazy_env_load`, `lazy_init_started`, `lazy_init_skipped`, `list_roots_failed`, `list_roots_response`, `lazy_init_loaded`, `lazy_init_no_env_in_roots`, `usage_log_status`. Independent of `ACCESS_VCS_ENABLE_LOGGING` so it answers the "why didn't logging work?" question even when usage logging is silent.

- **Location:** `~/.msaccess-vcs-mcp/logs/vcs-mcp-diagnostic-<instance>.jsonl`
- **Override:** `ACCESS_VCS_DIAGNOSTIC_LOG_DIR`
- **Opt out:** `ACCESS_VCS_DISABLE_DIAGNOSTIC_LOG=true`
- **Rotation:** 1 MB per file, 3 backups
- **Discoverable from agents:** `vcs_get_version_info()` returns the active `diagnostic_log_path`.

### Usage stream (`vcs-mcp-usage-<instance>.jsonl`) — default on

When `ACCESS_VCS_ENABLE_LOGGING=true` (the default), every tool call writes a structured entry. Set the env var to `false` to opt out.

- **Development installs:** logs to `{project_root}/logs/vcs-mcp-usage-<instance>.jsonl`
- **Package installs:** logs to `~/.msaccess-vcs-mcp/logs/vcs-mcp-usage-<instance>.jsonl`
- **Override:** `ACCESS_VCS_LOG_DIR`
- **Rotation:** `ACCESS_VCS_LOG_MAX_SIZE_MB` (default 10 MB), `ACCESS_VCS_LOG_BACKUP_COUNT` (default 5)

Each `tool_call` entry includes: `timestamp`, `version`, `event`, `tool`, `parameters`, `success`, `error`, `error_pattern`, `execution_time_ms`.

`<instance>` identifies the server process generation, separating concurrent
writers. Use `vcs_get_version_info()` for active diagnostic/usage paths and
the responding server's metadata rather than guessing filenames. Locked/unwritable streams
recover automatically on a later write and configuration reload; fallback to
structured stderr has rate-limited warnings and no growing memory queue. Logging does not block tool work with retry loops.

### Tiered audit posture

Three sensitivity tiers in the usage stream, each independently controlled:

1. **Audit metadata** — always written when `ENABLE_LOGGING=true`. Tool name, timing, success/error, sanitized parameter dict.
2. **Code-execution bodies** — `vcs_execute_sql`, `vcs_call_vba`, and `vcs_run_vba` write a `"code_execution"` event *before* execution begins. By default only `code_length` is recorded. Set `ACCESS_VCS_LOG_CODE_CONTENT=true` to record the full `code` field for forensic replay (off by default to limit business-data exposure). `code_length` lets analysts spot anomalies (e.g. "an agent ran a 4 KB VBA block") without seeing the body.
3. **Credential-shaped parameter keys** — any parameter whose name matches `password`, `secret`, `token`, `api_key`, `apikey`, `connection_string`, or `connectionstring` (case-insensitive) is replaced with `"<redacted>"` regardless of other switches. Defense in depth.

## Adding a New Tool

1. Write the handler function in `tools.py`
2. Decorate with `@vcs_tool("vcs_your_tool_name")` — this handles config reload, usage logging, and MCP registration automatically
3. Add tests in `tests/`

## Key Conventions

- All tool names use the `vcs_` prefix
- Tool handlers return `dict[str, Any]` with at least a `success` key
- Error results include `"error"` key (detected by usage logging)
- Async tools (`async def`) are supported by `@vcs_tool` transparently
- `Context` parameters from FastMCP are filtered out of usage logs automatically
