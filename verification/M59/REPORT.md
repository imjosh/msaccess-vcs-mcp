# M59 project-selection qualification

Checked 2026-10-14 against MCP baseline `f2fed88`. Changes cover README setup,
`docs/PROJECT_SELECTION.md`, the existing workflow preflight, and focused
configuration checks. No server implementation changed.

## Controlled automated evidence

Run from the MCP checkout with its existing virtual environment:

```powershell
.\venv\Scripts\python.exe -m pytest tests/test_config_env_loading.py tests/test_lazy_workspace_init.py -q --basetemp .pytest-m59-qualified --junitxml=verification/M59/results.xml
```

Result: **69 passed in 2.20 seconds**. [JUnit results](results.xml) include the
11 new `TestGlobalProjectRecipes` cases and existing environment-loading and
M58 SDK discovery checks. Two sandboxed attempts stalled and were stopped.
The completed run used approved execution outside the sandbox.

New checks call the registered public `vcs_get_version_info` wrapper with
disposable project files and a mocked installed add-in. They assert:

- Explicit selection outside the project, including a path with spaces, reports
  the chosen root and resolution method without requesting client roots.
- Initial base-file, launch-value, empty-launch-value, and personal-override
  precedence matches the documentation. Diagnostics identify the loaded files.
- A base-file reload replaces launch values, personal overrides remain last,
  and removing a key retains its previous process value.
- Newly created files and file deletion do not themselves trigger reload.
- The exact documented Claude Code Python argument maps either of two supplied
  project contexts and exits for absent or invalid context. Server startup is
  intercepted in this check.

The existing checks exercise native Windows paths through the real MCP SDK,
valid file URIs, failed discovery recovery, and selection stability. These are
controlled responses, not actual client runs. No Access database was created,
opened, attached to, or changed. No client configuration or login was modified.

## Client evidence and sources

Read-only version commands reported:

| Client | Reported version | Qualification |
| --- | --- | --- |
| Cursor | 3.24.12, commit `37b865941ccb31e2fc90c685157534efca87f960`, x64 | CLI printed an EPERM chdir error before version output; runtime loading unchecked |
| Claude Code | 2.1.276 | `--version` and `mcp add --help` only; no server registration or launch |
| Claude Desktop | Not checked | No runtime or version evidence |

Official documentation consulted on 2026-10-14:

- [Cursor MCP interpolation](https://cursor.com/docs/mcp#config-interpolation)
  documents `${workspaceFolder}` in `env` and global configuration placement.
- [Claude Code stdio project context](https://code.claude.com/docs/en/mcp#option-3-add-a-local-stdio-server)
  documents the spawned process's `CLAUDE_PROJECT_DIR`.
- [Claude Code scopes](https://code.claude.com/docs/en/mcp#mcp-installation-scopes)
  documents user registration and the default local scope.
- [MCP local-server setup](https://modelcontextprotocol.io/docs/2026-07-28/develop/connect-local-servers)
  documents Desktop's separate configuration and logs.

The 2026-10-09 ticket records an actual Cursor connection with 25 tools and an
empty configured database. It also records a native Windows roots validation
failure and written client recipes whose runtime loading checks were skipped.
Those historical observations do not qualify the recipes now documented.

## Remaining completion gaps

M59 remains in progress until actual client checks are recorded:

1. In Cursor, launch the global entry from two disposable project folders whose
   environment files name distinct absolute database paths. Call metadata from
   each connection and record root, method, database, package source, and loaded
   file diagnostics. Check multi-folder workspaces separately if used.
2. In Claude Code, qualify the user-scoped command from two disposable project
   sessions. Confirm the installed client supplies `CLAUDE_PROJECT_DIR` and that
   command arguments survive PowerShell and CLI registration unchanged.
3. In Desktop, record its version, launch the two named entries, and call metadata
   through each entry to confirm independent project configuration.

Use metadata only. The database filenames can refer to nonexistent disposable
paths; configuration selection does not require opening them. Preserve unrelated
registrations, credentials, and application databases. A client connection or
tool inventory alone does not close any of these checks. No actual Cursor
qualification of M58's adapter is claimed.
