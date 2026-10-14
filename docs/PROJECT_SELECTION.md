# Select a project for a globally registered MCP server

Register the server once, then select the project directory that supplies its
database configuration. A connected server or a populated tool list confirms
the transport. Verify the selected project and database before database work.

These recipes use stdio. Use an installed server compatible with the
[shipped workflow requirement](AGENT_WORKFLOWS.md#required-compatibility-preflight).
The metadata and discovery behavior below was checked against development commit
`f2fed88`; it is not a claim about a published package. For that checkout, replace
`uvx` and its arguments with the absolute path to `venv/Scripts/python.exe` and
`["-m", "msaccess_vcs_mcp.main"]`. Keep the project-selection environment unchanged.

## Prepare each project

1. Create `.env` in the project directory, beside the project's source folder.
   Use an absolute database path, including the filename:

   ```dotenv
   ACCESS_VCS_DATABASE=C:/Projects/MyApp/database.accdb
   ```

2. Keep `.env` and `.env.local` out of version control. If you need a personal
   database override, put it in that project's `.env.local`:

   ```dotenv
   ACCESS_VCS_DATABASE=C:/Users/YourName/dev/MyApp/database.accdb
   ```

3. Choose the client recipe below. Set `ACCESS_VCS_PROJECT_DIR` in the launch
   environment, before the server loads configuration. Putting the selector
   only inside an undiscovered `.env` cannot help the server find that file.

## Select the active project in Cursor

Merge this entry into `~/.cursor/mcp.json`, preserving other servers:

```json
{
	"mcpServers": {
		"msaccess-vcs-mcp": {
			"command": "uvx",
			"args": ["msaccess-vcs-mcp@latest"],
			"env": {
				"ACCESS_VCS_PROJECT_DIR": "${workspaceFolder}"
			}
		}
	}
}
```

[Cursor's MCP documentation](https://cursor.com/docs/mcp#config-interpolation)
supports `${workspaceFolder}` in `env` and describes global registration in
`~/.cursor/mcp.json`. Open the intended project folder, restart its server, and
run the metadata check below. The global recipe's runtime interpolation remains
unqualified here, particularly for multiple workspace folders. If metadata
shows a literal variable, the home directory, or another project, use an absolute
`ACCESS_VCS_PROJECT_DIR` in that project's `.cursor/mcp.json` and verify again.

## Select the session project in Claude Code

[Claude Code's MCP documentation](https://code.claude.com/docs/en/mcp#option-3-add-a-local-stdio-server)
defines `CLAUDE_PROJECT_DIR` in the spawned server's environment. `--scope user`
registers a server across projects in `~/.claude.json`; the default scope is local.
Start Claude Code from the intended project directory.

The bridge reads `ACCESS_VCS_PROJECT_DIR`, so map the client variable before
starting the bridge. With the server installed in the named virtual environment,
register this command in PowerShell, replacing the Python path:

```powershell
claude mcp add --scope user --transport stdio msaccess-vcs-mcp -- 'C:\Tools\msaccess-vcs-mcp\venv\Scripts\python.exe' -c 'import os, runpy; from pathlib import Path; project = os.environ.get("CLAUDE_PROJECT_DIR", ""); project and Path(project).is_dir() or exit("CLAUDE_PROJECT_DIR must name an existing project directory"); os.environ["ACCESS_VCS_PROJECT_DIR"] = project; runpy.run_module("msaccess_vcs_mcp.main", run_name="__main__")'
```

The command selects the client-supplied project explicitly and exits if the
context is missing or invalid. It carries no database default. Read
`CLAUDE_PROJECT_DIR` inside the spawned process as shown, rather than assuming
`${CLAUDE_PROJECT_DIR}` expands while Claude reads its JSON configuration.

Inspect `claude mcp get msaccess-vcs-mcp` before replacing an existing registration.
If it hardcodes `ACCESS_VCS_DATABASE`, replace that entry with this recipe.
Check for a same-named local or project registration that supersedes user scope.
Restart the server and verify metadata in each project. On clients without this
project variable, use a project registration with an absolute
`ACCESS_VCS_PROJECT_DIR` instead.

## Select a named project in Claude Desktop

Claude Desktop's local-server configuration is separate from Claude Code.
Open **Settings > Developer > Edit Config**. On Windows, the file is
`%APPDATA%\Claude\claude_desktop_config.json`, as described in the
[MCP local-server guide](https://modelcontextprotocol.io/docs/2026-07-28/develop/connect-local-servers).

Merge one named entry per project into `mcpServers`. Use literal absolute paths:

```json
{
	"mcpServers": {
		"msaccess-vcs-myapp": {
			"command": "uvx",
			"args": ["msaccess-vcs-mcp@latest"],
			"env": {"ACCESS_VCS_PROJECT_DIR": "C:/Projects/MyApp"}
		},
		"msaccess-vcs-otherapp": {
			"command": "uvx",
			"args": ["msaccess-vcs-mcp@latest"],
			"env": {"ACCESS_VCS_PROJECT_DIR": "C:/Projects/OtherApp"}
		}
	}
}
```

Restart Desktop. Ask for `vcs_get_version_info` from the intended named server,
then use that server's tools for the project. Each entry has its own process and
project directory. A chat's topic does not select a different project for an
already running entry. If Desktop cannot find `uvx`, use its absolute executable
path. Preserve unrelated registrations and login settings.

## Verify configuration before database work

1. Call `vcs_get_version_info()` through the selected client/server connection.
   This check never starts or attaches to Access and never opens a database.
2. Compare `project_configuration.project_root` with the intended project.
   Explicit selection reports `resolution_method: "ACCESS_VCS_PROJECT_DIR"`.
3. Compare `target_database` with the absolute path in the intended environment
   file. An empty string means the default database is unresolved.
4. Read server stderr for `Working directory`, `Resolved project root`, and
   `Loaded .env from` or `Loaded .env.local from`. Metadata names the project and
   database; stderr and `startup_env_load` or `lazy_env_load` diagnostics identify
   the files used. Metadata does not identify which file supplied each value.
5. Apply the [compatibility preflight](AGENT_WORKFLOWS.md#required-compatibility-preflight)
   before database operations. Repeat selection checks after reconnecting or
   changing project configuration.

If the database is empty, check that the selected directory exists and contains
`.env` or `.env.local`. Check the exact `ACCESS_VCS_DATABASE` spelling and value.
An inherited or client-provided empty value blocks the base `.env` on initial
load; remove that launch setting and restart. Check personal overrides too.
An explicit project without an environment file remains selected and emits
`explicit_project_no_env`. A `cwd_fallback_no_env_found` resolution with an empty
database indicates that automatic startup discovery found no configuration.

If the path is wrong, inspect the launch `env` and both environment files.
Startup's `no new variables were loaded` message can mean the launch environment
already supplied those keys. Check `server_runtime` to confirm the configured
Python executable and package source. Use the returned `diagnostic_log_path`
to locate this process's diagnostics. Desktop also captures server stderr in
`%APPDATA%\Claude\logs\mcp-server-SERVERNAME.log`.

## Configuration precedence and reload behavior

`ACCESS_VCS_PROJECT_DIR` selects the directory; `ACCESS_VCS_DATABASE` selects the
default database. They are separate settings. Use an existing absolute project
directory. An invalid selector warns and falls back to discovery.

| Load | Precedence for keys present in these sources, highest first |
| --- | --- |
| Initial load, including lazy workspace load | `.env.local`, existing process environment such as client `env`, `.env` |
| Reload after a tracked file's modification time changes | `.env.local`, `.env`, existing process environment |

The base `.env` initially fills unset keys. `.env.local` always overwrites keys
it contains, including values supplied by the client. During reload, both files
overwrite existing values in that order. A client-provided database is therefore
not a permanent override. Keep database defaults in project files for these
global recipes.

The next tool call checks modification times of files present at the last load.
Edits to a tracked `.env` or `.env.local` trigger reload and reset logging.
Adding a previously absent file or deleting a file does not itself trigger reload.
Removing a key does not clear the old process value. Restart after file creation,
file deletion, key removal, launch-environment changes, or project changes.
Reload retains the selected root, even if an environment file edits the selector.
One process's selected project does not switch with the chat or working directory.

## Diagnose automatic workspace discovery

Explicit selection in these recipes bypasses `roots/list` and works without M58's
compatibility adapter. Without explicit selection, startup first searches upward
from the working directory for environment files or project markers, then falls
back to that directory. It does not search the installed package for a database.
When startup has no environment file, later tool calls can request MCP roots and
select the first root containing `.env` or `.env.local`.

The observed Cursor failure returned a native Windows path where the MCP SDK
required a `file:` URI. Before M58, SDK validation failed before path conversion,
and later calls did not retry. M58 commit `f2fed88` adds a narrow adapter for
absolute Windows drive paths and bounded retries. Native UNC and device paths
remain unsupported by that adapter. See
[workspace discovery and recovery](CONTRIBUTING.md#workspace-discovery-and-recovery)
for the retry policy and diagnostics.

Controlled SDK tests qualify M58's repair; an actual Cursor runtime repair is
still unverified. Connection success alone never proves environment-file loading.
The [M59 verification report](../verification/M59/REPORT.md) separates controlled
configuration checks from client runtime evidence and outstanding manual checks.
