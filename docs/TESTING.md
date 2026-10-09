# Testing Guide for msaccess-vcs-mcp

Use the project virtual environment and Python >=3.11. Contributor setup, COM
ownership and server log discovery live in [CONTRIBUTING.md](CONTRIBUTING.md).

```powershell
.\venv\Scripts\Activate.ps1
pytest -m "not integration"
pytest tests/test_usage_logging.py -v
pytest --cov=msaccess_vcs_mcp --cov-report=html -m "not integration"
```

## Integration prerequisites

Integration runs need local Windows Access and the **supported server/add-in
combination**, established by [compatibility preflight](AGENT_WORKFLOWS.md#required-compatibility-preflight).
Use a separately authorized test installation; installing "latest" is not proof
of compatibility. Follow upstream Access support; 32-bit Access remains untested.
This guide defines no Windows version range or Python maximum.

For the add-in's own suite, follow [ADDIN_DEVELOPMENT.md](ADDIN_DEVELOPMENT.md#running-the-add-ins-own-tests)
and its maintained disposable complete-host lifecycle. User-project tests target
the intended user database or the confirmed rebuilt output. Never run the add-in
suite against the installed library, the primary development copy or `Testing/`.

Database writes and read-only SQL execution have separate permissions. Obtain the
needed authorization and configure the fixture's MCP options before the run;
setup never grants those permissions automatically. Keep production databases,
source edits, Trust Center settings and user-owned Access windows out of fixtures.

```powershell
pytest -m integration
```

## Executable export, merge and named-output recipe

The script below uses a persistent public MCP session. Every tool call is awaited;
normal server startup owns callbacks and mutual admission. Python imports of the
async tool handlers must also be awaited, but a bare import does not start the
callback server and whole exports/builds may return `completion_unconfirmed`.
The tool notation in [AGENT_WORKFLOWS.md](AGENT_WORKFLOWS.md) is illustrative MCP
client notation, not a synchronous Python program.

Provide a **closed disposable fixture** made from an authorized test project,
containing `TestQuery` with `SELECT 1 AS TestValue`. Its configured export folder
must be dedicated to this run, with no source edits to lose; the built output must
be a fresh absolute path in an existing folder. The fixture and built output need
read-only SQL permission so the oracle can execute the query. Preserve a seed and
evidence separately. This recipe modifies the fixture query to return 2.

To create a blank fixture instead, DAO `DBEngine.CreateDatabase` can create a new
unique file and query without an Access application process. Close only that DAO
database handle. If fixture preparation uses Access, use an isolated created
application, record its PID **and creation time**, and reverify both immediately
before `CloseCurrentDatabase`/`Quit`. Unknown identity leaves the process alone;
attaching `Dispatch("Access.Application")` followed by `Quit` is unsafe.

Save this script outside `tests/` (it is a deliberate integration run, not an
automatically collected test). Run it from the activated venv:

```powershell
python .\scratch\check_workflow.py C:\scratch\run\Fixture.accdb C:\scratch\run\Built.accdb
```

```python
import asyncio
import json
import re
import sys
from pathlib import Path

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client
from msaccess_vcs_mcp import __version__
from msaccess_vcs_mcp.cli import stdio_server_environment
from msaccess_vcs_mcp.compatibility import workflow_requirement


def terminal(result):
    assert isinstance(result, dict), result
    assert result.get("success") is True, result
    for flag in ("completion_unconfirmed", "cancelled", "execution_interrupted",
                 "interruption_uncertain", "policy_cleanup_error"):
        assert not result.get(flag), result
    assert not result.get("error_pattern"), result
    return result


async def check_workflow(database, output):
    database, output = Path(database), Path(output)
    assert database.is_absolute() and output.is_absolute()
    database, output = database.resolve(), output.resolve()
    assert database.is_file() and output.parent.is_dir()
    assert not output.exists(), "Use a fresh output for this recipe"
    env = stdio_server_environment()
    # Server connection cleanup closes only its identity-confirmed created hosts.
    # Fixture prerequisites exclude existing user-owned holders; never close them.
    env["ACCESS_VCS_LEAVE_ACCESS_OPEN"] = "false"
    params = StdioServerParameters(
        command=sys.executable, args=["-m", "msaccess_vcs_mcp"], env=env,
    )
    async with stdio_client(params) as streams:
        async with ClientSession(*streams) as session:
            await session.initialize()

            async def call(name, **arguments):
                reply = await session.call_tool(name, arguments)
                texts = [item.text for item in reply.content if item.type == "text"]
                assert not reply.isError and len(texts) == 1, reply
                result = json.loads(texts[0])
                print(name, json.dumps(result, ensure_ascii=False))  # retain evidence
                return result

            metadata = await call("vcs_get_version_info")
            observed = metadata.get("mcp_version")
            assert observed == __version__, metadata  # source-release provenance
            assert workflow_requirement().reason(observed) is None, metadata
            terminal(metadata.get("addin_compatibility"))
            assert metadata.get("async_available") is True, metadata
            # Repeat the preflight if this connection is replaced. Metadata does
            # not establish a VBA session: each dependent call admits automatically.

            async def query_value(target):
                result = terminal(await call(
                    "vcs_execute_sql", database_path=str(target),
                    sql="SELECT * FROM TestQuery", max_rows=2,
                ))
                assert not result.get("truncated"), result
                return result["rows"]

            assert await query_value(database) == [{"TestValue": 1}]
            folder = terminal(await call(
                "vcs_call_vba", database_path=str(database),
                function_name="VCS.API", args=["GetExportFolder"],
            ))
            source = Path(folder["result"]).resolve()
            assert source.is_relative_to(database.parent), "Use a dedicated fixture folder"
            # Confirm this is the fixture's dedicated configured folder before running.
            exported = terminal(await call(
                "vcs_export_database", database_path=str(database),
            ))
            assert Path(exported["export_path"]).resolve() == source, exported
            query_file = source / "queries/TestQuery.sql"
            original = query_file.read_text(encoding="utf-8-sig")
            modified, count = re.subn(r"\b1\s+AS\s+TestValue\b", "2 AS TestValue",
                                      original, flags=re.IGNORECASE)
            assert count == 1, original
            query_file.write_text(modified, encoding="utf-8-sig", newline="\r\n")
            merged = await call(
                "vcs_import_objects", database_path=str(database),
                source_dir=str(source), decision_policy="prefer_source",
            )  # this disposable-fixture run expressly selects source replacement
            if merged.get("log_path"):
                print("This merge attempt log:", merged["log_path"])
            terminal(merged)
            assert await query_value(database) == [{"TestValue": 2}]
            built = terminal(await call(
                "vcs_rebuild_database", source_dir=str(source), output_path=str(output),
            ))
            confirmed = Path(built["output_path"]).resolve()
            assert confirmed == output and confirmed.is_file(), built
            assert await query_value(confirmed) == [{"TestValue": 2}]
            # Do not export this output to guessed alternate folders: its configured
            # folder may be the original source. The query results are the content oracle.
            print("Verified fixture merge and confirmed rebuilt query result")


if __name__ == "__main__":
    asyncio.run(check_workflow(*sys.argv[1:]))
```

Stop on refusal, cancellation, timeout, cleanup error or uncertain completion.
Retain the printed result, operation ID/time and returned `log_path`. Inspect
recent calls and readiness through the same client, then reconcile contents before
retrying. If the session is gone, reconnect and repeat preflight; a new connection
does not turn the previous operation into a success. Never unwrap handlers or use
`AccessConnection` to get past the public admission/permission boundary.

The assertions establish a query-content round trip, not whole-database, layout
or binary equivalence. A filename, inventory entry or success print alone is not
an oracle. For further coverage, run relevant tests on the confirmed output and
check intended tests, assertions, failures, ERROR/EMPTY and stale results.

## Logs, negative cases and performance

Use this attempt's returned `log_path` or correlate timestamp, target and operation
identity with the [maintained log families](CONTRIBUTING.md#vcs-operation-logs-written-by-the-add-in-not-the-server).
Imports produce `Merge_*.log`, exports `Export_*.log`, full builds `Build_*.log`;
fixed `Export.log`/`Build.log` filenames and newest-file selection are insufficient.
`vcs_get_log` itself requires admission; permitted recent-call/status inspection
remains available after refusal. Discover active server JSONL paths from metadata.

Test missing/unsupported add-ins through mocked discovery in
`tests/test_version_compatibility.py`, rather than renaming an installed library.
Test write refusal using `ACCESS_VCS_DISABLE_WRITES=true` in an isolated test
configuration. Destination mismatch, callback-free starts and installed-target
protection have mocked public-boundary tests; preserve them when changing tools.

Measure elapsed time around awaited calls and report the terminal outcome with it.
Whole exports do not promise an object-count/type breakdown; do not divide timing
by an absent `exported_count` or infer incremental correctness from counts.
Compare actual changed content or query behavior for the requested scope.

Keep fixture files and logs until results have been recorded and all owned handles
are closed. Only the test's confirmed created processes may be closed. Ambiguous or
user-owned holders remain open and prevent disposal; report them. Add-in suite
whole-host disposal follows its maintained lifecycle, separately from user projects.

## Continuous integration

Use `pytest -m "not integration"` on runners without Access. Native integration
needs an authorized compatible Windows Access/add-in installation and isolated
fixtures. Passing mocked contracts does not qualify a desktop client's registration,
a live database build or the native behavioral gates. X14 owns client setup.
