# AI Agent Workflows for Microsoft Access Development

This guide documents common workflows for AI agents working on users' Microsoft
Access databases with msaccess-vcs-mcp. The Python-style examples illustrate MCP
tool calls; invoke them through your client's tool interface and read its live
schemas for parameter contracts. `read_file`, `write_file`, and Git calls below
represent your editor and version-control tools, not functions supplied by MCP.

The optional [access-vcs-sync](../skills/access-vcs-sync/SKILL.md) and
[access-vcs-recover](../skills/access-vcs-recover/SKILL.md) skills package sync and
recovery guidance for installation in an agent client. This guide remains usable
without installing them. For bridge development, read
[CONTRIBUTING.md](CONTRIBUTING.md); add-in rebuilds and its own tests are covered
in [ADDIN_DEVELOPMENT.md](ADDIN_DEVELOPMENT.md).

## Required compatibility preflight

This workflow owns the server requirement: **>=0.3.0 <0.4.0** for stable
releases, plus explicit development admission **0.3.0-dev.18** (build metadata
does not affect identity). The range is assigned for the next release; no
published 0.3.0 is claimed. The authoritative consumer declaration is the shipped
[`workflow_requirement.json`](../src/msaccess_vcs_mcp/workflow_requirement.json).
The current MCP server instructions distribute this preflight without a plugin.
The server's add-in requirement is independently owned by its release; consult
[release compatibility](RELEASE_COMPATIBILITY.md) for its contract and update steps.

Before database work, and after a reconnection, server replacement, or changed
connection identity, call the existing read-only `vcs_get_version_info()` tool.
Read `mcp_version`, even if `success` is false because installation diagnostics
failed. Apply strict SemVer comparison to this workflow's own range. Accept a
stable 0.3.x at or above 0.3.0, or exactly 0.3.0-dev.18 ignoring build metadata.
Other prereleases require a workflow requirement update. Database work proceeds
only after this check succeeds; a previously accepted connection does not admit
a newly connected server. When reconnection cannot be detected, check before
each new database workflow rather than retaining admission across workflows.

For an older server, tell the user: "The connected MCP server is {installed};
this workflow requires >=0.3.0 <0.4.0 (minimum 0.3.0), or the explicitly admitted
development build 0.3.0-dev.18. Update the MCP server to a compatible release or
that development build, restart its process, and reconnect the MCP client."
Stop database calls while that requirement is unmet. For a newer unsupported
major or 0.4+ minor, explain the supported combination and offer a workflow
that supports that server or reconnection to a supported server. Upgrading the
same server again does not resolve that mismatch.

If the version tool is absent, fails without `mcp_version`, or returns a
malformed/missing version, tell the user: "MCP server compatibility is
unconfirmed: installed version {value or unknown}; required >=0.3.0 <0.4.0,
minimum 0.3.0 (development admission 0.3.0-dev.18). Verify the server configured
for this connection, update to a compatible build if needed, restart and
reconnect, then repeat the read-only check." Stop dependent calls. Older servers
need not expose new requirement fields: their existing valid `mcp_version`
alone can be compared. Use the client's tool inventory to establish that the
version tool is absent; do not call unsupported APIs to probe for features.

The server independently admits its installed add-in. If metadata reports an
incompatible `addin_compatibility`, stop add-in work and relay its `error`,
`component`, `installed_version`, `required_range`, `minimum_version`, and
`recovery_action`. Relay the same details if an operation returns
`version_incompatible`, `version_unconfirmed` or `compatibility_session_invalid`; never retry the operation
through VBA or another tool to bypass the refusal. A failed check authorizes
this notification, not installing, rebuilding, downgrading, or upgrading.
Proceed after a compatible check without feature probes solely for release
compatibility. Requested mode/policy acknowledgments remain required per call.

After a refusal, server-permitted metadata/recent-call, dialog/status/recovery,
HTTP cancellation and raw DAO inspection retain their own guards. They do not
authorize replay, installation, rebuilding or overriding compatibility.
When using released skills, also check their exact [release provenance](../skills/release-policy.md).

## Overview

The MCP tool enables AI agents to work iteratively with Access databases by:
1. Exporting database objects to text-based source files
2. Modifying source files (queries, VBA modules, etc.)
3. Merging changes back into the database
4. Testing and iterating

## Establish the target and confirm outcomes

Locate and read the target project's exported `AGENTS.md`, normally in
`<database-name>.accdb.src`, plus the relevant `vcs-agent-docs/` references. These
are the authority for editing the actual export format. The paths below are
examples; a project can configure a different source folder.

Before exporting or importing, confirm that folder against the intended database:

```python
vcs_call_vba(
    database_path=r"C:\mydb.accdb",
    function_name="VCS.API",
    args=["GetExportFolder"],
)
```

Export writes only to the configured folder; omit deprecated `output_dir`.
Import's `source_dir` does not change the configured folder it reads. Preserve
pending source edits and unsaved database work before a replacement or export.

Check every operation's returned outcome before proceeding. Whole-project export,
merge, and rebuild require a terminal completion callback; a start acknowledgment
or `completion_unconfirmed` is not confirmed success. For uncertainty, inspect
recent calls and the matching operation log before retrying a mutation. Read
[DIALOGS.md](DIALOGS.md) for decision policies, cancellation, and recovery tools;
the portable skills carry additional conditional
[operation-choice](../skills/access-vcs-sync/references/operation-choice.md) and
[verification](../skills/access-vcs-sync/references/verification.md) guidance.

## Core Workflow Pattern

The usual edit-and-sync cycle is:

```
┌─────────────┐
│  Database   │
└──────┬──────┘
       │ Export
       ↓
┌─────────────┐
│Source Files │ ← AI Agent reads/edits
└──────┬──────┘
       │ Merge Build
       ↓
┌─────────────┐
│  Database   │ ← Testing
└──────┬──────┘
       │ Iterate
       ↓
```

## Common Workflows

### 1. Modify a Query

**Use case:** Update SQL logic, fix bugs, or optimize a query

**Steps:**
```python
# 1. Capture Access changes only after preserving pending source edits
vcs_export_database(database_path=r"C:\mydb.accdb")

# 2. Read the query source file
query_sql = read_file("C:\\mydb.accdb.src\\queries\\CustomerReport.sql")

# 3. Modify the SQL
# (AI agent makes changes to the SQL)

# 4. Write the updated query
write_file("C:\\mydb.accdb.src\\queries\\CustomerReport.sql", updated_sql)

# 5. Merge changes back into database
result = vcs_import_objects(
    database_path=r"C:\mydb.accdb",
    source_dir=r"C:\mydb.accdb.src",
    decision_policy="block",
)
# Proceed only after confirmed success; reconcile decisions or uncertainty first

# 6. Verify the SQL and run an authorized query check on C:\mydb.accdb
# A listed object name alone does not verify the change
```

**Tips:**
- Query files are in `queries/` folder with `.sql` extension
- Follow the exported query guidance for the editable SQL and companion files
- Test with sample data before committing

### 2. Add or Update VBA Code

**Use case:** Create new functions, fix bugs, or refactor VBA modules

**Steps:**
```python
# 1. Capture relevant Access changes before editing, if needed
vcs_export_database(
    database_path=r"C:\mydb.accdb",
    object_types=["modules"]
)

# 2. Read the module source file
module_code = read_file("C:\\mydb.accdb.src\\modules\\Utilities.bas")

# 3. Add or modify VBA code
# (AI agent makes changes to the module)

# 4. Write the updated module
write_file("C:\\mydb.accdb.src\\modules\\Utilities.bas", updated_code)

# 5. Merge changes back
result = vcs_import_objects(
    database_path=r"C:\mydb.accdb",
    source_dir=r"C:\mydb.accdb.src",
    decision_policy="block",
)
# Proceed only after confirmed success; reconcile decisions or uncertainty first

# 6. Compile to validate
result = vcs_compile_vba("C:\\mydb.accdb")
if not result["success"]:
    # Stop — ask the user to Debug → Compile in the VBE and paste the
    # code snippet around the highlighted error line. See agent_guidance.
    pass

# 7. Test the functions
# (Open VBE and test the new/updated functions)
```

**Tips:**
- Module files are in `modules/` folder with `.bas` (standard modules) or `.cls` (class modules) extensions
- Preserve the `Attribute VB_Name` header
- Use Option Explicit for type safety
- Follow the project's procedure-comment conventions
- On a compile failure, follow `agent_guidance` and the exported troubleshooting
  reference; use the highlighted error context before choosing a fix

### 3. Create a New Database Object

**Use case:** Add a new query, module, or other object

**Steps:**
```python
# 1. Export existing database if needed, after preserving pending source edits
vcs_export_database(database_path=r"C:\mydb.accdb")

# 2. Create new source file
# For a new query:
new_query = """SELECT CustomerID, CompanyName, ContactName
FROM Customers
WHERE Active = True
ORDER BY CompanyName
"""

write_file("C:\\mydb.accdb.src\\queries\\NewCustomerList.sql", new_query)

# 3. Merge into database
result = vcs_import_objects(
    database_path=r"C:\mydb.accdb",
    source_dir=r"C:\mydb.accdb.src",
    decision_policy="block",
)
# Proceed only after confirmed success; reconcile decisions or uncertainty first

# 4. Verify the object was created
objects = vcs_list_objects("C:\\mydb.accdb")
# Inspect the returned query inventory, then verify its actual SQL and behavior
```

**Tips:**
- Follow existing file naming conventions
- Use a compatible existing export as a pattern; preserve required companion files
- Use descriptive names
- Test immediately after creation

### 4. Bulk Export for Version Control

**Use case:** Initial export to git or periodic full export

**Steps:**
```python
# 1. Full export to the configured source folder
result = vcs_export_database(database_path=r"C:\mydb.accdb", full_export=True)

# 2. Check the outcome, export_path, and operation log before proceeding
# Inspect the exported files and Git diff for missing objects or unexpected changes

# 3. Commit to version control (using git tools)
git_add("C:\\mydb.accdb.src")
git_commit("Initial database export")
```

**Tips:**
- For an explicitly requested complete export, pass `full_export=True`
- Subsequent exports use "fast save" (only changed objects)
- Review Export.log for details
- Commit frequently for granular history

### 5. Pull Changes and Merge

**Use case:** Integrate changes from other developers

**Steps:**
```python
# 1. Commit any local changes first
vcs_export_database(database_path=r"C:\mydb.accdb")
git_add("C:\\mydb.accdb.src")
git_commit("My changes before pull")

# 2. Pull changes from remote
git_pull()

# 3. Review Git changes and any independent database edits before merging
# vcs_diff_database is only a limited query/module name inventory, not a content diff

# 4. Merge source changes into database
result = vcs_import_objects(
    database_path=r"C:\mydb.accdb",
    source_dir=r"C:\mydb.accdb.src",
    decision_policy="block",
)
# Proceed only after confirmed success; reconcile decisions or uncertainty first

# 5. Test the merged result
# (Verify database functions correctly)

# 6. After confirmed merge and verification, capture Access normalization if needed
vcs_export_database(database_path=r"C:\mydb.accdb")
git_add("C:\\mydb.accdb.src")
git_commit("Merged changes from team")
```

**Tips:**
- Preserve local source and database work before integrating remote changes
- Review diffs carefully
- Test thoroughly after merge
- Resolve Git conflicts in source files before import
- Database-versus-source conflicts require an authorized winner: default `block`
  reports `decision_required`; inspect the decisions instead of silently choosing
- Use `prefer_source` or `prefer_database` only when the user's intent authorizes
  that outcome; source-winning actions may include deletion
- See [DIALOGS.md](DIALOGS.md) for policy semantics and partial outcomes

### 6. Build Fresh Database from Source

**Use case:** Clean build, deployment, or distribution

**Steps:**
```python
# 1. Build from source files
result = vcs_rebuild_database(
    source_dir=r"C:\mydb.accdb.src",
    output_path=r"C:\builds\mydb_v1.0.accdb"
)

# 2. Verify build succeeded
if result.get("success"):
    print(f"Database built: {result['output_path']}")
    # Review the returned log_path/log_excerpt when available
else:
    print(f"Build failed or unconfirmed: {result.get('error')}")

# 3. Only after confirmed completion, verify and test the reported output_path
# Object-name inventory alone does not establish content or behavior
```

**Tips:**
- Build from source creates a fresh database; select a fresh or authorized output
- It uses an isolated Access host and refuses an already-open output, preserving
  user-owned windows; `requested_output_path` alone is not the resulting file
- `template_path` is currently accepted but not applied; report a required
  template as a limitation rather than promising it was used
- Use for deployments and releases
- Read this attempt's returned `log_path` or a correlated `Build_*.log` for issues
- Test thoroughly before distributing

### 6b. Rebuild the VCS add-in from source

For add-in source changes, follow
[Rebuilding the VCS add-in](ADDIN_DEVELOPMENT.md#rebuilding-the-vcs-add-in).
`vcs_rebuild_database` is for user projects; use `vcs_rebuild_addin` for the add-in.

### 6c. Run the add-in's own test suite

For add-in development, follow
[Running the add-in's own tests](ADDIN_DEVELOPMENT.md#running-the-add-ins-own-tests).
User-project tests target the user's database, as shown below.

### 7. Iterative Development Cycle

**Use case:** Rapid development with frequent testing

**Steps:**
```python
# Capture Access edits at the start, after preserving source edits, if needed
vcs_export_database(database_path=r"C:\mydb.accdb", object_types=["modules"])

# Edit one coherent change using the project's exported guidance
# Apply it and check the final result before running tests
vcs_import_objects(
    database_path=r"C:\mydb.accdb",
    source_dir=r"C:\mydb.accdb.src",
    decision_policy="block",
)

# Run relevant tests defined in this USER project, after confirmed import
vcs_run_tests(database_path=r"C:\mydb.accdb", filter="modTestUtilities")
# Inspect actual test keys, assertions, failures/errors, and EMPTY entries
# Verify the requested behavior, review the source diff, then commit as authorized
```

Export again only when needed to capture Access normalization or subsequent
Access edits; preserve source work first. A failed or uncertain operation ends
this iteration until its outcome is reconciled. Test filters must come from the
user's project, not the example name or the add-in's development suite.

### 8. Apply a single object or category

For a deliberate replacement of one named object:

```python
vcs_import_object(
    database_path=r"C:\mydb.accdb",
    object_type="module",
    object_name="Utilities",
)
```

This path bypasses index-based conflict detection even with `block`; establish
source-replacement intent first. Use conflict-aware merge when independent
Access edits must be reconciled. A single import is not a way around a previously
reported conflict.

For complete categories:

```python
vcs_import_objects(
    database_path=r"C:\mydb.accdb",
    source_dir=r"C:\mydb.accdb.src",
    object_types=["queries"],
    full_import=False,
    decision_policy="block",
)
```

Category import deletes objects absent from the category source and takes no
backup. Confirm complete source and intended deletions first. `full_import=True`
forces source replacement and skips conflict detection; it is ignored for a
whole-project merge. Use the operation-choice reference for scope and backup
tradeoffs.

## Best Practices

### File Organization

Follow the target export's `AGENTS.md` and file-type references for structure,
encoding, paired definitions/code, and the index. Preserve existing encoding and
BOM conventions when editing; leave the generated binary index to the add-in.
The actual exported project, rather than a generic directory sketch, determines
where objects live.

### Error Handling

```python
# Always check for errors
result = vcs_export_database(database_path=db_path)

if not result["success"]:
    print(f"Export failed: {result.get('error')}")
    # Inspect error_pattern, decisions, and the matching operation log
    # Reconcile completion_unconfirmed/timeouts before issuing another mutation
```

### Testing

After confirmed import or rebuild, read the user's exported testing guidance.
Compile VBA changes using `vcs_compile_vba` and follow its `agent_guidance` if
compilation fails. Run relevant tests on the user's database or confirmed rebuilt
output; the installed add-in supplies the runner as a library.

Check that returned test keys and assertions cover the intended change. A run
against the development add-in, zero selected tests, or an all-EMPTY result does
not validate the user project. Some tests mutate data or contact services; use
the project's fixture and authorization conventions. Verify affected queries,
layouts, and data behavior as needed, and report checks that remain outstanding.

### Version Control

Commit frequently with descriptive messages:
```bash
git commit -m "Add customer search query"
git commit -m "Fix date calculation in Utilities module"
git commit -m "Update invoice report layout"
```

## Troubleshooting

### Add-in Not Found

**Error:** `VCS add-in not found`

**Solution:**
1. Install MSAccess VCS add-in
2. Or set `ACCESS_VCS_ADDIN_PATH` in `.env`

### Import Fails

**Error:** Import/merge build fails

**Solution:**
1. Inspect the attempt's returned `log_path`; otherwise correlate the `Merge_*.log`
   with this database and timestamp using the exported troubleshooting reference.
2. Verify source files against the project's exported guidance.
3. Inspect `decision_required` and cleanup errors before proceeding.
4. A full rebuild needs its own intended outcome and authorized output; a failed
   merge alone is not a reason to replace the entire database.

### Objects Not Exporting

**Error:** Some objects missing from export

**Solution:**
1. Correlate the attempt's Export log and outcome with the intended database.
2. Check export scope and VCS options (`vcs-options.json`).
3. If an open object or native save dialog blocked the operation, inspect it using
   [DIALOGS.md](DIALOGS.md); preserve unsaved work before closing anything.

### Merge Conflicts

**Error:** Git merge conflicts in source files

**Solution:**
1. Resolve conflicts in source files (not in Access)
2. Use standard git conflict resolution
3. Test merged result in Access
4. Re-export to verify

### Busy server, timeout, or blocked dialog

Inspect before retrying:

```python
vcs_get_recent_calls(limit=10)
vcs_automation_status(database_path=r"C:\mydb.accdb")
vcs_list_dialogs(database_path=r"C:\mydb.accdb")
```

Match the operation ID, database, timestamps, and logs. Readiness is not proof a
previous mutation completed; recent-call absence is not proof it never started.
Follow [DIALOGS.md](DIALOGS.md) for supported dismissal and cancellation, or the
[recovery skill](../skills/access-vcs-recover/SKILL.md) for reconciliation steps.
Preserve user-owned Access windows. Retry only after the prior outcome and any
required decision are resolved.

## Advanced Patterns

### Conditional Logic Updates

When updating complex VBA logic:
1. Preserve and export the relevant current version if needed
2. Record non-obvious constraints in the project's existing comment conventions
3. Make incremental changes
4. Test each change
5. Commit working versions

### Schema Migrations

When changing table structure:
1. Read the target project's table-definition and data guidance
2. Edit the actual exported representation; do not assume it is SQL CREATE TABLE
3. Handle data migration separately
4. Test with sample data first
5. Document migration steps

### Form and Report Development

Forms and reports are exported but harder to edit as text:
1. Export for version control
2. Make visual changes in Access
3. Export again to capture changes
4. Commit with descriptive message

## Resources

- [MSAccess VCS Add-in Documentation](https://github.com/joyfullservice/msaccess-vcs-integration/wiki)
- Target export's `AGENTS.md` and `vcs-agent-docs/` — authoritative file-format,
  testing, and compile-error guidance for that project
- [Dialogs and noninteractive automation](DIALOGS.md) — policies and recovery
- [VBA Integration Guide](VBA_INTEGRATION.md)
- [AGENTS.md](https://github.com/joyfullservice/msaccess-vcs-integration/blob/main/Version%20Control.accda.src/AGENTS.md) - Comprehensive file structure guide

## Support

For issues or questions:
1. Check the [Wiki](https://github.com/joyfullservice/msaccess-vcs-integration/wiki)
2. Read the returned attempt `log_path`; correlate a missing path by time/operation identity
   using the [operation log families](CONTRIBUTING.md#vcs-operation-logs-written-by-the-add-in-not-the-server)
3. Open an issue on GitHub
4. Include error messages and context

## Automatic compatibility sessions (X17)

Workflow preflight is an independent workflow requirement. Operational commands automatically negotiate the server/add-in session even without agent preflight. Only the handshake transmits versions; every dependent command carries a validated session ID. Read-only metadata does not establish admission. See [release compatibility](RELEASE_COMPATIBILITY.md) for cache identity, invalidation, migration and completed X17 qualification.

## Optional skill installation and distribution

The maintained sources are this MCP repository's `skills/` directory. Release
the skills with the MCP version they describe and use that release's supported
add-in combination. Preserve the shared `skills/release-policy.md` and each skill's
references when delivering them. Individual directories are not independently
supported artifacts; install the skills supplied by the connected server's release.

The wheel bundles the complete tree under `msaccess_vcs_mcp/guidance/`: both
skills, shared release policy, a copy of the authoritative workflow declaration,
and generated `release.json`. The sdist contains their maintained sources; wheel
builds derive provenance from the package version. Preserve the complete tree.
For source checkouts, compare the server identity in `src/msaccess_vcs_mcp/__init__.py`
and read the original `workflow_requirement.json` beside it. Release packages
carry their own manifest; independently copied skill directories are unsupported.
Client integration packages are the preferred delivery route; their implementation
belongs to the separate X14 project. Skills remain optional for following this guide.
