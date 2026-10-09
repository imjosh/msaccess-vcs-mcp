# Completion, verification, and user-project tests

## Confirm completion

Whole-project export, merge, and build report success only from a terminal
completion callback. `started: true`, a dispatch acknowledgment, `async`/`sync`
markers, an Empty VBA return, progress messages, or file existence alone are
insufficient. `completion_unconfirmed` means the operation may have finished
either way. Inspect recent calls and the matching operation log before retrying.

Category and single-object calls return synchronous outcomes; `RunFilteredTests`
also returns final results. Still inspect `success`, `decisions`, errors, and
`policy_cleanup_error`, not just transport success or an imported count. A
`cancel_not_honored` result reports the actual outcome despite a cancellation
request. Interruption or uncertain delivery needs reconciliation before more work.

## Verify the requested result

Confirm database identity and, for a build, the actual reported output. Use this
operation's `log_path`/`log_excerpt` to check objects processed, skips, deletions,
and warnings. If the path is missing, use `vcs_get_log` with the correct family
and correlate the timestamp and database to the attempt; a latest log can be stale.
For log discovery and compile failures, follow the target export's
`vcs-agent-docs/troubleshooting.md` and the tool's `agent_guidance`.

Verify the changed content or behavior, not merely that an object name exists:
inspect a query's SQL, check relevant properties, run authorized read-only checks,
or re-export the affected object and review Access's normalization. A re-export
writes source, so preserve pending edits and account for the import's scope first.
VBA changes generally warrant compiling the target project. Layout changes may
need a visual check in Access; report that check as outstanding if unavailable.

`vcs_diff_database` currently compares names only for top-level `.sql` queries and
`.bas` modules. It does not compare content, recurse into module subfolders, or
verify other categories; `show_details=True` is not implemented. Use it only as
limited inventory evidence, not proof that edits reached Access.

## Relevant tests

Read the user's `vcs-agent-docs/testing.md` and their project test instructions.
After importing, run affected existing tests with `vcs_run_tests(database_path=...)`
against the **user's database**, or the confirmed rebuilt output being verified.
The installed add-in supplies the runner as a library; its installed or development
copy is not the test host for a user-project change.

Confirm the returned test keys include the intended modules/procedures and inspect
assertions, failures, errors, and EMPTY entries. A success in another database,
zero selected tests, or an all-EMPTY run does not verify the change. Partial runs
can leave older entries marked stale in the project's merged results state.

Select filters from the user's tests; the live schema defines comma-separated
syntax. Keep noninteractive execution and use a policy appropriate to test setup.
If `modTestAssert` is missing, install it through the supported add-in setup only
within the authorized project scope, or ask the user to do so. Tests can modify
data or contact external services; use the project's fixtures/backup conventions
and do not infer authorization for unrelated data operations from a sync request.

Report the database, source folder, applied/skipped objects, completion evidence,
tests actually run, and material gaps such as unavailable visual checks, missing
tests, unresolved decisions, uncertain completion, or cleanup failure.
