# Reconcile before retrying

## Collect evidence

Keep the original `success`, `error_pattern`, error text, `decisions`, paths,
call time, operation ID, and interruption/cleanup fields together. A client
timeout (`-32001`) discards a response; it does not establish that Access stopped.
`vcs_get_recent_calls` is gate-exempt. Match entries by tool, target, parameters,
and time. Its usage log may be disabled, missing, rotated, or have no entry yet
for a running call; absence is not proof of nonexecution.

Read the exact `log_path` from the attempt, using direct file access if the gate
is occupied. `log_excerpt` may already carry the error. Follow the target export's
troubleshooting reference for filesystem discovery. If responsive and the path is
lost, `vcs_get_log` retrieves the latest log: its live description identifies the
family for each operation. Correlate the log's time, database, and source folder;
a latest file may belong to another run.
`vcs_get_log` uses the COM gate, so it cannot unblock a call holding that gate.

Compare the operation's terminal evidence and the actual requested state. An
output file, object count, `started: true`, launch acknowledgment, or idle window
does not prove completion. Whole-project export, merge, and build use terminal
callbacks; without one, `completion_unconfirmed` is an unknown outcome, not a
failed operation safe to repeat. A synchronous single-object/category result or
test result is final only for the work it actually performed.

## Respond to the observed state

| State | Next action |
| --- | --- |
| `server_busy` | Read `busy_with`: the single server gate is shared across databases/clients. This refused request has not entered its handler. Let the in-flight call finish; follow `retry_after_seconds` without rapid repeated calls. |
| Add-in `operation_already_running` | Inspect the active call/root. Wait for its owner to finish; do not weaken its policy or issue a second mutation. |
| Client or server timeout; `completion_unconfirmed` | Inspect recent calls, this attempt's log, and readiness. Check for partial effects before any retry. |
| Confirmed terminal success | Verify the requested contents in the intended database; recover transport/reporting, without repeating the mutation. |
| Confirmed failure or cancellation | Identify partial writes and cleanup status, then correct the cause. Re-run only the still-needed, authorized work. |
| Evidence remains uncertain | Preserve diagnostics and stop further mutations; ask for the concrete missing observation or action. |

A busy gate on a different database is still busy. `ready: true` is target-specific
and may not predict that another database has released the shared gate. Use the
busy response and original call together. Repeated unchanged busy results warrant
dialog/status inspection, not retries in a loop or cancellation of another user's
operation. Inspection tools' `tool_timeout` and `worker_capacity_unavailable`
also require Access responsiveness to improve; flooding them consumes workers.

## Readiness and lifecycle

`vcs_automation_status` is gate-exempt and checks a confirmed Access process,
response, break mode, blocking dialogs, and gate state for that database.
`ready: true` does not erase `last_interruption` or establish the original result.
For `blocking_dialog` or `vba_break`, use the dialog reference. For ambiguous or
unconfirmed process identity, resolve the instance before any UI action.
`no_windows_to_probe` is inconclusive; `access_not_running` confirms absence of
the matched process. Reconcile logs/output before reopening after a process exits.

VBA workers may time out while Access still runs code or waits on a dialog.
After responsiveness returns, the server probes automatically on the next call;
it may recycle only confirmed server-owned instances under its ownership rules.
Keep a user-owned hung window intact and request the user's VBE/dialog action.
A server restart or larger timeout does not establish that an earlier mutation
failed; reconcile it first rather than masking the cause.

## Interruption and cleanup

`execution_interrupted` records an ended execution/error dismissal, not rollback.
`interruption_uncertain` means delivery or cancellation outcome was not confirmed.
Preserve the original failure/decisions even when readiness becomes true, and
check partial writes before a retry. An uncertain dismissal can have acted;
reinspect the window instead of repeating the click blindly.

Request-owned policies end with their operation. MCP sends no persistent
setter or cleanup call. An old pair may report `policy_cleanup_error`; preserve
that historical diagnostic and reconcile ownership before more mutations.
`ClearOperationPolicy` is now harmless and cannot resolve a blocked conflict.
Do not force global mode or remove request policy to bypass deletion protection.
Other cleanup errors remain secondary to the original operation result.

An incompatible request API requires the matching repaired add-in. `api_self_dispatch` is an add-in defect; changing the
host database or retrying cannot repair it. Respect installed-add-in and write
refusals; do not move to raw VBA/COM or alter server protections as a shortcut.
