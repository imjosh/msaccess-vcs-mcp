# Dialogs and noninteractive Access automation

This reference explains noninteractive workflows and recovery tools for agents
operating on users' Access databases, and the contracts contributors must preserve.
For common edit/export/import recipes, see [AGENT_WORKFLOWS.md](AGENT_WORKFLOWS.md).
The optional [sync](../skills/access-vcs-sync/SKILL.md) and
[recovery](../skills/access-vcs-recover/SKILL.md) skills provide portable workflow
entry points. Read live tool schemas for parameter contracts.

X17 automatically admits the installed/loaded add-in and validates dependent
commands. [Release compatibility](RELEASE_COMPATIBILITY.md) owns declarations,
cache invalidation, migration and permitted recovery after refusal. Capability
probing is optional. Request policies travel on individual operations.

Agent calls with policy parameters default to a scoped noninteractive mode. Interactive ribbon and
Immediate Window use is unchanged unless a decision policy is passed.

`DoCmd.SetWarnings` is not a dialog suppressor. It does not cover `MsgBox`,
the conflict form, VBA runtime dialogs, or compile errors.

## Request-owned policies

`vcs_import_objects`, `vcs_import_object`, `vcs_export_object`,
`vcs_export_database`, and `vcs_run_tests` default to `noninteractive=True`
and `decision_policy="block"`. Each operation receives the policy as its own
argument. The add-in retains it through queued continuations and releases it
when that operation ends. MCP sends no persistent policy setter or cleanup call.
Other requests cannot replace the active operation's policy, source, or journal.

For import and export tools, `noninteractive=False` omits the policy and preserves
legacy add-in behavior. It does not issue `SetInteractionMode(0)`. Legacy agent
calls can automatically apply the requested import or export action, including
deletions. Omitting a policy is not recovery for a blocked conflict or deletion.
Automation tests remain headless and reject `noninteractive=False`.

`SetOperationPolicy` is retired and returns `policy_activation_retired`.
`ClearOperationPolicy` is a harmless success. Standalone `SetInteractionMode(2)`
returns a refusal without changing the effective mode or active operation.
Older clients that require setter acknowledgments need migration to request
arguments. Older add-in pairs that bypassed single-object conflict checks are
historical limitations and do not provide this contract.

## Conflict outcomes

The following rules apply to detected conflicts in either single-object export
branch, single-object import, category calls, and full operations with a policy.
Unchanged destinations, identical content, and one-sided edits can transfer safely.
Missing or stale history does not authorize discarding independent changes.

| Detected conflict | `prefer_source` | `prefer_database` | `skip` or `decline` | `block` |
| --- | --- | --- | --- | --- |
| Import would overwrite an edited database object | Import source | Preserve object | Preserve object | Require decision |
| Export would overwrite an edited source file | Preserve file | Export database version | Preserve file | Require decision |
| Import would delete an edited database object with missing source | Require decision | Preserve object | Preserve object | Require decision |
| Export would delete a tracked, changed orphan file | Preserve file | Require decision | Preserve file | Require decision |
| Export finds a new, untracked source file | Preserve file | Preserve file | Preserve file | Preserve file |

`skip` and `prefer_database` differ on export. No noninteractive policy authorizes
deletion of a conflicting edited object or tracked, changed orphan file in this
stage. Review and resolve those deletions interactively, or preserve the item if
that matches the user's intent. Existing cleanup of unchanged indexed orphans is
unchanged. A conflict list containing a blocked deletion applies no resolutions
to that list, but earlier scanning and non-conflict work may already have occurred.

`decline` answers No, Cancel, or Abort on confirmations and preserves conflicts.
Yes/No/Cancel receives Cancel. Preference policies do not approve unrelated
confirmations. Informational OK-only notices can be acknowledged unattended.

## Inspect results before retrying

An actual transfer reports its transfer outcome. A preserving call reports
`skipped: true` and leaves the conflicting destination unchanged. Read the returned
`transferred`, `skipped`, `decisions`, and error fields rather than inferring success
from a journal resolution string. Journals can add resolution values and fields.
MCP preserves them through synchronous results and terminal callbacks.
Successful transfers update synchronization history. Skipped and blocked objects
remain unsynchronized so a later export can detect their conflicts after reopen.

For `decision_required`, inspect the journal's object, direction, requested action,
classification, policy, and recovery information. Reuse an existing user instruction
only when it covers that exact action and scope. Otherwise ask for the unresolved
choice. Never automatically retry with `prefer_source` or `prefer_database`, change
scope to bypass checks, or remove the policy to make a blocked deletion proceed.

In ordinary export/edit/import/test cycles, check the final transfer result before
running the relevant tests. Preserve independent edits. Reconcile skipped, blocked,
partial, failed, or uncertain results before the next mutation.

## Dispatch results and compatibility

The repaired intermediate API accepts trailing request policies:

```text
Export(policy), FullExport(policy), ExportVBA(policy)
Build(sourceFolder, policy), MergeBuild(policy)
ImportObject(type, name, policy), ExportObject(type, name, policy)
ImportByType(types, fullImport, policy), ExportByType(types, fullExport, policy)
RunFilteredTests(policy)
```

`APIAsync(callbackInfo, method, arg1, arg2, arg3)` carries up to three method
arguments. The add-in queues a typed JSON array, preserving booleans, empty strings,
and embedded pipes. A successful admission returns `async: true` and the supplied
public `operation_id`. Admission is separate from terminal completion. An immediate
`success: false` refusal causes MCP to unregister that callback and return without
waiting or dispatching a fallback operation.

`Export`, `FullExport`, and `ExportVBA` return final structured JSON.
`Build` and `MergeBuild` can return a `started: true` admission requiring a terminal
callback. Empty or unknown legacy export returns remain completion-unconfirmed.
Refusals and terminal outcomes preserve the operation ID and diagnostic fields.

The production MCP counterpart retains its existing release and compatibility-session
gates and session-envelope positions. Request-policy forwarding does not admit an
intermediate add-in through those gates. The historical-derived submission consumer
requires qualification with the exact repaired UP-004 add-in revision. Neither a
version string alone nor a successful historical setter call establishes this pair.

`RebuildAddIn(sourceFolder, callbackInfo, policy)` and fixed unattended
`BuildHeadless(sourceFolder, policy)` and `MergeHeadless(sourceFolder, policy)`
reject nonempty policies. Their recovery and compatibility contracts remain separate.

The add-in's return from `MergeBuild` is a start result, not the outcome:
the outcome arrives on the completion callback. When MCP had to start the
merge without one (callbacks unavailable, or the async start failed and MCP
fell back to a sync call), `vcs_import_objects` cannot confirm the outcome.
It returns `success: false` with `started: true` and
`completion_unconfirmed: true`. The merge may have finished either way, so
do not retry it blindly. Read `log_path`, or call `vcs_get_recent_calls()`
and `vcs_get_log(log_type="Merge")`. An immediate queue refusal is synchronous
and stops waiting. A later execution refusal arrives through that queued request's
own terminal emitter. MCP preserves either outcome without redispatch.

Whole-database export and build follow the same rule. Final export JSON
establishes completion: an add-in that advertises `request-policy/1` returns a
final structured result from `Export`, `FullExport` and `ExportVBA`. A build's
`started: true` only establishes admission. An older add-in's Empty return
establishes neither success nor failure. `vcs_export_database` and
`vcs_rebuild_database` report success only from final export JSON or a
terminal callback. Any path that has neither returns the unconfirmed-start
result above, with the operation named in `error` and
`vcs_get_log(log_type="Export")` or `"Build"`; reconcile its operation log
before retrying an uncertain call:

- the sync fallback (no callback server, an async start that raised, or a
  start result with neither `async` nor `sync`);
- an inline `{sync: true, result}` marker. MCP parses the nested `result`:
  a dispatcher refusal or a failure object with `success: false` fails and
  keeps its `error_pattern`, `decisions` and `decision_required`; final export
  JSON with `success: true` completes; Empty or anything else is unconfirmed.
  It is never run a second time.

A COM exception still fails. The integration helpers `export_source`,
`export_vba` and `build_from_source` never return `success: true` for the same
reason. `full_export=True` dispatches `FullExport` on the async, inline and
sync-fallback paths. A final-result sync API in the add-in is deferred (see
`DECISIONS.md`); until it exists, an agent that needs the outcome without the
callback server reads the log.

Operation log paths supplied by the add-in (`log_path` or `logPath`) remain
authoritative, including on refusals and unconfirmed starts. When no explicit
path is supplied, MCP can attach only the newest matching
`<source>/logs/<Export|Build|Merge>_*.log` whose modification time is at or
after the call started, and only with evidence that the operation ran.
Pre-start refusals and `started: false` never use that fallback. A decision
result keeps its own execution evidence; an unconfirmed start keeps its
uncertain verdict even when a current log is found. Root-level legacy
`Export.log` and `Build.log` are never inferred as operation logs. If no log
is attributable, `log_path` is null and `log_excerpt` is absent.

`vcs_rebuild_database` with an `output_path` calls `BuildAs(source, output)`
on every path (async, inline and sync fallback), so the add-in opens neither
the source-folder nor the save-as picker. The server version gate establishes
this release contract before target creation or callbacks; it no longer requires
an `APICapabilities` probe. Per-call acknowledgments remain mandatory.
`output_path` in the result is the path the completion callback
reports, never the request. It is None on every other result; an unconfirmed
start keeps the request as `requested_output_path`. The call is hosted in a
blank temporary database (Access exits if the add-in runs with no database
open). An uncertain result retains the original host and operation ID after the
caller returns. The Access gate stays occupied. Terminal settlement or observed
exit of the pinned process permits cleanup on the original apartment. The
response reports cleanup separately from the build outcome. Missing callbacks
and cancellation requests do not authorize closing a live build.

The host is recorded in the owned-instance registry before any build is
dispatched; if that record cannot be written, the build is refused. Registry
writers serialize through a lock file beside the registry, so concurrent
servers do not drop each other's records. When the exact process identity
cannot be pinned, nothing has been dispatched: the server requests `Quit`,
releases the target, and reports `recovery_required` because exit is
unconfirmed. A server that exits while a build is still pending leaves the
host running and its registry record in place. The next server does not
reconcile such records, so that host is never recycled automatically; close
it manually. If the registry cannot be read, a newly created window is left
unrecorded instead of rewriting the file; the server then treats that window as
the user's and never closes it.

`RunFilteredTests` is different: it runs the tests before it returns, and its
return is the final results JSON (or a refusal). `vcs_run_tests` returns a
refusal once, normalised, and does not call `RunFilteredTests` a second time.
A `runtime_error` in the results, or on the completion callback when the
run went async, is kept on the result with its `errorNumber`.

When a blocked decision and compile failure occur together (A35), the primary
result is `decision_required: true`, `error_pattern: decision_required`,
`success: false`, and `cancelled: false`. The translated decision text stays in
`error`; `run_error` retains the compile explanation and `run_error_pattern`
is `project_not_compiled`. These are runner diagnostics, separate from
`runtime_error`/`errorNumber`. Compile failure alone keeps the compile text and
`project_not_compiled` as the primary error, with `cancelled: false`.
Both transports preserve the decision journal, the operation's `log_path` and
`logPath`, results path or available inline partial results, and `results_error`.
If completion construction fails, `completion_error` and
`completion_error_number` retain that secondary fault; even passing saved
assertions cannot turn it into success. The add-in restores its owned state
before making its single terminal delivery attempt.

`vcs_run_tests` gives the same verdict whether the result came by callback,
as an inline sync marker, or through the sync fallback. Highest precedence
first: `decision_required`; a runtime error or the add-in's own
`success: false`; `cancelled`; `results_error`; then the add-in's `allPassed`,
which needs at least one passed test. An add-in without `allPassed` falls back
to nothing failed or errored and `passed > 0`. An all-EMPTY run is
`success: false` with an `error` that names the EMPTY count and no
`error_pattern`; passing tests mixed with EMPTY ones are still a success.
Results keep `decisions`, `log_path` and `logPath` when there is no
`decision_required`.

From VBA, the same switch is the optional policy argument:

```vba
VCS.MergeBuild "block"
VCS.RunFilteredTests "decline"
```

`VCS.MergeBuild` and `VCS.RunTests` with no policy stay interactive.
`VCS.RunTests` still uses silent mode internally so test code does not stop
on add-in message boxes, and it still shows the console. That is the add-in's
own behaviour for a person at the ribbon; a call that arrives through the MCP
server is automation and is headless (next section).

## Test runs are always headless

The add-in treats every API call as automation, and `ExecuteTests` forces a
headless run for that source whatever interaction mode was selected: no
console, no web runner, and `MsgBox2` prompts answered unattended. So
`vcs_run_tests(noninteractive=False)` cannot give an interactive run, and a
flag that did nothing would be a lie. It is refused before any add-in call
(no Access connection, no `SetInteractionMode`) with
`error_pattern: interactive_tests_unsupported` and `success: false`. Pass
`noninteractive=True` (the default) with a `decision_policy`, or run the tests
from the add-in's ribbon for the console. The `msaccess-vcs run-tests` CLI has
no `--interactive` flag for the same reason. Other tools' `noninteractive=False`
is unchanged. The decision is recorded as X11 in `DECISIONS.md`.

Live check (disposable blank database): with the helper missing,
the add-in's preflight installs `modTestAssert`, but under an ambient
interactive mode it then showed a "Test Helper Installed" `MsgBox2` that held the
call until it was dismissed. That is why the flag is refused rather than
documented as "headless but permissive": a visible prompt on an unattended call
holds the Access gate.

## What the add-in can prevent

These are suppressed or turned into structured results during a noninteractive
operation:

- Add-in `MsgBox2` prompts (information, warnings, Yes/No, OK/Cancel).
- The merge conflict form (`frmVCSConflict`).
- The folder picker used when no export folder is known.
- The main results window (`frmVCSMain`) staying on screen after the run.
- The extra "are you sure you want to cancel?" prompt when the window of a
  noninteractive run is closed: the add-in takes the close as the answer. A
  silent run keeps the prompt. An unattended one (automation) takes the default
  answer, yes, with no window; one a person started asks.

These are not prevented by the add-in. Use the inspector below:

- A `MsgBox` in the database's own VBA.
- Microsoft Access error and warning dialogs.
- VBA runtime dialogs (End / Debug / Help).
- VBA compile-error dialogs.
- VBA break mode. That is a paused project, not a dialog.
- Trust-center and macro-security prompts.
- "Save changes?" and other destructive confirms. They are never clicked
  automatically. When `vcs_export_object` or `vcs_import_object` has to close
  an open object and its native save prompt is answered Cancel, the call stops
  without touching the object and returns `success: false`, an `error` naming
  the object, `cancelled: true` and the log path (A30).

`ListAddinDialogs` / `DismissAddinDialog` on the add-in can see open `frmVCS*`
forms only while Access is responsive. A modal dialog blocks that call.

## List, dismiss, and check readiness

These tools do not use the Access COM gate. Call them while another tool is
still waiting.

The following are separate examples, not a sequence to execute. Use dialog IDs,
PID, and creation time returned by inspection; the numbers are placeholders.
Apply the authorization and outcome rules below before dismissing or cancelling.

```python
vcs_list_dialogs(r"C:\db.accdb")
vcs_list_dialogs(r"C:\db.accdb", pid=12345, create_time=133000000000000000)

vcs_dismiss_dialog(r"C:\db.accdb", "hwnd:100", button="OK")
vcs_dismiss_dialog(r"C:\db.accdb", "hwnd:100", button="End")
vcs_dismiss_dialog(r"C:\db.accdb", "hwnd:200", action="close")
vcs_dismiss_dialog(r"C:\db.accdb", "hwnd:200", action="cancel")

vcs_recover_dialogs(r"C:\db.accdb", policy="report")
vcs_recover_dialogs(r"C:\db.accdb", policy="safe")
vcs_recover_dialogs(r"C:\db.accdb", policy="end_runtime_error")

vcs_automation_status(r"C:\db.accdb")
```

Target selection reports `matched_by: owned_registry`, `window_title`, or `pid`.
A title is a selection hint, not proof of the open database. If title matches
conflict with a creation-time-confirmed owned instance, selection returns
`target_identity_conflict` without clicking or closing anything. Multiple
confirmed owned hosts remain ambiguous even when only one title matches.
Pass the intended `pid` and `create_time` explicitly to resolve ambiguity.
Process identity and the window are checked again immediately before action.

`action=close` closes a finished add-in window and does not cancel a running
operation. If the gate is busy with an operation on the same database
(compared after path normalisation), close is refused with
`operation_in_progress`. An operation on another database does not block it.
Closing the window of a noninteractive run cancels that run without any
confirmation prompt. `action=cancel` posts a close to the add-in window as a
cancel request and reports `cancel_requested: true`, not `interrupted`. While
an operation runs the add-in keeps its window open. An interactive run
(`noninteractive=False` on the import and export tools, and `vcs_export_database`) also asks
"Cancel Current Operation?": Yes stops it, No resumes it. The window staying
open can make the dismissal `dismiss_uncertain` even though the close was
posted. The waiting call's own result says how the request ended (see below).

Button names ignore the Win32 accelerator marker, so `End` matches a button caption of `&End`. `button="Debug"` is refused. `button="End"` stops the failed VBA call. Continue is never clicked automatically. The
result sets `failure_dialog_dismissed` or `interrupted`. The waiting tool
still returns its own error or timeout. Do not retry the mutation until
`vcs_automation_status` reports `ready: true`.

`ready` means every one of these is confirmed: the process is Access with a
readable creation time, it is running, it answered a window message (an
unknown answer is not ready), VBA is not in break mode, no blocking dialog is
open, and the Access gate is not busy with that database. When it is false,
`error_pattern` names the first missing condition. `access_not_running` (relaunch)
and `no_windows_to_probe` (wait) are different outcomes; the others are
`identity_unconfirmed`, `access_unresponsive`, `vba_break`, `blocking_dialog`
and `server_busy`.
`execution_interrupted` carries the dialog text captured before End. A cancel
request is never shown there; the call it was made against reports it.

Interruption records are keyed by process identity (PID plus creation time), so
a new process that reuses a PID starts clean. Dismissing a runtime or compile
error records one, from `vcs_dismiss_dialog` or `vcs_recover_dialogs` alike, and
so does clicking End on any dialog. Nothing else does: an `access_dialog` whose
text merely mentions an error is not a failure, and neither tool records it or
sets `failure_dialog_dismissed` for it.
The record is reserved before the click (or the cancel's close) is sent,
because ending the dialog lets Access resume and the blocked call can return
before the click does. A reservation made while a gated call on that database
was in flight belongs to that call. When the handler returns, it waits for a
click still being delivered, at most twice the dialog timeout plus a second,
and then uses up its records:
- Delivered: the result is forced to `success: false` and
  `execution_interrupted: true`. The primary `error_pattern` is
  `decision_required` if the result already had it (by `decision_required:
  true` or by pattern), otherwise `execution_interrupted`.
- Not sent (the button no longer belonged to the process, the window did not
  answer, or the button refused the action): the reservation is dropped and the
  result is unchanged. Nothing reached the dialog, so it did not release the
  call.
- Delivery uncertain (the button message timed out or failed part way), or the
  click did not report back by the deadline: the result is `success: false` with
  `interruption_uncertain: true`. A result that was a success gets
  `error_pattern: interruption_uncertain`; a result that already failed keeps
  its own pattern. It is not `execution_interrupted`, because the End may not
  have arrived. A click that reports back after the call finished finds nothing
  to settle, so a later call never inherits it.
- A cancel close that was posted is a request, not a confirmed interruption:
  the person may answer No. The call's result decides it:
  - `cancelled: true` (the add-in confirmed the cancel, in its sync result or
    its `cancelled` terminal callback): `success: false` and
    `execution_interrupted: true` with the same precedence as above. `cancelled`
    and the add-in's error text are kept.
  - A success: the result is unchanged and gains `cancel_not_honored: true`. The
    operation completed.
  - Any other failure: `interruption_uncertain: true` is added and the result
    keeps its own `error_pattern` (`decision_required`,
    `completion_unconfirmed`, or the error's own).
  A whole export (`vcs_export_database`) or a build that falls back to the sync
  API carries no `cancelled`. Its result is `completion_unconfirmed`, so a
  cancel there shows up only as `interruption_uncertain`. Single-object import
  and export report `cancelled: true` for a confirmed cancel. A delivered End
  on the same call still makes it `execution_interrupted`. A cancel posted with
  no call running on that database records nothing.

Decisions, transfer fields, `runtime_error`, and the original error text
are kept. This happens before the usage log is written, so the `tool_call`
entry, and `vcs_get_recent_calls`, record the same `success`, `error_pattern`,
`execution_interrupted` and `interruption_uncertain` the client got. A
reservation made after the handler returned did not interrupt it and is dropped
when the gate is released. A delivered click made with the gate free is shown as
`last_interruption` by the status tools until the process identity changes or
the next gated call on that database starts; an undelivered one records
nothing.

A dialog gets a known kind only from a positive signature (a Microsoft
Access or Visual Basic caption, run-time or compile error text, the End/Debug
button set, the single-OK-button rule below, or, for a window that is not a
standard dialog box, the add-in caption). A standard
dialog box with a custom caption is `kind: "vba_msgbox"` when its only
actionable button is OK (a Help button does not count); this is the shape of a
VBA `MsgBox "text", vbOKOnly, "Caption"`. Any other standard dialog that
matches no signature, including a custom-caption box with two or more buttons
or a single non-OK button, is `kind: "unknown"`. Both are reported with title,
text and buttons and are blocking.

A standard dialog box is a Win32 `#32770` window or an Office NetUI
`NUIDialog`. Access draws a `MsgBox` with the `@`-separated bold form (which is
what the add-in's `MsgBox2` shows when a person is watching) and its own error
dialogs as `NUIDialog`. That window has no Win32 buttons, so the inspector reads
its text and buttons through Microsoft Active Accessibility and presses a button
with its default action, after checking again that the same button is at the
same place. Only `NUIDialog` windows of an Access process are read. Both classes
get the same classification rules.

The window class is tested before the add-in caption. A `MsgBox2` box that
carries an add-in caption ("Version Control System", or a caption starting
`MSAccessVCS`) is still a `#32770` or `NUIDialog` window, so it is classified
by its buttons like any other dialog: OK-only is `vba_msgbox` (blocking,
`ready: false`, pressed by `safe` unless the text is destructive) and a
multi-button box is `unknown` (blocking, report-only). Only a window of another
class with an add-in caption, the add-in's main or progress form, is
`kind: "addin_window"` (`is_dialog: false`, never blocking).

`policy="safe"` clicks OK only on a `vba_msgbox` (an OK-only standard dialog)
whose text is not a save, discard, delete, or overwrite confirmation. Access
error and warning dialogs (`access_dialog`) are report-only, and compile errors
belong to `end_runtime_error`, so `safe` leaves both open. A single-button
`vba_msgbox` is closed because Access cannot continue until it is, and the
closed dialog is listed in `closed` with its `kind`, `title` and `message`. A
`vba_msgbox` with destructive words stays open and is listed in `skipped`.
`unknown` is never clicked automatically; dismiss it explicitly with
`vcs_dismiss_dialog(..., button=...)`. Everything skipped is returned in
`skipped` with its buttons. Debug, save and discard are never clicked, and the
click is a button message (or, on a `NUIDialog`, the button's accessible
default action), never a keystroke or coordinate click.

`vcs_dismiss_dialog` and `vcs_recover_dialogs` wait the same way after a
click: they poll until the dialogs they clicked have closed (dialogs the call
did not act on never extend the wait) or the timeout elapses, then inspect
once for the report. A dialog they clicked that is still open at the deadline
is `dismiss_uncertain`, for both tools. `end_runtime_error` does everything
`safe` does, and adds End on a runtime-error dialog and OK on an OK-only compile
error. It still never clicks Debug, and neither policy clicks an
`access_dialog`; an explicit `vcs_dismiss_dialog(..., button=...)` is
unchanged. The destructive-text check guards OK clicks
only: End on a runtime-error dialog does not depend on it, so a runtime error
whose text mentions "deleted" is still ended, and `safe` clicks nothing on it.

If two Access windows match the database, the tools return
`ambiguous_instance` and click nothing. Pass `pid`. If `create_time` does not
match that process, the tools return `process_identity_mismatch` and click
nothing.

An ownership-registry record selects a target only when its database path matches
and its stored creation time is readable and equals the Access process's observed
creation time. Records retained after a failed process query are bookkeeping,
not proof of identity. A missing, unreadable or different stamp cannot select a
registry-only target (`access_not_found`); title matching and explicit
PID/creation-time selection still follow their own identity checks.

Right before each click or close, the tools list windows again. The dialog
must still belong to the same process (PID and creation time), and its class,
title, full text, kind and buttons must match what the call inspected. Windows
reuses the handle of a closed box, so a different box can appear behind the
same handle and button handles. Any difference returns `dialog_changed` with
`success: false`, and nothing is clicked. `vcs_recover_dialogs` also runs its
policy again on the fresh window and clicks the button read from it, so `safe`
does not acknowledge a box whose text has become a delete or discard
confirmation by the time of that check. It still handles the other dialogs, and
lists the changed one in `skipped` with `reason: "dialog_changed"`. `closed`
always describes the box that was clicked.

## Limits

- Works on the interactive desktop session of the user running the MCP server.
  A dialog on another session, Secure Desktop, or an elevated window the
  server cannot query is reported as not found or unresponsive.
- 32-bit and 64-bit Access are both addressed by window messages. The server
  process does not need to match the Access bitness for `BM_CLICK`.
- A hung process (`access_unresponsive`) is not clicked. That is different
  from a modal dialog, which still answers window messages.
- Break mode has no button. Report it and reset it from the VBE, or stop the
  server-owned Access process with the existing ownership rules. Do not send
  keys.
- The check before a click narrows the gap between inspection and click, but
  cannot close it from outside Access. A box that is replaced after the last
  listing and before the button message arrives is still clicked.
  `vcs_dismiss_dialog` compares against its own inspection at the start of the
  call, not the `vcs_list_dialogs` result you read earlier; check `closed` to
  see which box was clicked.
- The default dialog wait is 5 seconds (`ACCESS_VCS_DIALOG_TIMEOUT_SEC`,
  capped at 30). If the window is still open, the result is
  `dismiss_uncertain` and the mutation is not retried.
- Each dialog tool call has one response deadline: the dialog wait plus 5
  seconds, covering the wait for a worker thread and the call itself. A call
  still running at the deadline returns `tool_timeout`; its thread cannot be
  stopped and runs on until Access answers. Each tool keeps at most two such
  threads and lets at most two more calls wait for one. A call that gets no
  thread by its deadline, or finds two already waiting, returns
  `worker_capacity_unavailable` without touching a window. Retry once Access
  answers. These threads never use capacity the Access gate needs, so gated
  tools keep answering, or return `server_busy`, on time.
- A gated tool does all its Access work on the gate's COM apartment thread,
  including connecting, loading the add-in and synchronous add-in calls. A
  dialog or status call therefore starts and returns on time while a gated
  call is blocked in any of those steps.
- Cancelling or timing out the caller does not free the Access gate. The same
  worker and call identity own it until the body, connection cleanup and
  apartment-loop shutdown finish. Other gated requests get `server_busy`
  within `ACCESS_VCS_BUSY_WAIT_SEC`; dialog, status and cancel tools remain
  callable. Async bodies receive cancellation at their next await, after any
  blocking COM step returns. Sync bodies continue to completion. Interruption
  records stay attached to that original call and are consumed before the
  worker releases its slot, including when no caller is still waiting.
- `vcs_rebuild_database` uses an isolated `DispatchEx` host. Cleanup closes only
  that host; an already-open output is refused without closing its holder.
  `tests/test_build_host_ownership.py` covers this boundary.

## Production request-policy migration gate

The production consumer keeps its existing SemVer range and
`msaccess-vcs.session/1` handshake. It additionally requires the admitted
`APIIdentity` response to contain
`request_policy_contract: "msaccess-vcs.request-policy/1"` before sending a
request-policy argument. This metadata is retained with that accepted session,
so the guard issues no additional native probe. Missing or unfamiliar values
return `request_policy_unsupported`, with `started: false`, before dispatch.
Legacy calls that omit the additional argument remain usable through normal
admission. Removing policy to bypass a refused conflict remains unsupported recovery.

The current unmodified production dev.18 add-in lacks this marker and revised
argument contract. It is explicitly refused for request-policy calls. It is not
qualified with the intermediate UP-004 add-in, which lacks mutual sessions.
UP-009 must integrate the request-policy methods before advertising the marker.
Its exact dispatch signatures must be:

```text
APIExecute(envelope, method, arg1, arg2, arg3)
APIExecuteAsync(envelope, callbackInfo, method, arg1, arg2, arg3)
API(method, arg1, arg2, arg3, strSessionEnvelope)
APIAsync(callbackInfo, method, arg1, arg2, arg3, strSessionEnvelope)
```

The session wrapper must pass `strSessionEnvelope` by name into raw dispatch.
A policy must never occupy the envelope slot. The session protocol declaration,
version checks, and identity validation remain separate admission requirements.

## Synchronous MCP identities

The repaired historical consumer calls
`APIMCP(operation_id, method, arg1, arg2, arg3)` for synchronous operation methods.
It generates a fresh UUID per request, independently of callback registration.
The add-in captures MCP as the initiating source and retains that public ID through
completion. External callers can continue to use `API` with add-in-generated IDs.
Metadata and inspection calls do not start an operation and retain their existing API.
The MCP never replays a synchronous operation after a COM exception. That failure
retains its ID and `operation_outcome_unconfirmed` for reconciliation.

The production counterpart requires its admitted request-policy marker before
using `APIExecuteMCP(envelope, operation_id, method, arg1, arg2, arg3)`. UP-009 must
implement that session-enforcing wrapper along with `APIMCP` before advertising
the marker. Old admitted production libraries retain their existing envelope-first
legacy calls when no new argument or result contract is needed. They do not establish
the repaired pair's MCP-origin guarantee and remain outside its qualification.
