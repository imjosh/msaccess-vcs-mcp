# Dialogs and noninteractive Access automation

Agent calls default to a scoped noninteractive mode. Interactive ribbon and
Immediate Window use is unchanged unless a decision policy is passed.

`DoCmd.SetWarnings` is not a dialog suppressor. It does not cover `MsgBox`,
the conflict form, VBA runtime dialogs, or compile errors.

## Run tests and merge without a dialog

```python
vcs_run_tests(r"C:\db.accdb", filter="modTestEncoding")
vcs_run_tests(r"C:\db.accdb", noninteractive=True, decision_policy="block")

vcs_import_objects(r"C:\db.accdb", r"C:\db.src")
vcs_import_objects(
    r"C:\db.accdb",
    r"C:\db.src",
    decision_policy="prefer_source",
)

vcs_import_object(r"C:\db.accdb", "module", "modHelpers")
vcs_export_object(r"C:\db.accdb", "form", "frmMain")
```

`noninteractive` defaults to true on all four tools. `decision_policy`
defaults to `block`. On `vcs_import_object` and `vcs_export_object` an
add-in error that would have been a message box (for example "Merging not
supported for add-in forms") comes back as `success: false` with the
message in `error` and the run's `log_path`.

| Policy | Confirmations | Merge conflicts |
| --- | --- | --- |
| `block` | Return `decision_required`. Do not approve. | Same. The database is not overwritten. |
| `decline` | Answer No, Cancel, or Abort. | Keep the database object. |
| `prefer_source` | Still `decision_required` (this policy is for conflicts). | Each conflict takes the action it asks for; source overwrites when it asks for none. |
| `prefer_database` | Still `decision_required`. | Keep the database object. Same as `skip`. |
| `skip` | Still `decision_required`. | Keep the database object and skip the source file. |

Pass `noninteractive=False` to select interactive mode explicitly and show
the normal prompts. MCP sends the mode before the operation starts, so the
run does not depend on the mode the add-in was last left in.

MCP sets the policy before the call and clears it in `finally`, so the
add-in's interaction mode is restored whether the operation finishes, fails,
or is cancelled. `vcs_run_tests` and a full `vcs_import_objects` merge pass the
policy inline to `RunFilteredTests` and `MergeBuild`. A category-scoped
`vcs_import_objects`, `vcs_import_object` and `vcs_export_object` set it with
`SetOperationPolicy`. A session policy set that way is caller-owned: the
add-in's `Finish` does not close it, and `ClearOperationPolicy` does. Clearing
is idempotent. A failed clear is attached as `policy_cleanup_error` (and
written to the usage log as `policy_cleanup_failed`); it never replaces the
operation's own result.

If the add-in refuses the interaction-mode request (for example while an
enclosing noninteractive scope is open), that refusal is the tool result, with
the add-in's `error`, and the operation is not started.

A blocked prompt is `success: false` with `error_pattern: decision_required`
and a `decisions` array. It is not a successful merge or a successful test run.

MCP passes these add-in refusals through unchanged, with `success: false` and
an `error`:

| `error_pattern` | Meaning |
| --- | --- |
| `invalid_decision_policy` | Unknown policy name. The message lists the valid names. MCP also refuses this before calling Access. |
| `operation_already_running` | Another add-in operation is running. Setting a policy or starting a merge is refused and nothing changed. A normal refusal, not a cleanup failure. |
| `merge_not_available` | The database has no merge to run (for example a blank database). It is not retried. Run a full build (`vcs_rebuild_database`). |
| `decision_required` | A prompt or merge conflict the policy did not cover. Carries `decisions`. |

The add-in's return from `MergeBuild` is a start result, not the outcome:
the outcome arrives on the completion callback. When MCP had to start the
merge without one (callbacks unavailable, or the async start failed and MCP
fell back to a sync call), `vcs_import_objects` cannot confirm the outcome.
It returns `success: false` with `started: true` and
`completion_unconfirmed: true`. The merge may have finished either way, so
do not retry it blindly. Read `log_path`, or call `vcs_get_recent_calls()`
and `vcs_get_log(log_type="Merge")`. A refusal arrives both as the sync return
and on the callback; MCP returns it once.

`RunFilteredTests` is different: it runs the tests before it returns, and its
return is the final results JSON (or a refusal). `vcs_run_tests` returns a
refusal once, normalised, and does not call `RunFilteredTests` a second time.
A `runtime_error` in the results, or on the completion callback when the
run went async, is kept on the result with its `errorNumber`.

From VBA, the same switch is the optional policy argument:

```vba
VCS.MergeBuild "block"
VCS.RunFilteredTests "decline"
```

`VCS.MergeBuild` and `VCS.RunTests` with no policy stay interactive.
`VCS.RunTests` still uses silent mode internally so test code does not stop
on add-in message boxes, and it still shows the console.

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
  automatically.

`ListAddinDialogs` / `DismissAddinDialog` on the add-in can see open `frmVCS*`
forms only while Access is responsive. A modal dialog blocks that call.

## List, dismiss, and check readiness

These tools do not use the Access COM gate. Call them while another tool is
still waiting.

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

`action=close` closes a finished add-in window and does not cancel a running
operation. If the gate is busy with an operation on the same database
(compared after path normalisation), close is refused with
`operation_in_progress`. An operation on another database does not block it.
Closing the window of a noninteractive run cancels that run without any
confirmation prompt. `action=cancel` asks the add-in window to stop the
operation and reports `interrupted: true`.

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
`execution_interrupted` carries the dialog text captured before End or cancel.

Interruption records are keyed by process identity (PID plus creation time), so
a new process that reuses a PID starts clean. Dismissing a runtime or compile
error records one, from `vcs_dismiss_dialog` or `vcs_recover_dialogs` alike, and
so does clicking End on any dialog. Nothing else does: an `access_dialog` whose
text merely mentions an error is not a failure, and neither tool records it or
sets `failure_dialog_dismissed` for it.
A record made while a gated call on that database was in flight belongs to that
call: when the handler returns, its result is forced to `success: false` and
`execution_interrupted: true`, and the record is removed. The primary
`error_pattern` is `decision_required` if the result already had it (by
`decision_required: true` or by pattern), otherwise `execution_interrupted`.
Decisions, `runtime_error`, `policy_cleanup_error` and the original error text
are kept. This happens before the usage log is written, so the `tool_call`
entry, and `vcs_get_recent_calls`, record the same `success`, `error_pattern`
and `execution_interrupted` the client got. A record attached after the handler
returned did not interrupt it and is dropped when the gate is released. A record
made with the gate free is shown as `last_interruption` by the status tools
until the process identity changes or the next gated call on that database
starts.

A dialog gets a known kind only from a positive signature (a Microsoft
Access or Visual Basic caption, run-time or compile error text, the End/Debug
button set, the add-in caption, or the single-OK-button rule below). A standard
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
