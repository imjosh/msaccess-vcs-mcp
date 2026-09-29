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
```

`noninteractive` defaults to true on both tools. `decision_policy` defaults
to `block`.

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

The add-in restores its interaction mode when the operation
finishes, fails, or is cancelled. A blocked prompt is `success: false` with
`error_pattern: decision_required` and a `decisions` array. It is not a
successful merge or a successful test run.

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
- The extra "are you sure you want to cancel?" prompt if that window is
  closed during a silent or noninteractive run.

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

A dialog gets a known kind only from a positive signature (a Microsoft
Access or Visual Basic caption, run-time or compile error text, the End/Debug
button set, the add-in caption). A standard dialog box that matches none is
`kind: "unknown"`, reported with its title, text and buttons. `unknown` is
blocking and is never clicked automatically; dismiss it explicitly with
`vcs_dismiss_dialog(..., button=...)`.

`policy="safe"` clicks OK only on a known-kind OK-only dialog whose text is
not a save, discard, delete, or overwrite confirmation. Everything else is
returned in `skipped` with its buttons.

`vcs_dismiss_dialog` and `vcs_recover_dialogs` wait the same way after a
click: they poll until the dialogs they clicked have closed (dialogs the call
did not act on never extend the wait) or the timeout elapses, then inspect
once for the report. A dialog they clicked that is still open at the deadline
is `dismiss_uncertain`, for both tools. `end_runtime_error` adds End on a
runtime-error dialog and still never clicks Debug.

If two Access windows match the database, the tools return
`ambiguous_instance` and click nothing. Pass `pid`. If `create_time` does not
match that process, the tools return `process_identity_mismatch` and click
nothing.

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
- The default dialog wait is 5 seconds (`ACCESS_VCS_DIALOG_TIMEOUT_SEC`,
  capped at 30). If the window is still open, the result is
  `dismiss_uncertain` and the mutation is not retried.
