# Dialogs and noninteractive Access automation

> Release compatibility policy (owner decision, 2026-10-02): every add-in release changes its version. The supported release version defines the API contract; capability probing is not required to establish release compatibility. The spec assumes the server checks the installed add-in version and refuses unsupported releases before starting operations. The minimum supported release version must be stated when the release is assigned; do not infer it from a development rebuild. Per-call mode and policy acknowledgments still confirm the requested state and remain required. This policy supersedes earlier statements requiring capability checks instead of a version gate. It is a specification change, not evidence that version enforcement is already implemented.

X16 implements the version gate before add-in-dependent tool bodies. See
[release compatibility](RELEASE_COMPATIBILITY.md) for assigned unpublished ranges,
development identities, stable errors and recovery exceptions.

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
defaults to `block`. `vcs_run_tests` accepts only true: an automation test run
is always headless (see "Test runs are always headless" below). On `vcs_import_object` and `vcs_export_object` an
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

Pass `noninteractive=False` to `vcs_import_objects`, `vcs_import_object` and
`vcs_export_object` to select interactive mode explicitly and show the normal
prompts. MCP sends the mode before the operation starts, so the run does not
depend on the mode the add-in was last left in. All three tools
require `SetInteractionMode(0)` to confirm `success: true` and numeric
`effective_mode: 0` before dispatching any operation or registering its callback.
The A24 add-in contract returns a JSON string:

```json
{"success":true,"requested_mode":0,"effective_mode":0}
{"success":false,"requested_mode":0,"effective_mode":2,"error_pattern":"interaction_mode_refused","error":"..."}
```

An enclosing caller-owned noninteractive scope (including one left open after
`policy_cleanup_error`) refuses interactive selection with
`interaction_mode_refused`. MCP returns that refusal unchanged and starts
nothing. It never clears or weakens the enclosing policy; its owner must clear
it with `ClearOperationPolicy`, or finish an active root, before retrying.

Explicit interactive requests require an **A24 build or later**. Earlier builds
return VBA `Empty` (Python `None`), which cannot confirm acceptance. Empty,
missing, malformed, or otherwise unconfirmed responses return
`success: false, error_pattern: interaction_mode_unconfirmed` without starting
work. Upgrade the add-in for unsupported responses. These checks confirm the
requested interaction state; they do not replace the release-version contract.
The add-in documents the same contract in
[`docs/noninteractive-dialogs.md`](../../msaccess-vcs-addin/docs/noninteractive-dialogs.md).

MCP sets the policy before the call and clears it in `finally`, so the
add-in's interaction mode is restored whether the operation finishes, fails,
or is cancelled. `vcs_run_tests` and a full `vcs_import_objects` merge pass the
policy inline to `RunFilteredTests` and `MergeBuild`. A category-scoped
`vcs_import_objects`, `vcs_import_object` and `vcs_export_object` set it with
`SetOperationPolicy`. A session policy set that way is caller-owned: the
add-in's `Finish` does not close it, and `ClearOperationPolicy` does. Clearing
is idempotent. A clear counts only on `{success: true}`; Empty or malformed is
a failure too. A failed clear is attached as `policy_cleanup_error` (and
written to the usage log as `policy_cleanup_failed`); it never replaces the
operation's own result.

MCP dispatches a noninteractive operation only after `SetOperationPolicy`
returns `{success: true, policy: <the requested policy>}`. An add-in refusal is
the tool result unchanged; Empty, malformed JSON, a missing or different
`policy` is `error_pattern: policy_unconfirmed`. Nothing starts and no clear is
sent. This acknowledgment confirms the requested policy. A policy method
that throws is a separate case and not handled.

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
| `operation_already_running` | Another add-in operation is running. Setting a policy or starting a merge is refused and nothing changed. A normal refusal, not a cleanup failure. The add-in's dispatcher refusal of a reentrant call (`VCS_API_REFUSED: ...` from `API`, or the marked `{success: false}` envelope from `APIAsync`) uses this pattern too, with `api_refused: true`, on every hard-coded call; no dependent call follows. |
| `api_self_dispatch` | The dispatcher refused a call that arrived back in the project that sent it. This is an add-in defect, so waiting and retrying cannot help. Also carries `api_refused: true`. |
| `invalid_addin_response` | A JSON-contract method (`ExportByType`, `ImportByType`, `ExportObject`, `ImportObject`, `ExecuteSQL`, `RunVBA`, `SetOption`) returned something that is not JSON. Never a success. Raw-contract calls (`GetOption` values, `GetLogContent` text, `vcs_call_vba`) still wrap a plain value as success. |
| `merge_not_available` | The database has no merge to run (for example a blank database). It is not retried. Run a full build (`vcs_rebuild_database`). |
| `decision_required` | A prompt or merge conflict the policy did not cover. Carries `decisions`. |
| `interaction_mode_refused` | The add-in could not make interactive mode effective. Its enclosing scope or active operation must be released by its owner; nothing starts. |
| `policy_unconfirmed` | MCP could not confirm `SetOperationPolicy` (Empty, malformed, no echoed `policy`, or a different policy). Nothing starts and no clear is sent; requires an add-in whose `SetOperationPolicy` returns `{success: true, policy}`. |
| `interaction_mode_unconfirmed` | MCP could not confirm interactive mode, including VBA Empty from an older add-in. Nothing starts; unsupported responses require an A24 or later build. |
| `interactive_tests_unsupported` | `vcs_run_tests(noninteractive=False)`. An automation test run is always headless. Refused by MCP before any add-in call; nothing started. |
| `version_incompatible` | The installed add-in is outside the server range or is an unadmitted prerelease. Carries installed version, required range, component and recovery action. |
| `version_unconfirmed` | Installed version is unknown, invalid or unreadable. No dependent work starts; metadata and Win32 recovery remain available. |
| `invalid_build_path` | `vcs_rebuild_database` was given a relative `output_path` (refused by MCP), or the add-in rejected the pair: a source folder without `vcs-options.json`, an output with no folder or extension, a missing output folder, or the add-in itself. Nothing started. |
| `export_folder_mismatch` | `vcs_export_database` was given an `output_dir` that is not the add-in's configured export folder. Refused by MCP before anything is exported; the result carries `configured_export_folder` and `requested_output_dir`, and `export_path` is `None`. |
| `export_folder_unavailable` | `vcs_export_database` could not read the add-in's export folder (`GetExportFolder` raised, was refused, or returned nothing or JSON). Nothing is exported. |

The add-in's return from `MergeBuild` is a start result, not the outcome:
the outcome arrives on the completion callback. When MCP had to start the
merge without one (callbacks unavailable, or the async start failed and MCP
fell back to a sync call), `vcs_import_objects` cannot confirm the outcome.
It returns `success: false` with `started: true` and
`completion_unconfirmed: true`. The merge may have finished either way, so
do not retry it blindly. Read `log_path`, or call `vcs_get_recent_calls()`
and `vcs_get_log(log_type="Merge")`. A refusal arrives both as the sync return
and on the callback; MCP returns it once.

Whole-database export and build follow the same rule. `Export`, `FullExport`,
`ExportVBA` and `Build` are Subs or form starts, so their API return is Empty
and says nothing about the outcome. `vcs_export_database` and
`vcs_rebuild_database` report success only from a terminal callback. Any path
that has none returns the unconfirmed-start result above, with the operation
named in `error` and `vcs_get_log(log_type="Export")` or `"Build"`:

- the sync fallback (no callback server, an async start that raised, or a
  start result with neither `async` nor `sync`);
- an inline `{sync: true, result}` marker. MCP parses the nested `result`:
  a dispatcher refusal or a failure object with `success: false` fails and
  keeps its `error_pattern`, `decisions` and `decision_required`; Empty or
  anything else is unconfirmed. It is never run a second time.

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
open), which is closed and removed when the call returns.

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

Decisions, `runtime_error`, `policy_cleanup_error` and the original error text
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
- Known gap (M43): `vcs_rebuild_database` attaches to an Access instance that
  is already running instead of starting its own. The build then fails ("You
  already have the database open"), and its cleanup closes that instance's
  database and quits it. Close other Access windows before a rebuild.
