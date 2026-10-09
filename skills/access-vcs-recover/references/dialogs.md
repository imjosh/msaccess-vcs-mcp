# Inspect dialogs and cancellation

## Bind the intended instance

Use `vcs_list_dialogs`, or `vcs_recover_dialogs(policy="report")`, to inspect without
clicking. These and `vcs_automation_status` remain callable while the COM gate is
blocked. Use the exact target database path. If multiple instances match, use a
PID and creation time from returned identity evidence; never guess a PID or
infer ownership from its number. Ambiguous or mismatched identity means no click.

Read title, text, kind, and buttons, and distinguish a blocking dialog from a
finished add-in results window. VBA break mode is a paused project, not a dialog;
ask the user to reset/resume appropriately in the VBE. A hung/unresponsive process
cannot be fixed by clicking. Resolve identity/readiness failures before acting.

## Choose an authorized action

| Situation | Supported action |
| --- | --- |
| Informational OK-only VBA message | `vcs_recover_dialogs(policy="safe")` can acknowledge nondestructive messages when continuing the authorized workflow is appropriate. |
| Runtime error with End/Debug | If ending this failed execution is authorized, use `end_runtime_error` recovery or dismiss the inspected dialog with End. Capture diagnostics first. |
| Compile error | `end_runtime_error` can acknowledge it, but that does not repair compilation. Follow the project's compile-error reference/tool guidance. |
| Access warning/error or unknown dialog | Read and report it; use explicit dismissal only when the intended button/action is understood and authorized. |
| Finished add-in results window | `vcs_dismiss_dialog(action="close")`; close is refused for a running operation on that database. |
| Running add-in progress window to stop | `action="cancel"` requests cancellation; inspect the original call's eventual outcome. |

`safe` does not click Access warnings/errors, unknown boxes, or destructive
save/discard/delete/overwrite prompts. `end_runtime_error` includes safe recovery
and additionally End on runtime errors and OK on compile errors; it can interrupt
code, even if the error text contains destructive words. It is not a harmless
acknowledgment of every dialog. Debug is refused; automatic Continue is unsupported.
For a choice outside existing authorization, present the specific text and effect
to the user before clicking. Do not bypass tool limits with keys or coordinates.

Use a freshly inspected `dialog_id` and the same process identity. Tools re-read
the window before acting, but a handle can be reused and a last-moment replacement
remains possible. On `dialog_changed`, inspect anew. Read `closed`/`skipped` to
see what was actually acted on. `dismiss_uncertain` or `tool_timeout` does not
prove a click failed; inspect again and preserve `interruption_uncertain`.

## Cancellation is a request

For an async operation with a known operation ID, `vcs_cancel_operation` records a
request that the add-in polls at a safe point. `success: true` there means only
`cancel_requested: true`. An unknown/already-finished response does not prove the
operation was stopped. Use the original call, recent calls, and operation log.

Dialog `action="cancel"` posts a close request to the running add-in window.
An interactive run asks for confirmation; No resumes it. A window can remain
open and return `dismiss_uncertain` even though the request was posted.

Only `cancelled: true` in the original operation result confirms cancellation.
`cancel_not_honored: true` means it ended without honoring the request; report the
actual outcome. A dialog cancel followed by another failure may carry
`interruption_uncertain`, while a confirmed cancel carries `execution_interrupted`.
Status inspection does not report a cancel request as a confirmed interruption.
Ending an error dialog can mark `execution_interrupted` even if the blocked call
returns before the click finishes; readiness afterwards does not turn it into
success. Preserve decisions, runtime errors, cleanup failures, and partial writes.

Recheck `vcs_automation_status` after UI recovery and reconcile the operation's
effects before retrying. If dialogs remain, process identity cannot be confirmed,
or the prior outcome stays unknown, stop mutations and request the specific user
action. Leave user-owned Access windows and unsaved work intact.
