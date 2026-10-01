<!-- BEGIN HEADER -->
# Decision Log

A reverse-chronological journal of architectural and strategic decisions.
Maintained by AI coding agents (and human developers) at the end of working
sessions. Each entry captures what was decided, what alternatives were
considered, and why — so future contributors never revisit dead ends or lose
context on trade-offs already evaluated.

Agents: read this file before working on any module referenced here.

### When to log

Log decisions that constrain future design, involved genuine alternatives,
or would be non-obvious to a future contributor. A good litmus test: does
the "What this rules out" section have something meaningful to say?

Do NOT log: bug fixes with obvious solutions, test-only refactors,
documentation updates, or minor config tweaks that don't affect
architecture.

### Entry format

Insert new entries directly below this header, newest first. Do not modify
or reorder existing entries except to add supersession notes (see below).
If a session produced multiple independent decisions, create a separate
entry for each.

```
---

## YYYY-MM-DD — [Short descriptive title]

**Trigger**: What problem, requirement, or situation prompted this work.

**Options explored**:
- For each option, name the approach, its strengths, and why it was or
  wasn't chosen. Include options that were tried and reverted.

**Decision**: What was chosen and the core trade-off.

**What this rules out**: Future directions now constrained or foreclosed.
What would trigger revisiting this decision.

**Relevant files**: Key files created or modified.
```

### Guidelines

- Focus on **why**, not what. The diff shows what changed; this log
  explains the reasoning.
- Capture rejected alternatives with equal care. Future agents need to
  know what was already tried.
- Be specific — name libraries, files, config choices, error messages.
- Aim for 10–50 lines per entry. Reference document, not narrative.
- Plain language. No jargon, no editorializing, no padding.

### Superseded entries

When a new decision invalidates, corrects, or replaces guidance in an older
entry, add a blockquote annotation to the affected older entry — do not
rewrite or delete its original text. Place the note immediately after the
entry's heading or after the paragraph containing the superseded claim.

> **⚠ Superseded** (YYYY-MM-DD): [Brief explanation of what changed and
> why.] See "[title of newer entry]" above.

Use **⚠ Partially superseded** when only specific claims are affected, and
**⚠ Superseded** when the entire entry's premise or decision has been
overturned. Always scan older entries for claims that conflict with the new
decision — agents reading the log linearly will otherwise encounter
contradictory guidance.
<!-- END HEADER -->

---

## 2026-09-30 — `vcs_cancel_operation` records a request; `cancelled` means the add-in confirmed it

**Trigger**: M37 (interface review F7, MCP half). The tool set a flag and returned `success: true, "Operation will stop at next safe point"`, yet the add-in did not poll the flag (A28 wires that) and a run that finished anyway returned its normal result with nothing saying a cancel had been asked for. The acknowledgment read as a stop.

**Options explored**:
- Keep the wording and only fix the docs. Rejected: the result is what agents read, and it still claimed a stop.
- Make the tool itself return `cancelled: true` or report a failed cancel after a wait. Rejected: the tool cannot know. Only the add-in's terminal callback does, and the tool must stay callable while Access is blocked.
- Restore the commented direct `Cancel` COM call beside the poll. Rejected: it takes the Access gate's COM apartment, can block on a busy instance, and gives a second channel to keep consistent. The HTTP poll is the one channel.

**Decision** (field names shared with A28 and M38):
- The tool's result carries `cancel_requested: true` (with `success: true`) when the request was recorded. It never carries `cancelled`. An unknown or already-finished operation id is `success: false`.
- `cancelled: true` on the original call appears only when the add-in confirms: a `cancelled` terminal callback, or `cancelled` in its final JSON. It keeps the M26 decision journal.
- When a run ends with a `complete` or `error` callback while a request was outstanding, the original call returns its real outcome plus `cancel_not_honored: true`.
- Neither field gets an `error_pattern`.
- `/cancel-status/{operation_id}` answers for that id only, and stops answering true when the operation's terminal callback arrives (the flag is cleared then, and the operation is unregistered once its caller has read the result). A request that arrives after the terminal callback is refused, so it can never mark a finished run or bleed into the next one.
- The commented-out direct `Cancel` API call is deleted.
- The `/cancel-status` and `/cancel/{id}` response bodies keep their `cancelled` key: the add-in's poller (`clsMCP.CheckCancelled`) reads it, and it names the wire flag, not the tool result.

**What this rules out**: Reporting a stop from the request alone; a second, COM-based cancel channel. The marker is an MCP result field, not an add-in field. Revisit if the add-in adds a cancel acknowledgment callback distinct from `cancelled`.

**Relevant files**: `operation_manager.py` (`PendingOperation.finished` / `cancel_not_honored`, `route_callback`, `request_cancel`, `wait_for_completion`), `tools.py` (`vcs_cancel_operation`, `_carry_cancel_outcome`), `tests/test_cancel_request.py`.

---

## 2026-09-30 — A dispatcher refusal is `operation_already_running`; non-JSON from a JSON-contract method is a failure

**Trigger**: Only `vcs_call_vba` recognized the add-in's `VCS_API_REFUSED: ...` return. Every hard-coded call saw the marked string as data: policy setup looked accepted, cleanup looked successful, and `vcs_check_vba_compiled` / `vcs_compile_vba` read it as truthy (interface review F9).

**Decision**:
- The refusal is detected once, at the integration boundary. `call_sync` returns the refusal as a JSON failure object (`success: false`, `error` = the text without the marker, `error_pattern: operation_already_running`, `api_refused: true`), so every existing consumer that parses JSON sees a structured failure. `call_async` does the same for a bare refusal string and for the `{success: false, error: <marked>}` envelope. The two boolean tools check for that object before reading the value.
- `error_pattern` is `operation_already_running`. The add-in's text (`modAPI.bas` `RefuseReentrantCall`) says "another API command is still running", the same meaning as its nested-start refusal. The add-in also has a self-dispatch variant of the text (a defect, not a caller mistake); it gets its own pattern, `api_self_dispatch`, because waiting and retrying cannot help. It is told apart by the text "arrived back in the project that sent it".
- A non-JSON or non-object return from a JSON-contract method (`ExportByType`, `ImportByType`, `ExportObject`, `ImportObject`, `ExecuteSQL`, `RunVBA`, `SetOption`) is `success: false` with `error_pattern: invalid_addin_response`. Raw-contract methods (`GetOption` values, `GetLogContent` text) and `vcs_call_vba` keep wrapping a plain value as success.
- A refused `SetOperationPolicy` or `SetInteractionMode` means the object call is not dispatched. A refused `ClearOperationPolicy` becomes `policy_cleanup_error`.

**What this rules out**: A live reentrant check (a second call while the gate is held hangs the MCP, X01); the table-driven unit test `tests/test_api_refusal.py` is the evidence. `vcs_call_vba` keeps its own check and result shape. M40 and M41 build on this detection.

**Relevant files**: `addin_integration.py` (`api_refusal_payload`, `call_sync`, `call_async`), `tools.py` (`_addin_json_result`, `vcs_check_vba_compiled`, `vcs_compile_vba`).

---

## 2026-09-30 — One test-run verdict on every transport; an all-EMPTY run is not a success

**Trigger**: M33 (interface review F1, F2, F11, F10). The sync test parser
recalculated `success` from `failed == 0`, `errored == 0` and `subs > 0`. It
ignored the add-in's `cancelled`, `results_error` and its own `success: false`,
so a cancelled passing subset or an unsaved results file was `success: true`
on the sync and inline-sync paths but `false` on the callback path. Both paths
counted a run whose procedures were all EMPTY (no assertions) as a pass,
although the add-in's `allPassed` requires at least one passed test.

**Options explored**:
- Keep recalculating from the summary and patch each missing field per path:
  the paths had already drifted once, and they would again.
- For an all-EMPTY run, add a new `error_pattern`: callers would have one more
  pattern to learn for what is simply a failed run. Rejected.
- Treat mixed passing and EMPTY tests as a failure: stricter than the add-in's
  own verdict, and a single unfinished test would fail a working suite.
- Follow the add-in's `allPassed` after the terminal fields (chosen).

**Decision**: `_test_run_verdict` in `tools.py` is the only place that sets a
test run's `success`. The sync parser, the inline sync marker and
callback-loaded results all use it. Highest precedence first:
`decision_required`; `execution_interrupted` (applied later by the gate, M28);
`runtime_error` or an explicit add-in `success: false`; `cancelled`;
`results_error`; then `allPassed`. An add-in older than `allPassed` falls back
to `failed == 0`, `errored == 0` and `passed > 0`. The summary's `passed` and
`failed` count assertions, so with no failure or error `passed > 0` means a
test passed. An all-EMPTY run is `success: false` with an `error` naming the
EMPTY count and no `error_pattern`. Passing tests mixed with EMPTY ones stay a
success. An explicit `success: false` is never raised to true. The helper also
keeps `decisions` (parsed from a JSON string) and sets both `log_path` and
`logPath`.

**What this rules out**: Transport-specific verdict code for tests. A caller
that relied on an all-EMPTY run passing now sees `success: false`. Revisit if
the add-in's `allPassed` rule changes.

**Relevant files**: `src/msaccess_vcs_mcp/tools.py`, `docs/DIALOGS.md`,
`tests/test_run_tests.py`.

---

## 2026-09-30 — Require confirmed interactive mode before dispatch

**Trigger**: M32. `SetInteractionMode` previously returned VBA Empty even when
an enclosing caller-owned policy prevented interactive selection. MCP treated
any response other than an explicit refusal as acceptance and started work.

**Options explored**:
- Accept Empty for older add-ins: preserves compatibility but cannot establish
  the requested mode, so the operation could run under a stale policy.
- Clear an enclosing policy before selection: violates the scope owner's lease,
  including a policy retained after failed cleanup.
- Require the A24 response capability (chosen): rejects unsupported builds
  without relying on a release number that rebuilding does not increment.

**Decision**: Every explicit interactive tool path uses `select_interactive_mode`
before dispatch or callback registration. Only `success: true` with numeric
`effective_mode: 0` permits work. Add-in refusals pass through unchanged; Empty,
missing, malformed, or otherwise unconfirmed responses produce
`interaction_mode_unconfirmed`. These requests never clear an enclosing policy.
The minimum is an A24 build or later, as documented in the add-in's
`docs/noninteractive-dialogs.md` and its A24 decision entry.

**What this rules out**: Interactive requests against older add-ins that cannot
confirm acceptance, and implicit recovery by weakening another caller's scope.
Noninteractive operation setup and cleanup retain their existing contracts.

**Relevant files**: `decision_policy.py`, `docs/DIALOGS.md`,
`tests/test_interactive_mode.py`, `tests/test_interactive_mode_live.py`.

---

## 2026-09-30 — Gate-exempt tools get bounded worker threads of their own; the gate waits without a thread

**Trigger**: M24 (review finding 2, follow-up to M01). `_run_exempt_in_worker` ran each sync gate-exempt tool through `asyncio.to_thread` under `wait_for`. Cancelling the await does not stop the thread, so a dialog call into a hung Access UI thread (an MSAA call on a `NUIDialog` is bounded only by that ceiling) stayed in the default executor. `AccessGate.run_exclusive` waited for its slot through `asyncio.to_thread(self._slot.acquire, True, wait_sec)` in the same executor. With the executor full, the slot wait queued before its timeout started, and gated calls neither ran nor returned `server_busy`. Reproduced with a one-thread executor.

**Options explored**:
- *A larger default executor*: rejected. It only moves the limit, and every hang still adds a thread.
- *A dedicated `ThreadPoolExecutor` for exempt tools*: rejected. Its internal queue has no limit, and its threads are joined at interpreter exit, so a hung worker blocks shutdown.
- *Kill or interrupt the hung worker*: not possible for a thread blocked in a Win32 or COM call.
- *Daemon thread per call under a per-tool budget, admission by polling on the loop, and a slot wait that polls on the loop* (chosen).

**Decision**: `exempt_workers.ExemptWorkers` gives each sync gate-exempt tool at most `MAX_WORKERS_PER_TOOL` (2) live daemon threads, counted until the thread returns, abandoned or not, and at most `MAX_WAITING_PER_TOOL` (2) waiting calls. A new call never overtakes a waiting one. The ceiling from M01 (dialog timeout plus `EXEMPT_WORKER_MARGIN_SEC`) is now one deadline for admission and the run together. No worker in time, or a full waiting line (refused at once), gives `worker_capacity_unavailable`, recoverable. A run past the deadline is `tool_timeout` as before. The budget is per tool, so hung `vcs_list_dialogs` calls do not lock out `vcs_get_recent_calls` or `vcs_dismiss_dialog`. `AccessGate._acquire_slot` polls a non-blocking acquire on the event loop, so `server_busy` arrives within `ACCESS_VCS_BUSY_WAIT_SEC` however full any executor is. The click stays in the worker that inspected the dialog, so M23's reservation deadline (twice the click timeout plus the settle margin) still holds. `Win32Backend.click` was already bounded (`SendMessageTimeoutW` with `SMTO_ABORTIFHUNG`, and a `WM_NULL` probe before an MSAA press). Tests now drive both paths against a real button window whose thread never pumps messages.

**What this rules out**: `asyncio.to_thread` or the default executor for gate-exempt tool bodies or for the gate's slot wait. A slot handoff can lag by up to one poll (20 ms). A refused or timed-out exempt call writes no usage entry until an abandoned worker finishes. Revisit the per-tool numbers if a client legitimately runs more than two concurrent calls of one dialog tool.

**Relevant files**: `exempt_workers.py`, `access_gate.py` (`_acquire_slot`), `tools.py` (`_run_exempt_in_worker`), `tests/test_exempt_workers.py`, `tests/test_access_gate.py`, `tests/conftest.py`, `AGENTS.md`, `docs/DIALOGS.md`.

---

## 2026-09-30 — Reserve an interruption before the click; the finished call waits for the click to settle

**Trigger**: M23. The interruption was recorded after `backend.click` returned. Ending a runtime-error dialog lets Access resume, so the blocked call could return and finish (no record yet: success) before the click returned. The record then found the gate free and became a free `last_interruption` with `busy_with: None`, or was dropped by the gate's cleanup. The existing tests released the call only after the dismissal finished, which hid the ordering.

**Options explored**:
- *Record before the click, as a confirmed interruption*: rejected. A click that was never sent would mark a call that nothing interrupted.
- *Keep recording after the click and have the finished call poll briefly*: rejected. No bound fits every delivery time, and a record arriving after the poll still lands on a free gate or a later call.
- *Reserve before the click, settle after, and make the finished call wait for a pending reservation up to its deadline* (chosen).
- *Keep `WindowBackend.click` a bool*: rejected. False covered both "nothing was sent" (wrong owner, unresponsive window, button refused the action) and "sent but not confirmed" (a timed-out `SendMessageTimeout`). The first must leave the call alone; the second must not leave a success.

**Decision**: `_press_fresh` reserves the interruption against the in-flight call on the same database after `_reverify` and the chooser, right before `backend.click`, when `_interrupts_execution` holds for the fresh kind and chosen button. `_close_fresh` does the same for a cancel. `WindowBackend.click` returns `CLICK_DELIVERED`, `CLICK_NOT_SENT` or `CLICK_UNCERTAIN`, and the reservation settles to `confirmed`, gone, or `uncertain` (a free reservation is dropped rather than kept uncertain). Records are keyed by reservation token and carry the process identity, so a second reservation never overwrites a confirmed one. `finish_gated_call(call_id, result)` waits on a condition for that call's pending reservations, bounded by each one's deadline (twice the click timeout plus `CLICK_SETTLE_MARGIN_SEC`: the Win32 click may spend one timeout on the responsiveness probe and one on the press). Confirmed wins: `execution_interrupted` as before. Otherwise any uncertain or unsettled reservation gives `success: false` and `interruption_uncertain: true`, with `error_pattern: interruption_uncertain` only when the result was a success and not `decision_required`. Precedence: `decision_required` > `execution_interrupted` > plain error > `interruption_uncertain`. For an async handler `_then` runs the finalization in a worker thread so the wait never blocks the event loop. `finish_gated_call(call_id, None)` in the gate cleanup still uses records up without waiting. `begin_gated_call` drops every record on its database that is not the starting call's: besides free records, that clears a reservation that read an earlier call as in flight just before the gate released it and landed after that call's cleanup. The public `vcs_dismiss_dialog` and `vcs_recover_dialogs` results are unchanged: any undelivered click is still `dismiss_uncertain`.

**What this rules out**: Recording an interruption after the action that causes it. Treating a click that was never sent as an interruption, or an unconfirmed one as a success. A late settle reaching a later call: a reservation belongs to one `call_id` and is gone once that call finishes. Revisit the deadline with M24 if a hung click is moved out of the worker that holds it.

**Relevant files**: `dialog_recovery.py` (`_reserve_interruption`, `_settle_interruption`, `begin_gated_call`, `finish_gated_call`, `_press_fresh`, `_close_fresh`, `Win32Backend.click`), `tools.py` (`_then`), `usage_logging.py` (`log_tool_call`), `tests/test_interruption_records.py`, `docs/DIALOGS.md`.

---

## 2026-09-30 — An interrupted call is finalized inside the logging layer; the log takes the result's own outcome

> **⚠ Partially superseded** (2026-09-30): the record is now reserved before the click, and the finished call waits for a click still in flight before it is finalized; the late-record rule is unchanged for reservations made after the handler returned. See "Reserve an interruption before the click; the finished call waits for the click to settle" above.

**Trigger**: M28. `with_logging` recorded the handler's result, then the gated wrapper applied the interruption record. A handler that returned success after its runtime error was ended logged `success: true`, and `vcs_get_recent_calls` showed that. Separately, a `decision_required` result that was also interrupted dropped the interruption: the flag was never added.

**Options explored**:
- *Wrap the whole gated run (gate wait, handler, finalize) in `with_logging`*: rejected. The logged time would include the gate wait, and `server_busy` answers, which never ran, would start producing `tool_call` entries.
- *Re-log after the gate*: rejected. Two entries per call, and recent-call history would show both.
- *Apply the record inside the logging layer* (chosen): `vcs_tool` wraps the handler with `finish_gated_call` and logs that, so one entry records the final result and its timing covers the handler and the finalization.
- *Let the post-gate cleanup keep applying late records to the result*: rejected. The client and the log would disagree again. A record attached after the handler returned did not interrupt it, so the cleanup in `finally` only uses it up.
- *Keep guessing the logged `error_pattern` from the error text*: rejected for results that carry their own. The guess turned `execution_interrupted` into whatever the text matched (for example `file_not_found`), and a `success: false` result without `error` was logged as a success.

**Decision**: `finish_gated_call` sets `success: false` and `execution_interrupted: true` on any matching interrupted result, then applies precedence to `error_pattern` only: `decision_required` stays primary (by flag or by pattern), otherwise `execution_interrupted`. `log_tool_call` treats `success: false` as a failure, uses the result's own `error_pattern` when it has one (the text guess is the fallback), and copies `execution_interrupted` and `policy_cleanup_error` into the entry.

**What this rules out**: Logging a gated handler's provisional result. Reading an interruption off a later call. Revisit the late-record rule with M23, which reserves the interruption before the dialog action releases the call.

**Relevant files**: `tools.py` (`vcs_tool`, `_then`), `dialog_recovery.py` (`finish_gated_call`), `usage_logging.py` (`log_tool_call`), `tests/test_interruption_records.py`, `docs/DIALOGS.md`.

---

## 2026-09-30 — Re-verify dialog content and re-run the policy before a click

**Trigger**: M22. The pre-click check compared process identity, the dialog handle, its PID and the inspected `ButtonInfo`, but not title, class or text. A modal box cannot change its own text, but it can be replaced: the inspected box closes, another opens, and Windows reuses the handles. Reproduced: an inspected "Hello" box became "Delete all records?" with the same handles, `recover_dialogs(policy="safe")` clicked OK, and the result reported "Hello".

**Options explored**:
- *Compare the kind only*: rejected. "Hello" and "Delete all records?" are both OK-only `vba_msgbox`.
- *Compare a content signature only*: catches a replaced box, but the safety rule would still be applied to the inspection, not to the box being clicked.
- *Re-run the policy on the fresh window only*: catches the destructive case, but lets a different harmless box through under an explicit dismiss, and a result could still describe the earlier box.
- *Both (chosen)*.

**Decision**: `_reverify` re-lists windows and returns the fresh window and its kind, or None. It requires the same PID and creation time and the same signature: class, title, full text (`window.texts`), classified kind, and the whole button list. `_press_fresh` is the only place a button is clicked: it re-verifies, runs the caller's chooser on the fresh window (`_button_named` for an explicit button, `_policy_button` for `recover_dialogs`), and clicks the `ButtonInfo` read from the fresh window. `_close_fresh` is the only place a window is closed. Any mismatch is `dialog_changed` with nothing clicked. The `closed` record and the interruption record are built from the fresh window. `recover_dialogs` sets top-level `success: false`, `error_pattern: dialog_changed` when a dialog it would have clicked changed; `dismiss_uncertain` takes precedence. `msaa.press` still re-checks role, name and state for NetUI buttons.

**What this rules out**: The race between the last check and the click cannot be closed from outside Access. A box can still be replaced after `_reverify` lists windows and before `BM_CLICK` or `accDoDefaultAction` arrives. The window is now two back-to-back calls instead of the whole inspection. `vcs_dismiss_dialog` compares against its own inspection at the start of the call, not the listing the agent read before deciding; the agent's think time is covered only in that a result describes the box actually clicked. Revisit if a caller-supplied expected signature is added to `vcs_dismiss_dialog`.

**Relevant files**: `dialog_recovery.py` (`_signature`, `_reverify`, `_press_fresh`, `_close_fresh`, `_policy_button`, `recover_dialogs`), `tests/test_dialog_identity.py`, `docs/DIALOGS.md`.

---

## 2026-09-30 — NetUI dialogs are read and pressed through MSAA with raw ctypes

**Trigger**: M19. The add-in's `MsgBox2` (the `@`-separated bold `MsgBox` form) and Access's own error dialogs appear as a top-level `NUIDialog` whose only child is a `NetUIHWND`. There are no Win32 `Static` or `Button` children, so the inspector dropped the window as ordinary: `vcs_list_dialogs` said `ready: true` while a box blocked the call, and `vcs_dismiss_dialog` returned `dialog_not_found`. A plain `MsgBox` (VBA or `Eval`) is still `#32770`.

**Options explored**:
- **`WM_CLOSE` on the `NUIDialog`**. It closes the box, but gives no text or buttons to classify, and on a Yes/No box it picks an answer nobody chose. Rejected.
- **Keystrokes (Enter, Alt+letter)**. The module's rule is no keystrokes or coordinate clicks; focus is not guaranteed. Rejected.
- **UI Automation via `comtypes` or `pywinauto`**. It works, but adds a dependency for about a dozen calls. Rejected for now.
- **MSAA through pywin32 late binding** (`ObjectFromLresult` to `IDispatch`, then `Invoke`). NetUI's `IDispatch::Invoke` returns `E_NOTIMPL`, so this fails on every property.
- **MSAA through direct `IAccessible` vtable calls with ctypes (chosen)**. `oleacc` and `oleaut32` ship with Windows. The object is fetched with `SendMessageTimeout(WM_GETOBJECT, SMTO_ABORTIFHUNG)` so a hung window is skipped. Visible static texts (role 41) and push buttons (role 43) are read, and `STATE_INVISIBLE` placeholders are skipped. A button is pressed with `accDoDefaultAction`.

**Decision**: `msaa.py` holds the ctypes `IAccessible` reader and `press`. `ButtonInfo` gains `path`: for a NetUI button, `hwnd` is the `NetUIHWND` host and `path` is the child-index path from its client object. `WindowBackend.click` takes the `ButtonInfo`, not a bare hwnd. The pre-click re-verify compares the whole `ButtonInfo`, and `press` checks again that the control at that path is a visible, enabled push button with the same name. `NUIDialog` gets the same classification rules as `#32770`. Only `NUIDialog` windows of an `MSACCESS.EXE` process are read, so other Office apps' dialogs are not touched.

**What this rules out**: The live backend now makes cross-process COM calls into Access's UI thread for NetUI dialogs. Each call is bounded only by the worker-thread ceiling of the gate-exempt tools, not per call; the `WM_GETOBJECT` fetch and a `responsive` probe before a press are the per-call guards. If Office changes the NetUI tree shape, or drops MSAA in favour of UIA only, revisit with a UIA backend.

**Relevant files**: `src/msaccess_vcs_mcp/msaa.py`, `src/msaccess_vcs_mcp/dialog_recovery.py` (`ButtonInfo`, `classify_window`, `Win32Backend.list_windows`/`click`, `_netui_contents`), `tests/test_dialog_classification.py`, `tests/test_dialog_live.py`, `docs/DIALOGS.md`.

---

## 2026-09-30 — MCP never calls the add-in's dialog APIs; Win32 only

**Trigger**: The add-in exposes `ListAddinDialogs` and `DismissAddinDialog` (`modDialogInspect`, on `VCS.API`). Using them for add-in windows looked cheaper than a second classifier.

**Options explored**:
- *Call the add-in APIs for `frmVCS*` forms, Win32 for the rest*: rejected. They run on the Access COM thread, which a blocking dialog occupies, so the call hangs at the moment it is needed. They see only add-in forms, and `close` refuses while an operation runs.
- *Try the add-in API first, fall back to Win32*: rejected. Two classifiers for one window, and the first attempt can hang.

**Decision**: The dialog tools use Win32 enumeration and window messages only, off the Access gate, keyed on process identity. Add-in windows are classified by caption like any other. The add-in APIs stay for callers that are inside Access.

> **⚠ Partially superseded** (2026-09-30): "Win32 only" meant "never through the add-in", and that still holds. Office NetUI boxes (`NUIDialog`) have no Win32 buttons, so the inspector also reads and presses them through MSAA, from outside Access and off the gate. See "NetUI dialogs are read and pressed through MSAA with raw ctypes" above.

**What this rules out**: A gate-exempt tool that calls `VCS.API`. Revisit only if the add-in can answer without the COM thread.

**Relevant files**: `dialog_recovery.py`, `access_gate.EXEMPT_TOOLS`, add-in `modDialogInspect.bas`, `docs/DIALOGS.md`.

---

## 2026-09-30 — The add-in's start result is non-final; MCP never reports a start as success

**Trigger**: The MCP spec assumed the sync return of `MergeBuild` carried the final result. It carries a start result (`started: true`); the outcome comes on the completion callback.

**Options explored**:
- *Treat the sync return as final*: rejected. It reports success for a merge that has not finished, or that then blocks on a prompt.
- *Poll the add-in for the outcome*: rejected. Polling is a COM call, and the gate and a blocking dialog are what this work avoids.
- *Return `success: true, started: true` when no callback exists*: rejected. Nothing can confirm the merge finished.

**Decision**: A refusal in the sync return is returned at once, normalised, and the duplicate on the callback is dropped (the operation is unregistered, so the first arrival wins). A started marker waits for the callback, and the payload goes through the shared decision normaliser. With no callback (none configured, or the async start failed and MCP fell back to a sync call) the result is `success: false, started: true, completion_unconfirmed: true` with a pointer to `log_path` and `vcs_get_recent_calls()`, and no `error_pattern`. `RunFilteredTests` is the exception: it runs the tests before returning, so its return is final. The policy for a full merge or test run is passed inline to `MergeBuild` and `RunFilteredTests`; only the scoped and single-object calls set it through `SetOperationPolicy` and clear it in `finally`.

**What this rules out**: Reading `success: true` from a start result as an outcome. Retrying an unconfirmed start blindly. Revisit if the add-in gains a synchronous final result for `MergeBuild`.

**Relevant files**: `tools.py` (`_merge_sync`, `_normalize_import_result`, `_call_under_policy`), `decision_policy.py` (`is_start_refusal`, `clear_operation_policy`), `docs/DIALOGS.md`.

---

## 2026-09-30 — `safe` clicks only `vba_msgbox`; `access_dialog` is report-only

**Trigger**: `SAFE_OK_KINDS` also held `access_dialog` and `vba_compile_error`, which the combined spec lists as report-only and as `end_runtime_error`'s.

**Options explored**:
- *Change the spec to match the code*: rejected (user decision). An Access error or warning dialog carries text the agent should read; clicking it away hides a failure.
- *Keep the compile-error click under `safe`*: rejected. Compile errors belong to `end_runtime_error`.

**Decision**: `safe` clicks OK only on an OK-only `vba_msgbox` without destructive text. `end_runtime_error` adds End on a runtime-error dialog (before, and independent of, the destructive-text check, which guards OK clicks only) and OK on an OK-only compile error. No policy clicks `access_dialog`. An explicit `vcs_dismiss_dialog(button=...)` is unchanged.

**What this rules out**: Auto-clicking Access error or warning dialogs. Revisit if a real Access dialog is found that blocks a run and carries no information.

**Relevant files**: `dialog_recovery.py` (`SAFE_OK_KINDS`, `auto_button`), `docs/DIALOGS.md`.

---

## 2026-09-29 — Click OK-only VBA MsgBox of any caption (`vba_msgbox`)

**Trigger**: After the classification work, a VBA `MsgBox` with a custom caption was kind `unknown` and never clicked, so an unattended run stalled on a dialog Access cannot continue past.

**Options explored**:
- *Caption allow-list only*: the user must guess every caption in advance. Not the fix. The single-OK signature below covers OK-only dialogs of any caption.
- *Click any OK-only dialog by button count alone*: rejected as a bare rule. The signature is a `#32770` dialog whose actionable buttons are exactly one OK (a Help button is ignored), and destructive text still blocks the click.
- *Click multi-button custom dialogs (OK/Cancel, Yes/No)*: rejected. Picking a button is a decision the tool cannot make.

**Decision**: New kind `vba_msgbox`, assigned only to the single-OK signature with any caption. Microsoft Access and Microsoft Visual Basic captions still get `access_dialog`. `safe` clicks it when the text has no save, delete, discard or overwrite words, and reports kind, title and text in `closed`. This is a deliberate exception to "recognised captions only" from the entry below.

**What this rules out**: Auto-clicking any dialog with two or more buttons, or a single non-OK button (Retry, Save). Keystrokes and coordinate clicks. Revisit if a real OK-only dialog is found whose OK is not safe to press.

**Relevant files**: `dialog_recovery.py` (`MSGBOX_KIND`, `SAFE_OK_KINDS`, `classify_window`), `docs/DIALOGS.md`, `specs/dialog-recovery-review-fixes.md`.

---

## 2026-09-29 — Dialog recovery rules: identity, positive signature, interruptions, off-event-loop

> **⚠ Partially superseded** (2026-09-30): in rule 3, the interruption is reserved before the click and counts only once delivery is confirmed; an unconfirmed delivery gives `interruption_uncertain`, not `execution_interrupted`. Records are kept per reservation, with the PID and creation time on each. See "Reserve an interruption before the click; the finished call waits for the click to settle" above.

> **⚠ Partially superseded** (2026-09-30): in rule 3, an interrupted `decision_required` result keeps `error_pattern: decision_required` but also gets `success: false` and `execution_interrupted: true`. See "An interrupted call is finalized inside the logging layer; the log takes the result's own outcome" above.

> **⚠ Partially superseded** (2026-09-29): rule 2 has one exception. A dialog with a single OK button and any other caption is kind `vba_msgbox` and is clicked by `safe`. See "Click OK-only VBA MsgBox of any caption (`vba_msgbox`)" above.

> **⚠ Partially superseded** (2026-09-30): in rule 1, the pre-click re-verify also compares class, title, text, kind and buttons, and re-runs the policy on the fresh window. See "Re-verify dialog content and re-run the policy before a click" above.

**Trigger**: Two-axis review of commit `e982918` found the dialog-recovery code did not meet the entry below. This records the four rules the fixes (M01 to M05) implement. `specs/dialog-recovery-review-fixes.md` holds the detail.

**Options explored**:
- *Treat an unreadable PID, name or creation time as "probably Access"*: rejected. A failed query is not permission; a reused PID could get a stranger's dialog dismissed.
- *Classify by caption alone or by absence of danger words*: rejected. Unknown is the default and a kind needs a positive signature.
- *Let a dismissed error dialog be reported as the operation's success*: rejected. The clicked-away error must survive to the final result.
- *Run gate-exempt tools on the event loop or the COM thread*: rejected. Those are what a blocking dialog or gated call occupies.

**Decision**:
1. **Identity**: no click unless the process is confirmed Access and its creation time is known and unchanged. Unconfirmed means no action, with a distinct `error_pattern`. Handles are re-verified just before the click.
2. **Positive signature**: a dialog is a known kind only by a positive signature. Everything else is `unknown` and is reported, never clicked.
3. **Interruptions** (M05): dismissing a runtime or compile error records an interruption keyed by PID plus creation time, carrying the in-flight gated call. When that call finishes it is forced to `success: false`, `execution_interrupted: true`, `error_pattern: execution_interrupted`, keeping the original text, and the record is removed. With no call in flight it shows as `last_interruption` until the identity changes or the next gated call on that database starts. Precedence: `decision_required` > `execution_interrupted` > plain error. A call on a different database does not adopt the record.
4. **Off the event loop**: every gate-exempt tool runs in a worker thread with a ceiling slightly above the dialog timeout, not only the four dialog tools.

> **⚠ Partially superseded** (2026-09-30): The worker is no longer `asyncio.to_thread`. Each tool has a bounded budget of daemon threads, and the ceiling also covers the wait for one. See "Gate-exempt tools get bounded worker threads of their own; the gate waits without a thread" above.

**What this rules out**: Treating a dismissed error dialog as success. Acting on an unverified process. Adding a gate-exempt tool that uses the COM thread. Revisit the interruption rule if records need to span databases.

**Relevant files**: `dialog_recovery.py`, `access_gate.py` (`InFlight.call_id`), `tools.py` (`vcs_tool` gated path), `docs/DIALOGS.md`.

---

## 2026-09-29 — Noninteractive add-in runs, Win32 dialog recovery off the Access gate

> **⚠ Partially superseded** (2026-09-29): "clicks only recognized OK-only dialogs" now also covers an OK-only MsgBox with any caption (`vba_msgbox`). See "Click OK-only VBA MsgBox of any caption (`vba_msgbox`)" above.

> **⚠ Partially superseded** (2026-09-30): "recognized OK-only dialogs" is now only `vba_msgbox` (plus an OK-only compile error under `end_runtime_error`). `access_dialog` is report-only. See "`safe` clicks only `vba_msgbox`; `access_dialog` is report-only" above.

**Trigger**: An agent driving Access gets stuck when the add-in or VBA opens a modal dialog, and a second MCP call cannot inspect that dialog because it waits on the same Access COM gate.

**Options explored**:
- *`DoCmd.SetWarnings False`*: rejected. It does not cover `MsgBox`, the conflict form, VBA End/Debug, or compile errors, and it hides failures.
- *Keep `eimSilent` and its caller-supplied default*: rejected for agent calls. Several prompts default to Yes, which approves a destructive choice with no record.
- *UI Automation only*: Win32 `BM_CLICK` on the button window already names the control. UI Automation would be a second path for the same buttons. Standard dialogs expose their captions to `GetWindowText`, so the Win32 path is the one implemented.
- *Put recovery on the Access gate*: rejected. The gate is what a modal dialog blocks.

**Decision**: Agent test and merge calls default to a scoped noninteractive mode (`decision_policy=block`). The add-in acknowledges OK-only prompts, applies an explicit conflict or decline policy, and otherwise returns `decision_required` without opening UI or choosing Yes. `Finish` restores the previous mode. Dialog inspection uses Win32 enumeration filtered to one Access PID plus creation time, and those tools are exempt from the Access gate. Automatic recovery clicks only recognized OK-only dialogs, plus End when `end_runtime_error` is requested. Debug, save/discard, and unknown dialogs are reported. Dismissing an error dialog records the failure and does not retry the mutation.

**What this rules out**: Treating a clicked error dialog as a successful operation. Silently overwriting merge conflicts on agent imports unless `decision_policy` is `prefer_source`. Sending keys, Enter, Escape, or coordinates. Closing a different Access instance because a PID was reused. Using `SetWarnings` as the suppression mechanism.

**Relevant files**: add-in `modDialogPolicy.bas`, `modUIUtil.MsgBox2`, `clsConflicts.ResolveOrPrompt`, `clsOperation`; MCP `dialog_recovery.py`, `access_gate.EXEMPT_TOOLS`, `tools.py` (`vcs_run_tests`, `vcs_import_objects`, `vcs_list_dialogs`, `vcs_dismiss_dialog`, `vcs_recover_dialogs`, `vcs_automation_status`).

---

## 2026-09-10 — Persist server-created Access; own it by PID

**Trigger**: Merge / test / edit loops were spawning a new `MSACCESS.EXE` per
tool call. Attach costs ~3 ms; spawn+quit costs ~3 s median. Ninety-seven
`vcs_run_vba` / import / test calls failed with "Cannot find Access instance"
because the worker only looks in the ROT. `_owns_app` is recomputed per
connection, so leaving a window open reclassified it as user-owned.

**Options explored**:
- *Keep quit-per-call, document `ACCESS_VCS_LEAVE_ACCESS_OPEN`.* Rejected.
  The flag already existed and was unused in this workspace; the reliability
  failures are the default path.
- *In-memory owned-PID set.* Rejected. The server restarts a median of every
  6.8 minutes; an in-memory set would orphan windows several times an hour.
- *Never Quit when UserControl is True.* Still rejected: owned instances
  deliberately set `UserControl` so they look interactive.
- *Persist by default + disk PID registry* (chosen). GetObject reattaches, so
  the bound is one window per distinct database, not an accumulating leak.
  Rebuild pre-flight closes registry-owned holders; user-owned still refuse.
  Stuck owned instances are recycled after a failed recovery probe.

**Decision**: Default `ACCESS_VCS_LEAVE_ACCESS_OPEN` to true. Record
`(pid, create_time, database_path)` in `~/.msaccess-vcs-mcp/owned-instances.json`.
`_owns_app` now means "the server created this process and may close it when
it must," not "quit at end of call."

**What this rules out**: Quitting a leftover owned instance at the end of
every tool call. Revisit if long-lived Access shows a correctness problem
(stale VBA project, memory growth) that in-process reset cannot fix.

**Relevant files**: `access_com/instance_registry.py`,
`access_com/connection.py`, `vba_worker_manager.py`, `tools.py`, `main.py`.

---

## 2026-09-10 — Ownership needs proof; closing needs identity

**Trigger**: Review of the persistence work above. Ownership had become the
authority to close somebody's Access process, but the code answered "is this
ours?" with heuristics that failed toward yes: `is_owned` accepted a record
with no creation stamp, a failed `tasklist` looked identical to "no Access
running" and silently wiped the registry, and closure resolved an instance by
*path* and quit whatever came back. Two rebuild-blocking bugs sat alongside
it: `GetObject(path)` launches Access when nothing holds the file, and that
process was never registered; and every tool call loads the add-in as a
library, which locks it, while the pre-flight matched only on open database.

**Options explored**:
- *Treat unconfirmed as owned, since the server usually is the one running
  Access.* Rejected. The cost of a false positive is closing a user's window
  with unsaved work; the cost of a false negative is a rebuild that refuses
  and tells the user which window to close.
- *Verify ownership by `UserControl` or window title instead of a registry.*
  Rejected again (as in the entry above): owned instances deliberately set
  `UserControl`, so it cannot distinguish them.
- *Leave a hung owned instance alone and let the rebuild refuse.* Rejected on
  the user's call: the instance is server-created, so unsaved state is
  acceptable loss, and refusing leaves a window nobody can act on.
- *Match the add-in by path at each of the fifteen `load_addin` call sites.*
  Rejected. Recording it once inside `load_addin` covers every path through
  the code, including ones added later.

**Decision**: Ownership requires a confirmed `(pid, create_time)` match;
every ambiguity resolves to "not ours". `list_access_pids_or_none()`
distinguishes a failed process query from an empty one, and the registry
keeps its records when liveness is unknowable. Closure re-verifies pid and
creation time against the record before touching an instance, runs the
graceful `Quit` in a worker thread under `ACCESS_VCS_CLOSE_TIMEOUT_SEC`
(a blocking COM call here would hang the Access gate), then force-terminates
a survivor. A record is dropped only once the process is confirmed gone.
`load_addin` records the loaded add-in so rebuild pre-flight can close
library holders, and that pre-flight moved inside the gate.

**What this rules out**: Any ownership signal that is not the durable
registry, and any close path that authorizes on a path match. Revisit the
terminate policy if owned instances ever hold state a user would miss —
today they do not, because the user's own windows are never server-created.

**Relevant files**: `access_com/instance_registry.py`,
`access_com/connection.py`, `access_com/process_qos.py`,
`addin_integration.py`, `tools.py`, `vba_worker_manager.py`.

---

## 2026-08-28 — Optional leave-open for COM-created Access

> **⚠ Superseded** (2026-09-10): Persist-by-default plus a durable PID
> registry replaced the opt-in flag as the default. The flag remains as an
> explicit `false` to restore quit-per-call. The 2026-08-28 leak concern
> assumed accumulation; GetObject reattach bounds the count to one window
> per database. See "Persist server-created Access; own it by PID" above.


**Trigger**: Sequential `msaccess-vcs run-tests` always Quit the Access the
CLI child created, so there was no way to time a boosted process and then
reattach to the same PID. QoS is process-lifetime; attach to a leftover
boosted instance is the comparison that matters.

**Options explored**:
- *Never Quit when UserControl is True.* Rejected as the default. Owned
  instances would leak Access processes after every CLI or MCP session.
- *ACCESS_VCS_LEAVE_ACCESS_OPEN* (chosen). When set, `close()` still releases
  COM but skips `CloseCurrentDatabase`/`Quit`. Attach later finds the same
  PID via GetObject. Default remains Quit for owned instances.

**Decision**: Opt-in env flag only. User-owned Access is unchanged.

**What this rules out**: Leaving every MCP-created Access running by default.

**Relevant files**: `access_com/connection.py`, `tests/test_access_visibility.py`.

---

## 2026-08-28 — CLI test output is compact pytest-style dots

**Trigger**: `msaccess-vcs run-tests` printed every MCP log fragment (Access
same-line console layout) and then the full per-test JSON. A 468-test run
took ~3 minutes, most of it HTTP. The last line was not a human summary.

**Options explored**:
- *Leave the stream as raw `Log.Add`.* Rejected after the first full run.
- *Live current-test name via `\\r` on stderr.* Rejected. Cursor's capture is
  not a TTY, so overwrite is unreadable.
- *Infer a dot from each start-of-test `(n/m)` progress.* Rejected. That is
  468 HTTP posts. The add-in now coalesces fast passes into log lines.
- *Print `.` from those batched log payloads; ignore `(n/m)`; compact JSON
  plus a human completion line* (chosen). Wrap dots at 80. Names remain the
  add-in's ≥ 1s PASS and FAIL/ERROR/EMPTY lines. `vcs_run_tests` as an MCP
  tool still returns the `tests` map.

**Decision**: The CLI is for watching. The MCP tool is for programmatic
reruns. Completion is a sentence such as `Tests passed. 468 subs, 2281
assertions, 3 empty in 14.42s`, not a sentinel.

**What this rules out**: Dumping the `tests` map to CLI stdout. Printing
`(n/m)` progress during `run-tests` (rebuild/export still do). Using progress
callbacks as the pytest dot stream. A `VCS_CLI_EXIT` token.

**Relevant files**: `cli.py` (`ProgressPrinter`, `compact_result_payload`,
`completion_message`), add-in `clsLog.cls`, `clsTestRunner.cls`.

---

## 2026-08-28 — Agent test runs stream through APIAsync and the CLI

**Trigger**: Agent-initiated test runs returned one JSON blob after the suite
finished. Cursor shows only "Running..." for MCP progress, so a long run looked
stuck. Builds already streamed: `APIAsync` + HTTP `Log.Add`/`Log.Progress` +
`msaccess-vcs rebuild-addin` printing each notification. Tests already posted
those callbacks when `MCP.IsActive`, but `vcs_run_tests` used blocking
`call_sync("RunFilteredTests")`, which never registered a callback.

**Options explored**:
- *Keep the sync COM call and tail `TestRun_*.log`.* Rejected. Duplicates the
  callback path builds already use, and the log is gitignored so agents miss it.
- *Register the callback then block in `call_sync`.* Rejected. The tool handler
  would hold COM for the whole suite and could not emit MCP progress until
  return.
- *Same async path as export* (chosen). `RunFilteredTests` joins the
  `APIAsync` timer list. MCP waits on callbacks. `msaccess-vcs run-tests`
  prints each line. Old add-ins still return `{sync: true, result: ...}`.

Failed tests complete as `eorFailed`, which posts type `error`. The results
file rides on every terminal callback as `results_path` so MCP can still
return per-test JSON for reruns. What the CLI prints is the compact stream
above, not every Access-console fragment.

**Decision**: `vcs_run_tests` follows the export pattern. The CLI is the live
output path; the MCP tool remains for short or programmatic runs. Tool
`success` still comes from the runner summary, not from `Operation.Result`.

**What this rules out**: Treating a finished-with-failures suite as a lost
result. Adding a second log-tailer for tests. Exempting `vcs_run_tests` from
the Access gate (the suite runs in the connected instance).

**Relevant files**: `tools.py` (`vcs_run_tests`), `cli.py`,
`operation_manager.py`, add-in `modAPI.bas`, `clsOperation.cls`,
`clsTestRunner.cls`.

---

## 2026-08-28 — Owned Access instances get UserControl after the database opens

**Trigger**: Headless MCP test runs created a new Access for
`Version Control.accda` and set `Visible = True`, but the window often never
appeared as a normal interactive app. Automation-created Access with
`UserControl` still False stays off the desktop even when Visible is True.
The common `EnsureDispatch` new-instance path never set `UserControl`; only
the isolated `DispatchEx` path restored it after open.

**Options explored**:
- *Set UserControl before OpenCurrentDatabase.* Rejected. AutoExec/`AutoRun`
  would see a person watching and open the installer form.
- *Visible only, as in the 2026-08-12 rule.* Insufficient for COM-created
  instances; that is the gap this closes.
- *UserControl = True after open, on owned instances only* (chosen). AutoExec
  already ran with the flag down. Attached user instances are left alone.

**Decision**: `_get_access_app` still shows the window after the database is
open, then sets `UserControl = True` when `_owns_app`.
`validate_access_installation()` stays hidden (no database, quit immediately).

**What this rules out**: Hiding MCP-created Access for tidiness. Setting
`UserControl` before a database opens. Forcing `UserControl` on instances the
server attached to.

**Relevant files**: `access_com/connection.py`
(`_ensure_owned_instance_interactive`), `tests/test_access_visibility.py`.

---

## 2026-08-28 — MCP-launched Access prefers a full-power core

**Trigger**: The add-in's agentic rebuild was slower than a ribbon rebuild of the
same source because Windows scheduled the COM-launched, windowless Access
process onto an LP-E core. Access is single-threaded, so that one core is the
operation.

**Options explored**:
- *Pin CPU affinity to P-cores.* Rejected. Topology differs by SKU, and pinning
  is the wrong default on a machine with only performance cores.
- *Promote every Access process the server sees.* Rejected. The server often
  attaches to a database the user already has open.
- *EcoQoS off + Above Normal for MCP-launched processes only* (chosen).
  `AccessConnection` applies it when `_owns_app` is true. `vcs_rebuild_database`
  applies it to an empty `EnsureDispatch` instance, not one that already has a
  database open. `vcs_rebuild_addin` snapshots `MSACCESS.EXE` PIDs before launch
  and promotes processes that appear afterward (builder, silent installer).
  The probe in `validate_access_installation` is not promoted: it quits
  immediately. Failures are swallowed.

**Decision**: Disable `PROCESS_POWER_THROTTLING_EXECUTION_SPEED` and set Above
Normal. Do not set an affinity mask. User-owned Access is unchanged.

**What this rules out**: Pinning cores. Changing QoS on an Access instance the
user launched. Assuming this closes the entire agent/ribbon gap.

**Relevant files**: `access_com/process_qos.py`, `access_com/connection.py`,
`rebuild_watcher.py`, `tools.py`.

---

## 2026-08-27 — Carry the existing HTTP stream through add-in self-rebuild

**Trigger**: The first `vcs_rebuild_addin` smoke test completed in Access but
showed only "Rebuild launched" and then hung. Even without the hang, watching
the coarse status file could report only `building`, `compiling`, and
`installing`, while the build log contained the useful per-object detail.

**Options explored**:
- *Tail `Build_*.log` from Python.* Viable fallback, but duplicates the stream
  already emitted by `Log.Add` and `Log.Progress`, requires discovering a
  timestamped file, and cannot preserve native progress values.
- *Have Worker.vbs implement HTTP posting.* Rejected. That would be a second
  callback implementation and a second formatter.
- *Pass the existing callback identity to builder Access* (chosen). The parent
  callback server outlives the launching Access instance, and the builder can
  invoke the existing `APIAsync(..., "Build", source)` path.

**Decision**: `vcs_rebuild_addin` registers an operation before launch and
passes its callback JSON as the second `RebuildAddIn` argument. The add-in
extracts the shell-safe URL and operation ID, carries them through Worker.vbs,
and reconstructs the callback JSON in builder Access. Build `log`, `progress`,
and terminal callbacks feed the same monotonic reporter as status phases. A
build `complete` callback ends only the callback pump; the status watcher
remains authoritative for compile/install and the final result. The original
`phaseStarted` now also crosses the worker boundary instead of being replaced.
`ReadDirectoryChangesW` uses overlapped I/O and `CancelIo`, avoiding the
deadlock caused by closing a handle with a synchronous read on another thread.

**What this rules out**: Tailing the build log as the primary live channel.
Reimplementing HTTP in Worker.vbs. Treating the build-phase `complete` callback
as proof that compile/install finished. Replacing durable terminal status with
an ephemeral callback.

**Relevant files**: `src/msaccess_vcs_mcp/tools.py`,
`operation_manager.py`, `rebuild_watcher.py`, `cli.py`,
`tests/test_rebuild_watcher.py`; add-in `clsVersionControl.cls` and
`clsWorker.cls`.

---

## 2026-08-27 — Real-time operation feedback: monotonic MCP progress, rebuild watcher, CLI

> **⚠ Partially superseded** (2026-08-27): builder HTTP callbacks now cross
> the disconnected worker boundary. The status watcher remains for
> compile/install and durable recovery, but is no longer the only live source.
> See "Carry the existing HTTP stream through add-in self-rebuild" above.

**Trigger**: Long Access operations looked idle in Cursor. VBA progress resets per category, so forwarding `current/total` as MCP `progress` violated the strictly-increasing rule. Full merge and database rebuild also dropped the MCP context (`ctx=None`). Add-in self-rebuild still required the agent to poll `rebuild-status.json` after `vcs_call_vba` returned. Cursor 3.13 often shows only "Running..." for progress notifications.

**Options explored**:
- *Forward VBA `current/total` as MCP progress.* Rejected. The spec requires each notification's `progress` to increase; category resets (28/30 queries, then 1/50 modules) are invalid and clients drop them.
- *MCP logging notifications (`ctx.log`).* Rejected. Deprecated 2026-07-28 / SEP-2577.
- *MCP Tasks.* Rejected. Extra protocol surface; does not fix Cursor display.
- *Have the add-in worker POST HTTP callbacks after Access quits.* Deferred. Would need a new worker contract; `rebuild-status.json` is already the durable source of truth.
- *New `vcs_rebuild_addin` that blocks on COM.* Rejected again. The host Access instance must quit so files can be replaced.
- *Agent-side Read polling of the status file.* Rejected as the normal workflow. Agents sat on timers after the work had finished.
- *Dedicated CLI that reimplements COM.* Rejected. A second operation path would drift from the MCP tools.

**Decision**: A shared `MonotonicProgressReporter` owns an incrementing sequence and keeps VBA counts in the message, with `total=None`. Full export, merge, and database rebuild all pass `ctx`. `vcs_rebuild_addin(source_dir)` launches via existing `RebuildAddIn` (gate held only for launch), then watches `rebuild-status.json` with `ReadDirectoryChangesW` (poll fallback), correlating `phaseStarted`. Cancel abandons the wait only — it does not kill `MSACCESS.EXE` or `wscript`. `vcs_call_vba` stays launch-only. `msaccess-vcs` is a stdio MCP client for the same tools, printing progress with immediate flush. Status-file polling remains recovery after a client timeout or stall.

**What this rules out**: Treating Cursor chat progress as a keepalive or as guaranteed UI. Killing Access/wscript when a client stops waiting. Making generic `vcs_call_vba` wait minutes for install. Adopting deprecated MCP logging notifications. A second COM implementation beside the MCP server.

**Relevant files**: `src/msaccess_vcs_mcp/operation_manager.py`, `rebuild_watcher.py`, `cli.py`, `tools.py` (`vcs_rebuild_addin`), `access_gate.py`, `tests/test_operation_manager.py`, `tests/test_rebuild_watcher.py`, `tests/test_cli.py`.

---

## 2026-08-21 — Self-heal corrupted pywin32 gen_py cache

**Trigger**: MCP startup died with `module 'win32com.gen_py.4AFFC9A0-...' has no attribute 'CLSIDToClassMap'` when `%TEMP%\gen_py` held a half-built Access type-library folder (only `__pycache__`, no wrapper `.py` files). Deleting the folder manually and retrying `EnsureDispatch` regenerated the wrappers and succeeded. The failure had recurred several times.

**Options explored**:
- **Do nothing** — users must manually delete `%TEMP%\gen_py`. Simple, but fatal: `validate_access_installation()` exits and Cursor reports `MCP error -32000: Connection closed`.
- **Wipe all of `gen_py` on every startup** — always safe, but forces a full makepy rebuild for every COM library the process touches.
- **Fall back to late-bound `Dispatch`** — avoids the cache, but breaks early-binding semantics (`Application.Run` return shape).
- **Targeted folder delete + one retry** — parse the type-library folder from the `AttributeError`, `rmtree` only that folder, drop `sys.modules` entries, `gencache.Rebuild()`, retry `EnsureDispatch` once.

**Decision**: Centralize early-bound dispatch in `ensure_dispatch()` (`access_com/connection.py`) with targeted purge and a single retry. Route every `EnsureDispatch("Access.Application")` call site through it. Log `gen_py_cache_rebuilt` to the diagnostic stream.

**What this rules out**: Startup wipe-all of `gen_py`. Late-bind fallback for Access Application dispatch. Folding this into `com_recovery.py` (that module is for live RPC disconnects, not local makepy cache repair). More than one automatic retry per call.

**Relevant files**: `src/msaccess_vcs_mcp/access_com/connection.py`, `config.py`, `validation.py`, `tools.py`, `usage_logging.py` (`gen_py_cache` error pattern), `tests/test_gen_py_cache.py`.

---

## 2026-08-21 — The installed add-in is never a target, for any tool

**Trigger**: Agents repeatedly passed the installed add-in as `database_path` — to
run the add-in's own tests, and to host `RebuildAddIn`. That file exists to be
loaded as a library: opening it as a database, or writing into it, resets a VBA
project while it is executing. A test run makes the point sharply, needing two
roles at once from two different files (the install supplies the runner and
`TestAssert`; the current database holds the code under test), so collapsing them
makes the runner scan the library's own components and has
`InstallTestAssertModule` write into the executing project. The server's own
instructions had drifted into recommending the wrong host, naming "a user database
Access already has open" for rebuilds — which sent agents hunting for a database to
borrow, and into the add-in repo's `Testing` folder when they could not find one.

**Options explored**:
- *Rely on the add-in's VBA guard alone.* Rejected. `ExecuteTests` does refuse a
  run there, but only once the install is already the current database — and
  opening it as a database is the thing being prevented, not just the run that
  follows. It also covers only test runs.
- *Per-tool guards.* Tried and withdrawn within a day. Two were written, for
  `vcs_run_tests` and for `RebuildAddIn` via `vcs_call_vba`, and they still left
  `vcs_export_database`, `vcs_run_vba`, and the rebuild's own `output_path`
  uncovered. Tailoring each message to name the right alternative host was the
  argument for them, and it turned out there is only one answer to name.
- *Compare full file names.* Rejected. An install configured for the compiled
  add-in is a `.accde` built from the same `.accda`, and only one of the two
  appears in `ACCESS_VCS_ADDIN_PATH`, so the other variant would slip past.
- *Check folder parameters too* (`source_dir`, `output_dir`). Rejected. An export
  folder beside the install is not the install, and refusing it would block reading
  the install's own exported source.

**Decision**: `_refuse_installed_addin_target` runs inside the `vcs_tool` wrapper,
ahead of the gate and of any COM work, and rejects `database_path`, `output_path`,
or `template_path` naming the install with `error_pattern:
installed_addin_refused`. The comparison, `_is_installed_addin_path`, ignores the
extension, mirroring the add-in's `modInstall.PathsMatchIgnoringExtension`. The
message names the development copy in the add-in's repository and rules out the
substitutes agents actually reached for: a user database, the repo's `Testing`
folder, a scratch `.accdb`. `vcs_get_version_info()` is the sanctioned way to learn
the installed version, and needs no path.

**What this rules out**: Any tool reaching the installed file, including ones not
yet written — a new tool inherits the refusal from the decorator. Hosting a rebuild
or a test run on a borrowed database. Asking a person to nominate a host. A
tool-specific exemption would have to be argued as an exemption, since there is no
longer a per-tool guard to quietly omit.

**Relevant files**: `src/msaccess_vcs_mcp/tools.py`
(`_refuse_installed_addin_target`, `_is_installed_addin_path`,
`_installed_addin_target_refusal`, `vcs_tool`), `AGENTS.md`, `README.md`,
`docs/AGENT_WORKFLOWS.md`, `tests/test_installed_addin_guard.py`. The add-in side is
`modInstall.CurrentDbIsInstalledAddIn` and the guard in
`clsVersionControl.ExecuteTests`.

---

## 2026-08-21 — The install path comes from the add-in's own settings key

**Trigger**: `get_default_addin_path()` built `%AppData%\MSAccessVCS\Version
Control.accda` from constants. Both halves can be wrong: the installer lets the
user choose a folder, and a compiled install is a `.accde` — the installer deletes
the `.accda` in that case, so the guessed path names a file that does not exist.
Refusing the installed add-in as a target made this load-bearing rather than
merely untidy, since a wrong path means the refusal protects nothing.

**Options explored**:
- *Access Menu Add-Ins `Library` value.* Rejected. It does hold the full path with
  the correct extension in one read, but it lives under
  `Office\{Application.Version}\...`, and nothing records which Office version ran
  the installer — an external reader has to enumerate versions and guess. Reading
  two values from one fixed key is less machinery and no less certain.
- *Ask a running Access instance for `GetInstalledAddInFileName`.* Rejected.
  Requires COM and a loaded add-in to answer a question needed before either.
- *Probe the disk for whichever of the two extensions exists.* Rejected as the
  primary source: it infers from a side effect instead of reading the recorded
  choice, and reports nothing when neither file is present.

**Decision**: Read `HKCU\Software\VB and VBA Program Settings\MSAccessVCS\Install`
— `Install Folder` and `Compile accde` — which is exactly what the add-in's
`modInstall.GetInstalledAddInFileName` joins. `Install Folder` is absent for a
default install because the installer deletes the value rather than writing the
default, so absence means `%AppData%\MSAccessVCS`, not "not installed".
`Compile accde` holds a VBA integer, so True arrives as `-1`. Cached for the life
of the process, with `reset_addin_path_cache()` for tests and reinstalls.
`ACCESS_VCS_ADDIN_PATH` still overrides everything.

**What this rules out**: Reading the install path from anywhere else — the Menu
Add-Ins registration, the trusted-location entry (folder only), or a reconstructed
`%AppData%` path. Treating a missing `Install Folder` as a missing install.
Assuming the extension.

**Relevant files**: `src/msaccess_vcs_mcp/config.py`
(`get_default_addin_path`, `_read_install_setting`, `reset_addin_path_cache`),
`tests/test_config.py`. The add-in side is `modInstall.GetInstallSettings` /
`GetInstalledAddInFileName`.

---

## 2026-08-21 — The vcs_call_vba timeout worker owns its COM apartment

**Trigger**: Giving `vcs_call_vba` a hard timeout meant dispatching
`Application.Run` on a daemon thread so the parent could stop waiting. Handing that
thread the caller's `app` proxy failed before reaching Access at all —
"CoInitialize has not been called" while the thread had no apartment,
`RPC_E_WRONG_THREAD` (0x8001010E) once it did. Which of the two surfaced looked
random, so the failure read as an add-in problem. `RebuildAddIn` was unreachable
through this tool for as long as the timeout existed.

**Options explored**:
- *Marshal the caller's pointer into the worker* (`CoMarshalInterThreadInterfaceInStream`
  / `CoGetInterfaceAndReleaseStream`). Rejected. A marshalled STA pointer serializes
  the call back onto the apartment that created it, so the calling thread blocks
  anyway and the timeout this function exists to impose never fires.
- *Run the call on the main thread and time out around it.* Rejected. There is
  nothing to time out around — a blocking COM call on the calling thread cannot be
  abandoned.
- *Re-acquire Access from the Running Object Table inside the worker* (chosen).
  Yields an apartment-local proxy, and `VCSAddinIntegration._find_access_in_rot`
  already existed for exactly this reason in the add-in probe.

**Decision**: `_run_application_call_with_timeout` takes `db_path` and the worker
calls `pythoncom.CoInitialize()`, then re-acquires the instance from the ROT. Without
`db_path` there is nothing to look up, so the caller's proxy is used as before — no
worse than it was, and the timeout may simply not fire.

**What this rules out**: Sharing a COM proxy across threads anywhere in this server.
Testing this path by stubbing `AccessConnection` alone — a test that does so silently
reaches the real ROT and passes or fails depending on whether Access happens to be
open, which is how `test_repo_addin_path_not_self_host_refused` became
environment-dependent. Stub `_find_access_in_rot` too; `_fake_access` in
`tests/test_call_vba_guards.py` wires both.

**Relevant files**: `src/msaccess_vcs_mcp/tools.py`
(`_run_application_call_with_timeout`), `src/msaccess_vcs_mcp/addin_integration.py`
(`_find_access_in_rot`), `tests/test_call_vba_guards.py`.

---

## 2026-08-12 — Access instances holding a database stay visible

> **⚠ Partially superseded** (2026-08-28): `Visible = True` is necessary but
> not sufficient for instances the server creates. Those also need
> `UserControl = True` *after* the database is open, or the window often never
> appears as a normal interactive app. See "Owned Access instances get
> UserControl after the database opens" above.

**Trigger**: COM automation starts Access hidden, and nothing in the server
ever showed the window. Access still asks questions only a person can answer —
trust prompts, conversion prompts, a VBA error breaking into the debugger — and
behind an invisible window those look like a hang. Every later call then blocks
on an instance the user cannot see, may not know exists, and has no way to
clear. The `.accda` work made this sharper by opening databases in instances
the server creates itself.

**Options explored**:
- *Leave instances hidden and detect the stall instead.* Rejected. The probe
  timeouts already report "Access is likely in break mode or blocked on a modal
  dialog", which is the right diagnosis and still leaves the user with no
  window to act on. Detection is not recovery.
- *Show only instances the server creates.* Rejected. An attached instance can
  raise the same dialog, and moniker binding launches a hidden Access when
  nothing had the file open, so "attached" does not imply "a person can see
  it".
- *Make visibility configurable.* Deferred. An opt-out re-creates the
  unresolvable-dialog failure for whoever sets it; wait for a concrete need
  (unattended CI is the plausible one).

**Decision**: Any instance the server drives a database through ends up
visible. Two helpers in `access_com/connection.py` carry the rule so it cannot
be half-applied: `ensure_access_visible(app)` (best-effort, never fails an
operation) and `open_current_database(app, path)`, which replaces every bare
`app.OpenCurrentDatabase(...)` call. Visibility comes *after* the open, because
showing a window can set `UserControl` and the add-in's `AutoRun` reads that
flag to decide whether a person is watching. `validate_access_installation()`
is the one exception: no database, quit immediately, so a window would only
flash.

**What this rules out**: Bare `OpenCurrentDatabase` calls — the AutoExec and
visibility rules travel with the helper, and a new call site that skips it
silently loses both. Hiding instances for speed or tidiness. Treating a probe
timeout as sufficient handling for a blocked dialog.

**Relevant files**: `src/msaccess_vcs_mcp/access_com/connection.py`
(`ensure_access_visible`, `open_current_database`), `validation.py`,
`vba_worker_manager.py`, `tools.py` (`vcs_rebuild_database`), `config.py`,
`tests/test_access_visibility.py`.

---

## 2026-08-12 — Opening .accda as the current database for add-in self-tests

> **⚠ Partially superseded** (2026-08-21): the `.accda` this opens is the
> *development copy* in the add-in's repository. The installed copy is now refused
> as a target by every tool, before anything opens it. See "The installed add-in is
> never a target, for any tool" above.

**Trigger**: The add-in's own tests only run when the add-in is the current
database, because the runner walks `CurrentVBProject`. Every `vcs_run_tests`
call against `Version Control.accda` failed with "Cannot find Access instance".
`GetObject(path)` binds a file moniker that resolves through the COM
registration for the extension; that registration opens .accdb and .mdb as the
current database but treats .accda as an add-in, so the bind fails. The
fallback instance then had no current database, and tier 2 of
`_find_access_in_rot` matches on `CurrentDb().Name`, so nothing was found.

**Options explored**:
- *Have a person pre-open the file* (the PowerShell `OpenCurrentDatabase`
  recipe in the add-in repo's `docs/agentic-rebuild.md`). Rejected: a human
  step per iteration, which is what this whole path exists to remove. That
  section of the add-in docs is stale as of this entry.
- *Register a file moniker for .accda.* Rejected: machine-wide COM
  registration change to fix one client.
- *Call `OpenCurrentDatabase` when the moniker bind fails* (chosen). Same
  approach `validation.py` already takes, and no new configuration.

**Decision**: `_open_as_current_database` runs only on the GetObject-failure
branch, and opens only on an instance we own, so a user's database is never
displaced. Two constraints surfaced in review. `UserControl` is lowered across
the open: `OpenCurrentDatabase` runs the target's AutoExec, and the add-in's
`AutoRun` opens its installer form when `Application.UserControl` says a person
is watching — which would strand the instance we are about to automate, and did
so whenever `_create_isolated_instance` supplied the process. A failed open is
also swallowed rather than raised, so `_get_current_db` can still reach its DAO
strategies for read-only callers.

**What this rules out**: Documenting a manual pre-open step for add-in tests.
Setting `UserControl` before a database opens — `_create_isolated_instance`
still needs it for teardown survival, so it has to be restored after AutoExec,
not before. Letting an `OpenCurrentDatabase` failure escape `_get_access_app`,
which would bypass the DAO fallback chain that predates this path.

**Relevant files**: `src/msaccess_vcs_mcp/access_com/connection.py`
(`_open_as_current_database`, `open_current_database`),
`tests/test_accda_current_database.py`.

---

## 2026-08-12 — Agentic add-in rebuild via existing vcs_call_vba

> **⚠ Partially superseded** (2026-08-27): the normal agent path is now
> `vcs_rebuild_addin`, which watches `rebuild-status.json` after launch. The
> rejection of "a tool that waits on COM" still holds; the new tool waits on
> the status file with the Access gate released. `vcs_call_vba` remains the
> launch-only escape hatch. Agent-side Read polling is recovery, not the
> primary workflow. See "Real-time operation feedback" above.
>
> **⚠ Partially superseded** (2026-08-21): the host is the development copy of
> the add-in in its repository, not an arbitrary open database — see "The
> installed add-in is never a target, for any tool" above. `vcs_call_vba` also has a timeout now
> (`ACCESS_VCS_CALL_VBA_TIMEOUT_SEC`), which closed the follow-up noted below and
> brought its own COM constraint; see "The vcs_call_vba timeout worker owns its
> COM apartment" above.

**Trigger**: Agents iterating on the VCS add-in source could not rebuild
`Version Control.accda` without a person, because server instructions told them
the add-in cannot be rebuilt via MCP (closing every Access instance would close
the user's other databases).

**Options explored**:
- *New `vcs_rebuild_addin` tool that waits on COM.* Rejected. The Access instance
  the MCP is talking to is deliberately quit; a blocking COM wait would hang.
  `vcs_call_vba` already dispatches `RebuildAddIn` through `CallByName`.
- *Have the server poll the status file.* Rejected. The agent already has a Read
  tool, and the status path is known from the source folder. No new MCP surface.

**Decision**: Document the existing `vcs_call_vba` → `VCS.API` → `RebuildAddIn`
path. The add-in writes `<source>/logs/rebuild-status.json` and refuses when
another Access instance holds a file the rebuild replaces. A COM error after
launch is expected.
`vcs_call_vba` still has no timeout; the worker sleeps before quit so the JSON
can return. Adding a timeout remains a follow-up.

**What this rules out**: Treating `vcs_rebuild_database` as the add-in rebuild
path. A dedicated rebuild-add-in tool unless `vcs_call_vba` grows a timeout.
Closing other Access windows from the MCP server stays out; the add-in reports
the offenders instead, for reasons recorded in the add-in repo's own decision
log. A second Access instance held by this server only blocks the rebuild if the
add-in is loaded in it, which happens as soon as any `vcs_*` call routes through
the add-in's API.

**Relevant files**: `src/msaccess_vcs_mcp/tools.py` (instructions, `vcs_call_vba`
example), `README.md`, `AGENTS.md`, `docs/AGENT_WORKFLOWS.md`.

---

## 2026-08-07 — Scoped object_types via ImportByType / ExportByType

**Trigger**: `vcs_import_objects` and `vcs_export_database` accepted `object_types` but largely ignored them. Import always ran a full `MergeBuild`. Export only special-cased a modules-only list into `ExportVBA` and otherwise exported everything. Agents passed `object_types=["modules"]` expecting a partial merge and got a whole-project one with no warning. The documented `overwrite` flag on import never mapped to any add-in behavior.

**Options explored**:
- *Keep ignoring object_types / document the lie*: rejected. Agents already rely on the parameter.
- *Add per-type MCP wrappers*: rejected. The add-in already exposes category-scoped APIs.
- *Route to ImportByType / ExportByType via call_sync* (chosen): no add-in rebuild; `modAPI.API` reaches any public `clsVersionControl` method through `CallByName`.

**Decision**: When `object_types` is set, call `ImportByType` / `ExportByType` synchronously and attach `log_path` like other sync results. When unset, keep the existing async full-project path (`MergeBuild` / `Export`/`FullExport`). Replace `overwrite` with `full_import`, which maps to `blnFullImport` (reload all files in the named categories vs only index-marked changes). Retire the modules-only `ExportVBA` shortcut so every scoped export uses one rule and reconciles deletions.

Scoped calls are sync-only because those methods are not in `APIAsync`'s command list. That is acceptable for category-sized work; revisit by adding them to the async list if blocking becomes painful. Single-type lists are passed as a bare string to avoid COM array-marshalling edge cases; multi-type lists remain Python lists.

**What this rules out**: Documenting or reintroducing an `overwrite` "skip existing" mode the add-in does not have. Treating a scoped merge as a safe partial without stating orphan deletion and no backup. Preferring `ExportVBA` for modules-only from these tools. Mapping "every category + full_import=True" as the recommended full rebuild path — use `vcs_rebuild_database` instead.

**Relevant files**: `src/msaccess_vcs_mcp/tools.py` (`vcs_import_objects`, `vcs_export_database`, `_scoped_types_arg`); add-in `clsVersionControl.ExportByType` / `ImportByType`, `modBuild.MergeScoped`.

---

## 2026-07-30 — Reaching the add-in API from MCP: vcs_call_vba, not vcs_run_vba

**Trigger**: An agent spent a long session trying to run the add-in's round-trip test harness (`VCS.RunRoundtripTests`) through `vcs_run_vba`. Every attempt returned an empty string. That looked like a broken add-in, then a stuck Access instance, and the workaround attempted next (`HandleRibbonCommand`) corrupted the host VBA project with error 2517 and required closing and reopening the database. The session ended by telling the user to paste a command into the Immediate window — for a capability the tools already had.

The cause is structural. `vcs_run_vba` is itself delivered through `modAPI.API`, which has a `Static IsRunning` re-entrancy guard. Submitted code therefore runs *inside* an API call, and anything it calls back into the API is nested by construction and refused. The guard returned `Empty` silently, which is indistinguishable from a method that legitimately returned nothing.

Three separate defects turned a one-line answer into a multi-hour dead end:

1. `vcs_call_vba` — the tool that does work — documented `"Version Control.API"` as its example. That never resolves. `Application.Run` matches a loaded VBA *project* name (`MSAccessVCS`), not the file name (`Version Control`). `"MSAccessVCS.API"` works but only once something has already loaded the add-in; only the full path is correct from a cold start.
2. `vcs_call_vba` did not unwrap the result. Early-bound `Run` returns a 31-element tuple — the return value followed by `Run`'s own 30 `Arg` slots, unused ones showing `DISP_E_PARAMNOTFOUND` (`-2147352572`). `VCSAddinIntegration.call_api_function` already handled this; the generic tool did not, so callers saw the payload buried in noise.
3. Nothing in the tool descriptions said which tool reaches the API, or that `vcs_run_vba` structurally cannot.

**Options explored**:
- *Document the full path and move on*: rejected. It puts an install-specific absolute path in every call and still leaves the silent-`Empty` trap for the next agent.
- *Relax the re-entrancy guard to allow nesting*: rejected. The guard protects `Operation` state owned by the outer call. The nested call is genuinely unserviceable; the defect is that it was refused silently, not that it was refused.
- *Resolve an alias qualifier server-side* (chosen): the server already knows `ACCESS_VCS_ADDIN_PATH`.

**Decision**: `vcs_call_vba` now accepts `"VCS.API"` (also `"Version Control.API"`, `"MSAccessVCS.API"`, case-insensitive) and rewrites the qualifier to the configured add-in's full path, which loads it on demand. Any other qualifier passes through untouched, so an explicit path or a user's own module still works. The result tuple is unwrapped to its first element, matching `call_api_function`. A "cannot find the procedure" failure now explains the project-name-versus-file-name rule. Server instructions name `vcs_call_vba` as the route to API methods and state that `vcs_run_vba` cannot be.

Paired with an add-in change: `modAPI.API` and `APIAsync` now return a message naming the refused method and pointing at `vcs_call_vba`, instead of returning `Empty`. `API` prefixes it with `API_REFUSED_PREFIX` (`"VCS_API_REFUSED: "`); `APIAsync` embeds it in its JSON. `vcs_call_vba` matches the prefix and reports `success: False` so a refusal is never mistaken for data.

That started as an `Err.Raise` and had to be changed after testing. An error raised inside a library database does not propagate across `Application.Run` into the calling project's handler — even with `On Error GoTo` active in the caller, Access shows a modal "Run-time error" dialog and blocks until a human dismisses it. Since this guard only trips on a nested call, which is by definition the case that crosses that boundary, raising was guaranteed to hit it. A blocking dialog is worse for automation than the silence it replaced, so the refusal travels as a marked return value instead.

**What this rules out**: Calling add-in API methods from inside `vcs_run_vba` — use `vcs_call_vba`. Adding tools that wrap individual API methods (`vcs_run_roundtrip_tests` and the like); the generic route now works and does not need per-method surface. Treating a `Run` result as a scalar anywhere else without unwrapping the tuple first. One accepted risk: a user module genuinely named `VCS` would have its qualifier rewritten — revisit if that ever surfaces.

**Relevant files**: `src/msaccess_vcs_mcp/tools.py` (`vcs_call_vba`, `_resolve_addin_function_name`, `_describe_run_failure`, server instructions); add-in `modules/API/modAPI.bas` (`RefuseReentrantCall`, `ERR_API_REENTRANT`).

---

## 2026-04-30 — EnsureDispatch ownership fix (DispatchEx fallback)

**Trigger**: Disabling the MCP tool in Cursor killed the user's Access window (with their open database). `AccessConnection._get_access_app()` called `EnsureDispatch("Access.Application")` when `GetObject(db_path)` failed, but `EnsureDispatch` can silently attach to an already-running user-owned Access instance instead of creating a new one. The code set `_owns_app = True` unconditionally, so `close()` called `_app.Quit()` on the user's session. This is the exact bug db-inspector-mcp fixed in their "DispatchEx fallback for COM instance conflicts" decision.

**Decision**: After `EnsureDispatch`, check `app.CurrentDb()` to detect whether the returned instance already has a database open. If it has *our* database, reuse and set `_owns_app = False`. If it has a *different* database, fall back to `DispatchEx("Access.Application")` which always spawns an isolated COM server process. If no database is open, it's a genuinely fresh instance and `_owns_app = True` is correct. `close()` only calls `Quit()` on instances the server actually created.

**What this rules out**: Setting `_owns_app = True` for any `EnsureDispatch` result without checking `CurrentDb()` first. Any future code path that creates an Access COM reference must verify ownership before storing it. If `DispatchEx` proves unreliable on specific Access/Windows configurations, revisit with the same guard pattern.

**Relevant files**: `src/msaccess_vcs_mcp/access_com/connection.py` (`_get_access_app`, `_create_or_reuse_instance`, `_create_isolated_instance`).

---

## 2026-04-30 — Isolate `vcs_run_vba` behind a timeout-controlled worker

> **⚠ Partially superseded** (2026-04-30): Subprocess isolation was tried and reverted the same day. The child process had to cold-start Python, import the package, initialize COM, and re-acquire Access via the ROT — adding seconds of overhead per call. Worse, `subprocess.Popen.communicate(timeout=45)` blocked the MCP event loop for the entire duration, preventing the server from processing other requests (including `ListToolsRequest`), which caused `BrokenResourceError` crashes when Cursor timed out. Reverted to a daemon-thread + `thread.join(timeout)` approach modelled on this project's existing `_probe_with_timeout` and db-inspector-mcp's `_run_dao_with_timeout`. The thread creates its own COM apartment via `pythoncom.CoInitialize()` and re-acquires Access through the ROT, so `thread.join(timeout)` fires even when the COM call blocks. A class-level `_active_worker` guard prevents zombie thread pile-up (same pattern as db-inspector-mcp). The recovery state machine, error classification, logging events, and timeout env vars are unchanged. `vba_worker.py` (the subprocess entry point) has been deleted.

**Trigger**: A long-running `vcs_run_vba` call left Cursor's MCP connection closed, followed by reconnect attempts timing out and later calls failing with `Not connected`. The existing add-in probe had `ACCESS_VCS_PROBE_TIMEOUT_SEC`, but the actual `Application.Run(..., "RunVBA", code)` call still happened synchronously inside the stdio MCP process. If Access entered VBA break mode, showed a modal dialog, or never returned from the submitted snippet, the server process could hang or die before it could return a structured error.

**Options explored**:
- **Keep synchronous COM and rely on the existing add-in probe**. Rejected: the probe only proves `GetVCSVersion` responds before dispatch; it does not bound the later arbitrary `RunVBA` call where the observed failure occurred.
- **Launch a short-lived Python COM worker subprocess per call**. Tried and reverted: cold-start overhead (Python import + COM init + ROT lookup) was too high, and `subprocess.communicate(timeout)` blocked the async event loop causing `BrokenResourceError` crashes.
- **Daemon thread + `thread.join(timeout)` (chosen)**. Same pattern as the existing add-in probe and db-inspector-mcp's DAO timeout. Worker thread creates its own COM apartment, re-acquires Access via ROT, and runs VBA. Main thread returns `TimeoutError` if the deadline expires; the daemon thread finishes naturally. No true cancellation, but no event-loop blocking either.
- **Use the async callback path (`APIAsync` + `OperationManager`)**. Not applicable to ad-hoc VBA snippets which don't have a detached async add-in contract.
- **Kill `MSACCESS.EXE` automatically**. Rejected: Access may be user-owned with unsaved work.

**Decision**: `vcs_run_vba` runs arbitrary VBA in a daemon worker thread with a hard timeout (`ACCESS_VCS_RUN_VBA_TIMEOUT_SEC`, default 45s, overridable per call via `timeout_seconds`). A per-database `COMRecoveryManager` classifies COM transport failures and runs a short `ACCESS_VCS_RECOVERY_PROBE_TIMEOUT_SEC` probe before later calls. Pre-dispatch failures are retried once after a successful probe; failures during `run_vba` are not auto-retried because the snippet may already have started.

**What this rules out**: Running `RunVBA` directly on the main thread without a timeout boundary. Do not solve this class of failure by killing `MSACCESS.EXE` unless ownership tracking proves the server created an isolated Access instance. The daemon-thread approach accepts that a timed-out worker thread stays alive until Access responds (or the process exits); this is the same trade-off db-inspector-mcp makes for DAO queries.

**Relevant files**: `src/msaccess_vcs_mcp/vba_worker_manager.py` (thread-based timeout, probe, retry policy, `_active_worker` guard), `src/msaccess_vcs_mcp/com_recovery.py` (classification and per-database state), `src/msaccess_vcs_mcp/tools.py` (`vcs_run_vba` delegates to the worker manager), `src/msaccess_vcs_mcp/usage_logging.py` (worker/recovery events), `src/msaccess_vcs_mcp/main.py` (startup/shutdown/fatal diagnostics), `.env.example`, `AGENTS.md`, `tests/test_vba_worker_manager.py`, `tests/test_usage_logging.py`.

---

## 2026-04-27 — Always-on diagnostic stream + tiered usage-log defaults

**Trigger**: When the server is launched from a user-level Cursor MCP config (`~/.cursor/mcp.json`), the working directory becomes the user's home folder, the upward `.env` walk in `_find_project_root` resolves to `C:\Users\<user>` (because of the `.cursor` directory there), no `.env` is found, and `ACCESS_VCS_ENABLE_LOGGING` defaults to `false` -- so the server falls silent at exactly the moment we most need its self-diagnostics. The lazy MCP-roots `.env` discovery in `_ensure_env_loaded` was added to fix this, but its only debugging output went to stderr (which Cursor's user-level MCP pipes to the server-output pane that the agent cannot read). We had no way to verify that `list_roots()` was even being called, let alone what it returned. Compounding this: the existing single `ACCESS_VCS_ENABLE_LOGGING` switch made auditability all-or-nothing, which conflicts with this server's central capability -- it makes destructive changes to live databases and a forensic record is exactly what an operator needs by default, *but* code-execution bodies (SQL/VBA fragments) frequently embed business data, table names, or PII that the same operator may not want persisted to disk.

**Options explored**:
- **Single switch, leave default off**. Status quo. Rejected: provides no audit trail by default for a tool that mutates live databases, *and* still doesn't solve the observability gap that prompted this work.
- **Single switch, flip default on, log code bodies in full**. Maximum forensic fidelity. Rejected: lumping audit metadata and SQL/VBA bodies behind one switch forces privacy-conscious users to choose between auditability and confidentiality.
- **Hash code bodies (SHA-256) instead of redacting them**. Provides "did this exact code run before?" lookup without storing plaintext. Rejected: agents often generate semantically identical SQL with trivial whitespace/parameter differences, so hash equivalence is too brittle to be useful, and a hash still leaks more about *which* code ran than `code_length` alone.
- **Per-tool granular logging knobs** (`LOG_SQL_BODIES` separate from `LOG_VBA_BODIES`, etc.). Rejected: complexity exceeds value; one body switch is enough until we have evidence of actual divergent needs.
- **Mirror tool calls into the diagnostic file**. Rejected: bloats lifecycle debugging with high-frequency event noise. The two streams have genuinely different lifetimes (lifecycle: rare, small) and review patterns (audit: frequent, large), and conflating them defeats the whole point of a separate always-on stream.
- **Two-stream design with tiered default-on usage logging (chosen)**. (1) Always-on diagnostic stream at `~/.msaccess-vcs-mcp/logs/vcs-mcp-diagnostic.jsonl`, gated only on `ACCESS_VCS_DISABLE_DIAGNOSTIC_LOG=true`, capturing server lifecycle events (`server_start`, `startup_env_load`, `lazy_env_load`, `lazy_init_started`, `lazy_init_skipped`, `list_roots_failed`, `list_roots_response`, `lazy_init_loaded`, `lazy_init_no_env_in_roots`, `usage_log_status`). (2) Usage stream defaults `ENABLE_LOGGING=true` so audit metadata is captured by default. (3) Code-execution bodies are redacted to `code_length` only; restore the old behavior with `ACCESS_VCS_LOG_CODE_CONTENT=true`. (4) Parameter keys matching `password|secret|token|api[_-]?key|connection[_-]?string` (case-insensitive) are auto-masked to `"<redacted>"` regardless of any switch -- defense in depth for accidentally-named-bad params.

**Naming sub-decision**: All log filenames carry the `vcs-mcp-` prefix (e.g. `vcs-mcp-usage.jsonl`, `vcs-mcp-diagnostic.jsonl`). Existing setups commonly point multiple MCP servers at a single shared `logs/` directory, where a generic name like `usage.jsonl` from one server collides with the same name from another. The prefix self-identifies the source without needing to inspect file contents.

**Discoverability sub-decision**: `vcs_get_version_info` returns `usage_log_path`, `diagnostic_log_path`, and `log_code_content` so an agent can `Read` either file directly and know whether code bodies are being captured -- without the agent having to inspect env vars itself.

**What this rules out**: Future entries should not add new "is logging on?" switches without first considering whether a *tier* of an existing stream covers the use case. Any new sensitive field that appears in logged data must either go through `_sanitize_parameters` (so the secret-key auto-mask catches it) or be added to the tiered model with its own explicit opt-in. The diagnostic stream is intentionally not a place to mirror tool-call events; future contributors who need richer audit data should extend the *usage* stream, not the diagnostic one. Reverting `ENABLE_LOGGING` to default-off would re-introduce the original "no audit trail by default" hazard and must be justified explicitly. The `code_length`-only default for code bodies is deliberate; do not "improve" it with a preview/first-N-chars representation -- length-only was chosen specifically because partial-content redactions leak just enough to enable per-keyword fishing while still failing to round-trip the actual code.

**Relevant files**: `src/msaccess_vcs_mcp/usage_logging.py` (new `_initialize_diagnostic_logging`, `log_diagnostic_event`, `get_diagnostic_log_path`, `is_diagnostic_logging_enabled`; `_write_log_entry` parameterized over handler; usage filename renamed to `vcs-mcp-usage.jsonl`; `log_code_execution` redaction; `_sanitize_parameters` secret-key masking; `_get_logging_config` adds `log_code_content` and flips `enabled` default), `src/msaccess_vcs_mcp/main.py` (`server_start` and `usage_log_status` events; diagnostic-log status print), `src/msaccess_vcs_mcp/config.py` (`startup_env_load` event in `_load_env_files`, `lazy_env_load` event in `_load_env_from_directory`; `ACCESS_VCS_ENABLE_LOGGING` default flipped, `ACCESS_VCS_LOG_CODE_CONTENT` added), `src/msaccess_vcs_mcp/tools.py` (`_ensure_env_loaded` rewired to emit diagnostic events on every branch; `vcs_get_version_info` exposes `usage_log_path`, `diagnostic_log_path`, `log_code_content`; FastMCP `instructions` text rewritten), `.env.example`, `AGENTS.md`, `tests/test_usage_logging.py` (new diagnostic-stream, body-redaction, and secret-key-mask tests), `tests/test_lazy_workspace_init.py` (extended to assert `lazy_init_*` events).

---

## 2026-04-25 — Honor the add-in lifecycle gate at every call site, with a hard probe timeout

**Trigger**: Commit 6342387 ("fix: surface add-in failures and tighten lifecycle checks") added a strict gate to `VCSAddinIntegration._call_addin_function` requiring both `self._app` AND `self._addin_loaded` to be truthy. The `_addin_loaded` flag is set only by `load_addin()`, which probes the add-in via `Application.Run("…\Version Control.API", "GetVCSVersion")`. However, every call site historically bypassed `load_addin()` and bare-assigned `addin._app = app` to skip a second COM round-trip. The gate refused every call, breaking all 13 user-facing tools plus the `_cleanup_session` atexit handler and `validate_components`. A separate concern surfaced during the fix: if Access is unresponsive (most commonly a developer left VBA paused in the VBE, but also modal dialogs and true hangs), even a "fast" probe will block the MCP server forever on the first tool call.

**Options explored**:
- **Loosen the gate to require only `_app`, restoring the pre-6342387 contract**. Smallest diff. But re-buries lifecycle errors -- a `vcs_run_vba` failure that's actually an add-in load problem still masquerades as a VBA bug, exactly the troubleshooting smell 6342387 was trying to fix.
- **Hybrid: keep the gate, auto-load inside `_call_addin_function` when `_addin_loaded` is False**. Lazy-load is convenient but conflates lifecycle with per-call dispatch and hides the probe cost in unpredictable places. Leaves no clean point to attach a hard timeout.
- **Honor the gate; switch every call site to `addin.load_addin(app, db_path=...)` (chosen)**. Pays one extra `Application.Run` per tool call (typically <10ms once Access has the add-in resident, ~50-200ms on the very first call while it loads). In exchange: lifecycle errors surface at the lifecycle boundary with a single, actionable message, and we have a natural place to attach the hang-protection timeout.

**Hang-protection sub-decision**: For the timeout on `GetVCSVersion`, evaluated `CoCancelCall` (requires server-side cooperation Jet/ACE doesn't implement), killing `MSACCESS.EXE` (loses unsaved work), subprocess isolation (overkill for a sub-second probe), and the disposable-worker-thread pattern that sibling project [`db-inspector-mcp`](C:/Repos/db-inspector-mcp/DECISIONS.md) already evaluated and shipped for DAO query timeouts. The worker-thread pattern is the only practical solution for in-process Access COM. Lifted with light edits: daemon thread + `thread.join(timeout)` + class-level `_active_probe_thread` guard against zombie pile-up + `pythoncom.CoInitialize()` per worker + ROT-based proxy re-acquisition (so the worker has its own apartment-local proxy instead of marshaling back to the main thread, which would defeat the timeout). Configurable via `ACCESS_VCS_PROBE_TIMEOUT_SEC` (default 10s -- short enough to feel responsive, long enough to absorb genuine first-load latency).

**Pre-flight sub-decision**: When `db_path` is supplied to `load_addin()`, run an `os.path.isfile(db_path)` check before any COM activity. Catches stale or typo paths in ~1ms instead of burning the full 10s timeout in ROT lookup for a database that simply isn't there. Defense in depth: today only `tools.py` call sites validate paths via `validate_database_path()`; `validation.py`, `_cleanup_session`, and the internal `get_version_info` did not.

**Instrumentation sub-decision**: Added `log_addin_probe(addin_path, duration_ms, success, timed_out, error)` to `usage_logging.py`, emitting one `"addin_probe"` JSONL event per probe in the same `usage.jsonl` stream as `tool_call` and `code_execution`. Lets us empirically verify the assumed "fast and cheap" cost (`rg '"event":"addin_probe"' logs/usage.jsonl`) and revisit the trade-off if subsequent probes turn out to be expensive in practice. Distinguishes timeouts from generic COM errors (`timed_out` boolean) so analytics can isolate true hangs.

**What this rules out**: Reverting to bare `addin._app = app` at any call site is now an anti-pattern -- it would re-break the gate and lose hang protection. A future contributor who sees the extra `Application.Run` per call should not "optimize" it away without first reading this entry. The probe timeout assumes a well-behaved Access instance can answer `GetVCSVersion` in <10s; if a legitimate use case exists where it cannot, raise `ACCESS_VCS_PROBE_TIMEOUT_SEC` rather than removing the timeout. The class-level `_active_probe_thread` guard is intentionally not cleared on timeout -- a lingering probe thread blocks all subsequent probes with a clear "previous probe still pending" error until Access responds (the daemon thread terminates) or the process exits. This is the desired behavior in production; tests use an autouse fixture to reset it. ROT-based re-acquisition only fires when `db_path` is provided; the `vcs_rebuild_database` and `get_version_info` paths pass `None` and fall back to sharing the main thread's proxy (best-effort, may not respect the timeout as reliably, but acceptable for paths that don't have a database in scope).

**Relevant files**: `src/msaccess_vcs_mcp/addin_integration.py` (new `load_addin(app, db_path=None)` signature with idempotent early return, pre-flight, instrumented probe; new `_probe_with_timeout` and `_find_access_in_rot` helpers; class-level `_active_probe_thread`; `get_version_info` now routes through `load_addin`), `src/msaccess_vcs_mcp/tools.py` (13 call sites updated; redundant inline `GetVCSVersion` pre-flights in `vcs_export_database` and `vcs_import_objects` removed in favor of `load_addin()` with the same friendly error wrapping), `src/msaccess_vcs_mcp/validation.py` (`validate_components` line 255), `src/msaccess_vcs_mcp/main.py` (`_cleanup_session` line 95), `src/msaccess_vcs_mcp/usage_logging.py` (`log_addin_probe`), `.env.example` (`ACCESS_VCS_PROBE_TIMEOUT_SEC` documented), `tests/test_addin_integration.py` (5 new tests: timeout, active-worker guard, db_path pre-flight, idempotency, invalid timeout fallback), `tests/test_usage_logging.py` (4 new tests for `log_addin_probe`).

---

## 2026-04-25 — Drop silent `Application.Eval` fallback in `_call_addin_function`

**Trigger**: Investigating failing add-in integration tests surfaced a long-standing usability and troubleshooting problem: when `Application.Run("…\Version Control.API", funcName)` failed twice in a row, `_call_addin_function` would silently fall through to `self._app.Eval("CallVcsApi(\"" & funcName & "\")")`, using a hard-coded default wrapper name. In any environment where the user had not implemented a `CallVcsApi` VBA function in the open database, this branch would still succeed against unit-test mocks and against any COM object whose `Eval` happened to return a value, producing false-success results for `vcs_export_*` and friends. The original `Run` error was buried in a three-layer retry/fallback message that obscured root cause during troubleshooting.

**Options explored**:
- **Keep current behavior, update tests to mock `Eval` raising**. Lowest-friction, but bakes the silent-success failure mode into production indefinitely.
- **Make the fallback opt-in via `ACCESS_VCS_API_WRAPPER`**. Removes the silent-success default but adds a configuration knob plus a code path almost no one will exercise; documenting *when* to set it requires explaining a real Access COM bug that most users will never hit.
- **Remove the fallback entirely (chosen)**. Single behavioral contract: if `Application.Run` fails twice, a `RuntimeError` is raised with the underlying COM error verbatim. Simpler call graph, fewer places to look during incident triage.

**Decision**: Removed the `Application.Eval` fallback and the `ACCESS_VCS_API_WRAPPER` env-var hook. The single `Run` retry is preserved -- it covers the documented Access first-call add-in load behavior -- but a second failure now surfaces directly as `RuntimeError("Failed to call add-in function '<name>': <error>")`. The decision was driven by simplification of code and of troubleshooting, not by performance or correctness in any narrow sense.

**What this rules out**: The MCP server will no longer auto-recover from genuine `Application.Run`-from-COM bugs by routing through a per-database VBA wrapper. Users who actually need that workaround must reintroduce it deliberately -- ideally as an explicit opt-in setting with a clear error message when it isn't configured -- not as a hidden default. Re-adding any silent fallback that swallows a documented error path should be rejected on review.

**Relevant files**: `src/msaccess_vcs_mcp/addin_integration.py` (`_call_addin_function`).

---

## 2026-04-25 — Multi-strategy `.env` discovery with workspace-roots lazy init

**Trigger**: A user reported that `ACCESS_VCS_ENABLE_LOGGING=true` in a client project's `.env` had no effect when `msaccess-vcs-mcp` was used from another project. Root cause: `_find_project_root()` only walked up from CWD or the installed package location. When the server was launched from a user-level `mcp.json` (`~/.cursor/mcp.json`) and CWD wasn't the project root, the upward walk failed, and the package-location fallback could match `msaccess-vcs-mcp`'s own `pyproject.toml` instead of the user's project. Compounding this, `_load_env_files()` always called `load_dotenv(..., override=False)`, so even a successful reload would silently fail to apply edited values. The sibling `db-inspector-mcp` project had already solved this with a layered resolution strategy.

**Options explored**:
- **Single env-var override (`ACCESS_VCS_PROJECT_DIR`) only**. Simple, works, but requires per-project user configuration. Doesn't help users who configure the server once at the user level and expect it to "just work" across projects.
- **MCP workspace-roots discovery only** (`ctx.session.list_roots()`). Automatic, no user config required. But only works for tools that accept a `Context` parameter and run async — most existing tools are sync without `ctx`.
- **Layered resolution: env-var → workspace roots → CWD walk → package walk → fallback (chosen)**. Mirrors `db-inspector-mcp`'s proven order. Each strategy compensates for the others' blind spots: explicit env-var for power users, workspace roots for IDE-launched servers, CWD walk for terminal-launched servers, package walk as a last-resort dev-install fallback.
- **Eager workspace-roots probe at startup**. Cleaner than lazy init, but FastMCP's `Context` is not available before the first tool call -- the MCP protocol handshake hasn't completed yet.

**Decision**: Ported `db-inspector-mcp`'s discovery machinery in full, with `ACCESS_VCS_` prefix substitution and three project-specific adaptations:

1. **Resolution-method tracking**. `_find_project_root()` and `initialize_from_workspace()` set a module-level `_project_root_method` (one of `RESOLUTION_PROJECT_DIR_ENV`, `RESOLUTION_WORKSPACE_ROOTS`, `RESOLUTION_CWD_ENV`, `RESOLUTION_CWD_MARKER`, `RESOLUTION_PACKAGE_ENV`, `RESOLUTION_PACKAGE_MARKER`, `RESOLUTION_CWD_FALLBACK`). Surfaced via `get_project_root_info()`, printed to stderr ("Resolved project root: X (via Y)"), and embedded in the `logging_initialized` JSONL event so users can audit *which* mechanism actually populated their config in any given session.
2. **mtime-based hot-reload with `override=True` on reload**. `_check_env_reload()` compares stored `.env`/`.env.local` mtimes against current values; when changed, the next `load_config()` triggers a reload using `override=True` so edited values actually replace old ones. Logging is reset only when a reload was detected (was previously reset on every `load_config()` call -- wasteful).
3. **Lazy init wired through the `vcs_tool` decorator, not individual tools**. The decorator inspects the wrapped handler's signature; for async handlers that accept `ctx: Context`, it calls `await _ensure_env_loaded(ctx)` before `load_config()`. Converted `vcs_get_version_info`, `vcs_list_objects`, `vcs_diff_database`, `vcs_import_objects`, and `vcs_rebuild_database` to async with optional `ctx` so the workspace-roots path triggers on the agent's first realistic call regardless of which tool that is.

**What this rules out**: Sync tools without `ctx` cannot trigger workspace-roots lazy init -- but this is fine because once any async-with-ctx tool runs, the project env is loaded for all subsequent calls (sync included). If a future agent goes straight to a sync tool first (e.g., `vcs_execute_sql` before any read tool), workspace-roots discovery won't fire and the user must rely on `ACCESS_VCS_PROJECT_DIR` or CWD-based discovery; converting more tools to async is the escape hatch. The `RESOLUTION_*` constants are part of an implicit public contract -- renaming them would silently break any log-analysis tooling that filters on `project_root_resolution`. The package-walk fallback can still match `msaccess-vcs-mcp`'s own dev tree when the CWD walk finds nothing; this is intentional for development installs and is the lowest-priority strategy.

**Relevant files**: `src/msaccess_vcs_mcp/config.py` (resolution-method tracking, `_check_env_reload`, `_load_env_from_directory`, `initialize_from_workspace`, `get_project_root_info`), `src/msaccess_vcs_mcp/tools.py` (`_lazy_init_attempted`, `_file_uri_to_path`, `_ensure_env_loaded`, `vcs_tool` decorator, 5 tool signatures), `src/msaccess_vcs_mcp/usage_logging.py` (`logging_initialized` event includes `project_root` + `project_root_resolution`), `.env.example`, `README.md` (new "Using msaccess-vcs-mcp from Another Project" section), `tests/test_config_env_loading.py` (17 tests), `tests/test_lazy_workspace_init.py` (9 tests).

---

## 2026-04-15 — Pre-execution audit logging for code execution tools

**Trigger**: The existing usage logging captures tool parameters but truncates all strings to 500 characters and only writes after execution completes. For the three tools that execute arbitrary code against databases (`vcs_execute_sql`, `vcs_call_vba`, `vcs_run_vba`), this leaves two gaps: (1) a complex SQL query or VBA code block may be truncated beyond usefulness in a forensic review, and (2) if the process crashes during execution, no record of what was attempted exists.

**Options explored**:
- **Raise the truncation limit globally**. Simple, but inflates every log entry (file paths, option names, etc.) unnecessarily. Rotation would trigger sooner.
- **Exempt specific parameter names from truncation** (e.g., `sql`, `code`). Mixes audit concerns into the general sanitization logic. Hard to extend cleanly.
- **Dedicated `log_code_execution()` function with a separate event type (chosen)**. Writes a `"code_execution"` event *before* execution begins, with the full untruncated code/SQL and the target database path. The existing `with_logging` decorator continues to write the post-execution `"tool_call"` event with truncated parameters, success/error, and timing. Two complementary records: the audit trail (what was attempted) and the outcome (what happened).

**Decision**: Added `log_code_execution(tool_name, database_path, code, code_type)` to `usage_logging.py`. Called from `vcs_execute_sql` (code_type=`"sql"`), `vcs_run_vba` (code_type=`"vba"`), and `vcs_call_vba` (code_type=`"vba_call"`) immediately after path validation but before any COM/database interaction. The `code` field is never truncated. The event goes to the same `usage.jsonl` file — no separate audit file — distinguished by `"event": "code_execution"`.

**What this rules out**: Code execution entries have no upper size limit on the `code` field. In practice, VBA code blocks and SQL queries are small (under 10 KB). If an agent somehow generates megabyte-scale code strings, log rotation handles it, but this is not a realistic concern. If a separate audit file is ever wanted (e.g., for compliance), the `log_code_execution` function could be retargeted without changing call sites.

**Relevant files**: `usage_logging.py` (`log_code_execution`), `tools.py` (call sites in `vcs_execute_sql`, `vcs_call_vba`, `vcs_run_vba`), `tests/test_usage_logging.py` (5 new tests).

---

## 2026-04-15 — Agents cannot enable McpAllowRunVBA programmatically

**Trigger**: The `vcs_set_option` tool allowed agents to set any VCS option, including `McpAllowRunVBA` which gates arbitrary VBA code execution via `vcs_run_vba`. An agent could autonomously enable this option and then run arbitrary code without user awareness, undermining the security boundary that `McpAllowRunVBA` was designed to provide.

**Options explored**:
- **No guard, rely on default-off**. `McpAllowRunVBA` defaults to False, but nothing prevented an agent from calling `vcs_set_option("db.accdb", "McpAllowRunVBA", True)` as its first action. The docstrings even showed this as an example.
- **Server-side blocklist in `vcs_set_option` (chosen)**. A case-insensitive check against a set of protected option names. Returns a descriptive error directing the user to enable the option manually via the VCS Options form.
- **VBA-side enforcement**. Have the add-in's `SetOption` method refuse `McpAllowRunVBA` when called from MCP. Harder to implement since VBA doesn't know the calling context, and the error would be less clear.

**Decision**: `vcs_set_option` blocks setting `McpAllowRunVBA` with a clear error message. The option requires explicit user consent via the VCS Options form in Access. Docstrings and README updated to stop suggesting agents can self-enable this option.

**What this rules out**: Agents cannot autonomously escalate to arbitrary VBA execution. If future protected options emerge (e.g., a hypothetical `McpAllowDDL`), add them to the `PROTECTED_OPTIONS` set in `vcs_set_option`.

**Relevant files**: `tools.py` (`vcs_set_option`), `README.md` (security section).

---

## 2026-04-15 — Object type normalization lives in VBA, not Python

**Trigger**: `vcs_export_object` and `vcs_import_object` only supported 6 core Access object types (query, form, report, module, table, macro). The add-in's `eDatabaseComponentType` enum defines 24+ types (relations, IMEX specs, VBE project, themes, etc.) that couldn't be exported individually. Additionally, the MCP tools used plural strings (`"queries"`) in `vcs_export_database` but singular (`"query"`) in `vcs_export_object`, creating inconsistency that confused AI agents.

**Options explored**:
- **Python-side normalization map**. A `normalize_object_type()` helper in the MCP server that maps plural/alias forms to canonical singular before passing to VBA. Only benefits MCP callers. Creates a second type map to maintain alongside VBA.
- **VBA-side normalization via `ResolveComponentType` (chosen)**. A `Select Case` function in `modContainers.bas` that accepts singular, plural, and alias forms (50+ strings) and maps to `eDatabaseComponentType`. Benefits all callers — MCP tools, direct `Application.Run` API calls, any future integration. Python becomes a transparent pass-through.
- **Accept both in Python AND VBA**. Redundant and creates maintenance burden keeping two maps in sync.

**Decision**: Type normalization lives entirely in VBA's `ResolveComponentType`. Python passes `object_type` strings through to VBA without validation. VBA returns structured error JSON for unrecognized types. `ExportObject` and `ImportObject` on `clsVersionControl` were extended to handle all 24 component types: core AccessObject types use the existing `ExportSingleObject` path; non-core types use `GetComponentClass` + `GetAllFromDB`. Single-file types (like `vbe_project`) don't require an `object_name` parameter.

**What this rules out**: Adding new component types requires updating VBA's `ResolveComponentType` — the Python MCP layer does not need changes. If a Python-only consumer needs early validation without a COM roundtrip, they would need to maintain their own type list, but this is unlikely since the VBA error response is fast and descriptive.

**Relevant files**: `modContainers.bas` (`ResolveComponentType`), `clsVersionControl.cls` (`ExportObject`, `ImportObject` rewritten), `tools.py` (docstrings updated, `object_name` made optional), `addin_integration.py` (no type-related changes).

---

## 2026-04-15 — Structured JSONL usage logging via composite decorator

> **⚠ Partially superseded** (2026-04-25): The "What this rules out" note about adopting `db-inspector-mcp`'s mtime-based hot-reload pattern "if performance becomes a concern" is now fact -- adopted for correctness rather than performance (the original `override=False` reload was silently failing to apply edits). The `logging_initialized` event also now includes `project_root` and `project_root_resolution` fields. See "Multi-strategy `.env` discovery with workspace-roots lazy init" above.

**Trigger**: Need to troubleshoot and evaluate how AI agents use the MCP tools in practice — which tools are called, with what parameters, how often they fail, and how long they take. The `db-inspector-mcp` sibling project already has a proven logging implementation that was requested as the reference pattern.

**Options explored**:
- **Python stdlib `logging` module**. Standard approach, but produces unstructured text. Not suitable for programmatic analysis of tool call patterns.
- **FastMCP middleware / hooks**. FastMCP doesn't expose a per-tool middleware layer. Would require monkey-patching internals.
- **Composite decorator with JSONL file logging (chosen)**. Follows the exact pattern from `db-inspector-mcp`: a `with_logging(name)` decorator that wraps each tool, writing one JSON object per line to a rotating file. A `vcs_tool("name")` composite decorator chains config reload → usage logging → `mcp.tool()` registration, replacing bare `@mcp.tool()` on all 17 tools. Controlled by `ACCESS_VCS_ENABLE_LOGGING` env var (default: off).

**Decision**: Adopted the `db-inspector-mcp` pattern with project-specific adaptations:
- Env var prefix changed from `DB_MCP_` to `ACCESS_VCS_` for consistency with existing config.
- Removed `database`/`dialect` fields from log entries (VCS tools pass `database_path` as a regular parameter, so it's captured in `parameters` automatically).
- Error pattern categories tailored to VCS-specific errors (COM errors, add-in errors, VBA compile errors, write-disabled, database busy) instead of SQL-specific patterns.
- `Context` objects from FastMCP are filtered out of logged parameters (not serializable).
- Lazy initialization: disabled state is not cached (`_logging_enabled` stays `None`) so the first tool call after `load_config()` populates the env can still enable logging. Failure state _is_ cached to avoid retry spam.

**What this rules out**: Log entries do not include tool return values — only parameters, success/failure, errors, and timing. If result logging is needed later, the `with_logging` decorator already receives the result for serialization checking and could be extended. The `vcs_tool` decorator calls `load_config()` on every tool invocation (one `stat()` call); if this becomes a performance concern, the hot-reload pattern from `db-inspector-mcp` (mtime-based gating) could be adopted.

**Relevant files**: `src/msaccess_vcs_mcp/usage_logging.py` (new), `src/msaccess_vcs_mcp/tools.py` (`vcs_tool` decorator, all 17 tools migrated), `src/msaccess_vcs_mcp/config.py` (logging env vars + `reset_logging` on reload), `src/msaccess_vcs_mcp/main.py` (startup status), `.env.example`, `tests/test_usage_logging.py` (36 tests), `AGENTS.md` (new), `.gitignore` (`logs/`).

---

## 2026-04-14 — Tool naming: `vcs_*` prefix

**Trigger**: Tools were originally named `access_*` (e.g., `access_export_database`). A separate MCP server for Access (`MCP-Access` by bclothier) also uses `access_*` prefixed tools. Both servers might be loaded in the same agent session. Additionally, these tools control the VCS add-in, not the Access application itself, so `access_*` was a misnomer.

**Options explored**:
- **Keep `access_*`**. Familiar, but inaccurate and collides with MCP-Access.
- **Use `vcs_*` (chosen)**. Accurately reflects these tools control the VCS add-in. No namespace collision.
- **Use `msaccess_vcs_*`**. Unambiguous but verbose — wastes tokens on every tool call.

**Decision**: All 9 existing tools renamed from `access_*` to `vcs_*`. All 8 new tools use `vcs_*`. Updated across `tools.py`, `README.md`, `validation.py`, and all docs.

**What this rules out**: Any external references to `access_export_database` etc. break. Acceptable since the server is pre-release with no external consumers.

**Relevant files**: `tools.py`, `README.md`, `validation.py`, `docs/*.md`.

---

## 2026-04-14 — Eight new tools for per-object development workflow

> **⚠ Partially superseded** (2026-04-15): `vcs_export_object` and `vcs_import_object` now support all 24 component types (not just the original 6 core types). Type normalization moved to VBA. See "Object type normalization lives in VBA, not Python" above.

**Trigger**: All existing tools operated at the whole-database level. Agents had no way to export/import a single object, execute SQL, run VBA, or control add-in options — capabilities essential for the tight edit-import-compile-test loop needed during add-in development and general database development.

**Options explored**:
- **Extend existing tools with filters** (e.g., `vcs_export_database` with `object_name` parameter). Conflates bulk and single-object semantics. The existing tools have async/callback infrastructure not needed for quick per-object calls.
- **Build intelligence into the MCP server** (Python-side VBE manipulation, SQL execution via separate ODBC connection). Creates tight coupling to Access internals in Python, duplicates logic better handled in VBA, and opens a second database connection causing file-locking conflicts.
- **Thin MCP tools that delegate to add-in API methods (chosen)**. Each new tool calls a corresponding method on `clsVersionControl` via the existing `API()` dispatcher. The add-in handles all Access interaction. The MCP layer just validates paths, parses JSON results, and formats responses.

**Decision**: 8 new tools added: `vcs_export_object`, `vcs_import_object`, `vcs_execute_sql`, `vcs_call_vba`, `vcs_run_vba`, `vcs_set_option`, `vcs_get_option`, `vcs_get_log`. Total: 17 tools. Each delegates to a public method on `clsVersionControl` in the add-in. The MCP server remains a lightweight wrapper — all business logic lives in VBA.

**What this rules out**: The MCP server does not do database introspection, schema analysis, or complex SQL. Those capabilities stay in `db-inspector-mcp`. If the VCS MCP needs to support operations not expressible through `clsVersionControl` API methods, the add-in must be extended first. This is intentional — it keeps the MCP layer thin and ensures all consumers of the add-in API get the same capabilities.

**Relevant files**: `tools.py` (8 new tool definitions), `addin_integration.py` (`call_sync` used by all new tools).

---

## 2026-04-14 — SQL execution via add-in DAO connection, not separate ODBC

**Trigger**: Usage logs from `db-inspector-mcp` showed 67% of all calls were just running SELECT queries (`db_count_query_results`, `db_preview`). Agents frequently need to inspect `MSysObjects`, `MSysQueries`, table data, and query results. Requiring a second MCP server for this basic need adds configuration overhead and creates a second connection to the same Access file (risking file-locking conflicts).

**Options explored**:
- **Keep SQL execution in db-inspector-mcp only**. Clean separation, but requires agents to have both MCPs configured. Two connections to the same `.accdb` file can cause locking issues. Extra overhead for the dominant use case.
- **Add ODBC connection in the VCS MCP server** (Python-side). Avoids VBA roundtrip but opens a second connection. Would need to handle Access SQL dialect quirks in Python.
- **Route through add-in's existing DAO connection (chosen)**. `ExecuteSQL` method on `clsVersionControl` uses `CurrentDb.OpenRecordset` — the same connection the add-in already holds. No file-locking conflict. Access SQL dialect handled natively. Read-only (SELECT only, enforced in VBA).

**Decision**: `vcs_execute_sql` tool calls the add-in's `ExecuteSQL` API method, which runs the query via `CurrentDb.OpenRecordset`, serializes results as JSON, and returns them. Non-SELECT statements are rejected. Results capped at `max_rows` (default 100). The db-inspector MCP remains available for cross-database comparison and heavy analytical work.

**What this rules out**: No write queries (INSERT/UPDATE/DELETE/DDL) through this tool. If agents need to modify data, they use `vcs_run_vba` or `vcs_call_vba` with appropriate VBA code. The SQL validation is simple (checks for `SELECT` prefix) — a determined agent could bypass it via `vcs_run_vba`, which is why `McpAllowRunVBA` defaults to off.

**Relevant files**: `tools.py` (`vcs_execute_sql`), `clsVersionControl.cls` (`ExecuteSQL`).

---

## 2026-04-14 — Two VBA execution tools with distinct roles

**Trigger**: Agents need to execute VBA code for testing and debugging. Two distinct use cases emerged: calling existing functions by name (safe, predictable) and executing arbitrary agent-generated code (powerful, risky). These have fundamentally different security profiles.

**Options explored**:
- **Single `vcs_run_vba` tool for both**. Agent passes either a function name or a code block, tool detects which. Blurs the security boundary — how do you gate "arbitrary code" while allowing "call existing function"?
- **Two tools with distinct roles (chosen)**. `vcs_call_vba` calls existing named functions via `Application.Run` — no temp module, no compilation, lower risk, separate permission (`McpAllowCallVBA`, default: True). `vcs_run_vba` executes agent-generated code via temp module lifecycle — compilation check, error capture, cleanup, higher risk, separate permission (`McpAllowRunVBA`, default: False).

**Decision**: `vcs_call_vba(database, function_name, args)` for existing functions; `vcs_run_vba(database, code)` for ad-hoc code. The add-in handles `RunVBA`'s full lifecycle (create, compile, execute, capture, cleanup). Error capture in `RunVBA` uses module-level variables with accessor functions rather than embedded JSON string construction in generated code — cleaner and avoids VBA quote-escaping nightmares.

**What this rules out**: `vcs_call_vba` is limited to public functions callable via `Application.Run` (max 3 args in current implementation). Private functions or functions requiring object parameters can't be called directly — use `vcs_run_vba` for those. If the 3-arg limit becomes a problem, the `InvokeTypes` + `pythoncom.Missing` padding pattern (used by MCP-Access) would support up to 30 args.

**Relevant files**: `tools.py` (`vcs_call_vba`, `vcs_run_vba`), `clsVersionControl.cls` (`RunVBA`).

---

## 2026-04-15 — Session-scoped option overrides for MCP/API callers

**Trigger**: `vcs_set_option` changes were silently discarded because the add-in reloads options from `vcs-options.json` at the start of every operation. The agent's overrides never persisted past the first subsequent export/build.

**Decision**: The MCP server generates a session ID at startup (`uuid4().hex[:8]`), registers it with the add-in via `RegisterSession`, and `vcs_set_option` now writes overrides to a session-scoped file (`mcp/options-{session_id}.json`) in the export folder. The add-in's operation entry points load these overrides after `LoadProjectOptions` when `Operation.Source` is API/MCP. On shutdown, `atexit` calls `EndSession` to clean up the file. A `vcs_end_session` tool is also available for explicit mid-session cleanup.

**What this rules out**: Session IDs don't persist across server restarts — the agent must re-set options if the server restarts. Stale override files are auto-cleaned after 30 days on the add-in side.

**Relevant files**: `tools.py` (`vcs_set_option`, `vcs_end_session`), `main.py` (session ID, atexit), `config.py` (`get_session_id`). Add-in side: see `DECISIONS.md` in `msaccess-vcs-addin`.
