# Dialog recovery: fixes from the review of `e982918`

Status: ready-for-agent

## Problem Statement

The dialog-recovery feature exists so an agent can get Access unstuck when a modal dialog blocks the Access gate. In its first cut, it can make things worse:

- Calling a dialog tool can freeze the whole MCP server. The tools skip the gate, but they do blocking Win32 work directly on the server's event loop. A dialog that hangs mid-click hangs the server with it, and every other Cursor window gets no response, including progress for the call the agent was trying to rescue.
- An agent that imports with the default `decision_policy="block"` can be told the merge succeeded when the add-in actually stopped at an uncovered conflict. Only one of the import paths reports `decision_required`.
- The server can click a button in a process it hasn't proved is the right Access instance. That happens when the creation time can't be read, when the process name is unknown, or when the window handle it recorded has since been reused. The project's rule is that every ambiguity resolves toward leaving the window alone, and this code resolves ambiguity toward acting.
- `vcs_automation_status` can report `ready: true` for a process that is dead or was never found. An agent trusts that and makes a call that will fail.
- `execution_interrupted` persists forever on a PID, so a later, unrelated process that reuses that PID inherits it. The call that was interrupted is never marked failed by the server.
- `vcs_recover_dialogs` waits about 50 ms after clicking instead of up to its timeout, so it reports dialogs as still open when they are just slow to close. It never returns `dismiss_uncertain`.
- Target resolution counts any window whose title contains the database's file name. That includes editors and browsers, so an agent with exactly one Access instance open gets `ambiguous_instance`.
- Automatic recovery treats every standard dialog box in the Access process as a recognised VBA message box, so the `safe` policy clicks OK on dialogs it doesn't actually recognise.

## Solution

From the agent's point of view, the four dialog tools stay callable, and the Access gate stays free, whatever state Access is in. Nothing about them can stall other requests. The dialog tools only click when three things are confirmed:

- The process is the named Access instance, by PID and creation time.
- The dialog is one they positively recognise.
- The window is still the one they inspected.

When any of those can't be confirmed, they report and do nothing. Every import path reports `decision_required` the same way. `ready: true` means what the documentation says. Ending a runtime error marks the interrupted call as failed and doesn't leak onto other processes.

## User Stories

1. As an AI agent, I want `vcs_list_dialogs` to return promptly while another Cursor window's call holds the Access gate, so that I can see what is blocking it.
2. As an AI agent in a second Cursor window, I want my unrelated tool calls to keep getting `server_busy` responses (or results) while another agent inspects dialogs, so that dialog inspection never looks like a server hang.
3. As an AI agent, I want in-flight operations to keep streaming progress while I inspect or dismiss a dialog, so that I don't mistake a working build for a stuck one.
4. As an AI agent, I want a click on a dialog whose thread has hung to come back within the dialog timeout, so that one hung dialog can't hang the MCP server.
5. As an AI agent, I want a click that couldn't be delivered in time reported as `dismiss_uncertain`, so that I know the dialog's state is unknown and nothing was retried.
6. As an AI agent importing with the default decision policy, I want every import path to return `error_pattern: decision_required` when the add-in stopped at an uncovered prompt. That covers the async, inline-sync, sync-fallback, no-callback and category-scoped paths, so that I never report a partial merge as success.
7. As an AI agent, I want the add-in's `decisions` payload returned on every import path that has one, so that I can tell the user which objects need a choice.
8. As an AI agent, I want the original import error returned when the add-in's cleanup of the operation policy also fails, so that the real cause isn't hidden.
9. As a maintainer, I want a cleanup failure after an import attached to the result as secondary information and written to the usage log, so that a policy left set in the add-in is visible.
10. As a user with Access open, I want the server to refuse to click in a process whose creation time it can't read, so that a reused PID can never get a stranger's dialog dismissed.
11. As a user, I want the server to refuse to click in a process whose executable name it can't confirm as Access, so that a failed process query isn't treated as permission.
12. As an AI agent, I want refusals caused by unconfirmed identity to carry a distinct `error_pattern`, so that I can tell "not Access" apart from "couldn't tell".
13. As an AI agent that didn't pass `create_time`, I want the server to record the creation time it observed and check it again before each click, so that the process can't change between inspection and click.
14. As an AI agent, I want the server to re-check, immediately before clicking, that the dialog and button handles I chose still belong to the same process and dialog, so that a closed dialog's reused handle isn't clicked.
15. As an AI agent, I want a changed or vanished target reported (for example as `dialog_changed`) with no click performed, so that I can re-inspect and decide again.
16. As an AI agent with one Access instance open, I want target resolution to consider only Access processes, so that an editor tab with the database's name doesn't make the target ambiguous.
17. As an AI agent with two Access instances holding databases of the same name, I want `ambiguous_instance` and no click, so that I pass `pid` explicitly.
18. As an AI agent, I want `vcs_automation_status` to return `ready: false` when the process is gone or has no windows, so that I don't send a call to a dead instance.
19. As an AI agent, I want `ready: true` only when the process is confirmed alive and responsive, VBA isn't in break mode, no blocking dialog is open, and the gate is free, so that the flag means exactly what the documentation says.
20. As an AI agent, I want a distinct `error_pattern` for "Access instance not running" and for "no windows to probe", so that I know whether to relaunch or wait.
21. As an AI agent, I want `vcs_recover_dialogs` to keep checking until the dialogs it clicked have closed or the timeout elapses, so that a slow close isn't reported as still open.
22. As an AI agent, I want `vcs_recover_dialogs` to stop early when only dialogs its policy doesn't cover remain, so that I'm not made to wait out the timeout for nothing.
23. As an AI agent, I want `vcs_recover_dialogs` to report `dismiss_uncertain` when a dialog it clicked is still open at the deadline, the same way `vcs_dismiss_dialog` does.
24. As an AI agent, I want the `safe` policy to click OK only on dialogs positively matched to a known kind, so that an unrecognised OK-only dialog is reported rather than dismissed.
25. As an AI agent, I want unrecognised dialogs reported with kind `unknown` along with their title, text and buttons, so that I can describe them to the user or dismiss one explicitly.
26. As an AI agent, I want dismissing a runtime or compile error to record an interruption whichever tool did the click, so that `vcs_dismiss_dialog` and `vcs_recover_dialogs` agree.
27. As an AI agent whose gated call was interrupted by ending a runtime error, I want that call's result to come back with `success: false` and `execution_interrupted: true`, so that the failure is recorded even if the COM error was swallowed.
28. As an AI agent, I want an interruption record tied to one process identity (PID plus creation time), so that a new process that reuses the PID starts clean.
29. As an AI agent, I want an interruption record used up by the gated call it interrupted, so that `execution_interrupted` doesn't stick to later, healthy calls.
30. As an AI agent, I want status tools to show `last_interruption` only while that record is still waiting to be used up, so that stale history doesn't read as current state.
31. As an AI agent, I want `action="close"` on a finished add-in window refused only while the gate is busy with that same database, so that an operation on another database doesn't block closing this one's window.
32. As an AI agent, I want `noninteractive=False` on `vcs_run_tests` and `vcs_import_objects` to explicitly select interactive mode, so that the behaviour doesn't depend on whatever mode the add-in was last left in.
33. As an operator, I want `ACCESS_VCS_DIALOG_TIMEOUT_SEC` picked up when `.env` changes, like every other setting, so that I don't need to restart the server to tune it.
34. As an operator, I want `ACCESS_VCS_DIALOG_TIMEOUT_SEC` listed in the AGENTS.md configuration section and in `.env.example`, so that I can find it.
35. As a maintainer, I want `vcs_dismiss_dialog` and `vcs_recover_dialogs` to share one sequence (resolve target, verify identity, act, wait, re-inspect), so that fixes to one apply to both.
36. As a maintainer, I want dialog IDs built and parsed in one place, so that the `hwnd:<n>` format can't drift between producer and consumer.
37. As a maintainer, I want comments that no longer match the code removed or corrected, so that the next reader isn't misled about the wait loop, interruption timing, or process-wide mode.
38. As a maintainer, I want parameters and response keys that never vary removed. Today these are the always-true `alive`, the `remaining` key that duplicates `dialogs`, `max_dialogs` (never passed) and `owner_hwnd` (collected, never used). That way the response shape only carries information.
39. As a maintainer, I want the undocumented `inspect` alias for `policy="report"` removed or documented, so that the recovery vocabulary is exactly what `docs/DIALOGS.md` lists.
40. As a user watching Access, I want no keystrokes, Enter, Escape or coordinate clicks introduced by any of these fixes, so that input only ever goes to a specific button window.
41. As a user, I want Debug never clicked, and save/discard dialogs never clicked automatically, as today, so that these fixes don't loosen what is safe.

## Implementation Decisions

**Tool dispatch for gate-exempt tools**

- Gate-exempt synchronous tools run in a worker thread from the tool wrapper, not on the event loop. They must not use the COM apartment thread, because that is exactly what a blocking dialog or a gated call is occupying. Win32 window enumeration and messaging don't need COM.
- Every tool in the exempt set gets this treatment, not only the four dialog tools, so a future exempt tool can't reintroduce the problem.
- The worker thread gets a hard ceiling slightly above the dialog timeout. If that ceiling is reached, the tool returns a recoverable timeout result and gives up on the thread. That mirrors how `vcs_run_vba` treats a hung worker.

**Window backend interface (the Win32 boundary)**

- The window backend interface gains a process-identity query. It returns the executable name and creation time for a PID, or "unknown" when either can't be read, plus whether the process is still running. The dialog logic stops calling the process registry helpers directly, so everything about the outside world goes through the backend.
- Clicking becomes bounded. The backend's click sends `BM_CLICK` with a timeout that stops waiting if the target thread is hung, and reports whether the message was handled in time. An undelivered click is a `dismiss_uncertain` outcome and is never retried.
- The real backend's click refuses a button handle whose owning process no longer matches the expected PID. This is checked at the lowest level, as a second line of defence behind the re-verification below.

**Process identity**

- A process may be acted on only when its name is confirmed as Access **and** its creation time is readable. If the caller passed `create_time`, it must match. Unknown name or unknown creation time means inspection may still report, but every click and close is refused with a new `error_pattern` for unconfirmed identity, distinct from `not_access_process` and `process_identity_mismatch`.
- When the caller didn't pass `create_time`, the observed creation time is fixed at resolution. Everything after that in the same call is checked against it.
- Immediately before each click or close, the server lists windows again and confirms three things: the dialog handle still exists, it still belongs to the same PID and creation time, and the button handle is still one of that dialog's buttons. Any mismatch returns `dialog_changed` with no action taken.

**Target resolution**

- Candidates are limited to windows owned by confirmed Access processes before title matching and counting. Ambiguity is decided only among Access instances.

**Readiness**

- `ready` requires all of the following. If any is missing, `ready` is false with a specific `error_pattern`:
  - confirmed identity
  - process running
  - responsiveness probed and true (unknown counts as not ready)
  - no break mode
  - no blocking dialog
  - gate free for that database
- "Process not running" and "no windows to probe" are separate outcomes.

**Classification and automatic recovery**

- A dialog gets a known kind only through a positive signature: a recognised caption, text or button set. It never gets one by elimination. Any standard dialog box that matches no signature is `unknown`.
- Exception, kind `vba_msgbox`: a standard dialog of any caption whose only actionable button is OK (Help ignored) is a VBA MsgBox and gets that kind. The button set is its positive signature. A custom-caption box with two or more buttons, or one non-OK button, stays `unknown`.
- The `safe` policy clicks only dialogs with a known kind (including `vba_msgbox`), and within those only OK-only ones without destructive text. A `vba_msgbox` with destructive text is reported, not clicked. `unknown` is always reported, never clicked automatically. Debug, save/discard, keystrokes and coordinate clicks remain excluded.
- Dismissing any dialog classified as a failure (runtime or compile error) records an interruption, whether the dismissal came from the explicit or the automatic path.

**Waiting after an action**

- `vcs_dismiss_dialog` and `vcs_recover_dialogs` share one post-action sequence. It polls until no dialog the call acted on remains, or until only dialogs the policy doesn't cover remain, or until the dialog timeout. It then re-inspects once for the returned report. A dialog the call acted on that is still open at the deadline yields `dismiss_uncertain`.

**Interruptions**

- Interruption records are keyed by process identity (PID plus creation time), not PID alone.
- A record also carries the gated call that was in flight when it was made, named the same way the gate's `busy_with` names it. When that gated call finishes, the tool wrapper checks for a matching record. If it finds one, it forces `success: false`, `execution_interrupted: true` and `error_pattern: execution_interrupted`, keeping the original error text if there was one. Then it removes the record.
- A record with no in-flight call to attach to, because the gate was free, is reported by status tools as `last_interruption` until the process identity changes or the next gated call on that database starts.

**Gate checks for `close`**

- `action="close"` on an add-in window is refused only when the gate's in-flight call is on the same database, compared after normalising paths.

**Import decision results**

- Add-in integration's merge build returns the add-in's parsed result, including `success`, `decision_required`, `decisions` and message. It no longer returns a hard-coded success.
- All import paths pass their final add-in result through the same decision normalisation already used for async completions. That includes the category-scoped sync import and the inline result carried by the add-in's `sync` marker. No path returns success without going through it.
- If clearing the operation policy fails during cleanup, that failure never replaces the operation's own error. It is attached as a secondary field and logged.

**Interaction mode**

- `noninteractive=False` sends an explicit interactive mode to the add-in rather than no argument. The documentation drops the word "previous" and describes the flag as selecting interactive mode.

**Configuration**

- The dialog timeout is read through the configuration layer, so the `.env` reload applies to it. Its default and ceiling are unchanged.

**Tidy-ups included because they touch behaviour**

- Dialog ID formatting and parsing live in one place.
- Always-constant fields and parameters are removed, and the `inspect` alias is removed. It is undocumented and new in `e982918`, so no caller depends on it.
- The stale comments noted in the review are corrected or deleted.

## Testing Decisions

- **One seam: the MCP tool boundary.** Tests call the `vcs_*` tools through the tool wrapper and assert only on two things: the returned dictionary, and what the fakes recorded (clicks, closes, add-in calls). Tests don't import or assert on internal helpers of the dialog-recovery module. The existing tests that target those internals move to the tool boundary.
- **Two test doubles behind the seam:**
  - The existing fake window backend, installed where the default backend is created. It is extended with scriptable process identity (name, creation time, running, each possibly unknown) and scriptable click behaviour: delivered, never delivered, or blocking until released. It can also change its window list between calls, to simulate reused handles and slow closes.
  - A mocked add-in integration, in the style of the existing import-tool patching helper, returning scripted results for merge build, the `sync` marker, category-scoped import, and policy set and clear.
- **What makes a good test here:** each user story above has at least one test, stated as an agent-visible outcome. For example: "an unconfirmed identity returns the unconfirmed-identity pattern and nothing is clicked", or "every import path returns `decision_required` when the add-in says so". Tests of the form "function X calls Y" are not wanted.
- **Event loop test:** a fake backend whose window listing blocks until released. While a dialog tool is in flight, a second coroutine on the same loop must make progress, and a gated tool call must get its normal response. The dialog tool then completes after release.
- **Hung click test:** a click that never returns. The tool returns `dismiss_uncertain` within the dialog timeout plus a small margin.
- **Interruption tests:**
  - A gated call is held open, a runtime error is ended through `vcs_dismiss_dialog`, and the held call's result comes back failed with `execution_interrupted`.
  - A new process identity with the same PID shows no interruption.
- **Import tests:** one test per import path, all asserting `decision_required`, plus one where both the operation and the policy cleanup raise.
- **Prior art:**
  - The existing dialog-recovery tests (the fake backend, window and button builders, and the "dialog tools run while the gate is held" test).
  - The scoped-sync import tests (patching the tool's collaborators).
  - The run-tests tool tests.
- **Live tests** against real Access stay marked `integration` and aren't the verification for these fixes. Add one live check for the bounded click only if it can be made reliable.

## Out of Scope

- Changes to the add-in's VBA, beyond confirming what it already returns (see Further Notes).
- UI Automation, or any input method other than `BM_CLICK` and `WM_CLOSE` on specific window handles.
- Splitting the dialog-recovery module by concern, or introducing a target type to bundle the parameters that travel together. Both were flagged as smells in the review; they're worth doing but aren't needed for these fixes.
- Rewriting the destructive-text heuristics beyond making "known kind" a positive match.
- Changing which dialogs are ever safe to click: Debug stays forbidden, and save/discard stays manual.
- Configuring an issue tracker for this repo.

## Further Notes

- This spec comes from a two-axis review of commit `e982918` ("inspect and dismiss Access dialogs off the COM gate"). The spec that commit implemented is its own `DECISIONS.md` entry dated 2026-09-29 and `docs/DIALOGS.md`. This spec doesn't change that decision; it makes the code meet it.
- **Add-in contract assumption.** This spec assumes the add-in's synchronous `MergeBuild` and `ImportByType` return values report `decision_required` and `decisions` the same way its async completion callback does, and that the `sync` marker carries the inline result. Confirm this in the add-in repository before implementing the import-path changes. If it isn't true, the add-in needs a matching change, which is out of scope here.
- The `DECISIONS.md` entry should get a short follow-up entry recording four rules: the identity rule (unconfirmed means no action), the "known kind by positive signature only" rule, how interruptions are used up, and the move of gate-exempt tools off the event loop.
- Publishing: the repo has no configured issue tracker and no `ready-for-agent` label, so this spec lives in the repo for now. Running `/setup-matt-pocock-skills` would configure one.
