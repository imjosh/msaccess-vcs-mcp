# Decisions and existing authorization

Choose a policy whose effect fits the user's intent for this scope. A request to
apply edited source can authorize replacement of those objects; it does not
automatically settle unrelated concurrent edits or all conflicts in a full merge.
Reuse explicit instructions already given rather than asking again. If the
winner is unspecified, `block` preserves the unresolved choice for the user.

For tools that expose these parameters, retain noninteractive execution unless
the user intends to handle prompts in Access. Read the live schema rather than
adding these parameters to tools that do not support them.

| Policy | Use when | Effect |
| --- | --- | --- |
| `block` | A winner or confirmation is not yet authorized | Uncovered choices return `decision_required`; OK-only add-in notices can be acknowledged. |
| `prefer_source` | The user authorized source to win conflicts throughout the selected scope | Takes each conflict's requested action, including deletion when requested; overwrites from source when no action is specified. |
| `prefer_database` / `skip` | Keep live database versions on conflicts | Keeps the database object and skips that source change. |
| `decline` | The user wants confirmations declined | Answers No, Cancel, or Abort; conflicts keep the database object. |

Conflict policies do not approve generic confirmations such as overwrite prompts.
`prefer_source` is not blanket approval. Policies also do not restore checks that
single-object import or category `full_import=True` bypass; select the operation
as well as the policy.

## When `decision_required` returns

Read and preserve `decisions`, affected object names, proposed actions, and any
`runtime_error`, log path, or cleanup error. Explain the concrete choice and its
scope. Apply an existing instruction only if it covers those actions; otherwise
ask the user which version/action to keep. Report the sync as blocked or partial,
not completed. Earlier work in the operation may already have occurred.

Before retrying with the resolved policy, inspect this attempt's log and current
state so completed work is not blindly repeated. Narrow the next operation if
the authorization covers fewer objects, while retaining needed conflict checks.
Source-file Git conflict markers require a source edit before importing again.

`noninteractive=False` selects normal prompts on supported import/export tools;
use it when the user will handle the choices in Access, not to evade a policy.
Automation test runs are always headless and reject that setting.

`policy_unconfirmed` or `interaction_mode_unconfirmed` means setup was not
confirmed and nothing started. Obtain a compatible add-in rather than assuming
the version string proves capability. `interaction_mode_refused` can mean another
caller's scope is active; let its owner release it. `policy_cleanup_error` is
secondary to the operation result and requires recovery before relying on restored
interaction state. Clearing another caller's scope is not conflict resolution.
