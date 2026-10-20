# Decisions and existing authorization

Choose a policy whose effect fits the user's intent for this scope. A request to
apply edited source can authorize replacement of those objects; it does not
automatically settle unrelated concurrent edits or all conflicts in a full merge.
Reuse explicit instructions already given rather than asking again. If the
winner is unspecified, `block` preserves the unresolved choice for the user.

For tools that expose these parameters, retain noninteractive execution unless
the user intends to handle prompts in Access. Read the live schema rather than
adding these parameters to tools that do not support them.

Read [the direction-aware policy table](../../../docs/DIALOGS.md#conflict-outcomes).
`prefer_source` imports source on conflicting import and preserves source on
conflicting export. `prefer_database` preserves the database on import and exports
database content on export. `skip` preserves the conflicting destination in either
direction. `decline` preserves conflicts and answers No, Cancel, or Abort.

No policy authorizes conflicting edited deletions in this stage. A preference
policy can still return `decision_required`. Single-object calls enforce their
request policy. `full_import=True` retains explicit policy protection.

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

Each operation owns its policy through completion and releases it itself.
MCP sends no persistent setup or cleanup. A refused request cannot change another
operation's policy, mode, or journal. Preserve unfamiliar journal fields and
resolution values, and inspect `skipped` and `transferred` before proceeding.
Never automatically retry with an overwrite policy or remove the policy to bypass
a blocked deletion. An incompatible API requires a matching repaired pair.
