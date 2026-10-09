---
name: access-vcs-sync
description: Synchronize a user's Microsoft Access database and VCS-exported source through MCP when exporting Access edits, importing source edits, merging changes, or building a fresh database. Includes verification and relevant user-project tests; excludes developing or rebuilding the VCS add-in itself.
---

# Access VCS sync

Before database operations, follow [the MCP release policy](../release-policy.md).

1. Identify the intended database and export folder. Locate and read the target
   project's exported `AGENTS.md`, normally `<database-name>.accdb.src/AGENTS.md`,
   and its references for the object types being edited. Follow project-specific
   guidance too. Read live MCP tool descriptions and schemas before choosing
   parameters; server/add-in versions can differ from these workflows.
2. Read [operation-choice.md](references/operation-choice.md). Choose the direction
   and smallest adequate scope from the user's requested outcome. Confirm the
   configured export folder and account for replacement, deletion, and backups
   before dispatch. Preserve source edits before any export.
3. For a tool with decision policies, read
   [decisions.md](references/decisions.md). Apply the user's stated intent and
   existing authorization; ask only for an unresolved choice. A blocked decision
   does not authorize another tool that bypasses conflict detection.
4. Dispatch the selected operation once. A launch acknowledgment is not completion.
   For a timeout, `server_busy`, dialog, or `completion_unconfirmed`, reconcile
   recent calls, operation logs, and readiness before any retry. Use
   `access-vcs-recover` when available; otherwise stop further mutations until
   the previous call's outcome and target state are known.
5. Read [verification.md](references/verification.md). Verify the requested changes
   in the intended database, run relevant user-project tests when needed, and
   report applied changes, skips, evidence, and material limitations. Source edits
   alone or a passing test in a different database do not complete a sync.

If MCP is unavailable or the operation lacks authorization, prepare the source
changes and give the user the appropriate add-in ribbon workflow. Editing an
exported project does not require installing this skill.
