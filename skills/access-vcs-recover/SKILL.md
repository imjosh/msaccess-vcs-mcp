---
name: access-vcs-recover
description: Recover a user's Access VCS MCP workflow after server_busy, a timeout, a blocked dialog, uncertain completion, interruption, or cleanup failure. Use before retrying a mutation or when requesting cancellation; excludes rebuilding the VCS add-in itself.
---

# Access VCS recovery

Follow [the MCP release policy](../release-policy.md) before retrying database
changes. Keep permitted inspection and recovery available after a refusal.

1. Preserve the original result, database/source paths, call time, operation ID,
   and any log path. Read live tool descriptions and schemas; inspect the target
   project's exported `AGENTS.md` and troubleshooting reference when available.
2. Read [reconcile.md](references/reconcile.md). Inspect recent calls and the
   matching operation log before retrying a mutation. Determine whether this call
   was refused, is still running, completed, failed, or remains uncertain.
3. Check `vcs_automation_status` on the intended instance. For a blocking dialog,
   break mode, ambiguous identity, or a cancellation request, read
   [dialogs.md](references/dialogs.md) before acting. Inspect first and apply only
   recovery actions covered by the user's request or existing authorization.
4. Recheck readiness and reconcile the original outcome with current database and
   source state. Readiness proves responsiveness, not successful completion.
   Retry only after accounting for prior effects and addressing the cause; retain
   the original scope and authorization. If evidence stays inconclusive, report
   the uncertainty and the specific user action or diagnostic needed.

Preserve user-owned Access windows. Use server-supported recovery rather than
killing Access, raw COM cleanup, keystrokes, blanket warning suppression, or
changing protections to get past a refusal. Recovery does not expand permission
to overwrite data, select a conflict winner, or rebuild a database.
