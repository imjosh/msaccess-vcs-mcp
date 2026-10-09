# Choose the operation and target

## Establish the project

Use an absolute database path and the actual export folder, not the installed
add-in or a convenient database already open. The `.accdb.src` naming convention
is a starting point; a project may configure another folder. Read its exported
`AGENTS.md` before editing. If absent, locate the project's existing export and
resolve the missing guidance; exporting merely to obtain docs can overwrite edits.

To confirm the configured folder without exporting, use `vcs_call_vba` against
the intended database, `function_name="VCS.API"`, `args=["GetExportFolder"]`.
Check the tool result and returned value against the source tree being edited.
This read does not authorize changing the project's export-folder option.
`source_dir` on import is validated and used for log lookup; the import itself
reads the configured folder. A passed path alone cannot retarget it.

Inspect pending source changes and live database changes relevant to the request.
Preserve edited source and unsaved Access work before a replacement or export.
The installed add-in is a library and is refused as a target; changing extensions
or switching to raw VBA is not a workaround for a refusal.

## Select by outcome

| Requested outcome | Tool and scope | Decision that matters |
| --- | --- | --- |
| Capture changes made in Access | `vcs_export_object` for one name; `vcs_export_database` for categories or the project | Export writes source, so protect pending source edits. |
| Replace one named object from edited source | `vcs_import_object` | Direct replacement bypasses index-based conflict detection; establish source-replacement intent first. |
| Apply changed source in complete categories | `vcs_import_objects` with `object_types`, normally `full_import=False` | Category scope reconciles orphans and takes no database backup. |
| Integrate project-wide source changes into the existing database | `vcs_import_objects` without `object_types` | Full merge checks conflicts and takes a database backup. |
| Create a fresh database from source | `vcs_rebuild_database` with source and an explicit absolute output | Verify the resulting file and its contents; a full build replaces an entire database. |

### Export

Omit the deprecated `output_dir`: export only writes to the configured folder.
A different path is refused as `export_folder_mismatch`; unavailable folder
discovery is `export_folder_unavailable`. Use the reported `export_path`.
Category export is synchronous; whole-project export follows completion callbacks.
`full_export` controls changed-only versus all objects in scope, not destination.

Whole-project export has no `decision_policy`/`noninteractive` parameters in the
current schema. Do not assume it suppresses dialogs like single-object export.
Resolve conflicts according to the request; use recovery inspection if blocked.

### Single-object import

MCP/API import disables the index for the call and replaces the object from source.
`decision_policy="block"` does not add a conflict check to that path. If preserving
independent database edits matters or the winning version is unknown, use a
conflict-aware merge instead. A known conflict or `decision_required` must not be
sidestepped with single-object import. Back up as needed; there is no full-merge
backup here. Single-file component types can replace an entire category despite
the tool name, and the importer may update dependent objects. Check the live
description and the affected files before treating it as a one-object change.

### Category import and full merge

A category import deletes database objects in scope that have no corresponding
source. Ensure the category source is complete and the deletions are intended;
use a named import for a deliberate single replacement instead of a partial tree.
Own a backup before a destructive category merge.

`full_import=True` reloads all source in the named categories and skips conflict
detection: source wins regardless of `decision_policy`. Use it only when that
replacement is intended, for example after verifying an index missed changes.
It is ignored without `object_types`; selecting every category with it is not
a substitute for a backed-up full merge or a fresh build. Leave the binary index
to the add-in rather than resetting it to force a result.

`merge_not_available` is a refusal, not a transient error. Consider a full build
only if a fresh database fits the requested outcome and output is authorized;
do not escalate a failed merge into an unrequested rebuild.

### Full rebuild

Select a fresh output or an explicitly authorized replacement; verify the folder
exists and the source contains `vcs-options.json`. The tool creates an isolated
Access host and refuses an output already open without closing its holder.
Use the callback-reported `output_path` for subsequent checks and tests. On an
unconfirmed start it is `None`; `requested_output_path` is only intent.
Named-output `BuildAs` is part of the admitted release contract; capability
discovery is optional and no mandatory probe/refusal is promised.
The current implementation accepts but does not apply `template_path`; if a
template is required, report that limitation and arrange a supported workflow.
