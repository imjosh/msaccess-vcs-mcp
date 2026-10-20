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
| Replace one named object from edited source | `vcs_import_object` | Explicit request policy protects independent destination edits; inspect skipped or blocked results. |
| Apply changed source in complete categories | `vcs_import_objects` with `object_types`, normally `full_import=False` | Category scope reconciles orphans and takes no database backup. |
| Integrate project-wide source changes into the existing database | `vcs_import_objects` without `object_types` | Full merge checks conflicts and takes a database backup. |
| Create a fresh database from source | `vcs_rebuild_database` with source and an explicit absolute output | Verify the resulting file and its contents; a full build replaces an entire database. |

### Export

Omit the deprecated `output_dir`: export only writes to the configured folder.
A different path is refused as `export_folder_mismatch`; unavailable folder
discovery is `export_folder_unavailable`. Use the reported `export_path`.
Category export is synchronous; whole-project export follows completion callbacks.
`full_export` controls changed-only versus all objects in scope, not destination.

Whole-project and category exports accept explicit request policies and default
to `block`, as single-object export does. Preserve source edits before exporting.

### Single-object import

Single-object import enforces the explicit request policy. Default `block`
protects independent database edits. Naming an object identifies scope rather than
overwrite authorization. Read transferred, skipped, and blocked outcomes. Own a
backup as needed because there is no full-merge backup on this path. Some component
types can affect dependent objects or an entire category, so inspect their scope.

### Category import and full merge

A category import deletes database objects in scope that have no corresponding
source. Ensure the category source is complete and the deletions are intended;
use a named import for a deliberate single replacement instead of a partial tree.
Own a backup before a destructive category merge.

`full_import=True` reloads all source in the named categories while retaining
explicit policy protection. Use it when the full source scope is intended.
Conflicting edited deletions require interactive review or preservation.
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
