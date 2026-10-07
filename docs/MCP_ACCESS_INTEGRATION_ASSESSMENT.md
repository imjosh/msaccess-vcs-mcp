# MCP-Access integration assessment

Research snapshot: **7 October 2026**. Recommendation: **keep the msaccess-vcs-mcp feature branch as the host; migrate selected MCP-Access capability implementations through the host's operation, connection, recovery, admission, and VCS boundaries.** Do not run both session managers inside one server or transplant the donor dispatcher wholesale.

This document separates **observed implementation**, **inference from source**, and **proposed integration behavior**. Classifications and acceptance criteria are proposals; the code observations beneath them are evidence. This is an integration assessment, not an implementation specification or a claim that the integration works.

## Snapshot and scope

| Repository | Inspected revision | Role |
|---|---|---|
| msaccess-vcs-mcp | `167843211e9c01001a5e60bb8866784ac5864a0a` | Host; new worktree on `codex/mcp-access-integration-assessment` |
| unmateria/MCP-Access | `9f219a4bccc961236070a107092796eac6b22c13` | Donor; package metadata reports 0.7.63 |
| msaccess-vcs-addin | `dab554ac4fdacc23edc6843a268684dc0b90c3ef` | Add-in source context; source inspection does not establish the version installed in Access |

The host worktree is `C:\Users\Josh\Documents\GitHub\work\msaccess-vcs\worktrees\mcp-access-integration-assessment`. The original MCP checkout has nine modified files outside this committed baseline: generated COM-wrapper repair, runtime identity metadata, version changes, and associated tests/documentation. Those changes were not copied into this worktree. Reconcile them before implementation; particularly preserve the repair's single-native-instance behavior rather than reintroducing a second Access process during Dispatch retry. The add-in has an uncommitted change in `modules/Tests/Components/modTestTableData.bas`; this assessment does not use that modified test as pinned evidence. Original repositories were read-only research context.

Sources are checked-out first-party code. Links below are GitHub permalinks generated from known remotes for locally inspected source, not a claim that these URLs were fetched successfully. They pin repository, commit, file, and line range; host/add-in links use their configured `imjosh` origin repositories rather than assuming upstream contains these feature commits. An unpublished revision can be inspected locally even if its GitHub permalink is unavailable. The SHAs, not current default branches or README feature counts, define this assessment's scope. Package metadata is cited for dependency/version claims only. [D-package], [H-package]

**Validation performed:** source inspection and static registration/dependency reconciliation. No tests, live Access operations, database opening, process interaction, registry writes, configuration changes, or installed-add-in verification were performed. Runtime claims quoted in donor comments remain claims of that implementation, not results reproduced here.

## Decisions and integration constraints

- **Use donor:** bring a donor implementation or pure algorithm into the host with its behavior intact. This can add a capability or replace an existing host implementation.
- **Adapt donor:** bring donor capability logic into the host, changing it to use host-owned lifecycle, effects, result, and synchronization boundaries.
- **Use host:** provide the capability through an existing host implementation; do not import the donor implementation.
- **Defer:** leave the donor capability out of the initial integration until a named contract or prerequisite is settled.

The gate serializes operations; it does not by itself select the same native Access instance. Host tools commonly construct a per-call `AccessConnection`, while donor tools use a process-global `_Session._app`. Host connection cleanup releases its COM references while, by default, preserving server-created Access processes for later reuse. Consequently, “one executor” or “one visible database” is insufficient evidence for “one Application instance.” [H-gate], [H-connection]

These labels describe the choice for each donor implementation. Host capabilities to retain are listed separately. A tool classified as **Adapt donor** can contain algorithms classified as **Use donor**; those algorithms can retain their behavior while the tool integration changes.

Proposed common boundary for adopted live capabilities:

1. Validate target and effect permissions before touching COM or files. If the capability will export through the add-in, complete add-in admission before mutation.
2. Acquire the existing gate and resolve one host-owned `AccessConnection`. Verify database identity from live state; capture PID/window/process creation identity when needed.
3. Pass `app`, `db`, and explicit operation state into capability helpers; helpers must not launch, switch, close, or independently connect Access.
4. Persist the intended change, invalidate derived data, and synchronize the affected VCS categories while holding the same operation ownership.
5. Normalize the result, including mutation and synchronization outcomes separately. Release references and ownership only when the worker and cleanup actually finish.

For VBA development, the durable mutation is the existing source import/build operation. This boundary does not justify exposing direct VBE writers. This is a proposed seam, not a demand for a new global singleton. Do not recursively call registered `vcs_*` wrappers from a gated capability: use internal integration methods under the already-held gate. The host wrapper gates tools and uses an explicit `ADDIN_DEPENDENT_TOOLS` membership list; registering a new decorator alone does not add compatibility admission. [H-wrapper], [H-admission]

## Complete donor tool inventory

Static audit found **69 unique `Tool(name=...)` declarations and 69 matching dispatcher branches**, with no missing or extra names. Three tools are filtered from discovery and refused at dispatch when `MCP_ACCESS_ALLOW_CODE_EXEC` is off, leaving **66 advertised by default**. This count is a source audit, not a live `tools/list` call. [D-tools], [D-dispatch], [D-server-list], [D-codegate]

Every registered donor tool appears exactly once below, grouped only where the decision and dependencies are comparable. Names are current donor names, not proposed public names. Adopted host tools should follow the host's `vcs_` naming and structured-success conventions. [H-wrapper]

Dependency notation:

- **S:** donor `_Session.connect` or direct singleton-state manipulation, including session-global launch/switch/recovery side effects.
- **P:** parsed controls cache, keyed `type:name`.
- **C:** live CodeModule COM-object cache, keyed `type:name`; **not** cached VBA text.
- **W:** donor global watchdog inherited by connected tools. **Ws/Wc/Wr/Wm/Wo:** additional save, compile, run/eval, maintenance, or open/cancel watchdog.
- **V:** VBA project object-model access; **E:** generated/existing code or macro execution.
- **X:** filesystem export/import/output path; **M:** database/object mutation; **U:** window or global user-input effects.
- **G:** donor operator code-execution switch. Only the three named tools have this dispatcher gate. All S tools inherit launch/open policy described later.

| Donor tools (count) | Implementation choice | Actual implementation and donor dependencies | Host equivalent / adaptation boundary |
|---|---|---|---|
| `access_list_objects` (1) | Use host | S,W; enumerates CurrentData/CurrentProject collections. [D-objects] | Retain `vcs_list_objects` and DAO discovery; adapt object-type aliases only if needed. |
| `access_table_info` (1) | Adapt donor | S,W; DAO TableDefs/Fields, falls back to SELECT COUNT for unknown counts (including linked/remote tables), returns linked Connect unmasked. [D-tableinfo] | Dedicated schema reader through admitted host DB; bound record-count work and default credential masking. Host DAO get_table_schema is an existing helper, not equivalent to this full public tool. |
| `access_get_code` (1) | Adapt donor | S,W,X; SaveAsText to temporary file, encoding detection, strips selected binary sections for forms/reports. [D-getcode] | Useful live inspection; explicitly distinguish donor display text from canonical add-in source. Use host connection and bounded temp I/O. |
| `access_set_code` (1) | Defer | S,C,P,W,Ws,V,M,X; form/report VBA-only route preserves layout via VBE; full-definition route restores binary blocks, backs up, LoadFromText, then injects VBA. [D-setcode] | Full-definition and VBA-only shortcuts duplicate canonical source edit/import. Do not expose another durable code-write route; internal encoding/binary/injection helpers may serve controlled import/build. Revisit only for a demonstrated need that source import/build cannot meet. |
| `access_export_structure` (1) | Use host | S,C,W,V,X; bespoke structure document plus VBE module text, optional output file. [D-structure] | Retain `vcs_export_database` as canonical source export. A human-oriented structure summary can later reuse inventory helpers; donor summary is not a VCS round trip. |
| `access_close` (1) | Use host | Direct `_Session.quit`, cache clearing, attached-vs-created teardown. [D-close] | Use host ownership-aware close lifecycle. `vcs_end_session` clears option overrides and **does not close Access**; do not alias the semantics. [H-end-session] |
| `access_vbe_module_info`, `access_vbe_get_proc`, `access_vbe_get_lines` (3) | Adapt donor | S,C,W,V; inspect actual imported procedures/lines. Donor preparation closes/saves forms/reports; resolver recovery can create/save code-behind via HasModule. [D-vbe-read], [D-vbe-module], [D-vbe-read-close], [D-vbe-read-init] | First VBE diagnostic slice after import/failed compile, once object/procedure/line context is known. Replace mutation-prone preparation: preserve or refuse dirty open state, report no-module without creating it. |
| `access_vbe_find`, `access_vbe_search_all` (2) | Defer | S,C,W,V; live text/regex search, partial errors; resolver shares initialization side effects. [D-vbe-search], [D-vbe-module] | Native repository search is the default; diagnostic triad handles targeted live verification. Reuse matching/context algorithms internally; reconsider public live search only for a demonstrated loaded-code-only diagnostic gap. |
| `access_search_queries` (1) | Adapt donor | S,W; searches QueryDefs.SQL. [D-query-search] | DAO-based read capability; no VBE prerequisite. Reuse host DAO identity and result limits. |
| `access_find_usages`, `access_find_definition` (2) | Adapt donor | S,C,P,W,V; Usages combines textual VBA search, QueryDefs.SQL and form/report control expressions; definitions scans current-DB declarations/continuations, excluding locals. No referenced-library definition lookup. [D-definition-scope], [D-usages], [D-definition] | Secondary source-backed structured Access navigation across code/query/control bindings and declarations; label provenance/partial results and heuristic scope. Normal repository search remains sufficient for simple lookups. No semantic call graph/safe rename; current donor does not resolve referenced libraries. |
| `access_vbe_replace_lines`, `access_vbe_replace_proc`, `access_vbe_patch_proc`, `access_vbe_append` (4) | Defer | S,C,P,W,Ws,V,M; DeleteLines/InsertLines, procedure/declaration handling, post-edit health checks, explicit save. Patch-proc prevalidates atomic matching but COM failure recovery remains best effort. [D-vbe-write], [D-patch] | Durable edits belong in canonical source followed by host import/merge or rebuild. Omit public live writers from routine integration; pure patch/parsing algorithms may support source edits/controlled imports. A future exception requires demonstrated necessity, isolation, preimage/drift and cleanup safeguards. |
| `access_vbe_check_syntax` (1) | Defer | S,C,W,V; reads modules and calls structural/block parsers from compile.py. [D-syntax] | Defer the standalone COM-backed tool. Adapt pure block/structure parsers as optional source-file preflight/post-import hints; no full grammar/type resolver or compiler replacement. Donor all-project scan filters component types 1/100, omitting class modules type 2. [D-syntax-scan] |
| `access_list_controls`, `access_get_control` (2) | Adapt donor | P; on cache miss: S,W,X via ac_get_code/SaveAsText. On hit: returns cached parse without connecting requested DB. [D-controls-read] | Reuse parser against fresh host export or explicitly identified source text. Redesign cache identity; this capability alone cannot prove shared live COM. |
| `access_search_controls` (1) | Adapt donor | S,P,W,X indirectly; scans object exports and selected properties, with decode/continuation handling. [D-controls-search] | Retain scanner and field selection; use one validated connection/export provider and partial-error results. |
| `access_create_control`, `access_delete_control`, `access_set_control_props`, `access_set_form_property`, `access_set_multiple_controls` (5) | Adapt donor | S,P,C,W,M; opens design, manipulates live controls/properties, saves/closes, invalidates cache. Creation/property writes can attach lint and tab-parent hints, adding V/E via auto measurements. [D-controls-create], [D-controls-write], [D-controls-batch] | Export whole form/report. Preserve property failures and created names; gate mutation before Design view. Replace hidden code execution in post-write lint with heuristic mode initially. |
| `access_get_form_property` (1) | Adapt donor | S,W,U; opens Design view and calls save-and-close in finally, despite being a getter. [D-form-get] | Preserve/read prior open/view/dirty state and use no-save cleanup where appropriate; do not market as effect-free. |
| `access_manage_tab_order` (1) | Adapt donor | S,P,C,W,U,M for set/auto; read/set/auto-renumber modes, section-aware exclusions, mutation modes invalidate caches. [D-taborder] | Classify by requested action; require writes only for set/auto, export whole object after changes, preserve view state. |
| `access_export_text`, `access_import_text` (2) | Use host | S,P,C,W,X; raw SaveAsText/LoadFromText with encoding and code-behind splitting; import mutates. [D-rawtext] | Retain `vcs_export_object`/`vcs_import_object` and add-in canonical representation. Raw text troubleshooting is separate optional scope, not a replacement for VCS artifacts. |
| `access_get_db_property`, `access_get_field_properties`, `access_list_startup_options` (3) | Adapt donor | S,W; reads DAO properties and selected startup options. [D-properties] | Add dedicated inspection through host DAO. Startup settings are metadata; opening the DB can still run startup logic. |
| `access_set_db_property`, `access_set_field_property` (2) | Adapt donor | S,W,M; DAO property write/CreateProperty fallback and coercion. [D-properties] | Export DbProperty or TableDef category as applicable; startup/security properties need effect policy and readback. |
| `access_list_linked_tables` (1) | Adapt donor | S,W; DAO Connect/SourceTableName; password masking is opt-in, default false. [D-linked-read] | Default credential masking; host logger must also redact `new_connect` and nested/returned connection strings. |
| `access_relink_table` (1) | Defer | S,W,M; ADODB preflight, LoginTimeout injection, DAO RefreshLink or delete+TransferDatabase with stored-login handling, best-effort restore. [D-relink] | Settle credential, remote-server, link deletion/recreation, schema refresh and TableDef synchronization policies first. |
| `access_list_relationships`, `access_list_indexes` (2) | Adapt donor | S,W; DAO Relations/Indexes inspection. [D-relations-read], [D-indexes] | Safe initial schema readers once common connection boundary exists. |
| `access_create_relationship`, `access_delete_relationship`, `access_manage_index` (3) | Adapt donor | S,W,M; DAO append/delete; relationship delete asks confirm; index supports create/delete. [D-relations-write], [D-indexes] | Export DbRelation collection or TableDef including indexes; explicit removal semantics for deleted relations. |
| `access_list_references` (1) | Adapt donor | S,W,V; current-project references, tolerates broken references. [D-references] | Use project matching by live filename; add-in is a loaded library, not the edit target. |
| `access_manage_reference` (1) | Adapt donor | S,C,W,V,M,X; add by file/GUID or remove, clears CodeModule cache. [D-references] | Validate library paths, export DbVbeReference metadata, invalidate all affected project handles; do not edit installed add-in references. |
| `access_execute_sql`, `access_execute_batch` (2) | Adapt donor | S,W,M depending SQL; DAO read/action branches, ODBC dbSeeChanges retry, lexical destructive-prefix confirmation. Batch is sequential and can commit earlier statements before failure. [D-sql], [D-batch] | Retain `vcs_execute_sql` for SELECT; add action/batch only behind explicit write/effect policy. SQL syntax classification and broad source synchronization required; no atomic-batch claim. |
| `access_manage_query` (1) | Adapt donor | S,W,M; create/update/delete/rename QueryDefs and SQL retrieval. [D-query-manage] | Export query definition; rename/delete must remove prior canonical source. Distinguish getter action from mutation. |
| `access_search_data` (1) | Adapt donor | S,W; bounded text-field search, skips linked tables, parameterized/fallback escaped patterns. [D-data-search] | Preserve limits/local-table scope and sensitive result policy; this is data access, not source export. |
| `access_create_database` (1) | Defer | Direct S,M,X; creates file, closes current singleton DB, reopens/exclusivity logic. [D-createdb] | Requires nonexistent-target validator, creation ownership and initial VCS source policy; do not disturb user database in reused app. |
| `access_create_table`, `access_alter_table` (2) | Adapt donor | S,W,M; DAO fields/defaults/validation/attributes; destructive field deletion needs confirm. [D-tables] | Export TableDef plus affected relations and configured TableData; admit exact supported types and rename behavior. |
| `access_create_form`, `access_build_form` (2) | Adapt donor | S,P,C,W,M; live create/design/save, layout planner/themes/defaults/control helpers and post-build lint. [D-createform], [D-buildform] | Retain planner algorithms; adapt execution, created-name persistence, partial builds, whole-form export and heuristic-only lint initially. |
| `access_clone_object` (1) | Adapt donor | S,P,C,W,M,X; raw export retains binary sections, changes identifier, routes import via ac_set_code. [D-clone] | New destination identity and overwrite guard; export cloned object; retain original unchanged and expose partial failure. |
| `access_delete_object` (1) | Defer | S,P,C,W,Ws,M; confirms, saves modules, deletes live object, clears caches. AC_TYPE excludes tables. [D-delete] | Deletion requires canonical source/index cleanup or tombstone contract; exporting a now-missing object cannot perform it. |
| `access_compile_vba` (1) | Use host | S,C,W,Wc,V,M,U; auto-decompile once, dirty standard module, focus current-project VBE pane, compile-menu trigger, parser fallbacks. [D-compile] | Retain `vcs_compile_vba`/`vcs_check_vba_compiled`. Adapt passive compile-selection/context helper into host failure diagnostics with project/correlation checks; retain user-guidance fallback. Reuse passive parsers separately; no automatic decompile import. [D-compile-location] |
| `access_run_vba` (1) | Use host | S,W,Wr,V/E,G; calls existing function via Application.Run; name misleading relative to host snippet tool. [D-run] | Existing host equivalent is `vcs_call_vba`, not `vcs_run_vba`; reconcile explicit execution authorization first. |
| `access_eval_vba`, `access_run_macro` (2) | Defer | S,W,Wr,E,G; Eval can fall back to temporary module execution (V,M); macro may execute RunCode. [D-eval], [D-run] | Settle uniform operator execution controls, temporary-module recovery and arbitrary-effect/source-sync policy; reuse hardened host execution infrastructure. |
| `access_output_report`, `access_transfer_data` (2) | Defer | S,W,X,E/M depending report/action; report OutputTo; data import/export Excel/CSV via DoCmd. [D-output] | Settle output path/overwrite policy, event execution, imported table schema/data effects, and configured TableData scope. |
| `access_compact_repair`, `access_decompile_compact` (2) | Defer | S,C,P,W,Wm,M,X,U; close/reopen, temporary replacement/backup, /decompile process, taskkill and PID-difference cleanup. [D-maintenance] | Separate ownership-aware maintenance contract, proven backup/replacement and no user-process closure. |
| `access_lint_form` (1) | Adapt donor | P plus export S,W,X; pure geometry rules; default auto mode can create temp module and run WizHook (V,E,M); optional screenshot adds U/X. [D-lint] | Retain pure rules; inject exported text, default heuristic; treat WizHook/screenshot branches as explicit separate effects. |
| `access_screenshot` (1) | Adapt donor | S,W,U,X; captures HWND via Win32/Pillow; optional object opening fires events except design mode; open watchdog sends ESC. Returns path/dimensions dict. [D-screenshot] | Start with current-window capture only, correct host PID. Add native MCP image result, path policy, bounded capture and view-state preservation. |
| `access_ui_click`, `access_ui_type` (2) | Defer | S,W,U,E/M; foreground switch, global mouse/key injection, unbounded SendMessage for text; no bounded semantic effect. [D-input] | Needs explicit user-input authorization, fresh PID/window/coordinate validation, bounded messages and arbitrary-effect synchronization contract. |
| `access_tips` (1) | Adapt donor | Static dictionary, no Access connection. [D-tips] | Retain useful guidance content only after rewriting session, security, canonical-source and host tool names. Donor workflow prompt is not host policy. |

All connected donor rows marked W inherit the global watchdog; this does not mean each capability creates its own thread. Parser cache-hit paths and static tips can avoid connecting entirely. The above dependencies identify what must be removed or injected, not just imports to copy.

## Source-first VBA workflow and VBE priorities

**Proposed default:** edit canonical exported files with native repository tools, review the source diff, import/merge affected objects or rebuild through existing VCS tools, then compile and run relevant tests. Use live diagnostics when imported state differs or compilation fails. Source remains the durable edit surface; a routine live patch → save → auto-export sequence adds a parallel development path and drift obligations without replacing a missing host workflow. [H-vcs-bulk], [H-rebuild-db], [H-vcs-object], [H-compile], [H-options-tests]

The revised choice defers **eight public entrypoints**: four VBE writers plus access_set_code, two standalone live code searches and the COM-backed syntax tool. It retains three adapted diagnostic primitives and secondary composite navigation. A read-like name alone is not enough reason to adopt a duplicate API.

| Priority / disposition | Tools or algorithms | Workflow value and limits |
|---|---|---|
| **Highest-priority internal compile-diagnostic adaptation** | `_get_vbe_error_location(app)` from compile.py; no standalone donor tool | Passively reads ActiveCodePane selection/module and nearby lines. Adapt into host `vcs_compile_vba` failure results on the same connection after failure, validating current-DB project and correlation to that compile. Current host/add-in compilation returns success/Boolean and lacks failing module/line; stale/unrelated selection is not authoritative. The helper needs no focus/selection/save mutation. [D-compile-location], [H-compile], [A-compile-result] |
| **First VBE diagnostic slice — Adapt donor** | `access_vbe_module_info`, `access_vbe_get_proc`, `access_vbe_get_lines` | Inspect actual imported/loaded code, declarations and a known procedure/line range. Get-proc/get-lines require caller-supplied identities; module-info derives names with regex and uses VBE bounds/parser fallbacks. Return DB/project/object provenance and live coordinates; do not promise automatic compile-error location. [D-vbe-read] |
| **Secondary structured navigation — Adapt donor** | `access_find_usages`, `access_find_definition` | Usages searches VBA text, query SQL and form/report control/event expressions; definitions scans current-DB declarations and continuation groups, explicitly excluding procedure locals. Adapt algorithms to source-backed structured Access navigation wherever possible. Results can hit strings/comments or be incomplete; no semantic call graph/reference completeness/safe rename. The donor does **not** search referenced libraries or type libraries. Simple lookups/refactors still use native repository search. [D-usages], [D-definition], [D-definition-scope] |
| **Defer standalone live search APIs** | `access_vbe_find`, `access_vbe_search_all` | Native repository search is normal and the diagnostic triad handles targeted live verification. Matching/context/continuation algorithms remain useful internally, including composite navigation; exposing separate whole-project live search requires a demonstrated loaded-code-only diagnostic need. [D-vbe-search] |
| **Adapt optional pure source preflight; defer standalone COM tool** | Parsers behind `access_vbe_check_syntax` | Extract block/structure routines to accept canonical text without opening VBE. These are heuristics, not full VBA grammar/type resolution or successful compilation; donor's all-project scan excludes class-module component type 2. Real `vcs_compile_vba` remains authority. [D-syntax], [D-syntax-scan], [H-compile] |
| **Defer public live VBA/object writers** | `access_vbe_replace_lines`, `access_vbe_replace_proc`, `access_vbe_patch_proc`, `access_vbe_append`, `access_set_code` | These duplicate source edit/import and create another synchronization surface. Pure text algorithms and internal injection/encoding helpers may support controlled import/build, without registering public writers. Revisit an exception only after demonstrating why source import/build cannot meet that task. [D-vbe-write], [D-patch], [D-setcode] |

**Observed read-side effects require adaptation before the diagnostic slice.** Get-lines/get-proc/module-info call `_close_form_design_view` before resolving the module. The helper attempts form/report DoCmd.Close with `AC_SAVE_YES` without a prior dirty/view check and suppresses failures. On resolution failure, `_force_vbe_init` opens Design view, can flip `HasModule=False` to True, then closes/saves; standard-module recovery toggles VBE visibility. A supposed read can save unrelated user edits or create/save code-behind. [D-vbe-read], [D-vbe-read-close], [D-vbe-read-init]

**Proposed adapted-reader contract:** identify the current host project, read with fresh operation-local handles, return no-module/unavailable without manufacturing code-behind, and preserve or explicitly refuse dirty/open objects. Do not replace every save-close with no-save-close: that can discard user edits. Required transitions need an explicit preservation/refusal contract. Do not run donor watchdog, trust-registry changes or automatic decompile to make a read succeed. Cache-hit proxies currently return unchecked; fresh text reads are not proof of cache-owner liveness. [D-vbe-module], [D-vbe-text], [D-watchdog], [D-registry], [D-compile]

**Debugger limitation:** the inventory exposes no standalone active-error/current-pane-selection, breakpoint, stepping, locals or watch tool. Internal `_get_vbe_error_location` does passively read selection/context; adapting it into host compile failures is useful, but is not an already exposed debugger API. Capture it on the same host connection immediately after failed compilation; verify the pane belongs to the current DB and correlate selection with that failure. Stale/unrelated selection must be omitted or labeled unverified, not presented as the error. Retain existing log/dialog context and user-guidance fallback when unavailable; known object/procedure/line context then makes triad readback useful. Live CodeModule line numbers may differ from serialized source headers/offsets; compare text rather than assuming identical line numbers. [D-tools], [D-vbe-read], [D-compile-location], [H-compile], [A-compile-result], [H-dialog-tools]

### Practical source edit → import → diagnostic example

1. Edit the canonical class-module or form code-behind source and review its repository diff. Import the corresponding object with existing `vcs_import_object`/`vcs_import_objects`, or rebuild when required. Compile, then run the relevant tests if compilation succeeds. [H-vcs-object], [H-vcs-bulk], [H-rebuild-db], [H-compile], [H-options-tests]
2. Compilation fails. Retain that failure. In the proposed adaptation, host compile captures best-effort passive selection/context immediately on its existing connection; only a current-project selection correlated to this failure supplies location. If absent/unverified, retain existing log/dialog inspection and user-highlighted-location guidance. Once context is known, adapted module-info/get-proc/get-lines shows imported code; no speculative live editing or runtime debugger controls are implied. [H-compile], [H-dialog-tools], [D-compile-location], [A-compile-result], [D-vbe-read]
3. Compare live code with the intended source, accounting for serialized headers/declarations/line offsets. If import omitted or transformed code, report that mismatch and investigate import. If imported code matches the erroneous source, fix source, review the diff, re-import, compile and rerun relevant tests. The repair stays in source.
4. If source has newer edits since import, or the user changed loaded code, report provenance and source drift; preserve both versions for deliberate reconciliation. Never auto-export over edited source as reader cleanup. Export is an explicit synchronization decision.

Scratch experiments or proven importer defects can motivate revisiting a live writer only after demonstrating a specific unmet need. Define isolated target, lifetime, preimage/drift protection, failure/cleanup and intentional source reconciliation. Defer does not authorize routine production work through scratch VBE. No exception, implementation or runtime behavior is validated by this research.

## Host capabilities to retain

The host has **25 registered `vcs_tool` handlers** at this snapshot. The grouping below covers them all. All capabilities in this section remain host implementations; identified guard and adapter work remains necessary.

| Host capabilities | Retained behavior and reason | Evidence |
|---|---|---|
| `vcs_export_database`, `vcs_import_objects`, `vcs_rebuild_database`, `vcs_export_object`, `vcs_import_object`, `vcs_diff_database` | Retain canonical source/build workflows, noninteractive policies, export-folder checks and operation results. Direct donor changes should feed these internal add-in boundaries, not create a parallel source format. | [H-vcs-bulk], [H-rebuild-db], [H-vcs-object] |
| `vcs_list_objects`, `vcs_execute_sql` | Retain existing discovery and SELECT pathway; richer navigation/schema readers and explicit action SQL are adaptations, not reasons to swap session stacks. | [H-vcs-bulk], [H-sql] |
| `vcs_get_version_info`, `vcs_get_recent_calls`, `vcs_get_log`, `vcs_cancel_operation` | Retain compatibility/runtime diagnostics and executed-call visibility; callers need these after timeout or uncertain completion. | [H-diagnostics], [H-options-tests] |
| `vcs_check_vba_compiled`, `vcs_compile_vba`, `vcs_call_vba`, `vcs_run_vba`, `vcs_run_tests`, `vcs_rebuild_addin` | Retain compile/execution/test/build lifecycle and existing timeout/recovery behavior. Donor compile and eval are not drop-in equivalents. | [H-compile], [H-call], [H-run], [H-options-tests], [H-rebuild] |
| `vcs_set_option`, `vcs_get_option`, `vcs_end_session` | Retain session override scoping and protected `McpAllowRunVBA`; “session” here is distinct from donor persistent COM session. | [H-options-tests], [H-end-session] |
| `vcs_list_dialogs`, `vcs_dismiss_dialog`, `vcs_recover_dialogs`, `vcs_automation_status` | Retain gate-exempt bounded inspector/recovery paths and structured interrupted/uncertain outcomes. | [H-dialog-tools], [H-recovery], [H-recovery-live] |

Retain host `AccessGate`, add-in admission, connection/instance registry, callback/operation management, config reload, logging and installed-add-in target refusal. Adapt new capability registration and effect metadata around these. Replace donor executor, singleton ownership, watchdog stack, atexit teardown and server/dispatcher transport with host infrastructure. [H-gate], [H-wrapper], [H-admission], [H-connection], [H-target-refusal]

Host's outer gate and worker timeout lifetimes require one additional distinction. Named-function calls use a separate COM-initialized daemon in `_run_application_call_with_timeout`; it can return a timeout while the daemon still lives, allowing the outer gate to finish. The VBA worker manager has its own active-worker/probe accounting but that is not a universal guard on all future direct COM calls. **Proposal:** direct mutation services must refuse or wait while a timed-out subworker may still be executing against the target. Do not assume the outer gate alone excludes every native Access operation. [H-call-timeout], [H-worker-manager]

## Session, cache and helper dependency map

### Singleton and second-instance hazards

**Observed:** donor `_Session` holds app, DB path, PID, attachment flag, CodeModule cache, decompile history and watchdog state globally. `connect` resolves the requested path, probes the current proxy/DB, launches if absent and switches databases if recorded paths differ. Launch first attempts one `GetActiveObject("Access.Application")` ROT entry, accepts idle or matching target, otherwise DispatchEx; exclusive mode avoids attaching. It does not enumerate every Access instance to find the target. [D-session], [D-launch]

**Inference:** loading donor helpers unchanged in host creates two independent connection owners and potentially two independent native instances. A gate shared only by host wrappers does not constrain the donor global thread, separate standalone server, decompile processes or imported atexit handlers. The donor singleton also switches its attached instance's current DB on subsequent requests, even though initial launch refuses to attach a different database. Host-created versus donor-created ownership flags cannot safely substitute for the host registry's PID plus creation-time evidence. [D-session], [D-launch], [D-switch], [D-close], [H-connection]

**Proposal:** make connection identity explicit per operation. Use a capability-scoped resolver that receives the same admitted host app/db used for subsequent export. Any longer-lived app registry must be an intentional host change with liveness, database identity, apartment, ownership and disconnect rules; do not achieve persistence accidentally through imported `_Session`.

### Actual caches and invalidation

| State | Observed behavior | Integration consequence |
|---|---|---|
| `_parsed_controls_cache` | Plain global map, key `form:name`/`report:name`; contains parsed SaveAsText controls and form text consumed by lint. Cache helper checks key **before connecting db_path**. | Cross-database incorrect reads are possible even in a single donor server: cache A/form/F, then request B/form/F; hit bypasses B's connection/switch. Treat this as a source-inferred regression path to test. [D-cache], [D-controls-read], [D-lint-model] |
| `_Session._cm_cache` | Caches live CodeModule COM proxies. _get_code_module cache hits return the proxy without a liveness test; CountOfLines/Lines are read later. Cache misses may invoke form/report initialization that creates/saves code-behind. | Do not retain proxies across host connection/app generations or thread changes. Start with operation-local handles; diagnostic readers must not create modules or save/close unrelated dirty objects. [D-vbe-module] |
| VBA text | `_cm_all_code` explicitly reads live CodeModule.Lines every call; an old text cache was removed. | Do not claim donor VBA text caching or copy stale-text invalidation rules from tips/comments. [D-vbe-text] |
| `_decompiled_dbs` | Tracks per-session paths already decompiled; compile consults it and cleanup clears it. | This is lifecycle/maintenance state, not a reusable source-analysis cache. Remove with donor auto-decompile. [D-session], [D-compile] |
| Object/all invalidators | Object invalidation removes both P and C keys; all invalidation clears both. DB switch, force cleanup and quit clear both; form/code/control mutators use targeted invalidation; create/clone/delete often clear all; reference changes clear C. | Host imports/rebuilds, user edits, direct VBA/SQL/schema mutations and recovery must participate; donor-only invalidators cannot see them. [D-cache], [D-switch], [D-close], [D-controls-write], [D-references] |

**Proposal:** disable cross-call caching in first slices. Later cache immutable parse results by normalized database identity, app generation or confirmed export hash, object type/name and content hash. Switch/reopen/recovery/build/import/reference changes invalidate the appropriate generation. Do not retain stale parses merely because the name still exists; external user edits have no donor mutation event.

### Helper modules worth separating

| Donor helper cluster | Implementation choice and scope |
|---|---|
| `constants.py`, `design_defaults.py`, parser/scanner portions of `controls.py`, geometry/rules in `lint.py` | **Use donor algorithms; adapt donor boundaries:** object/control constants, twip snapping, wrapped-property and octal-escape decoding, tab-parent extraction, section geometry, deterministic layout rules. Algorithms should accept text/models and avoid importing core for its session side effects. [D-controls-parser], [D-lint-model], [D-helpers], [D-buildform] |
| `helpers.py` | **Adapt donor:** encoding-aware read/write, strict encoding errors, binary extraction/restoration, code-behind splitting and property coercion. Binary restoration currently takes app and performs SaveAsText; isolate it from pure text routines. Canonical add-in serializers remain authoritative. [D-helpers] |
| `vbe.py` pure string/procedure routines and `compile.py` structural checks | **Use donor pure algorithms / adapt diagnostic integration:** passive compile-selection/context with project/failure validation [D-compile-location], matching, continuation context, declaration/procedure-kind handling and source syntax heuristics. Public live VBE writers/search/syntax remain **Defer** as classified above. Internal injection helpers may support controlled import/build without public writers. Replace mutation-prone reader preparation and save watchdog with host-owned policy; identify the current project and return no-module/unavailable without manufacturing code-behind. [D-vbe-module], [D-patch], [D-definition], [D-syntax], [D-project] |
| `build_form.py` layout planner/themes | **Use donor planner; adapt donor executor:** receives fields/actions, emits layout/control specifications, then live creation adds session/control/lint dependencies. Preserve the useful plan/execution separation. [D-buildform] |
| `ui.py` capture | **Adapt donor** window bitmap capture/Pillow cleanup and dimensions; **defer** foreground/global input helpers. Capture has a Pillow dependency absent from host metadata. [D-screenshot], [D-input], [D-package], [H-package] |
| `tips.py`, `tools.py` schemas/coercion, `server.py` prompt | **Adapt donor** useful documentation/schema constraints, but register with host decorator and host schemas/results. Retain argument sanitization intent. Do not import donor server/executor or reflect untrusted path text into instructions. [D-tips], [D-tools], [D-server-list] |

Donor imports `core.py` widely, including from apparently textual helpers. Core performs process DPI configuration, creates a ThreadPoolExecutor and registers global SHIFT release at import. This is another reason to extract algorithms instead of declaring the whole package a runtime dependency. [D-core-init]

## Watchdogs and timeout semantics

**Observed donor chain:** global session watchdog resolves PID, observes persistent dialogs, then calls `_dismiss_dialogs_by_pid`. Spawned instances use 3 seconds grace; attached instances use 5 seconds and require a tool to appear in flight. It tracks windows by HWND, then dismisses the PID's dialog set. The shared helper prefers Cancel, then other priorities such as End/OK, and falls back to WM_CLOSE. It records a dismissal timestamp/title and the server appends a textual diagnostic. That is not a structured interruption verdict or confirmation that the intended operation completed. [D-watchdog], [D-dismiss], [D-server-result]

| Operation-local donor path | Dependency to replace |
|---|---|
| DB switch/open | Core Wo thread polls after open grace and dismisses dialogs; synthetic SHIFT and AutomationSecurity changes accompany open. [D-switch] |
| VBE module save; save-all before delete | Ws wraps DoCmd.Save and calls shared dismissal helper; may fall back to saving individual modules. [D-vbe-save], [D-saveall] |
| Compile | Wc reads dialog text before dismissal; dirtying, decompile history, code-pane focus and menu trigger are intertwined. [D-compile] |
| Run/Eval | Wr polls dialogs; a parameter described as timeout enables watchdog behavior around a blocking COM invocation. Do not assume a hard deadline that stops VBA. [D-run], [D-eval] |
| Compact/repair/decompile | Wm wraps close/save/compact calls; subprocess paths dismiss by PID and terminate processes. [D-maintenance] |
| Screenshot with object open | Wo sends ESC after deadline; the blocked DoCmd call must still return before TimeoutError is raised. A current-window capture need not use this path. [D-screenshot] |

**Proposal:** replace every automatic broad dismissal with host policy-based inspector/recovery and operation interruption records. Keep dialog APIs gate-exempt so diagnosis remains possible while a live capability holds the gate. Preserve bounded worker capacity/deadlines and reinspection before action. Retain one interrupted or uncertain final result for client and logs; do not append “a dialog was dismissed” to an otherwise successful JSON string. [H-recovery], [H-recovery-actions], [H-recovery-outcomes], [H-recovery-live], [H-dialog-tools], [H-wrapper]

Host caller cancellation does not release the gate while its worker/Runner cleanup remains active. Preserve that property when introducing COM mutations, save/export cleanup and image generation. A timeout/cancel request is not evidence that Access stopped, and source synchronization must not race the still-running call. [H-gate]

## Security and configuration reconciliation

### Donor switches and automatic settings

| Setting/effect | Observed donor behavior | Proposed disposition |
|---|---|---|
| `MCP_ACCESS_ALLOW_CODE_EXEC` | Default false; gates and hides only run_vba, eval_vba and run_macro; read on each dispatch/discovery. | **Adapt donor** operator-controlled execution admission across actual effects, not just three tool names. [D-codegate], [D-server-list], [D-dispatch] |
| `MCP_ACCESS_SHIFT_BYPASS` | Default on, false-like values opt out; synthetic SHIFT is global input. | **Use host:** apply the host opening policy initially. Make any startup-bypass feature explicit and bounded, with documented limitations. [D-shift], [H-open] |
| `MCP_ACCESS_EXCLUSIVE` | Default off; skips attach when on, lock-file prechecks/open checks; shared mode adds advisory. | **Defer automatic adoption**; design-lock requirements are real but need host ownership/reuse mode contract, verified rather than inferred from a requested flag. [D-exclusive], [D-launch], [D-switch] |
| AccessVBOM | Separate `enable_vba_trust.ps1` sets HKCU Access Security/AccessVBOM=1 for found Office versions. VBE readers/writers depend on VBA project model access. | **Defer registry helper execution**; report prerequisites, never change trust automatically as part of normal tool use. [D-trust], [D-vbe-module] |
| Recovery registry | Launch/switch automatically delete Access Resiliency DisabledItems/StartupItems and set DisableAllCallersWarning/DoNotShowUI. | **Remove donor behavior from automatic path**; this alters user-wide Office settings, not capability-local state. [D-registry], [D-launch], [D-switch] |
| AutomationSecurity | Switch attempts force-disable value 3 before open, then sets 1 afterwards rather than restoring captured previous value. | **Use host:** apply host policy; any extension must capture/restore prior state and avoid claiming this blocks all startup macro execution. [D-switch] |
| Paths | Singleton resolves db_path; switch tests file existence. No donor universal root allowlist, installed-add-in protection or global write-disable check is present in inspected dispatch/session paths. | **Use host:** apply host-aware target and auxiliary-file validation, explicitly adding file creation/output/reference paths. [D-dispatch], [D-session], [D-switch], [H-security], [H-target-refusal] |

“Trusted path” has two meanings that must remain separate: permission to operate on a filesystem path, and Access trusting a location/code project. Neither donor execution gate nor host path validation establishes the other. No generic allowlist is inferred from a Trust Center setting.

### Existing and new guard gaps

**Observed host:** `ACCESS_VCS_DISABLE_WRITES` is enforced by `check_write_permission`, but the decorator does not call it universally. Current handlers call it for imports and rebuilds; call/run VBA, compile, set-option and other potential effects do not obtain a universal read-only guarantee from this switch. The protected `McpAllowRunVBA` option cannot be changed via `vcs_set_option`, but that is a narrower control. [H-security], [H-wrapper], [H-run], [H-options-tests]

**Observed donor:** broad mutation tools remain available with the execution switch off. They can change modules, event expressions, query SQL, startup properties and references. `access_lint_form(measure="auto")` can create and execute temporary measurement code despite not appearing in CODE_EXEC_TOOLS. Screenshot normal/preview/object-opening and report/UI paths can run existing events or macros. **Inference:** the donor switch closes three direct execution entrances; it is not a universal sandbox against execution or writes. A confirmation boolean supplied by an agent is not operator opt-in. [D-codegate], [D-lint], [D-screenshot], [D-properties], [D-input]

**Proposal:** declare effects centrally: read metadata/data, write database design/data, execute arbitrary/existing code, operate UI, modify user settings, and write external files. Determine dynamic effects by action/mode, then admit them before acquiring or opening a target. Start with fail-closed writes/code execution, heuristic lint and current-window screenshot. Reconcile existing `vcs_call_vba` flexibility with add-in `McpAllowRunVBA` rather than introducing a second overlapping switch with undocumented bypasses.

Host valid input extensions are .accdb/.accda/.mdb; .accde/.mde are excluded by current validator. Installed add-in targeting is refused separately, comparing paths while ignoring extension. **Proposal:** keep these restrictions; do not broaden them silently for donor support. Future compiled-file read-only tools require a separate compatibility policy and must never permit design/VBE mutation. New donor-like parameter aliases such as `db_path`/output/input/library paths must be covered explicitly; the host's current refusal binds known target parameter names. [H-security], [H-target-refusal]

Credential handling needs more than donor's opt-in masking or host's current named-key redaction. Donor linked-table readers default to clear Connect strings, relink returns old/new connection strings, and table_info can return Connect. Host redaction recognizes connection_string/connectionstring but donor uses new_connect and nested results. **Proposal:** mask by default in output and logs, restrict intentional secret readback, redact nested values and log effect metadata without executing payload bodies unless operator-enabled. [D-linked-read], [D-relink], [D-tables], [H-logging]

## Canonical source synchronization contract

**Normal VBA workflow:** source edit → host import/merge or rebuild → compile/tests → optional live read/compare. Public donor VBE writers and access_set_code are deferred; live-code export is not the default development path. Diagnostic source drift must not trigger automatic overwrite of newer source. [H-vcs-object], [H-compile], [D-vbe-read]

**Adopted live design/schema effects and any separately justified future code-write exception:** the following contract scopes their synchronization obligations; it does not imply adopting public code writers. **Proposal:** live write capability success must describe two distinct outcomes: (1) database mutation persisted, (2) canonical VCS source synchronized. Automatic export alone cannot make a DB/filesystem change atomic. If export fails after save, return failure with `database_changed: true`, `source_sync: "failed"`, affected identities and log/retry context; never retry the mutation blindly. Export cancellation or timeout yields an unknown/unconfirmed state when completion is not known.

Existing host object export/import uses add-in methods, policy confirmation and cleanup rather than raw SaveAsText as the VCS format. Keep that boundary. Source category selection and deletion cleanup need explicit mappings (verified details below); `vcs_export_object` is not a universal way to export every metadata collection or erase missing objects. [H-vcs-object]

| Direct change | Proposed source boundary | Why an object-only export is insufficient |
|---|---|---|
| Form/report controls, properties, tab order, create/build/clone | Entire containing Form/Report canonical component, including layout, module and add-in-managed sidecars/binary handling | Preserve existing code-behind; its durable edits use source import by default. Donor stripped display text/raw exports are not canonical source. [D-setcode], [D-helpers], [H-vcs-object] |
| Standard/class VBA and form/report code edits | Normal path: canonical source component imported/merged or rebuilt; no public live writer adopted | A separately justified future live-write exception needs save/readback and explicit whole-component reconciliation; never overwrite newer source on diagnostic cleanup. [D-vbe-write], [D-vbe-save], [H-vcs-object] |
| Query create/update | DbQuery definition | Rename/delete also need old source removal; action SQL can alter objects beyond the query. [D-query-manage] |
| Table fields/indexes/link definitions | DbTableDef and dependent categories; inspect configured TableData rather than assuming all data is versioned | Indexes live in TableDef; field rename/delete affects relation and query dependencies. [D-tables], [D-indexes] |
| Relations/references/database properties | DbRelation/DbVbeReference/DbProperty category or component mapping | These are metadata collections with add-in-defined persistence formats, not ordinary form/module names. [D-relations-write], [D-references], [D-properties] |
| Data mutation/import | TableData only where add-in export configuration includes it, plus new/changed TableDef when applicable | Row data can be outside source control; record that boundary explicitly. [D-batch], [D-output] |
| Arbitrary SQL/VBA/macro/UI | Full affected-category export or full database export after effects are known; otherwise defer | Cannot safely infer object scope from arbitrary code, events or mouse/key input. |
| Delete/rename | Add-in-authoritative source/index cleanup or deliberate tombstone/removal workflow | An absent object cannot be exported; stale source could rebuild a deleted object. Never infer artifact names and unlink them ad hoc. [D-delete] |

Hold the gate through mutation, save, invalidation and export completion. Admission and a known export folder must be established before mutation; otherwise the capability can knowingly create an unsynchronizable change. Reject or explicitly support databases without a VCS source tree. User edits performed outside this server remain outside its gate; source-hash/preimage and state readback checks reduce but cannot eliminate that race.

## MCP results and dependency packaging

Donor dispatch mixes JSON strings, raw code text, `OK:`/`ABORTED:`/`NOOP:` text and exceptions wrapped into text. Its server always returns `TextContent`; screenshot writes a PNG and returns a path/dimensions dictionary serialized as text. It does **not** return native MCP `ImageContent`. Appending watchdog warnings can also turn a JSON-looking response into non-JSON text. [D-dispatch], [D-server-result], [D-screenshot], [D-patch]

**Proposal:** retain host dict results with explicit `success`, `error_pattern`, mutation/source-sync status, partial errors, cleanup outcome and bounded results. Convert donor helper outcomes before transport rather than guessing success from nonempty strings. Return native image content plus structured metadata using the host's installed FastMCP/MCP facilities, extending logging/interruption handling deliberately if the result envelope is not a dict. Verify that actual MCP clients display the image; a valid disk path is not image delivery.

Host requires Python >=3.11 and mcp[cli]>=1.28,<2; donor Python >=3.10, mcp>=1,<2 and Pillow>=10. pywin32 overlaps. No demonstrated dependency-version conflict requires wholesale donor installation. **Proposal:** extract only needed code; add Pillow only with capture capability, retain host's MCP minimum and virtual-environment workflow. Donor declares MIT in metadata, but no tracked LICENSE/COPYING/NOTICE file was found in this checkout, while host tracks LICENSE. Resolve original attribution/license notice text before substantial copying; metadata alone does not provide that text. [D-package], [H-package]

## Incremental migration recommendation

1. **Baseline reconciliation and capability seam.** Review/retain original uncommitted host repairs on an intentional implementation baseline. Define effect metadata, admitted connection handoff, host project matching, structured results and immutable export-provider seam. Remove donor core imports from extracted pure algorithms. No new Access session owner.
2. **First milestone: control inspection plus same-instance current-window screenshot.** Adapt list/get controls using fresh host-managed object text. Capture the existing HWND without opening a form/report. Include an existing VCS operation in the proof. Parser output demonstrates text analysis; live identity and cross-operation observations demonstrate the shared native session.
3. **Source-first diagnostic slice, then secondary navigation.** Adapt passive compile-selection/context into existing host failure results with current-project/correlation validation and user fallback, then module-info/get-proc/get-lines to inspect actual imported state without hidden module creation/save/close; demonstrate source edit/import/compile/readback using existing host imports. Composite usages/definitions algorithms are secondary structured navigation, source-backed wherever possible. Reuse pure syntax parsers optionally. Standalone live find/search-all/syntax APIs remain deferred. Query/control/schema readers and heuristic lint remain separately assessed; operation-local handles precede cross-call caches.
4. **One synchronized write slice.** Start with a narrowly scoped control property update on a disposable fixture, save and whole-form add-in export under one gate. Return separate mutation/source results; test export failure before adding whole-form creation. This design/schema slice does not add a parallel live VBA edit path.
5. **Expand bounded design/schema capabilities.** Add form builder/clone, queries/tables/indexes/relationships/references only after their canonical category/export/deletion mapping is complete. Durable VBA edits continue through source import/build. Public VBE patch/append/replace and access_set_code stay deferred; reuse internals only when controlled import/build needs them.
6. **Separate deferred decisions.** Public live VBE/code writers and standalone live code-search/syntax APIs, maintenance/decompile, relink/credentials, file creation, data/report transfer, delete, Eval/macros and global UI remain out until their explicit prerequisites and permission/source-sync contracts are accepted.

### Concrete acceptance criteria

These are implementation acceptance requirements, not tests executed for this assessment.

- **Registry coverage:** enumerate adopted public tools and effects; disabled writes/exec are refused before COM/temp-module creation; direct tool-name calls cannot bypass discovery filtering. Gate-exempt diagnostics remain bounded during a blocked capability.
- **Identity proof:** in real Access, verify normalized live DB path, HWND PID and process creation identity for adapted live operation and VCS export/call; native instance reuse is observed, not inferred from executor. A timed-out named-function daemon or VBA worker cannot race a newly admitted direct mutation. An unrelated user-owned instance is neither switched nor closed.
- **Apartment/ownership proof:** all app/db/CodeModule proxies remain on the admitted apartment or an explicitly reacquired/marshalled worker boundary; timed-out subworkers remain accounted for before later mutations; cancellation retains ownership through active work and cleanup. No second executor/global watchdog is imported.
- **Cross-database correctness:** two fixture databases with identical object names but different controls/code never reuse each other's results; user edit, host import, rebuild, disconnect and reopened app generation invalidate relevant derived data.
- **Source-first proof:** source edit/import or build/compile/test/readback completes without deferred VBE/set-code writers. Readers identify DB/project/object and live coordinates; mismatches/newer source are compared and never automatically exported over.
- **Compile context:** failed host compilation returns validated best-effort current-project location/context where available without focus/selection/save changes; stale/wrong-project/unrelated selection is omitted or unverified, and absent location retains user-guidance fallback. This is compile diagnosis, not breakpoint/step/locals debugging.
- **Reader effects:** HasModule=False remains unchanged and returns no-module; dirty/open objects are preserved or refused, never silently saved/closed/discarded. A read does not manufacture code-behind or invoke donor init/decompile/registry recovery. Unknown error locations are not fabricated.
- **Navigation/parser limits:** composite results label source/live provenance, partial errors and heuristic scope; no semantic rename/reference guarantee or referenced-library lookup is credited. Optional pure source syntax accepts text without COM and does not count as successful compilation; omitted class-module coverage is explicit.
- **Parser fidelity:** wrapped properties, octal escapes, section/tab ancestry, binary blocks and class module headers round-trip or remain display-only with that limitation stated.
- **Screenshot transport:** actual MCP client receives a PNG image and metadata for the confirmed target window; no object opens, input injection or trust/registry edits occur in milestone one; output path validation/cleanup is explicit.
- **Write durability and sync:** real fixture saved, re-read/reopened, exported and diffed; unchanged objects remain unchanged. Export failure reports database_changed/source_sync failed. Importing exported source into a clean fixture reproduces the intended design/code.
- **Partial outcomes:** invalid control property, COM save failure, multi-property update failure, batch SQL failure and cleanup failure cannot become generic success. Helper text does not imply edit success; any later approved internal patch reuse preserves abort/no-op semantics.
- **Recovery:** report-only/destructive dialogs remain policy constrained; explicit runtime End records interruption, uncertain delivery remains failure, and timeout is not reported as rollback. No broad Cancel/WM_CLOSE fallback.
- **Security/compatibility:** installed add-in and compiled targets remain refused; new db/output/input/library aliases receive target/path protections; configuration flags without observed enforcement are not credited as guards; unsupported installed add-in refuses before mutation; heuristic lint works with code execution off; linked credentials are absent from default tool output/logs.

### Open decisions to settle before a spec

1. Public tool names/aliases: host-only `vcs_` names, or explicitly versioned donor compatibility aliases with equivalent admission?
2. Source-first diagnostics: define provenance, source/live comparison, header/line mapping and deliberate drift reconciliation; never overwrite newer source automatically. For adopted design/schema writes, decide synchronization versus explicit database-only mode and safe export retry.
3. Connection persistence: operation-local connection with validated process reuse, or a deliberate host-owned persistent app registry? Who owns app generations and cache invalidation?
4. Code execution: operator controls for snippet, named function, macro, Eval, measurement code, form/report events and query expressions. Which current escape hatches need compatible tightening?
5. Export mapping: exact add-in categories/flags for TableDef, TableData, relations, references and properties; deletion/rename cleanup; unsupported component cases.
6. Diagnostic-reader state: preserve or refuse a user's open/dirty object; report absent code-behind without setting HasModule. What context can logs/dialogs provide, and when must the user supply a highlighted error location?
7. Exclusive design operations: lock acquisition/proof, reuse behavior and release; no import of exclusive mode as a global startup default.
8. Referenced libraries: current donor definitions search only the current DB; library/type-library navigation would be a new optional capability. Keep existing reference introspection useful for compiler prerequisites; require explicit scope/path policy if library navigation is proposed.
9. Native image results: concrete FastMCP result type, size limits, log envelope and client rendering compatibility.
10. Ownership-safe maintenance and auxiliary I/O: creation, backups, replacement, credential storage, report/data output and global input authorization.
11. Deferred live-code/search/syntax APIs: require a demonstrated source-import/build or targeted-readback limitation before reconsidering them. Any writer exception needs isolated experimental/import-repair scope, preimage, cleanup and explicit source reconciliation; available parsers/injection helpers alone are not justification.

### Verified add-in export boundaries and side effects

The add-in's implementation lives in `modules/API/clsVersionControl.cls`; `modAPI.bas` is the dispatch layer. Contrary to a narrow core-only assumption, host `vcs_export_object` already accepts noncore types including `relation`, `vbe_reference`, `db_property` and `table`. Add-in aliases include `relation(s)`, `vbe_reference(s)`, `db_property(s)` and table aliases; generic `relationships`/`references` aliases are not established by that resolver. [H-vcs-object], [A-aliases]

| Existing add-in persistence | Observed implementation | Adaptation consequence |
|---|---|---|
| Core Form/Report/Module/Query export | ExportObject lookup requires the object to exist; direct API/MCP export disables VCSIndex persistence. [A-object-export] | It works for existing object readback, but cannot export a deleted object or be assumed to reconcile tracking/index metadata. |
| DbVbeReference | `vbe-references.json`; `SingleFile=True`. [A-reference-file] | Export is collection-oriented. Last-reference removal/empty collection may produce no iteration to rewrite file; specifically test it. |
| DbProperty | `dbs-properties.json`; `SingleFile=True`. [A-property-file] | Same collection/empty-result concerns; a property change is not an independent file per property. |
| DbRelation | `relations/<safe filename>.json`; noncore lookup compares requested name with FSO.GetBaseName of exported item keys. [A-relation-file], [A-object-export] | Logical relation name and safe filename can differ; map identity through add-in helpers. |
| DbTableDef | `tbldefs` folder, indexes serialized in table definition, linked Connect sanitized/environment-aware; modification detection uses dates and hashes. [A-tabledef], [A-tabledef-state] | Preserve add-in serializer and configured data behavior. Force scoped export after direct DAO changes when timestamp detection is uncertain. |
| Scoped category reconciliation | ExportByType routes ExportScoped. It scans category state, preserves conflict policy, calls ClearOrphanedSourceFiles, then exports; supports force/full flags and can refuse option-hash drift. [A-by-type], [A-scoped] | Candidate mechanism for delete/rename reconciliation: forced category export via existing API with conflict semantics. Do not invent direct source-file deletion or unconditional overwrite. |

**Material side effect:** scoped API/MCP export closes open objects of the category with `acSaveYes`, which can save unrelated user edits. Category export for delete/index reconciliation is therefore broader than “only synchronize my object.” Define prior open/dirty-state behavior and report the scope. Single-object export's disabled index persistence versus category export's broader effects is an actual design tradeoff requiring a spec decision. [A-scoped]

Host SQL delegates to add-in ExecuteSQL, whose current implementation checks `McpAllowExecuteSQL`, accepts a SELECT prefix and opens a read-only snapshot. Keep this existing read pathway; action/batch SQL requires a different explicit contract. `McpAllowImport` and `McpAllowCallVBA` appear in clsOptions definitions/defaults/config serialization but no enforcement use was found by searching the VBA modules. Treat that as a static inspection finding, not a promised permission boundary. [H-sql], [A-sql], [A-options], [A-options-default], [A-options-config]

Microsoft's first-party STA documentation requires per-thread COM initialization, message handling and marshaling across apartments; that supports explicit proxy-lifetime boundaries, not a blanket claim that a shared executor makes all subworkers safe. The Access OpenCurrentDatabase documentation specifies default shared opening and distinguishes opening an Access current database from DAO OpenDatabase. These API contracts reinforce the identity/exclusivity decisions; they do not replace real-instance validation. [Microsoft STA documentation](https://learn.microsoft.com/windows/win32/com/single-threaded-apartments), [Microsoft OpenCurrentDatabase documentation](https://learn.microsoft.com/en-us/office/vba/api/access.application.opencurrentdatabase).

The assessment supports incremental migration from the stabilized host feature branch. For VBA, prioritize live diagnostic readback of imported code while source edit/import/compile/testing stays the durable workflow. Composite navigation and pure parsers complement it; public VBE/code writers and duplicate live search/syntax APIs remain deferred. Form/control/schema capabilities retain their separate assessment and host-owned effects/synchronization boundaries.

## Pinned source references

[D-package]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/pyproject.toml#L5-L33
[H-package]: https://github.com/imjosh/msaccess-vcs-mcp/blob/167843211e9c01001a5e60bb8866784ac5864a0a/pyproject.toml#L5-L46
[H-gate]: https://github.com/imjosh/msaccess-vcs-mcp/blob/167843211e9c01001a5e60bb8866784ac5864a0a/src/msaccess_vcs_mcp/access_gate.py#L128-L270
[H-connection]: https://github.com/imjosh/msaccess-vcs-mcp/blob/167843211e9c01001a5e60bb8866784ac5864a0a/src/msaccess_vcs_mcp/access_com/connection.py#L412-L848
[H-wrapper]: https://github.com/imjosh/msaccess-vcs-mcp/blob/167843211e9c01001a5e60bb8866784ac5864a0a/src/msaccess_vcs_mcp/tools.py#L1070-L1218
[H-admission]: https://github.com/imjosh/msaccess-vcs-mcp/blob/167843211e9c01001a5e60bb8866784ac5864a0a/src/msaccess_vcs_mcp/compatibility.py#L28-L34
[D-tools]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/tools.py#L1-L1645
[D-dispatch]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/dispatcher.py#L84-L783
[D-server-list]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/server.py#L23-L108
[D-codegate]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/security.py#L19-L57
[D-objects]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/code.py#L119-L151
[D-tableinfo]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/database.py#L284-L335
[D-getcode]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/code.py#L277-L302
[D-setcode]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/code.py#L405-L598
[D-structure]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/code.py#L681-L788
[D-close]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/core.py#L1153-L1194
[H-end-session]: https://github.com/imjosh/msaccess-vcs-mcp/blob/167843211e9c01001a5e60bb8866784ac5864a0a/src/msaccess_vcs_mcp/tools.py#L3822-L3874
[D-vbe-read]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/vbe.py#L722-L884
[D-vbe-module]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/vbe.py#L158-L281
[D-query-search]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/vbe.py#L1188-L1217
[D-usages]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/vbe.py#L1220-L1333
[D-definition]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/vbe.py#L1830-L2252
[D-vbe-write]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/vbe.py#L924-L1038
[D-patch]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/vbe.py#L1336-L1652
[D-syntax]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/vbe.py#L1655-L1777
[D-controls-read]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/controls.py#L238-L321
[D-controls-search]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/controls.py#L349-L546
[D-controls-create]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/controls.py#L618-L808
[D-controls-write]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/controls.py#L847-L1117
[D-controls-batch]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/controls.py#L1385-L1428
[D-form-get]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/controls.py#L1119-L1161
[D-taborder]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/controls.py#L1200-L1383
[D-rawtext]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/controls.py#L873-L1033
[D-properties]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/properties.py#L42-L186
[D-linked-read]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/relations.py#L72-L117
[D-relink]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/relations.py#L15-L259
[D-relations-read]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/relations.py#L262-L286
[D-indexes]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/relations.py#L481-L551
[D-relations-write]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/relations.py#L289-L367
[D-references]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/relations.py#L370-L478
[D-sql]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/sql.py#L13-L126
[D-batch]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/sql.py#L129-L230
[D-query-manage]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/sql.py#L234-L298
[D-data-search]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/sql.py#L322-L474
[D-createdb]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/database.py#L36-L104
[D-tables]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/database.py#L108-L335
[D-createform]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/code.py#L603-L678
[D-buildform]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/build_form.py#L163-L558
[D-clone]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/code.py#L813-L901
[D-delete]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/code.py#L244-L270
[D-compile]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/compile.py#L631-L841
[D-run]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/vba_exec.py#L220-L331
[D-eval]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/vba_exec.py#L335-L473
[D-output]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/export.py#L24-L113
[D-maintenance]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/maintenance.py#L22-L348
[D-lint]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/lint.py#L1184-L1389
[D-screenshot]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/ui.py#L20-L231
[D-input]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/ui.py#L238-L400
[D-tips]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/tips.py#L358-L375
[H-vcs-bulk]: https://github.com/imjosh/msaccess-vcs-mcp/blob/167843211e9c01001a5e60bb8866784ac5864a0a/src/msaccess_vcs_mcp/tools.py#L1220-L1853
[H-rebuild-db]: https://github.com/imjosh/msaccess-vcs-mcp/blob/167843211e9c01001a5e60bb8866784ac5864a0a/src/msaccess_vcs_mcp/tools.py#L1854-L2067
[H-vcs-object]: https://github.com/imjosh/msaccess-vcs-mcp/blob/167843211e9c01001a5e60bb8866784ac5864a0a/src/msaccess_vcs_mcp/tools.py#L2680-L2844
[H-sql]: https://github.com/imjosh/msaccess-vcs-mcp/blob/167843211e9c01001a5e60bb8866784ac5864a0a/src/msaccess_vcs_mcp/tools.py#L2846-L2891
[H-diagnostics]: https://github.com/imjosh/msaccess-vcs-mcp/blob/167843211e9c01001a5e60bb8866784ac5864a0a/src/msaccess_vcs_mcp/tools.py#L2376-L2536
[H-options-tests]: https://github.com/imjosh/msaccess-vcs-mcp/blob/167843211e9c01001a5e60bb8866784ac5864a0a/src/msaccess_vcs_mcp/tools.py#L3472-L3820
[H-compile]: https://github.com/imjosh/msaccess-vcs-mcp/blob/167843211e9c01001a5e60bb8866784ac5864a0a/src/msaccess_vcs_mcp/tools.py#L2538-L2677
[H-call]: https://github.com/imjosh/msaccess-vcs-mcp/blob/167843211e9c01001a5e60bb8866784ac5864a0a/src/msaccess_vcs_mcp/tools.py#L2894-L3090
[H-run]: https://github.com/imjosh/msaccess-vcs-mcp/blob/167843211e9c01001a5e60bb8866784ac5864a0a/src/msaccess_vcs_mcp/tools.py#L3354-L3469
[H-rebuild]: https://github.com/imjosh/msaccess-vcs-mcp/blob/167843211e9c01001a5e60bb8866784ac5864a0a/src/msaccess_vcs_mcp/tools.py#L2069-L2318
[H-dialog-tools]: https://github.com/imjosh/msaccess-vcs-mcp/blob/167843211e9c01001a5e60bb8866784ac5864a0a/src/msaccess_vcs_mcp/tools.py#L3877-L4046
[H-recovery]: https://github.com/imjosh/msaccess-vcs-mcp/blob/167843211e9c01001a5e60bb8866784ac5864a0a/src/msaccess_vcs_mcp/dialog_recovery.py#L499-L528
[H-recovery-live]: https://github.com/imjosh/msaccess-vcs-mcp/blob/167843211e9c01001a5e60bb8866784ac5864a0a/src/msaccess_vcs_mcp/dialog_recovery.py#L1621-L1717
[H-target-refusal]: https://github.com/imjosh/msaccess-vcs-mcp/blob/167843211e9c01001a5e60bb8866784ac5864a0a/src/msaccess_vcs_mcp/tools.py#L3140-L3208
[H-call-timeout]: https://github.com/imjosh/msaccess-vcs-mcp/blob/167843211e9c01001a5e60bb8866784ac5864a0a/src/msaccess_vcs_mcp/tools.py#L3288-L3349
[H-worker-manager]: https://github.com/imjosh/msaccess-vcs-mcp/blob/167843211e9c01001a5e60bb8866784ac5864a0a/src/msaccess_vcs_mcp/vba_worker_manager.py#L552-L590
[D-session]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/core.py#L339-L423
[D-launch]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/core.py#L515-L622
[D-switch]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/core.py#L926-L1150
[D-cache]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/core.py#L1257-L1267
[D-lint-model]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/lint.py#L227-L407
[D-vbe-text]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/vbe.py#L283-L301
[D-controls-parser]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/controls.py#L30-L232
[D-helpers]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/helpers.py#L19-L381
[D-project]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/core.py#L1206-L1254
[D-core-init]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/core.py#L20-L84
[D-watchdog]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/core.py#L429-L512
[D-dismiss]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/vba_exec.py#L79-L228
[D-server-result]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/server.py#L115-L157
[D-vbe-save]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/vbe.py#L26-L101
[D-saveall]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/code.py#L154-L239
[H-recovery-actions]: https://github.com/imjosh/msaccess-vcs-mcp/blob/167843211e9c01001a5e60bb8866784ac5864a0a/src/msaccess_vcs_mcp/dialog_recovery.py#L1176-L1287
[H-recovery-outcomes]: https://github.com/imjosh/msaccess-vcs-mcp/blob/167843211e9c01001a5e60bb8866784ac5864a0a/src/msaccess_vcs_mcp/dialog_recovery.py#L249-L314
[D-shift]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/security.py#L60-L105
[H-open]: https://github.com/imjosh/msaccess-vcs-mcp/blob/167843211e9c01001a5e60bb8866784ac5864a0a/src/msaccess_vcs_mcp/access_com/connection.py#L151-L210
[D-exclusive]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/security.py#L108-L143
[D-trust]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/enable_vba_trust.ps1#L7-L31
[D-registry]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/core.py#L880-L908
[H-security]: https://github.com/imjosh/msaccess-vcs-mcp/blob/167843211e9c01001a5e60bb8866784ac5864a0a/src/msaccess_vcs_mcp/security.py#L7-L126
[H-logging]: https://github.com/imjosh/msaccess-vcs-mcp/blob/167843211e9c01001a5e60bb8866784ac5864a0a/src/msaccess_vcs_mcp/usage_logging.py#L521-L565
[A-aliases]: https://github.com/imjosh/msaccess-vcs-addin/blob/dab554ac4fdacc23edc6843a268684dc0b90c3ef/Version%20Control.accda.src/modules/Core/modContainers.bas#L187-L218
[A-object-export]: https://github.com/imjosh/msaccess-vcs-addin/blob/dab554ac4fdacc23edc6843a268684dc0b90c3ef/Version%20Control.accda.src/modules/API/clsVersionControl.cls#L1205-L1363
[A-reference-file]: https://github.com/imjosh/msaccess-vcs-addin/blob/dab554ac4fdacc23edc6843a268684dc0b90c3ef/Version%20Control.accda.src/modules/Components/clsDbVbeReference.cls#L488-L554
[A-property-file]: https://github.com/imjosh/msaccess-vcs-addin/blob/dab554ac4fdacc23edc6843a268684dc0b90c3ef/Version%20Control.accda.src/modules/Components/clsDbProperty.cls#L600-L666
[A-relation-file]: https://github.com/imjosh/msaccess-vcs-addin/blob/dab554ac4fdacc23edc6843a268684dc0b90c3ef/Version%20Control.accda.src/modules/Components/clsDbRelation.cls#L357-L398
[A-tabledef]: https://github.com/imjosh/msaccess-vcs-addin/blob/dab554ac4fdacc23edc6843a268684dc0b90c3ef/Version%20Control.accda.src/modules/Components/clsDbTableDef.cls#L135-L198
[A-tabledef-state]: https://github.com/imjosh/msaccess-vcs-addin/blob/dab554ac4fdacc23edc6843a268684dc0b90c3ef/Version%20Control.accda.src/modules/Components/clsDbTableDef.cls#L1527-L1587
[A-by-type]: https://github.com/imjosh/msaccess-vcs-addin/blob/dab554ac4fdacc23edc6843a268684dc0b90c3ef/Version%20Control.accda.src/modules/API/clsVersionControl.cls#L1605-L1659
[A-scoped]: https://github.com/imjosh/msaccess-vcs-addin/blob/dab554ac4fdacc23edc6843a268684dc0b90c3ef/Version%20Control.accda.src/modules/Core/modExport.bas#L470-L633
[A-sql]: https://github.com/imjosh/msaccess-vcs-addin/blob/dab554ac4fdacc23edc6843a268684dc0b90c3ef/Version%20Control.accda.src/modules/API/clsVersionControl.cls#L2461-L2480
[A-options]: https://github.com/imjosh/msaccess-vcs-addin/blob/dab554ac4fdacc23edc6843a268684dc0b90c3ef/Version%20Control.accda.src/modules/Infrastructure/clsOptions.cls#L72-L75
[A-options-default]: https://github.com/imjosh/msaccess-vcs-addin/blob/dab554ac4fdacc23edc6843a268684dc0b90c3ef/Version%20Control.accda.src/modules/Infrastructure/clsOptions.cls#L185-L188
[A-options-config]: https://github.com/imjosh/msaccess-vcs-addin/blob/dab554ac4fdacc23edc6843a268684dc0b90c3ef/Version%20Control.accda.src/modules/Infrastructure/clsOptions.cls#L1018-L1027
[D-vbe-read-close]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/vbe.py#L266-L280
[D-vbe-read-init]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/vbe.py#L215-L263
[D-vbe-search]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/vbe.py#L1040-L1185
[D-definition-scope]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/vbe.py#L1910-L2015
[D-syntax-scan]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/vbe.py#L1678-L1750
[D-compile-location]: https://github.com/unmateria/MCP-Access/blob/9f219a4bccc961236070a107092796eac6b22c13/mcp_access/compile.py#L19-L45
[A-compile-result]: https://github.com/imjosh/msaccess-vcs-addin/blob/dab554ac4fdacc23edc6843a268684dc0b90c3ef/Version%20Control.accda.src/modules/API/clsVersionControl.cls#L2908-L2952
