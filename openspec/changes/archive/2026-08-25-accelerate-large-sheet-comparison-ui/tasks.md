## 1. Baseline and Exact Oracle

- [x] 1.1 Record the existing dirty worktree and implementation hashes, define read-only real fixtures for Skill, WorldMonster, Dungeon, Language, and the composite-key IdleBuilding Sheet, and guarantee all generated mutations stay in disposable directories.
- [x] 1.2 Implement a fresh-process legacy exact-result manifest that serializes logical row/column mappings, typed cell/formula/cache differences, structure, conflicts, and only-difference membership for 2-way and 3-way fixtures.
- [x] 1.3 Implement cold and warm performance and memory capture for startup, selected-Sheet readiness, exact comparison, cached revisit, 1,000-row actions, undo/redo, scrolling, and save as separate phases.
- [x] 1.4 Add adversarial Oracle fixtures for duplicate/missing keys, composite keys, blank continuation groups, equal-count insert/delete, reorder, formula/cache changes, column structure, stale generation, and cancellation.

## 2. Immutable Per-Sheet Snapshot

- [x] 2.1 Add immutable typed cell, field, record, row, render-span, and Sheet snapshot data structures with explicit parser/generation/mutation version keys.
- [x] 2.2 Build paired sequential read-only value/formula ingestion for one requested Sheet, preserving current literal, date, boolean, error, blank, normal/array/data-table formula, cached-value, and external-link semantics.
- [x] 2.3 Add a generation-safe selected-Sheet snapshot cache and lifecycle that leaves unopened worksheet XML unscanned and drops stale results without touching Tk.
- [x] 2.4 Remove comparison-path dependence on normal-mode value/formula workbook objects for large selected Sheets while retaining a guarded legacy feature-flag fallback.

## 3. Schema-Aware Exact Comparison

- [x] 3.1 Parse and normalize field declarations/type rows, align logical columns by validated schema identity, and emit explicit insert/delete/reorder/unresolved structural results.
- [x] 3.2 Build validated composite record identities from applicable `@id`/`@const` fields, preserve record order, and group blank-key continuation rows with their owning record.
- [x] 3.3 Implement unique row-hash anchors plus bounded deterministic fallback for invalid/no-key regions, with explicit ambiguity and no unbounded quadratic matching.
- [x] 3.4 Produce exact 2-way and Base-anchored 3-way cell/structure/conflict results entirely from immutable snapshots with no worksheet access.
- [x] 3.5 Compare new manifests against the frozen legacy Oracle across all real and adversarial fixtures and resolve every non-conservative mismatch before enabling the new path by default.

## 4. Virtualized Result Presentation

- [x] 4.1 Separate the complete logical display result from a fixed virtual window with logical-to-visible row/column mappings, selection anchors, overscan, and a hard 320-row Tk pane limit.
- [x] 4.2 Replace prefix append/load-more scrolling with a logical scrollbar supporting wheel, page, thumb drag, minimap, first/middle/last jumps, and direct navigation to unrendered difference blocks.
- [x] 4.3 Precompute pane strings, row/header state, diff spans, padding, structural markers, and block membership off the Tk thread, then publish each viewport with bounded bulk text/tag operations.
- [x] 4.4 Coalesce high-rate scroll events to the newest logical position and prove scroll callbacks perform zero worksheet reads, formula normalization, alignment, or diff computation.
- [x] 4.5 Preserve synchronized horizontal geometry, C-area behavior, selection, copy, hover, row headers, touched/resolved state, only-difference counters, minimap scale, and next/previous block navigation across recycled windows.

## 5. Accelerated Operations and Undo

- [x] 5.1 Add a typed Sheet-local operation overlay keyed by stable logical record/field identity while retaining physical save coordinates and existing manual operation records.
- [x] 5.2 Route single-cell and row adoption through bounded overlay recomputation and visible-window publication without full alignment or full Text rebuild.
- [x] 5.3 Batch region operations, including 1,000-row and offscreen targets, into one overlay transaction, one affected-result update, and one bounded UI publication.
- [x] 5.4 Route content-only undo/redo through overlay deltas and local result updates; advance topology generations only for row/column/Sheet structure changes.
- [x] 5.5 Rebuild only the affected Sheet after structural operations, reject actions against stale logical mappings, and keep unrelated Sheet snapshots reusable.

## 6. Save and Lifecycle Integration

- [x] 6.1 Feed overlay and structural operations into existing ZIP/native Excel save staging without changing immutable sources or bypassing formula-cache decisions.
- [x] 6.2 Preserve atomic replacement, package validation, Excel reopen, failure recovery, XLSM/VBA handling, comments, links, row/column metadata, and formula/reference behavior.
- [x] 6.3 Update Sheet readiness, progress, cancellation, retry, interaction gates, and feature-flag rollback so view-only exact comparison does not require eager editable-workbook materialization.

## 7. Verification and Acceptance

- [x] 7.1 Add GUI tests proving a 20,000-difference result keeps every Tk pane at or below 320 rows and supports rapid wheel/thumb/minimap/block navigation with correct logical targets.
- [x] 7.2 Add instrumentation tests proving scroll and cached revisit perform zero worksheet reads and measuring P95 viewport render, heartbeat, selected-Sheet comparison, peak RSS, action, undo, and redo thresholds from the specs.
- [x] 7.3 Run the existing smoke and GUI suites covering row/column alignment, only-diff, formula/cache behavior, structural actions, undo/redo, SVN merge semantics, XLSM, save fidelity, and Excel reopen; fix every regression in scope.
- [x] 7.4 Run real-workbook end-to-end acceptance on disposable Mine/Base/Theirs copies, publish normalized Oracle/performance evidence, and enable the new default only when every correctness and safety gate passes.
- [x] 7.5 Validate the OpenSpec change, document any calibrated threshold or fallback decision in the design/validation evidence, and leave the change apply-complete without archiving it.

## 8. Exact Whole-Workbook Sheet Status and Interaction Gates

- [x] 8.1 Add a generation-safe per-Sheet state registry with pending, calculating, exact-same, exact-changed, unresolved, and failed states; never infer unchanged from missing or incomplete data.
- [x] 8.2 Prioritize the selected Sheet, then background-compare every remaining supported Sheet so workbook navigation eventually exposes a complete exact changed/unchanged map and progress.
- [x] 8.3 Add prominent Sheet/workbook calculation UI, exact-state navigation badges, whole-workbook progress, and a calculating surface that replaces provisional/stale comparison rows.
- [x] 8.4 Centralize readiness guards for copy, overwrite, accept, region, structural, undo/redo, and save; reject non-ready operations without queuing them and show a modal with Sheet, state, stage, progress, reason, and retry condition.
- [x] 8.5 Add 2-way and 3-way GUI tests proving no provisional result is shown, no rejected action mutates state, stale generations cannot unlock controls, and exact completion atomically publishes the correct Sheet status/result.

## 9. Complete Real-Workbook Corpus Benchmark and 15-Second Gate

- [x] 9.1 Build a recursive read-only inventory of every file under `C:\GM15\design\sheets\develop`, explicitly report temporary/unsupported files, and create all test sides and mutations only in disposable roots.
- [x] 9.2 Add fresh-process 2-way and 3-way corpus workers that record per-file startup/catalog/whole-summary/RSS and every Sheet's dimensions, formulas, engine/fallback, request-to-final-exact time, final state, revisit time, and error.
- [x] 9.3 Run identical-copy 2-way and Mine/Base/Theirs self-comparisons for every supported file/Sheet, require a zero-difference/conflict Oracle, and publish machine-readable results plus separate slowest-Sheet rankings.
- [x] 9.4 Optimize measured ingestion, scheduling, comparison, adapter, and publication bottlenecks iteratively until every supported Sheet reaches final exact readiness within 15 seconds in both modes; preserve the 1.5/2.5-second figures as reported stretch targets.
- [x] 9.5 Re-run the complete corpus after each optimization class and retain before/after evidence, timeouts, failures, fallback reasons, and unsupported cases without excluding slow outliers.
- [x] 9.6 Keep direct parser/comparator Oracle timing separate from a fresh-child `SowMergeApp` runtime corpus; enforce the 15-second gate on constructor-to-current-generation terminal full comparison detail and physical targets without proactively loading editable workbooks, and report first mutation-backend load/action separately.

## 10. Accuracy, Operations, and Excel Compatibility Release Gates

- [x] 10.1 Create deterministic disposable value, formula/cache, row, and column mutations for the slowest real Sheets and require exact legacy/new Oracle parity for 2-way/3-way results, conflicts, only-diff membership, and physical operation targets.
- [x] 10.2 Exercise cell, row, region, structural, undo, redo, and save workflows on slow/representative real workbooks; prove rejected non-ready actions make no changes and accepted exact-ready actions produce the expected overlay/manual/native records.
- [x] 10.3 Validate every supported source workbook through disposable no-op package/reopen checks and validate mutated representative XLSX/XLSM outputs for atomic recovery, formulas/references/caches, VBA, comments, links, relationships, and row/column metadata using openpyxl and real Excel where available.
- [x] 10.4 Run the complete existing smoke/GUI/save/SVN/XLSM matrix, fix every in-scope regression, verify no source workbook changed, validate the OpenSpec strictly, and leave the change unarchived until all 15-second, Oracle, operation, and Excel compatibility gates pass.

## 11. Changed-Revision Interaction Regression Gate

- [x] 11.1 Preserve the update75 `Dungeon` revision-39265/39264 diagnostic evidence and add bounded telemetry for UI activity, heartbeat, CPU/RSS, viewport P95, edit-loader caller/reason, and C-area render work.
- [x] 11.2 Remove every view-only path from editable-workbook materialization; use immutable snapshot/render/cache data or a clear placeholder for missing values.
- [x] 11.3 Restrict editable-workbook materialization to explicit mutation/save demand, start it asynchronously under one owner, reject the first operation with the existing explanatory modal, and require retry after readiness.
- [x] 11.4 Make remaining-Sheet computation cooperative and idle scheduled: require a 1–2 second quiet window, yield/reschedule for UI activity, and preempt hidden work for a newly selected Sheet.
- [x] 11.5 Deduplicate hover payload calculation and suppress C-area full-row rebuild/tagging for same-logical-row hover.
- [x] 11.6 Add deterministic 2-way/3-way GUI regressions for view-only interactions, loader prohibition, quiet-window/preemption, C-area deduplication, two-dimensional wide-column virtualization (first/middle/last and combined row/column windows), heartbeat, and viewport metrics.
- [x] 11.7 Add a read-only/disposable `Dungeon` 39265/39264 2-way/3-way release gate covering `Dungeon` and `MonsterGroup` exact readiness, hover/wheel/thumb/tab interaction, wide-column/off-screen physical and Base operation targets, no editable load, 15-second final results, heartbeat/viewport limits, and exact operation targets.
  - [x] 11.7a Add pure immutable-snapshot duplicate-field proof tests before enabling real inputs: accepted two-stage two-way and Mine/Base/Theirs same-ordinal duplicate runs create non-actionable blank proof anchors first, then either verify an existing exact top cache or rebuild a pending complete bounded interval. Require complete row keys/pairs, unique-or-`START`/`END` anchors, all-blank proof value/formula tokens, equal full digests, exactly one same-logical proof slot, a non-ambiguous/non-structural all-side final bijection, and—on the rebuild path—same unique ordered/count non-proof interval fields with preserved formula cell diffs/physical Base targets. Cover an exact three-way top cache with asymmetric Mine/Base versus Theirs/Base child gaps. Reject nonblank or formula-bearing proof columns, digest-unequal runs, width/missing members, run-count or ordinal mismatch, missing/reordered anchors or non-proof interval sequence, row-key/pair ambiguity, 3-way Base mismatch, inherited formula mismatch with any non-proof unresolved/ambiguous/structural/non-bijective slot, unrelated formula or cache causes, crosswired or non-retained proof slots, structural/unresolved columns, incomplete side bijection, and candidate-builder/rebuild exceptions. Every negative case must remain `UNRESOLVED` with no action target or legacy fallback.
- [x] 11.8 Run the complete Oracle, operation, save, XLSX/XLSM, real-Excel, smoke, and GUI matrix; strictly validate OpenSpec and package only a new EXE for user testing without commit, push, or archive.
