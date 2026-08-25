## Context

The application currently has two competing data paths. Startup and mutation keep normal-mode openpyxl value/formula workbooks, while background exact comparison opens another 4-6 read-only workbooks and builds per-Sheet caches. Large-Sheet rendering limits the initial prefix, but the Tk `Text` documents and their per-row/per-cell tags grow whenever more results are materialized. A full refresh deletes and reinserts the rendered documents and reapplies tags, so work scales with the number of displayed differences rather than the visible viewport.

Read-only profiling of the current project data shows that a normal Skill workbook load costs about 182-185 MB per value/formula copy, whereas opening read-only and streaming one large Sheet costs roughly 4 MB. The sampled Sheets above 3,000 rows also expose stable schema markers (`@id`, `@const`, or composite `@id` fields), making deterministic record alignment preferable to workbook-wide materialization and content-fuzzy alignment.

The repository already has strict guarantees for formula/cache identity, row and column structure, generation-safe background work, atomic save, native Excel replay, package validation, and Excel reopen. The new engine must improve performance without weakening those guarantees or overwriting the existing dirty worktree.

## Goals / Non-Goals

**Goals:**

- Make first useful comparison depend on the selected Sheet, not the total workbook cell count.
- Produce an immutable, formula-aware semantic snapshot that is the only comparison/render input.
- Prove new results against the current exact implementation before changing the default path.
- Keep Tk document size and tag count bounded by the viewport for full and only-difference modes.
- Make scrolling, selection, navigation, and hover independent of workbook I/O and total difference count.
- Make content-only operations update an overlay and a bounded set of result rows; reserve full invalidation for topology changes.
- Preserve every existing save/reopen and merge semantic guarantee.

**Non-Goals:**

- Per-cell author attribution, SVN history replay, or provenance storage.
- Physically splitting Sheets or workbooks, changing source workbook layout, or changing data-export contracts.
- Replacing Excel-native structural replay where it is required for formula/reference fidelity.
- Treating style-only or recalculation-only package noise as a user-authored data edit.

## Decisions

### 1. Establish a frozen legacy Oracle before switching engines

Add a fresh-process harness that serializes the current exact comparison into a normalized manifest containing Sheet identity, logical column slots, aligned record/row identities, change kinds, physical coordinates, typed values, raw formulas, cached values, conflicts, and only-difference membership. Real source workbooks are read-only; all mutations are made to disposable copies outside the repository.

The new engine must match the manifest exactly for supported unambiguous inputs and must be at least as conservative for ambiguous inputs. This is preferred over using screenshots or row counts because those can miss wrong-cell alignment and formula/cache regressions.

### 2. Introduce immutable per-Sheet semantic snapshots

Use paired sequential read-only streams initially: one cached-value workbook and one formula workbook per comparison side. They preserve current openpyxl type/date/formula semantics while avoiding normal-mode cell object materialization. Each selected Sheet is consumed once into compact immutable rows; all later comparison, rendering, scrolling, and hover operations use the snapshot only. A later direct-OOXML decoder is allowed only if it proves byte-for-semantic parity against this reader.

Snapshots contain typed value/formula tokens, compact row tuples, logical field identities, physical coordinates, record keys, row hashes, rendered fragments, and diff spans. Cache keys include file signature, side, Sheet, parser version, generation, and mutation version. Stale generations cannot publish.

This is preferred over pandas/Polars because those tools discard Excel formula/cache distinctions and workbook structure, and preferred over physical Sheet splitting because normal workbook loads still materialize the same cells.

### 3. Align declared schema identities before fuzzy content

Columns use the normalized first-row field declaration plus second-row type declaration as their primary identity; order changes remain explicit structural changes. Rows use every applicable `@id`/`@const` component and validate uniqueness independently on each side. Blank-key continuation rows attach to the preceding keyed record and align inside that bounded group. Record order is compared separately so a reorder is not hidden by a dictionary join.

When declared identities are absent or invalid, use unique row-hash anchors followed by bounded deterministic alignment inside unmatched windows. Any window that exceeds the configured product/side limit or remains duplicate-ambiguous becomes an explicit unresolved block; it is never silently auto-merged. This avoids worst-case quadratic behavior and wrong-row actions.

### 4. Separate the complete logical result from the rendered viewport

`_full_display_rows` (or its replacement) remains the complete ordered logical result. A new virtual window stores `window_start`, `window_length`, logical-to-visible mappings, selection anchors, and bounded overscan. Tk receives at most 320 logical result rows per pane regardless of the complete result size.

The vertical scrollbar represents the complete logical result rather than the physical Tk document. Wheel, page, thumb, minimap, and next/previous-difference commands update a desired logical position. Events are coalesced with `after_idle`/a short frame budget, then all pane text is replaced in one batch and tags are applied in bulk. Intermediate rows are never materialized merely to jump to a distant result.

A fixed virtual Text window is preferred over a new Canvas grid because it preserves existing text selection, copying, horizontal alignment, fonts, and tag semantics with less UI risk.

### 5. Publish render-ready rows from background computation

Background snapshot comparison prepares each row's pane strings, diff-cell spans, row/header state, block membership, and structural markers. The Tk thread only selects a viewport slice, installs bounded strings/tags, and restores logical selection/scroll state. No `Worksheet.cell`, worksheet iteration, formula normalization, alignment, or diff computation is permitted in scroll callbacks.

### 6. Use an operation overlay and topology-aware invalidation

Content operations record typed cell/formula/cache replacements in a Sheet-local overlay keyed by logical record and field identities, with physical coordinates retained for save. A cell or row action recomputes only the affected record rows and neighboring block boundaries. Region operations batch overlay writes and publish one comparison/UI update. Undo and redo swap overlay deltas and reuse the same bounded path.

Row/column/Sheet insertions, deletions, or reorders advance the topology generation and rebuild the affected Sheet snapshot off the Tk thread. They do not invalidate unrelated Sheets. The existing editable/native backend may be loaded lazily only when a structural/save operation requires it; it is not a prerequisite for view-only exact comparison.

### 7. Preserve the existing save safety boundary

The operation overlay feeds the existing manual operation records, ZIP cell patch path, Excel-native structural replay, formula-cache handling, immutable source staging, package validation, atomic replacement, and Excel reopen gate. Successful in-memory comparison is never treated as proof that a saved workbook is valid. Save failures leave the source and overlay recoverable.

### 8. Measure cold, warm, interaction, and mutation phases separately

Benchmarks run in fresh subprocesses and report wall time, P50/P95, peak RSS delta, parsed rows/cells, worksheet access counts, Tk document line/tag counts, heartbeat gaps, and operation/save phases. Required real fixtures include Skill, WorldMonster, Dungeon, Language, and at least one composite-key Sheet. Synthetic fixtures cover 20,000 changed rows, duplicates, blank continuation groups, formulas, column structure, cancel/stale generations, and ambiguous fallback.

### 9. Publish only generation-matched final Sheet states

Each workbook owns a generation-safe Sheet-state registry with explicit states: `PENDING`, `CALCULATING`, `EXACT_SAME`, `EXACT_CHANGED`, `UNRESOLVED`, and `FAILED`. A Sheet may contribute to the workbook changed/unchanged summary only after its current generation reaches one of the two exact terminal states. `PENDING`, `CALCULATING`, stale, cancelled, unresolved, and failed results are never displayed or counted as unchanged.

The selected Sheet has first scheduling priority. After it reaches a terminal state, the worker continues through every remaining supported Sheet so the navigation can eventually answer which Sheets changed across the whole workbook. The navigation shows per-Sheet state badges plus workbook progress (`exact/total`, changed count, calculating Sheet). Selecting a non-ready Sheet shows a calculation surface instead of provisional cell differences.

Whole-workbook status and selected-Sheet operation readiness are deliberately separate deliverables. A hidden large Sheet may publish an exact generation-matched `EXACT_SAME`/`EXACT_CHANGED` summary from a bounded cache once all semantic and structural differences have been decided, without retaining every display string. If the user later selects that Sheet and complete per-cell operation data is not materialized, the selected Sheet returns to `CALCULATING`, shows only the calculation surface, and asynchronously upgrades to the full immutable logical result. The earlier exact navigation badge remains a trustworthy workbook-summary fact; it is not exposed as actionable cell data until the full upgrade finishes.

Comparison-detail readiness is also distinct from editable-workbook backend readiness. The selected Sheet's 15-second user-visible gate ends only after its current-generation terminal state, full prepared logical rows, typed formulas/caches, column/Base mappings, and physical operation targets are atomically installed and the calculation surface is gone. It does not proactively load normal-mode editable workbooks. If an exact result is visible but the mutation/save backend is still deferred, operation controls remain explainable and handler-guarded: the first attempted mutation/save starts the single-owner backend load, shows the modal reason, performs no mutation, and requires a retry after backend readiness. First-operation load and action latency are measured separately.

Every mutation entry point uses one shared readiness guard. If the current Sheet/generation is not exact-ready, no action is queued or partially executed. A prominent persistent banner remains visible, and an attempted copy/overwrite/accept/region/structural/undo/redo/save action opens a modal containing the Sheet name, current stage/progress, why exact data is required, and the instruction to retry after completion. This deliberately favors trustworthy final data over early previews.

Guarded operation widgets remain clickable while a Sheet is non-ready so a real user click reaches the shared explanatory modal; visual disabled styling alone is insufficient because Tk would swallow the event. Handler-side readiness is authoritative, including structural and undo/redo paths, and tests assert that the click creates no overlay, manual/native operation, topology change, stack replay, or delayed action.

### 10. Benchmark the complete real workbook/Sheet corpus

Inventory recursively enumerates all files under `C:\GM15\design\sheets\develop`, ignores lock/temp artifacts such as `~$*`, and records unsupported extensions explicitly. Every supported `.xlsx`/`.xlsm` file is copied to a disposable root and measured in fresh processes in both modes:

- 2-way: two independent disposable copies of the same source workbook;
- 3-way: independent Mine/Base/Theirs disposable copies of the same source workbook.

Self-comparison is intentional: it provides a deterministic zero-difference Oracle while exercising actual workbook parsing, formulas, links, styles metadata, Sheet topology, and 3-way reconciliation. For every file, capture application startup/catalog time, selected-first-Sheet time, whole-workbook exact-summary time, RSS, failures, and fallbacks. For every Sheet, capture rows/columns/formulas, request time, exact-ready time, cache source, final exact state, revisit time, and Oracle equality. Results are written as machine-readable JSON plus a slowest-Sheet report.

Corpus evidence has two non-interchangeable tiers. The direct snapshot/legacy Oracle tier isolates parser, alignment, formula-token, fallback, and exactness costs and is the deterministic correctness baseline. The user-visible runtime tier creates a fresh real `SowMergeApp` child for every file/mode/Sheet, selects that Sheet through the production entry point, and measures from constructor request through current-generation terminal comparison detail, including workbook catalog, scheduling, fallback, adapter, complete physical targets, and bounded Tk publication. It SHALL NOT request editable-workbook preload inside that timing or memory window. Only the user-visible runtime tier can satisfy the 15-second opening/readiness gate; direct-comparator timing is diagnostic and cannot be presented as application opening speed. Mutation-backend load and first accepted operation are separate reported phases.

The release acceptance maximum is 15 seconds from Sheet request/selection to final exact result for every supported Sheet in both 2-way and 3-way modes on the reference workstation. The earlier 1.5/2.5-second values remain optimization/stretch targets, not release blockers, because the user explicitly accepts 15 seconds. Any Sheet above 15 seconds, timeout, crash, unresolved self-comparison, or false changed state remains an optimization blocker.

### 11. Optimize from measured slowest Sheets without weakening fidelity

Optimization proceeds from the corpus ranking rather than representative guesses. First remove accidental broad-workbook materialization and scheduling contention. Then optimize the dominant measured phase using compact snapshots, shared tokens, stable-row hash skips, direct selected-Sheet OOXML pull parsing, or a local rebuildable cache only when that phase is proven dominant. A direct XML decoder cannot become authoritative until its typed value/formula/cache/structure manifest matches the paired-openpyxl Oracle.

The benchmark additionally seeds deterministic value, formula, row, and column mutations into disposable copies of the slowest Sheets and representative XLSM/link/comment/metadata workbooks. These variants must match the legacy Oracle and their accepted operations must survive existing ZIP/native save staging, package validation, atomic replacement/failure recovery, openpyxl reopen, and real Excel reopen where available. Performance work stops or rolls back on any comparison, targeting, undo/redo, or saved-package mismatch.

Acceptance targets on the current reference workstation are:

- every supported real Sheet in the complete corpus reaches a final exact result within 15 seconds in both 2-way and 3-way modes; 1.5 seconds for Skill and 2.5 seconds for 18k-21k-row narrow fixtures remain stretch targets recorded in evidence;
- three-way Skill comparison peak RSS at least 50% below the measured legacy full-workbook path and no more than 400 MB;
- virtual scroll render P95 no more than 33 ms, no heartbeat gap above 200 ms, and no Tk pane above 320 logical rows;
- cached Sheet revisit P95 no more than 100 ms;
- single content action P95 no more than 250 ms, 1,000-row region apply no more than 2.0 s before save, and undo/redo P95 no more than 500 ms;
- exact Oracle parity and all existing fidelity suites remain mandatory even if a timing target is met.

### 12. Real changed-revision interaction is a release gate

The update75 live trace for `Dungeon.xlsx` revision 39265 versus 39264 is authoritative regression evidence. `Dungeon` reached its terminal result in 10.038 seconds, but a hover callback observed a cached-value blank and synchronously loaded four editable workbooks for 32.495 seconds on the Tk thread, producing a 33.128-second heartbeat gap. The process exceeded 700 MB RSS. At the same time the scheduled `MonsterGroup` background comparison remained CPU-bound, so even otherwise small interactions showed persistent latency. Identical-copy corpus runs used an identity fast path and therefore did not exercise this changed-revision fallback.

View-only callbacks — hover, click inspection, wheel/page/thumb/minimap, tab/viewport activation, and C-area refresh — may consume only immutable snapshots, prepared render fragments, and small UI caches. They MUST NOT call `ws_*_edit`, `_ensure_edit_loaded`, or `_request_edit_preload`; a missing cached value renders an explicit non-actionable placeholder instead. Editable workbooks are single-owner asynchronous resources for an explicit mutation or save request only. The first request remains rejected with the existing explanatory modal while that backend loads, and requires retry after readiness.

Remaining-Sheet work is deliberately opportunistic: it starts only after a 1–2 second quiet window following selected-Sheet readiness, checks for recent UI activity at bounded checkpoints, yields/reschedules while interaction is active, and gives a newly selected Sheet priority over hidden work. Hover payload generation is computed once per target; C-area same-logical-row hover does not rebuild/re-tag the complete row. Diagnostics record heartbeat, CPU, RSS, edit-loader reason/caller, viewport P95, and C-area render timing.

The release gate uses read-only revision inputs or disposable copies of `Dungeon.xlsx` revisions 39265/39264. In both 2-way and 3-way modes, `Dungeon` and `MonsterGroup` must reach exact final results within 15 seconds; continuous hover/wheel/thumb/tab interaction must show no editable loader, heartbeat gap above 200 ms, or viewport render P95 above 33 ms. Exact Oracle parity, operation targets, package validation, XLSX/XLSM fidelity, and real Excel reopen remain hard constraints.

### 13. Wide sheets use a two-dimensional logical viewport

The vertical viewport alone is insufficient for a dense wide sheet: a 20-row
window with 69 long columns still creates a large Tk `Text` mutation and native
paint burst.  Wide sheets therefore retain the complete immutable comparison
model but render only a two-dimensional `(row window, column window)` surface.
This is rendering virtualization, not a data truncation or a new comparison
mode.

The complete raw fragments, typed values/formulas, logical field identities,
physical Mine/Base/Theirs coordinates, difference maps, Base mappings, and
operation projections remain authoritative and complete.  A wide virtual
surface materializes only the current visible logical columns plus a bounded
two-to-four-column formatting overscan (normally about 8--12 columns on the
69-column Dungeon sheet).  Its column spans map the short rendered line back to
the original logical columns.  Copy, hover, cell focus, C-area inspection,
column actions, and save/undo targets resolve through that mapping and then use
the complete raw/projection model; none may depend on recycled Text positions.

The horizontal scrollbar and horizontal minimap represent the complete logical
column sequence.  Their first/middle/last requests are coalesced to the newest
logical column window, just like vertical requests.  A combined vertical and
horizontal request is published from the two bounded logical windows without
reading a worksheet, re-aligning records, or recomputing differences.  An
off-screen logical cell is first brought into its row and column window before
Text selection is applied.  Narrow sheets retain the existing full-column path.

Publication remains opaque and atomic.  While a current generation is
calculating, the calculation overlay hides every pane.  The parent first
installs the complete immutable model and mapping, then builds the complete
initial bounded `(row, column)` surface, including headers, tags, C-area
geometry, scrolling state, and operation identities.  Only then may it publish
the exact terminal state and uncover the pane.  It must never progressively
reveal partial rows or columns.  A stale generation or tab change invalidates a
queued window before it mutates Tk.

Validation covers first/middle/last horizontal thumb positions, combined
vertical/horizontal navigation, off-screen cell focus/copy/operation targets,
three-way Base mappings, same-row hover de-duplication, and narrow-sheet
regression.  Every view-only route remains zero-worksheet-read.  The 20,000
changed-row and changed-revision gates require viewport P95 <= 33 ms and no UI
heartbeat interval above 200 ms without reducing default column widths or
discarding displayed data.

### 14. Duplicate declared-field identities require a bounded all-side proof

A repeated normalized `(declaration, type)` field identity is normally an
alignment ambiguity and remains fail-closed.  The engine may clear only that
specific duplicate-field reason, and only after producing an immutable,
snapshot-only candidate mapping.  It must not enable legacy fallback, clear a
row-alignment failure, or downgrade any unrelated column ambiguity.

`START` and `END` are virtual, unique schema anchors for this proof.  They are
valid only as the outer boundaries of a side's complete declared field list;
they never stand in for a missing interior field.  For every duplicate member
on Mine, Theirs, and (when present) Base, the candidate is valid only when all
of the following hold:

1. Declared record keys are complete and unique on every side, and their
   record grouping is validated without a leading blank continuation.
2. The resulting row-pair relation covers every physical row on every side,
   has no one-sided pair, and contains no ambiguous/unmatched block.
3. The proof runs in two stages.  Stage one produces only immutable blank
   duplicate anchor pairs; it publishes no exact result and no operation
   target.  Stage two first verifies an already exact top-level logical-column
   cache: every proof coordinate has exactly one same-logical Mine/Theirs
   (and, when present, Base) slot; the model and every slot are non-ambiguous
   and retained; `unresolved_cols` and `structural_diff_cols` are empty; and
   every side is a complete physical bijection.  That verification may retain
   an exact three-way top cache even when its Mine-to-Base and Theirs-to-Base
   child anchor gaps are asymmetric; it never infers a new mapping from those
   child gaps.  Only when the top cache is pending does stage two rebuild the
   complete bounded column-alignment interval using the proof pairs as fixed
   anchors.  Both paths require the same final resolved/bijective cache.  This
   final validation is deliberately separate from candidate construction so a
   two-way duplicate does not bypass column-cache evidence.
4. For each corresponding occurrence of the duplicated identity, that
   occurrence lies in the same interval bounded by the nearest unique identity
   or virtual `START`/`END` on every side.  The interval width, duplicate-run
   size, and zero-based ordinal in that run are identical.  This comparison is
   per occurrence across sides; it does not require different occurrences of
   the same identity to share one interval.  An interior occurrence and a
   separate tail occurrence bounded by `END` can therefore both be proven.
5. Every cell's cached value and formula token in the corresponding physical
   column is blank on every compared side, and its complete
   typed/formula-aware column digest is byte-identical on every compared side.
   Digest equality is an additional proof, not permission for nonblank
   duplicate content.  Equality of a header, a blank sample, or a cached value
   alone is insufficient.
6. On the pending-cache rebuild path, every non-proof member inside each
   rebuilt anchor interval has one unique declared `(declaration, type)`
   identity on every side, the same complete ordered identity sequence and
   member count, and one retained, bijective Mine/Base/Theirs slot.  Formula
   changes in those non-proof members are ordinary cell differences and are
   retained in the exact result; they do not make a blank duplicate column
   contentful.  The already-exact top-cache path has already proven this
   complete slot-level mapping and may not substitute child-gap provenance for
   that proof.
7. A proof slot may retain `blank-column` plus an inherited
   `formula-identity-mismatch` cause only when that mismatch came from the
   same pre-rebuild interval and every non-proof interval member satisfies
   condition 6.  Any formula token in a proof column, any non-proof unresolved,
   ambiguous, structural, reordered, missing, extra, or non-bijective member,
   or any formula cause not isolated to such otherwise resolved members remains
   terminally unresolved.

The proof is all-or-nothing for a duplicate run.  Nonblank proof content,
a changed digest, missing or extra proof or interval member, different
interval/ordinal, reordered anchor or non-proof sequence, invalid row key,
Base disagreement, candidate-builder exception, a non-proof formula-structure
or mapping failure, or any final cache failure leaves the whole Sheet
`UNRESOLVED`, non-actionable, and without physical operation targets.  The
result records the surviving reason(s) so diagnostics distinguish duplicate
identity from row, digest, interval, or column-cache failure.

The implementation constructs no worksheet objects and reads no editable
workbook.  It receives only immutable snapshots and the already-built
column-cache evidence.  Candidate state is scoped to the current parser/file/
topology/mutation generation and is discarded on any refresh.  Pure 2-way and
3-way tests cover both the exact-top-cache and pending-rebuild stage-two paths,
including asymmetric Mine/Base versus Theirs/Base child gaps whose already
exact top cache is accepted, formula-different non-proof cells that remain
visible, and crosswired, structural, or non-retained top-cache negatives before
this rule is allowed to affect a real revision harness.

## Risks / Trade-offs

- [Risk] Replacing worksheet objects with snapshots can miss uncommon formula or cell types. → Keep paired openpyxl read-only streams as the reference decoder and add formula/array/data-table/date/error/external-link fixtures before any direct XML fast path.
- [Risk] Schema markers are duplicated or edited. → Validate uniqueness on every side, include all declared identity components, preserve order changes, and fall back conservatively with visible ambiguity.
- [Risk] Virtualization can break selection, row actions, minimap jumps, or copied text. → Maintain logical identities independent of visible line numbers and add first/middle/last-window GUI tests for every navigation/action entry point.
- [Risk] Replacing viewport text can still cause scroll jitter. → Coalesce events, preserve logical top/selection, render fixed-size batches, and test thumb dragging plus high-rate wheel input.
- [Risk] Operation overlays diverge from the eventual workbook. → Revalidate overlay operations against the save target generation and retain package/reopen validation as the final authority.
- [Risk] A full rewrite is too large for one safe switch. → Ship behind an internal feature flag, land Oracle and virtualization in staged tasks, and keep the legacy engine as rollback until all gates pass.
- [Risk] Existing user changes overlap implementation areas. → Inspect the dirty diff first, avoid reverting unrelated hunks, and keep edits narrowly scoped.
- [Risk] Whole-workbook status scanning can compete with the selected Sheet. → Always preempt/reprioritize for the selected Sheet, use bounded workers, and resume background Sheets only after selected exact readiness.
- [Risk] Users may act on a visually stable but stale generation. → Cover non-ready/stale content with the calculation state, centralize readiness guards, show a modal on attempted mutation, and never queue the rejected action.
- [Risk] Corpus self-comparison does not exercise changed-cell rendering or save. → Pair it with deterministic disposable mutations on the slowest Sheets and run the full save/reopen/Excel compatibility matrix.

## Migration Plan

1. Capture reproducible legacy correctness/performance manifests without changing runtime behavior.
2. Add snapshot and semantic comparison types behind a disabled feature flag and prove Oracle parity.
3. Enable the new comparison path for selected large Sheets while retaining legacy save/mutation fallback.
4. Add virtual viewport rendering for both legacy and new snapshots, then make it the default for large result sets.
5. Add operation overlays and bounded mutation refresh, preserving structural fallback and save validation.
6. Run the full real/synthetic, GUI, merge, formula, XLSM, save, and Excel-reopen matrix; switch the feature flag default only after all gates pass.
7. Keep one release-level rollback switch. Reverting the flag restores the legacy path without migrating workbook data.

## Open Questions

- The exact row/byte threshold for choosing snapshot/virtual mode will be calibrated from the new baseline; correctness must not depend on the threshold.
- Direct OOXML decoding may be evaluated later, but is not required for the first accepted implementation.
# Applied implementation notes

The snapshot engine is enabled only for selected Sheets of at least 2,000
rows.  It uses a conservative unresolved gate and legacy fallback; this is a
correctness decision, not a relaxation of any performance threshold.  The
initial-sheet path uses read-only catalog handles and defers editable workbook
materialization until an edit/save demand.  Actual real READY timings remain
recorded in validation evidence and are not treated as passing the original
target.
