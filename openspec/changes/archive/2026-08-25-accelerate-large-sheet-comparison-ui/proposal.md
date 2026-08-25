## Why

Large project workbooks under `C:\GM15\design\sheets\develop` contain individual Sheets from roughly 3,000 to more than 20,000 rows. The current application repeatedly materializes full value/formula workbooks and lets rendered Tk text and tag counts grow with the number of displayed differences, causing high startup memory, multi-second exact comparison, slow large merge actions, and severe scrolling stalls once about 1,000 or more changed rows are visible.

## What Changes

- Add an exact, per-Sheet large-workbook comparison path that lazily streams only the requested worksheet, constructs a compact immutable semantic model, and avoids retaining duplicate full-workbook value/formula objects for comparison.
- Add a reproducible real-workbook baseline and old/new exact-difference Oracle covering cell values, formulas, row/column structure, 2-way and 3-way results, operation outcomes, memory, and UI responsiveness.
- Prefer declared schema identities such as `@id` and `@const` (including composite identities and keyed continuation groups) for linear-time record alignment, with bounded deterministic fallback and explicit ambiguity instead of unbounded fuzzy matching.
- Virtualize large full and only-difference result views so Tk holds only the visible logical window plus bounded overscan, regardless of whether the Sheet contains 1,000 or 20,000 changed rows.
- Precompute render text and difference spans off the Tk thread, coalesce scroll events, batch widget mutations, and guarantee that scrolling performs no workbook reads or comparison work.
- Apply cell/row/region operations to an in-memory operation overlay and update only affected comparison rows, blocks, and visible viewport state; preserve structural invalidation only for row/column/Sheet topology changes.
- Preserve current atomic save, native Excel structural replay, formula-cache handling, package validation, undo/redo, stale-generation rejection, and Excel reopen guarantees.
- Compute and expose an exact whole-workbook Sheet status map: every supported Sheet is explicitly pending, calculating, exact-same, exact-changed, unresolved, or failed, and no pending/provisional result is presented as final.
- Block copy, overwrite, accept, region, structural, undo/redo, and save operations whenever the relevant Sheet generation is not exact-ready; show a prominent persistent state banner and an immediate modal explanation when the user attempts a blocked operation.
- Add a recursive, read-only corpus benchmark for every supported Excel file under `C:\GM15\design\sheets\develop`, covering 2-way and 3-way startup plus every Sheet's request-to-exact-ready time, and optimize until every supported Sheet completes within the user-accepted 15-second maximum on the reference workstation.
- Treat exact comparison, exact operation targeting, package validity, atomic recovery, and Excel/XLSM reopen compatibility as release-blocking constraints that cannot be traded for speed.
- Add a release-blocking changed-revision gate using disposable copies of `Dungeon.xlsx` revisions 39265 and 39264: both `Dungeon` and the measured `MonsterGroup` outlier must publish exact 2-way/3-way results within 15 seconds while hover, wheel, thumb drag, and tab changes remain responsive.
- Make every view-only callback (hover, click inspection, wheel/page/thumb/minimap navigation, tab activation, and viewport publication) strictly snapshot/render/cache-only. Missing view data is represented by a non-actionable placeholder; it must never trigger editable-workbook materialization.
- Run remaining-Sheet work only after a short UI quiet window, cooperatively yield/preempt it for interaction and selection, and retain explicit diagnostics for heartbeat gaps, CPU/RSS, edit-loader callers, and C-area render cost.
- Exclude per-cell author attribution and SVN history reconstruction from this change.

## Capabilities

### New Capabilities

- `large-sheet-semantic-comparison`: Lazy per-Sheet OOXML ingestion, stable schema-aware alignment, exact 2-way/3-way comparison, Oracle parity, and bounded performance behavior for large Sheets.
- `virtualized-difference-rendering`: Fixed-cost viewport rendering, smooth scrolling, navigation, selection, and result presentation for extremely large difference sets.
- `large-sheet-operation-performance`: Operation overlays and bounded recomputation/rendering for cell, row, region, column, Sheet, undo, and redo workflows without weakening save fidelity.

### Modified Capabilities

- `only-diff-loading-readiness`: Exact only-difference publication and readiness must use the compact comparison snapshot and virtualized result window without materializing all result rows in Tk.

## Impact

- Primary implementation: `sow_merge_tool.py`, especially workbook loading, Sheet cache construction, row/column comparison, `SheetView` rendering/scrolling, operation refresh, and save staging integration.
- New reproducible performance/Oracle harnesses using disposable fixtures derived from the real workbooks in `C:\GM15\design\sheets\develop`; source workbooks remain read-only and unchanged.
- Existing smoke, GUI, formula-cache, row/column alignment, only-diff, SVN merge, XLSM, undo/redo, and native-save suites remain required regression gates.
- New whole-corpus evidence records every discovered workbook and Sheet, 2-way/3-way timings, final exact state, engine/fallback reason, Oracle result, peak memory, and the slowest-Sheet ranking; unsupported or failed files are reported explicitly rather than skipped.
- No workbook format change, physical Sheet split, external database migration, command-line breaking change, or new author-data dependency.
- New real changed-revision evidence records `Dungeon`/`MonsterGroup` final exact timing, viewport P95, heartbeat, CPU/RSS, edit-loader reason/caller counts, C-area render cost, interaction traces, and Oracle/operation/save compatibility outcomes.
