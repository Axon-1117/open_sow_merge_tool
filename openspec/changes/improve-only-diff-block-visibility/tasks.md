## 1. Stable Difference-Block Model

- [x] 1.1 Add an internal block record and cache that groups complete `_full_display_rows` snapshots by consecutive logical pair index and exposes `pair_idx -> block_idx` lookup.
- [x] 1.2 Key and invalidate the block cache with only-diff snapshot and row-model versions so structural edits cannot leave stale block targets.
- [x] 1.3 Compute dynamic pending/processed status from existing in-memory visual-diff maps while preserving stable snapshot block numbering.
- [x] 1.4 Add focused unit tests for separated pair ranges, touched resolved rows, structural slots, and blocks beyond an 800-row render limit.

## 2. Block Presentation and Status

- [x] 2.1 Add a fixed-width toolbar indicator for calculating, empty, active/total, pending, and processed block states.
- [x] 2.2 Resolve the active block from explicit selection first and the top visible pair otherwise, without allowing cell hover to cause toolbar flicker.
- [x] 2.3 Add synchronized block-start spacing tags to all active data and row-number panes without inserting synthetic text lines.
- [x] 2.4 Add a compact textual block marker to the left row-number area and preserve it through normal rendering, row-only refresh, hover-arrow display, cache replay, and incremental row loading.
- [x] 2.5 Hide all block-only presentation when only-diff mode is disabled and restore it immediately when only-diff mode is enabled with precise data ready.

## 3. Navigation and Region Actions

- [x] 3.1 Replace rendered-line-only previous/next state with full block-model navigation for toolbar buttons and existing keyboard shortcuts.
- [x] 3.2 Add a cached-row materialization path that can reveal a target block beyond the current render limit without worksheet rescan or horizontal scroll reset.
- [x] 3.3 Make region-level mine/theirs/Base adoption resolve scope from the stable block model, skip already resolved members, and retain existing formula, undo, progress, and batch structural-copy behavior.
- [x] 3.4 Verify selection restoration and active block status after single-cell, single-row, region, insert/delete, undo, and only-diff toggle operations.

## 4. GUI, Performance, and Real-File Validation

- [x] 4.1 Add 2-way GUI coverage for multiple compacted blocks, visible separators, exact block count, selection/viewport active state, and endpoint button states.
- [x] 4.2 Add 3-way GUI coverage for Base-only differences, structural rows, synchronized spacing across six panes, and processed block status.
- [x] 4.3 Add a large-sheet GUI regression with blocks before and after the initial 800 rendered rows and verify cross-limit navigation without `rescan=True`.
- [x] 4.4 Replay copies of representative workbooks from `C:\GM15\design\sheets\develop` or `release`, including a workbook with multiple regions, and record block count, navigation correctness, region-adoption scope, and timing.
- [x] 4.5 Run compile, smoke, only-diff, minimap, hover/C-area, row-alignment, formula-cache, and region-copy regressions; confirm block updates add no worksheet reads during scrolling or selection.
- [x] 4.6 Update release notes and version/build metadata when the implementation is approved for packaging.
