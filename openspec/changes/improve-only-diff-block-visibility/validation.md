## Validation Record

Date: 2026-07-22

### Focused behavior

- Added `_gui_self_test_diff_blocks.py` covering 2-way and 3-way block counts, markers, synchronized spacing, processed status, stale-hover region anchoring, undo, horizontal scroll preservation, and navigation beyond the initial 800 rendered rows.
- Added cached-model coverage to `_smoke_test_review_regressions.py`, including stable numbering, structural/touched rows, unrendered blocks, and a guard that fails on any app/workbook access during cached block updates.
- Cross-limit navigation materializes the complete target block from `_full_display_rows` and does not call `refresh(..., rescan=True)`.

### Real workbook replay

Workbook source: `C:\GM15\design\sheets\develop\WorldMonster.xlsx`

- Base: SVN `r36162`, exported read-only through TortoiseSVN.
- Theirs: a copy of the current working-copy workbook (working copy revision `36321`, file changed revision `36163`).
- Mine: a package-preserving copy of Base with an independent local edit at `WorldMonster@design!B3000`.
- All generated workbooks were isolated under `tmp/test_tmp`; the source workbook was not modified.

Measured result on `WorldMonster@design`:

- Window construction: 6.347 seconds.
- Precise only-diff calculation: 10.317 seconds.
- Exact snapshot: 1,201 difference rows in 2 logical pair-contiguous blocks (`1400-2599`, `2999`).
- Cross-block cached navigation: 6.775 seconds, with `rescan=False` only.
- Region adoption: 1,200 rows / 8,955 recorded cells in 7.178 seconds.
- Formula handling: one dependency notice was generated; formulas were retained while theirs cached results were adopted.
- Accuracy: the adjacent block remained unchanged and the independent mine edit at row 3000 was preserved.

### Regression matrix

Passed compile and all targeted suites:

- Core smoke, review regressions, 2-way row replay, 3-way row alignment, large 3-way open/only-diff, large structural row insertion, tail append split, sheet-level operations, and XLSM support.
- Formula-cache save/undo, Base insertion/alignment, minimap, only-diff toggle, main diff-cell rendering, hover/C-area selection, row-header width, horizontal synchronization, progress feedback, initial cache padding, sheet cache reuse, and 3-way sheet state.

### Packaging

Packaging was approved after implementation validation. Release metadata was advanced to `2026-07-22.update54` / `new131-onlydiff-block-navigation-fix`; the packaging workflow must publish and verify these exact identifiers.
