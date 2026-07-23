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

## Independent Acceptance Follow-up

Date: 2026-07-22

The feature, performance, test, and UX roles replayed the shipped change against its scenarios instead of relying on the completed task checkboxes. The first pass found four gaps:

- block keyboard shortcuts were bound only in conflict mode and returned a tuple from their lambdas;
- 3-way selection and cross-limit navigation rebuilt Base text through worksheet reads;
- cross-limit materialization appended data panes but not row-number panes, so the target block marker was not visible;
- value-empty coordinate padding before a real one-sided tail append entered only-diff as structural rows.

The follow-up implementation now:

- binds `Ctrl+N`, `Ctrl+P`, and `Shift+F4` in ordinary 2-way/3-way and conflict views through handlers that return `"break"`, while plain `F4` retains hover-panel pinning;
- builds and invalidates `pair_text_base` alongside mine/theirs caches, keeping selection and navigation at `Worksheet.cell=0`, `iter_rows=0`, and `rescan=0` after precise data is ready;
- batch-appends left/Base/right row-number text with real row numbers, `[block]` markers, `blockstart` tags, and synchronized vertical views;
- normalizes one-sided tail alignment by removing only the value/formula-empty coordinate prefix before a later real tail append, while retaining internal blank insertions and uncached formulas.

Strict real-file replay used isolated copies derived from `C:\GM15\design\sheets\develop\WorldMonster.xlsx`:

- window construction: 3.357 seconds;
- precise only-diff calculation: 1.819 seconds;
- exact snapshot: 854 rows in 4 blocks (`99-548`, `799-1199`, `1999-2000`, `2999`);
- cross-limit navigation: 6.201 seconds, `rescan=False`;
- all active data and row-number panes contained the target line, the target marker/tag was visible, vertical views were synchronized, and horizontal position was preserved;
- region adoption changed only block 1 (450 rows), preserved the adjacent block and the independent mine edit, and completed in 4.527 seconds.

The isolated Guide fixture now reports `blank_diff_count=0`; the final structural block contains only the real theirs rows 150/151. Focused tests retain internal one-sided blank rows and formulas with missing cached values.

Compile, strict OpenSpec validation, diff-block, review, row-alignment, Base-insert, large only-diff, minimap, horizontal synchronization, row-header, hover/C-area, formula-cache, bounds, and Guide padding regressions passed. A terminal-only Excel COM replay remains environment-dependent (`HRESULT 80070520` in a non-interactive logon session) and is covered separately by the successful interactive/real-output checks.

Acceptance result: the change is ready to archive. Cross-limit navigation remains perceptibly slow at about 6.2 seconds on this real fixture; viewport virtualization or direct-window materialization remains a performance follow-up rather than an unverified correctness claim.
