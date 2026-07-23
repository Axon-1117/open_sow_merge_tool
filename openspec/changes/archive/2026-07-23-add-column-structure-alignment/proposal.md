## Why

The comparer currently matches columns only by physical position. Inserting or deleting a column therefore shifts every following value, producing cascades of false differences and false conflicts, while users have no column-level adopt operation that can preserve workbook structure safely.

## What Changes

- Detect and align inserted, deleted, and retained logical columns before cell comparison in 2-way and 3-way modes.
- Represent column insert/delete operations explicitly so comparison, presentation, undo, and final save share one structural model.
- Show column-level difference blocks and allow users to adopt or retain an inserted/deleted column range without treating every shifted cell as an unrelated edit.
- Extend 3-way conflict rules so independent column structure changes merge cleanly and overlapping structural/value changes are reported conservatively.
- Preserve formulas, styles, widths, hidden state, comments, hyperlinks, validation, merged cells, external links, macros, and other advanced workbook content through the existing native-save safety policy.
- Add real-workbook and synthetic regressions for multiple inserted/deleted columns, column-local value edits, formulas, large sheets, undo, and save/reopen fidelity.

## Capabilities

### New Capabilities

- `column-structure-alignment`: Defines logical column alignment, column structural differences and conflicts, column-level adoption, undo, and fidelity-preserving save behavior.

### Modified Capabilities

None.

## Impact

- Primary implementation areas: row/cell comparison caches, `SheetView` column mappings and hit testing, 2-way/3-way conflict scanning, manual operation records, undo, Excel COM/native OOXML replay, and only-diff/minimap presentation in `sow_merge_tool.py`.
- Test impact: new focused logic, GUI, performance, and real-file replay coverage; existing row-alignment, formula-cache, sheet-level, XLSM, only-diff block, minimap, and save regressions must remain green.
- No CLI argument, settings-schema, workbook-format, or external dependency changes are intended.
