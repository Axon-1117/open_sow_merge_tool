## Why

Chinese cell content can be corrupted when a save uses the Excel native replay path, and Excel may report a repaired workbook after that save. Separately, the comparison layer treats two visually empty OOXML representations as different, causing false markers in `GunshipsMaster`.

## What Changes

- Decode native replay operation payloads explicitly as UTF-8 for both 2-way and 3-way saves.
- Keep Unicode text, cell types, formulas, and workbook package structure intact through native and fallback save paths.
- Treat an empty string cell and an absent/empty cell as equivalent for value-level diff classification when neither side carries a formula identity.
- Recover read-only worksheet bounds by scanning sheet XML when a valid XLSX omits its optional `<dimension>` element, so populated columns are not reported as deleted.
- Add regression coverage for 2-way/3-way Unicode saves, Excel-reopen/package validation, and the reported `GunshipsMaster` cells.

## Capabilities

### New Capabilities

- `workbook-save-and-diff-fidelity`: Preserve Unicode workbook content during save and classify semantically empty cells consistently.

### Modified Capabilities

- None.

## Impact

- `sow_merge_tool.py` native Excel replay payload loading, read-only worksheet bounds, and value comparison helpers.
- New focused smoke/regression coverage and release validation.
- No new runtime dependency.
