## Why

In "只看差异" mode, rows from separate worksheet locations are compacted into adjacent screen lines with no visible boundary or exact block count. Users cannot tell how many independent difference regions exist, which region is active, or whether navigation and region-level adoption cover the same rows.

## What Changes

- Introduce a precise, cached difference-block model built from the complete only-diff snapshot rather than only the currently rendered rows.
- Add visible spacing and a block marker at each block boundary without inserting synthetic worksheet rows.
- Show the current block, total block count, and pending block count in the Sheet toolbar.
- Make previous/next difference navigation work across the full snapshot, including blocks beyond the large-sheet initial render limit.
- Keep block numbering stable after rows are adopted, while marking fully resolved blocks as processed.
- Use the same block model for display, navigation, and region-level adoption so their boundaries cannot disagree.
- Preserve existing 2-way and 3-way comparison, save, formula, row-alignment, minimap, and CLI semantics.

## Capabilities

### New Capabilities
- `only-diff-block-navigation`: Defines exact difference-block grouping, visual boundaries, stable block status, full-snapshot navigation, and region-action consistency in only-diff mode.

### Modified Capabilities

None.

## Impact

- Primary implementation area: `SheetView` only-diff snapshot, render tagging, row-number gutter, block navigation, and region-copy handlers in `sow_merge_tool.py`.
- GUI tests must cover 2-way and 3-way views, large sheets with more than 800 rendered difference rows, structural differences, resolved/touched rows, and keyboard navigation.
- No external API, CLI argument, settings schema, workbook format, or packaging dependency changes are expected.
