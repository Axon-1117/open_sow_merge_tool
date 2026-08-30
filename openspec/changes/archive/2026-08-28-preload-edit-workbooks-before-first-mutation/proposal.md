## Why

The large-Sheet snapshot path deliberately waits for the first mutation or save before loading editable workbooks, so every fresh session's first row overwrite is rejected even when the selected Sheet is already exact-ready. This makes a completed comparison appear broken and forces a predictable second click.

## What Changes

- Start the single-owner editable-workbook loader automatically after the initial selected Sheet publishes a current, exact, actionable view.
- Keep comparison startup snapshot-first: editable loading must not delay the initial window or selected-Sheet exact publication.
- Keep mutation safety fail-closed while editable loading is still in progress; never queue or replay a rejected row/cell/region/structure operation.
- Refresh mutation readiness when the background preload completes so the next user action can execute immediately.
- Add regression coverage for automatic preload ownership, first row overwrite after preload, and safe rejection during the remaining loading window.

## Capabilities

### New Capabilities

None.

### Modified Capabilities

- `large-sheet-semantic-comparison`: Editable workbooks are warmed automatically after the initial selected Sheet becomes exact-ready instead of waiting exclusively for the first mutation/save demand.

## Impact

- Primary implementation: `sow_merge_tool.py` startup lifecycle, exact-result publication, editable preload ownership, and readiness UI refresh.
- Tests: `_gui_self_test_loading_readonly_gate.py` and focused startup/mutation regressions.
- Performance: selected-Sheet opening remains snapshot-only; background CPU and memory can rise after the exact initial view is available.
