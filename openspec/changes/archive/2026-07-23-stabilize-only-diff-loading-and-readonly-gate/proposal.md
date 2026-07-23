## Why

Large three-way workbooks can become unresponsive when the user enables only-diff while editable workbook preload is finishing: the edit-ready callback performs a synchronous exact rescan on the Tk thread. The same lifecycle renders a temporary full-sheet page before exact only-diff data is ready, permits mutations against provisional state, and can make a maximized window appear to change state while Windows cannot repaint it.

## What Changes

- Introduce an explicit per-Sheet loading/readiness lifecycle with generation-safe background results.
- Keep a Sheet read-only until its row model, exact requested comparison, logical column mapping, and editable workbooks are ready.
- Make only-diff transitions explicit: accept one request, lock the control while calculating, show immediate status feedback, and atomically publish the exact result.
- Show a cancellable modal progress dialog for a user-triggered exact-only-diff calculation; while it is active, block the current Sheet and Sheet switching.
- Keep the only-diff checkbox at one fixed toolbar location and present changing difference-block counts in a separate stable status area.
- Let the C-area row comparison consume the full available lower-pane width so it exposes more logical columns than any single main pane.
- Eliminate large-Sheet `refresh(rescan=True)` and exact comparison work from the Tk thread.
- Reuse formula-aware background cache data for exact only-diff instead of reopening and rescanning the same workbooks.
- Prioritize the active Sheet, prevent competing exact workers, and cache hidden-Sheet results without rendering them.
- Preserve the main-window state across loading, mode transitions, and asynchronous result application.
- Add correctness, responsiveness, stale-result, mutation-gate, and real WorldMonster regression coverage.

## Capabilities

### New Capabilities

- `only-diff-loading-readiness`: Defines exact only-diff publication, per-Sheet readiness and mutation gates, background scheduling, progress feedback, and window-state stability.

### Modified Capabilities

None.

## Impact

- Primary implementation area: `SowMergeApp` background cache/preload scheduling and `SheetView` only-diff lifecycle, controls, mutation handlers, and rendering in `sow_merge_tool.py`.
- GUI and smoke tests will cover every write entry point, stale generations, hidden-Sheet completion, active-Sheet priority, main-thread heartbeat, maximized-window preservation, C-area viewport width, fixed control geometry, modal progress, cancellation, and Sheet-switch locking.
- Real-file acceptance will use isolated copies of `WorldMonster.xlsx`; no user workbook will be modified.
- No CLI, workbook format, persisted merge output, or external dependency change is expected.
