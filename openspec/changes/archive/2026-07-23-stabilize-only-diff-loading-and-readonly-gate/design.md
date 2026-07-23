## Context

`SheetView` currently uses independent booleans such as `_data_ready`, `_diff_partial`, `_only_diff_async_building`, and global editable-workbook readiness. They do not form one interaction contract: loading UI can still expose mutation entry points, only-diff may show a temporary full page, and edit preload completion calls `refresh(rescan=True)` for every materialized view on the Tk thread.

The 2026-07-23 WorldMonster reproduction exposed a deterministic race. The user enabled only-diff at 12:14:39.817 while Base edit preload was still running; preload completed at 12:14:40.627 and the edit-ready UI callback started `refresh(rescan=True, only_diff=True)` at 12:14:40.697. The Tk thread returned at 12:14:52.943, producing 12.246 seconds without event processing. A prior run where the same click occurred after edit preload used the asynchronous path and completed in about 3.35 seconds.

Large-Sheet cache computation already opens value and formula-aware read-only workbooks and builds row/column alignment, but it deliberately does not publish a complete row-difference map. Cache application nevertheless marks the map exact, then an only-diff request can reopen the same workbooks for a second scan. Hidden exact workers can also compete with the selected Sheet and render after they finish.

## Goals / Non-Goals

**Goals:**

- Make readiness and mutation permission explicit per Sheet.
- Keep every mutation path disabled until exact requested comparison data and editable workbooks are ready.
- Keep Tk callbacks responsive and perform no large-Sheet rescan on the Tk thread.
- Give an only-diff request immediate, unambiguous feedback, lock repeated toggles, and publish the exact view atomically.
- Provide determinate progress and an explicit Cancel action for a user-triggered exact-only-diff calculation.
- Prevent current-Sheet interaction and Sheet switching while that modal calculation is active.
- Keep the only-diff control position stable and use a separate status strip for dynamic difference information.
- Use the full lower-pane width for the C-area comparison viewport.
- Reuse the first formula-aware background pass when only-diff is already requested.
- Run at most one priority exact-only-diff worker and prefer the selected Sheet.
- Prevent hidden-Sheet completion from rebuilding visible Tk widgets.
- Preserve the user's main-window state through startup and asynchronous transitions.

**Non-Goals:**

- Changing row/column alignment semantics, conflict classification, formulas, save/replay behavior, or workbook output.
- Making approximate or partial difference rows writable.
- Parallelizing multiple XML scans; openpyxl parsing is CPU/GIL and I/O intensive, so uncontrolled concurrency is counterproductive.
- Adding a new persistent settings schema or CLI option.

## Decisions

### Use one per-Sheet lifecycle and one mutation guard

Each `SheetView` owns a lifecycle generation and state:

`LOADING -> DIFFING -> EDIT_LOADING -> READY -> BUSY`

Failures and cancellations enter `FAILED` or `CANCELED`; shutdown enters `CLOSING`. A central `can_mutate()` predicate additionally requires current row/column generations, an exact requested comparison, editable workbook readiness, and no active operation.

Both widget state and every command/event handler call the same guard. This is preferred over button-only disabling because row headers, main panes, C-area bindings, keyboard shortcuts, save, and undo are independent mutation entry points.

### Treat only-diff as a request with a locked, cancellable transition

A READY full-view Sheet accepts one enable request immediately, leaves the checkbox label and position unchanged, checks and disables it, opens a modal progress dialog within 100 ms, and makes the Sheet read-only. The dialog shows the active stage, processed rows, total rows, percentage, and an explicit Cancel button. It retains the stable current view until the exact snapshot is ready, then swaps the row set once and unlocks controls.

When a Sheet opens with only-diff already requested, the checkbox starts checked and locked. Cache application must not temporarily set the user-facing variable to full mode. If no exact rows are ready, the view shows a clearly labelled read-only preview/loading state rather than pretending the checked mode is complete.

The progress dialog owns the input grab while the user-triggered calculation is active. The selected Sheet is pinned in both the notebook and the bottom Sheet navigation; attempts to switch are rejected until success, failure, or cancellation. Repeated checkbox toggles are rejected while DIFFING. Cancel advances the build generation, immediately closes the dialog, restores the stable full view and full-mode preference, and guarantees that stale worker results cannot render or unlock a later request. Failure restores a stable view, keeps mutations disabled when readiness is invalid, and exposes retry; it never queues a delayed write action.

### Separate stable controls from dynamic status

The only-diff checkbox uses a constant label and a dedicated toolbar slot. Calculation feedback is never appended to the checkbox text. Difference-block ordinal and pending-count text lives in an always-reserved status strip below the toolbar, so checking or unchecking only-diff cannot move the checkbox or neighboring controls.

### Let the C-area use the lower pane width

The C-area header and comparison `Text` widgets expand horizontally with their notebook page. They no longer copy the width of one main pane. Horizontal synchronization keeps the main pane as the canonical logical start, while C's wider viewport reveals additional columns to the right. C's scrollbar reports and changes C's own wider viewport without silently moving the main panes. Tests verify that the C body/header remain aligned and that C displays a larger viewport than one main pane whenever the window provides that space.

### Make background cache completeness explicit

Cache payloads carry independent flags for formula awareness, row-model exactness, Mine/Theirs diff exactness, Base diff exactness, and only-diff snapshot exactness. `_pair_diff_full_exact` is set only when the full pair map was actually produced.

When only-diff is requested before a large Sheet's first background pass begins, that pass uses its already materialized aligned value/formula tuples to compute the complete row-difference maps and exact only-diff indices. When the request arrives after the first pass, the existing dedicated exact worker remains a fallback.

This conditional exact pass avoids imposing full exact comparison cost on every unopened Sheet while eliminating duplicate workbook opens for the selected/default-only-diff Sheet.

### Never rescan a large Sheet on the Tk thread after edit preload

The editable-workbook completion callback only updates readiness and refreshes interaction gates. Formula-aware background caches are authoritative for data computed before edit preload completes. If a visible row presentation still needs formula text, a bounded background reconciliation payload is produced and atomically applied; the callback never calls `refresh(rescan=True)`.

Foreground manual refresh and structural operations still request exact recomputation, but their heavy work must run behind the existing progress/background mechanism before UI application.

### Coordinate one active exact worker

`SowMergeApp` owns the current priority exact request `(sheet, generation)`. Starting a request cancels or invalidates an older hidden request before launching another. Outside a modal user-triggered exact calculation, tab selection reprioritizes queued Sheet cache work. During that modal calculation, tab selection is pinned to the owner Sheet. Exact completion for a hidden Sheet updates its cache/state only; rendering waits until that Sheet is selected.

Generation checks occur both before publishing cache data and before changing widget/readiness state.

### Preserve window state instead of forcing maximization

The promoted startup root stays withdrawn while the main widget tree is built. After idle layout, the application deiconifies and applies the intended startup state once. Only-diff, cache application, and edit-ready callbacks do not call root geometry/state APIs.

Tests capture state and geometry before a transition and require the same state afterward. The application does not blindly maximize after every refresh because that would override a user-selected normal window.

### Use structural responsiveness gates

Correctness is never traded for an empty/approximate only-diff result. Performance acceptance focuses on eliminating duplicate scans and Tk work:

- checkbox/tab callbacks return within 100 ms;
- Tk heartbeat has no gap above 200 ms during background exact work;
- large-Sheet Tk-thread `refresh(rescan=True)` count is zero;
- one priority exact worker exists globally;
- hidden completion performs zero visible Text renders;
- one known-at-start exact request performs at most one formula-aware workbook scan.

## Risks / Trade-offs

- [Risk] Central gating misses an event binding and allows a provisional write. -> Route all mutation handlers through the guard and add direct handler tests, not only widget-state tests.
- [Risk] Disabling controls until editable workbooks are ready increases perceived waiting. -> During ordinary loading keep scrolling, selection, copying, Sheet switching, and status/diagnostic viewing available; only an explicit modal exact-only-diff request temporarily blocks current-Sheet interaction and Sheet switching, with progress and Cancel always visible.
- [Risk] Conditional exact comparison increases the first background pass for a requested large Sheet. -> Skip provisional full rendering and duplicate reopen/rescan, prioritize the selected Sheet, and retain the fast existence-only path for unrequested Sheets.
- [Risk] A stale worker clears the current Sheet's busy state. -> Require matching `(sheet, generation)` for data publication and lifecycle transitions.
- [Risk] Removing edit-ready rescan leaves formula text provisional. -> Mark formula awareness explicitly and reconcile only missing visible presentation data in background.
- [Risk] Startup withdraw/deiconify changes automated GUI timing. -> Keep `root.update()` compatible and extend progress/window tests.
- [Risk] A canceled openpyxl worker cannot stop inside every XML operation. -> Invalidate its generation immediately, check cancellation between bounded row blocks, and prevent it from rendering or unlocking controls.
- [Risk] A modal dialog outlives its Sheet or a stale worker closes a newer dialog. -> Key dialog updates, unlocking, and closure by `(sheet, generation)` and release the grab during Sheet/window destruction.
- [Risk] Expanding C changes horizontal scroll clamping at the right edge. -> Keep the main pane as the canonical logical start and test left, middle, and right synchronization with the wider C viewport.

## Migration Plan

No data migration is required. Introduce lifecycle/gating first, then cache exactness and scheduling, then remove the edit-ready rescan and enable the locked only-diff transition. Keep existing synchronous/manual paths as correctness fallbacks only where they are not reachable from Tk callbacks.

Rollback removes the lifecycle integration and restores the previous callbacks; workbook data and saved output formats are unchanged.

## Open Questions

None required for implementation.
