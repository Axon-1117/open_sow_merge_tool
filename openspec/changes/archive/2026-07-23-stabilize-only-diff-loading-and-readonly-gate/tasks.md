## 1. Readiness lifecycle and mutation safety

- [x] 1.1 Add per-Sheet lifecycle generation/state helpers and a single mutation-readiness predicate.
- [x] 1.2 Drive row, region, cell, column, missing-Sheet, undo, refresh, and save entry points through the mutation guard.
- [x] 1.3 Refresh all affected widget states atomically when the selected Sheet lifecycle changes.
- [x] 1.4 Keep viewing, selection, copying, scrolling, Sheet switching, and diagnostics available in read-only states.

## 2. Only-diff transition behavior

- [x] 2.1 Represent only-diff as a requested mode without temporarily changing the user-facing checkbox to full mode.
- [x] 2.2 Lock the checkbox and show immediate calculating/read-only feedback while an exact snapshot is pending.
- [x] 2.3 Atomically publish matching exact rows and unlock controls only after full readiness.
- [x] 2.4 Handle failure, cancellation, retry, and repeated-toggle attempts without delayed mutations.

## 3. Background cache and scheduling

- [x] 3.1 Add explicit formula, row-model, pair-diff, Base-diff, and only-diff completeness metadata to Sheet caches.
- [x] 3.2 Stop marking large existence-only caches as full exact pair-difference maps.
- [x] 3.3 Reuse a known only-diff request in the first formula-aware background pass to avoid a second workbook scan.
- [x] 3.4 Coordinate one priority exact worker, invalidate stale generations, and prefer the selected Sheet.
- [x] 3.5 Cache hidden-Sheet exact results without rebuilding visible Tk widgets.

## 4. Main-thread and window stability

- [x] 4.1 Remove edit-ready `refresh(rescan=True)` calls and replace them with bounded prepared-data/readiness application.
- [x] 4.2 Ensure large-Sheet comparison, formula reconciliation, and exact only-diff generation never execute on the Tk thread.
- [x] 4.3 Keep the promoted startup root hidden until layout is ready, then apply its startup window state after idle.
- [x] 4.4 Add lightweight timing, heartbeat, and window-state diagnostics for loading and only-diff transitions.

## 5. Automated verification

- [x] 5.1 Add lifecycle and direct-handler tests proving no mutation path can bypass a non-READY Sheet.
- [x] 5.2 Add only-diff transition tests for feedback, locked toggles, exact publication, stale generations, failure, and retry.
- [x] 5.3 Add scheduling/cache tests for one exact worker, active-Sheet priority, hidden result caching, and completeness flags.
- [x] 5.4 Extend GUI progress tests to assert Tk heartbeat, zero edit-ready large rescan, and maximized/normal window preservation.
- [x] 5.5 Run existing comparison, row/region/column, formula, undo, save, and packaging-relevant regression suites.

## 6. Real-file acceptance and OpenSpec validation

- [x] 6.1 Replay isolated WorldMonster Mine/Base/Theirs copies with only-diff requested during editable preload.
- [x] 6.2 Verify callbacks and heartbeat meet the responsiveness gates and source workbook hashes remain unchanged.
- [x] 6.3 Verify final only-diff rows/cells match the exact oracle and READY operations retain save/undo correctness.
- [x] 6.4 Run strict OpenSpec validation and record implementation/validation evidence.

## 7. Follow-up UX acceptance gaps

- [x] 7.1 Expand the C-area comparison body/header to the full lower-pane width and add viewport-width plus horizontal-alignment regression coverage.
- [x] 7.2 Add a generation-safe modal exact-only-diff progress dialog with measurable progress, an explicit Cancel action, and stale-result protection.
- [x] 7.3 Lock the owner Sheet, notebook tabs, bottom Sheet navigation, and current-Sheet interaction while the modal calculation is active; restore them on every exit path.
- [x] 7.4 Keep the only-diff checkbox label and geometry fixed and move dynamic difference-block information into a separately reserved status area.
- [x] 7.5 Replay automated visual/interaction acceptance, including C-area visible-column count, checkbox coordinates in every state, repeated click rejection, tab lock, Cancel, responsiveness, and isolated WorldMonster behavior.

## Validation evidence

- `python _gui_self_test_loading_readonly_gate.py`: 7/7 lifecycle, direct-handler, failure/retry, stale generation, serial broker, saved-baseline, missing-Sheet, hidden-cache, and priority assertions passed.
- `python _gui_self_test_progress_feedback.py`: heartbeat, zero edit-ready rescan, maximized startup, and normal-window preservation passed.
- Comparison and mutation regressions passed: `_smoke_test.py`, `_gui_self_test_diff_blocks.py`, `_gui_self_test_region_mode_interaction.py`, `_gui_self_test_logical_column_actions.py` (31/31), `_smoke_test_formula_cache_undo.py`, `_smoke_test_2way_formula_cache_save.py`, `_smoke_test_sheet_level_ops.py`, `_smoke_test_large_3way_only_diff.py`, and `_smoke_test_large_only_diff_row_insert.py`.
- Isolated WorldMonster Mine/Base/Theirs baseline passed with only-diff requested before editable preload completed and ordinary (non-modal) preload Sheet switching: callback `85.83 ms`, maximum heartbeat gap `109.23 ms`, one exact difference row, unchanged maximized state/geometry, valid saved package, correct undo/reapply/reopen results, and unchanged source/input SHA-256 hashes.
- `_gui_self_test_c_hscroll_bar.py` now verifies that C occupies the lower-pane width, exposes a wider viewport than one main pane, keeps header/body widths aligned, follows main scrolling without letting its own scrollbar move the main panes, and passed.
- `_gui_self_test_only_diff_progress_modal.py` passed real widget invocation, fixed checkbox geometry, modal grab, measurable progress, repeated-click rejection, notebook and bottom-navigation locking, Cancel restoration, stale-generation rejection, and successful retry/publication.
- Isolated real WorldMonster cancel/retry cold-path acceptance passed twice after prebuilding the reusable modal dialog. Success callbacks were `31.30/43.96 ms`, dialogs visible at `63.72/77.05 ms`, and maximum heartbeat gaps `132.29/144.78 ms`. Cancel at `27.33%` responded in `27.15/32.08 ms`, stopped the broker in `22.25/21.17 ms`, restored the stable full view and Sheet switching, and rejected stale publication. Retry reached `0–100%`, produced one exact difference row, rejected Sheet switching while active, kept checkbox drift at `0 px`, and retained C/main viewport widths `1483/705 px`.
- Full compatibility regressions passed after the follow-up: loading/read-only gate `7/7`, logical column actions `31/31`, progress/window feedback, difference blocks, region interactions `6/6`, C-area cell alignment, only-diff, Sheet diff state, formula/save/undo/Sheet-level smoke tests, and large 2-way/3-way only-diff tests.
- Final `2026-07-23.update57` / `new134-only-diff-loading-progress-ux` packaged EXE launched with the isolated WorldMonster Mine/Base/Theirs inputs, completed PyInstaller extraction and workbook loading, promoted the Tk root to maximized state, and reached `READY`. The release, dist, and shared-publish EXE copies all match SHA-256 `6C48B940C26985F5CE5E3DE30BE2297CBB890659E4E511FDA5DB5E2440985A6C`; the release ZIP and shared-publish ZIP match `994BC57126A6F28E55D6BD359BF7C1A4EE731F15E08F2374FC41325F776A6A44`.
- `openspec validate stabilize-only-diff-loading-and-readonly-gate --type change --strict --json --no-interactive`: 1 change passed, 0 failed.
- `git diff --check`: passed (line-ending conversion warnings only).
