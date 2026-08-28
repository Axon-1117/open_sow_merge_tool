## 1. Automatic Editable Preload

- [x] 1.1 Add an app-level automatic preload request that accepts only the selected current full-detail immutable view and reuses the existing single-owner loader.
- [x] 1.2 Trigger the automatic request from the centralized Sheet lifecycle transition into editable-deferred readiness without changing explicit mutation/save fallback behavior.
- [x] 1.3 Preserve post-load view readiness publication and loader-failure retry eligibility.

## 2. Regression Coverage

- [x] 2.1 Add focused lifecycle tests proving partial, hidden, unresolved, stale, and already-loading views do not create an automatic owner.
- [x] 2.2 Add a GUI regression proving an exact initial view starts exactly one background loader and the first row overwrite after readiness does not show the editable-loading modal.
- [x] 2.3 Retain coverage proving an action during `EDIT_LOADING` is rejected with no write, queue, or automatic replay.

## 3. Verification

- [x] 3.1 Run Python compile checks and the focused automatic-preload/loading-readiness tests.
- [x] 3.2 Run the relevant smoke/GUI row-overwrite regression and confirm no source workbook is modified.
- [x] 3.3 Validate the OpenSpec change strictly and review the scoped diff against the pre-existing dirty worktree.
