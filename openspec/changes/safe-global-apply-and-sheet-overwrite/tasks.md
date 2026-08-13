## 1. Global preflight and user feedback

- [x] 1.1 Classify unresolved columns by actual exact cell/formula differences and allow equal unresolved no-op columns to be skipped.
- [x] 1.2 Collect blocking ambiguous columns with causes, physical mappings, and affected cell samples.
- [x] 1.3 Add a modal blocker dialog with cancel behavior and explicit left/right whole-Sheet choices with risk confirmation.

## 2. Whole-Sheet overwrite integration

- [x] 2.1 Route confirmed whole-Sheet choices through the existing Sheet-copy/manual-operation bookkeeping for both 2-way directions.
- [x] 2.2 Refresh/invalidate projections and preserve one-step undo/rollback semantics after Sheet replacement.
- [x] 2.3 Keep 3-way direction restrictions and missing-Sheet behavior unchanged.

## 3. Regression coverage

- [x] 3.1 Add 2-way tests for equal blank/duplicate unresolved columns, real ambiguous differences, blocker details, cancellation, and both overwrite directions.
- [x] 3.2 Add save/reopen tests proving whole-Sheet overwrite preserves Chinese text and produces an Excel-valid package without repair signals.
- [x] 3.3 Run focused GUI/save tests, compile checks, and OpenSpec validation.

## 4. Release

- [x] 4.1 Update version/build metadata and release notes.
- [x] 4.2 Build and verify the executable and release archive.
