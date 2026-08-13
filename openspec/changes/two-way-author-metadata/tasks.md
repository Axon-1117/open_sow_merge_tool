## 1. 2-way identity wiring

- [x] 1.1 Construct and resolve the raw 2-way launch context before file normalization.
- [x] 1.2 Pass the resolved context into `SowMergeApp` without changing comparison paths or 3-way behavior.
- [x] 1.3 Render 2-way identity labels and hover details with author/revision metadata and safe fallbacks.

## 2. Regression coverage

- [x] 2.1 Add tests for resolved 2-way Mine/older-side author metadata and unavailable-author fallback.
- [x] 2.2 Run compile, focused metadata/UI tests, and OpenSpec strict validation.

## 3. Release

- [x] 3.1 Update version/build metadata and release notes.
- [x] 3.2 Build and verify the executable and release archive.
