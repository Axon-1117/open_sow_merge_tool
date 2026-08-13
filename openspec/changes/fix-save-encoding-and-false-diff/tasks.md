## 1. Root-cause implementation

- [x] 1.1 Replace locale-dependent PowerShell operation JSON reading with explicit UTF-8 decoding in the Excel native replay script.
- [x] 1.2 Normalize exact empty-string values with absent cells at the shared comparison-key boundary while preserving formula identity semantics.

## 2. Regression coverage

- [x] 2.1 Add focused 2-way and 3-way Unicode save tests covering cell-only XML patching, structural fallback output, package validation, and exact text round-trip.
- [x] 2.2 Add comparison tests for `None`/`""` equivalence, the reported `GunshipsMaster` coordinate set, and conservative formula/non-empty text differences.
- [x] 2.3 Run compile, focused save/diff tests, existing row/column/3-way regressions, and OpenSpec validation.

## 3. Release

- [x] 3.1 Update version/build metadata and release notes after the fixes pass validation.
- [x] 3.2 Build the executable, synchronize the local publish directory, and verify hashes/package integrity.

## 4. Missing worksheet dimension regression

- [x] 4.1 Recover missing read-only worksheet dimensions through an actual XML scan before computing comparison bounds.
- [x] 4.2 Add a regression fixture that removes `<dimension>` while preserving populated cells, and verify the recovered bounds.
- [x] 4.3 Run focused regression coverage and rebuild/publish the executable.
