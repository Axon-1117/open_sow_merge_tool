## Context

The save pipeline has three relevant paths: cell-only OOXML patching, Excel COM replay for structural operations, and an openpyxl fallback when Excel COM is unavailable. The native path serializes cell operations to a UTF-8 JSON file, then embeds a PowerShell script that reads and converts that JSON before assigning values to Excel. The current read uses PowerShell's default text encoding, which is not guaranteed to match the producer's UTF-8 bytes.

The comparison pipeline uses `_merge_cmp_value` in quick scans, exact row comparison, column signatures, and 2-way/3-way diff maps. `None` and `""` are currently different keys even though both render as empty in the grid. The reported `GunshipsMaster@design` cells contain exactly this representation mismatch between the compared workbooks.

## Goals / Non-Goals

**Goals:**

- Make native replay JSON decoding deterministic and Unicode-safe on Windows PowerShell.
- Verify saved output remains a valid OOXML package that Excel can reopen without repair.
- Make semantically empty literal cells compare equal in both 2-way and 3-way paths.
- Preserve formula identity checks so a formula, formula cache, or formula-vs-literal change is not hidden.

**Non-Goals:**

- Changing Excel's calculation behavior or formula cache policy.
- Rewriting existing workbook content that is not part of an explicit save operation.
- Treating whitespace-only text as blank.

## Decisions

1. **Read operation JSON with an explicit UTF-8 reader.** Use .NET `File.ReadAllText(path, UTF8)` in the generated PowerShell script before `ConvertFrom-Json`. This removes the locale/code-page dependency while keeping the existing JSON contract and works for both 2-way and 3-way native replay.

2. **Normalize only empty string at the typed comparison boundary.** Make the shared comparison key map `None` and the exact empty string to the same blank key. Keep formula identity comparison ahead of that key, so distinct formulas and formula-vs-literal cases remain differences. Do not trim or normalize non-empty text.

3. **Use package validation plus reopen checks as save gates.** Keep the existing OOXML validator and Excel reopen gate for native replay; add regression assertions that Chinese text survives and the output opens through openpyxl/package inspection for fallback and cell-only paths.

4. **Calculate missing read-only worksheet dimensions before comparison.** OOXML permits omission of `<dimension>`; Excel scans those sheets, whereas openpyxl read-only mode reports both bounds as `None`. Centralize a force-calculated fallback and use it on comparison/cache scan paths so a metadata omission is never converted to a 1×1 worksheet.

## Risks / Trade-offs

- [Risk] A literal zero-length string can be semantically distinct from a truly absent cell to an Excel formula such as `ISBLANK`. → The UI diff contract is display/value comparison; formula identity is still preserved, and whitespace/non-empty text remains distinct.
- [Risk] Excel COM can be unavailable in a headless/logon-limited test session. → Keep deterministic fallback and OOXML tests, and run native replay coverage when COM is available.
- [Risk] Changing the shared key affects row/column signatures as well as exact cell comparison. → This is intentional so blank padding cannot create structural false diffs; existing formula and non-empty-type regression tests remain required.
- [Risk] Calculating an omitted dimension requires a full XML scan. → Only do it when openpyxl reports a missing bound; normal workbooks retain their declared-dimension fast path.

## Migration Plan

No data migration is required. Ship the code change, run the focused save/diff regressions, then package the executable. Existing workbooks are read unchanged; only future comparisons and saves use the corrected behavior.

## Open Questions

None.
