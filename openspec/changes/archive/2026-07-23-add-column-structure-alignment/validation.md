## Validation Record

Date: 2026-07-23

Release: `2026-07-23.update55` / `new132-column-structure-alignment`

Implementation freeze before the release-constant edit: SHA-256 `24F274969E4A87A74F17D7CE316A7CC1C2B59106F849CC348B87861080975C50`. Final release source SHA-256: `40A3AC7A2BE850A64DBBB94F71E6CB179B8DED34A84493C5439ADC99DB1BA496`.

### Correctness and safety

- Logical immutable column slots are shared by 2-way/3-way comparison, rendering, hit testing, only-diff, minimap, actions, undo, and save.
- Common Base-relative insertions merge only when an exact ordered full-column proof exists; ambiguous, competing, delete-versus-modify, formula/cache, and low-confidence cases remain visible and conservative.
- Column actions record ordered insert/delete/copy operations, restore formulas/content/metadata/mappings/selection on undo, and fail closed on unsupported or unsafe replay.
- Formula translation covers proven local and target-sheet A1 references, ranges, whole columns, absolute markers, insert/delete contraction, and `#REF!`; external, structured, named, and unrelated-sheet references remain conservative.
- Native replay preserves full-column content/metadata, uses immutable uniquely named source staging, validates the OOXML package, reopens it in Excel, and atomically exposes the output only after all gates pass. No unsafe openpyxl fallback is allowed for column structure.
- Content-only cell/region edits advance edit versions while retaining proven immutable column geometry. Ordinary undo refreshes exact affected pairs; row/column/Sheet structure still invalidates and rebuilds the full model.

### Automated regression matrix

Final headless run on the implementation freeze:

- 52 command gates passed, 0 failed.
- 106 independently named tests/gates passed; two environment/optional observations were reported during the headless run (live SVN CLI absent and real Excel disabled there).
- Logical column actions: 31/31; logical geometry: 15/15.
- Native-save mock gates: 7/7; non-COM fast mutation parity: 2/2; sheet-cache reuse: 3/3.
- Row alignment, 3-way/Base insertion, tail append, formula-cache save/undo, only-diff blocks, minimap, hover/C-area, horizontal synchronization, row headers, progress, sheet-level operations, XLSM, SVN artifact detection, large sheets, stable source copies, mapping guards, and save failure/retry paths passed.
- `py_compile`, strict OpenSpec validation, and `git diff --check` passed.

Excel-dependent gates were then run serially in the desktop interactive session:

- Real Guide common inserted columns: passed native save and Excel reopen.
- 3-way independent tail append: passed, including native save.
- Real native XLSX/XLSM column replay: 8/8 passed.
- Fast in-memory versus real Excel native final-state parity: 3/3 passed.
- Formula-bearing manual row insertion and COM blank-cell save: passed.
- Real Skill formula/column convergence: passed; first insert was 432.5 ms, formula capture 163.5 ms, refresh 125.3 ms, with zero residual visual/structural differences.

The same real-native suite returned `HRESULT 80070520` from a sandbox logon token, failed closed, removed no target, and left no Excel process. The unchanged suite passed 8/8 when rerun with the desktop interactive token; the failure is recorded as environment reliability evidence rather than suppressed or counted as a product pass.

### Real workbook acceptance

All mutations were made on isolated copies under `C:\Users\dd\AppData\Local\Temp\sow_ux_5_3_20260723_001`. The originals under `C:\GM15\design\sheets\develop` were read-only inputs and retained identical SHA-256 hashes before and after acceptance:

- `Guide.xlsx`: `F973E286C8261F04AE85AFC8B4D054D23E9A7026F2385B8BD760A2266B273A36`
- `Skill.xlsx`: `A5727975EC55C285BE94BDE4F5A0E1F0805FCBAF8489614BB68071ABF98D63CD`
- `Dungeon.xlsx`: `5B5016A90D277332052B270DDA46BD031308D70090281E778E94C13A88E68AE6`
- `WorldMonster.xlsx`: `58148A08C6A4BCD7983E2DB869A169E5809001A279E41AB3258E93DE953D1BCF`

Guide replayed independent multi-row edits/inserts/deletes, two inserted and two deleted columns, and four same-cell conflicts. Its output contract verified every authorized marker and conflict choice, all 885 inserted-column rows, expected shapes, and zero unauthorized payload differences. Skill and Dungeon each replayed two inserted and two deleted columns through Excel; final outputs were cell-for-cell equal to Theirs. Skill retained 7,961 formulas with a stable 0 -> 0 same-formula cache difference; the cache-restoration branch is separately covered by deterministic same-formula/different-formula/idempotence tests. Dungeon retained real Excel external-link/formula rewrites as visible semantic differences before adoption. WorldMonster preserved the uncommitted Mine edit at `B3000` and adopted only the committed Theirs edit at `C3000`.

All four outputs passed package validation and independent Excel reopen. Canonical isolated SVN conflict artifacts correctly identified Base r100, local-uncommitted Mine, and Theirs r101; `svn.exe` was absent, so no live repository was mutated. Full evidence is in `C:\Users\dd\AppData\Local\Temp\sow_ux_5_3_20260723_001\acceptance_report.json`.

### Performance acceptance

Final isolated UX measurements after optimization:

| Workbook | Full stable projection | Precise only-diff | Actions | Native save |
| --- | ---: | ---: | ---: | ---: |
| Guide | 4.091 s | 58 ms | apply/undo/redo 465/412/487 ms | 6.024 s |
| Skill | 20.099 s | 1.937 s | insert/undo/redo/delete 410/269/393/480 ms | 8.688 s |
| Dungeon | 14.210 s | 3.996 s | insert/undo/redo/delete 783/533/776/1,479 ms | 8.873 s |
| WorldMonster | 18.224 s | 3.048 s | apply/undo/redo 207/205/204 ms | 2.367 s |

WorldMonster precise only-diff improved from fresh-process P95 17.04 s to 3.05 s; apply/undo/redo improved from 2.30/4.34/2.25 s to about 0.205 s. Skill and Dungeon retained every formula/cache semantic gate while native saves fell below the calibrated 11-second formula-dense target. Small native replay remains capped at 7.5 seconds; formula-dense native replay uses P95 <= 11 seconds and a 12-second single-run hard limit. Immutable staging, full-column metadata copy, OOXML validation, and Excel reopen were not weakened to obtain these results.

### Script audit

- `_open_language_target_sheet.py` is an obsolete one-off interactive diagnostic hard-coded to `Language.xlsx`, an old r29098 side file, and one target Sheet. It can resolve SVN Base and open the GUI, but is not a reusable regression and should not be run casually against the working-copy path.
- `_proto_virtual_text_render.py` is an isolated 200,000-row Tk text-window virtualization prototype. It is not imported or wired into the product; it documents the fixed-window/proxy-scrollbar approach and the remaining selection/hover/navigation integration work.

Both pre-existing untracked scripts were preserved unchanged.

### Packaging

- PyInstaller build completed with dependency installation skipped because the locked environment was unchanged.
- Stable and release executables are byte-identical, SHA-256 `473B8860AC4FFE9AE620CA5502D653E16F603A409FED821B1FB9B37822EC4754`.
- Release ZIP SHA-256: `3C9A55E4935D9AA13C02B6D514A5EBCEC43E05CE29B6EC1D505E3CD0520E1DAC`.
- Archived executable: `dist/archive/sow_merge_tool_new132-column-structure-alignment_20260723_052437.exe`.
- Historical release ZIP: `release/sow_merge_tool_release_2026-07-23.update55_new132-column-structure-alignment.zip`.
- Publish synchronization, Git staging, commit, and push were intentionally not performed.

Acceptance result: the change is ready to archive.
