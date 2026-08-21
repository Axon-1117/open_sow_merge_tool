# Large-Sheet Oracle and Phase Validation

Captured on 2026-08-16 (Asia/Shanghai), before enabling any new comparison
path.  All mutations used `tmp/test_tmp/sow_large_sheet_*`; the source root
`C:\GM15\design\sheets\develop` was only read/copied.

## Oracle contract

- `_large_sheet_legacy_oracle.py` runs the existing application in a child
  Python process, waits for exact READY state, and serializes logical columns,
  aligned rows, physical coordinates, typed cached/formula tokens, row/column
  structure, direct three-way conflicts, and exact only-difference membership.
- `_large_sheet_snapshot_oracle.py` normalizes a frozen manifest and compares a
  future immutable-snapshot manifest without Tk or worksheet access.  It
  reports every mismatch; it never silently treats a changed row/cell mapping
  as equivalent.
- `_large_sheet_oracle_fixtures.py` contains real read-only fixture definitions
  for Skill, WorldMonster, Dungeon, Language, and IdleBuilding composite-key
  data.  The IdleBuilding Sheet is resolved by declared name or, safely, the
  largest qualifying Sheet, avoiding locale rendering differences.

Fresh-process 2-way verification passed for `equal_count_insert_delete`; the
legacy output compared exactly with its normalized copy.  Fresh-process 3-way
formula/cache verification passed and emitted the required `conflicts` field.

## Adversarial coverage

The disposable fixture generator self-test passed for all required cases:

- duplicate/missing declared keys;
- composite keys;
- blank continuation groups;
- equal-count insertion/deletion;
- reorder;
- same-formula different cached values;
- a direct Mine/Base/Theirs conflict;
- inserted column structure;
- stale generation and cancellation publication-gate cases.

## Fresh-process phase baseline

Evidence JSON files: `phase-baseline-synthetic.json` and
`phase-baseline-skill.json`.  Values below are one cold sample each; they are
baselines, not acceptance claims or P95 calibrations.

| Fixture | Startup | Selected Sheet READY | Cached revisit | Scroll | RSS delta |
| --- | ---: | ---: | ---: | ---: | ---: |
| synthetic 3,001 rows | 701 ms | 929 ms | 64 ms | 122 ms | 27–29 MB |
| Skill / SkillLogicBuff@design | 3,515 ms | 9,895 ms | 55 ms | 134 ms | 806–808 MB |

The scroll measurement replaces `ws_a_val`/`ws_b_val` with assertions after
READY; it completed without any worksheet read.  It is intentionally above the
eventual 33 ms target and documents the legacy baseline.

`action_1000`, undo, redo, and save remain explicit optional phase names in
`_performance_large_sheet_phases.py`.  They require the operation-overlay
implementation's stable target/hook and have not been represented by a fake
operation benchmark.  They remain pending under task 1.3 / task 7.2.

## Reproduction

```powershell
python _large_sheet_oracle_fixtures.py
python _large_sheet_legacy_oracle.py --mine <disposable-mine.xlsx> --theirs <disposable-theirs.xlsx> --sheet Data --out <legacy.json>
python _large_sheet_snapshot_oracle.py --compare --legacy <legacy.json> --candidate <snapshot.json> --out <result.json>
python _performance_large_sheet_phases.py --runs 3 --real Skill --out <evidence.json>
```
# Final implementation evidence (2026-08-16)

## Latest calibration

- No-hint Skill selected READY samples are about 4.94--5.48s with RSS
  +58--59MB, versus the legacy recorded 9.9s / +806MB baseline. WorldMonster
  is about 12.89s / +148MB.
- Fresh synthetic 20k cold viewport publish remains variable at roughly
  90--145ms and heartbeat 155--276ms; it does **not** stably meet the 33ms /
  200ms acceptance gate. Warm publication is about 1ms. The specification
  threshold is deliberately unchanged.
- The noninteractive region harness temporarily suppresses only its formula
  cache prompt scheduler and restores it in `finally`; case12--16 passed.

- Default selected-Sheet snapshot gate is enabled only at `>=2000` rows.  Any
  unresolved key/schema/base alignment falls back to the legacy worker.
- Exact Oracle parity passed for all synthetic adversaries and disposable real
  Skill/WorldMonster self pairs, plus Skill cell, row-insert, and tail-column
  variants.  Immutable 3-way Skill readiness was also exercised.
- Fresh-process samples: Skill 2-way READY about 4.95s and RSS +69MB; World
  READY about 10.15s and RSS +148MB; Skill 3-way READY about 7.26s and RSS
  +92MB.  The original READY target is therefore explicitly **not met** for
  all real fixtures; RSS is materially reduced from the recorded +806MB
  legacy Skill baseline.
- Skill virtual scroll 3-run actual publish P95 was 2.092ms and interaction
  heartbeat max 68.2ms; the outer ~125ms scroll timing contains its fixed
  120ms pump.  Region cases passed individually with 30s caps.  Native save
  replay passed 7 cases with one optional real-Excel case skipped.

## Final fixed-hash fidelity acceptance (2026-08-17)

This section is a strict final-validation record, not a claim that a
ZIP/openpyxl check is equivalent to Microsoft Excel. Production was frozen at
`07900BE94653B0701BAEFBDF0482CF23646480DF755BB8AFE76EB78EA4E4BCF9`.
The coordinated GUI benchmark harness was frozen at
`0A787498B0166B8BD9E0884A7E71E716B844962E9D0A246282D797F05E3C908E`.
All mutations were made to disposable workbooks; `C:\GM15\design\sheets\develop`
remained read-only.

- Full fidelity v2: `_large_sheet_excel_fidelity_gate.py --real Skill --timeout 45`
  passed in 266.9 s. Its temporary report was
  `%TEMP%\fidelity_v2_skill_07900BE.json` (schema
  `large-sheet-excel-fidelity-gate-v2`, `status: ok`). It proves pair-free
  stable physical identity (including negative tamper cases) and direct
  2-way plus frozen true-3-way legacy Oracle parity for the real
  `SkillLogicBuff@design` value, formula/cache, row, and structural-column
  variants. The value variant produced the expected 3-way conflict; the
  other frozen variants had the expected zero conflicts. Only-difference
  membership and physical targets are included in the manifest comparison.
- Special-formula gate passed in direct 2-way and frozen 3-way: ArrayFormula
  copy was rejected, DataTable and external-formula typed tokens were
  preserved by the Oracle comparison, and cache-only save plus ZIP/openpyxl
  reopen succeeded. It is intentionally a synthetic special-formula fixture;
  a formula text containing an external reference is not evidence of an
  `externalLink` package part.
- Actual application operations passed on disposable workbooks after explicit
  editable-backend preload: cell and region apply/undo, column insert/undo,
  column delete/undo, and every injected mutating-stage atomicity failure.
  The 1,000-row virtualized region apply/undo/redo regression passed in
  107.4 ms after preload. Deferred first-click/modal/no-mutation behavior is
  separately covered by the operation worker; completed-operation timing does
  not include editable-backend loading.
- Atomic recovery gates passed: injected ZIP save failure and injected
  native/reopen failure left source and user-target SHA-256 values unchanged,
  retained the operation records, cleaned staged output, and succeeded on a
  later retry after removal of the injection.
- Native save/reopen matrix passed 8/8, including the optional real Microsoft
  Excel COM native column replay. It covers XLSX route fidelity, XLSM macro
  and untouched ZIP-part preservation, column/row/cell replay ordering,
  native failure non-replacement, and reopen-validation fail-closed behavior.
  Independent `XLSM` support, XLSX save/diff fidelity, and 3-way formula-cache
  undo/save regressions also passed.
- The full 193-workbook no-op inventory was rehashed after validation:
  `records=193`, `missing=0`, `mismatches=0`. Its original no-op evidence is
  ZIP validity plus openpyxl reopen. The real source corpus contains zero
  XLSM workbooks, so this is not and must not be represented as real-corpus
  XLSM or real-Excel-COM no-op coverage.

The full gate deliberately reports actual App operations, atomic failure/retry,
and native replay as `not_covered`; those are supplied by the independent
actual-App, fault-injection, native replay, and COM checks above rather than
being silently inferred from an overlay-unit result. No final gate result
claims a full real-Excel COM reopen of all 193 source workbooks.

### Task 10.x acceptance recommendation

- **10.1 — recommend check:** the real large Skill Sheet has deterministic
  value/formula-cache/row/column mutations, exact direct 2-way and frozen
  true-3-way legacy Oracle parity, conflicts, only-diff membership, and
  stable physical target checks.
- **10.2 — do not check yet:** cell, region, structural, save, and non-ready
  no-mutation paths are covered by disposable representative fixtures and the
  slow 1,000-row region case, but final validation found a real 2-way row
  redo defect: an inserted row is absent after apply → undo → redo. This must
  be fixed and independently rerun before accepting the row/redo requirement.
- **10.3 — do not mark as universally complete without an explicit scope
  decision:** all 193 supported source workbooks have no-op ZIP/openpyxl and
  post-run hash evidence; representative XLSX/XLSM, atomic recovery, package
  metadata, and an available real-Excel COM native replay pass. Missing are
  real-corpus XLSM (none exist) and real-Excel-COM no-op reopen of all 193.

## Superseding final-candidate rerun (2026-08-17)

The earlier fixed-hash record above is historical only. After the row-redo and
cross-workbook style fixes, the final candidate was frozen at
`C771AF4904284147E031304472FB45EAEED3BE1C89606CB0446CF7156F2004EA`.

- Special formula direct 2-way/frozen 3-way, actual App cell/region/column
  actions and mutation-stage atomicity, and 1,000-row region apply/undo/redo
  all passed (`86.7 ms` for the post-preload region apply).
- ZIP and native/reopen failure/retry, XLSM support, XLSX save/diff fidelity,
  and native column save/reopen passed. The native matrix passed 8/8,
  including its real Microsoft Excel COM column replay.
- Full `Skill` v2 gate passed again in 267.9 s; temporary report:
  `%TEMP%\fidelity_v2_skill_C771AF49.json`. It reports `status: ok`, schema
  v2, stable pair-free physical identity `ok`, and direct 2-way plus frozen
  true-3-way parity for all four real-Skill variants.
- Post-run read-only corpus rehash: `sources=193`, `missing=0`,
  `mismatches=0`; no fidelity-gate Python process remained.

`manual_merge_row_insert` remains **INCONCLUSIVE**, not PASS, if its legacy
COM execution exceeds 180 seconds. Its equivalent new cross-workbook
row-style/metadata test and real-Excel COM output test are positive evidence,
but they do not retroactively turn that older timed-out script into a PASS.

## Final frozen compatibility rerun (production 02C7392F, 2026-08-17)

This section supersedes the prior production hashes for compatibility
acceptance. Production SHA-256 was identical before and after the matrix:
`02C7392F5EB888490C4683D6B453374AE8F7B5E7CD84880B632E81C561022089`.
No production code was changed during this rerun. All workbook mutation and
save tests used disposable copies; the real source corpus remained read-only.

- `py_compile` passed for production plus the Oracle/fidelity/row/save
  harnesses. The special formula fixture passed exact direct 2-way and frozen
  true-3-way parity for ArrayFormula, DataTable, and external-formula typed
  tokens; ArrayFormula copy remained rejected and cache-only save passed.
- Full real-Skill fidelity v2 passed in 267.5 s. Temporary report:
  `%TEMP%\fidelity_v2_skill_02C7392F.json`. It reports schema v2,
  `status: ok`, pair-free stable physical identity `ok`, four real-Skill
  value/formula-cache/row/column variants, direct 2-way parity, and frozen
  true-3-way legacy parity.
- Actual `SowMergeApp` cell/region/column actions and injected mutation-stage
  atomicity passed. The virtualized 1,000-row region apply/undo/redo passed in
  92.0 ms after explicit deferred-edit preload.
- Two-way row insert apply/undo/redo/save passed. Cross-workbook rich row
  style replay passed in 2-way and 3-way, including formula, comment,
  hyperlink, row height/hidden/outline metadata, save/reopen, failure/retry,
  source hashes, and real Microsoft Excel COM reopen.
- Added `_smoke_test_row_delete_redo_fidelity.py` as a test-only gate. It
  passed real-App 2-way and 3-way row delete apply/undo/redo/save/reopen,
  adjacent formulas, comment/hyperlink/row metadata restoration, manual row
  journals, row-model topology advancement, three-way Base targeting, and a
  two-way injected save failure with unchanged source hashes/retained journal
  followed by successful same-batch retry.
- ZIP failure/retry, native/reopen failure/retry, XLSM support, XLSX
  save-and-diff fidelity, shared/formula-cache undo-save, and native column
  save/reopen passed. The native column matrix passed 8/8 with the optional
  real Microsoft Excel COM case executed (0 skipped).
- The legacy `_smoke_test_manual_merge_row_insert.py` produced no output and
  exceeded its 180-second hard cap, so it remains **INCONCLUSIVE**. Its setup
  waits for `_edit_loaded_event` without requesting the deferred editable
  backend, so it may block on the current readiness modal before reaching the
  COM save path. The tracked process was terminated. Passing replacement
  row-style/metadata/COM gates are independent evidence and do not convert
  this old script into PASS.
- Final read-only audit: `sources=193`, `missing=0`, `mismatches=0` against
  the stored inventory; `tracked_python_orphans=0`. The real corpus still has
  no XLSM source, and this rerun does not claim real-Excel-COM no-op reopen of
  all 193 workbooks.

No OpenSpec task checkbox was changed by this compatibility rerun.
