# Final independent regression matrix (2026-08-17)

Production was frozen at SHA-256
`02C7392F5EB888490C4683D6B453374AE8F7B5E7CD84880B632E81C561022089`
before this run and rehashed unchanged afterward.  The source corpus
`C:\GM15\design\sheets\develop` was read-only; all workbook variants and
save outputs were created in disposable test roots.

## Corpus and Oracle acceptance

- `large_sheet_corpus_exact_run6_final.json`: 193 supported source workbooks,
  386 2-way/3-way runs, and 1,592 Sheet results; all PASS and exact-same with
  zero differences/conflicts.  The direct parser/legacy Oracle maximum was
  9,548.993 ms.
- `large_sheet_corpus_gui_final_full3.json`: the same 193/386/1,592 complete
  fresh `SowMergeApp` GUI corpus results; all PASS, terminal `EXACT_SAME`,
  prepared comparison detail, no calculation surface, and physical targets.
  No selected Sheet exceeded the 15-second gate; maximum
  constructor-to-final-exact time was 9,141.179 ms.  Editable backends were
  correctly still `EDIT_DEFERRED` in the view-only timing window.
- `_smoke_test_snapshot_oracle_parity.py` passed all adversarial Oracle cases.
  `_smoke_test_real_snapshot_direct_oracle.py` passed real Skill and
  WorldMonster self comparisons plus disposable Skill cell, row, and column
  changes against the direct legacy manifest.
- `worldmonster-excel-fidelity-gate.json` independently passed deterministic
  real WorldMonster value, formula/cache, row (at row 18,435), and column
  variants.  Each records exact direct 2-way and frozen-legacy 3-way results,
  conflicts, only-difference membership, and stable physical targets.

## Operation, save, and compatibility acceptance

`terra-final-regression/summary.json` records a serial 36-script independent
matrix.  Every script ran in a new Python process with separate stdout/stderr,
a hard per-script timeout, and exact PID-tree cleanup; all 36 passed.  It
covers row/column alignment and replay, only-diff and virtual rendering,
formula/cache, cell/row/region/structural operations, undo/redo, atomic ZIP
and native failure/retry, XLSX fidelity, XLSM/VBA, SVN merge/conflict/author
diagnostics, exact Sheet readiness, and binary-identical fast paths.

The manual row-insert smoke test was updated only in its setup: it now requests
the intentionally deferred editable backend and waits for exact Sheet readiness
before exercising structural replay.  It retains prompt-scheduler restoration
in `finally`; production behavior was not changed.  It passed in the final
matrix together with row-delete redo and cross-workbook style/metadata replay.

Existing full final compatibility evidence is retained in
`oracle-and-phase-validation.md`: real Skill v2 2-way/frozen-3-way Oracle
parity, Array/DataTable/external formula handling, cache-only save, actual App
operations, 1,000-row region apply/undo/redo, atomic failure recovery,
representative XLSX/XLSM package preservation, and available real Excel COM
reopen all passed.

## Source integrity and static checks

- `terra-final-regression/live-source-rehash.json`: all 193 live supported
  source workbooks rehashed against the final corpus inventory; 0 missing and
  0 mismatches.
- `python -m py_compile` passed for production, both corpus harnesses, and
  updated smoke tests.
- `git diff --check` passed (only pre-existing LF/CRLF warnings).
- `openspec validate accelerate-large-sheet-comparison-ui --strict` passed.

The change remains unarchived.  Direct parser timing is reported separately
from the fresh-GUI user-visible timing and is not used to satisfy the 15-second
opening gate.
