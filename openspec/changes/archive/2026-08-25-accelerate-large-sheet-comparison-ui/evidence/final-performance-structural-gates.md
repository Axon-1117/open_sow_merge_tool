# Final Performance, Cache, and Structural-Invalidation Gates

Captured 2026-08-17 (Asia/Shanghai).  Production remained frozen at
`02C7392F5EB888490C4683D6B453374AE8F7B5E7CD84880B632E81C561022089`.
No production comparison, operation, or save code was changed for this audit.
Every generated workbook was created under a Python temporary directory; the
production safe-save helper's one OS-temp staged output was package-validated,
read-only reopened, and removed by the harness.  The source corpus under
`C:\GM15\design\sheets\develop` was only copied/read.

## Cold/warm phases and memory (task 1.3)

Machine-readable evidence:

- `phase-final-20k.json`: three fresh children, synthetic 20,000-difference
  workbook.
- `phase-final-worldmonster-3way.json`: one fresh child, disposable three-way
  `WorldMonsterSurvivor@design` real fixture.
- `phase-final-save-synthetic.json`: three fresh children, actual accepted
  cell action + production safe-save + package validation + read-only reopen.

| Phase | Result | Acceptance interpretation |
| --- | ---: | --- |
| 20k selected detail (six fresh samples recorded in the phase workers) | worst 6.389 s | below 15 s release gate |
| 20k cached revisit P95 | 4.847 ms | below 100 ms |
| 20k viewport publish P95 | 0.888 ms | below 33 ms |
| 20k scroll viewport publish P95 | 0.778 ms | below 33 ms |
| 20k heartbeat max-gap P95 | 92.532 ms | below 200 ms |
| WorldMonster 3-way selected-ready | 8.037 s | below 15 s release gate |
| WorldMonster 3-way peak RSS | 218,308,608 B (208.2 MiB) | below 400 MiB Skill-path ceiling; corpus evidence remains authoritative for all Sheets |
| WorldMonster 3-way cached revisit | 70.339 ms | below 100 ms |
| WorldMonster 3-way scroll viewport publish | 10.847 ms | below 33 ms |
| Safe-save P95 (synthetic 3,001-row operation) | 82.610 ms | separate production save phase; package/reopen passed |

The single WorldMonster cached-revisit run recorded a first direct viewport
publication of 50.712 ms and a 153.443 ms heartbeat gap.  This is retained as
calibration data; it is not substituted for the required 20,000-difference
synthetic scroll benchmark, which passed its 33 ms/200 ms P95 limits above.

The real GUI corpus gate remains separately recorded in
`large_sheet_corpus_gui_final_full3.json`; it is the authoritative all-files,
all-Sheets 15-second runtime result, rather than a replacement for these
phase measurements.

## Virtualized interaction and no-work guards (tasks 4.4, 7.2)

`_gui_self_test_large_virtual_viewport.py` passed on a disposable 20,000-row
all-difference workbook.  It asserts all of the following while normal and
read-only worksheet access, snapshot ingestion, formula normalization,
alignment, and logical-row comparison functions are replaced with immediate
failures:

- rapid thumb requests coalesce to the newest logical target before the 16 ms
  callback;
- page, wheel, thumb, minimap, and precomputed difference-block navigation
  target virtual rows directly;
- no pane exceeds the configured 20-row viewport limit; and
- the final bounded publication is below 33 ms.

`_gui_self_test_sheet_cache_reuse.py` passed with `Worksheet.cell`,
`Worksheet.iter_rows`, and `ReadOnlyWorksheet.iter_rows` forbidden during the
unchanged selected-Sheet snapshot revisit.  It also confirms unopened sibling
worksheet XML is not parsed and mutable workbooks remain deferred until an
operation demands them.

`_performance_test_column_structure_guards.py` passed.  Its cached replay
paths reported zero worksheet cell/row reads and zero signature/alignment/cache
rebuild calls; it also verifies the one-owner lazy editable backend promotes
the edit and cached-value pairs exactly once.

## Real 1,000-row operation timings (tasks 1.3, 7.2)

Three serial disposable runs of `_gui_self_test_large_overlay_batch.py` each
applied a virtualized 1,000-row region, then undid and redid it through the
real `SowMergeApp` path after explicit deferred-backend promotion:

| Sample | Apply | Undo | Redo |
| ---: | ---: | ---: | ---: |
| 1 | 100.7 ms | 63.7 ms | 71.2 ms |
| 2 | 91.7 ms | 112.7 ms | 58.3 ms |
| 3 | 159.5 ms | 68.8 ms | 93.3 ms |
| nearest-rank P95 | 159.5 ms | 112.7 ms | 93.3 ms |

The test enforces 2,000 ms for apply and 500 ms for undo, confirms a single
refresh/transaction, verifies an offscreen target, preserves the 20-row
viewport bound, and checks the exact value after apply, undo, and redo.

## Structural isolation and stale targets (task 5.5)

`_gui_self_test_structural_sheet_cache_isolation.py` passed with disposable
two-Sheet workbooks.  A real selected `S1` column adoption rebuilt its own
snapshot; the existing `S2` immutable snapshot object, view object, exact
entry, and ready state were unchanged.  The test then advanced `S1` to a
stale logical mapping and proved a copy click returned false without adding a
manual cell operation, overlay delta, undo record, or extra topology advance.

## Harness provenance

- `_performance_large_sheet_phases.py`
  `C62DEF91B8E95F213F427D5ABEA1E9731808F812D35E7DA8AA5B846773E56D08`
- `_gui_self_test_large_virtual_viewport.py`
  `AAEC9DF5CD9748F82A838F8C8FE5E5E00C3A298F5A3BB51B82D748B086141131`
- `_gui_self_test_large_overlay_batch.py`
  `B88B8073FFB9E4DE92F2EA16ADE69D9B39E96556F8CB8383D8774ED774F6825F`
- `_gui_self_test_sheet_cache_reuse.py`
  `E00D884E0DC7CD30882F872E2A8E923241DD987E93871E201FF7BE4CE402EA81`
- `_gui_self_test_structural_sheet_cache_isolation.py`
  `4574BCB7CCCF8E484C3F3199CEA8762A7BF3501D5EB375A612E0C8E6BF30B829`
- `_performance_test_column_structure_guards.py`
  `8FC672E3B7217731A9675A8D18397B4F1A4BA61237130EBCA97C7072C5559DEC`
