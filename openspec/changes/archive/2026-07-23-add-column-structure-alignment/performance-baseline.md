## Column-Structure Performance Baseline

Captured on 2026-07-22 with three fresh-process repetitions per case and phase. The reproducible driver is `benchmarks/column_structure_baseline.py`; raw evidence is written outside the repository to `C:\tmp\column_structure_baseline_v2\baseline_results.json`.

### Environment and reproduction

- Windows 11 `10.0.26200`, Intel Core i7-14700K (20 physical / 28 logical cores), 63.77 GiB RAM.
- CPython 3.14.2 64-bit at `C:\Python314\python.exe`, openpyxl 3.1.5, psutil 7.2.2.
- Git HEAD `8088e5607853ffe6445260264b799c92a512be85`; measurements intentionally include the dirty task implementation and record its SHA-256 in the JSON report. The final report matches the current benchmark script SHA-256 `f307034f...8157` and implementation SHA-256 `44cfd451...17ff`.
- Run: `python benchmarks/column_structure_baseline.py --repeats 3 --real-root "C:\GM15\design\sheets\develop" --output-dir "C:\tmp\column_structure_baseline_v2" --scale-widths "68,128,255,256,257,512,513"`.
- Each case/phase/repetition runs in a fresh subprocess with `PYTHONHASHSEED=0`; P95 uses nearest rank, so with three samples it equals the maximum.

The real workbooks are opened read-only for cache, signature, and mapping phases. Structural variants for those phases are tuple-spliced in memory, preserving retained formula text without writing a workbook. The action/save phase uses disposable copies under `C:\tmp` and measures openpyxl throughput only; it is not a formula-cache, formula-reference, OOXML-fidelity, or mapping-correctness oracle.

`cold first-ready proxy` begins after the fresh worker has started Python, the benchmark script, openpyxl, and psutil. It includes the first import of `sow_merge_tool`, read-only workbook open, sequential row-cache capture, in-memory two-column insertion plus one-column deletion, two signature snapshots, and current task-1.3 mapping. GUI construction/rendering is not yet implemented and is excluded explicitly.

### Reproducible results

All values below are median / P95 milliseconds. Mapping-memory and cold-memory columns are fresh-process peak RSS deltas.

| Case / sheet | Shape | Signature total | 2-way mapping | Cold first-ready proxy | Mapping RSS | Cold RSS | Column action | Save |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| Synthetic | 4,000x96 -> 4,000x97 | 1,670.17 / 1,771.78 | 2.972 / 3.134 | 1,732.24 / 1,739.71 | 13.66 / 14.31 MB | 35.06 / 35.23 MB | 524.07 / 529.15 | 1,185.39 / 1,193.89 |
| `Guide.xlsx` / `TGuideStep@design` | 884x33 -> 884x34 | 36.30 / 36.37 | 0.489 / 0.516 | 239.70 / 252.40 | 3.53 / 3.70 MB | 11.39 / 11.63 MB | 36.10 / 36.58 | 256.57 / 265.36 |
| `Skill.xlsx` / `SkillTimeline@design` | 3,926x33 -> 3,926x34 | 224.56 / 253.09 | 0.663 / 0.753 | 765.31 / 773.83 | 5.93 / 6.05 MB | 16.31 / 16.43 MB | 169.65 / 175.99 | 1,772.94 / 1,884.01 |
| `Dungeon.xlsx` / `Dungeon@design` | 2,426x68 -> 2,426x69 | 365.10 / 370.14 | 1.802 / 1.918 | 1,301.12 / 1,332.73 | 6.01 / 6.11 MB | 36.49 / 37.10 MB | 196.63 / 202.69 | 1,267.30 / 1,296.18 |

### Width safety curve

These measurements use prebuilt snapshots with no exact intrinsic-key pairs but one unique high-similarity candidate per column, exercising the O(columns squared) path of the real task-1.3 API. The configured automatic-alignment ceiling is 256 columns; 257 and wider sheets take the deterministic `column-limit-exceeded` fallback.

| Width | Mapping median / P95 | Peak RSS median / P95 | Result |
| ---: | ---: | ---: | --- |
| 68 | 14.252 / 14.320 ms | 0.676 / 0.676 MB | 68 high-similarity anchors |
| 128 | 43.723 / 44.773 ms | 1.137 / 1.180 MB | 128 high-similarity anchors |
| 255 | 168.790 / 168.795 ms | 2.316 / 2.410 MB | 255 high-similarity anchors |
| 256 | 168.799 / 169.497 ms | 2.258 / 2.297 MB | 256 high-similarity anchors |
| 257 | 0.632 / 0.636 ms | 0.191 / 0.191 MB | full unresolved fallback |
| 512 | 1.197 / 1.204 ms | 0.324 / 0.340 MB | full unresolved fallback |
| 513 | 1.195 / 1.201 ms | 0.309 / 0.309 MB | full unresolved fallback |

The independent test matrix counted exactly 65,536 similarity evaluations for 256 high-similarity/no-exact columns, confirming the optimized O(columns squared) path; it also measured the separate all-exact fast path at about 12.3 ms for 256 columns with zero similarity calls. Before the optimization, the old repeated-ranking implementation took about 7.6 seconds at 512 columns; that regression is now prevented by the 256-column guard and focused smoke tests.

### Historical physical-index failure baseline

The earlier isolated fixture inserted two populated columns near the first third and deleted one original column near the second third. Its logical result is two structural blocks spanning three changed columns, but the physical-index comparator marked every row as changed: Guide 2,822 cells / 884 rows, Skill 8,497 cells / 3,926 rows, and Dungeon 29,059 cells / 2,426 rows. The original diagnostic JSON remains at `C:\tmp\column_alignment_baseline\baseline_results.json`; it is retained as failure evidence, not as the reproducible task-1.3 performance report.

### Acceptance guards for later integration

- Correctness: recover two explicit inserted/deleted structural blocks spanning three structural columns; formula-cache loss, duplicate columns, and blank tails may add `unresolved` blocks but MUST NOT create false inserted/deleted identities.
- Populated-cache mapping P95: representative real workbooks <= 10 ms; 256 exact columns <= 25 ms; independent 256-column high-similarity/no-exact case <= 250 ms.
- Width guard: 257 or more columns MUST take deterministic fallback in <= 5 ms on this reference machine unless a later change supplies a separately measured safe algorithm.
- Cold first-ready proxy P95: Guide <= 0.5 s, Skill <= 1.5 s, Dungeon <= 2.5 s, synthetic <= 3.0 s.
- Signature P95: Guide <= 100 ms, Skill <= 500 ms, Dungeon <= 750 ms, synthetic <= 2.5 s.
- Additional signature/mapping RSS <= 32 MB; cold proxy RSS <= 64 MB.
- After mapping, selection, hover, scrolling, navigation, and block presentation MUST perform no worksheet cell reads, worksheet iteration, signature rebuild, or alignment rescan.
- Interactive selection/hover P95 <= 50 ms; scroll/navigation P95 <= 250 ms; column action P95 <= 750 ms.
- Fidelity-preserving save is workload-tiered after direct Excel replay measurement: cell-only/ZIP and non-native low-formula saves remain <= 6 s; small workbooks that require cross-workbook native full-column replay have P95 <= 7.5 s; formula-dense multi-sheet native replay has P95 <= 11 s with a 12 s single-run hard limit. Package validation, immutable source staging, full-column metadata copy, and the unique Excel reopen gate remain mandatory in every tier.

### Post-integration native-save calibration

The original 6-second save guard above was proposed before native column replay existed; its baseline table measures openpyxl throughput and explicitly is not a fidelity oracle. Final profiling on the real `Skill.xlsx` acceptance fixture (2.93 MB, 14 sheets, 470,827 materialized cells, 87,649 formulas) measured 10.0-10.3 seconds end to end while preserving full-column metadata and reopening the output in the same Excel instance. The real `Guide.xlsx` native path measured 5.978 seconds and 6.798 seconds on equivalent accepted runs, establishing the small-native jitter band. The dominant costs were Excel process/workbook lifecycle, full-column cross-workbook copy, and full OOXML validation; `SaveCopyAs` itself was only 0.37-0.40 seconds.

Safety experiments confirmed that neither double garbage collection nor explicit COM range/column release reduced total latency. Bounded `UsedRange` copy was no faster and would weaken metadata fidelity. Forcing PowerShell termination left an orphan Excel process and is prohibited. The tiered guard records the measured physical cost instead of weakening correctness or disguising cleanup time.

### Post-integration interaction optimization

Fresh-process profiling on `WorldMonster.xlsx` found that precise only-diff display repeatedly normalized the same formula identity and tokenized identity column projections. A normal cell adopt also invalidated and rebuilt column geometry even though it changed no row/column topology, while ordinary undo unnecessarily reran full row alignment. The final implementation uses bounded formula/canonicalization caches keyed by immutable mappings, bypasses tokenization for identity mappings, renders one side without self-comparison, retags proven column geometry after content-only edits, and refreshes ordinary undo through existing row pairs. Structural row/column/Sheet actions retain full invalidation and rebuild.

Measured `WorldMonster.xlsx` results improved from fresh-process P95 17.04 s / 2.30 s / 4.34 s / 2.25 s for only-diff / adopt / undo / redo to 3.05 s / 0.207 s / 0.205 s / 0.204 s in the final isolated UX replay. `Skill.xlsx` structural actions completed in 0.269-0.480 s and native save in 8.688 s. `Dungeon.xlsx` retained 1,055+ real formula/external-link semantic differences while precise only-diff fell to 3.996 s and native save to 8.873 s. All native saves therefore met their calibrated workload tier without weakening package validation, full-column copy, or Excel reopen gates.
