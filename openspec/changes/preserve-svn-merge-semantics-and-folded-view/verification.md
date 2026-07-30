# Verification evidence

Date: 2026-07-28

## Independent functional coverage

- SVN launch roles, CLI/auto-detected raw-Base preservation, and two-way sidecar normalization: 5/5 passed.
- Complete OOXML package/equivalence matrix, including unavailable-role evidence: 6/6 passed.
- Whole-workbook convergence and semantic pre-merge: 3/3 passed.
- SVN Author labels and diagnostic evidence: 4/4 passed.
- Adaptive three-way GUI, including the Mine/Theirs folding case: 4/4 passed.
- Python compilation passed.

## Complete suite closure

- Audited inventory: 66 executable regression scripts (39 smoke + 27 GUI).
- Final result: 62 PASS / 4 SKIP / 0 FAIL.
- The original complete inventory run was followed by an independent replay of
  all 18 scripts adjacent to the final logical-column and only-diff changes;
  all 18 passed on the final source state.
- `python -m py_compile sow_merge_tool.py` and `python -m compileall -q .`
  passed.
- `git diff --check` passed.
- `openspec validate preserve-svn-merge-semantics-and-folded-view` passed.
- Real `Guide.xlsx` common-column acceptance passed the complete sequence:
  apply two columns, rebuild mapping without changing `row_pairs`, suppress
  accepted row-offset noise, restore a later real edit as a visible
  difference, undo, reapply, native Excel save, package validation, and reopen.
  The native save was run outside the sandbox on per-run temporary copies.
- Four intentional skips remain:
  - `_gui_self_test_logical_column_actions.py`: safety skip because it directly
    operates on fixed real workbooks instead of per-run copies.
  - `_gui_self_test_real_skill_column_formula_convergence.py`: safety skip
    because it directly opens a fixed UX Skill workbook.
  - `_gui_self_test_real_item_blank_row_delete.py`: required Item Base fixture
    is unavailable.
  - `_gui_self_test_real_link_theirs_row_delete.py`: required Link Base/Theirs
    fixtures are unavailable.

## Existing targeted regressions

- `_smoke_test_svn_conflict_detection.py`: passed; live SVN CLI case skipped because `svn.exe` is not installed.
- `_smoke_test_merge_three_way_row_align.py`: passed.
- `_smoke_test_3way_alignment.py`: passed.
- `_smoke_test_xlsm_support.py`: passed.
- `_gui_self_test_sheet_diff_state_3way.py`: passed.
- `_gui_self_test_loading_readonly_gate.py`: 7/7 passed.

## Real-workbook package baselines

All runs were read-only.

- Four distinct `WorldMonster.xlsx` workbooks from local branches:
  - 4 identities / 6 pairs.
  - Complete matrix: 210.8 ms.
  - Every package was readable; one Mine/Theirs pair converged by complete member-byte equality.
  - Full startup analysis: 220.1 ms; converged to the common Mine/Theirs candidate with zero unresolved conflicts.
- `develop/WorldMonster.xlsx`:
  - 2,381,699-byte ZIP package.
  - 21 members / 21,865,418 uncompressed payload bytes.
  - One-pass fingerprint: 62.0 ms.
- `develop/Gunships护山神兽.xlsx`:
  - 27 members / 1,319,663 uncompressed payload bytes.
  - One-pass fingerprint: 13.4 ms.
- Real `Platform.xlsm` pair from `version_certificate` and `online`:
  - Complete member-byte equality.
  - Pair comparison: 27.9 ms.
- Real VBA-bearing `Battle/StageCfgs.xlsm`:
  - Full startup analysis: 31.5 ms.
  - Candidate package was valid and retained `xl/vbaProject.bin`.
- Real WC metadata for `develop/Gunships护山神兽.xlsx`:
  - Mine resolved locally from `.svn/wc.db` as revision 36737, Author `jun.yin`.

These measurements cover the new role/equivalence stage only. Existing workbook parsing and exact sheet-difference computation retain their separate performance telemetry.

## Final real-data replay before update63 publication

Date: 2026-07-29

- Replayed six cross-branch histories from real SVN revisions:
  `WorldMonster`, `Gunships护山神兽`, `Guide`, `Item`, `Building`, and `Skill`.
  WorldMonster converged by complete Mine/Theirs package equality; Guide,
  Item, and Building retained their semantic-premerge outcomes and candidates.
- Effective-column-width preflight safely declined the legacy physical-column
  writer for the structural Gunships and Skill histories. Gunships completed
  startup in 1.04 seconds and Skill in 8.50 seconds, both with a valid
  byte-identical Mine candidate for manual three-way review. The earlier Skill
  baseline was 202.7 seconds.
- Replayed the fixed real `release/Building.xlsx.merge-left.r37073` and
  `.merge-right.r37074` sidecars: semantic premerge applied 11 changes with
  zero unresolved conflicts and produced a valid candidate.
- Replayed an update conflict from real release revisions 37137/37138:
  the non-overlap case retained a local-only worksheet and produced zero
  unresolved conflicts; an edit overlapping the real `Building@design!H6`
  change produced exactly one unresolved conflict.
- Replayed a large WorldMonster block operation on isolated copies:
  1,201 precise difference rows in two blocks; applying the first block wrote
  8,955 cells, preserved the adjacent block, and preserved an independent
  local Mine edit.
- WorldMonster only-diff progress/cancel acceptance passed with a stable
  checkbox position, responsive heartbeat, locked sheet switching, rejected
  stale publication, and restored state after cancellation.
- Native Excel/COM save of a temporary copy of real
  `version_certificate/Battle/StageCfgs.xlsm` produced a valid reopenable XLSM.
  Excel rewrote the binary VBA stream, but the VBA project/component semantic
  model remained identical.
- SHA-1 comparison against `.svn/pristine` confirmed that develop/release
  copies of WorldMonster, Gunships, Guide, Item, Building, and Skill were
  unchanged after all replay runs.

## Source-delta refinement verification before update64 publication

Date: 2026-07-30

- Replayed the actual source snapshots
  `Building.develop.at37073.xlsx` and
  `Building.develop.at37074.xlsx` against the current release Target Working.
  The final result was exactly `incoming=1`, `applied=0`,
  `already_present=1`, `target_retained=14`, `unresolved=0`, and
  compatibility `merged_count=0`.
- The Building candidate SHA-256 was byte-for-byte identical to Target
  Working (`20c141...ae4f`), and all three replay inputs retained their
  original SHA-256, size, and modification time.
- The release-r37073 control case detected the unrelated `xl/styles.xml`
  package difference and failed closed with one unresolved incoming item;
  its candidate also remained byte-identical to Target Working.
- Source-delta regressions cover exact SVN conflict metadata with Chinese
  repository paths, exact source-path author lookup, applied/already/third
  cell states, same-shape row replacement, same-width column replacement,
  target-only row/column/Sheet retention, blank writes and clears, shared
  strings, formula caches, document properties, styles, and VBA/member
  changes.
- Non-interactive inventory: 40/40 scripts passed. The live SVN CLI subcase
  was skipped because `svn.exe` is not installed; its offline conflict and
  conflict-data fixtures passed.
- GUI inventory: 25/27 scripts passed. Two real Item/Link scripts skipped
  before entering product code because their historical SVN sidecar fixtures
  were unavailable.
- The real WorldMonster block replay passed with 1,201 difference rows,
  two blocks, navigation without rescan, 8,955 adopted cells, and preserved
  adjacent/local edits.
- The real Skill column convergence replay passed after explicitly verifying
  that pre-READY writes are rejected without state or file changes. After
  READY, apply/undo/reapply/delete converged to zero visual, structural, and
  unresolved differences; the first column insertion took 451 ms.
- `python -m compileall -q .`, `git diff --check`, and strict OpenSpec
  validation passed on the frozen product source.

## Structural-conflict and compact-layout verification before update65

Date: 2026-07-30

- Replayed the actual develop conflict inputs for `Gunships护山神兽.xlsx`.
  Source Before/After each contain 26 effective columns in
  `GunshipsModify@design`; Target Working contains 24 and is missing the two
  independent logical columns L14 (`part_level_icon`) and L20 (`map_model`).
  READY automatically selects L14 and enables keep Target Working, adopt
  Source Before, and adopt Source After. The latter two plan one-column native
  `insert_copy` operations at anchor 14; `GunshipsConfig@column` remains a
  normal E10 cell difference with all structural-column actions disabled.
- Workbook-level `<workbook>` markers are no longer returned as cell
  locations. The unsupported `xl/worksheets/sheet2.xml` fallback resolves
  through Source Before's OOXML relationship map to
  `GunshipsModify@design`, so the outcome action opens that real Sheet in full
  three-way review without inventing a row/column address.
- At 1450x860 the permanent top diagnostics and duplicate upper Sheet tab row
  are absent, the lower strip is the only visible Sheet navigator, the hidden
  C2 page no longer reserves height, and the main grid measured 353 px
  (previous field observation: about 117 px). Compact pane labels retained
  role, workbook/revision, and Author without clipping; hover detail retained
  the full input path and author lookup reason.
- Focused GUI acceptance passed 3/3; conflict navigation passed 7/7; logical
  column actions passed 35/35; logical column geometry, C-area scrolling,
  C-area cell alignment, adaptive workspace (5/5), and the main smoke suite
  passed. Cross-branch source-delta passed 12/12 in an isolated rerun.
- `python -m py_compile`, `git diff --check`, and strict OpenSpec validation
  passed.
- Real input SHA-256 remained unchanged:
  Target Working `B2D4ECC4...5CC500AC`, Source Before
  `A400769A...C4E80293`, Source After `A165CBD2...492030DE`.

## Action-first structural workflow and stable-author verification for update66

Date: 2026-07-30

- Replayed the actual develop conflict inputs for `Gunships护山神兽.xlsx`.
  The fixed first-row action group requires 920 px after moving difference
  navigation onto the existing structural-action row. At both 1450x860 and
  1024x760, every first-row action, difference-navigation child, and all three
  structural-column buttons were mapped at their requested width with no
  overlap or overflow. Injecting a twelve-times-expanded structural summary
  did not move any action control.
- The real `GunshipsModify@design` workflow automatically selected L14, kept
  all three column actions enabled after applying it, then selected L20 and
  visibly reported it as pending. Only after applying L20 did the UI report
  `列结构处理完成` and disable the three buttons.
- The release working-copy node had already advanced beyond r37348. Exact
  repository URL-plus-revision lookups nevertheless resolved Source Before
  r37347 to `cheng.zhu2` and Source After r37348 to `rongheng.xue`.
  TortoiseSVN-only lookup ran in an isolated child process with an eight-second
  timeout and successful results were cached in memory.
- The focused real Gunships suite passed 3/3, SVN author diagnostics passed
  8/8, logical-column actions passed 35/35, focused merge acceptance passed
  3/3, logical-column geometry and adaptive three-way workspace passed, the
  cross-branch source-delta suite passed 12/12, conflict navigation passed 7/7,
  and the main smoke test passed.
- `python -m py_compile` and `git diff --check` passed. Real input SHA-256
  remained unchanged: Target Working
  `B2D4ECC4ABDD34A48D734453BF1B9491A45FDC0B9AFE27CF7667EE4E5CC500AC`,
  Source Before
  `A400769ABEFABFB1D93B79CA44A955501F247B71615ED0BCE53D0079C4E80293`,
  and Source After
  `A165CBD2F890F64BCDCCC5DCEEA98CCC5FF8ABCED94735B51971E9B9492030DE`.

## Centered actions, Global Mode, and unified geometry verification for update67

Date: 2026-07-30

- The complete Previous/Block status/Next group is horizontally centered in
  its own 33 px navigation row at both 1450x860 and 1024x760. The full
  structural decision text sits immediately before the three column buttons;
  only the selected logical range (for example `L20`) is red and bold.
- Both adoption menus expose Row, Region, and Global modes. Global Mode uses
  the complete exact current-Sheet diff maps, rejects stale, structural, or
  ambiguous models before the first write, confirms the source and cell
  count, and commits all safe cells as one undoable transaction. In
  three-way conflict mode the left action uses Source Before and the right
  action uses Source After; undo and injected post-write failures restore
  values, manual-operation maps, dirty state, conflict maps, and the
  user-touched flag.
- The per-Sheet logical width tuple is fixed at 18 characters and reused by
  all main panes, headers, cached rows, prescan, projection rebuild, and
  restore paths. C-area headers now use Excel labels `A` through `Z`, `AA`,
  and so on while retaining structural markers.
- The dedicated section-10 GUI suite passed 8/8. Its real develop/release
  workbook replay was read-only, and production SHA-256 remained
  `7887C347C848687BF236467A1C99CC91A2F052F485FC7B97412F47FC7FD76586`
  before and after every final GUI suite.
- The many-small-block benchmark applied 350 interleaved conflict blocks
  across 700 rows in 1.067 seconds, performed one consolidated refresh,
  created one undo entry, and restored the exact pre-action state in one
  undo.
- Final affected regressions passed: Gunships 3/3, Region Mode 16/16,
  logical-column actions 35/35, logical-column geometry 15/15, adaptive
  three-way workspace 5/5, focused merge acceptance 3/3, conflict navigation
  8/8, cross-branch source delta 12/12, SVN Author diagnostics 8/8, OOXML
  equivalence 6/6, automatic merge semantics 4/4, SVN merge roles 5/5,
  diff-block presentation, C-area hover/alignment, click-x stability,
  bottom-bar alignment, and the main smoke suite.
- `python -m py_compile`, `git diff --check`, and strict OpenSpec validation
  passed on the frozen production source.

## CJK pixel geometry and compact action-layout verification for update68

Date: 2026-07-30

- The source workbook uses one fixed width for `GunshipsModify@design` column
  A. The former renderer instead padded with Python `len()` under Consolas 11;
  Tk measured Latin/space at 8 px and the fallback Chinese glyph at 15 px.
  Consequently the first separator was at x=146 for `id@id`, `uint32`, and
  `1`, x=167 for `小等级`, and x=174 for `此行不填`.
- The renderer now uses NFC-normalized East Asian display width, a common
  `新宋体 11` editor font, and zero-pixel index placeholders after two-slot
  glyphs. The real A1:A5 values and mixed Latin/accent/combining cases each
  occupy exactly 18 Tk indices and 144 px; every measured boundary across all
  three panes and the header differs by no more than 1 px.
- Cursor, selection, conflict, and difference tags retain the editor font and
  therefore cannot move geometry. The Section 11 GUI suite passed 4/4.
- Difference navigation and the structural status plus three column buttons
  share one responsive row. At 1450x860 navigation remains absolutely
  centered; at 1024x760 it shifts left with a 10 px safety gap while no
  control overlaps or clips.
- Root utilities appear at the far left in the required order. All visible
  structural-column guidance uses Excel names such as `A`, `Z`, `AA`, and
  `T`; internal `L<n>` identifiers remain diagnostic-only.
- Final regressions passed: Region Mode 16/16, Section 10 8/8, focused
  Gunships 3/3, focused merge 3/3, logical-column actions 35/35, logical
  geometry 15/15, adaptive workspace 5/5, diff-block presentation, C-area
  hover/alignment, click-x stability, conflict navigation 8/8, bottom-bar
  alignment, and the main smoke suite.
- `python -m py_compile`, `git diff --check`, and strict OpenSpec validation
  passed. Frozen production SHA-256:
  `6DE63D384CEF35A677AA5E05A877111629C85A5F8BC6CEA5F719467D5CF1C671`.
- The update68 release was built and published on 2026-07-30. Local release
  files and the configured `C:\GM15\design\design\常用软件\excel_merge_tool`
  destination matched for all nine published files. The executable SHA-256 is
  `D34EFB30EC7D3D2D0D535857C88825593585389EB3DA05835910456F4CC73245`;
  both compatible ZIP names use SHA-256
  `0E44053A1CE9689BD070FBB2F3D7C92AF6179891F2AAE7B94A20D3C59EF0D4FD`.
