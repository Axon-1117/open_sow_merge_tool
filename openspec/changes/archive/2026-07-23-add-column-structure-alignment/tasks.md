## 1. Column Identity and Alignment Model

- [x] 1.1 Add immutable logical column-slot/block records, physical-to-logical lookup maps, confidence/ambiguity state, and row/column model version keys.
- [x] 1.2 Build cache-backed multi-signal column signatures from sequential row data without worksheet random reads.
- [x] 1.3 Implement deterministic 2-way column sequence alignment with retained, inserted, deleted, and unresolved ranges plus a safe physical-index fallback.
- [x] 1.4 Implement Base-anchored mine/theirs column mapping and deterministic placement of independent side-only insertions.
- [x] 1.5 Add focused logic tests for middle/tail insertions and deletions, duplicate/blank columns, bounded ambiguity, formulas, and independent 3-way changes.

## 2. Comparison, Conflict, and Presentation Integration

- [x] 2.1 Make background and foreground cell/formula comparison consume logical column slots and stop shifted columns from producing cascading differences.
- [x] 2.2 Extend 3-way scanning with retained/modified/deleted/inserted column states and conservative delete-versus-modify or competing-insertion conflicts.
- [x] 2.3 Render missing-side placeholders, structural column block markers, and aligned column headers/widths in 2-way and 3-way views.
- [x] 2.4 Route selection, hit testing, hover/C-area, horizontal synchronization, only-diff block data, and minimap column markers through logical slots.
- [x] 2.5 Invalidate and rebuild column mappings after row/column structural changes before stale render, conflict, action, undo, or save paths can continue.
- [x] 2.6 Add GUI regressions for column-shift noise suppression, 2-way/3-way pane alignment, mixed row/column structure, horizontal navigation, and cache replay.

## 3. Column-Level Actions and Undo

- [x] 3.1 Add discoverable column-block selection and Mine/Base/Theirs adopt/retain actions with clear unresolved-mapping diagnostics.
- [x] 3.2 Extend manual operation records with ordered insert/delete/copy column operations including target slot, source side, source physical columns, and metadata scope.
- [x] 3.3 Apply consecutive column operations in batches, update slot mappings atomically, and preserve adjacent logical columns.
- [x] 3.4 Add one-step undo/rollback for column actions that restores workbook content, formulas, metadata, mappings, selection, and difference state.
- [x] 3.5 Add action tests for inserted/deleted blocks, partially resolved conflicts, mixed cell/row/column actions, failure rollback, and repeated undo/adopt cycles.

## 4. Fidelity-Preserving Save and Reopen

- [x] 4.1 Add Excel native replay for batched insert/delete/copy column operations with full-column content and metadata transfer.
- [x] 4.2 Remap explicit cell and formula-cache operations after column structure changes and retain existing formula translation/special-formula safety checks.
- [x] 4.3 Preserve widths, hidden state, styles, comments, hyperlinks, validation, merged cells, conditional formatting, external links, macros, and untouched sheets according to the selected source policy.
- [x] 4.4 Reject unsafe fallback, validate the temporary OOXML package, and atomically replace the merged target only after native replay and reopen checks succeed.
- [x] 4.5 Add XLSX/XLSM save-reopen tests covering formulas, shared/array/data-table formulas, advanced metadata, untouched-part comparison, native failure, and recoverable retry.

## 5. Performance, Real-File Acceptance, and Release

- [x] 5.1 Establish pre-change timing and memory baselines for signature/mapping, first interactive state, column action, and save on synthetic plus representative real workbooks.
- [x] 5.2 Add performance guards proving mapping uses sequential caches and that scrolling, selection, navigation, and block presentation perform no worksheet rescans.
- [x] 5.3 Replay isolated copies of `Guide.xlsx`, `Skill.xlsx`, and at least one wide/formula-heavy project workbook with multi-column insert/delete plus independent edits and conflicts.
- [x] 5.4 Run the full row-alignment, formula-cache, only-diff block, minimap, hover/C-area, sheet-level, XLSM, SVN, progress, and save regression matrix.
- [x] 5.5 Record validation evidence and update release/version metadata only after correctness, fidelity, UX, and performance acceptance is approved.
