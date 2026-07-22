## Context

`SheetView` already distinguishes contiguous difference runs by logical pair index for navigation and region operations. In only-diff mode, however, unrelated worksheet regions are compacted into adjacent `tk.Text` lines. The view exposes no block count or boundary, and `_compute_diff_blocks()` currently operates on the rendered `display_rows`, so large-sheet blocks beyond the initial 800-row window are absent from navigation state.

The existing snapshot behavior intentionally keeps touched rows visible after adoption. The new presentation must preserve that stability, row/pair mappings, formula and structural merge semantics, synchronized scrolling, and the fast-open strategy for large workbooks.

## Goals / Non-Goals

**Goals:**

- Make every only-diff block visually distinguishable and expose exact current, total, and pending block status.
- Count and navigate blocks from the complete precise snapshot, including unrendered large-sheet rows.
- Keep block numbering stable while users resolve differences.
- Use one block model for presentation, navigation, and region-level adoption.
- Add no workbook reads to scrolling, selection, or block presentation.

**Non-Goals:**

- Changing row alignment, conflict detection, diff-cell colors, minimap scale, formula handling, or save semantics.
- Adding context/unchanged worksheet rows to only-diff mode.
- Inserting synthetic text rows into data panes.
- Changing CLI arguments, persisted settings, or workbook output formats.

## Decisions

### Build a stable block model from the full only-diff snapshot

Introduce an internal block record containing a stable ordinal, ordered pair indices, start/end pair indices, current pending status, and optional rendered line bounds. Build records by scanning `_full_display_rows` and starting a new block whenever the next pair index is not the previous pair index plus one.

The model is keyed by the only-diff snapshot/source version and row-model version. It is rebuilt when alignment or snapshot membership changes, but ordinary scrolling and selection only read cached maps such as `pair_idx -> block_idx`.

This is preferred over `_compute_diff_blocks()` on `display_rows` because rendered rows are capped for large sheets and therefore cannot provide an exact total.

### Keep snapshot identity separate from pending status

Block membership and numbering come from the stable only-diff snapshot, including touched rows retained after adoption. Pending status is derived separately: a block is pending when at least one member currently satisfies `_pair_has_visual_diff()`.

This prevents later blocks from being renumbered after an earlier block is resolved. It also allows the toolbar to report both stable total blocks and dynamic pending blocks.

### Add spacing and a gutter marker without synthetic rows

Apply a dedicated block-start text tag to the first rendered row of every block after the first. The tag adds identical vertical spacing to left/Base/right data widgets and their row-number widgets. Add a compact block marker in the left row-number area for the first row of each block.

No blank or label lines are inserted into widget text. Therefore `display_rows[line - 1]`, click hit testing, hover extraction, row operations, selection restoration, and minimap mappings retain their current one-screen-line-to-one-pair invariant.

The marker is textual as well as spatial so the distinction does not depend on color perception. Block tags are applied by a shared bulk helper in full render, cached render, append-more-rows, and row-only refresh paths.

### Add one toolbar status model

Add a fixed-width block indicator near the existing previous/next buttons. Its states are:

- `差异块 计算中...` while precise snapshot data is unavailable.
- `差异块 -/0` when no block exists.
- `差异块 N/T · 待处理 P` when the active block is known.
- A processed indication when the active stable block has no remaining visual differences.

An explicit selected pair determines the active block. Without explicit selection, the top visible pair determines it. Hover alone does not continuously switch the toolbar indicator, avoiding label flicker as the pointer moves across cells.

### Navigate by pair target, not rendered line target

Previous/next commands select a block from the complete model, then resolve its first pair to a screen line. If the pair is outside `display_rows`, a helper increases the render limit or appends the required cached only-diff rows before scrolling. The helper must not rescan worksheet data and must preserve synchronized horizontal views.

The same path serves toolbar buttons and existing keyboard shortcuts.

### Make region actions consume the stable block record

Region adoption resolves the selected pair to a stable block and iterates that block's pair indices. Rows with no current visual difference are skipped, allowing a partially resolved block to remain one visual and operational unit. Existing cell-copy, formula-cache, batch insert/delete, undo, progress, and save-recording helpers remain responsible for each applicable row.

This replaces separate boundary inference in presentation and operation paths, eliminating disagreement about which rows a displayed region contains.

### Cache and invalidate without workbook access

Block construction is O(k), where k is the number of rows retained by the only-diff snapshot. Pending-status refresh may scan block members through existing in-memory diff maps but must not call worksheet cell APIs. Cache invalidation is tied to existing only-diff source and row-model version changes.

## Risks / Trade-offs

- [Risk] Missing a block-start tag on one pane could desynchronize vertical scrolling. -> Mitigation: apply and clear spacing through one helper across all active data and row-number widgets, with 2-way and 3-way geometry tests.
- [Risk] Row-header hover arrows could overwrite the block marker. -> Mitigation: derive row-header text from pair and block metadata in one formatter so normal, hover, and restored states preserve the marker contract.
- [Risk] A stale block cache could target the wrong region after structural insertion. -> Mitigation: include row-model/snapshot versions in the cache key and rebuild pair-to-block mappings after structural edits before re-enabling actions.
- [Risk] Materializing a distant block could accidentally render every prior large-sheet row. -> Mitigation: reuse cached only-diff rows and extend only to the target position; do not trigger `rescan=True`.
- [Trade-off] Stable snapshot blocks can remain visible after they are fully resolved. -> This matches current snapshot behavior and gives users reliable progress; processed styling and pending count make the state explicit.
- [Trade-off] A compact gutter marker conveys less context than a synthetic `... skipped rows ...` line. -> It avoids widespread line-mapping risk; original row numbers and block navigation still provide location context.

## Migration Plan

The change is internal and requires no data migration. Ship behind the existing only-diff mode behavior, verify both modes, and roll back by removing the block presentation/model integration if GUI synchronization regressions occur. Workbook merge and save data remain unaffected.

## Open Questions

None required for implementation. Exact spacing and marker colors may be tuned during GUI replay, but the marker must remain readable without relying on color alone.
