# Design

## Classification

The column model remains conservative. A slot can stay `unresolved` even when the current values on both sides are equal. Global value replacement classifies unresolved slots by their exact current diff state instead of blocking on the presence of any unresolved slot:

1. Unresolved slot with no source/destination cell difference: harmless no-op; exclude it from global candidates.
2. Unresolved slot with a real value/formula difference: blocking ambiguity; do not write any cell in that Sheet.
3. Missing physical column or one-sided structural row/column: blocking structural difference.

The comparison must use the same formula-aware cell comparison and exact-diff maps used by the existing global preflight. Empty `None`/`""` normalization remains in force. The ambiguous slot is not promoted to a trusted mapping merely because it is currently equal.

## Blocker dialog

When blocking ambiguity exists, show a modal dialog owned by the Sheet view. It must show the Sheet name, direction, logical Excel column labels, physical source/target columns where available, cause text, and a small count/sample of affected difference cells. The safe default is to return to the view without writing.

The dialog offers:

- Return/cancel.
- Apply the entire left Sheet to the right side (for `A2B`).
- Apply the entire right Sheet to the left side (for `B2A`).

Whole-Sheet choices require a second confirmation stating that the target Sheet will be replaced, including target-only cells/rows/columns, formulas, styles, merges, comments, hyperlinks, validation and Sheet layout. The action is one undoable Sheet operation and must reuse the existing Sheet-copy operation bookkeeping.

## Whole-Sheet operation

Use the existing source-to-target Sheet replacement path so the operation is represented as `manual_sheet_ops` and is replayed by the existing save pipeline. Refresh the current view as a structural change, invalidate stale row/column projections, and keep the normal mutation guard/interactive mutation lease.

The full-Sheet path must not use the ambiguous logical projection. It must copy the source Sheet structure and cell payload as a Sheet unit. Existing save validation and explicit UTF-8 native replay JSON reading remain the save gates.

## Safety and compatibility

- No partial cell writes occur before a blocking dialog decision.
- Cancel leaves workbooks, undo state, and conflict state unchanged.
- A whole-Sheet choice is explicitly destructive and requires confirmation.
- Tests must verify Chinese text round-trips, OOXML package validity, and Excel/native replay reopen behavior without repair warnings.
