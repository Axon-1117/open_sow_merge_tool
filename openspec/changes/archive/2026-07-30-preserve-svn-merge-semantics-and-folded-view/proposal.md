## Why

The merge tool currently replaces TortoiseSVN's original merge-left/base input with the target working copy's WC BASE. That can erase the source revision delta, manufacture misleading `Base = Theirs` or `Mine = Base` relationships, and make production conflict classification differ from the SVN operation that produced the conflict.

Excel users also need to recognize update conflicts versus cross-branch merges immediately, understand who authored every input revision, and avoid spending screen space on a third pane when two inputs are proven semantically identical.

## What Changes

- Preserve TortoiseSVN's original Base/Mine/Theirs inputs and separately retain the target working-copy pristine identity instead of replacing Base.
- Classify update conflicts and cross-branch merges from the raw SVN inputs and expose the four relevant identities: source/base, working Mine, incoming Theirs, and target WC pristine.
- Query and display SVN author information for each available version identity.
- Compute and log a complete workbook semantic-equivalence matrix, with explicit comparison coverage and fail-closed behavior.
- Automatically converge unambiguous three-way cases and automatically merge non-overlapping workbook changes while retaining manual review for real conflicts.
- Add an expandable folded-three-way workspace that hides a proven redundant pane without discarding its data.
- Distinguish surrounding workspace colors while preserving the current white spreadsheet canvases and all cell-difference/selection rendering:
  - two-way comparison: current gray;
  - update conflict: pink;
  - cross-branch merge: light green.
- Explain automatic convergence, automatic merge, and remaining-conflict outcomes through concise dialogs and diagnostic logs.
- In cross-branch cherry-picks, name merge-left/merge-right as Source Before/Source After, derive the incoming delta only from that pair, and project only that delta onto Target Working while preserving unrelated target-branch differences.
- Use working-copy conflict metadata as the primary source-path/revision evidence when available, and report applied, already-present, target-retained, and unresolved counts separately.
- Make structural column conflicts immediately actionable, prevent workbook-level review markers from being treated as Sheet cells, and compact the workspace to one visible Sheet navigator with legible Author-first identities and a taller main grid.
- Keep merge controls visible ahead of verbose difference summaries, automatically advance through every remaining structural column choice, and resolve historical source-revision authors even after the working copy advances.
- Center difference-block navigation, place the pending structural-column cue directly beside its actions with a red logical-column token, add an atomic current-Sheet global cell-adoption mode, and use stable per-Sheet column widths plus Excel-style column labels in both the main and hover comparison views.
- Make fixed-width rendering pixel-stable for mixed Latin/CJK data, combine difference navigation and structural-column actions into one space-efficient row, left-align the root utility actions, and remove internal `L<n>` notation from all user-facing column guidance.

## Capabilities

### New Capabilities

- `svn-merge-input-semantics`: Correct SVN input-role preservation, scenario classification, target-pristine identity, author attribution, semantic-equivalence evidence, logging, and safe automatic convergence.
- `adaptive-three-way-workspace`: Mode-specific workspace cues, expandable folded-three-way presentation, preserved spreadsheet rendering, and user-facing automatic-merge outcome explanations.

### Modified Capabilities

None.

## Impact

- Primary implementation: `sow_merge_tool.py`.
- Test impact: three-way scan/merge unit tests, launch-path regression tests, real workbook fixtures, Tk GUI self-tests, author-query tests, and performance baselines.
- SVN integration: TortoiseSVN `/base`, `/mine`, `/theirs`, `/merged` launch arguments plus read-only working-copy metadata and author lookup.
- No change to the persisted Excel file formats; `.xlsx` and `.xlsm` output must preserve currently supported workbook structures and VBA payloads.
