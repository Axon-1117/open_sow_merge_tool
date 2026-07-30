## Context

TortoiseSVN already launches the tool with four merge arguments: `/base`, `/mine`, `/theirs`, and `/merged`. For update conflicts, Base is the old common revision; for cross-branch merges, Base is the source-left revision used to compute the incoming change. The current startup path then prefers a newly exported WC BASE from the target working copy and overwrites that original Base before conflict scanning. Because all three-way classification is Base-anchored, the replacement changes the meaning of every downstream result.

Excel workbooks are SVN binary files. SVN can replace an unmodified whole file, but it cannot apply row, column, cell, formula, or workbook-structure changes to a diverged target workbook. The tool therefore needs both the true source delta and the target state to perform its own semantic merge.

The application is a single-process Tk desktop tool. Workbook loading and diff/merge work can be expensive, `.xlsm` must retain VBA, and the UI must not expose editable controls until role discovery and semantic comparison finish.

## Goals / Non-Goals

**Goals:**

- Preserve the original SVN merge inputs and make their roles explicit.
- Retain target WC pristine as a fourth diagnostic/decision identity without substituting it for Base.
- Distinguish ordinary two-way comparison, update conflict, cross-branch merge, and unknown three-way launch modes.
- Attach the best available SVN revision and author information to every displayed identity.
- Produce a complete, auditable equivalence matrix for Base, Mine, Theirs, and target pristine.
- Automatically converge proven whole-workbook identities and automatically pre-merge supported non-overlapping logical changes.
- Keep unresolved overlapping changes in the existing manual workflow.
- Fold only a proven redundant pane and allow the user to expand it without losing state.
- Make the surrounding workspace visually identify the launch mode while preserving all spreadsheet-cell rendering.

**Non-Goals:**

- Changing SVN server configuration, repository MIME properties, merge tracking, or marking conflicts resolved without an explicit user save/exit action.
- Treating a clean B working copy as proof that it equals the A source-left revision.
- Claiming semantic equality when comparison coverage is incomplete.
- Adding support for legacy `.xls` or `.xlsb` formats.
- Replacing the current row/column logical alignment algorithm wholesale.

## Decisions

### Preserve a five-path launch context

Use a launch-context record with:

- `source_base_path`: the original `/base` supplied by TortoiseSVN;
- `mine_path`: the original `/mine` working side;
- `theirs_path`: the original `/theirs` incoming/source-right side;
- `target_pristine_path`: an optional read-only export of `args.merged@BASE`;
- `merged_path`: the output working file.

The first three are the only inputs to three-way conflict classification. Target pristine is used for diagnostics, authorship, and the separate statement “working copy is locally clean.” It SHALL NOT replace source Base.

The previous revision-export fallback derived a repository path by stripping `.merge-left.rN` from a target-branch path. That is unsafe for cross-branch merges. A complete sidecar supplied by SVN is used directly after stability/package validation. Revision re-export is allowed only when the exact repository URL for that identity is known.

For cross-branch merge conflicts, the scenario-specific display and decision roles are:

- `Source Before`: merge-left, the source-path snapshot at the exclusive start of the merge range;
- `Source After`: merge-right, the source-path snapshot at the inclusive end of the range;
- `Target Working`: the target-branch working file;
- `Target Pristine`: the target working-copy pristine file used only to identify local uncommitted edits.

Source Before is not described as a branch common ancestor or as target WC BASE. When the working-copy conflict record is available, its source repository path and revisions are primary identity evidence; filename suffix parsing is a fallback.

### Classify the launch from raw evidence

- `.merge-left.rN` plus `.merge-right.rM` means cross-branch merge.
- `.rOLDREV` plus `.rNEWREV`, excluding merge-left/right, means update conflict.
- Missing Theirs means two-way comparison.
- Ambiguous complete three-way arguments mean unknown three-way mode and use neutral styling.

Classification, raw paths, parsed revisions, and the reason are logged before any normalized copies are made.

### Use fail-closed package equivalence

Whole-workbook equivalence is determined by comparing the sorted OOXML package member set and every member's uncompressed bytes. This ignores ZIP container timestamps/order while covering formulas, cached values, styles, relationships, names, validations, conditional formatting, tables, drawings, external links, workbook metadata, and VBA payloads. Exact package equality is conservative: logically equivalent but differently serialized XML can be reported unequal, but differing workbook content is never reported equal.

The matrix stores, for every pair:

- raw file SHA-256 equality;
- complete package equality;
- readiness/error state;
- elapsed time and comparison reason.

Automatic whole-workbook convergence requires complete package equality. A partial openpyxl comparison is never sufficient for that decision.

### Separate convergence from semantic pre-merge

Whole-workbook convergence rules:

- `Mine ≡ Base` and `Theirs !≡ Base` → initialize result from Theirs.
- `Theirs ≡ Base` and `Mine !≡ Base` → initialize result from Mine.
- `Mine ≡ Theirs` → initialize result from their common workbook.
- Cross-branch `Mine ≡ source-left` → initialize result from source-right/Theirs.
- Cross-branch `source-left ≡ source-right` → initialize result from Mine.

If no whole-workbook rule applies, reuse the existing Base-anchored logical row/column/cell classifier to pre-merge changes that are proven one-sided or identical. Unsupported workbook-level package differences, ambiguous structural mappings, and cells changed differently on both sides remain unresolved.

Automatic work initializes the in-memory/output candidate and reports what happened. It does not silently invoke SVN “resolved.” The user retains the final Save Merged/exit confirmation.

### Drive cross-branch cherry-picks from the source delta

For a cross-branch merge, first derive the supported logical incoming delta from `Source Before → Source After`. Project only that delta onto a candidate initialized from Target Working:

- target value equals Source Before → apply Source After;
- target value equals Source After → classify as already present and leave it unchanged;
- target value differs from both → retain an unresolved conflict;
- target-only differences outside the incoming delta → preserve unchanged and do not count as automatic merges.

The same rule applies to supported logical row and column changes after alignment. Ambiguous structural projection remains manual. A branch common ancestor is neither required nor substituted for Source Before because it would expand a single-revision cherry-pick into unrelated branch history.

The outcome records distinct counts for incoming changes, applied changes, already-present changes, target-retained differences, and unresolved conflicts. `merged_count` remains a compatibility aggregate where required, but user-facing branch-merge explanations use the distinct counters.

### Fold presentation without deleting data

The application always retains all three source models. If the matrix proves one pane redundant, that pane is hidden by default and replaced by a compact strip explaining the equality and offering “展开三方”. Expanding/restoring panes preserves sheet, scroll position, selected cell, pending merge choices, and output data.

The fold decision is presentation-only. Auto-convergence and pre-merge decisions are recorded separately.

### Apply colors only to surrounding chrome

Use centralized workspace colors:

- two-way: existing gray;
- update conflict: soft pink;
- cross-branch merge: soft green;
- unknown three-way: existing gray.

Only root/frame chrome, gaps, title strips, and non-grid containers inherit the mode color. Spreadsheet canvases/text widgets stay white, and existing difference, conflict, cursor, selection, row/column action, and disabled-state colors remain unchanged. Contrast is validated at the pane boundary and for all labels placed on the colored chrome.

### Resolve SVN authors with layered evidence

Each identity carries path, display revision, repository-relative identity when known, author, author source, and lookup status.

Lookup order:

1. SVN metadata already available from the target working-copy database for WC identities.
2. An available `svn info --xml`/`svn log --xml` client using an exact peg URL/path and revision.
3. Existing TortoiseSVN log-cache metadata when it can be read without UI.
4. `未知` with a logged lookup reason.

For Mine, the displayed SVN author is the author of its underlying WC BASE; when Mine has local changes, the label also states “本地未提交修改” instead of attributing those edits to a repository author.

### Explain automatic outcomes once

After startup analysis, show one modal summary for three-way launches:

- scenario and identity labels;
- equivalence facts used;
- number of automatically merged and unresolved items;
- whether a pane was folded;
- whether the result was initialized from Mine, Theirs, or a semantic pre-merge.

For cross-branch merges, labels and counts use Source Before, Source After, Target Working, and Target Pristine. The dialog distinguishes “applied” from “already present” and never presents unrelated target-branch differences as changes merged from the selected source revision.

The modal is informational and dismissible. If unresolved conflicts remain, its primary action enters the first unresolved conflict; otherwise it focuses the Save Merged action. Detailed evidence stays in the log.

### Treat workbook markers as review evidence, not cell addresses

Conflict accounting may retain workbook-level and structural markers such as
`<workbook>` so unsupported OOXML representation changes remain visible and
fail closed. These markers are not Sheet names. Navigation filters candidates
against the real displayed Sheet catalog and positive row/logical-column
coordinates. It skips Sheets with no real conflict and selects the first
navigable cell in workbook order.

If unresolved accounting contains only workbook-level or structural evidence,
the startup dialog changes its primary action to full three-way/manual review.
It never passes a pseudo Sheet to the Notebook selector.

### Make structural choices discoverable without weakening write safety

After a Sheet reaches READY, the first unresolved structural column block is
selected automatically when the user has not already made an explicit
selection. This enables the semantic choices relevant to the current scenario,
including keeping Target Working and adopting Source Before or Source After.

Automatic selection is presentation state only. The existing authoritative
projection revalidation, ambiguous-mapping confirmation, native Excel replay,
rollback snapshot, and save-time validation remain mandatory. A disabled
choice must carry an explicit reason; “no hidden selection” is not an
acceptable explanation when a unique actionable block is already known.

### Use one compact Sheet host and semantic identity strip

Keep `ttk.Notebook` as the internal lazy Sheet host but apply a client-only
style that hides its duplicate tab row. The existing lower Sheet buttons
remain the single visible navigator and continue to select Notebook pages.

Remove raw SVN arguments, current-read paths, build text, and aggregate Sheet
counts from the permanent top area; those details are already logged or
available through diagnostics. Keep only the compact action toolbar and the
load/progress status needed for feedback.

Each pane identity label prioritizes:

1. scenario-specific semantic role;
2. workbook basename and revision;
3. `Author = ...`.

The full local/repository path is attached as hover help and remains in the
diagnostic log. This prevents path length from clipping the author.

The C-area row comparison remains content-sized. The hover panel uses one
heading line that includes the current Sheet/column/cell identity and reserves
only the height required by its visible source rows and scrollbar. Freed
vertical pixels belong to the expanding main spreadsheet paned window.

### Prioritize actions and advance structural work

Place the three structural-column decisions in a fixed, non-shrinking action
group. Show only the selected block and remaining-block count in the permanent
status text; retain the complete structural summary and reason in hover help
and diagnostics. General Sheet statistics must yield horizontal space before
any merge action is clipped.

After a successful retain/apply transaction and its authoritative projection
refresh, clear the completed selection and search the refreshed model for the
next actionable structural block. Select that block using the new projection
generation and keep actions enabled. Disable actions only when the refreshed
model proves there are no remaining structural blocks, and say so explicitly.

### Make historical Author attribution stable

An exact repository revision has one revision-property author even when a
working-copy file later advances. Retain exact path matching as the first local
lookup, then use repository-UUID plus revision evidence from an available SVN
client/local metadata source. Cache only successfully verified
repository-UUID/revision/author results, never guessed basename matches.
Failures remain non-blocking and are logged with the attempted evidence chain.

### Keep action guidance spatially attached

Use one compact action row for the Previous Difference/status/Next Difference
group and the structural actions. Reserve the three column buttons on the right
and place a compact tagged status immediately before them. Render the first
Excel-style logical-column token (`T`, `T:U`) with a red foreground while
keeping the complete cause and structure summary in hover diagnostics.

### Add an atomic current-Sheet global cell action

Extend the existing row/region scope variable with `global`. Global scope uses
the same side direction as the split button but derives its targets from the
authoritative full-Sheet logical diff map, never from filtered/visible rows.
Before writing, build and validate the complete cell plan, reject structural or
ambiguous prerequisites, and capture one workbook/comparison snapshot. Apply
the batch as one transaction and push one undo record; any failure restores the
snapshot and leaves the Sheet unchanged. Show a confirmation containing Sheet,
source role, and target-cell count.

### Use one bounded uniform column width per Sheet

Choose one fixed readable width for the Sheet rather than deriving a different
width for each column from cell content. Materialize an equal-valued immutable
width tuple for all logical slots and reuse it for all main panes, column
headers, minimap/cursor lines, and C-area comparisons. Projection changes may
change the tuple length but ordinary content, filtering, scrolling, hover, and
selection do not alter the width value.

All user-facing logical-column headers use `openpyxl.utils.get_column_letter`
(`A..Z`, `AA..`) while internal diagnostic keys may retain `L<n>`.

### Make the fixed-width model pixel-stable for CJK

Use one installed CJK monospace editor font whose Latin/space glyph is exactly
one half of its Chinese glyph. Format values by normalized East Asian display
width rather than Python code-point count: narrow characters occupy one slot,
East Asian Wide/Fullwidth characters occupy two, Latin-script Ambiguous
characters plus `™`/`®` occupy one, other Ambiguous characters occupy two, and
combining marks occupy zero. Truncation reserves the display width of its
marker. Keep the rendered string's Tk index horizon compatible with existing
fixed spans by inserting zero-width index placeholders for double-width
glyphs.

The main panes, Base pane, C-area rows, and every column header use the same
font and formatter. Selection/difference tags may change color, underline, or
background but never replace the editor font with a metric-changing bold font.
This retains the existing cached spans and hit-testing model while making
GunshipsModify A1:A5 land on identical pixel boundaries.

### Combine navigation and structural actions responsively

Use one fixed-height merge-action surface. Keep the structural status and
three column buttons packed at the right edge. Place the difference-navigation
group at the absolute center while it fits; when its right edge would collide
with the structural cluster, shift it left just enough to preserve a fixed
gap. This removes one redundant vertical row without sacrificing narrow-window
action visibility.

General transient Sheet feedback remains on the left and must yield before it
can overlap either action group.

### Use stable utility order and Excel column names

Configure the root utility toolbar with the four buttons in columns zero
through three and put the expanding grid weight after them, producing the
left-to-right order `重算并刷新`, `导出诊断包`, `复制反馈信息`, `检查更新`.

Centralize user-facing logical-column and range formatting through
`get_column_letter`. Column-action statuses, hover help, confirmations,
blockers, conflict locations, and completion messages call that formatter.
Internal cache keys and diagnostic evidence may continue to use `L<n>`.

## Risks / Trade-offs

- [Package equality is conservative and may miss logically equivalent workbooks] → Treat false negatives as full/manual three-way display; never relax equality using incomplete comparison.
- [Opening and hashing four large workbooks can increase startup time] → Stream hashes, cache by stable path/size/mtime, reuse already stabilized copies, and perform analysis behind the existing startup progress/read-only gate.
- [Automatic logical pre-merge can expose unsupported package-level changes] → Gate pre-merge on supported change coverage and leave unsupported differences unresolved.
- [Author lookup may require credentials/network or an unavailable SVN CLI] → Use layered local metadata and show `Author = 未知` without blocking merge.
- [Color propagation can accidentally recolor spreadsheet canvases] → Centralize chrome-only application and add pixel/widget-color GUI assertions.
- [Legacy tests currently encode WC BASE substitution] → Replace them with raw-source-left expectations and add production launch-path tests.

## Migration Plan

1. Introduce launch-context, scenario classification, authorship records, and equivalence logging without changing UI behavior.
2. Stop substituting WC BASE and update scan/merge tests to use the original Base.
3. Enable conservative convergence and supported semantic pre-merge behind the startup analysis gate.
4. Add folded-pane state, outcome dialog, and mode-specific chrome colors.
5. Run smoke, unit, GUI, real-workbook, `.xlsm`, and performance acceptance suites.
6. Package to a separate acceptance build. Rollback is switching back to the previous release executable; no workbook or repository migration is required.
