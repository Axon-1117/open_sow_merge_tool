## Context

The current comparison model aligns rows but compares cells at the same physical column index. A structural change before column `C` therefore makes every following column appear different, and the manual-operation model can replay row and sheet structure but has no equivalent column operation. The change crosses comparison, 3-way conflict detection, rendering, hit testing, region actions, undo, formula/cache handling, and fidelity-preserving save paths.

Representative workbooks contain formulas, merged cells, validation, comments, hyperlinks, hidden columns, widths, external links, macros, and sheets with duplicate or blank headers. Alignment must therefore be conservative: a visible warning and explicit unresolved mapping is safer than silently pairing the wrong business fields.

## Goals / Non-Goals

**Goals:**

- Align logical columns before cell comparison so insert/delete operations do not create cascading false differences.
- Use one column-slot model for 2-way display, 3-way conflict detection, column-level adoption, undo, and save.
- Preserve formula and workbook fidelity through existing native-save safety rules.
- Keep column analysis sequential and cache-backed, with predictable behavior on large sheets.
- Provide conservative, testable behavior for ambiguous or independently inserted columns.

**Non-Goals:**

- Inferring business-schema renames when a column's identity and all contents changed without reliable anchors.
- Replacing the existing row-alignment model or redesigning the workbook/sheet navigation UI.
- Automatically resolving delete-versus-modify conflicts.
- Adding a non-Excel dependency or changing CLI arguments and persisted settings.

## Decisions

### Build side-to-side mappings through logical column slots

Introduce a column-slot record containing a stable logical ordinal and optional physical column index for mine, Base, and theirs. A 2-way view builds slots directly between A and B. A 3-way view first maps each side to Base, anchors retained Base columns, then places side-only insertions between their neighboring Base anchors.

Rendering, comparison, selection, minimap data, and copy actions consume slot ordinals instead of assuming one physical index applies to every side. Missing columns render as structural placeholders rather than shifting later columns.

This is preferred over special-casing the first mismatching header because sheets can have multiple header rows, blank headers, duplicate labels, and formulas.

### Derive column identity from cached multi-signal signatures

Build signatures from existing sequential row caches after row alignment. Signals include normalized values/formula identity from bounded header and representative data rows, non-empty patterns, number-format/style fingerprints where available, and neighboring-column context. Sequence alignment supplies candidate mappings; exact anchors and high-confidence candidates are accepted automatically.

Ambiguous duplicate/blank columns remain explicitly unresolved. The UI reports the affected range and permits an explicit keep/adopt decision; it does not silently shift the remainder based on a low-confidence guess. Signatures and mappings are versioned with row-model and edit versions so structural changes cannot reuse stale slots.

### Use Base-anchored 3-way structural rules

For each Base column, each side is classified as retained, modified, or deleted. Side-only insertions are anchored before/after a retained Base slot. Independent operations merge when their anchors and affected Base columns do not overlap. Insertions from Mine and Theirs at the same anchor share one logical slot only when a unique full-column value/formula digest from the sequential cache proves exact correspondence (including an adopted full-column copy); a partially matching block merges only its proven ordered matches. Delete-versus-modify, incompatible replacements of the same Base column, and every unproved or content-different insertion remainder are structural conflicts requiring user choice.

Column alignment happens before cell conflicts are scanned. Value/formula conflicts are then compared within the same logical slot, preventing a preceding insertion from manufacturing false conflicts.

### Add explicit column operations and column-level actions

Extend the manual-operation model with ordered `insert_cols`, `delete_cols`, and copy/adopt metadata containing target slot, count, source side, and source physical columns. The view exposes column-block selection and Mine/Base/Theirs column-level actions while retaining existing cell, row, region, and sheet actions.

Column actions operate on the stable slot model, update physical-to-logical mappings atomically, invalidate affected row/text caches, and push one undo record per user action. A failed action restores the entire prior column model and workbook state.

### Replay structural columns natively and validate the output

Column insertion/deletion or whole-column adoption forces Excel COM/native replay for workbooks whose formulas or advanced objects cannot be proven safe under openpyxl. Native copy transfers full column content and metadata, then explicit cell/formula-cache operations apply final values. Formula references are adjusted by Excel's structural operations; manually copied formulas use the existing translation and special-formula safety checks.

The merged file is written to a temporary path, reopened/validated as an OOXML package, and atomically replaces the target only after success. Native failure never falls back to a path known to lose formulas, macros, external links, validation, merged ranges, or other advanced content.

### Bound performance with sequential caches and conservative limits

Column signature construction performs a single pass over already captured row data and is `O(rows × columns)` for the selected sheet, without high-row random worksheet reads. The column count is small relative to rows in representative files; pairwise dynamic programming is limited to the column axis. Cached mappings are reused by scrolling, selection, navigation, and rendering.

Performance tests record mapping time, first interactive time, structural-action latency, save latency, and peak memory on synthetic and real workbooks. If a sheet exceeds configured safe limits or mapping confidence is insufficient, the system presents an unresolved structural range instead of performing an expensive or speculative alignment.

## Risks / Trade-offs

- [Risk] Duplicate or content-identical columns can be paired incorrectly. -> Require high-confidence anchors, use neighbor context, expose ambiguity, and test duplicate/blank-header cases.
- [Risk] Row and column alignment can invalidate each other's signatures. -> Build column signatures from a stable row-pair snapshot and version both models; rebuild in a defined order after structural edits.
- [Risk] Excel structural replay can change formulas or advanced objects unexpectedly. -> Force native replay, validate output, compare critical OOXML parts, and keep atomic rollback.
- [Risk] Column placeholders can break text hit testing and horizontal synchronization. -> Route display positions through slot mappings and add 2-way/3-way GUI geometry tests.
- [Risk] Full-column copy can be slow on wide or heavily formatted sheets. -> Batch consecutive operations, pause low-priority scanning, show progress, and establish real-file performance gates.
- [Trade-off] Conservative ambiguity leaves some ranges unresolved. -> This avoids silent data corruption and gives the user an explicit structural decision.

## Migration Plan

1. Add the model and detection behind an internal feature flag while keeping physical-index comparison as a rollback path.
2. Enable it first for read-only comparison and diagnostics, then for 2-way actions, 3-way conflicts, undo, and finally native save.
3. Run the complete existing regression matrix plus new real-workbook column scenarios before enabling by default.
4. Roll back by disabling column alignment and column actions; workbook and settings formats remain unchanged.

## Open Questions

- Whether the first release should include a manual column-mapping dialog for ambiguous ranges or only block the action with diagnostics. The implementation may start with blocking plus clear diagnostics, provided it never silently aligns an ambiguous range.
