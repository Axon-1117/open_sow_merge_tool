## ADDED Requirements

### Requirement: Content operations use a Sheet-local overlay
The system SHALL represent accepted cell, row, and region content changes as typed overlay deltas on the immutable Sheet snapshot and SHALL update only affected comparison rows and block boundaries.

#### Scenario: Single-cell adoption
- **WHEN** the user adopts one value or formula in a large READY Sheet
- **THEN** the system records one undoable overlay delta, recomputes the affected logical row, updates the visible viewport if necessary, and completes within 250 ms P95 before save

#### Scenario: Thousand-row region adoption
- **WHEN** the user adopts a contiguous region containing 1,000 affected rows
- **THEN** the system batches overlay writes and comparison updates, publishes one bounded UI update, and completes within 2.0 seconds P95 before save

#### Scenario: Operation targets an unrendered row
- **WHEN** a block or region operation includes rows outside the virtual viewport
- **THEN** every logical target is updated in the overlay without materializing those rows in Tk

### Requirement: Undo and redo reuse bounded comparison updates
The system SHALL undo or redo content overlays without rescanning unrelated workbook rows or rebuilding unrelated Sheet snapshots.

#### Scenario: Undo content-only operation
- **WHEN** the latest operation changes no row, column, or Sheet topology
- **THEN** undo restores its overlay delta, updates only affected result records, and completes within 500 ms P95

#### Scenario: Redo content-only operation
- **WHEN** a previously undone content operation is redone
- **THEN** redo reapplies the same logical targets and bounded updates without full row or column alignment

#### Scenario: Undo structural operation
- **WHEN** undo changes row, column, or Sheet topology
- **THEN** the affected Sheet advances its topology generation and rebuilds off the Tk thread while unrelated Sheets remain valid

### Requirement: Structural operations invalidate only necessary topology
The system SHALL distinguish content changes from row, column, and Sheet topology changes and SHALL never use a stale logical mapping as a mutation target.

#### Scenario: Row structure changes
- **WHEN** rows are inserted, deleted, or reordered
- **THEN** the affected Sheet's record alignment, complete result, block map, and virtual mappings are regenerated before further mutations are enabled

#### Scenario: Column structure changes
- **WHEN** columns are inserted, deleted, reordered, or accepted
- **THEN** the affected Sheet's field identities, formula projections, column actions, and visible snapshot are regenerated before further mutations are enabled

#### Scenario: Another Sheet is unchanged
- **WHEN** a topology operation completes on one Sheet
- **THEN** exact snapshots and virtual windows for unrelated unchanged Sheets remain reusable

### Requirement: Accelerated operations preserve save fidelity
The system MUST replay accepted overlays and structural operations through the established safe-save boundary and MUST NOT treat UI success as save validation.

#### Scenario: Cell-only overlay is saved
- **WHEN** a workbook with content-only overlays is saved
- **THEN** the output contains the accepted typed values, formulas, and cached-value decisions and passes package validation, atomic replacement, and required Excel reopen checks

#### Scenario: Structural overlay is saved
- **WHEN** accepted operations require native Excel structural semantics
- **THEN** the system stages an immutable source copy, uses the existing native replay path, validates the package, and preserves formula/reference behavior

#### Scenario: Save fails
- **WHEN** ZIP patching, native replay, validation, replacement, or reopen fails
- **THEN** the original workbook remains recoverable and the operation overlay remains available for retry or diagnosis

### Requirement: Every operation is guarded by final exact readiness
The system SHALL use one centralized operation guard and SHALL NOT execute or defer a user mutation against a pending, calculating, stale, unresolved, cancelled, timed-out, or failed Sheet result.

#### Scenario: Copy or overwrite is requested too early
- **WHEN** the user requests a cell, row, block, region, or Sheet copy/overwrite before exact readiness
- **THEN** no overlay/manual/native operation is recorded and a modal identifies the Sheet, state, calculation stage, and retry condition

#### Scenario: Structural action is requested too early
- **WHEN** the user requests row, column, or Sheet insertion/deletion/adoption while the logical mapping is not exact-current
- **THEN** no topology generation is mutated and the same readiness modal explains that stable logical targets are unavailable

#### Scenario: Save is requested with a non-ready affected Sheet
- **WHEN** any Sheet contributing operations to the save target is not exact-ready for its current generation
- **THEN** save staging does not begin and the modal lists the blocking Sheets and reasons

### Requirement: Performance optimization cannot weaken Excel compatibility
The system MUST treat exact operation targeting and Excel-compatible saved output as hard release gates for every performance iteration.

#### Scenario: Slow-Sheet operation variant is accepted
- **WHEN** deterministic cell, formula, row, or column changes are applied to a disposable variant of a slow real Sheet
- **THEN** the overlay/manual operation result matches the exact Oracle before save and the saved workbook preserves the accepted result after package validation and reopen

#### Scenario: Representative XLSM or metadata-rich workbook is saved
- **WHEN** the corpus includes VBA, comments, links, formulas, row/column metadata, or workbook relationships
- **THEN** the optimized path preserves those supported features through atomic staging and Excel reopen exactly as required by the existing save contract
