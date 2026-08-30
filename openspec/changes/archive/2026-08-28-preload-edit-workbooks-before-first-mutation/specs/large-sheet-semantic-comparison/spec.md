## MODIFIED Requirements

### Requirement: Whole-workbook Sheet status is exact and explicit
The system SHALL eventually compute a generation-matched exact status for every supported Sheet and SHALL NOT represent pending, calculating, stale, unresolved, cancelled, or failed work as unchanged.

#### Scenario: Workbook is opened
- **WHEN** workbook metadata and Sheet names become available
- **THEN** every Sheet is initially marked pending or calculating, the selected Sheet is prioritized, and changed/unchanged badges are withheld until an exact result is published

#### Scenario: Background workbook scan progresses
- **WHEN** the selected Sheet reaches a terminal state
- **THEN** remaining Sheets are compared in the background with visible exact-count and changed-count progress until the whole-workbook summary is complete

#### Scenario: Hidden large Sheet has an exact bounded summary
- **WHEN** background comparison proves the current generation's semantic and structural Sheet status but full per-cell display and operation data has not been materialized
- **THEN** navigation may publish the exact-same or exact-changed badge from a bounded cache, but that cache SHALL NOT be exposed as actionable comparison rows or start editable-workbook preload

#### Scenario: User selects a summary-exact Sheet without complete operation detail
- **WHEN** a user selects a Sheet whose whole-workbook summary is exact but whose full logical display and physical operation targets are incomplete
- **THEN** the selected Sheet enters calculating state, shows only the calculation surface, asynchronously materializes the complete immutable result, and unlocks operations only after that full result is generation-current

#### Scenario: Comparison detail is exact before editable backend load
- **WHEN** the current-generation terminal result, complete prepared rows, formula/cache data, column/Base mappings, and physical operation targets are installed while normal-mode editable workbooks are not yet ready
- **THEN** the exact result SHALL be visible without waiting for editable materialization, the system SHALL automatically start one background editable-workbook loader, and mutation/save handlers SHALL remain guarded only until that loader publishes readiness

#### Scenario: First row overwrite occurs after automatic preload
- **WHEN** the initial selected Sheet is exact-ready and the automatically started editable-workbook loader has completed before the user requests a row overwrite
- **THEN** the first row-overwrite request SHALL execute normally without an editable-loading retry modal

#### Scenario: User acts while automatic preload is still running
- **WHEN** a mutation or save is requested after automatic preload starts but before editable workbooks are ready
- **THEN** the system SHALL reject that action without mutation, queuing, or automatic replay and SHALL explain that editable workbooks are still loading

#### Scenario: Sheet has no differences
- **WHEN** the current generation's exact 2-way or 3-way comparison completes with no semantic or structural difference
- **THEN** the Sheet is marked exact-same and may be counted as unchanged

#### Scenario: Sheet has differences
- **WHEN** the current generation's exact comparison contains any semantic, formula/cache, conflict, row, column, or Sheet-structure difference
- **THEN** the Sheet is marked exact-changed and its navigation badge is visually prominent

#### Scenario: Comparison cannot prove a result
- **WHEN** comparison is unresolved, cancelled, stale, timed out, or failed
- **THEN** the Sheet is marked with that non-exact state, is not counted as unchanged, mutation controls remain locked, and the unresolved result SHALL NOT start automatic editable preload

### Requirement: Complete real corpus meets the accepted exact-readiness maximum
The system SHALL benchmark every supported workbook and every Sheet under `C:\GM15\design\sheets\develop` in fresh-process 2-way and 3-way runs using disposable copies and SHALL reach final exact readiness within 15 seconds per Sheet.

#### Scenario: Corpus inventory runs
- **WHEN** the benchmark enumerates the source directory recursively
- **THEN** it records every file, explicitly classifies unsupported or temporary files, and never writes beneath the source directory

#### Scenario: Two-way corpus run
- **WHEN** a supported source workbook is measured in 2-way mode
- **THEN** each Sheet is requested, reaches a final exact state within 15 seconds, produces the correct self-comparison zero-difference Oracle, and records startup, Sheet, whole-workbook, memory, and engine/fallback metrics

#### Scenario: Three-way corpus run
- **WHEN** a supported source workbook is measured with independent disposable Mine/Base/Theirs copies
- **THEN** each Sheet reaches a final exact state within 15 seconds with no false conflict/change and records the same per-file and per-Sheet metrics

#### Scenario: Slowest Sheets are ranked
- **WHEN** a complete corpus run finishes
- **THEN** machine-readable evidence and a human-readable report list every timeout/failure and rank the slowest Sheets separately for 2-way and 3-way optimization

#### Scenario: Direct Oracle timing is kept separate from application opening timing
- **WHEN** parser/comparator workers and the production GUI entry point are benchmarked
- **THEN** evidence labels them as separate tiers, and only a fresh `SowMergeApp` child measured from constructor request through selected-Sheet terminal full comparison detail and complete physical targets is used to pass the 15-second user-visible opening gate; the gate SHALL NOT wait for the post-publication automatic editable preload to finish

#### Scenario: Automatic mutation backend warmup is measured separately
- **WHEN** an exact selected Sheet publishes while editable workbooks remain deferred
- **THEN** the corpus records view-only opening time separately, then measures automatic single-owner backend load, edit-ready latency and memory, first accepted operation, undo/redo, and save as mutation-readiness phases
