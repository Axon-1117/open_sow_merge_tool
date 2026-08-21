## ADDED Requirements

### Requirement: Large Sheets load through selected-Sheet snapshots
The system SHALL construct comparison data for a large Sheet from sequential read-only value and formula streams and SHALL NOT require normal-mode materialization of every Sheet in every workbook side before presenting an exact view.

#### Scenario: User opens one large Sheet
- **WHEN** a workbook contains multiple Sheets and the user selects one large Sheet
- **THEN** the system parses the selected Sheet into an immutable semantic snapshot while unopened Sheets remain unparsed beyond workbook metadata

#### Scenario: Formula-aware snapshot is complete
- **WHEN** a selected Sheet contains literals, formulas, missing or present cached values, dates, booleans, errors, blanks, or external references
- **THEN** the snapshot preserves the typed comparison information required by the existing exact formula/cache semantics

#### Scenario: Cached Sheet is revisited
- **WHEN** the user returns to a Sheet whose file signature and comparison generation are unchanged
- **THEN** the system reuses the exact snapshot without worksheet iteration and completes the revisit within the recorded 100 ms P95 acceptance threshold

### Requirement: Schema-aware record and field alignment is deterministic
The system MUST use validated declared field and record identities before content similarity and MUST refuse an automatic mapping when evidence is ambiguous.

#### Scenario: Sheet has a unique declared key
- **WHEN** all applicable `@id` or `@const` components form a unique record key on each compared side
- **THEN** the system aligns records by the composite key in linear expected time and reports insertions, deletions, edits, and order changes separately

#### Scenario: Keyed records contain continuation rows
- **WHEN** blank-key rows follow a unique keyed record
- **THEN** the system associates those rows with that record and performs only bounded local alignment inside the continuation group

#### Scenario: Declared key is duplicated or missing
- **WHEN** declared identities cannot uniquely align an unmatched region
- **THEN** the system uses unique row-hash anchors and bounded deterministic fallback, and marks any remaining large or duplicate-ambiguous block unresolved rather than guessing

#### Scenario: Columns are inserted, deleted, or reordered
- **WHEN** field declaration and type rows identify logical columns whose physical positions differ
- **THEN** the system aligns those columns by schema identity while preserving insertion, deletion, reorder, and ambiguity as explicit structural results

### Requirement: Duplicate declared field identities are admitted only by an all-side bounded proof
The system SHALL treat a repeated normalized declared field identity as
unresolved unless an immutable snapshot-only proof identifies every occurrence
without weakening any row or column safety gate. `START` and `END` MAY be used
only as virtual unique outer schema anchors. The proof SHALL clear only the
duplicate-field-identity cause; every other unresolved cause remains terminal
and non-actionable.

#### Scenario: Same-ordinal duplicate run is proven in two-way comparison
- **WHEN** both sides have complete validated declared row keys, full
  one-to-one row coverage, a duplicate field run bounded by the same unique or
  virtual `START`/`END` anchors, equal run width and ordinal, all-blank
  value/formula tokens, and equal complete typed/formula-aware column digests
- **THEN** the engine may create non-actionable proof anchors and clear only
  the duplicate-field identity cause after either verifying an existing exact,
  non-structural, fully bijective top-level cache with one same-logical slot
  per proof pair, or rebuilding the complete bounded interval when that cache
  is pending. The rebuild path requires every non-proof interval field to have
  the same unique ordered declaration/type sequence and count on both sides;
  formula changes in those non-proof fields remain ordinary exact cell
  differences

#### Scenario: Same-ordinal duplicate run is proven in three-way comparison
- **WHEN** Mine, Base, and Theirs satisfy the same row, anchor, width, ordinal,
  digest, and final cross-side-bijection proof for every duplicate member
- **THEN** it may first verify a fully resolved, non-structural, all-side
  bijective top-level cache with one same-logical Mine/Base/Theirs proof slot,
  even if Mine-to-Base and Theirs-to-Base child anchor gaps are asymmetric. If
  that top cache is pending, it may instead use the bounded interval rebuild
  only when Mine, Base, and Theirs have the complete same unique ordered
  non-proof interval sequence; either path preserves Base physical coordinates
  and every ordinary 3-way difference/conflict rule

#### Scenario: Exact top cache rejects crosswire or structural evidence
- **WHEN** duplicate proof coordinates exist but an otherwise non-pending
  top-level cache has a crosswired proof triple, a non-retained/ambiguous slot,
  a structural column, an unresolved column, or an incomplete side bijection
- **THEN** the proof SHALL remain terminally `UNRESOLVED`; it may not use the
  exact-top-cache path or infer a replacement from asymmetric child gaps

#### Scenario: Duplicate proof is incomplete or disagrees
- **WHEN** any candidate has nonblank content or an unequal digest, missing/
  extra run member, unequal interval width or ordinal, missing/reordered anchor,
  incomplete/ambiguous row key or row pair, unresolved/ambiguous/non-bijective
  non-proof interval slot, Base mismatch, a formula token in a proof column,
  a formula-structure/mapping failure outside the proven blank columns, or
  candidate-builder exception
- **THEN** the Sheet SHALL remain `UNRESOLVED`, expose no automatic mutation or
  physical operation target, and SHALL NOT invoke a legacy mapping fallback

#### Scenario: Formula-different unique interval members remain content diffs
- **WHEN** a blank duplicate proof anchors an interval whose non-proof members
  are unique, same ordered/count on every side, and fully retained/bijective,
  but one or more of those members has a different formula identity
- **THEN** an inherited `formula-identity-mismatch` on a blank proof slot may
  be discharged only by the two-stage rebuild, while the non-proof formula
  cells remain visible exact differences and no structural mapping is hidden

#### Scenario: Non-duplicate ambiguity remains blocked
- **WHEN** a Sheet has any row, schema, column-cache, parser, cancellation, or
  generation failure other than the narrowly proven duplicate-field identity
- **THEN** the duplicate proof SHALL NOT clear or mask that failure and the
  Sheet remains terminally non-actionable

### Requirement: New exact results match a frozen Oracle
The system SHALL compare the new semantic result against a normalized legacy exact Oracle before the new path becomes the default.

#### Scenario: Unambiguous 2-way fixture
- **WHEN** the old and new engines compare the same unambiguous 2-way real or synthetic fixture
- **THEN** their logical rows, columns, cell differences, formulas, cached-value decisions, structural changes, and only-difference membership are identical

#### Scenario: Unambiguous 3-way fixture
- **WHEN** the old and new engines compare the same Mine/Base/Theirs fixture
- **THEN** their source changes, local changes, conflicts, auto-merge eligibility, and target cells are identical

#### Scenario: Ambiguous fixture
- **WHEN** the legacy engine produces a mapping but the new engine lacks sufficient stable evidence
- **THEN** the new engine may report an unresolved block but SHALL NOT silently choose a different row or cell as an automatic mutation target

#### Scenario: Oracle uses real source workbooks
- **WHEN** baseline or Oracle tests use files under `C:\GM15\design\sheets\develop`
- **THEN** those source files are opened read-only and every mutation/save occurs only on a disposable copy outside the source directory

### Requirement: Large comparison has bounded performance and memory
The system SHALL record cold and warm performance in fresh processes and SHALL meet the accepted large-workbook thresholds without weakening correctness.

#### Scenario: Representative Skill Sheet
- **WHEN** the accepted Skill large-Sheet fixture is compared on the reference workstation
- **THEN** exact selected-Sheet comparison completes within the 15-second release maximum and evidence also reports progress against the 1.5-second stretch target

#### Scenario: Representative 18k-to-21k-row Sheet
- **WHEN** the accepted WorldMonsterSurvivor or Language fixture is compared on the reference workstation
- **THEN** exact selected-Sheet comparison completes within the 15-second release maximum and evidence also reports progress against the 2.5-second stretch target

#### Scenario: Three-way Skill memory
- **WHEN** three Skill sides are compared through the new large-Sheet path
- **THEN** peak RSS is at least 50 percent below the measured legacy path and does not exceed 400 MB on the reference workstation

#### Scenario: Stale comparison finishes
- **WHEN** a file, topology, mutation, cancellation, or retry advances the Sheet generation during background comparison
- **THEN** the stale result cannot replace the current snapshot, render into the selected view, or unlock mutation controls

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
- **THEN** navigation may publish the exact-same or exact-changed badge from a bounded cache, but that cache SHALL NOT be exposed as actionable comparison rows

#### Scenario: User selects a summary-exact Sheet without complete operation detail
- **WHEN** a user selects a Sheet whose whole-workbook summary is exact but whose full logical display and physical operation targets are incomplete
- **THEN** the selected Sheet enters calculating state, shows only the calculation surface, asynchronously materializes the complete immutable result, and unlocks operations only after that full result is generation-current

#### Scenario: Comparison detail is exact before editable backend load
- **WHEN** the current-generation terminal result, complete prepared rows, formula/cache data, column/Base mappings, and physical operation targets are installed but normal-mode editable workbooks remain deferred
- **THEN** the exact result is visible and the calculation surface is removed, while mutation/save handlers remain guarded and an attempted operation opens the readiness modal, starts one backend load, and performs no mutation until a later retry

#### Scenario: Sheet has no differences
- **WHEN** the current generation's exact 2-way or 3-way comparison completes with no semantic or structural difference
- **THEN** the Sheet is marked exact-same and may be counted as unchanged

#### Scenario: Sheet has differences
- **WHEN** the current generation's exact comparison contains any semantic, formula/cache, conflict, row, column, or Sheet-structure difference
- **THEN** the Sheet is marked exact-changed and its navigation badge is visually prominent

#### Scenario: Comparison cannot prove a result
- **WHEN** comparison is unresolved, cancelled, stale, timed out, or failed
- **THEN** the Sheet is marked with that non-exact state, is not counted as unchanged, and mutation controls remain locked

### Requirement: Complete real corpus meets the accepted exact-readiness maximum
The system SHALL benchmark every supported workbook and every Sheet under `C:\GM15\design\sheets\develop` in fresh-process 2-way and 3-way runs using disposable copies and SHALL reach final exact readiness within 15 seconds per Sheet.

#### Scenario: Corpus inventory runs
- **WHEN** the benchmark enumerates the source directory recursively
- **THEN** it records every file, explicitly classifies unsupported or temporary files, and never writes beneath the source directory

#### Scenario: Two-way corpus run
- **WHEN** a supported source workbook is measured in 2-way mode
- **THEN** each Sheet is requested, reaches a final exact state within 15 seconds, produces the correct self-comparison zero-difference Oracle, and records startup, Sheet, whole-workbook, memory, engine, and fallback metrics

#### Scenario: Three-way corpus run
- **WHEN** a supported source workbook is measured with independent disposable Mine/Base/Theirs copies
- **THEN** each Sheet reaches a final exact state within 15 seconds with no false conflict/change and records the same per-file and per-Sheet metrics

#### Scenario: Slowest Sheets are ranked
- **WHEN** a complete corpus run finishes
- **THEN** machine-readable evidence and a human-readable report list every timeout/failure and rank the slowest Sheets separately for 2-way and 3-way optimization

#### Scenario: Direct Oracle timing is kept separate from application opening timing
- **WHEN** parser/comparator workers and the production GUI entry point are benchmarked
- **THEN** evidence labels them as separate tiers, and only a fresh `SowMergeApp` child measured from constructor request through selected-Sheet terminal full comparison detail and complete physical targets is used to pass the 15-second user-visible opening gate, without proactively loading editable workbooks

#### Scenario: First mutation backend load is measured separately
- **WHEN** an exact Sheet is visible while editable workbooks remain deferred
- **THEN** the corpus records view-only memory separately, then measures the explanatory first click, single-owner backend load, retry, accepted operation, undo/redo, and save as mutation phases rather than application opening time
