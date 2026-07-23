## Purpose
Define logical column-structure alignment, conflict handling, user actions, save fidelity, and responsiveness guarantees for two-way comparison and three-way Excel merge workflows.

## Requirements

### Requirement: Logical column alignment precedes cell comparison
The system SHALL align retained, inserted, and deleted logical columns before comparing cell values or formulas, and SHALL compare each side through the resulting logical column slots.

#### Scenario: Inserted columns do not shift later differences
- **WHEN** one side inserts two columns before an otherwise unchanged range
- **THEN** the system reports one inserted column block and does not report every retained column to its right as changed

#### Scenario: Deleted columns do not shift later differences
- **WHEN** one side deletes consecutive columns from the middle of a sheet
- **THEN** the system reports one deleted column block and realigns the retained columns after that block

#### Scenario: Independent value edit remains attached to its logical column
- **WHEN** one side inserts a column before an existing logical column and the other side edits a value in that existing column
- **THEN** the value edit is compared in the retained logical slot rather than against the inserted column

### Requirement: Ambiguous column mappings are conservative
The system MUST NOT silently accept a column mapping when duplicate, blank, or conflicting signals make the logical identity ambiguous.

#### Scenario: Duplicate blank columns cannot be distinguished
- **WHEN** multiple adjacent columns have indistinguishable headers, content, formulas, and neighboring context
- **THEN** the system marks the range as structurally unresolved and requires an explicit user decision before column adoption or merged save

#### Scenario: Reliable neighboring anchors bound an ambiguous range
- **WHEN** exact retained columns exist on both sides of an ambiguous inserted/deleted range
- **THEN** the system confines the unresolved state to that range and keeps later anchored columns aligned

### Requirement: Three-way column structure conflicts use Base identity
The system SHALL map mine and theirs to Base independently and SHALL distinguish non-overlapping structural changes from conflicting changes to the same Base column or insertion anchor.

#### Scenario: Independent insertion and value modification merge cleanly
- **WHEN** theirs inserts columns at one Base anchor and mine modifies a retained Base column outside that inserted range
- **THEN** the system reports no structural conflict and preserves both changes

#### Scenario: Column deletion conflicts with modification
- **WHEN** one side deletes a Base column and the other side changes a value, formula, or relevant metadata in that same Base column
- **THEN** the system reports an unresolved delete-versus-modify structural conflict

#### Scenario: Different independent insertions remain distinct
- **WHEN** mine and theirs insert different columns at non-overlapping Base anchors
- **THEN** the merged logical model retains both inserted blocks in deterministic Base-relative order

#### Scenario: Competing insertions at one ambiguous anchor require a decision
- **WHEN** both sides insert incompatible column blocks at the same Base anchor and their correspondence cannot be proven
- **THEN** the system reports a structural conflict instead of pairing them by physical index

#### Scenario: Proven common insertions share logical slots
- **WHEN** Mine and Theirs contain uniquely identifiable, exactly equal value/formula columns at the same Base insertion anchor
- **THEN** the system pairs each proven ordered match into one common logical slot while leaving any content-different remainder as a competing-insertion conflict

### Requirement: Column structure is visible and actionable
The system SHALL render missing-side placeholders and column-block markers from the logical slot model and SHALL provide Mine, Base, and Theirs actions for the selected structural column block where applicable.

#### Scenario: Inserted block is visible without cell-difference flooding
- **WHEN** a side contains an inserted column block
- **THEN** the view identifies the block as a column insertion, shows placeholders on the missing side, and keeps following column headers and cells aligned

#### Scenario: Adopt a selected inserted column block
- **WHEN** the user selects an inserted column block and invokes the target-side adopt action
- **THEN** the system inserts exactly that block with its values, formulas, widths, hidden state, styles, comments, hyperlinks, and validation metadata

#### Scenario: Adopt a selected deleted column block
- **WHEN** the user accepts a source-side deletion for a selected Base column block
- **THEN** the system records deletion of exactly those logical columns and leaves adjacent blocks unchanged

#### Scenario: Row and column structural differences coexist
- **WHEN** a sheet contains inserted/deleted rows and inserted/deleted columns
- **THEN** row navigation, column selection, only-diff presentation, and region actions resolve through both stable mappings without changing an adjacent logical row or column

### Requirement: Column actions are atomic and undoable
The system SHALL apply each user-initiated column block action atomically and SHALL restore workbook data, mappings, selection, and difference state when the action is undone or fails.

#### Scenario: Undo inserted column adoption
- **WHEN** the user adopts an inserted column block and then invokes undo
- **THEN** the original physical columns, logical slot mapping, formulas, metadata, selection, and difference markers are restored

#### Scenario: Native replay fails during save
- **WHEN** Excel native column replay or output validation fails
- **THEN** the target file remains unchanged, the temporary output is not presented as successful, and the unresolved user operations remain recoverable in the open session

### Requirement: Saved workbooks preserve column fidelity
The system SHALL save accepted column insertions and deletions through a path that preserves formulas and advanced workbook content, and MUST validate the merged package before atomic replacement.

#### Scenario: Formula references follow an inserted column
- **WHEN** an adopted column insertion shifts formulas with relative, absolute, shared, array, or data-table semantics
- **THEN** supported formulas retain Excel-equivalent references and unsupported unsafe transformations stop the save with a clear error

#### Scenario: Advanced column metadata survives save and reopen
- **WHEN** a structural column action affects a sheet with widths, hidden state, styles, comments, hyperlinks, validation, merged cells, conditional formatting, external links, or macros
- **THEN** the merged workbook reopens successfully and unaffected advanced content remains equivalent to the source policy selected by the user

#### Scenario: Untouched OOXML parts remain stable
- **WHEN** a column action changes only one sheet
- **THEN** critical package parts and untouched sheets that do not require Excel normalization remain unchanged or pass an explicitly documented native-normalization comparison

### Requirement: Column alignment meets responsiveness gates
The system MUST build and use column mappings from sequential in-memory caches without adding high-row random worksheet reads during scrolling, selection, navigation, or column-block presentation.

#### Scenario: Large real sheet mapping
- **WHEN** column alignment runs on a representative large workbook from the project data set
- **THEN** mapping completes within an acceptance threshold recorded before implementation and subsequent UI updates use cached slot mappings

#### Scenario: Column block action remains responsive
- **WHEN** the user adopts a consecutive structural column block
- **THEN** the UI reports progress, batches the operation, and meets the established interaction and save thresholds without blocking unrelated cancellation or close handling

### Requirement: Existing comparison and merge behavior remains compatible
The system SHALL preserve existing row alignment, formula-cache, only-diff block, minimap, hover/C-area, sheet-level, XLSM, and SVN merge behavior for workbooks without column structural changes.

#### Scenario: Workbook has no column structure change
- **WHEN** the compared sheets have the same logical columns
- **THEN** existing cell, row, region, sheet, formula, save, and navigation results remain unchanged

#### Scenario: Column mapping cache becomes stale
- **WHEN** a row or column structural edit changes the model version
- **THEN** stale column slots are invalidated before rendering, conflict scanning, adoption, undo, or save continues
