# Global replacement safety

## Requirements

### Requirement: Equal unresolved columns do not block safe global replacement

The system SHALL exclude unresolved logical columns whose source and destination cells are equal for every exact mapped row from global write candidates and SHALL NOT reject the Sheet solely because those columns are unresolved.

#### Scenario: Blank separator columns are equal

- **WHEN** a 2-way Sheet contains equal blank separator columns and differences only in confidently mapped columns
- **THEN** global replacement applies the safe differences and does not write the blank columns

### Requirement: Real ambiguous differences block cell-level global replacement

The system SHALL block cell-level global replacement before the first write when an unresolved or ambiguous logical column contains a real source/destination value or formula difference.

#### Scenario: Repeated template column differs

- **WHEN** a repeated-signature column has a real difference and the mapping cannot uniquely identify its counterpart
- **THEN** no cell-level global write occurs and the user receives a blocker dialog naming the Sheet, column, cause, and affected difference sample

### Requirement: The blocker provides explicit whole-Sheet choices

The system SHALL offer confirmed whole-Sheet overwrite in either supported 2-way direction from the ambiguity blocker.

#### Scenario: User chooses right-to-left whole-Sheet overwrite

- **WHEN** the user confirms the right-side Sheet overwrite
- **THEN** the complete right Sheet replaces the left Sheet through the existing Sheet operation bookkeeping, the action is undoable as one operation, and the view is refreshed as a structural change

### Requirement: Whole-Sheet overwrite warns about data loss

The system SHALL require an explicit confirmation that explains target Sheet content and structure may be replaced or lost before whole-Sheet overwrite begins.

#### Scenario: User cancels destructive confirmation

- **WHEN** the user cancels the whole-Sheet confirmation
- **THEN** no workbook or undo state changes occur

### Requirement: Saved whole-Sheet output is Unicode-safe and Excel-valid

The system SHALL preserve Chinese text through whole-Sheet overwrite and SHALL produce output that passes the existing OOXML validation and reopen checks without an Excel repair warning.

#### Scenario: Save after confirmed overwrite

- **WHEN** a confirmed whole-Sheet overwrite is saved
- **THEN** Chinese values round-trip exactly and the saved workbook/package remains valid for Excel reopening
