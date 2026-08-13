## ADDED Requirements

### Requirement: Native replay preserves Unicode payloads

The system SHALL decode native replay operation payloads as UTF-8 before assigning text values to Excel, regardless of the Windows PowerShell system code page.

#### Scenario: 2-way save with Chinese cell content

- **WHEN** a 2-way save replays a changed Chinese text cell through Excel native replay
- **THEN** the saved workbook contains the exact original Unicode text
- **AND** the saved package passes OOXML validation and can be reopened without a repair prompt

#### Scenario: 3-way save with Chinese cell content

- **WHEN** a 3-way merge save replays changed Chinese text from Mine, Base, or Theirs
- **THEN** the saved workbook contains the exact original Unicode text
- **AND** the saved package passes OOXML validation and can be reopened without a repair prompt

### Requirement: Fallback and cell-only saves preserve workbook text

The system SHALL preserve Unicode text and valid OOXML structure when native replay is unavailable and the save uses the openpyxl fallback or cell-only XML patch path.

#### Scenario: Native Excel is unavailable

- **WHEN** a structural save falls back to openpyxl
- **THEN** Chinese text remains unchanged after reopening the output
- **AND** the output is not replaced if package validation fails

#### Scenario: Cell-only XML patch

- **WHEN** a 2-way or 3-way save changes only cell values
- **THEN** the XML patch writes the exact Unicode text and preserves untouched workbook parts

### Requirement: Semantically empty cells do not produce value diffs

The system SHALL compare an absent/empty cell and an exact empty-string cell as equal when neither side has a formula or special-formula identity.

#### Scenario: Reported GunshipsMaster cells

- **WHEN** the `GunshipsMaster@design` comparison sees `None` on one side and `""` on the other at `F5`, `H5:H7`, or `H9:H11`
- **THEN** those cells SHALL not be marked as differences

#### Scenario: 2-way and 3-way blank equivalence

- **WHEN** the same blank representation mismatch is evaluated in a 2-way or 3-way row comparison
- **THEN** the value-level comparison SHALL report equality
- **AND** non-empty text, whitespace-only text, distinct formulas, and formula-vs-literal changes SHALL remain differences

### Requirement: Worksheets without declared dimensions are compared by actual content

The system SHALL calculate the real worksheet range before comparison when a valid XLSX omits the optional worksheet `<dimension>` element and read-only parsing exposes an unknown row or column bound.

#### Scenario: Revision export lacks a dimension element

- **WHEN** a revision-exported XLSX contains populated cells but its worksheet XML omits `<dimension>`
- **THEN** the comparison SHALL scan the actual worksheet data to establish its row and column bounds
- **AND** populated columns SHALL not be reported as deleted solely because the declared dimension is absent
