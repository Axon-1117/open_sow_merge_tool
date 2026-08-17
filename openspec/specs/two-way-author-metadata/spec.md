# 2-way author metadata

## Requirements

### Requirement: Resolve author metadata for 2-way SVN inputs

The system SHALL preserve the raw 2-way comparison identities and resolve
author/revision metadata through the same exact SVN/WC lookup path used by
3-way startup before constructing the comparison UI.

#### Scenario: SVN 2-way comparison has author metadata

- **WHEN** a 2-way comparison is launched with a revision-side input and a
  working-copy input
- **THEN** the app context contains Mine and the older-side identities with
  their proven revisions and resolved authors when SVN metadata is available

#### Scenario: Local 2-way comparison has no SVN metadata

- **WHEN** a 2-way comparison is launched from ordinary local workbook paths
- **THEN** comparison still opens and unavailable author fields use the
  existing `未知` fallback without inventing metadata

### Requirement: Display author metadata in both 2-way panes

The system SHALL show each visible 2-way pane's semantic role, author, and
revision in the identity bar, and SHALL expose the full identity and lookup
reason through the existing hover detail.

#### Scenario: 2-way identity bar shows both authors

- **WHEN** Mine and the older-side identities have resolved authors
- **THEN** the left and right identity labels include the corresponding
  author and revision instead of only the workbook path

#### Scenario: Unavailable author remains diagnosable

- **WHEN** one 2-way identity author cannot be resolved
- **THEN** its label shows `未知` and its hover detail preserves the reason,
  while the other pane remains fully populated
