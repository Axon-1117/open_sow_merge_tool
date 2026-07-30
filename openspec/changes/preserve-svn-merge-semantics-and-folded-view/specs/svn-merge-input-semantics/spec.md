## ADDED Requirements

### Requirement: Preserve original SVN merge roles
The system SHALL retain the original TortoiseSVN Base, Mine, and Theirs inputs as the authoritative three-way comparison roles and SHALL retain target WC pristine as a separate optional identity.

#### Scenario: Cross-branch source-left is preserved
- **WHEN** TortoiseSVN launches the tool with `.merge-left.rN`, target working file, and `.merge-right.rM`
- **THEN** the scanner SHALL retain `.merge-left.rN` as Source Before, `.merge-right.rM` as Source After, and SHALL NOT replace Source Before with target WC pristine

#### Scenario: Working-copy conflict metadata is available
- **WHEN** the target working copy records exact merge source paths and revisions in its conflict metadata
- **THEN** those repository identities SHALL be used as the primary Source Before/Source After evidence and filename suffix parsing SHALL remain a fallback

#### Scenario: Update old revision is preserved
- **WHEN** TortoiseSVN launches the tool with `.rOLDREV`, local Mine, and `.rNEWREV`
- **THEN** the scanner SHALL use `.rOLDREV` as the common Base even if the post-update WC BASE equals `.rNEWREV`

### Requirement: Classify SVN merge scenario
The system SHALL classify each launch as two-way comparison, update conflict, cross-branch merge, or unknown three-way from raw launch evidence.

#### Scenario: Merge-left and merge-right identify branch merge
- **WHEN** raw Base and Theirs use `.merge-left` and `.merge-right` names
- **THEN** the system SHALL classify the launch as a cross-branch merge and record the evidence

#### Scenario: Old and new revision files identify update
- **WHEN** raw Base and Theirs use ordinary old/new revision sidecar names
- **THEN** the system SHALL classify the launch as an update conflict and record the evidence

### Requirement: Retain four input identities
The system SHALL retain source Base, Mine working, Theirs incoming, and target WC pristine identities independently for diagnostics and decision-making.

#### Scenario: Clean target is not source equality
- **WHEN** Mine equals target WC pristine but differs from source Base
- **THEN** the system SHALL report the working copy as locally clean and SHALL NOT report Mine equal to source Base

#### Scenario: Cross-branch roles are displayed without generic Base ambiguity
- **WHEN** the launch is a cross-branch merge
- **THEN** the four identities SHALL be described as Source Before, Source After, Target Working, and Target Pristine rather than describing Source Before as a branch common ancestor or target Base

### Requirement: Complete equivalence matrix
The system SHALL calculate and log pairwise workbook package equivalence for all available input identities without treating partial workbook comparisons as complete equality.

#### Scenario: OOXML container metadata differs only
- **WHEN** two packages contain the same member paths and identical uncompressed member bytes but ZIP timestamps or entry order differ
- **THEN** the system SHALL report the workbooks package-equivalent

#### Scenario: VBA payload differs
- **WHEN** two `.xlsm` packages differ in their VBA payload
- **THEN** the system SHALL report them not equivalent

#### Scenario: Comparison cannot complete
- **WHEN** an input is incomplete, unreadable, or comparison is cancelled
- **THEN** the system SHALL record an unknown/error matrix result and SHALL NOT use that pair for automatic convergence

### Requirement: Safe automatic convergence
The system SHALL automatically initialize the merged result when complete package equivalence proves an unambiguous whole-workbook result.

#### Scenario: Mine equals Base
- **WHEN** Mine is package-equivalent to Base and Theirs is not
- **THEN** the merged result SHALL be initialized from Theirs and the decision SHALL be logged

#### Scenario: Theirs equals Base
- **WHEN** Theirs is package-equivalent to Base and Mine is not
- **THEN** the merged result SHALL be initialized from Mine and the decision SHALL be logged

#### Scenario: Mine equals Theirs
- **WHEN** Mine and Theirs are package-equivalent
- **THEN** the merged result SHALL be initialized from their common content

#### Scenario: Empty branch delta
- **WHEN** cross-branch source Base and Theirs are package-equivalent
- **THEN** the merged result SHALL remain Mine

### Requirement: Safe semantic pre-merge
The system SHALL automatically apply supported one-sided and identical logical workbook changes and SHALL retain conflicting or unsupported changes for manual review.

#### Scenario: Independent cell changes
- **WHEN** Mine and Theirs change different logical cells relative to the original Base
- **THEN** both changes SHALL be present in the initialized merged result without an unresolved conflict

#### Scenario: Same cell changed differently
- **WHEN** Mine and Theirs change the same logical cell to different results
- **THEN** the cell SHALL remain unresolved and SHALL be available to conflict navigation

#### Scenario: Unsupported package difference
- **WHEN** a package-level change cannot be represented by the supported semantic merge model
- **THEN** the system SHALL not silently discard or auto-accept that change

### Requirement: Source-delta-driven cross-branch merge
The system SHALL derive a cross-branch incoming change set only from Source Before to Source After and SHALL project only that change set onto Target Working.

#### Scenario: Target still contains Source Before value
- **WHEN** an incoming logical value changes from Source Before to Source After and Target Working still equals Source Before
- **THEN** the candidate SHALL receive Source After and the change SHALL be counted as applied

#### Scenario: Target already contains Source After value
- **WHEN** an incoming logical value changes from Source Before to Source After and Target Working already equals Source After
- **THEN** Target Working SHALL remain unchanged and the change SHALL be counted as already present rather than applied

#### Scenario: Target contains a third value
- **WHEN** an incoming logical value changes from Source Before to Source After and Target Working differs from both
- **THEN** the logical location SHALL remain unresolved

#### Scenario: Target has unrelated branch differences
- **WHEN** Target Working differs from Source Before outside the Source Before-to-Source After change set
- **THEN** those target-only differences SHALL be preserved unchanged and SHALL NOT be counted as source changes or automatically merged items

#### Scenario: Structural source delta cannot be mapped safely
- **WHEN** an incoming row or column insertion/deletion cannot be projected onto Target Working with an unambiguous logical mapping
- **THEN** the structural change SHALL remain unresolved without writing by an assumed physical coordinate

### Requirement: Cross-branch outcome counters
The system SHALL report incoming, applied, already-present, target-retained, and unresolved counts separately for cross-branch merges.

#### Scenario: Selected revision is already present in target
- **WHEN** every supported incoming change already equals Target Working
- **THEN** the dialog and log SHALL report zero applied changes, the correct already-present count, zero unresolved conflicts, and an unchanged Target Working candidate

### Requirement: SVN version authorship
The system SHALL display a revision and SVN author status for every available input identity.

#### Scenario: Committed revision author is available
- **WHEN** exact SVN metadata for an input revision is available
- **THEN** its label SHALL include the filename/revision and `Author = <svn-author>`

#### Scenario: Mine contains local edits
- **WHEN** Mine differs from its target WC pristine
- **THEN** the label SHALL show the underlying SVN BASE author and SHALL also state that Mine contains local uncommitted changes

#### Scenario: Author cannot be resolved
- **WHEN** no exact local or remote author metadata is available
- **THEN** the label SHALL show `Author = 未知` and the log SHALL record the lookup failure reason

#### Scenario: Working copy advanced beyond source revision
- **WHEN** Source Before or Source After identifies an exact repository revision but the current working-copy node has advanced to a later revision
- **THEN** author lookup SHALL use stable revision-level repository evidence or a previously verified repository/revision cache rather than degrading solely because the current WC row no longer matches

### Requirement: Auditable merge decisions
The system SHALL log raw identities, normalized stable copies, scenario evidence, author sources, equivalence matrix results, automatic actions, and unresolved counts.

#### Scenario: Production launch is diagnosed
- **WHEN** a three-way launch completes startup analysis
- **THEN** one diagnostic block SHALL contain enough evidence to reproduce why the tool folded, converged, pre-merged, or remained manual
