# Only-diff block navigation

## Requirements

### Requirement: Exact full-snapshot difference blocks
When precise only-diff data is ready, the system SHALL group the complete only-diff snapshot into blocks using logical pair continuity, not screen-line adjacency or the current render limit.

#### Scenario: Separate worksheet regions remain separate blocks
- **WHEN** the only-diff snapshot contains pair indices `10, 11, 50, 51`
- **THEN** the system reports two blocks, `10-11` and `50-51`, even though those rows are adjacent on screen

#### Scenario: Block count includes rows beyond the initial render limit
- **WHEN** a large sheet has difference blocks after the first 800 rendered rows
- **THEN** the reported total includes those blocks without requiring the user to scroll or load every preceding screen row

#### Scenario: Count waits for precise data
- **WHEN** precise only-diff data is still being calculated
- **THEN** the block indicator reports a calculating state instead of presenting a provisional count as exact

### Requirement: Visible and color-independent block boundaries
In only-diff mode, the system SHALL render a visible boundary before every block after the first and SHALL identify the block independently of background color.

#### Scenario: Adjacent filtered rows belong to different blocks
- **WHEN** the last displayed row of one block is immediately followed by the first displayed row of another block
- **THEN** the main panes show synchronized vertical spacing and the left row-number area identifies the new block number

#### Scenario: Two-way and three-way panes remain aligned
- **WHEN** a block boundary is displayed in 2-way or 3-way mode
- **THEN** every visible data pane and row-number pane applies the same vertical spacing so rows remain horizontally aligned

#### Scenario: Full view remains unchanged
- **WHEN** the user disables only-diff mode
- **THEN** block-only spacing and block markers are hidden and the normal full-sheet row layout is restored

### Requirement: Stable block counter and status
The Sheet toolbar SHALL show the active block number, total snapshot block count, and number of blocks that still contain unresolved visual differences.

#### Scenario: Selected row determines active block
- **WHEN** the user explicitly selects a row that belongs to a difference block
- **THEN** the toolbar shows that block as the active block

#### Scenario: Viewport determines active block without selection
- **WHEN** there is no explicit selection
- **THEN** the toolbar shows the block containing the top visible difference row

#### Scenario: Resolved block keeps its identity
- **WHEN** all remaining differences in a block are adopted and its touched rows remain visible in snapshot mode
- **THEN** the block keeps its original number, is marked processed, and the pending block count decreases without renumbering later blocks

### Requirement: Full-snapshot block navigation
Previous and next block actions, including their keyboard shortcuts, SHALL navigate using the complete block model and SHALL preserve horizontal scroll position.

#### Scenario: Navigate to an unrendered large-sheet block
- **WHEN** the next block is outside the current 800-row render window
- **THEN** the system materializes enough rows to display the target block and scrolls to its first row without a full worksheet rescan

#### Scenario: Navigate while horizontally scrolled
- **WHEN** the user navigates between blocks while a nonzero horizontal scroll position is active
- **THEN** the target block becomes vertically visible and the horizontal position remains unchanged across all synchronized panes

#### Scenario: Navigation state at endpoints
- **WHEN** the active block is the first or last block
- **THEN** the unavailable previous or next action is disabled while the other action remains available when applicable

### Requirement: Region adoption uses displayed block boundaries
Region-level adoption SHALL resolve its scope from the same stable block model used by the visual boundary and navigation features.

#### Scenario: Adopt the selected displayed block
- **WHEN** the user selects a row in block 3 and invokes a region-level mine, theirs, or Base action
- **THEN** the operation considers only pair rows belonging to block 3 and does not modify rows in adjacent blocks

#### Scenario: Previously resolved rows inside a block
- **WHEN** a stable block contains touched rows that no longer have visual differences
- **THEN** region adoption skips those resolved rows while continuing through the remaining unresolved rows in the same block

#### Scenario: Structural difference block
- **WHEN** a block contains one-sided inserted or deleted row slots
- **THEN** region adoption retains the existing batch insert/delete semantics for those rows and does not merge the block with a noncontiguous neighboring block

### Requirement: Block presentation does not degrade large-sheet performance
The system MUST derive and render block metadata without additional worksheet scans or high-row random cell reads.

#### Scenario: Build block metadata from cached differences
- **WHEN** precise only-diff pair indices are available
- **THEN** block construction performs a linear pass over those cached indices and does not read workbook cells

#### Scenario: Scroll and selection updates
- **WHEN** the user scrolls, selects a row, or navigates between blocks
- **THEN** the active block indicator updates from cached mappings without triggering difference recomputation
