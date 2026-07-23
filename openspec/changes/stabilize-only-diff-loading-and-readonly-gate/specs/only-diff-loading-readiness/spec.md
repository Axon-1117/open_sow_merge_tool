## ADDED Requirements

### Requirement: Sheet readiness controls mutation
The system SHALL maintain readiness per Sheet and SHALL permit a workbook mutation only when that Sheet's requested comparison, row model, logical column mapping, and editable workbooks are ready for the current generation.

#### Scenario: Loading Sheet remains view-only
- **WHEN** a Sheet is loading, calculating exact differences, reconciling formulas, failed, canceled, or applying another mutation
- **THEN** scrolling, selection, text copying, Sheet switching, and status inspection remain available while all mutation entry points are rejected

#### Scenario: Modal exact-only-diff calculation owns interaction
- **WHEN** the user explicitly starts an exact-only-diff calculation from a READY full view
- **THEN** the current Sheet and Sheet switching are unavailable until the calculation succeeds, fails, or is canceled, while the progress dialog remains operable

#### Scenario: Readiness is isolated per Sheet
- **WHEN** one Sheet is READY and another Sheet is still calculating
- **THEN** the READY Sheet can be edited while selecting the calculating Sheet presents a read-only view

#### Scenario: Hidden mutation binding cannot bypass the gate
- **WHEN** a non-READY Sheet receives a row-header click, pane double-click, C-area action, keyboard command, toolbar action, undo, or save request
- **THEN** no workbook, manual operation record, touched row, modified flag, or undo entry changes

#### Scenario: Exact state unlocks atomically
- **WHEN** the current generation has exact requested comparison data, current mappings, and editable workbooks
- **THEN** the Sheet enters READY and all applicable controls are enabled in one UI update

### Requirement: Only-diff transition is explicit and atomic
The system SHALL accept at most one only-diff transition while exact data is pending and SHALL provide immediate visible feedback without presenting provisional results as exact.

#### Scenario: User enables only-diff from a ready full view
- **WHEN** the user enables only-diff and no valid exact snapshot exists
- **THEN** the checkbox becomes checked and locked without changing label or position, a modal progress dialog appears within 100 ms, the Sheet becomes read-only, and the current stable view remains until exact results are published

#### Scenario: Exact progress is measurable
- **WHEN** a user-triggered exact-only-diff calculation processes a large Sheet
- **THEN** the dialog identifies the Sheet and stage and reports processed rows, total rows, percentage, and elapsed time without blocking Tk repaint

#### Scenario: Sheet opens with only-diff requested
- **WHEN** a Sheet is first opened with only-diff requested
- **THEN** the checkbox is checked and locked and any temporary preview is explicitly labelled read-only rather than changing the user-facing mode to full

#### Scenario: Exact snapshot becomes ready
- **WHEN** the matching exact-only-diff generation completes successfully
- **THEN** the full exact row set is rendered atomically, the calculating status is cleared, and controls are unlocked only if all readiness conditions hold

#### Scenario: Repeated toggle while calculating
- **WHEN** only-diff is already calculating
- **THEN** another toggle does not start, cancel, or queue a second transition

#### Scenario: User cancels exact calculation
- **WHEN** the user presses Cancel in the only-diff progress dialog
- **THEN** the build generation is invalidated, the dialog and Sheet lock are released, the stable full view and unchecked preference are restored, and a stale worker result cannot alter the UI

#### Scenario: Exact calculation fails
- **WHEN** exact-only-diff calculation fails
- **THEN** the stable view remains read-only, the failure is shown non-modally, and retry is available without executing any delayed mutation

### Requirement: Heavy comparison work stays off the Tk thread
The system MUST perform large-Sheet row alignment, exact difference generation, formula reconciliation, and workbook scanning outside the Tk thread.

#### Scenario: Editable workbooks finish loading
- **WHEN** editable workbook preload completes for one or more materialized large Sheets
- **THEN** the Tk callback updates readiness and applies prepared data without calling `refresh(rescan=True)` or scanning workbook rows

#### Scenario: Only-diff calculation is active
- **WHEN** exact differences are calculated for a large Sheet
- **THEN** checkbox, Cancel, and rejected tab callbacks return within 100 ms and the Tk heartbeat has no interval above 200 ms

#### Scenario: Background result publication
- **WHEN** a background comparison completes
- **THEN** the Tk thread performs only bounded cache/state installation and visible rendering

### Requirement: Background comparison is reusable and generation-safe
The system SHALL distinguish cache completeness, reuse formula-aware data when an exact request is known, and reject results from stale generations.

#### Scenario: Large cache is not exact
- **WHEN** a large-Sheet cache only confirms that some difference exists
- **THEN** it is not marked as a complete pair-difference or only-diff snapshot

#### Scenario: Only-diff requested before first compute
- **WHEN** a large Sheet is queued with only-diff already requested
- **THEN** its first formula-aware background pass produces the exact only-diff snapshot without reopening the same workbook set for a second full scan

#### Scenario: User changes Sheet during computation
- **WHEN** a result completes for a Sheet or generation that is no longer active
- **THEN** it may be cached for that Sheet but cannot render into, unlock, or otherwise alter the selected Sheet

#### Scenario: Comparison is superseded
- **WHEN** row structure, column structure, manual refresh, retry, or cancellation advances a Sheet generation
- **THEN** every older queued or running result is ignored

### Requirement: Active Sheet owns exact-compute priority
The system SHALL run at most one priority exact-only-diff worker and SHALL prefer the currently selected Sheet.

#### Scenario: Switch away from a calculating Sheet
- **WHEN** ordinary background comparison is active and no modal exact-only-diff request owns the selected Sheet
- **THEN** selecting a different Sheet reprioritizes it and the hidden request is canceled or downgraded

#### Scenario: Modal calculation pins selected Sheet
- **WHEN** a user-triggered exact-only-diff progress dialog owns one Sheet
- **THEN** notebook tabs and bottom Sheet navigation cannot select a different Sheet until the owner calculation exits

#### Scenario: Hidden Sheet finishes
- **WHEN** an exact result for a hidden Sheet is accepted
- **THEN** it updates cache/readiness only and performs no visible Text widget rebuild

#### Scenario: Queued Sheet is selected
- **WHEN** a not-yet-started Sheet is selected
- **THEN** it moves to the front of background comparison work and visible loading feedback appears within 100 ms

### Requirement: Window state remains stable
The system SHALL preserve the user's main-window state during startup promotion, loading, only-diff transitions, and asynchronous result application.

#### Scenario: Maximized window enables only-diff
- **WHEN** the main window is maximized before only-diff calculation begins
- **THEN** it remains maximized with unchanged restored geometry after calculation and rendering complete

#### Scenario: Normal window enables only-diff
- **WHEN** the user has placed the main window in normal state
- **THEN** no loading or refresh callback maximizes, normalizes, or repositions it

#### Scenario: Startup progress becomes main UI
- **WHEN** the startup progress root is promoted to the main application
- **THEN** the incomplete normal-size main window is not exposed before layout and intended startup state are applied

### Requirement: Exactness and save behavior remain compatible
The system SHALL preserve existing comparison, merge, undo, formula, structural replay, and save correctness while changing loading behavior.

#### Scenario: Exact only-diff matches oracle
- **WHEN** background exact-only-diff completes for a 2-way or 3-way workbook
- **THEN** its row set and per-cell differences equal the synchronous exact comparison oracle

#### Scenario: WorldMonster conflict workflow
- **WHEN** isolated WorldMonster Mine/Base/Theirs copies are opened, only-diff is requested during editable preload, and the user changes Sheets
- **THEN** the UI remains responsive, no provisional write is possible, the final exact views are correct, and the source files remain unchanged

#### Scenario: Save after READY mutations
- **WHEN** the user performs supported row, region, cell, column, or Sheet actions after READY and saves
- **THEN** existing native replay, atomic replacement, formula-cache, and reopen-validation guarantees remain in force

### Requirement: C-area maximizes comparison visibility
The system SHALL let the C-area row comparison use all available horizontal space in its lower pane while preserving logical horizontal synchronization with the main views.

#### Scenario: C-area is wider than one main pane
- **WHEN** the window has enough width to show multiple main panes
- **THEN** the C-area comparison body occupies the available lower width and exposes more logical columns than any one main pane

#### Scenario: C-area header stays aligned
- **WHEN** the main or C-area horizontal position changes
- **THEN** the C-area header and body share the same logical start and matching viewport geometry without truncating the body to one main-pane width

### Requirement: Only-diff control geometry remains stable
The system SHALL keep the only-diff checkbox at a fixed toolbar location independent of its value and changing difference statistics.

#### Scenario: Only-diff state changes
- **WHEN** only-diff changes between unchecked, calculating, checked, failed, or canceled states
- **THEN** the checkbox label and screen position remain unchanged

#### Scenario: Difference-block status changes
- **WHEN** difference-block ordinal, total, pending count, or calculation status changes
- **THEN** that information appears in a separate reserved status area and does not reflow the only-diff control
