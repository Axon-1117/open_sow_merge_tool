## MODIFIED Requirements

### Requirement: Only-diff transition is explicit and atomic
The system SHALL accept at most one only-diff transition while exact data is pending and SHALL provide immediate visible feedback without presenting provisional results as exact. The complete exact only-difference result SHALL be stored independently from the bounded virtual viewport.

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
- **THEN** the full exact logical row set and block map are published atomically, only the bounded current viewport is rendered, the calculating status is cleared, and controls are unlocked only if all readiness conditions hold

#### Scenario: Exact result contains thousands of differences
- **WHEN** the exact only-difference result contains more rows than the virtual window limit
- **THEN** counters, navigation, minimap, and operations address the complete result while Tk materializes only the current bounded viewport

#### Scenario: Repeated toggle while calculating
- **WHEN** only-diff is already calculating
- **THEN** another toggle does not start, cancel, or queue a second transition

#### Scenario: User cancels exact calculation
- **WHEN** the user presses Cancel in the only-diff progress dialog
- **THEN** the build generation is invalidated, the dialog and Sheet lock are released, the stable full view and unchecked preference are restored, and a stale worker result cannot alter the logical result or virtual viewport

#### Scenario: Exact calculation fails
- **WHEN** exact-only-diff calculation fails
- **THEN** the stable view remains read-only, the failure is shown non-modally, and retry is available without executing any delayed mutation

### Requirement: Heavy comparison work stays off the Tk thread
The system MUST perform large-Sheet snapshot ingestion, row and column alignment, exact difference generation, formula reconciliation, operation recomputation, and workbook scanning outside the Tk thread. The Tk thread SHALL perform only bounded viewport publication and interaction-state updates.

#### Scenario: Comparison snapshot finishes loading
- **WHEN** a selected large-Sheet snapshot completes
- **THEN** the Tk callback installs immutable prepared state without calling `refresh(rescan=True)`, scanning workbook rows, or materializing the complete result in Tk

#### Scenario: Only-diff calculation is active
- **WHEN** exact differences are calculated for a large Sheet
- **THEN** checkbox, Cancel, and rejected tab callbacks return within 100 ms and the Tk heartbeat has no interval above 200 ms

#### Scenario: Background result publication
- **WHEN** a background comparison completes
- **THEN** the Tk thread installs the generation-safe complete result and renders no more than the bounded virtual window

#### Scenario: User scrolls a prepared result
- **WHEN** a user scrolls, drags the scrollbar, uses the minimap, or navigates to an unrendered difference
- **THEN** the Tk thread reads only prepared snapshot/render data and performs no worksheet access, formula normalization, alignment, or difference computation

#### Scenario: View-only cache data is missing
- **WHEN** hover, click inspection, wheel, page, thumb, minimap, tab activation, C-area refresh, or viewport publication cannot find a prepared value or render fragment
- **THEN** the callback shows a non-actionable placeholder or defers presentation and SHALL NOT call `ws_*_edit`, `_ensure_edit_loaded`, or `_request_edit_preload`

#### Scenario: Remaining Sheet work competes with interaction
- **WHEN** the selected Sheet is exact-ready and hidden Sheets remain pending
- **THEN** remaining work starts only after a 1–2 second UI quiet window, yields at bounded checkpoints for recent UI activity, and a newly selected non-ready Sheet preempts hidden work

### Requirement: Non-ready Sheets never expose provisional comparison results
The system MUST present only generation-matched final exact comparison results as actionable data and MUST clearly cover or replace stale/provisional content while calculation is active.

#### Scenario: User selects a calculating Sheet
- **WHEN** the selected Sheet has not reached exact-same or exact-changed for the current generation
- **THEN** the view shows an unmistakable calculating surface with Sheet name, stage, progress, and elapsed time instead of temporary comparison rows

#### Scenario: User attempts an operation while calculating
- **WHEN** the user invokes copy, overwrite, accept, region, structural, undo/redo, or save before the relevant Sheet generation is exact-ready
- **THEN** the action performs no mutation, queues no delayed action, and immediately opens a modal explaining that exact comparison is still calculating and why the operation is unavailable

#### Scenario: Guarded controls remain explainable
- **WHEN** an operation is unavailable only because exact comparison or full operation detail is not ready
- **THEN** its visible control remains invokable so the click reaches the shared modal, while the handler-side guard remains authoritative and prevents every content, topology, overlay, undo/redo, and save mutation

#### Scenario: Calculation is unresolved or failed
- **WHEN** the user attempts an operation on an unresolved, failed, cancelled, timed-out, or stale Sheet
- **THEN** the modal names the state and reason, the prominent persistent state remains visible, and the action remains blocked until a successful retry publishes an exact result

#### Scenario: Exact result is atomically published
- **WHEN** the current generation reaches exact-same or exact-changed
- **THEN** the calculation surface is replaced once by the complete logical result plus bounded viewport and operations unlock only after every existing formula/cache, topology, and save-readiness gate also passes
