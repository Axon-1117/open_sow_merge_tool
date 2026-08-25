## ADDED Requirements

### Requirement: Result rendering is bounded by the viewport
The system SHALL keep the complete logical result independently from the Tk documents and SHALL render only the visible logical window plus bounded overscan.

#### Scenario: Sheet contains 20,000 changed rows
- **WHEN** full or only-difference mode contains 20,000 logical result rows
- **THEN** each main Tk pane contains no more than 320 logical rows while the scrollbar, counters, minimap, and block totals represent the complete result

#### Scenario: User scrolls to a distant window
- **WHEN** the user moves from the first result rows to a middle or final viewport
- **THEN** the system renders the target slice directly without materializing intervening rows

#### Scenario: Small result set
- **WHEN** the complete result fits inside the configured virtual window
- **THEN** the view preserves current small-Sheet behavior, selection, copy, highlighting, and pane alignment

### Requirement: Scrolling performs no comparison or workbook I/O
The system MUST serve wheel, page, thumb, minimap, and programmatic scroll requests exclusively from the immutable comparison/render snapshot.

#### Scenario: High-rate wheel scrolling
- **WHEN** multiple wheel events arrive faster than the render frame budget
- **THEN** the system coalesces them to the newest logical position and performs no worksheet cell access, worksheet iteration, formula normalization, alignment, or diff recomputation

#### Scenario: Scroll reaches the end of a large result
- **WHEN** the viewport approaches or reaches the final logical row
- **THEN** the system replaces the fixed viewport rather than appending another 500-row batch to the Tk documents

#### Scenario: Virtual scroll performance
- **WHEN** automated scrolling traverses a 20,000-difference fixture on the reference workstation
- **THEN** viewport render P95 is at most 33 ms and the Tk heartbeat has no interval above 200 ms

#### Scenario: Changed-revision interaction remains view-only
- **WHEN** the disposable `Dungeon.xlsx` revision-39265 versus revision-39264 fixture is exercised in 2-way or 3-way mode with continuous hover, wheel, thumb, and tab interaction
- **THEN** viewport render P95 is at most 33 ms, no heartbeat exceeds 200 ms, no editable-workbook load is requested, and both `Dungeon` and `MonsterGroup` publish current-generation exact results within 15 seconds

### Requirement: Logical interaction survives virtualization
The system SHALL resolve every visible action through stable logical row and column identities rather than permanent Tk line numbers.

#### Scenario: Selection crosses virtual windows
- **WHEN** a selected row or cell leaves and later re-enters the viewport
- **THEN** its logical selection remains stable and is restored at the correct record and field

#### Scenario: Difference navigation targets an unrendered block
- **WHEN** next/previous difference, block navigation, or minimap selection targets an unrendered result
- **THEN** the system moves the virtual window directly to that target and selects it without scanning or rendering intermediate results

#### Scenario: Row or region action in a virtual window
- **WHEN** the user invokes a supported action on a visible virtualized row or block
- **THEN** the action targets the corresponding complete-result identity and never a recycled visible line from another window

#### Scenario: Text copy in a virtual window
- **WHEN** the user copies visible cell or row text
- **THEN** the copied content and pane identity match the logical result currently displayed

### Requirement: Virtual views expose duplicate declared fields only after bounded proof
The system SHALL not render an actionable virtual result for a Sheet whose
duplicate normalized field identity remains unresolved.  It MAY expose a
duplicate occurrence only after the immutable snapshot comparison proves the
same occurrence across every side by validated row keys and complete row pairs,
same unique-or-virtual-`START`/`END` anchor interval, equal interval width and
ordinal, all-blank cached-value/formula tokens, equal complete
typed/formula-aware column digest, and a two-stage final cache proof.  Stage
one installs no actionable result. Stage two either verifies an already exact,
non-structural, cross-side-bijective top cache with exactly one same-logical
slot per proof occurrence, or—only when that top cache is pending—rebuilds the
bounded interval and requires every non-proof interval field to be unique in
the same ordered declaration/type sequence and count across sides. An exact
three-way top cache may pass despite asymmetric Mine/Base and Theirs/Base child
anchor gaps; that exception never permits child-gap mapping inference. The
proof clears only the duplicate-field reason; it does not mask any other
ambiguity.

#### Scenario: Interior and END-bounded duplicate occurrences are proven
- **WHEN** a two-way or Mine/Base/Theirs snapshot has one duplicate occurrence
  between matching unique schema anchors and another corresponding occurrence
  at the matching `END` boundary, with complete validated rows, equal
  per-occurrence interval/run/ordinal, all-blank values/formulas, equal
  all-side full-column digests, a same unique ordered/count non-proof interval,
  and a final resolved/bijective column cache
- **THEN** the virtual view may render the exact logical fields and preserves
  their Mine/Base/Theirs physical operation mappings without treating the two
  occurrences as one shared interval

#### Scenario: Duplicate proof has unequal content or membership
- **WHEN** a duplicate run has nonblank or digest-unequal content, a missing or
  extra member, unequal run width/count/ordinal, or a missing or reordered
  unique/virtual boundary
- **THEN** the calculation surface remains non-actionable, the Sheet is
  `UNRESOLVED`, and no recycled virtual line, copy target, or operation target
  is exposed as an exact mapping

#### Scenario: Formula change inside a resolved proof interval
- **WHEN** an otherwise proven blank duplicate interval has unique, retained,
  bijective non-proof members whose formula identities differ across sides
- **THEN** the duplicate anchors may be used only for the stage-two rebuild,
  the formula cells remain visible exact differences, and an inherited formula
  mismatch cannot clear any unresolved, structural, reordered, or non-bijective
  interval member

#### Scenario: Exact three-way top cache has asymmetric child gaps
- **WHEN** duplicate proof triples are each represented by exactly one retained
  Mine/Base/Theirs slot in an otherwise non-ambiguous, non-structural,
  all-side-bijective top cache, while the Mine/Base and Theirs/Base child
  anchor intervals differ
- **THEN** the virtual view may use the existing exact top cache and preserve
  the proof triples; it SHALL NOT derive or replace any mapping from the child
  interval difference

#### Scenario: Row, Base, cache, or builder evidence fails
- **WHEN** declared row keys or row pairs are incomplete or ambiguous, the
  three-way Base differs, any final column or non-proof interval slot is
  unresolved, structural, reordered, or non-bijective, a proof column contains
  a formula token, or the duplicate-proof builder raises or returns incomplete
  evidence
- **THEN** the view SHALL retain the original unresolved reason(s), perform no
  legacy mapping fallback, and publish no exact/actionable virtual surface

### Requirement: Tk mutations and tags are batched
The system SHALL construct pane text and tag ranges before entering Tk and SHALL use bounded bulk widget operations for a viewport publication.

#### Scenario: Viewport contains many changed cells
- **WHEN** every row in the viewport has multiple changed cells
- **THEN** the system inserts pane text in bounded batches and applies row/cell tags in bulk rather than issuing per-cell workbook or render computations

#### Scenario: View mode changes
- **WHEN** the user switches between full and only-difference mode with exact snapshots available
- **THEN** the system replaces only the bounded viewport, preserves logical position where applicable, and does not rebuild a document proportional to the complete result count

#### Scenario: Hover repeats on the same logical row
- **WHEN** pointer motion repeatedly addresses the same logical row and column in the C-area or main panes
- **THEN** tooltip payload is computed at most once per target and the C-area does not delete, reinsert, or re-tag the complete row solely for that repeat hover

### Requirement: Wide result surfaces virtualize logical columns without data loss
The system SHALL retain the complete immutable logical column model while a
wide virtual surface materializes only the current visible logical columns plus
bounded formatting overscan.  It SHALL NOT reduce the default column width,
truncate the authoritative data model, or publish partially rendered columns as
an exact result.

#### Scenario: Wide changed sheet reaches exact readiness
- **WHEN** a wide Sheet such as 69-column `Dungeon` reaches a current-generation exact result
- **THEN** the calculation surface remains opaque until its complete first bounded row-and-column window, headers, tags, C-area geometry, scroll state, and physical/Base operation mappings are installed, after which the exact state is published once

#### Scenario: Horizontal thumb targets the complete logical width
- **WHEN** the user moves the horizontal scrollbar or minimap to its first, middle, or final position
- **THEN** the system coalesces to the requested logical-column window, renders only that bounded window, and preserves full-width logical coordinates for copy, selection, tooltip, and operations

#### Scenario: Combined row and column navigation
- **WHEN** high-rate wheel, vertical thumb, horizontal thumb, or minimap input changes both desired row and desired logical-column windows
- **THEN** only the newest bounded two-dimensional window is published, without worksheet access, formula normalization, record alignment, or intervening Text materialization

#### Scenario: Off-screen logical target is focused or operated
- **WHEN** focus, copy, hover, adoption, a column action, or a three-way Base action addresses an off-screen logical column
- **THEN** the view first maps that logical column into a bounded visible window and the action uses the complete physical/Base projection rather than a recycled Text position

#### Scenario: Wide-sheet interaction gate
- **WHEN** 20,000 changed rows or the changed-revision `Dungeon` fixture is exercised through first/middle/last horizontal positions and combined vertical/horizontal navigation
- **THEN** all view-only paths perform zero worksheet reads, viewport P95 is at most 33 ms, no heartbeat interval exceeds 200 ms, and the default rendered field values remain readable at their existing widths
