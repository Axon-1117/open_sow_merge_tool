## 1. SVN launch context and identities

- [x] 1.1 Add explicit merge-scenario, version-identity, and launch-context data models with centralized workspace colors
- [x] 1.2 Classify two-way, update-conflict, cross-branch-merge, and unknown-three-way launches from raw arguments
- [x] 1.3 Stop WC BASE from replacing the original Base and remove the unsafe cross-branch revision re-export path
- [x] 1.4 Retain target WC pristine as a separate fourth input identity and pass the complete context into the UI
- [x] 1.5 Resolve and format SVN revision/author metadata with local-first fallbacks and logged unknown states

## 2. Equivalence, convergence, and automatic merge

- [x] 2.1 Implement complete OOXML package comparison that covers `.xlsx` and `.xlsm` members while ignoring ZIP container metadata
- [x] 2.2 Build and log the complete pairwise equivalence matrix with hashes, readiness, elapsed time, and reasons
- [x] 2.3 Implement safe whole-workbook convergence rules and initialize the merged candidate from the proven result
- [x] 2.4 Integrate Base-anchored semantic pre-merge for supported one-sided/identical changes while retaining unresolved and unsupported differences
- [x] 2.5 Produce one startup outcome model containing automatic action, folded identity, merged counts, unresolved counts, and fallback reasons

## 3. Adaptive merge workspace

- [x] 3.1 Apply gray, soft-pink, and soft-green colors to chrome-only containers for two-way, update, and branch-merge modes
- [x] 3.2 Display scenario-aware Base/Mine/Theirs/target-pristine labels with revision and author information
- [x] 3.3 Implement redundant-pane folding and an always-available `展开三方`/fold toggle without discarding models
- [x] 3.4 Preserve sheet, scroll, selection, pending choices, and initialized result across fold transitions
- [x] 3.5 Extend the startup read-only gate through role analysis, author lookup, equivalence comparison, and automatic processing
- [x] 3.6 Show the required automatic-outcome dialog and route its primary action to first conflict or Save Merged

## 4. Verification

- [x] 4.1 Add role-classification and production launch-path regressions proving raw update Base and branch merge-left are preserved
- [x] 4.2 Add package-equivalence matrix tests including ZIP metadata differences, workbook-part differences, errors, and `.xlsm` VBA differences
- [x] 4.3 Add automatic convergence and non-overlapping/overlapping semantic merge tests
- [x] 4.4 Add GUI tests for mode colors, white spreadsheet canvases, folded/expanded layouts, and state preservation
- [x] 4.5 Add author-label and diagnostic-log tests for resolved and unavailable metadata
- [x] 4.6 Run real-workbook `.xlsx/.xlsm` acceptance cases and capture startup/equivalence performance baselines
- [x] 4.7 Run the complete smoke, unit, GUI, compile, and OpenSpec validation suites

## 5. Source-delta-driven branch merge refinement

- [x] 5.1 Read exact cross-branch source path/revision roles from WC conflict metadata with sidecar-name fallback
- [x] 5.2 Add scenario-specific Source Before/Source After/Target Working/Target Pristine labels and diagnostics
- [x] 5.3 Project only the Source Before-to-Source After logical delta onto a Target Working candidate, preserving unrelated target differences
- [x] 5.4 Separate incoming, applied, already-present, target-retained, and unresolved outcome counters and update the startup dialog/log

## 6. Refinement verification

- [x] 6.1 Add synthetic cell/row/column tests for applied, already-present, third-value conflict, unrelated target changes, and ambiguous structure
- [x] 6.2 Replay the real Building r37073/r37074 conflict expecting incoming=1, applied=0, already-present=1, unresolved=0; then run affected and complete regressions

## 7. Structural-conflict navigation and compact workspace

- [x] 7.1 Replay the real Gunships r37347/r37348 conflict and record source roles, target/source column geometry, structural blocks, and navigable conflicts
- [x] 7.2 Auto-select the first actionable structural column block and keep explicit Target Working/Source Before/Source After column choices available after readiness checks
- [x] 7.3 Exclude workbook-level pseudo locations from cell navigation, skip no-difference sheets, and provide a manual-review action when no real cell is navigable
- [x] 7.4 Replace long inline identity paths with compact role, filename/revision, and Author labels while retaining full paths in hover details and diagnostics
- [x] 7.5 Remove the duplicate upper Sheet tab strip and redundant top diagnostics, compact the lower comparison areas, and allocate released vertical space to the main grid

## 8. Structural-conflict and layout verification

- [x] 8.1 Add focused GUI regressions for automatic structural-block selection, enabled column choices, pseudo-location filtering, and real-conflict navigation
- [x] 8.2 Add GUI geometry regressions for one visible Sheet navigator, legible Author labels, a single-line hover heading, and increased main-grid height
- [x] 8.3 Replay the real Gunships r37347/r37348 merge and run affected smoke, logical-column, adaptive-workspace, compile, and OpenSpec validation suites

## 9. Action-first structural workflow and stable authors

- [x] 9.1 Reflow the structural toolbar so all three column buttons remain fully visible and verbose structural summaries move to compact secondary detail
- [x] 9.2 Automatically select the next actionable structural block after each successful decision and distinguish remaining work from completion
- [x] 9.3 Resolve and cache exact repository-revision authors independently of the current WC node revision
- [x] 9.4 Add narrow-window, multi-column progression, advanced-WC author, and real Gunships r37347/r37348 regressions

## 10. Centered actions, global adoption, and unified column geometry

- [x] 10.1 Center the complete difference-navigation group and move tagged structural status directly before the three column buttons with red logical-column text
- [x] 10.2 Add left/right Global Mode that atomically applies every safely mapped cell difference on the current Sheet as one undoable action
- [x] 10.3 Build and reuse one bounded uniform column-width model per Sheet across every logical column, main row, pane, and column header
- [x] 10.4 Replace C-area internal L-number headers with Excel-style labels matching the main view
- [x] 10.5 Add layout, global-action atomicity/undo, stable-width, Excel-label, real-workbook, performance, and regression verification

## 11. CJK pixel geometry and compact action layout

- [x] 11.1 Replace code-point padding with fixed-index East Asian display-width formatting and use one CJK monospace font across main panes, C-area, and headers
- [x] 11.2 Prevent selection/difference tags from changing font metrics and prove GunshipsModify A1:A5 pixel boundaries match their headers and all three panes
- [x] 11.3 Combine difference navigation and structural-column controls into one collision-safe responsive action row
- [x] 11.4 Move the four root utility buttons to the far left and replace every user-facing structural-column `L<n>` label with Excel-style names
- [x] 11.5 Run real Gunships, mixed-language geometry, narrow/wide layout, column-action, hover, navigation, compile, and strict OpenSpec regressions
