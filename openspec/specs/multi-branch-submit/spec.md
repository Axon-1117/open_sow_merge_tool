# Multi-branch submit

## Scenario: select branches and files

- **WHEN** the user opens branch submit mode
- **THEN** the tool discovers the SVN working-copy root, dynamically shows all same-repository top-level branches including `master` in descending SVN change-time order, and requires one source, one or more targets, and one or more safe `.xlsx` changes

## Scenario: recursive multi-file workbench

- **WHEN** the user starts from an `.xlsx`, folder, or folder background
- **THEN** the tool recursively scans that folder, shows all SVN changes with status metadata, defaults safe modified/added/deleted Excel files on, keeps unversioned Excel off, and displays non-Excel changes read-only

## Scenario: optional source delta preflight

- **WHEN** the user chooses “预检查（可选）” and source and all selected targets are clean enough to analyze
- **THEN** the tool obtains source pristine and working content, computes source-before to source-after changes, and previews a per-target action matrix without changing target files

## Scenario: direct submission without preflight

- **WHEN** the user clicks “开始提交” without opening the optional preflight dialog
- **THEN** the same checks run in the background immediately before the source commit; safe unique-key tail-row changes are prepared automatically, already-applied files are skipped, and ambiguous changes open the Excel merge tool instead of being silently copied

## Scenario: fail closed

- **WHEN** status collection fails, a selected target is dirty, conflict/switch/property/external/unsupported structure is detected, or a target changed since analysis
- **THEN** the affected action is blocked and no commit dialog is opened for that unsafe action; a manual action remains available through the Excel merge tool

## Scenario: conservative fast path

- **WHEN** a source delta is limited to value/formula changes or unique-key rows appended at the source tail, with stable headers and default styles
- **THEN** the tool classifies the target in read-only OOXML, projects only the proven cells/rows into a candidate in place of a full openpyxl rewrite, and reports the automatic/already/manual counts

## Scenario: manual merge fallback

- **WHEN** a target cell has an independent value, a row is inserted in the middle, styles/structure changed, or the unique key cannot be proved stable
- **THEN** the tool marks the file “需人工合并”, launches the existing Excel merge UI with source-before/source-after aliases, and accepts the result only after the target hash changes and SVN status is rechecked

## Scenario: added, deleted, and renamed files

- **WHEN** an added or deleted Excel file is selected
- **THEN** add requires an absent target path and versioned parent, delete requires target content equal to the deletion baseline, and a detected rename is blocked with a Repair Move instruction

## Scenario: source commit first

- **WHEN** the user starts a ready batch
- **THEN** the source branch commit dialog opens first with the frozen user message, and targets remain pending until every file is reconciled against SVN status, pristine, hashes, and revisions after the dialog closes

## Scenario: partial source selection

- **WHEN** the user commits only some selected source files
- **THEN** propagation stops, the committed files are split into an explicit child batch, and the remaining source files return to the workbench

## Scenario: per-target confirmation

- **WHEN** a target is ready after revalidation
- **THEN** the tool writes a durable prepare intent and backup, applies only the planned candidate, opens one TortoiseSVN commit dialog with the frozen message and tracking footer, and marks each file committed only after post-commit reconciliation

## Scenario: cancellation and recovery

- **WHEN** any source or target commit is cancelled or fails
- **THEN** the tool stops subsequent targets, persists per-file facts, keeps completed files immutable, and on resume reconciles actual SVN state before opening another dialog or restoring only hash-matching uncommitted candidates

## Scenario: legacy compatibility

- **WHEN** the tool is launched with existing two-way, three-way, or TortoiseSVN diff/merge arguments
- **THEN** those paths retain their current behavior and do not enter branch submit mode

## Scenario: Explorer context menu

- **WHEN** the current user invokes “多分支 SVN 提交” from an `.xlsx`, folder, or folder background
- **THEN** the tool opens branch-submit mode, infers the working-copy/source/scope, and recursively scans the selected folder without changing legacy TortoiseSVN registrations

## Scenario: context-menu uninstall

- **WHEN** the current user runs the dedicated context-menu uninstaller
- **THEN** only the three owned `SowMultiBranchSVNSubmit` shell trees are removed and other Explorer or TortoiseSVN settings remain untouched
