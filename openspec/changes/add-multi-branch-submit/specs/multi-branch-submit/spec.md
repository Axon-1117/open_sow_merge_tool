# Multi-branch submit

## Scenario: select branches and files

- **WHEN** the user opens branch submit mode
- **THEN** the tool discovers the SVN working-copy root and dynamically shows same-repository top-level branches, keeps `master` visible but disabled, and requires one source, one or more targets, and one or more safe `.xlsx` changes

## Scenario: recursive multi-file workbench

- **WHEN** the user starts from an `.xlsx`, folder, or folder background
- **THEN** the tool recursively scans that folder, shows all SVN changes with status metadata, defaults safe modified/added/deleted Excel files on, keeps unversioned Excel off, and displays non-Excel changes read-only

## Scenario: source delta preflight

- **WHEN** source and all selected targets are clean enough to analyze
- **THEN** the tool obtains source pristine and working content, computes source-before to source-after changes, and previews per-target applied, already-present, retained, unresolved, and unsupported counts without changing target files

## Scenario: fail closed

- **WHEN** status collection fails, a selected target is dirty, conflict/switch/property/external/unsupported structure is detected, or a target changed since analysis
- **THEN** the batch is blocked and no commit dialog is opened until the condition is resolved and analysis is rerun

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
