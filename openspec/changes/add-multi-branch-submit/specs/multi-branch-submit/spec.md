# Multi-branch submit

## Scenario: select branches and files

- **WHEN** the user opens branch submit mode
- **THEN** the tool discovers the configured SVN working-copy root, shows only whitelisted branches (`develop`, `release`, `sandbox` by default), excludes `master`, and requires one source, one or more targets, and one or more `.xlsx` files

## Scenario: source delta preflight

- **WHEN** source and all selected targets are clean enough to analyze
- **THEN** the tool obtains source pristine and working content, computes source-before to source-after changes, and previews per-target applied, already-present, retained, unresolved, and unsupported counts without changing target files

## Scenario: fail closed

- **WHEN** a selected target has SVN conflict, an unselected local change would be included, an unsupported structural change is detected, or a target changed since analysis
- **THEN** the batch is blocked and no commit dialog is opened until the condition is resolved and analysis is rerun

## Scenario: source commit first

- **WHEN** the user starts a ready batch
- **THEN** the source branch commit dialog opens first with the user message prefilled, and targets remain pending until source revision and file hashes are verified after the dialog closes

## Scenario: per-target confirmation

- **WHEN** a target is ready after revalidation
- **THEN** the tool applies only the planned candidate, opens one TortoiseSVN commit dialog for that target with the source message and tracking footer, and marks the target committed only after post-commit verification

## Scenario: cancellation and recovery

- **WHEN** any source or target commit is cancelled or fails
- **THEN** the tool stops subsequent targets, persists the batch as resumable, keeps completed target states, and on resume rechecks revisions/hashes before continuing

## Scenario: legacy compatibility

- **WHEN** the tool is launched with existing two-way, three-way, or TortoiseSVN diff/merge arguments
- **THEN** those paths retain their current behavior and do not enter branch submit mode

## Scenario: Explorer context menu

- **WHEN** the current user installs the `.xlsx` context menu and invokes “多分支 SVN 提交” on a workbook
- **THEN** the tool opens branch-submit mode, infers the SVN working-copy root and source branch from that file, and preselects it without changing legacy TortoiseSVN registrations

## Scenario: context-menu uninstall

- **WHEN** the current user runs the dedicated context-menu uninstaller
- **THEN** only the `SowMultiBranchSVNSubmit` shell key is removed and other Explorer or TortoiseSVN settings remain untouched
