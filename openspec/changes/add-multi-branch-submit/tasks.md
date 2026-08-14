# Tasks

- [x] Add branch-submit domain model, settings, batch persistence, and safe path validation.
- [x] Add fail-closed recursive SVN status provider using `svn --xml` or an isolated TortoiseSVN DLL child process.
- [x] Add dynamic same-repository branch discovery, recursive change scanning, source-delta preflight, add/delete gates, and per-file action matrix.
- [x] Add Tortoise-style Tk workbench with unlimited branch checkboxes, favorites/search, multi-file filters, recent messages, progress/cancel, diff and path actions.
- [x] Add write-ahead batch journal, candidate/backup hashes, source partial child batches, target partial stop, restart reconciliation, guarded restore, and corrupt-state reporting.
- [x] Add `--branch-submit` and no-argument mode chooser without changing existing TortoiseSVN argument paths.
- [x] Add adversarial tests for dirty targets, rename blocks, partial selection, server-success/client-error, process death after source/target commit, write-intent crash, guarded restore, and corrupt state.
- [x] Run first-principles design review, implementation review, and adversarial interruption/concurrency review.
- [x] Build and deploy portable release EXE/ZIP packages, update documentation/version, verify installed hashes, and keep release actions independent of PR workflow.
- [x] Add installable/uninstallable `.xlsx`, folder, and folder-background Explorer context menus with exact three-key uninstall scope.
- [x] Enable `master`, sort branch candidates by newest SVN change time, and move the static Explorer verb from forced top placement to the natural slot before SVN extension commands.
