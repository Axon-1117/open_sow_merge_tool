# Validation and adversarial review

## Design-layer review

- The implementation treats `source pristine -> source working` as the only incoming change set; it never copies the complete source workbook into a target.
- Target-only values are preserved by the existing `_cross_branch_source_delta_premerge` implementation, and unsupported structure/package changes are blocked.
- Source commit is required before target processing. Every target is re-projected after its update, so a stale preview cannot be silently committed.
- Batch state is written atomically after source and target transitions. Completed targets remain complete, and a cancelled target can resume when its working file still matches the stored candidate hash.
- SVN credentials and source documents are not persisted outside the user-selected working copy and local batch metadata.

## Execution-layer evidence

- `python _smoke_test_branch_submit.py` — passed: whitelist/path safety, source-delta preview, blocked target, atomic batch state, UTF-16 no-BOM pathfile, UTF-8 log message, source drift rejection.
- `python _smoke_test_cross_branch_source_delta.py` — all 12 tests passed.
- `python _smoke_test_svn_merge_role_semantics.py` — all 7 tests passed.
- `python _smoke_test_svn_conflict_detection.py` — passed; live SVN case skipped because `svn.exe` is not installed on this machine.
- `python -m py_compile branch_submit.py sow_merge_tool.py _smoke_test_branch_submit.py` — passed.
- `dist/sow_merge_tool.exe --help` — passed and exposes `--branch-submit`.
- `build_exe.ps1 -SkipPublish` — passed; single-file EXE, release ZIP and SHA256 manifest generated.

## Adversarial review

- Source file changed after preview: blocked before the source commit dialog and persisted as `failed`.
- Target conflict in one selected branch: that branch is `blocked`; no commit dialog is opened for it, while ready branches remain visibly distinguishable and the batch cannot be started from the UI.
- Target commit cancellation: post-commit pristine/hash verification prevents false `committed`; remaining targets are not processed and the batch is resumable.
- Process restart after a target cancellation: a matching stored candidate is recognized and is not overwritten by a second update/replay.
- Existing TortoiseSVN diff/merge argument paths: role-semantics regression suite passed after the no-argument mode chooser was restricted to the default picker.

## Known validation boundary

The current machine has TortoiseSVN GUI (`TortoiseProc.exe`) but no `svn.exe`; real networked commit/cancel/update dialogs therefore remain an acceptance step on a disposable SVN test repository. The implementation fails closed if post-dialog pristine/hash verification cannot prove success.

