# Validation and adversarial review

## Design-layer review

- Incoming work is always `source pristine -> source working`; a complete source workbook is never copied over a target branch.
- Cross-branch submission is modeled as a recoverable sequence, not an atomic transaction. `BranchSubmitBatch` stores per-file source and target facts and stops after any partial or unknown result.
- Status is fail-closed. The normal provider is `svn status --xml`; this workstation used the isolated TortoiseSVN DLL child, whose crash, timeout, callback error, or invalid JSON blocks the batch.
- Every target must be clean before tool intervention. Modify, add, and delete have separate compatibility gates; rename and unsupported workbook structures are blocked.
- A prepare intent is durably saved before each target write. Restore requires the current file to match the recorded candidate hash, and committed repository revisions are never rolled back automatically.
- TortoiseProc exit codes are advisory only. Source and target reconciliation uses actual recursive SVN status, pristine content, working hashes, and WC node revision data.

## Execution-layer evidence (2026-08-14)

- `python _smoke_test_branch_submit.py` — 12 adversarial tests passed, including partial source selection, server success with client error, source/target process death, write-intent crash, restore guard, corrupt state, folder handoff, and real TortoiseSVN DLL status.
- All 42 repository `_smoke_test*.py` scripts passed (0 failures). Existing Excel merge, formula/cache, structural replay, SVN role, conflict, `.xlsm`, and cross-branch source-delta behavior remained green.
- `python _gui_self_test_branch_submit_workbench.py` — passed with 32 branch candidates and 200 changed items.
- Real read-only WC discovery at `C:\sow_main\excel` found 22 same-repository branches, enabled `master` as source and target, and sorted them by the newest SVN `changed_date` (`develop`, `release`, `master`, ...). Repository UUID remained `509b88cb-e3bb-49fc-85e7-49e888d66b00`.
- The built and installed EXE's hidden status child returned the two current unversioned Excel lock files under `release`; the installed GUI opened with title `Excel 合并器 · 多分支 SVN 提交` and was cancelled without a business commit.
- `build_exe.ps1 -SkipPublish` generated the `2026-08-14.update76` single EXE and ZIP. The installed EXE SHA256 is `ABB0C6FCDC9338CE1DD34A0921753DDCB9BEDCBB255BE77983E1FAE66DB7B09A`; all 12 deployed files match the release directory.
- Context-menu round trip installed exactly three owned keys and then removed exactly those keys. An adjacent sentinel key survived. Final commands point to `C:\sow_main\excel\excel_merge_tool\sow_merge_tool.exe` with `%1`, `%1`, and `%V` respectively. Reinstall removed the legacy `Position=Top` value from all three keys so Windows uses normal static-verb enumeration before the SVN extension group.

## Repeated attack cases

- Service succeeded but client reported failure/process died: resume reconciles committed pristine/hash/revision and does not reopen the commit dialog.
- User unchecked part of the source set: propagation stops and creates an explicitly loadable committed-only child batch.
- Tool died after durable prepare intent and working-copy mutation: startup audits the intent, recognizes only an exact candidate/deletion scene, and otherwise marks the action `unknown`.
- Target committed only some files: later targets remain untouched; resume skips committed files and revalidates only pending candidates.
- Target content changed after preview: update and fresh projection rerun; dirty or incompatible content blocks the whole batch before source propagation.

## Remaining acceptance boundary

This workstation has TortoiseSVN 1.14.9 runtime libraries but no `svn.exe`/`svnadmin.exe`. The isolated native status path was exercised against the real business WC, while destructive commit/cancel tests used disposable WC fixtures and fake commit reconciliation. No business workbook was submitted. A final end-user acceptance pass may use a disposable network SVN repository to observe the native TortoiseSVN dialogs end to end.
