# SOW Excel Merge Tool

Windows desktop tool for semantic Excel merge and recoverable multi-branch
SVN submission. The user-facing executable remains `sow_merge_tool.exe`.

## Local workflow

```powershell
.\tools\bootstrap.ps1
.\tools\test.ps1 -Profile Fast
.\tools\test.ps1 -Profile Full
.\tools\release.ps1 -DeployPath 'C:\sow_main\excel\excel_merge_tool'
```

Development happens directly on `master`. The repository does not use pull
requests for the release flow.

Regression and GUI scripts live under `tests/legacy`; they are not runtime
entry points and are invoked through `tools/test.ps1`.

The Feishu requirements document and real business workbooks remain
read-only during validation.
