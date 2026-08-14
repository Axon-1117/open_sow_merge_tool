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

Development happens on `SWolf`. After the local gate and post-build
adversarial review pass, `SWolf` is merged directly into `master`; this
repository does not use pull requests for the release flow.

The Feishu requirements document and real business workbooks remain
read-only during validation.
