# 回归脚本

本目录保存 Excel 合并、SVN 提交和 Native 工作台回归脚本。它们不是用户运行时入口，统一由仓库根目录的 `tools/test.ps1` 调度。

常用命令：

```powershell
.\tools\test.ps1 -Profile Fast
.\tools\test.ps1 -Profile Full
.\tools\test.ps1 -Profile Integration
.\tools\test.ps1 -Profile Adversarial
.\tools\test.ps1 -Profile Native
```

`Fast` 只运行关键烟测；`Full` 才运行本目录全部 43 个烟测。真实 SVN 仓库测试位于 `tests/integration`，由 `Integration`、`Full` 和 `Adversarial` 调度；不要把它重新加入本目录造成重复执行。
