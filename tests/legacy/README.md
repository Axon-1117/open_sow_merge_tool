# 历史回归脚本

本目录保存历史 Excel 合并、SVN 提交和 GUI 回归脚本。它们不是用户运行时入口，统一由仓库根目录的 `tools/test.ps1` 调度。

常用命令：

```powershell
.\tools\test.ps1 -Profile Fast
.\tools\test.ps1 -Profile Full
.\tools\test.ps1 -Profile Native
```
