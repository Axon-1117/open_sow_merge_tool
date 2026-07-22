# sow_merge_tool 使用指南

## 0. 当前版本
- 版本号：`2026-07-22.update50`
- Build Tag：`new127-row-formula-structural-merge-fix`

### 本次修复
- 深入修复 2-way/3-way 的单元格覆盖、插行、批量插行和双方独立尾部新增；插入目标行会记录真实来源侧与来源行，撤销和最终保存按相同结构操作回放。
- 公式复制现在按 Excel 语义平移相对引用；相同公式且计算结果不同会保留公式、采用来源缓存值并提示同步依赖数据，公式和缓存完全相同则真正跳过，不再生成误导提示。
- 公式结构已纳入 3-way 冲突判断；不同公式即使当前缓存结果相同也不会被误判为相等，数组公式和数据表公式无法安全移动时会阻止覆盖。
- 2-way 整 Sheet 采用/删除也按 A/B 目标侧进入结构化原生回放，修复复杂 Sheet 只保留基础单元格、可能丢失高级对象的问题。
- Excel 原生插行回放由仅粘贴格式改为完整粘贴，保留来源行的批注、超链接、数据验证和其他单元格元数据，随后再以明确操作写入最终值和公式。
- Sheet 级自动删除现在同时核对公式、缓存值、样式、批注、超链接、合并区域、冻结窗格和行列尺寸；无法证明 mine 与 base 完全一致时保守提示冲突，不再静默删除本地格式修改。
- 带公式、宏、图表、数据验证、条件格式、合并单元格或外部对象的结构保存必须通过 Excel 原生回放；原生保存失败时停止替换目标文件，避免生成可打开但引用已损坏的工作簿。
- 大表尾部新增判定和公式缓存预检改为批量顺序读取，移除 read-only worksheet 高行号逐行随机回扫；点击采用期间后台差异计算会主动让路并拒绝旧缓存覆盖当前视图。
- `WorldMonster.xlsx` 真实三方副本回放：冲突扫描约 9.44 秒，目标 Sheet ready 约 14.27 秒，右侧区域采用 59 个差异单元格约 0.70 秒；等待 3 秒无回退，输出公式、缓存和 ZIP 校验通过。
- `Language.xlsx` 19,934 行真实副本回放：冲突扫描约 6.94 秒，目标 Sheet ready 约 20.12 秒，只看差异切换约 0.06 秒，无误报差异或冲突。
- 修复大表只看差异模式采用 theirs 后约 7 秒界面又显示 mine 旧值的问题；根因是区域操作失效快照后重新启动了读取磁盘旧 mine 的 only-diff 后台构建，并在完成后覆盖主视区行文本。
- 当前 Sheet 一旦发生用户操作，only-diff 后台任务会在启动和结果应用两处被阻止；后续刷新改为从已编辑的内存工作簿与当前差异映射重建文本，不再读取磁盘旧 mine。
- 修复大公式表点击“使用右侧区域”后先延迟生效、随后又恢复 mine 旧值的问题；日志确认旧版本会在打开后依次后台重算 mine/theirs/base，并把较早启动的 mine 结果回灌到用户操作后的界面。
- merge 模式不再自动启动 Excel 全表重算；自动重算策略现在严格遵守配置，存在未保存覆盖操作时也会阻止手动重算，避免磁盘旧文件覆盖内存合并结果。
- 后台 Sheet 差异缓存现在只要检测到该 Sheet 有任何用户操作就拒绝应用，不再要求初始 `_data_ready` 已完成，消除首屏后台结果延迟覆盖用户操作的竞态。
- 为已对齐的 `theirs -> mine` 区域增加批量写回路径和每 200 行进度反馈；真实 `WorldMonster@design` 完整 1200 行差异块约 1.5 秒，`WorldMonsterSurvivor@design` 完整 6000 行差异块约 2.0 秒。
- 区域采用不再受主视区 800 行渲染上限截断；只要属于同一逻辑连续差异块，屏幕外的尾部行也会一并采用并进入保存记录。
- 公式缓存保存增加严格 `cache-only` 模式，只更新 `<v>` 并原样保留现有 `<f>` 及共享公式组；真实区域采用后的 merged 输出已通过公式、缓存、ZIP 和共享公式结构校验。
- 修复 TortoiseSVN 手动 merge 大文件时偶发 `File is not a zip file` 的问题；此前只要异步导出文件已经出现就会立即复制，可能把仍在写入的 BASE 截成半个 ZIP。
- BASE 现在优先从 working copy 的 `.svn\\pristine` 同步读取；Tortoise 异步导出必须等待文件大小稳定，并通过完整 ZIP、`[Content_Types].xml`、`xl/workbook.xml` 和压缩成员校验后才允许进入比较。
- 所有 `.merge-left/.merge-right/.r####` sidecar 和 stable 临时副本增加统一完整性门禁，并避免 stable 临时文件被重复复制；不完整文件会给出明确的 SVN 临时文件错误，不再把底层 `BadZipFile` 直接暴露给用户。
- 使用 `C:\\GM15\\design\\sheets\\release` 当前 `WorldMonster.xlsx` 冲突现场回放通过：BASE/mine/theirs 三方完整性校验及 3-way 冲突扫描均成功。
- 修复 `WorldMonster.xlsx` 在“只看差异”下采用右侧区域后，相同公式的计算结果会被刷新流程恢复为 mine 旧缓存的问题；现在保留原公式、采用 theirs 当前缓存值，并提示同步合并公式依赖数据。
- 修复公式工作簿保存后 Excel 提示修复、修复后公式缓存结果丢失的问题；单元格覆盖优先使用经过结构校验的 OOXML 定点写回，保留共享公式元数据和未改动公式缓存，无法安全保存时会停止而不是生成风险文件。
- 公式与缓存结果采用独立操作记录，插入/删除行时会同步迁移记录；保存后仍保留公式，且不会为了采用缓存结果主动触发整本工作簿重算。
- 合并输出默认恢复为临时文件写入后原子替换，避免保存中断留下半写入文件。
- 修复 `Language.xlsx` 这类大表里点击“使用右侧区域”/“使用左侧区域”时明显卡顿的问题；区域模式下连续的单边新增行现在会合并成一次批量插入，不再逐行 `insert_rows()`，真实回放里 `theirs -> mine` 区域 adopt 从约 12.8 秒下降到约 3 秒。
- 修复 3-way 悬停完整对比区的来源标签口径，三方模式现在统一显示为 `base[行号] / mine[行号] / theirs[行号]`，不再显示 `A[...] / B[...]`。
- 修复 3-way 悬停/C 区等 Base 取值路径仍复用旧的 `A-side row` 口径的问题；涉及结构漂移或单边缺失行时，Base 现在统一按真实 `_base_row_for_pair()` 映射读取。
- 修复 `Language.xlsx` 这类 working copy 冲突文件在无法调用 `svn.exe` 或 Tortoise `cat BASE` 时，Base 会错误回退到目录里较小 `.r####` 文件的问题；现在会直接从 working copy 的 `.svn\\wc.db + pristine` 读取真正的 pristine BASE。
- 修复 `Language.xlsx / default@design@na_TLanguageCn` 这类“远端尾部大块新增 + 本地尾部独立新增”场景下，本地新增行被错误吸附到远端旧行的问题；尾部低相似度 paired 行现在会按保守规则拆成独立块。
- 修复拆分后的 paired twin rows 仍同时显示同一个 Base 行的问题；Base 现在只会保留在更接近 base 的那一侧，另一侧显示为空。
- 修复 3-way 下“悬停完整对比”面板在部分大表/首屏布局场景中被主视区挤压、导致高度异常的问题；现在会为底部对比区域预留稳定高度，确保 BASE / mine / theirs 三行能完整显示。
- 修复 3-way 尾部独立新增块被错误配对为同一行冲突的问题；对相对 base 的双方独立尾部追加，改为按 `theirs` 在前、`mine` 在后拆分成独立块，配合现有 `B2A` 插行语义可保留双方新增。
- 修复 10000+ 行大表时主视区/C 区行号栏宽度不足、只能显示 4 位数字的问题，行号宽度现在会随位数自动同步。
- 修复 `Language.xlsx` 这类 3-way 大表冲突在 merge 模式下首屏可能长时间无法打开的问题；大表 only-diff 首屏不再走 read-only worksheet 的高行号逐格随机访问慢路径。
- 为大表 only-diff 预计算增加块级行缓存，改为按 block 读取并在内存中比较 A/B/Base，避免 `refresh(rescan=True)` 卡到分钟级。
- 优化 3-way 大表的 Base 差异缓存与首屏渲染，只为当前候选/可见行补齐 base 口径，不再为整张大表做逐 pair 全量扫描。
- 修复 3-way 模式下 Base 列在结构漂移后可能读错 base 行的问题，`BASE2A` 现在会按真实 `mine -> base` 映射读取与覆盖。
- 修复 3-way 冲突扫描在插入/删除行后按原始行号误报冲突的问题，改为基于行映射比较。
- 修复 3-way 保存时可能丢失单边新增/删除整 sheet 的问题，并补齐整 sheet adopt/delete 保存链路。
- 缺失整 sheet 不再只显示摘要页签，改为空白对照表视图，并支持整表采用。
- 修复“只看差异”模式下新增整行，尤其是插入的空白行，未被正确筛选出来的问题。
- 修复差异 minimap 在只看差异模式下按过滤列表压缩定位、导致红块位置明显偏移的问题，改为按整张 sheet 的真实行位次定位。
- 修复主视区与 C 区的差异单元格高亮渲染，确保 diffcell 红底在主视区正确生效。
- 优化主视区 hover 驱动逻辑：无显式选中时，悬停任意单元格都会驱动 C 区和悬停对比区；有显式选中时，C 区保持锁定。
- 支持在主视区和 C 区右键取消显式单元格选中，恢复到 hover-driven 状态。
- 修复切换“只看差异”时旧的 `selcell` 蓝框残留问题，并保留仍然可见的差异行选中状态。
- 修复切换“只看差异”后旧 hover 状态可能继续指向被隐藏普通行、导致 C 区显示错误内容的问题。
- 修复首次打开 sheet 时后台 cache 应用漏同步列边界、导致主视区局部误出现灰色背景的问题；刷新本 Sheet 后恢复正常的问题也一并消除。
- 修复在“只看差异”模式下取消勾选后，主视区需要额外滚动一次鼠标才会恢复完整内容的问题；现在会立即重绘到正确的全量首屏。

## 1. 适用范围
- 用于 Excel `.xlsx` / `.xlsm` 文件的对比与冲突合并（SVN/TortoiseSVN 工作流）。
- `.xlsm` 会保留宏工作簿容器并按原扩展名保存。

## 2. 环境要求
- Windows 10/11
- 无需安装 Python 或其他运行库（已打包为单文件 EXE）

## 3. 直接运行
双击：
`sow_merge_tool.exe`

## 3.1 打包发布与 GitHub 提交
常用发布命令：

```powershell
powershell -ExecutionPolicy Bypass -File build_exe.ps1
```

如果本次发布需要在打包完成后顺手提交到 GitHub，可改用：

```powershell
powershell -ExecutionPolicy Bypass -File build_exe.ps1 -GitCommit -GitPush `
  -GitInclude sow_merge_tool.py,build_exe.ps1,release/README.md,release/sow_merge_tool.exe,release/sow_merge_tool_release.zip
```

可选参数：
- `-GitCommitMessage "release: 2026-07-22.update50 (new127-row-formula-structural-merge-fix)"`：自定义提交信息。
- `-GitInclude <path1,path2,...>`：必须显式指定要暂存的发布文件，避免把工作区中无关改动一并提交。
- `-GitRemote origin`：指定推送远端，默认 `origin`。
- `-GitBranch main`：指定推送分支；不传时默认使用当前检出的分支。

说明：
- Git 提交流程默认不会自动启用，避免误把临时改动直接推送。
- 开启后只会暂存 `-GitInclude` 明确列出的路径，然后提交并推送。
- 如果没有可提交的变更，脚本会跳过 commit，并保留已完成的打包/发布结果。

## 4. 命令行参数
### 4.1 普通对比
```bat
sow_merge_tool.exe --base "A.xlsx" --mine "B.xlsm"
```

### 4.2 SVN 冲突合并（推荐）
```bat
sow_merge_tool.exe --base "BASE.xlsx" --mine "MINE.xlsx" --theirs "THEIRS.xlsx" --merged "MERGED.xlsx"
```

### 4.3 单文件自动识别冲突
如果传入的是冲突文件路径（同目录包含 `.mine` / `.rXXXX`），工具会自动识别：
```bat
sow_merge_tool.exe "C:\path\conflict.xlsx"
```

## 5. 冲突合并流程
1. 进入冲突界面后，点击行号箭头或“采用对方(B)”/“保留我的(A)”进行覆盖。
2. 完成后点击“保存Merged并退出”。
3. 如果提示 Excel 占用，请关闭 Excel 再保存。

## 6. 常见问题
### 6.1 保存失败（Permission denied）
- 通常是 Excel 或 SVN 正在占用目标文件。
- 关闭 Excel 再保存即可。

### 6.2 打不开 / 无反应
- 确保 TortoiseSVN 已正确注册 diff/merge 工具（见下一节）。

## 7. TortoiseSVN 注册表配置
（管理员不需要，仅限当前用户）

### 7.1 Diff
```bat
reg add "HKCU\Software\TortoiseSVN\DiffTools" /v .xlsx /t REG_SZ /d "\"D:\\Tools\\sow_merge_tool\\dist\\sow_merge_tool.exe\" --base \"%base\" --mine \"%mine\" --title \"%bname\"" /f
reg add "HKCU\Software\TortoiseSVN\DiffTools" /v .xlsm /t REG_SZ /d "\"D:\\Tools\\sow_merge_tool\\dist\\sow_merge_tool.exe\" --base \"%base\" --mine \"%mine\" --title \"%bname\"" /f
```

### 7.2 Merge
```bat
reg add "HKCU\Software\TortoiseSVN\MergeTools" /v .xlsx /t REG_SZ /d "\"D:\\Tools\\sow_merge_tool\\dist\\sow_merge_tool.exe\" --base \"%base\" --mine \"%mine\" --theirs \"%theirs\" --merged \"%merged\" --title \"%bname\"" /f
reg add "HKCU\Software\TortoiseSVN\MergeTools" /v .xlsm /t REG_SZ /d "\"D:\\Tools\\sow_merge_tool\\dist\\sow_merge_tool.exe\" --base \"%base\" --mine \"%mine\" --theirs \"%theirs\" --merged \"%merged\" --title \"%bname\"" /f
```

### 7.3 备用（XLSX / XLSM 节点）
```bat
reg add "HKCU\Software\TortoiseSVN\DiffTools\XLSX" /v command /t REG_SZ /d "D:\\Tools\\sow_merge_tool\\dist\\sow_merge_tool.exe" /f
reg add "HKCU\Software\TortoiseSVN\DiffTools\XLSX" /v args /t REG_SZ /d "--base %base --mine %mine --title %bname" /f
reg add "HKCU\Software\TortoiseSVN\MergeTools\XLSX" /v command /t REG_SZ /d "D:\\Tools\\sow_merge_tool\\dist\\sow_merge_tool.exe" /f
reg add "HKCU\Software\TortoiseSVN\MergeTools\XLSX" /v args /t REG_SZ /d "--base %base --mine %mine --theirs %theirs --merged %merged --title %bname" /f
reg add "HKCU\Software\TortoiseSVN\DiffTools\XLSM" /v command /t REG_SZ /d "D:\\Tools\\sow_merge_tool\\dist\\sow_merge_tool.exe" /f
reg add "HKCU\Software\TortoiseSVN\DiffTools\XLSM" /v args /t REG_SZ /d "--base %base --mine %mine --title %bname" /f
reg add "HKCU\Software\TortoiseSVN\MergeTools\XLSM" /v command /t REG_SZ /d "D:\\Tools\\sow_merge_tool\\dist\\sow_merge_tool.exe" /f
reg add "HKCU\Software\TortoiseSVN\MergeTools\XLSM" /v args /t REG_SZ /d "--base %base --mine %mine --theirs %theirs --merged %merged --title %bname" /f
```

## 8. 日志
日志路径：
```
%TEMP%\sow_merge_tool_debug.log
```

---

如需更新版本，请替换 `sow_merge_tool.exe` 后重新注册即可。
