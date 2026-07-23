# sow_merge_tool 使用指南

## 0. 当前版本
- 版本号：`2026-07-23.update56`
- Build Tag：`new133-region-mode-guided-apply`

### 本次修复
- 区域模式不再要求用户预先点中差异行：没有有效选区或当前行无差异时，首次点击“使用左侧区域/使用右侧区域”会自动定位、高亮并滚动到最近的可应用差异块，状态栏明确提示再次点击；第二次点击才写入，避免隐式修改未确认区域。
- 已明确选中且方向可用的差异区域仍保持单击采用；显式选中方向不可用的差异块时会原地提示，不跨块写入。整张 Sheet 无可用区域时仅显示非模态状态，不弹窗、不响铃、不产生撤销记录。
- Full 模式只在当前已渲染范围内寻找候选，避免破坏大表增量渲染；Only-diff 使用完整差异块模型，可从已处理快照块定位到下一待处理块。区域定位只读取内存差异映射，不读取 worksheet、不触发重算或映射重建。
- 区域候选按 A2B/B2A/BASE2A 分别校验来源行、方向差异列和不可写列映射；20 万行空方向差异图增加 O(1) 短路，合成基准由约 112–153ms 降至约 0.005ms。
- 完成逻辑列结构对齐：Mine/Base/Theirs 共用不可变逻辑列槽，支持插入、删除、复制、保留、原子回滚与精确撤销；公式引用、样式、批注、验证、条件格式、合并区域和跨 Sheet 引用通过 Excel 原生整列回放并在保存后重新打开校验。
- “只看差异”升级为精确差异块视图：按完整 only-diff 快照和逻辑 pair 连续性统计差异块，工具栏显示当前块、总块数、待处理数量及已处理状态，不再把原表中相隔很远的区域显示成无法区分的一段。
- 每个差异块在主视区和行号区使用同步间距，左侧行号栏增加 `[块号]` 文本标记；2-way/3-way 的全部可见窗格保持行对齐，关闭“只看差异”后这些块专属标记会立即隐藏。
- 上一处/下一处差异改为使用完整快照导航，能跳到大表首屏 800 条渲染上限之外的差异块；目标块按缓存完整物化，不触发 worksheet 重扫，并保持主视区与 C 区的横向滚动位置。
- 区域采用与显示、计数、导航统一使用同一差异块边界；块内已经处理的行会跳过，相邻块不会被误改。修复工具栏区域操作受旧悬停状态影响、可能忽略用户当前显式行/单元格选择的问题。
- 差异块元数据只读取内存中的 `_full_display_rows` 和差异列缓存，滚动、选择及块状态更新不会增加高行号 worksheet 随机读取；块编号在已处理行继续保留时保持稳定。
- 真实 `WorldMonster.xlsx` 三方副本回放：精确 only-diff 得到 1201 条差异行、2 个逻辑块；跨块导航未触发 `rescan=True`，1200 行区域采用记录 8955 个公式缓存单元格，相邻块和 mine 本地修改均保持不变。
- 修复 Excel COM 原生结构回放写入空白单元格时把 JSON `null` 直接赋给 `Value2` 可能失败的问题；空白操作现在显式调用 `ClearContents()`，并继续区分空字符串、数字零和布尔值。
- 修复 read-only worksheet 裁剪边界可能只按尾行估算列数、遗漏前部较宽列的问题；后台 Sheet 差异标记复用顺序读取的 trim 行缓存，减少重复 XML 扫描。
- benchmark 现在强制导入脚本同目录的当前源码并校验模块路径，避免误测旧安装目录中的历史版本。
- 启动 merge 模式时的“检测到冲突/未检测到直接冲突”提示改为主窗口内非阻塞通知，不再使用带 `grab/wait_window` 的确认弹窗；提示会自动消失，也可手动关闭，工作簿加载与用户操作可直接继续。
- 修复 `Dungeon.xlsx` 大表在区域模式下首次点击“使用右侧区域”没有效果、必须先执行一次单行采用后区域操作才生效的问题；区域锚点现在与单行模式一致，优先使用已选中或悬停的 pair，不再误读 800 行首屏后的第 801 行 Tk 尾部哨兵。
- 使用真实 `Dungeon.xlsx` 三方文件副本回放：在 `display_rows=800 / insert_line=801` 的故障条件下，首次区域操作正确识别 `pair 152-1351` 共 1200 行，样本值由 mine 的 `411` 更新为 theirs 的 `167`，并完整记录普通单元格及公式缓存操作。
- 修复启动进度交互优化后主窗口无法全屏/最大化的问题；启动进度窗切换为主界面后会恢复窗口缩放能力，并在 Windows 下自动最大化。
- 修复主视区加载和差异计算明显变慢的问题；后台计算结果会保留给尚未打开的 Sheet 复用，小/中型 Sheet 改为顺序批量读取行缓存，避免 read-only worksheet 逐格随机访问和首次切换页签的重复计算。
- 3-way 的 `mine -> base` 精确差异列一并进入后台 Sheet 缓存，主视区纯渲染不再等待可编辑工作簿加载；Base 行文本在一次重绘内批量读取并复用，不再为差异着色重复回扫 XML。
- “只看差异”切换改为非阻塞状态机：数据未 ready 时只记录用户选择并提升当前 Sheet 优先级，不再在 Tk 回调内执行同步全表 rescan；当前 Sheet 精确差异计算期间会暂停未打开 Sheet 的低优先级扫描。
- 下方 Sheet 页签升级为两阶段着色：ZIP Sheet 指纹预检在不解压工作表 XML 的情况下快速标记浅黄色候选，精确单元格/公式比较后确认亮黄色或清除候选；迟到的预检结果不会覆盖精确结论。
- Sheet 导航新增“浅黄=预检，亮黄=确认”提示。真实 `WorldMonster.xlsx` 回放中，候选 Sheet 在窗口出现后约 0.66 秒完成预标记，第二个差异 Sheet 无需等待完整精确扫描后才出现差异提示。
- 真实性能回放：`WorldMonster.xlsx` 首个 Sheet ready 约 6.89 秒，取消/勾选“只看差异”回调约 0.130/0.016 秒，精确 only-diff 约 6.62 秒；`Language.xlsx` 20,340 行首屏约 16.02 秒，勾选回调约 0.021 秒，精确 only-diff 约 8.30 秒。
- 初始化项目级 OpenSpec 目录与 Codex OpenSpec skills，后续规格变更可在仓库内直接使用 propose/apply/archive 工作流。
- 新增完整的耗时任务进度反馈：启动时显示 SVN 来源解析、三方冲突扫描和 mine/base/theirs 工作簿加载阶段，主界面持续显示当前后台计算的 Sheet 与总体进度。
- merged 保存现在从等待可编辑数据、重放整 Sheet/插行/公式缓存操作、写入目标文件到 OOXML/ZIP 完整性校验全程显示阶段、进度和已用时间，不再只在最后复制文件时短暂显示静态进度窗。
- 启动加载和 merged 保存改为后台执行，Tk 事件循环保持响应；保存期间会暂停低优先级 Sheet 扫描，优先完成用户操作并阻止重复点击。
- 启动进度窗与主应用复用同一个 Tk 解释器，修复 Python 3.14 下临时窗口跨线程析构可能导致的 `Tcl_AsyncDelete`，并保持 F4 等根窗口快捷键正常工作。
- `Language.xlsx` 真实三方副本回放：冲突扫描约 7.53 秒，首个大 Sheet ready 约 16.84 秒，保存并校验约 1.71 秒；merged 输出完整可读。
- `WorldMonster.xlsx` 真实三方副本回放：冲突扫描约 9.46 秒，首个 Sheet ready 约 8.06 秒，保存并校验约 1.94 秒；merged 输出完整可读。
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
- `-GitCommitMessage "release: 2026-07-22.update52 (new129-fast-ui-and-sheet-premark-fix)"`：自定义提交信息。
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
