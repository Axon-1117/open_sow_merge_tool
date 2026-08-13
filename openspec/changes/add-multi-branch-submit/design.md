# Design

## First-principles invariants

1. 用户真正要同步的是源分支的 delta，而不是源工作簿的最终字节。
2. 目标分支已有的、与源 delta 无关的修改必须保留。
3. SVN 远端在分析后发生变化时，旧候选失效，必须重新分析。
4. 提交成功只能由提交后状态验证确认，TortoiseProc 进程退出不是成功信号。
5. SVN 跨分支提交不是原子事务；批次必须保存每个分支、每个文件的事实状态，并在部分成功时停止后续传播。
6. 未解决冲突、不支持结构变化、工作副本异常都 fail closed。
7. 工具写入目标工作副本前必须先落盘意图、备份和候选哈希；恢复只处理仍能证明属于本批次的现场。

## Data flow

`Source pristine -> Source working` 形成 incoming delta；对每个 target 执行 `target HEAD + incoming delta` 预合并；所有目标通过后先打开 source commit，再按用户顺序更新、重新验证、写入候选、打开 target commit，并逐文件核对 working/pristine/revision。新增要求目标不存在且父目录已版本化；删除要求目标语义仍等于源删除前基线；重命名 fail closed 并提示 Repair Move。

批次 JSON 位于 `%LOCALAPPDATA%/SowMergeTool/branch_submit/batches/<id>/batch.json`，只保存路径、哈希、版本、状态和可恢复元数据，不保存密码。

源提交只完成部分文件时，原批次标记为 superseded，并生成仅包含已提交文件的子批次；用户明确载入子批次后才能继续目标传播。目标分支部分提交时停止后续分支，续跑会先对账，跳过已经提交的文件。

## SVN integration

优先使用 `svn status --xml --verbose --depth infinity --ignore-externals`。没有 `svn.exe` 时，同一安装包以隐藏子进程加载 TortoiseSVN 的 `libsvn_tsvn.dll` 执行 `svn_client_status6`，通过 JSON 返回结果；子进程崩溃、超时或 JSON 无效都会阻断整批。`.svn/wc.db` 只用于已由 SVN 状态证明后的仓库身份、节点和 pristine 对账，不用文件时间戳猜测状态。

分支候选必须同时满足顶层工作副本节点、同 repository root、同 UUID、合法目录且非重解析链接；工具目录与非分支目录不会仅凭文件夹名进入候选。状态扫描忽略 externals，冲突、obstructed、replaced、switched、属性修改和未知状态 fail closed。

提交说明在预检查时冻结；目标说明为冻结说明加追踪尾注：`[MultiBranchSync] batch=<id> source=<branch>@r<revision>`。源提交和每个目标提交都由用户在确认窗口中最终确认。
