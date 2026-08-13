# Design

## First-principles invariants

1. 用户真正要同步的是源分支的 delta，而不是源工作簿的最终字节。
2. 目标分支已有的、与源 delta 无关的修改必须保留。
3. SVN 远端在分析后发生变化时，旧候选失效，必须重新分析。
4. 提交成功只能由提交后状态验证确认，TortoiseProc 进程退出不是成功信号。
5. 批次在每个持久化节点可恢复；取消/失败停止后续分支，避免不受控的部分提交。
6. 未解决冲突、不支持结构变化、工作副本异常都 fail closed。

## Data flow

`Source pristine -> Source working` 形成 incoming delta；对每个 target 执行 `target working + incoming delta` 预合并；所有目标通过后先打开 source commit，再按用户顺序更新、重新验证、写入候选、打开 target commit，并验证工作副本干净和 revision 前进。

批次 JSON 位于 `%LOCALAPPDATA%/SowMergeTool/branch_submit/batches/<id>/batch.json`，只保存路径、哈希、版本、状态和可恢复元数据，不保存密码。

## SVN integration

优先使用 `svn.exe` 获取机器可读 status/info；当前常见安装只有 TortoiseProc 时，通过只读 `.svn/wc.db`/pristine 检查和 TortoiseSVN automation 完成 update/commit。`/pathfile` 临时文件使用 UTF-16 无 BOM；`/logmsgfile` 预填提交说明。

提交说明为源说明加追踪尾注：`[MultiBranchSync] batch=<id> source=<branch>@r<revision>`。源提交和每个目标提交都由用户在确认窗口中最终确认。

