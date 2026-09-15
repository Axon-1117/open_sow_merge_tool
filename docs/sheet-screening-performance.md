# Sheet 前置筛选性能验证（update115）

使用真实 SVN 历史 Building.xlsx r41332 → r41333 的安全副本。只读查询确认 r41333 修改 `/sheets/v1.2/Building.xlsx`。SVN 日志只提供文件范围，Sheet 筛选依据实际工作簿内容。

同机隐藏窗口对照：旧 ZIP CRC 筛选进入精确比较 17/20 页，总耗时 22.41 秒；新筛选进入精确比较 1/20 页（Building@design），总耗时 12.07 秒，减少约 46%。这是单次配对实测，不是 p95；实际办公负载影响耗时。

筛选先完成所有 Sheet 判断，再放行精确计算。忽略保存的视图、滚动位置、尺寸提示、Excel 生产版本及绝对保存路径；解析共享字符串实际内容，保留公式、缓存值、样式、关系和结构。异常时返回未知，继续精确比较。精确比较与原有保存规则保持一致。

诊断入口：`tools/benchmark_startup.py`。设置 `SOW_PROFILE_BUILDING` 和 `SOW_PROFILE_RIGHT` 为历史版本文件；设置 `SOW_PROFILE_BUILDING_UI=1`、`SOW_PROFILE_FULL_COMPARE=1` 获取全流程耗时及筛掉的 Sheet 数。输入均创建临时副本。

残余瓶颈：唯一变化页的精确计算和 UI 回调仍出现超过 200ms 的心跳间隔，本轮未达到原计划的心跳目标。后续应独立优化精确比较及结果应用，不能以首屏时间代替最终完成时间。

验收：121 项 pytest 与 Full 所选烟测通过。Full 的真实 SVN 集成因新增 Added.xlsx 无 BASE revision 被 freshness 门禁阻断；在未修改的 01b15cc 基线副本上复现同样失败。此问题不来自本轮 Sheet 筛选。候选包可构建，部署等待集成门禁恢复。
