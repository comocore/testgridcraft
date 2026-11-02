# PF：性能

回答"打开大表快不快，改一个数要等多久，占多少内存？"。依据：README "Fast. Copy-on-write workbooks…, incremental recalculation, and a virtualised grid"；`docs/target-app-parity.md` "Performance ~35%"；`docs/gaps.md` 第 2 条 "Large workbooks don't open; recalc is single-threaded"（issue #175，约 100 MB 的 XLSX）。

## 怎么量

- 条件：接电源，关掉低电量模式；关掉其他应用；GridCraft 和 Excel 不同时开着。每项量 5 次，记在 `gridcraft` 和 `excel` 两个数组里，第一次（冷启动或冷打开）也算在内，note 里说明。
- 计时：屏幕录制（Mac：⌘⇧5；Windows：Win+Alt+R），同时在旁边开着 `tools/stopwatch.html`，逐帧看起点和终点的数字。起点是按下回车或双击的那一帧，终点在每个用例里写明。
- 内存：Mac 用"活动监视器"的"内存"列，Windows 用任务管理器"详细信息"里的"内存（专用工作集）"，在终点后 10 秒读数。
- 文件：`PF-ROWS-1M.xlsx` 要用 `node tools/make-corpus.mjs --big` 生成。

## 用例

| ID | 优先级 | 系统 | 阶段 | 来源 | 用例 | 怎么测、看什么 |
|---|---|---|---|---|---|---|
| PF-START-01 | P0 | Mac Win | M | — | 启动 | 从点图标到出现可以输入的空白工作簿（秒） |
| PF-OPEN-01 | P0 | Mac Win | M | G | 打开 10 万行 | `PF-ROWS-100K.xlsx`：从打开到第一屏数据显示（秒） |
| PF-CALC-01 | P0 | Mac Win | M | G | 改一个数后重算 | `PF-CALC-20K.xlsx` 打开后，把 A2 改成 1000 回车：到 B20001 的累计和 E 列更新为止（秒） |
| PF-MEM-01 | P0 | Mac Win | M | G | 内存 | `PF-ROWS-100K.xlsx` 打开后的内存（MB） |
| PF-OPEN-02 | P1 | Mac Win | M | G | 打开 100 万行 | `PF-ROWS-1M.xlsx`（约 40 MB）：从打开到第一屏数据（秒）；打不开就记提示原文，等级写在 note 里 |
| PF-SAVE-01 | P1 | Mac Win | M | G | 保存 10 万行 | `PF-ROWS-100K.xlsx` 改一个格后另存（秒） |
| PF-SORT-01 | P1 | Mac Win | M | G | 排序 10 万行 | `PF-ROWS-100K.xlsx` 按 Col 3 降序排序（秒） |
| PF-SIZE-01 | P2 | Mac Win | M | G | 文件大小 | `PF-ROWS-100K.xlsx` 不改动另存后的大小（MB），和 Excel 另存的对比 |
