# ED：编辑和格式

回答"日常的录入、填充、排序、筛选、格式，和 Excel 用起来一样吗？会不会悄悄改坏数据？"。依据：README "Familiar. Ribbon, formula bar, Name Box, sheet tabs, keyboard shortcuts, fill handle, point-mode formula entry…"；`docs/target-app-parity.md` 编辑一项 "~90% breadth, ~70% ready"；0.14.2 修复了"删除时清掉筛选隐藏的行"（PR #1422）。

每个用例在 Excel 里做同样的操作对照。操作步骤和结果写进 note，要具体到单元格。

## 用例

| ID | 优先级 | 系统 | 阶段 | 来源 | 用例 | 怎么测、看什么 |
|---|---|---|---|---|---|---|
| ED-FILL-01 | P0 | Mac Win | E | — | 填充柄 | 空白表里拖填充柄：`1, 2` → 序列；`2026-01-31` 按天、按工作日、按月（2 月底是 28 日吗）；`第1组`、`Item 1`；`一月`（Excel 的中文序列）；`周一`；按住 Ctrl（Mac 上是 Option）拖动；双击填充柄填到相邻列的末尾。和 Excel 逐项对比 |
| ED-PASTE-01 | P0 | Mac Win | E | G | 复制和粘贴 | `XL-BASIC.xlsx`：在 GridCraft 内部复制带公式和格式的区域，粘贴后相对引用是否调整；选择性粘贴的值、格式、转置；从 GridCraft 复制到 Excel、从 Excel 复制到 GridCraft（格式和公式是否保留）；从网页表格、Word 或 WPS 文字复制一张表粘进来 |
| ED-FILT-01 | P0 | Mac Win | E S R | G | 筛选后删除和填充 | `ED-LIST.xlsx`：筛选 Region = North，选中 A2:F201 按 Delete，再取消筛选：别的地区的行应该完好（0.14.2 起）。再筛选一次，在可见行里向下填充（Ctrl+D）：隐藏行不应被填。和 Excel 结果逐行对比，保存重开 |
| ED-SORT-01 | P0 | Mac Win | E | G | 排序 | `ED-LIST.xlsx`：按 Region 升序、再按 Amount 降序排序（「数据 › 排序」）；只选中一列时排序的提示；合计行（第 202 行）不应被排进去；撤销后恢复原顺序。中文姓名的顺序见 ZH-SORT-01 |
| ED-FMT-01 | P1 | Mac Win | E S A | — | 设置单元格格式 | 新表里设置：会计格式、百分比、自定义 `0.0,"千"`、日期 `yyyy/m/d aaa`、边框（粗外框、斜线）、填充、合并后居中、自动换行、缩小字体填充、文字方向。保存，A：Excel 打开后是否一样 |
| ED-CF-01 | P1 | Mac Win | E S A | G | 条件格式 | `ED-LIST.xlsx` 的 Amount 列：数据条、三色刻度、图标集、"大于 5000"突出显示、公式规则 `=$C2="North"`。保存后 A：Excel 里规则和显示是否一样 |
| ED-TABLE-01 | P1 | Mac Win | E S A | G | 表格（超级表） | `ED-LIST.xlsx` 套用表格格式，打开汇总行，新增一列 `=[@Amount]*1.13`（结构化引用），加一个切片器。保存后 A |
| ED-UNDO-01 | P1 | Mac Win | E V | G | 撤销到底 | 在 `XL-BASIC.xlsx` 副本上做 30 步混合操作（输入、删除行列、格式、排序、插入工作表），然后一直撤销到头，另存。V：和原文件 diff 应该没有值和公式的差别 |
| ED-KEY-01 | P1 | Mac Win | E | — | 常用快捷键 | 逐个试下面 24 个，在 note 里列出不能用的（Mac 上用对应的 ⌘ 或 ⌃ 组合，以 Excel for Mac 为准）：方向键跳到数据边缘、扩展选区、F2、F4、Ctrl+;、Ctrl+Shift+;、Ctrl+D、Ctrl+R、Ctrl+1、Alt+=、Ctrl+Z、Ctrl+Y、Ctrl+F、Ctrl+H、Ctrl+Space、Shift+Space、Ctrl+-、Ctrl+Shift+=、Ctrl+Page Down、Ctrl+Home、Ctrl+End、Ctrl+Enter、Alt+Enter、F9 |
| ED-FIND-01 | P2 | Mac Win | E | G | 查找和替换 | `ED-LIST.xlsx`：查找"张"、按格式查找、通配符 `张?`、替换全部（是否只在选区内）、在公式里查找 |
| ED-VALID-01 | P2 | Mac Win | E | G | 数据验证 | `XL-BASIC.xlsx` 的 A4:A13 下拉列表：输入列表外的值被拒绝、提示原文；新建一个"整数 1–100"的验证和输入提示 |
