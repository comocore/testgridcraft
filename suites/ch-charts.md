# CH：图表和对象

回答"能做出常用的图表吗？Excel 里做的图表打开、保存之后还对吗？"。图表是上游评估里最不成熟的部分：`docs/target-app-parity.md` "Charts and sparklines ~50% breadth, ~30% ready"；`docs/gaps.md` 第 4 条 "Charts lose options on round trip; shallow chart model"；`docs/chart-parity.md` 没有三维图、曲面图和地图。依据的版本：0.15.0。

## 用例

| ID | 优先级 | 系统 | 阶段 | 来源 | 用例 | 怎么测、看什么 |
|---|---|---|---|---|---|---|
| CH-NEW-01 | P0 | Mac Win | E S A | G | 新建常用图表 | 在 `ED-LIST.xlsx` 旁边建一个按地区汇总的小表（5 行），用它插入：簇状柱形图、折线图、饼图、条形图、散点图（Amount 对 Units）、柱形加折线的组合图（次坐标轴）。改标题、图例位置、加数据标签。保存后 A：Excel 里每张图的类型、系列、标题、图例和数据标签是否一样 |
| CH-RT-01 | P0 | Mac Win | O D S A | M | Excel 做的图表 | `CH-EXCEL.xlsx`（recipes.md §6，8 张图）。D：逐张和 Excel 并排截图，记下画错和画不出的。保存后 A：在 Excel 里逐张看类型、坐标轴刻度、数据标签、趋势线、颜色还在不在 |
| CH-SPARK-01 | P1 | Mac Win | E S A | G | 迷你图 | `ED-LIST.xlsx` 汇总表上插入折线、柱形、盈亏迷你图，标出高点和低点。保存后 A |
| CH-PIC-01 | P1 | Mac Win | E S A | — | 图片、形状和文本框 | 插入一张 PNG 图片、一张放进单元格的图片、一个箭头形状、一个中文文本框、一个复选框。移动行列后它们跟着走吗；保存后 A |
| CH-MOVE-01 | P2 | Mac Win | E | G | 移动和调整大小 | 拖动图表、用八个控制点改大小（0.15.0 之后的 main 分支才有统一的控制点，记下当时版本的行为）、按 Delete 删除、撤销 |
