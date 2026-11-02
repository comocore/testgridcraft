# PV：数据透视表和数据工具

回答"数据透视表能用吗？总数对不对？Excel 里的透视表打开后还能刷新吗？"。依据：README "Pivot tables. Field list, filters, layouts, date grouping and show-values-as, saved to and read from XLSX"；`docs/target-app-parity.md` "PivotTables ~65% breadth, ~45% ready … Round-trip never checked in Excel"；`docs/gaps.md` 第 8 条。

透视表的数字和 Excel 用同一份数据（`ED-LIST.xlsx`）做的透视表逐格对比，任何一格不同都加 `wrongValue`。

## 用例

| ID | 优先级 | 系统 | 阶段 | 来源 | 用例 | 怎么测、看什么 |
|---|---|---|---|---|---|---|
| PV-NEW-01 | P0 | Mac Win | E S A | G | 新建数据透视表 | `ED-LIST.xlsx` 的 A1:F201：行 = Region，列 = Date 按月分组，值 = Amount 求和、Units 平均值，筛选 = Name。和 Excel 做的同一张透视表逐格对比总计和小计。改一个源数据、刷新。保存后 A：Excel 打开不提示修复，能刷新，刷新后数字不变 |
| PV-RT-01 | P0 | Mac Win | O D E S A | M | Excel 做的透视表 | `PV-EXCEL.xlsx`（recipes.md §7：两张透视表、一个计算字段、值显示方式"占总计的百分比"、一个切片器、一张透视图）。D：数字和布局与 Excel 一样。E：在源数据里改两个值，刷新，数字和 Excel 刷新后的一样。保存后 A |
| PV-SLICER-01 | P1 | Mac Win | E S A | G | 切片器和透视图 | 在 PV-NEW-01 的透视表上加切片器（Region）和一张透视柱形图，用切片器筛选。保存后 A |
| PV-WHATIF-01 | P1 | Mac Win | E | — | 单变量求解和规划求解 | 单变量求解：`=PMT(5%/12,B1,-300000)` 要等于 2000 时 B1 是多少。规划求解：recipes.md §8 的小型线性规划。结果和 Excel 的一样 |
| PV-SUB-01 | P2 | Mac Win | E | G | 分类汇总和删除重复项 | `ED-LIST.xlsx` 按 Region 排序后分类汇总 Amount；分级显示的 1、2、3 级；对 Name 列删除重复项，提示删了几行，和 Excel 一样吗 |
