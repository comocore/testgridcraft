# FX：公式和计算

回答"算出来的数和 Excel 一样吗？"。数字算错比功能缺失更糟，所以凡是和 Excel 结果不同的，都加 `wrongValue` 标记，不论差多少。依据：README "500+ worksheet functions"；`docs/target-app-parity.md` "513 / 523 (98.1%; 504 working)"，以及 "No conformance suite against Excel-computed oracles; edge-case bugs keep surfacing"；`docs/gaps.md` 第 6、15 条。

参考答案来自 Excel：用 Excel 打开生成的探针、让它算一遍、另存为 `corpus/made/*.excel.xlsx`（recipes.md §1），文件里就存着 Excel 算出的值。比较时：

```sh
node tools/diff-xlsx.mjs corpus/made/FX-CORE.excel.xlsx runs/<本批>/outputs/FX-CORE.gridcraft.xlsx --values > runs/<本批>/evidence/FX-CORE-01-V-diff.txt
```

`outputs/FX-CORE.gridcraft.xlsx` 是在 GridCraft 里打开 Excel 的副本、不做修改、另存得到的：GridCraft 打开时会全部重算，存下的就是它算出的值。

浮点数在最后一两位上的差别（`--tolerance 1e-9` 以内）不算；显示出来的数字不同，或者错误类型不同（`#NUM!` 和 `#VALUE!`），都算。

## 用例

| ID | 优先级 | 系统 | 阶段 | 来源 | 用例 | 怎么测、看什么 |
|---|---|---|---|---|---|---|
| FX-CORE-01 | P0 | Mac Win | O D S V | M | 69 个单格公式 | 打开 `FX-CORE.excel.xlsx`，D 列和 Excel 截图并排看一遍；另存为 `outputs/FX-CORE.gridcraft.xlsx`，V：`diff-xlsx --values`。每个不同的值在 note 里写行号、公式、Excel 的值、GridCraft 的值。1 处不同记 `minor`，2–5 处 `major`，更多 `fail`；都加 `wrongValue` |
| FX-LEDGER-01 | P0 | Mac Win | O D E V | M | 引用区域的公式 | 打开 `FX-LEDGER.excel.xlsx`，Summary 表的 10 个值和 Excel 一样（D、V 同上）。E：把 Ledger!D5 改成 5000、删除第 10 行、在第 20 行插入一行并填数，在 Excel 里做同样的三步，对比 Summary 的每个值 |
| FX-PREC-01 | P0 | Mac Win | E | — | 精度和长数字 | 在空白工作簿里依次输入：`=0.1+0.2=0.3`、`=1-0.9`、`123456789012345678`、`0.1234567890123456789`、`=2^53+1`、`1E+308`、`'123456789012345678`（带撇号）。和 Excel 逐个对比显示的值、编辑栏里的值。长整数超过 15 位变成 0 是 Excel 的行为，一样就是 `pass` |
| FX-DYN-01 | P0 | Mac Win | E | G | 输入动态数组 | 打开 `ED-LIST.xlsx`，在 recipes.md §3 写的单元格里输入那 8 个公式（SEQUENCE、SORT、UNIQUE、FILTER、XLOOKUP、LET、MAP 和 LAMBDA、TEXTSPLIT）。看：溢出范围、蓝色框线、被挡住时的 `#SPILL!`、`H1#` 引用；结果和 Excel 一样 |
| FX-LAMBDA-01 | P1 | Mac Win | E S R A | — | 名称里的 LAMBDA | 名称管理器里定义 `Tax = LAMBDA(x, x*0.13)`，在单元格里用 `=Tax(100)`；保存、重开仍然有效；A：Excel 打开后名称和结果都在 |
| FX-ITER-01 | P1 | Mac Win | E | — | 循环引用和迭代计算 | A1 `=B1+1`、B1 `=A1`：看警告和状态栏；在「选项 › 公式」打开迭代计算（最多 100 次，最大误差 0.001）：结果和 Excel 一样 |
| FX-AUDIT-01 | P2 | Mac Win | E | G | 公式审核 | `FX-LEDGER.xlsx` 的 Summary：追踪引用单元格和从属单元格、公式求值、监视窗口、显示公式（Ctrl+\`）。能不能用，箭头指得对不对 |
| FX-VOL-01 | P2 | Mac Win | O D | M | 易失函数和没有的函数 | `FX-MISSING.excel.xlsx`（recipes.md §1）：NOW、RAND 打开时重算是正常的；CUBEVALUE、WEBSERVICE、IMAGE、GETPIVOTDATA、INFO、CELL 这些 GridCraft 没有、只是占位或结果因电脑而异的函数显示什么，是否盖掉了 Excel 存的值 |
| FX-CLI-01 | P1 | Mac Win | T | M | 上游的检查工具 | `gridcraft-cli check corpus/made/FX-CORE.excel.xlsx corpus/made/FX-LEDGER.excel.xlsx --json > evidence/FX-CLI-01-check.json`。它报告的不一致，和 FX-CORE-01、FX-LEDGER-01 里 diff-xlsx 找到的是否一致；不一致的地方逐个列出 |
