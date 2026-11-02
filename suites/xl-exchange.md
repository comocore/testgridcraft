# XL：和 Excel 互换文件

回答"别人发来的 Excel 文件能不能正常打开？GridCraft 存出来的文件，Excel 打开会不会提示修复、内容会不会变？"。这是读者最关心、也是上游自己排在第一位的差距：`docs/gaps.md` 第 1 条 "XLSX output not proven in Excel or on real files"（"nobody has yet confirmed that Excel opens our output without a repair prompt across a real corpus"）；第 3 条 "Some unmodelled XLSX content is dropped on save"；第 7 条 `.xls` 只读一部分、加密工作簿不能带密码保存。依据的版本：0.15.0。

所有用例都用 `work/` 里的副本。GridCraft 存出的文件放 `outputs/`，名为 `<文件名>.gridcraft.xlsx`；Excel 的参考副本见 [corpus/recipes.md](../corpus/recipes.md)。

A 阶段（交给 Excel）的做法：用 Excel 打开 GridCraft 存的文件。出现"我们发现……中的部分内容有问题，是否尝试恢复"之类的提示就是 `repair`，阶段至少 `major`，把提示和修复报告的原文抄进 note（报告里会列出被删掉的部分）。没有提示时，对照原文件看内容。

V 阶段（检查文件）：`node tools/diff-xlsx.mjs <原文件> outputs/<…>.gridcraft.xlsx > evidence/<ID>-V-diff.txt`，在记录的 `diff` 里写文件名。工具只列出变化，变化是不是问题由测试者判断：例如 `_xlfn.` 前缀、默认格式写法不同不算；值、公式、合并单元格、名称、图表和数据透视表这类部件少了，算 `dataLoss`。

## 用例

| ID | 优先级 | 系统 | 阶段 | 来源 | 用例 | 怎么测、看什么 |
|---|---|---|---|---|---|---|
| XL-RT-01 | P0 | Mac Win | O D S A V | G | 不改动直接往返 | 打开 `XL-BASIC.xlsx`，和 Excel 里打开的同一个文件并排对比三个工作表（隐藏的 Hidden 表应仍隐藏）：合并、冻结、筛选按钮、数字格式、条件格式（负数标红）、下拉列表、链接、名称 Rate。不做任何修改，另存为 `outputs/XL-BASIC.gridcraft.xlsx`。A：Excel 打开有没有修复提示，三个表和原来是否一样。V：先比较原文件和 Excel 自己往返的副本 `XL-BASIC.excel.xlsx`（recipes.md §2），知道哪些差别 Excel 也有；再比较原文件和 GridCraft 存的文件 |
| XL-RT-02 | P0 | Mac Win | O E S A V | G | 编辑后往返 | 打开 `XL-BASIC.xlsx`：C4 改成 99；在第 8 行上方插入一行，填一行数据；把 F1 的 Rate 改成 0.1；A1 改成红色字；删除"格式 Formats"表的第 20 行。保存为 `outputs/XL-BASIC-edit.gridcraft.xlsx`。看 E 列和合计是否随之更新（Excel 里做同样的修改对照）。A、V 同上：除了这几处修改，其他应该不变 |
| XL-RT-03 | P0 | Mac Win | O D S A V | M | Excel 做的综合工作簿 | `XL-EXCEL.xlsx`（recipes.md §4，含表格、各类条件格式、数据验证、批注和备注、图表、迷你图、图片、保护的工作表、页面设置）。O：打开时的提示照抄。D：逐项对照 recipes.md 里的清单，记下哪些显示不同、哪些没了。保存为 `outputs/XL-EXCEL.gridcraft.xlsx`。A：Excel 打开有没有修复提示，清单逐项看还在不在。V：diff-xlsx 的 parts 一节 |
| XL-RT-04 | P1 | Mac Win | O D S A V | M | 动态数组和 LAMBDA | `FX-DYN.excel.xlsx`（recipes.md §3）。D：每个溢出区域的值和 Excel 一样，溢出范围相同。保存后 A：在 Excel 里这些公式仍然溢出，没有被加上 `@`，也没有变成旧式数组公式 `{=…}` |
| XL-XLS-01 | P0 | Mac Win | O D S | M | 老的 .xls 文件 | `XL-OLD.xls`（recipes.md §5，XL-BASIC 另存为 Excel 97-2003）。D：值、公式、数字格式、字体、填充、边框、对齐是否保留（0.15.0 起应保留）；条件格式、数据验证、链接不读是已知的，记下有没有提示。S：按保存时，应该要求另存为 .xlsx，不能写回 .xls |
| XL-ODS-01 | P1 | Mac Win | O D | M | .xlsb 和 .ods | `XL-BASIC.xlsb` 和 `XL-BASIC.ods`（recipes.md §5）。文档说只导入数据、公式取保存的值：核对值对不对，公式是否变成了值，有没有提示说明这一点 |
| XL-CSV-01 | P0 | Mac Win | O D | G | CSV 的日期顺序 | 依次打开 `CSV-DATES-DMY.csv`、`-MDY`、`-AMBIG`、`-MIXED`、`-ISO`。0.15.0 起：DMY 里 03/04/2026 应是 4 月 3 日，MDY 里是 3 月 4 日；AMBIG 应提示"猜的是月在前"；MIXED 应保持文本并提示；ISO 三种写法都应是日期。提示原文照抄。用格式 `yyyy-mm-dd` 看清每个日期 |
| XL-CSV-02 | P1 | Mac Win | E X | G | 另存为 CSV | 打开 `ZH-TEXT.xlsx` 的第一个表，另存为 CSV（`outputs/ZH-TEXT.gridcraft.csv`）。看：文件编码（`file -I` 或 VS Code 右下角）、有没有 BOM、写的是显示的值还是原始值、日期怎么写；再用 Excel 直接双击打开这个 CSV，中文是否乱码 |
| XL-GS-01 | P2 | Mac Win | O D | R | Google 表格导出的文件 | real-world.md 的 R-GS（在 Google 表格里用 FILTER、QUERY、GOOGLETRANSLATE 写的公式，下载为 .xlsx）。0.15.0 起：FILTER 应还原为原公式并算出同样的值；QUERY、GOOGLETRANSLATE 保留包装、显示 Google 存的值 |
| XL-REAL-01~03 | P0 | Mac Win | O D S A V | R | 真实工作簿（三份） | real-world.md 的 R-XL-1 到 R-XL-3。O：打开用时大概多久、有什么提示；D：和 Excel 对照，记下每处可见的差别；保存后 A、V |
| XL-REAL-04~06 | P1 | Mac Win | O D S A V | R | 真实工作簿（另外三份） | real-world.md 的 R-XL-4 到 R-XL-6，做法同上 |
