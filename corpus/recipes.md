# 在 Excel 里做的文件（M 类）

这些文件要在 Excel 里做，因为测试的就是"Excel 做的文件到了 GridCraft 里会怎样"。都放进 `corpus/made/`，在准备批 `prep1`、`prep2` 里做一次。每做好一个，在 `baseline.yaml` 里记下 Excel 的版本。

除非写明，都用 Excel 的"另存为 › Excel 工作簿 (.xlsx)"。Mac 上菜单在"文件"里，Windows 上在"文件 › 另存为"。

## §1 计算的参考副本

用 Excel 打开下面的生成文件，等它算完（状态栏不再显示"计算"），不做任何修改，另存为：

| 打开 | 另存为 | 用于 |
|---|---|---|
| `generated/FX-CORE.xlsx` | `made/FX-CORE.excel.xlsx` | FX-CORE-01、FX-CLI-01 |
| `generated/FX-LEDGER.xlsx` | `made/FX-LEDGER.excel.xlsx` | FX-LEDGER-01、FX-CLI-01 |
| `generated/ZH-TEXT.xlsx` | `made/ZH-TEXT.excel.xlsx` | ZH-FX-01 |

`ZH-TEXT` 的 `TEXT(…,"aaaa")` 和 `[DBNum1]` 的结果取决于 Excel 的语言和区域，`baseline.yaml` 里记下 Excel 显示的是"星期一"还是别的。

再新建一个工作簿，在 A1:A8 依次输入下面的公式，另存为 `made/FX-MISSING.excel.xlsx`（FX-VOL-01）：

```
=NOW()
=RAND()
=CUBEVALUE("x","y")
=WEBSERVICE("https://example.com/")
=IMAGE("https://example.com/logo.png")
=GETPIVOTDATA("Amount",A1)
=INFO("osversion")
=CELL("filename",A1)
```

## §2 Excel 自己的往返

用 Excel 打开 `generated/XL-BASIC.xlsx`，不修改，另存为 `made/XL-BASIC.excel.xlsx`。Excel 另存时自己也会改写文件（补上缓存值、换一种写法），XL-RT-01 的 V 阶段先比较 `XL-BASIC.xlsx` 和这个副本，知道哪些差别 Excel 自己也有，再看 GridCraft 的结果。

## §3 动态数组

用 Excel 打开 `generated/ED-LIST.xlsx`，在下面的单元格输入公式（每个都会溢出到旁边的格子），另存为 `made/FX-DYN.excel.xlsx`（XL-RT-04；FX-DYN-01 在 GridCraft 里输入同样的 8 个公式）：

| 单元格 | 公式 |
|---|---|
| H1 | `=SEQUENCE(5)` |
| J1 | `=SORT(UNIQUE(C2:C201))` |
| L1 | `=UNIQUE(B2:B201)` |
| N1 | `=FILTER(A2:E201,(C2:C201="North")*(E2:E201>5000),"none")` |
| T1 | `=XLOOKUP(10,A2:A201,E2:E201)` |
| T2 | `=LET(total,SUM(E2:E201),n,COUNT(E2:E201),total/n)` |
| T3 | `=MAP(SEQUENCE(3),LAMBDA(v,v*10))` |
| V1 | `=TEXTSPLIT("a,b,c",",")` |

## §4 综合工作簿 XL-EXCEL.xlsx

用 Excel 打开 `generated/ED-LIST.xlsx`，做下面 14 项，另存为 `made/XL-EXCEL.xlsx`。这张清单就是 XL-RT-03 逐项检查的清单。

1. A1:F201 套用表格格式（"中等色 2"），打开汇总行，Amount 列汇总为"求和"。
2. Amount 列加数据条；Units 列加三色刻度；新增一列 G "Flag" `=IF([@Amount]>5000,1,IF([@Amount]>2000,2,3))`，加三个符号的图标集。
3. Region 列加数据验证：序列 `North,South,East,West,Central`。
4. B2 加一条批注（新建批注），再回复一条；B3 加一条备注（新建备注）。
5. A1 的表头加超链接到 `https://example.com/`。
6. 定义名称 `Totals`，引用 Amount 列（`=Table1[Amount]`）。
7. 新建工作表 "Summary"：A2:A6 五个地区，B2 `=SUMIFS(Table1[Amount],Table1[Region],A2)` 向下填充；插入簇状柱形图和饼图。
8. Summary 的 C2:C6 插入折线迷你图，数据是 List 表里各地区前 12 行的 Amount（随便选，能显示就行）。
9. Summary 插入一张自己截的小图片（截一块空白窗口即可，不要用有版权的图）。
10. 保护 Summary 工作表，不设密码。
11. List 表页面设置：横向、调整为 1 页宽、顶端标题行 `$1:$1`。
12. List 表冻结首行。
13. List 表把第 10 到 20 行组合（数据 › 组合）。
14. 文件属性里的作者改为 `GetGridCraft test`。

## §5 其他格式

用 Excel 打开 `generated/XL-BASIC.xlsx`，分别另存为：

| 格式 | 文件 | 用于 |
|---|---|---|
| Excel 97-2004 工作簿（.xls） | `made/XL-OLD.xls` | XL-XLS-01 |
| OpenDocument 电子表格（.ods） | `made/XL-BASIC.ods` | XL-ODS-01 |
| Excel 二进制工作簿（.xlsb） | `made/XL-BASIC.xlsb` | XL-ODS-01 |

另存时 Excel 会提示某些功能不兼容，照常保存，把提示原文记进 `baseline.yaml`。

## §6 图表 CH-EXCEL.xlsx

用 Excel 打开 `generated/ED-LIST.xlsx`，新建工作表 "Charts"，A2:A6 五个地区，B 列各地区 Amount 求和（SUMIFS），C 列各地区 Units 求和。用它做 8 张图，另存为 `made/CH-EXCEL.xlsx`：

1. 簇状柱形图，带数据标签；
2. 堆积条形图（B、C 两列）；
3. 带数据标记的折线图，加一条线性趋势线并显示公式；
4. 饼图，数据标签显示百分比；
5. 散点图（List 表的 Units 对 Amount），纵轴最小值固定为 0、最大值 10000；
6. 组合图：B 列柱形、C 列折线放在次坐标轴；
7. 瀑布图；
8. 三维簇状柱形图。

## §7 数据透视表 PV-EXCEL.xlsx

用 Excel 打开 `generated/ED-LIST.xlsx`（A1:F201，不含合计行），另存为 `made/PV-EXCEL.xlsx` 前做：

1. 新工作表上的透视表 1：行 = Region，列 = Date 按"月"分组，值 = Amount 求和，值显示方式"总计的百分比"。
2. 同一张表上的透视表 2：行 = Name，值 = ID 计数、Units 平均值。
3. 透视表 1 加计算字段 `PerUnit = Amount / Units`（Units 为 0 的行会出错误值，保留）。
4. 透视表 1 插入切片器 Region。
5. 用透视表 1 插入一张数据透视图（簇状柱形）。

## §8 规划求解

PV-WHATIF-01 用的小型线性规划（Excel 要先在加载项里启用"规划求解"）：

| 单元格 | 内容 |
|---|---|
| B1、B2 | 变量 x、y，初值 0 |
| B4 | 目标 `=3*B1+5*B2`，求最大值 |
| B6 | 约束 `=B1`，≤ 4 |
| B7 | 约束 `=2*B2`，≤ 12 |
| B8 | 约束 `=3*B1+2*B2`，≤ 18 |

非负，求解方法"单纯线性规划"。正确答案 x = 2、y = 6、目标 36。

## §9 加密的工作簿

用 Excel 打开 `generated/XL-BASIC.xlsx`，设置打开密码后另存：

- Mac：文件 › 密码…（或"审阅 › 保护工作簿"旁的"用密码加密"，视版本而定），打开密码 `grid-0.15`，另存为 `made/SC-ENC.xlsx`。
- 再做一个，打开密码 `表格密码`，另存为 `made/SC-ENC-ZH.xlsx`。
- Windows：文件 › 信息 › 保护工作簿 › 用密码进行加密。

这两个文件的内容是公开的测试数据，密码写在这里没有关系。

## §10 保护的工作表

用 Excel 打开 `generated/XL-BASIC.xlsx`，在 Data 表选中 B4:B13，设置单元格格式 › 保护，取消"锁定"；然后审阅 › 保护工作表，密码 `prot`，其余选项保持默认。另存为 `made/SC-PROT.xlsx`。

## §11 带 VBA 的工作簿

新建工作簿，打开"开发工具"选项卡（Excel 偏好设置 › 功能区和工具栏，勾选"开发工具"），打开 Visual Basic 编辑器：

在 ThisWorkbook 里：

```vb
Private Sub Workbook_Open()
    MsgBox "VBA ran on open"
End Sub
```

插入一个模块：

```vb
Sub Hello()
    Range("A1").Value = "Hello from VBA"
    MsgBox "Hello from VBA"
End Sub
```

在 Sheet1 上插入一个表单控件按钮，指定宏 `Hello`。另存为"Excel 启用宏的工作簿 (.xlsm)"：`made/SC-VBA.xlsm`。
