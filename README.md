# GridCraft 实测方案

Hands-on tests of [GridCraft](https://github.com/storytold/gridcraft), the open-source spreadsheet app, with Microsoft Excel on the same machine as the reference: the method, the cases, the test workbook generator, a workbook comparer, and the raw records behind the test pages on [getgridcraft.com](https://getgridcraft.com). The method is written in Chinese; case IDs, file names and record fields are the same on every page. GetGridCraft is an independent site, not affiliated with ArtCraft or Microsoft.

这里是 [getgridcraft.com](https://getgridcraft.com) 测试页面背后的方法、用例、测试文件生成脚本、工作簿比较工具和原始记录。GetGridCraft 是独立网站，和 ArtCraft、微软都没有关系。许可见 [corpus/README.md](corpus/README.md#公开和许可)。

> 版本 1.0 · 2026-11-02 · 适用于 GridCraft 0.15.0 及以后
>
> 本方案和全部用例在开始测试之前提交。之后的修改记在文末的"修订记录"里，测试页面会链接到这里，读者可以看到我们有没有在测完之后改标准。

## 1. 要回答的问题

| 读者的问题 | 套件 |
|---|---|
| 别人发来的 Excel 文件能不能打开？GridCraft 存的文件，Excel 打开会不会提示修复？老的 .xls 呢？CSV 里的日期会不会颠倒？ | [XL：和 Excel 互换文件](suites/xl-exchange.md) |
| 算出来的数和 Excel 一样吗？长数字会不会变？动态数组能用吗？ | [FX：公式和计算](suites/fx-formulas.md) |
| 填充、粘贴、排序、筛选、格式用起来和 Excel 一样吗？会不会悄悄改坏数据？ | [ED：编辑和格式](suites/ed-editing.md) |
| 图表做得出来吗？Excel 里的图表保存后还对吗？ | [CH：图表和对象](suites/ch-charts.md) |
| 数据透视表的数对不对？能刷新吗？ | [PV：数据透视表和数据工具](suites/pv-data.md) |
| 能打印、导出 PDF 吗？ | [PR：打印和 PDF](suites/pr-print.md) |
| 有中文界面吗？中文输入法、中文显示、中文 CSV、身份证号、中文 PDF 怎么样？ | [ZH：中文](suites/zh-chinese.md) |
| 设了密码的文件能打开吗？保护、链接、宏、VBA 安全吗？会不会泄露作者名？ | [SC：密码、宏、链接和隐私](suites/sc-security.md) |
| 命令行和 MCP 是不是像文档写的那样工作？AI 智能体能碰到哪些文件？ | [MC：MCP 和命令行](suites/mc-mcp-cli.md) |
| 安装、更新、卸载顺不顺利？会不会联网？ | [IN：下载、安装和网络](suites/in-install.md) |
| 打开大表快不快，重算要等多久，占多少内存？ | [PF：性能](suites/pf-performance.md) |
| 会不会崩溃、丢数据？坏文件会不会把它弄崩？ | [ST：稳定性和数据安全](suites/st-stability.md) |

上游自己的文档列出了很多"没验证过"的地方：`docs/gaps.md` 的第 1 号差距是 "XLSX output not proven in Excel or on real files"，第 6 号说计算正确性要等到有"Excel 算出的标准答案"测试集才能下结论，`docs/target-app-parity.md` 写着 PivotTable 的往返 "never checked in Excel"。各套件开头写了它依据的是哪几条。

## 2. 原则

- **只发布测到的。** 没测的写"未测"，不拿官方说法顶替。官方说法和实测不一致时，两者都写，注明版本和日期。
- **结果和环境绑在一起。** 每个结果都写明 GridCraft 版本、系统、机器、Excel 版本和测试日期。
- **先定方法，再测。** 等级定义、用例和步骤在测试前定好。
- **以同一台机器上的 Microsoft Excel 为参考。** 计算结果的标准答案是 Excel 算出、存在文件里的值；显示效果只在同一台机器上比较。没有 Excel 订阅时，可以用免费的 Excel 网页版做"打开有没有修复提示"和看值的检查，`run.yaml` 里写明；LibreOffice Calc 是第二参考，只在用例写明时用。
- **人来判定。** 脚本只负责生成文件、比较文件（`diff-xlsx.mjs`）和计时；每一个等级都由测试者看过之后给出。MCP 套件里 AI 智能体本身是被测对象，判定仍由人做。
- **数字和文件要有证据。** 计算和往返的用例，除了看屏幕，还要用 `diff-xlsx.mjs` 比较文件，输出存进记录。
- **拿不准时选更差的一级**，并在备注里写原因。
- **不打总分，不打星。** 只报告各等级的数量和具体问题。
- **可以复现。** 测试文件、生成脚本、原始记录和截图都保留，能公开的公开（§10）。

## 3. 对象、环境和准备

### 3.1 版本

测试分批做（§4.1），每一批开始时用当时最新的正式版，一批之内不换版本。GridCraft 几天就出一个新版本，每个结果都记着它是在哪个版本上测的，页面上照实写；已经测过的批次由回归批跟进（§8）。

安装包只从官方 GitHub Release 下载，下载后用 `SHA256SUMS.txt` 核对，把文件名和 SHA-256 写进 `run.yaml`。

### 3.2 机器

| 机器 | 用途 | 参考软件 |
|---|---|---|
| Mac（Apple M4，32 GB） | 主测试机 | Microsoft Excel for Mac（Microsoft 365） |
| Windows 电脑 | 第二台测试机 | Microsoft Excel for Windows（Microsoft 365） |

两台机器的具体配置第一次测试时写进 `run.yaml`。

### 3.3 每台机器的准备

1. **新建一个专门测试用的系统用户**（如 "Tester"），在这个用户下测试，让 GridCraft 以默认设置启动，截图里也不会出现自己的名字和文件。
2. 系统语言用简体中文，区域格式用"中国大陆"。英文页面需要的少量英文界面截图，测完后在「视图 › 界面语言」里临时切换补拍。Excel 也用中文界面，写进 `run.yaml`。
3. 关掉 Excel 和系统的自动更新，避免中途换版本。实在更新了，写进 `run.yaml` 的 `changes`。
4. 记下系统里装的中文字体（`run.yaml` 的 `fonts.chinese_installed`）。**不要为了测试另装字体**：读者的电脑也不会为了 GridCraft 去装字体。
5. 截图前，GridCraft 和 Excel 都设 100% 缩放，窗口一样大，同一个工作表、同一个位置。
6. 截图保存位置设到本批的 `evidence/` 目录（Mac：⌘⇧5 › 选项 › 存储位置）。
7. 性能测试前接上电源，关掉低电量模式（pf-performance.md 有完整条件）。
8. 测试前把 GridCraft 的设置文件夹改名（Mac：`~/Library/Application Support/GridCraft`；Windows：`%APPDATA%\GridCraft`），`run.yaml` 里 `settings: fresh`。标题栏的"自动保存"保持关闭（默认），除非用例要求打开。
9. 装好 Node.js 24 或更新版本（运行 `tools/` 里的脚本）和网络监视工具（Mac：LuLu），记进 `run.yaml` 的 `tools`。

### 3.4 测试文件

在准备批 `prep1`、`prep2` 里做一次（[batches.md](batches.md)）：

1. 生成：`node tools/make-corpus.mjs`，输出在 `corpus/generated/`（加 `--big` 再生成 100 万行的文件）。生成的文件里的公式没有缓存值，Excel 打开时会全部计算。
2. 用 Excel 逐个打开 `corpus/generated/` 里的 .xlsx（`ST-BAD-*` 和 `PF-*` 除外），对照每个文件第一行或说明列里写的预期。和预期不符的记在 `corpus/baseline.yaml`，告诉编辑修生成脚本。打开时 Excel 提示修复，说明是生成脚本的问题，同样告诉编辑。
3. 按 [corpus/recipes.md](corpus/recipes.md) 用 Excel 另存参考副本、做综合工作簿、加密文件和带 VBA 的文件，放进 `corpus/made/`。
4. 按 [corpus/real-world.md](corpus/real-world.md) 下载 R 类文件，记下来源和 SHA-256。

## 4. 套件、优先级和用时

- **P0**：第 1 轮先测的范围，各测试页面的第一版由这些用例的结果组成。
- **P1**：第 1 轮完整测试的范围，P0 测完后再做。
- **P2**：有时间再做，或者读者问到再做。

| 套件 | 用例 | P0 | P1 | P2 |
|---|---|---|---|---|
| XL 和 Excel 互换文件 | 15 | 8 | 6 | 1 |
| FX 公式和计算 | 9 | 4 | 3 | 2 |
| ED 编辑和格式 | 11 | 4 | 5 | 2 |
| CH 图表和对象 | 5 | 2 | 2 | 1 |
| PV 数据透视表和数据工具 | 5 | 2 | 2 | 1 |
| PR 打印和 PDF | 4 | 1 | 2 | 1 |
| ZH 中文 | 9 | 6 | 2 | 1 |
| SC 密码、宏、链接和隐私 | 6 | 4 | 2 | 0 |
| MC MCP 和命令行 | 6 | 3 | 2 | 1 |
| IN 下载、安装和网络 | 11 | 6 | 4 | 1 |
| PF 性能 | 8 | 4 | 3 | 1 |
| ST 稳定性和数据安全 | 12 | 6 | 6 | 0 |
| 合计 | 101 | 50 | 39 | 12 |

用例数包括只在一个系统上测的用例，`new-run.mjs` 会按系统列出准确的清单。上表用 `node tools/count-cases.mjs` 统计（和 `new-run.mjs` 用同一个解析器），改了用例要重新统计。

### 4.1 分批

全部测试拆成每批不超过约 1 小时的小批，测完一批就上线一批（[batches.md](batches.md)）。每批的用时里包含约 10 分钟的固定部分（开始、复判上一批抽到的用例、检查和提交），测试本身约 45–50 分钟。用时是估计值，前几批做完后按实际情况更新 batches.md。一批做不完，就停在用例之间，下一批先补上；不为了赶时间降低判定标准。

## 5. 判定

### 5.1 等级

| 等级 | 含义 | 例子 |
|---|---|---|
| `pass` 一致 | 和 Excel 看不出差别；功能按预期完成 | 同样的值、格式和布局；字体渲染的细微差别不算 |
| `minor` 小差别 | 内容完整、数字正确，有看得见但不影响使用的差别 | 行高差几个像素；图表颜色略有不同；一个不常用的格式显示不同 |
| `major` 明显差别 | 数字没错，但显示或结果错到影响使用；或者要绕路才能完成 | 日期显示成数字；条件格式没生效；Excel 打开时提示修复但修复后内容完整 |
| `fail` 失败 | 打不开；数字算错；内容丢失；崩溃、卡死；功能没有或者用不了 | 一个公式结果和 Excel 不同；图表或透视表保存后没了；Excel 修复时删掉了内容 |
| `na` 不适用 | 这台机器或这个用例不涉及 | 只有 Windows 才有的安装选项，在 Mac 上 |
| `blocked` 受阻 | 前一步失败，这一步没法测 | O 是 `fail`，后面的阶段都记 `blocked` |
| `skip` 没测 | 这次没测，note 写原因 | 没有打印机 |

**数字算错一律至少 `fail`。** 一个读者会信以为真的错误数字，比一个缺失的功能更糟。只有用例里写明的分级（例如 FX-CORE-01 按错误个数分级）例外。

### 5.2 标记

| 标记 | 什么时候用 |
|---|---|
| `wrongValue` | 计算结果、透视表的数、显示的数字和 Excel 不同。页面上单独列出，并向上游报告（§9） |
| `dataLoss` | 保存或往返之后，值、公式、格式、批注、图表、透视表、名称等内容消失或改变，而且没有任何提示。页面上单独列出 |
| `repair` | Excel 打开 GridCraft 存的文件时提示有问题、要求修复 |
| `crash` | 崩溃，或者 30 秒没有反应。同时在 ST-LOG-01 里记一笔 |
| `leak` | 存出的文件里留下了不该留的东西：本该加密却没加密而且覆盖了原文件、真实姓名写进了作者属性、删除的数据还在文件里 |
| `warning` | GridCraft 打开或保存时给出提示。note 里照抄提示原文 |

### 5.3 阶段

| 代号 | 阶段 | 做什么 |
|---|---|---|
| O | 打开 | 在 GridCraft 里打开文件 |
| D | 显示 | 和 Excel 同一个位置、同样缩放对比 |
| E | 操作 | 做用例写明的编辑或操作 |
| S | 保存 | 保存或另存为，记下用的是哪一个、存成什么格式 |
| R | 重开 | 在 GridCraft 里重新打开存出的文件，和保存前对比 |
| A | 交给 Excel | 用 Excel 打开 GridCraft 存出的文件，看有没有修复提示、内容对不对 |
| V | 检查文件 | 用 `diff-xlsx.mjs` 比较原文件（或 Excel 的参考副本）和 GridCraft 存出的文件 |
| X | 导出 | 导出或转换（CSV、PDF），检查得到的文件 |
| T | 任务 | 按步骤完成任务 |
| M | 测量 | 记录数值 |

### 5.4 规则

- 一个阶段只给一个等级，有几处问题时按最重的给。
- 拿不准时选更差的一级，note 里写原因。
- 因为本机没有某个字体、GridCraft 用了替代字体而造成的差别，照常给等级（读者遇到的就是这个结果），note 里写"字体替换：原字体 → 实际字体"。Excel 也替换了的，写两边各用了什么。
- 上游文档写明不支持的功能，照样测、照样给等级。有没有提示、会不会悄悄出错，比"能不能做"更重要。
- **V 阶段**：`diff-xlsx.mjs` 列出的每一类变化都要看一遍。值和公式的变化、部件的减少（图表、透视表、批注……）是 `dataLoss`；写法上的差别（`_xlfn.` 前缀、空格、等价的格式代码）不算。工具不比较字体、填充、边框、列宽和图表细节，这些靠 A 阶段的目视检查。
- 先写 note，再给等级。note 要写到不看截图也能明白：哪个工作表、哪个单元格、和 Excel 比怎么不同。

### 5.5 复判

每批收尾时用 `check-run.mjs --sample 10` 随机抽 10% 的用例（至少 1 个），在之后某一批的开头不看原记录重新判定，和第一次判定至少隔一天。做法见 [runs/README.md](runs/README.md)。页面上报告两次判定的一致率。

## 6. 证据

**截图**

- 文件名：`<用例ID>-<阶段>-<软件>[-<序号>].png`，软件写 `gridcraft`、`excel`、`libreoffice`、`system`（系统弹窗）。例：`XL-RT-03-A-excel.png`、`IN-SIG-01-T-system-1.png`。
- 对比截图两边同一个工作表、同一位置、同缩放、窗口一样大，只截窗口。
- 等级是 `minor`、`major`、`fail` 的阶段，至少一对截图。P0 用例即使 `pass` 也截一对，页面展示用。
- 截图里不能有个人信息：用户名、邮箱、真实文件名和路径。

**录屏**：启动、打开计时、卡顿和崩溃用录屏，命名同截图，扩展名 `.mov` 或 `.mp4`。

**文件**：GridCraft 保存或导出的文件放 `outputs/`，名为 `<文件名>.gridcraft.<扩展名>`，例如 `XL-BASIC.gridcraft.xlsx`；Excel 为对照存的文件名为 `<文件名>.excel.<扩展名>`。

**检查输出**：`diff-xlsx.mjs`、`gridcraft-cli check` 等的输出存成 `evidence/<ID>-<说明>.txt`，例如 `node tools/diff-xlsx.mjs corpus/generated/XL-BASIC.xlsx runs/…/outputs/XL-BASIC.gridcraft.xlsx > runs/…/evidence/XL-RT-01-V-diff.txt`。

**日志**：每批测完把 GridCraft 日志复制到 `logs/<日期>.log`；崩溃后马上复制 `gridcraft.log` 和 `gridcraft.1.log`。位置：Mac `~/Library/Application Support/GridCraft/logs/`，Windows `%APPDATA%\GridCraft\logs\`。

## 7. 记录

每一批一个记录目录，格式和工具见 [runs/README.md](runs/README.md)。每批的流程（详见 [batches.md](batches.md)）：

1. `node tools/new-run.mjs --os mac --version 0.15.0 --batch xl1 --from <上一批的目录>`。
2. 复判上一批抽到的用例。
3. 按本批的用例测，边测边填对应的 `.yaml`。
4. `node tools/check-run.mjs <本批目录> --sample 10`，修掉所有 error，记下要复判的用例，提交。

## 8. 回归测试

每个新版本发布后三天内，在 Mac 上做一个回归批 `reg`（约 50 分钟）：五个覆盖面最广的核心用例 XL-RT-01、XL-RT-03、FX-CORE-01、ZH-SHOW-01、ST-SAV-02。一周里出了几个版本，只回归最新的那个。

发布说明里提到修了我们判过非 `pass` 的用例时，把这些用例另排回归批（`--batch reg --add …`），每批照样控制在 1 小时内。

结果写进该版本的发布说明页（"实测变化"），测试页面上标出哪些结果已经在新版本上复测过。每隔大约一个月，或者上游宣布进入 Beta 时，在最新版本上重做一轮 P0。

## 9. 向上游报告问题

1. 先在 [GridCraft 的 issue](https://github.com/storytold/gridcraft/issues) 里搜，有人报过就在那条下面补充我们的版本、系统和测试文件。
2. 没有就新建：写清步骤、预期（Excel 的结果）、实际结果、版本、系统，附上最小的测试工作簿（我们自己做的文件以 CC BY 4.0 提供）和截图、日志片段。上游的 `gridcraft-cli check` 输出也可以附上。
3. **安全问题不发公开 issue**：用 GitHub 仓库的 Security › Report a vulnerability 私下报告，`issue` 字段写 `private report`。上游修好、发布之后，页面再写细节。
4. 把 issue 地址填进记录的 `issue` 字段。问题修好后，在回归测试里复测。

issue 只用来说明问题，不在 issue 里推广 GetGridCraft。

## 10. 公开

- 测试文件：见 [corpus/README.md](corpus/README.md)。我们做的文件以 CC BY 4.0 公开；第三方文件只公开来源和 SHA-256。
- 原始记录：`runs/` 里的 `run.yaml` 和各套件的记录、截图、检查输出，整体公开。私有样本只公开汇总。还没修好的安全问题，记录里只写"已私下报告"，细节等上游发布修复后再补。
- 页面：在 [getgridcraft.com](https://getgridcraft.com) 上。每批测完后 1–2 天内写进页面上线。页面写明测试日期、版本、机器和 Excel 版本，计划和已测的用例数，链接到本方案和原始记录。

## 11. 局限

- 一位测试者，两台机器。判定有主观成分，用复判来控制。
- 参考是测试机上的那个 Excel 版本。Excel 自己在不同版本、不同平台上也有差别（比如 Mac 版没有 Power Pivot），结果只对写明的版本有效。
- 测试文件和真实工作簿的样本有限，不代表所有文件。生成的探针刻意简单，真实文件的问题要靠 R 类文件发现。
- `diff-xlsx.mjs` 只比较它列出的那些方面。它说"没有差别"，只说明值、公式、数字格式、合并、名称和部件没变，不证明字体、图表细节、打印设置都没变。
- 结果只对写明的 GridCraft 版本有效。

## 修订记录

| 日期 | 版本 | 修改 |
|---|---|---|
| 2026-11-02 | 1.0 | 第一版：12 个套件，101 个用例，测试文件生成脚本、工作簿比较工具和记录工具 |
