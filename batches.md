# 分批计划

> 方案 1.0 · 2026-11-02。每一批在一台机器上做，尽量不超过 1 小时；测完一批，就把这一批的结果写进页面上线。

## 每一批怎么做

每批的用时里都包含约 10 分钟的固定部分：

1. **开始（约 3 分钟）**：确认装的是不是最新的正式版（§版本），运行 `node tools/new-run.mjs --os mac --version <版本> --batch <批次> --from <上一批的目录>`。`--from` 会把上一批 `run.yaml` 里的机器、参考软件和字体信息带过来，只要核对有没有变。把 GridCraft 的设置文件夹改个名（README.md §3.3），让它以默认设置启动。
2. **复判（约 4 分钟）**：上一批结束时抽到的用例，不看原记录，重新判定，写进上一批记录的 `recheck`（README.md §5.5）。第一批没有这一步。
3. **测试（约 45–50 分钟）**：按下表的用例做。测试文件都从 `corpus/` 复制到本批的 `work/` 里再用，原件不动。
4. **收尾（约 3 分钟）**：`node tools/check-run.mjs <本批目录> --sample 10`，修掉所有 error，记下抽到的复判用例，提交。

做不完就停在用例之间，没做的留空，下一批先补上。不要为了赶时间降低标准：拿不准仍然选更差的一级。每一批都自动带上 ST-LOG-01（崩溃记录）。

### 版本

每一批开始时用当时最新的正式版，一批之内不换版本。上一批之后有了新版本，就先装新版本（下载、核对 SHA-256，约 5 分钟，算在这一批里），这一批的结果记在新版本上。已经测过的批次不因为新版本重测，由回归批（`reg`）跟进。

### 上线

- 测完一批后 1–2 天内上线：编辑把记录写进"测完上线"一栏的页面，过 content-gate，核对页面和记录一致，部署。
- 第一次出现的页面（`/tests`、`/tests/excel-files`、`/tests/formulas`、`/zh/tests/chinese`、`/tests/performance`），第一批就要能单独回答一个读者的问题，下表的顺序是按这个排的。`/zh/tests/chinese` 只有中文版。
- 页面写明"计划 M 个用例，已测 N 个"，没测的列为"未测"。每个结果标出测试时的 GridCraft 版本和 Excel 版本。
- 中英文都出的页面，两种语言同一次更新。

## 准备（每台机器一次，不上线）

| 批次 | 内容 | 用例 | Mac 用时 | 测完上线 |
|---|---|---|---|---|
| `prep1` | 测试用户、系统语言和区域、LuLu；`node tools/make-corpus.mjs`；用 Excel 打开生成的探针并另存参考副本（recipes.md §1、§2、§5）；下载 R 类文件并核对 SHA-256（real-world.md） | — | 约 55 分钟 | 不上线 |
| `prep2` | 在 Excel 里按 recipes.md 做文件：动态数组（§3）、综合工作簿（§4）、图表（§6）、透视表（§7）、加密和保护（§9、§10）、带 VBA 的 .xlsm（§11） | — | 约 60 分钟 | 不上线 |

Windows 上可以直接复制 Mac 上做好的 `corpus/made/`（它们不进仓库的话，用 U 盘或网盘），不用重做 `prep2`。

## 第 1 轮：P0（每台机器 12 批）

Mac 先做，Windows 再按同样的批次做一遍。只在一个系统上测的用例（比如 IN-SIG-01 只在 Mac，IN-INS-01 只在 Windows），`new-run.mjs` 会按系统挑出来。用时是 Mac 上的估计，第一批做完后按实际改。

| 批次 | 内容 | 用例 | Mac 用时 | 测完上线 |
|---|---|---|---|---|
| `in1` | 下载、签名、安装、首次启动和网络 | IN-DL-01 IN-SIG-01 IN-SIG-02 IN-INS-01 IN-1ST-01 IN-NET-01 | 约 45 分钟 | 新页面 `/tests`（方法、机器、进度）；`/download/*` 换成实际看到的步骤；`/safety` 加上签名和网络连接的实测 |
| `xl1` | 往返、编辑后往返、老 .xls、CSV 日期 | XL-RT-01 XL-RT-02 XL-XLS-01 XL-CSV-01 | 约 55 分钟 | 新页面 `/tests/excel-files` |
| `xl2` | Excel 做的综合工作簿和三份真实文件 | XL-RT-03 XL-REAL-01 XL-REAL-02 XL-REAL-03 | 约 60 分钟 | `/tests/excel-files`；`/compare/microsoft-excel` 的"还缺什么"加上实测 |
| `fx1` | 单格公式、区域公式、精度、动态数组 | FX-CORE-01 FX-LEDGER-01 FX-PREC-01 FX-DYN-01 | 约 55 分钟 | 新页面 `/tests/formulas` |
| `zh1` | 中文界面、输入法、显示 | ZH-UI-01 ZH-IME-01 ZH-SHOW-01 | 约 45 分钟 | 新页面 `/zh/tests/chinese`（只有中文版） |
| `zh2` | 中文 CSV、身份证号、中文 PDF 和导出 PDF | ZH-CSV-01 ZH-IDS-01 ZH-PDF-01 PR-PDF-01 | 约 50 分钟 | `/zh/tests/chinese`；`/tests/excel-files` 加 PDF |
| `ed1` | 填充、粘贴、筛选后删除、排序 | ED-FILL-01 ED-PASTE-01 ED-FILT-01 ED-SORT-01 | 约 50 分钟 | 新页面 `/tests/editing` |
| `ch1` | 图表和数据透视表 | CH-NEW-01 CH-RT-01 PV-NEW-01 PV-RT-01 | 约 60 分钟 | `/tests/editing`；`/compare/microsoft-excel` 的图表和透视表加上实测 |
| `sc1` | 加密、保护、链接、VBA | SC-ENC-01 SC-PROT-01 SC-LINK-01 SC-VBA-01 | 约 50 分钟 | `/safety` 的密码、链接和 VBA 加上实测 |
| `mc1` | 命令行、Claude Code 接入、文件边界 | MC-CLI-01 MC-MCP-01 MC-MCP-02 | 约 50 分钟 | 改写 `/mcp` |
| `st1` | 坏文件、磁盘满、保存时强制退出、没保存的修改 | ST-BAD-01 ST-BAD-02 ST-SAV-01 ST-SAV-02 ST-REC-01 | 约 50 分钟 | `/safety` 加"会不会丢数据"；`/tests` |
| `pf1` | 启动、打开 10 万行、重算、内存 | PF-START-01 PF-OPEN-01 PF-CALC-01 PF-MEM-01 | 约 50 分钟 | 新页面 `/tests/performance` |

## 第 1 轮：P1（每台机器 11 批）

| 批次 | 内容 | 用例 | Mac 用时 | 测完上线 |
|---|---|---|---|---|
| `xl3` | 动态数组往返、.xlsb 和 .ods、存 CSV、一份真实文件 | XL-RT-04 XL-ODS-01 XL-CSV-02 XL-REAL-04 | 约 60 分钟 | `/tests/excel-files` |
| `xl4` | 另外两份真实文件、上游的检查工具 | XL-REAL-05 XL-REAL-06 FX-CLI-01 | 约 50 分钟 | `/tests/excel-files`、`/tests/formulas` |
| `fx2` | LAMBDA、循环引用、中文函数、中文排序 | FX-LAMBDA-01 FX-ITER-01 ZH-FX-01 ZH-SORT-01 | 约 45 分钟 | `/tests/formulas`、`/zh/tests/chinese` |
| `ed2` | 格式、条件格式、表格 | ED-FMT-01 ED-CF-01 ED-TABLE-01 | 约 50 分钟 | `/tests/editing` |
| `ed3` | 撤销到底、快捷键 | ED-UNDO-01 ED-KEY-01 | 约 45 分钟 | `/tests/editing` |
| `ch2` | 迷你图、图片、切片器、单变量求解和规划求解 | CH-SPARK-01 CH-PIC-01 PV-SLICER-01 PV-WHATIF-01 | 约 55 分钟 | `/tests/editing` |
| `sc2` | GridCraft 宏、作者信息、打印、分页 | SC-MACRO-01 SC-AUTH-01 PR-PRINT-01 PR-BREAK-01 | 约 55 分钟 | `/safety`、`/tests/editing` |
| `mc2` | run 脚本、控制通道、更新提示 | MC-CLI-02 MC-CTL-01 IN-UPD-01 | 约 50 分钟 | `/mcp`、`/download/*` |
| `in2` | 为所有用户安装、便携版、卸载 | IN-INS-02 IN-PORT-01 IN-UNI-01 | 约 40 分钟 | `/download/windows` |
| `st2` | 坏的 XML、解压炸弹、极端公式、文件被占用、只读、多窗口 | ST-BAD-03 ST-BAD-04 ST-BAD-05 ST-SAV-03 ST-SAV-04 ST-WIN-01 | 约 55 分钟 | `/safety`、`/tests` |
| `pf2` | 100 万行、保存、排序 | PF-OPEN-02 PF-SAVE-01 PF-SORT-01 | 约 45 分钟 | `/tests/performance` |

P2 用例不排进批次，有空或读者问到时用 `--add` 加进某一批。

## 回归

每个新版本发布后三天内，在 Mac 上做一个回归批（约 50 分钟），覆盖面最广的五个用例：

| 批次 | 内容 | 用例 | Mac 用时 | 测完上线 |
|---|---|---|---|---|
| `reg` | 新版本回归 | XL-RT-01 XL-RT-03 FX-CORE-01 ZH-SHOW-01 ST-SAV-02 | 约 50 分钟 | 该版本的发布说明页（"实测变化"）；各测试页标出已复测的版本 |

发布说明里提到修了我们判过非 `pass` 的用例时，用 `--add` 把它们加进回归批，或者另排一批，照样控制在 1 小时内。
