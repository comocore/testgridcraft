# 测试文件

| 目录或文件 | 内容 | 进仓库 |
|---|---|---|
| `generated/` | `tools/make-corpus.mjs` 生成的探针工作簿和 CSV（G 类），以及 `manifest.json` | 否，随时可以重新生成 |
| `made/` | 按 [recipes.md](recipes.md) 在 Excel 里做的文件（M 类）：参考副本 `*.excel.xlsx`、综合工作簿、加密和带 VBA 的文件 | 是（加密文件的密码写在 recipes.md 里，它们只是测试用的假数据） |
| `real/` | R 类第三方真实工作簿，按 [real-world.md](real-world.md) 下载 | 否，只记下载地址和 SHA-256 |
| `baseline.yaml` | 用 Excel 核对探针的记录：哪个文件、哪个 Excel 版本、通过与否、有什么问题 | 是 |

## 生成

```sh
node tools/make-corpus.mjs         # 全部，约 3 秒，约 7 MB
node tools/make-corpus.mjs --big   # 另加 100 万行的 PF-ROWS-1M.xlsx（约 40 MB）
```

同样的脚本、同样的 Node 版本，每次生成的文件逐字节相同（ZIP 里的时间戳固定为 2026-01-01）；`manifest.json` 记着每个文件的大小、SHA-256 和用途。填充用的数据由脚本按固定种子生成，姓名是常见的中文名，金额和日期都是编的，不含任何第三方数据。

生成的工作簿里，公式只写公式、不写缓存值，并设了"打开时完整重算"（`fullCalcOnLoad`），所以 Excel 打开时会自己算一遍。比较计算结果用的标准答案，是 Excel 算完后另存的副本（recipes.md §1）。2007 年以后加入 Excel 的函数（CONCAT、IFS、STDEV.S、SEQUENCE……）按文件格式的要求写成 `_xlfn.` 前缀，不然 Excel 会显示 `#NAME?`。

| 文件 | 内容 | 用于 |
|---|---|---|
| `XL-BASIC.xlsx` | 合并单元格、冻结、自动筛选、定义名称、超链接、条件格式、下拉列表、十几种数字格式和单元格样式、一个隐藏的工作表 | XL-RT-01、XL-RT-02、ED-*、SC-AUTH-01 |
| `FX-CORE.xlsx` | 69 个单格公式：数学、精度、统计、逻辑、文本、日期、财务、查找、错误值 | FX-CORE-01 |
| `FX-LEDGER.xlsx` | 40 行流水账和一张用 SUMIFS、COUNTIF、MAXIFS、INDEX/MATCH、SUMPRODUCT 汇总的表 | FX-LEDGER-01、FX-AUDIT-01 |
| `ZH-TEXT.xlsx` | 中文姓名、自动换行的长中文、中英文混排、宋体、全角标点、繁体、日文韩文、生僻字、中文日期格式、¥ 会计格式；一张中文函数表 | ZH-SHOW-01、ZH-FX-01、ZH-PDF-01、XL-CSV-02 |
| `ZH-CSV-UTF8-BOM.csv`、`ZH-CSV-UTF8.csv`、`ZH-CSV-GBK.csv` | 同一张中文表的三种编码 | ZH-CSV-01 |
| `ZH-IDS.csv` | 编造的 18 位身份证号（含 X 结尾）、0 开头的号码、带前导零的工号、19 位卡号 | ZH-IDS-01 |
| `CSV-DATES-*.csv` | 日在前、月在前、两可、混合、年在前和中文写法的日期 | XL-CSV-01 |
| `ED-LIST.xlsx` | 200 行名单（编号、中文姓名、地区、日期、金额、数量）和一行 SUBTOTAL 合计 | ED-*、CH-*、PV-*、FX-DYN-01、ZH-SORT-01 |
| `PR-PAGES.xlsx` | 150 行、打印时跨好几页的报表，每 10 行一句中文 | PR-* |
| `SC-LINKS.xlsx` | 各种超链接（https、http、mailto、file、UNC、自定义协议、javascript）和 HYPERLINK、WEBSERVICE、FILTERXML | SC-LINK-01、IN-NET-01 |
| `PF-ROWS-100K.xlsx`、`PF-CALC-20K.xlsx`、`PF-WIDE.xlsx`、`PF-ROWS-1M.xlsx` | 10 万行数据、2 万行连锁公式、1000 列、100 万行（`--big`） | PF-*、ST-SAV-02 |
| `ST-BAD-*.xlsx` | 截断、空、不是 ZIP、XML 坏掉、解压后约 400 MB 的工作表、极端公式 | ST-BAD-* |

## 用 Excel 核对

准备批里用 Excel 逐个打开 `generated/` 里的 .xlsx（`ST-BAD-*` 和 `PF-*` 除外），对照文件里写的预期，并确认 Excel 没有提示修复。和预期不符的，说明探针本身有问题，在 `baseline.yaml` 里记下来，这个文件先不用，告诉编辑去修生成脚本：

```yaml
excel: Microsoft Excel for Mac 16.103
date: 2026-11-03
files:
  - file: XL-BASIC.xlsx
    ok: true
  - file: ZH-TEXT.xlsx
    ok: false
    note: Excel 提示修复，修复报告说删除了某某部分
```

## 公开和许可

`made/` 里的文件，生成脚本的输出，`runs/` 里的记录，以及方案和用例文档，以 [CC BY 4.0](https://creativecommons.org/licenses/by/4.0/deed.zh-hans) 公开（[LICENSE-CC-BY-4.0](../LICENSE-CC-BY-4.0)），署名写"GetGridCraft test corpus, getgridcraft.com"。`tools/` 里的脚本以 MIT 或 Apache-2.0 双许可公开（[LICENSE-MIT](../LICENSE-MIT)、[LICENSE-APACHE](../LICENSE-APACHE)），和 GridCraft 本身一样，任选其一。

第三方真实工作簿不放进仓库、不提供下载，只公开它的来源地址、下载日期和 SHA-256，读者可以自己下载同一个文件来复现。美国政府机构发布的文件属于公有领域，以后可以随测试数据一起提供。

私有样本（工资表、客户名单等）永远不公开，截图只留在本地。
