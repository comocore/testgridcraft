# 真实工作簿（R 类）

第三方工作簿不进仓库（`corpus/real/` 被 `.gitignore` 排除），只记来源和 SHA-256。准备批里下载到 `corpus/real/`，用 `shasum -a 256` 核对；和下表不同（网站更新了文件），记下新的 SHA-256 和下载日期，改这张表，在提交说明里写明。

| 代号 | 文件 | 来源 | 下载日期 | SHA-256 | 用于 |
|---|---|---|---|---|---|
| R-XL-1 | `NST-EST2024-POP.xlsx` | https://www2.census.gov/programs-surveys/popest/tables/2020-2024/state/totals/NST-EST2024-POP.xlsx （美国人口普查局各州人口估计，合并单元格的多级表头、千分位格式、脚注；公有领域） | 2026-11-02 | `4a98cd464d3f9e122e74b47b38f1584e78e9077f4b36be1cabf7a79ab5942852` | XL-REAL-01 |
| R-XL-2 | `MER-T01.01.xlsx` | https://www.eia.gov/totalenergy/data/browser/xls.php?tbl=T01.01 （美国能源信息署《月度能源评论》表 1.1，两张表、约 9400 个单元格、按年和按月的数据；公有领域。每月更新，SHA-256 会变） | 2026-11-02 | `a56c2871a16ec7c0db87a6e0ae0d59a999157b32ba24c91048a87cb0e33cf675` | XL-REAL-02 |
| R-XL-3 | 待定 | 一份带公式和图表的公开工作簿，最好是 Excel 做的财务或预算模型（政府机构、大学课程网站发布的模板），至少 3 张工作表、几百个公式、2 张以上图表 | | | XL-REAL-03 |
| R-XL-4 | 待定 | 一份带数据透视表的公开工作簿（例如统计机构发布的、内含透视表的数据文件） | | | XL-REAL-04 |
| R-XL-5 | 待定 | 一份中文的公开 .xlsx：统计公报、政府采购或预算公开的附表，含合并单元格和中文数字格式 | | | XL-REAL-05 |
| R-XL-6 | 待定 | 一份大而复杂的公开工作簿：几十张工作表、条件格式、表格（ListObject）、定义名称，5–20 MB | | | XL-REAL-06 |
| R-ZH-XLS | 待定 | 一份中文的 Excel 97-2003 `.xls`，例如统计局或地方政府网站下载的年度数据表 | | | ZH-XLS-01 |
| R-GS | 自己做 | 在 Google 表格里新建一张表：A1:A10 填数字，B1 `=FILTER(A1:A10,A1:A10>5)`，C1 `=QUERY(A1:A10,"select A where A > 5")`，D1 `=GOOGLETRANSLATE("你好","zh","en")`，然后"文件 › 下载 › Microsoft Excel"，存为 `corpus/made/R-GS.xlsx` | | | XL-GS-01 |

"待定"的几行，在准备批里选定后补全这张表；选不到合适的，对应用例记 `skip` 并写原因。选文件时不要选含个人信息的。
