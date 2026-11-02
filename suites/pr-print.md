# PR：打印和 PDF

回答"能打印、导出 PDF 吗？分页和 Excel 一样吗？"。依据：README "Page Break Preview, print settings, and PDF export and printing"；`crates/pdf/src/lib.rs`（0.15.0）：PDF 只用 14 种标准字体、不嵌入字体，WinAnsi 以外的字符（包括中文）用不嵌入的 Adobe-GB1 字体写入；上游 issue #446 "PDF export draws non-Latin text as ?"（2026-11-02 仍未关闭）。

中文 PDF 的用例在 ZH-PDF-01。

## 用例

| ID | 优先级 | 系统 | 阶段 | 来源 | 用例 | 怎么测、看什么 |
|---|---|---|---|---|---|---|
| PR-PDF-01 | P0 | Mac Win | E X | G | 导出 PDF | `PR-PAGES.xlsx`：页面布局里设横向、A4、所有列放在一页、第 1 行为打印标题、页脚加页码。导出 PDF 到 `outputs/PR-PAGES.gridcraft.pdf`；Excel 用同样设置导出 `PR-PAGES.excel.pdf`。对比：页数、每页的行范围、标题行是否每页重复、页码；用"预览"和 Chrome 各打开一次 |
| PR-PRINT-01 | P1 | Mac Win | T | G | 打印对话框 | 「文件 › 打印」：能列出系统里的打印机吗，份数、双面、灰度、打印范围；打印到 PDF 打印机（Windows："Microsoft Print to PDF"；Mac：存储为 PDF）。没有打印机时只做到对话框 |
| PR-BREAK-01 | P1 | Mac Win | E S A | G | 分页预览和页面布局视图 | `PR-PAGES.xlsx`：分页预览里拖动分页线、插入手动分页符；页面布局视图里拖页边距；保存后 A：Excel 的分页和 GridCraft 一样吗 |
| PR-HF-01 | P2 | Mac Win | E S A | — | 页眉和页脚 | 页眉左中右三段、首页不同、奇偶页不同、插入日期和文件名代码。导出 PDF 看效果；保存后 A |
