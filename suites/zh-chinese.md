# ZH：中文

回答中文读者最先问的几件事："界面有中文吗？能用中文输入法吗？中文表格显示正常吗？中文 CSV 会不会乱码？身份证号会不会变？导出的 PDF 里中文还在吗？"。依据（0.15.0）：

- 界面：`crates/ui-egui/src/i18n/zh.rs` 约 440 条翻译，覆盖功能区和部分菜单，对话框里的内容（如「设置单元格格式」的各页、「查找和替换」）大多仍是英文；`ROADMAP.md` 写简体中文 "menus only, ~12%"，比代码落后。
- 输入法：`docs/localization-parity.md` "IME input goes through egui/winit and is unverified for Chinese, Japanese and Korean"；issue #2345（Wayland 上开始编辑时第一个按键绕过输入法）。
- 显示：issue #142 "Chinese text in files is not supported"、#2110（中文行高）、#2621（macOS 上中英文混排上下错位）、#1497（Arch Linux 装了字体仍显示方框）、#2330（编辑栏里中文的排版），2026-11-02 都还开着。
- CSV：`docs/file-format-parity.md` "other invalid UTF-8 remains Windows-1252"，也就是说 GBK 编码的 CSV 很可能会被当成西欧编码读成乱码。
- PDF：见 [pr-print.md](pr-print.md) 的依据。

**测试机的系统语言用简体中文**（README §3.3），GridCraft 第一次启动时会跟随系统语言。

## 用例

| ID | 优先级 | 系统 | 阶段 | 来源 | 用例 | 怎么测、看什么 |
|---|---|---|---|---|---|---|
| ZH-UI-01 | P0 | Mac Win | T | — | 中文界面 | 全新设置下第一次启动：界面是不是中文。截图"开始""插入""公式""数据"四个选项卡，以及「设置单元格格式」「排序」「查找和替换」「选项」四个对话框，在 note 里写每处还是英文的地方（大概数一下）。在「视图 › 界面语言」里切换语言是否马上生效 |
| ZH-IME-01 | P0 | Mac Win | E | — | 中文输入法 | 用系统自带的拼音输入法：选中单元格直接打"北京"（不按 F2），第一个字母有没有丢、候选框在不在光标旁；在编辑栏里打字；双击单元格在中间插入中文；在「查找」框、工作表改名、批注里打中文；打中文标点。Mac 上 Windows 上各做一次 |
| ZH-SHOW-01 | P0 | Mac Win | O D | G | 中文显示 | `ZH-TEXT.xlsx` 和 Excel 并排：字体是否正常（不是方框或问号）、自动换行的行高、中英文混排是否上下对齐、全角标点、繁体、日文韩文、生僻字、"宋体"那一格用了什么字体替换（Mac 上没有宋体）、`yyyy年m月d日` 和 ¥ 会计格式 |
| ZH-CSV-01 | P0 | Mac Win | O D | G | 中文 CSV | 依次打开 `ZH-CSV-UTF8-BOM.csv`、`ZH-CSV-UTF8.csv`、`ZH-CSV-GBK.csv`。哪个能正确显示中文，哪个乱码；乱码时有没有提示、有没有办法选编码。Excel 打开同样三个文件对照（Excel 对不带 BOM 的 UTF-8 也常乱码，照实记） |
| ZH-IDS-01 | P0 | Mac Win | O D | G | 身份证号和长数字 | 打开 `ZH-IDS.csv`：身份证号是否变成科学计数法、后三位是否变成 0；以 X 结尾的那个；手机号 0 开头的那个；工号 000123 的前导零；银行卡号。Excel 打开同一个文件对照。再试：有没有办法按文本导入（GridCraft 0.15.0 的 `file.open` 有 `csvAsText` 参数，界面里有没有对应的选项） |
| ZH-PDF-01 | P0 | Mac Win | X | G | 导出中文 PDF | `ZH-TEXT.xlsx` 导出 PDF 到 `outputs/ZH-TEXT.gridcraft.pdf`，分别用"预览"（Mac）、Adobe Acrobat Reader、Chrome、WPS（有的话）打开：中文显示了吗，是方框、问号还是空白；能不能选中和复制中文 |
| ZH-FX-01 | P1 | Mac Win | O D V | M | 中文函数和格式 | `ZH-TEXT.excel.xlsx`（recipes.md §1）的"函数"表：LENB、`TEXT(…,"aaaa")`、`[DBNum1]`、`[DBNum2]` 和 Excel 的结果对比；V：`diff-xlsx --values` |
| ZH-SORT-01 | P1 | Mac Win | E | G | 中文排序 | `ED-LIST.xlsx` 按 Name 升序：GridCraft 的顺序和 Excel 的默认顺序（拼音）一样吗；Excel 的"笔划排序"选项 GridCraft 有没有 |
| ZH-XLS-01 | P2 | Win | O D | R | 中文的老 .xls | real-world.md 的 R-ZH-XLS（政府网站下载的中文 .xls）：中文是否正常，数字格式和合并单元格对不对 |
