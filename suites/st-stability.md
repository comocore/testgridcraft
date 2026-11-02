# ST：稳定性和数据安全

回答"会不会崩溃、丢数据？坏文件会不会把它弄崩？"。依据（0.15.0）：仓库的 `CLAUDE.md` "Never crash … People trust GridCraft with their numbers; a crash loses their work"；`crates/engine/src/io.rs` 的 `write_file`：先写 `<文件名>.sctmp` 临时文件，再改名替换原文件；GridCraft 没有崩溃恢复，"自动保存"默认关闭，打开后只对已存成 .xlsx 的工作簿生效（`crates/ui-egui/src/app/files.rs`）；未关闭的崩溃 issue，例如 #3092（Windows 关闭窗口时崩溃）、#2042 和 #2839（macOS 文件对话框打开时崩溃）。

所有用例都用文件的副本做。

## 用例

| ID | 优先级 | 系统 | 阶段 | 来源 | 用例 | 怎么测、看什么 |
|---|---|---|---|---|---|---|
| ST-BAD-01 | P0 | Mac Win | O | G | 截断的文件 | `ST-BAD-TRUNC.xlsx`：应给出看得懂的错误，不崩溃、不卡住。提示原文照抄；Excel 的表现也记下 |
| ST-BAD-02 | P0 | Mac Win | O | G | 空文件和假文件 | `ST-BAD-EMPTY.xlsx`（0 字节）和 `ST-BAD-NOTZIP.xlsx`（文本文件）：各记一次提示原文 |
| ST-BAD-03 | P1 | Mac Win | O | G | 坏的 XML | `ST-BAD-XML.xlsx`：不崩溃；能打开的话显示了什么 |
| ST-BAD-04 | P1 | Mac Win | O M | G | 解压炸弹 | `ST-BAD-BOMB.xlsx`（一个工作表解压后约 400 MB）：不崩溃、不卡死；M 记打开过程中的内存峰值（MB） |
| ST-BAD-05 | P1 | Mac Win | O | G | 极端公式 | `ST-BAD-DEEP.xlsx`：3000 层括号、1000 层 SUM、整张表求和、170 亿格的 SEQUENCE、超长 REPT。应给出错误值或结果，不崩溃，30 秒内能操作 |
| ST-SAV-01 | P0 | Mac | T | G | 磁盘满时保存 | 磁盘工具 › 文件 › 新建映像 › 空白映像，20 MB，APFS。把 `XL-BASIC.xlsx` 复制进去，再用大文件把映像塞到只剩约 1 MB。在 GridCraft 里打开映像里的文件，插入一个 `PF-WIDE.xlsx` 的工作表内容（复制粘贴）后保存：应提示空间不足；原文件仍能用 Excel 正常打开；映像里不留 `.sctmp` 文件 |
| ST-SAV-02 | P0 | Mac Win | T | G | 保存时强制退出 | 打开 `PF-ROWS-100K.xlsx` 的副本，改一个格，按保存后立刻强制退出（Mac：⌥⌘Esc；Windows：任务管理器）。原文件完好（Excel 能打开），旁边有没有留下 `.sctmp`。做 3 次 |
| ST-SAV-03 | P1 | Win | T | G | 文件正被 Excel 打开 | Excel 打开一个副本，GridCraft 也打开同一个文件，改后保存：各有什么提示，最后文件里是谁的内容 |
| ST-SAV-04 | P1 | Mac Win | T | G | 只读文件和没有写权限的文件夹 | 副本设成只读后保存；保存到没有写权限的文件夹。应有清楚的提示，不能悄悄失败 |
| ST-REC-01 | P0 | Mac Win | T | G | 崩溃时没保存的修改 | a：打开 `XL-BASIC.xlsx` 的副本，改三处，不保存，强制退出，再打开 GridCraft：文档说没有恢复，确认是不是这样（没有恢复是文档写明的行为，照实记，不判 `fail`；但如果它声称恢复而内容不对，`fail`）。b：打开标题栏的"自动保存"，改一处，等 5 秒，强制退出：重新打开后这处修改在不在 |
| ST-WIN-01 | P1 | Mac Win | T | G | 多窗口和关闭 | 同一个工作簿开两个窗口（「视图 › New Window」），各改一处，关掉一个窗口、再关最后一个：什么时候问保存；Windows 上关闭窗口时有没有崩溃（issue #3092）；退出后进程是否还在 |
| ST-LOG-01 | P0 | Mac Win | T | — | 整轮测试的崩溃记录 | 不是单独做的用例。每一批里每一次崩溃、卡死（30 秒没反应）都在这一条的 `events` 里记一行：时间、在做哪个用例、操作、日志文件名。崩溃后马上复制日志（`gridcraft.log` 和 `gridcraft.1.log`） |
