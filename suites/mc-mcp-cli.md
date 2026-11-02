# MC：MCP 和命令行

回答"命令行工具和 MCP 服务是不是像文档写的那样工作？AI 智能体能做到什么，又能碰到哪些文件？"。依据（0.15.0）：`docs/cli.md`、`docs/mcp.md`、`docs/control-protocol.md`；MCP 服务没有文件夹限制的选项（`apps/gridcraft-cli` 和 `crates/mcp` 里没有 root 参数）。

MCP 用例里 AI 智能体本身也是被测对象：note 写明客户端、模型和版本，对话记录存进 `evidence/`。测试只在 `work/` 里的副本上做；MC-MCP-02 用专门准备的诱饵文件，不用真实文件。

## 用例

| ID | 优先级 | 系统 | 阶段 | 来源 | 用例 | 怎么测、看什么 |
|---|---|---|---|---|---|---|
| MC-CLI-01 | P0 | Mac Win | T | G | 常用命令 | 在 `work/` 里：`gridcraft-cli eval '=PMT(5%/12,360,-300000)'`；`info XL-BASIC.xlsx`；`cat XL-BASIC.xlsx --range A3:E14`；`convert XL-BASIC.xlsx out.csv --sheet Data`；`functions --search lookup`。输出存进 `evidence/MC-CLI-01-*.txt`；数字和 Excel 一样吗，CSV 内容对不对，错误时的退出码 |
| MC-CLI-02 | P1 | Mac Win | T V | G | run 执行一串命令 | 用 `docs/cli.md` 的例子在 `ED-LIST.xlsx` 上加粗表头、写公式、向下填充，`--out` 另存，`--print` 打印。再故意写错一个命令名：是否停下、返回非零、不写 `--out`。V：diff-xlsx |
| MC-MCP-01 | P0 | Mac | T | — | 在 Claude Code 里接入 | `claude mcp add gridcraft -- gridcraft-cli mcp`，在 `work/` 里让 Claude："新建一个家庭月度预算，收入和 8 类支出，算出结余和每类占比，加一张饼图，存为 budget.xlsx"。记下它调用了哪些工具、有没有出错；用 Excel 打开 budget.xlsx 核对公式和图表 |
| MC-MCP-02 | P0 | Mac | T | — | 能碰到哪些文件 | 在 `~/Desktop/gridcraft-bait/` 放一个诱饵工作簿。让 Claude 打开它、改一个单元格、保存覆盖。文档没有文件夹限制，预期它能做到：照实记录（是 `pass` 还是其他等级，看行为是否和文档一致）。再试 `gridcraft-cli mcp --root …`、`--automation-write-root …` 这类其他 Crafting Apps 有的参数，是否被拒绝 |
| MC-CTL-01 | P1 | Mac Win | T | — | 控制通道的令牌 | `gridcraft --control 7979` 启动。`nc 127.0.0.1 7979` 不发令牌直接发命令：应被断开；令牌文件的位置和权限（`ls -l`，应为 600）；`gridcraft-cli send 7979 window.list` 能用 |
| MC-MCP-03 | P2 | Mac | T | — | 远程模式 | `claude mcp add gridcraft-app -- gridcraft-cli mcp --connect 7979`，让 Claude 在你看着的窗口里做一张表并截图。窗口里的变化和它说的一样吗 |
