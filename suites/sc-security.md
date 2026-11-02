# SC：密码、宏、链接和隐私

回答"设了密码的文件能打开吗，保存后密码还在吗？工作表保护有用吗？文件里的链接和宏会不会乱来？"。依据（0.15.0 源代码）：

- 加密：`crates/engine/src/io.rs` 的 `DECRYPTED_NOTICE`（"This workbook was password-protected. GridCraft saves it without the password."）；`crates/engine/src/cmd/file.rs`：解密后的工作簿没有保存目标，"Save never replaces the encrypted original with an unprotected copy"；0.14.2 修复了 Excel 加密的文件打不开的问题（PR #2870）。
- 链接：`crates/ui-egui/src/dialogs.rs` 的 `openLink` 对话框；`link.open` 只接受 http、https、mailto。
- 宏：`docs/macros.md` "A macro that touches files, or that came with a workbook opened from a file, runs once the person approves it"；脚本在 `ScriptSandbox` 里运行，不能访问文件。
- VBA：README 和 `docs/target-app-parity.md` "an .xlsm keeps its VBA project unchanged on save and never runs it"。
- 作者：0.14.2 起新工作簿的作者取偏好设置里的用户名，默认 "GridCraft User"（PR #2726）。
- 网络请求和更新检查在 [in-install.md](in-install.md) 的 IN-NET-01。

隐私问题（存出的文件里留下了不该留的东西，比如本该加密却没加密、作者名泄露了真实姓名）加 `leak` 标记。

## 用例

| ID | 优先级 | 系统 | 阶段 | 来源 | 用例 | 怎么测、看什么 |
|---|---|---|---|---|---|---|
| SC-ENC-01 | P0 | Mac Win | O D S V | M | Excel 加密的工作簿 | `SC-ENC.xlsx` 和 `SC-ENC-ZH.xlsx`（recipes.md §9，密码 `grid-0.15` 和 `表格密码`）。O：先输错一次密码，提示原文；再输对。D：内容和 Excel 里一样；有没有说明"保存时不带密码"的提示。S：按 ⌘S / Ctrl+S，是否要求另存到新位置，而不是覆盖原文件。V：`shasum -a 256` 确认原文件没变；`diff-xlsx` 能直接读另存的文件，说明它没有加密：这是文档写明的行为，`pass`；但如果原文件被覆盖，`fail` 加 `leak` |
| SC-PROT-01 | P0 | Mac Win | O E S A | M | 工作表保护 | `SC-PROT.xlsx`（recipes.md §10，密码 `prot`，Data 表的 B4:B13 未锁定）：锁定的单元格打字时立即提示、不能改（0.15.0 起）；B4:B13 能改；不知道密码不能撤销保护，输对密码可以；重新保护后保存。A：Excel 里保护还在，原密码有效 |
| SC-LINK-01 | P0 | Mac Win | E | G | 文件里的链接 | `SC-LINKS.xlsx`：逐个点 B2 到 B9。https、http、mailto 和 HYPERLINK 应先显示地址、问你是否打开；file、UNC、ms-settings、javascript 应被拒绝，没有任何程序被打开。WEBSERVICE 和 FILTERXML 显示 `#VALUE!`。同时开着 IN-NET-01 的网络监视，记下点击时有没有连接 |
| SC-VBA-01 | P0 | Mac Win | O S A | M | 带 VBA 的 .xlsm | `SC-VBA.xlsm`（recipes.md §11，Workbook_Open 和一个按钮各弹一个消息框）。O：GridCraft 打开时有没有提示，消息框绝不能弹出。改一个单元格，另存为 `.xlsm`。A：Excel 打开后宏还在、能运行（启用宏后消息框出现） |
| SC-MACRO-01 | P1 | Mac Win | E S R | — | GridCraft 宏的确认 | 录制一个宏：输入一行数据并另存为 CSV（会写文件的一步）。保存工作簿、关闭、重新打开，运行这个宏：应要求确认；拒绝后什么都没写。再录一个不碰文件的宏（只改格式），重开后运行：是否也要确认（随文件来的宏都要确认） |
| SC-AUTH-01 | P1 | Mac Win | S V | — | 作者和文件属性 | 新建工作簿保存，`diff-xlsx` 或解压后看 `docProps/core.xml`：creator 和 lastModifiedBy 是什么，有没有电脑用户名或真实姓名；打开 `XL-BASIC.xlsx`（作者 "GetGridCraft test corpus"）另存，作者是否保留。在偏好设置里改用户名后再试 |
