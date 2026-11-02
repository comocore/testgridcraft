# IN：下载、安装和网络

回答"装起来顺不顺利？签名对不对？它会不会联网，联到哪里？"。依据（0.15.0）：README "Downloads"；`packaging/windows/gridcraft.wxs` 和 `INSTALLATION.md`（0.13.0 起默认为当前账户安装）；`packaging/macos/Info.plist.in`（macOS 11 起）；更新检查在 `apps/gridcraft/src/main.rs`（`start_update_check`）和 craft-libs 的 `crates/update`；README "no telemetry"。

网站编辑在 2026-11-02 核对过 0.15.0 的签名（getgridcraft.com 的安全页）；这里由测试者在自己的机器上再做一遍，看到的是读者实际会看到的。

网络监视：Mac 用 [LuLu](https://objective-see.org/products/lulu.html)（免费、开源，每个新连接都会弹窗询问），Windows 用"资源监视器 › 网络"。版本写进 `run.yaml` 的 `tools`。

## 用例

| ID | 优先级 | 系统 | 阶段 | 来源 | 用例 | 怎么测、看什么 |
|---|---|---|---|---|---|---|
| IN-DL-01 | P0 | Mac Win | T | — | 下载和校验 | 从 GitHub 发布页下载安装包和 `SHA256SUMS.txt`，`shasum -a 256 -c`（Windows：`Get-FileHash`）核对；记下下载用时和网络环境（中国大陆网络下 GitHub 下载是否很慢，照实记） |
| IN-SIG-01 | P0 | Mac | T | — | Mac 签名和公证 | 挂载 DMG，`spctl -a -vv`、`codesign -dv --verbose=4`、`xcrun stapler validate`；拖到"应用程序"后第一次打开时 macOS 显示的对话框（截图） |
| IN-SIG-02 | P0 | Win | T | — | Windows 签名和 SmartScreen | .msi 属性里的数字签名；双击后 SmartScreen 是否拦截、显示的发布者 |
| IN-INS-01 | P0 | Win | T | — | 只为我安装 | 默认选项安装：是否要管理员权限；装到哪里；开始菜单；桌面快捷方式默认不勾；.xlsx 的"打开方式"里有没有 GridCraft，默认程序有没有被改 |
| IN-INS-02 | P1 | Win | T | — | 为所有用户安装 | 先卸载，再选 Everyone on this computer：要管理员、装到 Program Files；然后再运行一次安装包选"只为我"：应被拒绝并说明原因 |
| IN-PORT-01 | P1 | Win | T | — | 便携版 | 解压运行：设置和日志写到哪里（`%APPDATA%\GridCraft`）；不写注册表的打开方式 |
| IN-1ST-01 | P0 | Mac Win | T | — | 第一次启动 | 全新设置下启动：界面语言、浅色深色是否跟随系统、有没有欢迎页或提示、「文件 › 新建」里的示例工作簿能否打开 |
| IN-NET-01 | P0 | Mac Win | T | G | 网络连接 | 开着网络监视启动 GridCraft：连了哪些地址（预期只有 api.github.com 查更新）；新建、编辑、保存一个工作簿、打开 `SC-LINKS.xlsx` 时有没有新连接；断网后启动：能正常用，有没有报错弹窗 |
| IN-UPD-01 | P1 | Mac Win | T | — | 更新提示 | 装上一个旧版本（例如 0.14.2），启动：是否提示更新；选"稍后"；再选"更新并重启"：更新后版本号、设置是否保留；有未保存的工作簿时是否先问 |
| IN-UNI-01 | P1 | Mac Win | T | — | 卸载 | 卸载后留下了什么：设置文件夹、日志、注册表的打开方式（Windows） |
| IN-WEB-01 | P2 | Mac | T | — | 自己部署网页版 | 解压 `gridcraft-web-<版本>.zip`，`python3 -m http.server` 起一个本地服务，用 Chrome 打开：示例能用吗，打开和保存文件；开发者工具的"网络"里有没有请求第三方地址 |
