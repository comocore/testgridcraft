# 测试记录

每一批测试一个目录：`<日期>-<mac|win>-<GridCraft 版本>-<批次>`，如 `2026-11-03-mac-0.15.0-xl1`。一批只测一个版本、一台机器，批次和用例见 [batches.md](../batches.md)。

```sh
# 一批的记录；--from 带上上一批的机器、参考软件和字体信息
node tools/new-run.mjs --os mac --version 0.15.0 --batch xl2 --from runs/2026-11-03-mac-0.15.0-xl1
# 回归批再加几个用例
node tools/new-run.mjs --os mac --version 0.16.0 --batch reg --add FX-DYN-01,ZH-IME-01
# 检查并汇总，同时列出要复判的用例
node tools/check-run.mjs runs/2026-11-03-mac-0.15.0-xl1 --sample 10
```

不带 `--batch` 时，还可以按套件和优先级生成（`--priority P0 --suites xl,fx`），目录名里没有批次，只在批次以外临时加测时用。

## 目录里有什么

| 文件或目录 | 内容 |
|---|---|
| `run.yaml` | 批次和环境：系统、机器、GridCraft 安装包和 SHA-256、Excel 的版本和语言、中文字体、工具。开始前填好，中途的变化写进 `changes` |
| `xl.yaml` … `st.yaml` | 本批涉及的套件各一个文件，已列出本机要测的用例，只要填空。每批都有 `st.yaml` 里的 ST-LOG-01，记这一批里的崩溃和卡死 |
| `work/` | 测试文件的副本，在这里操作。不提交（`.gitignore`） |
| `evidence/` | 截图、录屏、`diff-xlsx.mjs` 等检查工具的输出。文件名规则见 README.md §6 |
| `outputs/` | GridCraft 保存或导出的文件（`XL-BASIC.gridcraft.xlsx` 等），以及 Excel 为对照存的文件（`*.excel.*`） |
| `logs/` | 复制出来的 GridCraft 日志 |

## 一条记录

```yaml
  - id: XL-RT-01
    p: P0
    title: "不改动直接往返"
    O: pass
    D: minor
    S: pass
    A: pass
    V: pass
    diff: XL-RT-01-V-diff.txt
    flags: []
    note: "D：'格式 Formats' 表 B14 自动换行后行高比 Excel 矮一行，文字被截掉最后两个词。A：Excel 打开无修复提示，三个表和原文件一致。V：只有 _xlfn 写法和 docProps 的差别，Excel 自己往返也有。"
    evidence: [XL-RT-01-D-gridcraft.png, XL-RT-01-D-excel.png, XL-RT-01-V-diff.txt]
    issue: ""
    checked: 2026-11-03
```

（上面的内容只是格式示例，不是测试结果。）

- 每个阶段填一个等级：`pass minor major fail na blocked skip`。没测的留空。
- `flags` 可填 `wrongValue`（数和 Excel 不同）、`dataLoss`（保存或往返后内容丢了或变了，没有提示）、`repair`（Excel 打开 GridCraft 存的文件时要求修复）、`crash`（崩溃或 30 秒没反应）、`leak`（文件里留下了不该留的东西）、`warning`（GridCraft 打开或保存时给出提示）。
- 任何 `minor`、`major`、`fail`、`blocked`、`skip` 或标记，`note` 必填：写清是哪个工作表、哪个单元格、什么样，让没看截图的人也明白。
- `diff`：有 V 阶段的用例，写 `diff-xlsx.mjs` 输出文件的名字。有 `dataLoss` 或 `wrongValue` 标记时必填。
- `evidence` 写 `evidence/` 里的文件名。有差异的用例至少放一对截图（GridCraft 和 Excel）。私有样本（`-PRIV-`）不放任何截图。
- `issue`：向上游报告后填 issue 的地址；私下报告的安全问题写 `private report`。
- `checked`：判定的日期。

性能（PF）的记录是测量值：

```yaml
  - id: PF-OPEN-01
    p: P0
    title: "打开 10 万行"
    unit: "s"
    gridcraft: [1.8, 1.5, 1.5, 1.6, 1.5]
    excel: [2.4, 2.1, 2.0, 2.2, 2.1]
    other: []
    other_app: ""
    M:
    flags: []
    note: "终点：第一屏数据显示、可以选中单元格。第一次是冷打开。"
    evidence: [PF-OPEN-01-gridcraft-1.mov]
    issue: ""
    checked: 2026-11-05
```

（上面的数字只是格式示例，不是测量结果。）

## 复判

`--sample 10` 按本批的目录名作种子，从已判定的用例里随机列出 10%（至少 1 个）。在之后某一批的开头、至少隔一天，不看原记录，重新判定，写进原来那一批记录的 `recheck`：

```yaml
    recheck: { date: 2026-11-04, D: major }
```

两次不同时，再看一遍，在 note 里写"复判：minor → major，原因……"，并把阶段的等级改成最终结果。页面上报告复判的一致率。

## 提交

每批测完运行一次 `check-run.mjs`，没有 error 再提交。截图用 PNG；录屏超过 50 MB 的不进仓库，放在外部存储，在 note 里写位置。没修好的安全问题，提交前把 note 和 evidence 里的细节换成"已私下报告"，原始材料留在本地，修复发布后再补上（README §10）。
