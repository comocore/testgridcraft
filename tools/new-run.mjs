#!/usr/bin/env node
// Creates the record files for one test run: runs/<date>-<os>-<version>[-<batch>]/.
//
//   node tools/new-run.mjs --os mac --version 0.15.0 --batch xl2 --from runs/2026-11-03-mac-0.15.0-xl1
//   node tools/new-run.mjs --os mac --version 0.16.0 --batch reg --add FX-DYN-01,ZH-IME-01
//   node tools/new-run.mjs --os win --version 0.15.0 --priority P0 --suites xl,fx
//
// --batch lists the cases of one batch in batches.md (and ignores --priority and --suites);
// --add puts more cases into it. --from copies the machine, reference and font details from an earlier run.
// Without --batch, --priority P0 lists only P0 cases, P1 lists P0 and P1, and the default lists all.

import { mkdir, writeFile, readFile, access } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { parse, parseDocument } from 'yaml';
import { SUITES, loadBatches, loadCases, unitOf } from './lib/cases.mjs';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const arg = (name, fallback) => {
  const i = process.argv.indexOf(`--${name}`);
  return i > 0 ? process.argv[i + 1] : fallback;
};

const os = arg('os');
const version = arg('version');
if (!['mac', 'win'].includes(os) || !version) {
  console.error(
    'usage: new-run.mjs --os mac|win --version <GridCraft version> [--batch <id> [--add ID,...]] [--from <earlier run>] [--date YYYY-MM-DD] [--priority P0|P1|P2] [--suites xl,fx,...]',
  );
  process.exit(2);
}
const date = arg('date', new Date().toISOString().slice(0, 10));
const maxPriority = Number(arg('priority', 'P2').slice(1));
const batchId = arg('batch');
const batch = batchId && (await loadBatches(join(root, 'batches.md'))).get(batchId);
if (batchId && !batch) {
  console.error(`no batch "${batchId}" in batches.md`);
  process.exit(2);
}
const added = arg('add', '').split(',').filter(Boolean);
const from = arg('from');
const prev = from && parse(await readFile(join(from, 'run.yaml'), 'utf8'));
const suites = batch ? Object.keys(SUITES) : arg('suites', Object.keys(SUITES).join(',')).split(',');
const runId = `${date}-${os}-${version}${batch ? `-${batch.id}` : ''}`;
const dir = join(root, 'runs', runId);

try {
  await access(dir);
  console.error(`${dir} already exists`);
  process.exit(1);
} catch {}

const q = (s) => JSON.stringify(s ?? '');
const all = await loadCases(join(root, 'suites'));
for (const id of added) {
  if (!all.some((c) => c.id === id)) {
    console.error(`--add: no case ${id} in suites`);
    process.exit(2);
  }
}
// Every batch that uses GridCraft also keeps the crash log of ST-LOG-01.
const inBatch = batch && new Set([...batch.cases, ...added, ...(batch.cases.length ? ['ST-LOG-01'] : [])]);
const cases = all.filter((c) =>
  c.systems.includes(os) && (inBatch ? inBatch.has(c.id) : suites.includes(c.suite) && Number(c.priority.slice(1)) <= maxPriority),
);

await mkdir(join(dir, 'evidence'), { recursive: true });
await mkdir(join(dir, 'outputs'), { recursive: true });
await mkdir(join(dir, 'logs'), { recursive: true });

const template = `# 本轮测试的环境。开始前填好，中途有变化记在 changes 里。说明见 runs/README.md。
run: ${runId}
batch: ${batch ? batch.id : '""'}
started: ${date}
finished:
tester: ""              # 内部代号；页面上只写 GetGridCraft editors
machine:
  os: ""                # 例：macOS 26.6.2 (25G83)；Windows 11 Pro 24H2 (26100.4061)
  model: ""             # 例：Mac mini (M4, 2024)
  cpu: ""
  memory_gb:
  gpu: ""
  display: ""           # 分辨率和缩放，例：3840×2160，缩放 200%
  system_language: ""
  region: ""            # 区域格式（日期、数字的写法），例：中国大陆
  input_methods: []     # 例：["简体拼音", "ABC"]
  power: ""             # 例：接电源，低电量模式关
gridcraft:
  version: ${version}
  installer: ""         # 文件名
  installer_sha256: ""
  cli: ""
  cli_sha256: ""
  install_path: ""
  process_name: ""      # 活动监视器或任务管理器里显示的名称
  ui_language: ""
  settings: fresh       # fresh：开始前设置文件夹是空的
  autosave: off         # 标题栏的"自动保存"开关
reference:
  excel: ""             # 例：Microsoft Excel for Mac 16.103（Microsoft 365，"Excel › 关于 Excel"里的版本）；用网页版写 Excel for the web 和日期
  excel_language: ""    # Excel 的界面语言，例：简体中文
  libreoffice: ""       # 例：LibreOffice 26.2.1；没有就留空
  others: []            # 例：["WPS Office 12.1", "Numbers 14.4"]
fonts:
  chinese_installed: [] # 系统里装的中文字体，例：["苹方", "宋体-简", "华文黑体"]
  extra_installed: []   # 为测试另装的字体：应该为空
tools: []               # 例：["Node 26.9.0", "LuLu 3.1.5"]
changes: []             # 例：["2026-11-04 Excel 自动更新到 16.104"]
`;

// --from: the machine doesn't change between batches, so start from what the last run recorded.
// The installer and its checksum carry over only when the version is the same.
let runYaml = template;
if (prev) {
  const doc = parseDocument(template);
  const keep = [['tester'], ['machine'], ['reference'], ['fonts'], ['tools'], ['gridcraft', 'install_path'], ['gridcraft', 'process_name'], ['gridcraft', 'ui_language'], ['gridcraft', 'autosave']];
  if (prev.gridcraft?.version === version) keep.push(['gridcraft', 'installer'], ['gridcraft', 'installer_sha256'], ['gridcraft', 'cli'], ['gridcraft', 'cli_sha256']);
  for (const path of keep) {
    const value = path.reduce((o, k) => o?.[k], prev);
    if (value !== undefined && value !== null) doc.setIn(path, value);
  }
  runYaml = doc.toString();
}
await writeFile(join(dir, 'run.yaml'), runYaml);

for (const suite of suites) {
  const list = cases.filter((c) => c.suite === suite);
  if (!list.length) continue;
  const lines = [
    `# ${suite.toUpperCase()}：${SUITES[suite].name} · ${runId}`,
    `# 用例和步骤见 suites/${SUITES[suite].file}；等级、标记和证据的规则见 README.md §5–§6。`,
    '# 等级：pass minor major fail na blocked skip    标记：dataLoss wrongValue repair crash leak warning',
    `suite: ${suite}`,
    'cases:',
  ];
  for (const c of list) {
    lines.push(`  - id: ${c.id}`, `    p: ${c.priority}`, `    title: ${q(c.title)}`);
    if (suite === 'pf') {
      lines.push(`    unit: ${q(unitOf(c.id))}`, '    gridcraft: []', '    excel: []', '    other: []', '    other_app: ""', '    M:');
    } else {
      for (const s of c.stages) lines.push(`    ${s}:`);
    }
    if (c.id === 'ST-LOG-01') lines.push('    events: []    # - { time: "2026-11-03 14:05", case: XL-RT-03, action: "", log: logs/… }');
    if (c.stages.includes('V')) lines.push('    diff: ""      # tools/diff-xlsx.mjs 的输出文件，放在 evidence/');
    lines.push('    flags: []', '    note: ""', '    evidence: []', '    issue: ""', '    checked:');
  }
  await writeFile(join(dir, `${suite}.yaml`), `${lines.join('\n')}\n`);
  console.log(`${suite}.yaml  ${list.length} cases`);
}
console.log(`→ ${dir}`);
