#!/usr/bin/env node
// Checks the records of a test run and prints a summary.
//
//   node tools/check-run.mjs runs/2026-11-03-mac-0.15.0-xl1
//   node tools/check-run.mjs <run> --sample 10     also list 10% of the judged cases to judge again
//   node tools/check-run.mjs <run> --strict        exit with an error on warnings too
//
// Errors are records that break the rules in README.md §5–§7; warnings are things to finish.

import { readFile, readdir, access } from 'node:fs/promises';
import { basename, dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { parse } from 'yaml';
import { FLAGS, SUITES, VERDICTS, loadCases, rng } from './lib/cases.mjs';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const runDir = process.argv[2];
if (!runDir) {
  console.error('usage: check-run.mjs <run folder> [--sample <percent>] [--strict]');
  process.exit(2);
}
const sampleAt = process.argv.indexOf('--sample');
const samplePct = sampleAt > 0 ? Number(process.argv[sampleAt + 1] ?? 10) : 0;
const strict = process.argv.includes('--strict');

const errors = [];
const warnings = [];
const exists = (p) => access(p).then(() => true, () => false);
const catalogue = new Map((await loadCases(join(root, 'suites'))).map((c) => [c.id, c]));

const load = async (file) => {
  try {
    return parse(await readFile(file, 'utf8'));
  } catch (e) {
    console.error(`${basename(file)}: ${e.message}`);
    process.exit(1);
  }
};

const run = await load(join(runDir, 'run.yaml'));
for (const key of ['machine.os', 'machine.model', 'gridcraft.version', 'gridcraft.installer_sha256', 'reference.excel']) {
  const v = key.split('.').reduce((o, k) => o?.[k], run);
  if (v === undefined || v === null || v === '') warnings.push(`run.yaml: ${key} is empty`);
}

const evidenceFiles = new Set(await readdir(join(runDir, 'evidence')).catch(() => []));
const summary = [];
const judged = [];

for (const suite of Object.keys(SUITES)) {
  const file = join(runDir, `${suite}.yaml`);
  if (!(await exists(file))) continue;
  const doc = await load(file);
  const counts = Object.fromEntries(VERDICTS.map((v) => [v, 0]));
  let open = 0;
  let dataLoss = 0;
  let wrong = 0;
  for (const rec of doc.cases ?? []) {
    const where = `${suite}.yaml ${rec.id}`;
    const c = catalogue.get(rec.id);
    if (!c) {
      errors.push(`${where}: not in suites (renamed or removed?)`);
      continue;
    }
    const stages = suite === 'pf' ? ['M'] : c.stages;
    const given = stages.map((s) => rec[s]).filter((v) => v !== null && v !== undefined && v !== '');
    for (const s of stages) {
      const v = rec[s];
      if (v === null || v === undefined || v === '') continue;
      if (!VERDICTS.includes(v)) errors.push(`${where}: ${s} is "${v}", not one of ${VERDICTS.join(' ')}`);
      else counts[v]++;
    }
    for (const key of Object.keys(rec)) {
      if (/^[A-Z]$/.test(key) && !stages.includes(key)) errors.push(`${where}: stage ${key} doesn't belong to this case (${stages.join(' ')})`);
    }
    if (suite === 'pf') {
      const values = [...(rec.gridcraft ?? []), ...(rec.excel ?? []), ...(rec.other ?? [])];
      if (values.some((x) => typeof x !== 'number')) errors.push(`${where}: measurements must be numbers`);
      if ((rec.gridcraft ?? []).length && rec.gridcraft.length !== 5) warnings.push(`${where}: ${rec.gridcraft.length} GridCraft measurements, not 5`);
      if ((rec.other ?? []).length && !rec.other_app) errors.push(`${where}: "other" has values but other_app is empty`);
      if (!values.length && !given.length) open++;
    } else if (given.length < stages.length) {
      open++;
    }

    const flags = rec.flags ?? [];
    for (const f of flags) if (!FLAGS.includes(f)) errors.push(`${where}: unknown flag "${f}"`);
    if (flags.includes('dataLoss')) dataLoss++;
    if (flags.includes('wrongValue')) wrong++;
    const bad = given.some((v) => ['minor', 'major', 'fail', 'blocked', 'skip'].includes(v));
    if ((bad || flags.length) && !String(rec.note ?? '').trim()) errors.push(`${where}: a difference, skip or flag needs a note`);
    const shown = given.some((v) => ['minor', 'major', 'fail'].includes(v));
    const privateCase = /-PRIV-/.test(rec.id);
    if (shown && !privateCase && !(rec.evidence ?? []).length) warnings.push(`${where}: no screenshot for a difference`);
    if (privateCase && (rec.evidence ?? []).length) errors.push(`${where}: private samples keep no evidence in the repository`);
    for (const e of rec.evidence ?? []) if (!evidenceFiles.has(basename(String(e)))) errors.push(`${where}: evidence/${e} not found`);
    if (rec.issue && !/^https:\/\/github\.com\/storytold\/gridcraft\/(issues|pull|security\/advisories)\/[\w-]+$/.test(rec.issue) && rec.issue !== 'private report') warnings.push(`${where}: issue should be a link to a GridCraft issue, or "private report"`);
    if ((flags.includes('dataLoss') || flags.includes('wrongValue')) && c.stages.includes('V') && !rec.diff) warnings.push(`${where}: a data loss or wrong value needs the diff-xlsx output in "diff"`);
    if (rec.diff && !evidenceFiles.has(basename(String(rec.diff)))) errors.push(`${where}: evidence/${rec.diff} not found`);
    if (given.length && !rec.checked) warnings.push(`${where}: "checked" (the date it was judged) is empty`);
    if (rec.recheck) {
      for (const [s, v] of Object.entries(rec.recheck)) {
        if (s === 'date') continue;
        if (!VERDICTS.includes(v)) errors.push(`${where}: recheck ${s} is "${v}"`);
        else if (rec[s] && rec[s] !== v && !String(rec.note).includes('复判')) errors.push(`${where}: recheck of ${s} differs (${rec[s]} → ${v}); say which one stands in the note ("复判：…")`);
      }
    }
    if (given.length) judged.push({ suite, id: rec.id, recheck: Boolean(rec.recheck) });
    if (rec.id === 'ST-LOG-01') for (const e of rec.events ?? []) if (!e.time || !e.action) errors.push(`${where}: each event needs a time and an action`);
  }
  summary.push({ suite, cases: (doc.cases ?? []).length, open, dataLoss, wrongValue: wrong, ...counts });
}

console.log(`\n${basename(runDir)}`);
console.table(summary);

const rechecked = judged.filter((j) => j.recheck);
if (rechecked.length) console.log(`Judged again: ${rechecked.length} cases`);

if (samplePct) {
  const rand = rng([...basename(runDir)].reduce((h, ch) => Math.imul(h ^ ch.codePointAt(0), 16777619) >>> 0, 2166136261));
  const pool = judged.filter((j) => !j.recheck);
  const n = Math.ceil((pool.length * samplePct) / 100);
  const pick = [];
  while (pick.length < n && pool.length) pick.push(pool.splice(Math.floor(rand() * pool.length), 1)[0].id);
  console.log(`\nJudge these again, at least a day later, without looking at the first result (${n} of ${judged.length}):\n${pick.sort().join('\n')}`);
}

for (const e of errors) console.log(`error    ${e}`);
for (const w of warnings) console.log(`warning  ${w}`);
console.log(`\n${errors.length} errors, ${warnings.length} warnings`);
process.exit(errors.length || (strict && warnings.length) ? 1 : 0);
