#!/usr/bin/env node
// Compares two XLSX files: the workbook before (usually the original or Excel's copy) and after
// (usually what GridCraft saved). It lists what changed; a person decides whether that is a problem.
//
//   node tools/diff-xlsx.mjs corpus/generated/XL-BASIC.xlsx runs/…/outputs/XL-BASIC.gridcraft.xlsx
//   node tools/diff-xlsx.mjs corpus/made/FX-CORE.excel.xlsx runs/…/outputs/FX-CORE.gridcraft.xlsx --values
//   node tools/diff-xlsx.mjs a.xlsx b.xlsx --json > evidence/XL-RT-01-V-diff.json
//
// --values   compare only cached values (for calculation checks; formulas and formats are ignored)
// --tolerance <n>   relative tolerance for numbers (default 1e-9; absolute below 1)
// --max <n>  examples listed per category (default 20)
// --json     machine-readable output
//
// Exit code: 0 when nothing differs, 1 when something does, 2 on a usage or read error.

import { readFile } from 'node:fs/promises';
import { isCompoundFile, readXlsx } from './lib/xlsx.mjs';

const args = process.argv.slice(2);
const flag = (name) => args.includes(`--${name}`);
const opt = (name, fallback) => {
  const i = args.indexOf(`--${name}`);
  return i >= 0 ? args[i + 1] : fallback;
};
const files = args.filter((a, i) => !a.startsWith('--') && !['--tolerance', '--max'].includes(args[i - 1]));
if (files.length !== 2) {
  console.error('usage: diff-xlsx.mjs <before.xlsx> <after.xlsx> [--values] [--tolerance 1e-9] [--max 20] [--json]');
  process.exit(2);
}
const tolerance = Number(opt('tolerance', '1e-9'));
const max = Number(opt('max', '20'));
const valuesOnly = flag('values');

async function load(path) {
  const buf = await readFile(path);
  if (isCompoundFile(buf)) return { error: 'an OLE compound file: encrypted with a password, or an .xls' };
  try {
    return readXlsx(buf);
  } catch (err) {
    return { error: err.message };
  }
}
const [a, b] = await Promise.all(files.map(load));
for (const [i, wb] of [a, b].entries()) {
  if (wb.error) {
    console.error(`${files[i]}: ${wb.error}`);
    process.exit(2);
  }
}

const report = {};
const add = (category, line) => {
  report[category] ??= { count: 0, examples: [] };
  report[category].count++;
  if (report[category].examples.length < max) report[category].examples.push(line);
};

// Formulas are compared without the prefixes Excel adds for newer functions, and without spaces.
const normFormula = (f) => f?.replace(/_xlfn\.(_xlws\.)?|_xlpm\.|\s+/g, '').replace(/^=/, '').toUpperCase();
const sameNumber = (x, y) => {
  const scale = Math.max(1, Math.abs(x), Math.abs(y));
  return Math.abs(x - y) <= tolerance * scale;
};
const show = (cell) => (cell === undefined ? '(empty)' : cell.v === undefined ? '(no value)' : JSON.stringify(cell.v));
function sameValue(x, y) {
  const vx = x?.v;
  const vy = y?.v;
  if (vx === undefined || vx === '') return vy === undefined || vy === '';
  if (typeof vx === 'number' && typeof vy === 'number') return sameNumber(vx, vy);
  if (x.t === 'e' || y?.t === 'e') return x.t === y?.t && vx === vy;
  return vx === vy;
}

// Sheets, by name and order.
const namesA = a.sheets.map((s) => s.name);
const namesB = b.sheets.map((s) => s.name);
for (const n of namesA) if (!namesB.includes(n)) add('sheets', `missing after: "${n}"`);
for (const n of namesB) if (!namesA.includes(n)) add('sheets', `new after: "${n}"`);
if (namesA.filter((n) => namesB.includes(n)).join('\u0000') !== namesB.filter((n) => namesA.includes(n)).join('\u0000')) add('sheets', `order: ${namesA.join(', ')} → ${namesB.join(', ')}`);

for (const sa of a.sheets) {
  const sb = b.sheets.find((s) => s.name === sa.name);
  if (!sb) continue;
  const where = (r) => `${sa.name}!${r}`;
  const refs = new Set([...sa.cells.keys(), ...sb.cells.keys()]);
  for (const r of refs) {
    const ca = sa.cells.get(r);
    const cb = sb.cells.get(r);
    if (!sameValue(ca, cb)) add('values', `${where(r)}: ${show(ca)} → ${show(cb)}`);
    if (valuesOnly) continue;
    if (normFormula(ca?.f) !== normFormula(cb?.f)) add('formulas', `${where(r)}: ${ca?.f ?? '(none)'} → ${cb?.f ?? '(none)'}`);
    if ((ca?.fmt ?? 'General') !== (cb?.fmt ?? 'General') && (ca?.v !== undefined || cb?.v !== undefined)) add('numFmts', `${where(r)}: ${ca?.fmt ?? 'General'} → ${cb?.fmt ?? 'General'}`);
  }
  if (valuesOnly) continue;
  const ma = [...sa.merges].sort().join(' ');
  const mb = [...sb.merges].sort().join(' ');
  if (ma !== mb) add('merges', `${sa.name}: ${ma || '(none)'} → ${mb || '(none)'}`);
  for (const key of ['state', 'condFormats', 'validations', 'protected', 'frozen', 'rightToLeft']) {
    if (sa[key] !== sb[key]) add('sheetProps', `${sa.name} ${key}: ${sa[key]} → ${sb[key]}`);
  }
  const la = sa.hyperlinks.map((h) => `${h.ref}→${h.target}`).sort().join(' ');
  const lb = sb.hyperlinks.map((h) => `${h.ref}→${h.target}`).sort().join(' ');
  if (la !== lb) add('hyperlinks', `${sa.name}: ${la || '(none)'} → ${lb || '(none)'}`);
}

if (!valuesOnly) {
  const key = (d) => `${d.name}=${d.ref}`;
  const na = new Set(a.names.map(key));
  const nb = new Set(b.names.map(key));
  for (const n of na) if (!nb.has(n)) add('names', `missing after: ${n}`);
  for (const n of nb) if (!na.has(n)) add('names', `new after: ${n}`);
  if (a.workbookProtected !== b.workbookProtected) add('sheetProps', `workbook protection: ${a.workbookProtected} → ${b.workbookProtected}`);

  // Parts that carry content GridCraft might drop: charts, PivotTables, tables, comments, macros…
  const kinds = [
    ['charts', /^xl\/charts\/chart\d+\.xml$/],
    ['chartEx', /^xl\/charts\/chartEx\d+\.xml$/],
    ['drawings', /^xl\/drawings\/drawing\d+\.xml$/],
    ['pictures', /^xl\/media\//],
    ['tables', /^xl\/tables\//],
    ['pivotTables', /^xl\/pivotTables\//],
    ['pivotCaches', /^xl\/pivotCache\//],
    ['slicers', /^xl\/slicers?\//],
    ['slicerCaches', /^xl\/slicerCaches\//],
    ['timelines', /^xl\/timelines?\//],
    ['comments', /^xl\/comments\d+\.xml$/],
    ['threadedComments', /^xl\/threadedComments\//],
    ['vba', /^xl\/vbaProject\.bin$/],
    ['externalLinks', /^xl\/externalLinks\//],
    ['connections', /^xl\/connections\.xml$/],
    ['queryTables', /^xl\/queryTables\//],
    ['customXml', /^customXml\//],
    ['dataModel', /^xl\/model\//],
    ['controls', /^xl\/(ctrlProps|activeX)\//],
    ['embeddings', /^xl\/embeddings\//],
  ];
  for (const [kind, re] of kinds) {
    const ca = a.parts.filter((p) => re.test(p)).length;
    const cb = b.parts.filter((p) => re.test(p)).length;
    if (ca !== cb) add('parts', `${kind}: ${ca} → ${cb}`);
  }
  for (const p of ['creator', 'lastModifiedBy', 'title']) {
    if ((a.props[p] ?? '') !== (b.props[p] ?? '')) add('docProps', `${p}: "${a.props[p] ?? ''}" → "${b.props[p] ?? ''}"`);
  }
}

const categories = Object.keys(report);
if (flag('json')) {
  console.log(JSON.stringify({ before: files[0], after: files[1], valuesOnly, tolerance, differing: categories, report }, null, 1));
} else {
  console.log(`before: ${files[0]}\nafter:  ${files[1]}${valuesOnly ? '\n(values only)' : ''}\n`);
  if (!categories.length) console.log('No differences found in what this tool compares.');
  for (const c of categories) {
    console.log(`## ${c}: ${report[c].count}`);
    for (const line of report[c].examples) console.log(`  ${line}`);
    if (report[c].count > report[c].examples.length) console.log(`  … and ${report[c].count - report[c].examples.length} more`);
    console.log('');
  }
  console.log('Compared: sheets, cell values, formulas, number formats, merges, sheet settings, hyperlinks, names, parts (charts, PivotTables, tables, comments, VBA, links…) and document properties.');
  console.log('Not compared: fonts, fills, borders, column widths, chart and PivotTable details, print settings. Check those by eye.');
}
process.exit(categories.length ? 1 : 0);
