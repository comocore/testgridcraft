#!/usr/bin/env node
// Counts the cases per suite and priority, for the table in README.md §4.

import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { SUITES, loadCases } from './lib/cases.mjs';

const cases = await loadCases(join(fileURLToPath(new URL('..', import.meta.url)), 'suites'));
const rows = Object.entries(SUITES).map(([suite, { name }]) => {
  const list = cases.filter((c) => c.suite === suite);
  const n = (p) => list.filter((c) => c.priority === p).length;
  return { suite: suite.toUpperCase(), name, cases: list.length, P0: n('P0'), P1: n('P1'), P2: n('P2') };
});
console.table(rows);
const sum = (k) => rows.reduce((s, r) => s + r[k], 0);
console.log(`total ${sum('cases')}: P0 ${sum('P0')}, P1 ${sum('P1')}, P2 ${sum('P2')}`);
