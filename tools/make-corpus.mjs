#!/usr/bin/env node
// Generates the test workbooks and text files in corpus/generated/, byte for byte the same on every run.
//
//   node tools/make-corpus.mjs           # everything except the 1-million-row file
//   node tools/make-corpus.mjs --big     # also PF-ROWS-1M.xlsx (about 40 MB)
//
// Formulas are written without cached values and with fullCalcOnLoad, so Excel calculates them on open.
// The Excel copies that serve as the reference (corpus/made/*.excel.xlsx) are made by hand, see corpus/recipes.md.

import { createHash } from 'node:crypto';
import { mkdir, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { crc32, deflateRawSync } from 'node:zlib';
import { buildXlsx } from './lib/xlsx.mjs';
import { writeZip } from './lib/zip.mjs';

const root = fileURLToPath(new URL('..', import.meta.url));
const out = join(root, 'corpus', 'generated');
const big = process.argv.includes('--big');
await mkdir(out, { recursive: true });

const manifest = [];
async function emit(name, data, about) {
  const buf = Buffer.isBuffer(data) ? data : Buffer.from(data, 'utf8');
  await writeFile(join(out, name), buf);
  manifest.push({ name, bytes: buf.length, sha256: createHash('sha256').update(buf).digest('hex'), about });
}

// A deterministic pseudo-random sequence, so lists look varied but never change between runs.
function rng(seed) {
  let s = seed >>> 0;
  return () => ((s = (s * 1664525 + 1013904223) >>> 0) / 2 ** 32);
}
const f = (formula, s) => ({ f: formula, ...(s && { s }) });
const header = (...labels) => labels.map((v) => ({ v, s: 'header' }));

// ---- XL: a workbook with the everyday parts of a file ------------------------------------------

const regions = ['North', 'South', 'East', 'West', 'Central'];
const zhRegions = ['华北', '华南', '华东', '西南', '华中'];
await emit(
  'XL-BASIC.xlsx',
  buildXlsx({
    title: 'XL-BASIC',
    names: [
      { name: 'Rate', ref: "Data!$F$1" },
      { name: 'Amounts', ref: "Data!$D$4:$D$13" },
    ],
    sheets: [
      {
        name: 'Data',
        colWidths: [14, 12, 14, 14, 12, 10],
        freeze: 'A4',
        autoFilter: 'A3:E13',
        merges: ['A1:D1'],
        rows: [
          [{ v: 'XL-BASIC · Expected: A1:D1 merged, row 3 frozen with filter buttons, F1 = 0.06 named Rate', s: 'bold' }, null, null, null, 'Rate', { v: 0.06, s: 'pct' }],
          [],
          header('Region', 'Date', 'Units', 'Amount', 'With rate'),
          ...Array.from({ length: 10 }, (_, i) => [
            regions[i % 5],
            { v: 46023 + i * 17, s: 'date' },
            { v: 10 + i * 7, s: 'int' },
            { v: 1234.5 * (i + 1) - (i % 3) * 2000, s: 'red' },
            f(`D${i + 4}*(1+Rate)`, 'money'),
          ]),
          ['Total', null, f('SUBTOTAL(109,C4:C13)', 'int'), f('SUBTOTAL(109,D4:D13)', 'red'), f('SUM(E4:E13)', 'money')],
          [],
          ['Average of Amounts', null, null, f('AVERAGE(Amounts)', 'money')],
          ['Link', { v: 'example.com' }],
        ],
        links: [{ ref: 'B17', url: 'https://example.com/' }],
        condFormats: [
          '<conditionalFormatting sqref="D4:D13"><cfRule type="cellIs" dxfId="0" priority="1" operator="lessThan"><formula>0</formula></cfRule></conditionalFormatting>',
        ],
        validations: [
          '<dataValidation type="list" allowBlank="1" showErrorMessage="1" sqref="A4:A13"><formula1>"North,South,East,West,Central"</formula1></dataValidation>',
        ],
      },
      {
        name: '格式 Formats',
        colWidths: [30, 22],
        rows: [
          header('What', 'Cell'),
          ['General number', 1234.5678],
          ['#,##0', { v: 1234567.891, s: 'int' }],
          ['#,##0.00', { v: -1234.5, s: 'money' }],
          ['Negative red', { v: -42, s: 'red' }],
          ['0.00%', { v: 0.12345, s: 'pct' }],
          ['0.000%', { v: 0.000123, s: 'pct3' }],
          ['yyyy-mm-dd', { v: 46327, s: 'date' }],
          ['yyyy"年"m"月"d"日"', { v: 46327, s: 'zhDate' }],
          ['hh:mm:ss', { v: 0.75, s: 'time' }],
          ['¥ accounting', { v: 9876.5, s: 'yuan' }],
          ['Leading zeros 0000000000', { v: 42, s: 'zeros' }],
          ['Text format @ with a number', { v: '00123', s: 'text' }],
          ['Wrapped text', { v: 'A long line that should wrap inside its cell, keeping every word visible.', s: 'wrap' }],
          ['Yellow fill, thin border', { v: 'filled', s: 'yellow' }],
          ['Bold', { v: 'bold', s: 'bold' }],
          ['Italic blue', { v: 'italic', s: 'italicBlue' }],
          ['Rotated 45°', { v: 'rotated', s: 'rotated' }],
          ['TRUE', true],
          ['Error from formula', f('1/0')],
        ],
      },
      {
        name: 'Hidden',
        hidden: true,
        rows: [['This sheet is hidden. It should stay hidden after a save.']],
      },
    ],
  }),
  'Merges, frozen rows, AutoFilter, a defined name, a hyperlink, conditional formatting, a list validation, number formats, styles and a hidden sheet',
);

// ---- FX: formulas whose results Excel provides --------------------------------------------------

const fxCore = [
  ['Math', 'SUM(1,2,3)'],
  ['Math', 'ROUND(2.5,0)'],
  ['Math', 'ROUND(-2.5,0)'],
  ['Math', 'ROUND(1.005,2)'],
  ['Math', 'MROUND(7.5,5)'],
  ['Math', 'MOD(-7,3)'],
  ['Math', 'MOD(7,-3)'],
  ['Math', 'INT(-3.5)'],
  ['Math', 'TRUNC(-3.5)'],
  ['Math', 'POWER(2,0.5)'],
  ['Math', 'SUMPRODUCT({1,2,3},{4,5,6})'],
  ['Math', 'FACT(20)'],
  ['Math', 'COMBIN(49,6)'],
  ['Precision', '0.1+0.2=0.3'],
  ['Precision', '0.1+0.2-0.3'],
  ['Precision', '1-0.9'],
  ['Precision', '123456789012345678'],
  ['Precision', '1E+308*10'],
  ['Stats', 'AVERAGE(1,2,3,4,100)'],
  ['Stats', 'MEDIAN(1,2,3,4,100)'],
  ['Stats', 'STDEV.S(2,4,4,4,5,5,7,9)'],
  ['Stats', 'PERCENTILE.INC({1,2,3,4,5},0.3)'],
  ['Stats', 'RANK.EQ(3,{1,3,3,5})'],
  ['Stats', 'NORM.S.DIST(1.96,TRUE)'],
  ['Logic', 'IF(1>2,"yes","no")'],
  ['Logic', 'IFS(FALSE,1,TRUE,2)'],
  ['Logic', 'IFERROR(1/0,"div")'],
  ['Logic', 'SWITCH(3,1,"a",3,"c","other")'],
  ['Logic', 'AND(TRUE,1,"x"="x")'],
  ['Text', 'CONCAT("a","b",1)'],
  ['Text', 'TEXTJOIN("-",TRUE,"a","","b")'],
  ['Text', 'UPPER("straße")'],
  ['Text', 'PROPER("hello WORLD")'],
  ['Text', 'TEXT(1234.567,"#,##0.0")'],
  ['Text', 'TEXT(0.5,"0%")'],
  ['Text', 'TEXT(46327,"dddd, mmmm d, yyyy")'],
  ['Text', 'VALUE("1,234.5")'],
  ['Text', 'SUBSTITUTE("a-b-c","-","+",2)'],
  ['Text', 'TRIM("  a   b  ")'],
  ['Text', 'REPT("ab",3)'],
  ['Date', 'DATE(2026,2,29)'],
  ['Date', 'DATE(1900,2,29)'],
  ['Date', 'EOMONTH(DATE(2026,1,31),1)'],
  ['Date', 'EDATE(DATE(2026,1,31),1)'],
  ['Date', 'DATEDIF(DATE(2020,2,29),DATE(2026,2,28),"Y")'],
  ['Date', 'NETWORKDAYS(DATE(2026,11,1),DATE(2026,11,30))'],
  ['Date', 'WORKDAY(DATE(2026,12,24),3)'],
  ['Date', 'WEEKDAY(DATE(2026,11,2),2)'],
  ['Date', 'WEEKNUM(DATE(2026,1,1),21)'],
  ['Date', 'YEARFRAC(DATE(2026,1,31),DATE(2026,2,28),0)'],
  ['Date', 'TIME(25,61,61)'],
  ['Finance', 'PMT(0.05/12,360,-300000)'],
  ['Finance', 'FV(0.04/12,120,-500)'],
  ['Finance', 'NPV(0.08,-1000,300,400,500)'],
  ['Finance', 'IRR({-1000,300,400,500})'],
  ['Finance', 'RATE(60,-200,10000)'],
  ['Lookup', 'VLOOKUP("c",{"a",1;"b",2;"c",3},2,FALSE)'],
  ['Lookup', 'VLOOKUP(2.5,{1,"x";2,"y";3,"z"},2,TRUE)'],
  ['Lookup', 'HLOOKUP("b",{"a","b";1,2},2,FALSE)'],
  ['Lookup', 'INDEX({10,20,30},2)'],
  ['Lookup', 'MATCH(0.3,{0.1,0.2,0.3},0)'],
  ['Lookup', 'CHOOSE(2,"a","b","c")'],
  ['Errors', 'NA()'],
  ['Errors', 'SQRT(-1)'],
  ['Errors', '"a"+1'],
  ['Errors', 'ISERROR(1/0)'],
  ['Info', 'ISNUMBER("1")'],
  ['Info', 'N("text")'],
  ['Info', 'TYPE({1,2})'],
];
await emit(
  'FX-CORE.xlsx',
  buildXlsx({
    title: 'FX-CORE',
    sheets: [
      {
        name: 'Formulas',
        colWidths: [6, 12, 52, 24],
        freeze: 'A2',
        rows: [
          header('#', 'Area', 'Formula (as text)', 'Result'),
          ...fxCore.map(([area, formula], i) => [i + 1, area, `=${formula}`, f(formula)]),
        ],
      },
    ],
  }),
  `${fxCore.length} single-cell formulas: maths, precision, statistics, logic, text, dates, finance, lookups, errors. Excel's results come from the copy made in Excel (recipes.md)`,
);

// Formulas that refer to cells, including ranges, names and structured edits, for recalculation tests.
const r = rng(7);
const ledgerRows = Array.from({ length: 40 }, (_, i) => [
  { v: 46023 + Math.floor(r() * 300), s: 'date' },
  regions[Math.floor(r() * 5)],
  ['Rent', 'Food', 'Travel', 'Office', 'Other'][Math.floor(r() * 5)],
  { v: Math.round(r() * 200000) / 100, s: 'money' },
  f(`IF(D${i + 2}>1000,"large","small")`),
]);
await emit(
  'FX-LEDGER.xlsx',
  buildXlsx({
    title: 'FX-LEDGER',
    names: [{ name: 'Amounts', ref: 'Ledger!$D$2:$D$41' }],
    sheets: [
      {
        name: 'Ledger',
        colWidths: [12, 10, 10, 12, 10],
        freeze: 'A2',
        rows: [header('Date', 'Region', 'Category', 'Amount', 'Size'), ...ledgerRows],
      },
      {
        name: 'Summary',
        colWidths: [34, 16],
        rows: [
          header('Measure', 'Value'),
          ['SUM(Amounts)', f('SUM(Amounts)', 'money')],
          ['SUMIFS North + Travel', f('SUMIFS(Ledger!D:D,Ledger!B:B,"North",Ledger!C:C,"Travel")', 'money')],
          ['COUNTIF > 1000', f('COUNTIF(Ledger!D2:D41,">1000")')],
          ['AVERAGEIF Food', f('AVERAGEIF(Ledger!C2:C41,"Food",Ledger!D2:D41)', 'money')],
          ['MAXIFS South', f('MAXIFS(Ledger!D2:D41,Ledger!B2:B41,"South")', 'money')],
          ['Last date', f('MAX(Ledger!A2:A41)', 'date')],
          ['Rows with "large"', f('COUNTIF(Ledger!E2:E41,"large")')],
          ['3-D style: SUM of two cells', f("Ledger!D2+Ledger!D3", 'money')],
          ['INDEX/MATCH largest region', f('INDEX(Ledger!B2:B41,MATCH(MAX(Ledger!D2:D41),Ledger!D2:D41,0))')],
          ['SUMPRODUCT January', f('SUMPRODUCT((MONTH(Ledger!A2:A41)=1)*Ledger!D2:D41)', 'money')],
        ],
      },
    ],
  }),
  '40 ledger rows and a summary sheet of SUMIFS, COUNTIF, AVERAGEIF, MAXIFS, INDEX/MATCH and SUMPRODUCT over them',
);

// ---- ZH: Chinese text, formats and files --------------------------------------------------------

const zhNames = ['张伟', '王芳', '李娜', '刘洋', '陈静', '杨磊', '赵敏', '黄强', '周杰', '吴婷', '徐明', '孙丽'];
await emit(
  'ZH-TEXT.xlsx',
  buildXlsx({
    title: '中文测试',
    sheets: [
      {
        name: '中文 工作表',
        colWidths: [10, 10, 40, 18, 18],
        rows: [
          header('姓名', '地区', '备注', '金额', '日期'),
          ...zhNames.map((name, i) => [
            name,
            zhRegions[i % 5],
            i % 4 === 0 ? { v: '这是一段比较长的中文备注，用来检查自动换行时每一行的高度和断行位置是否正确。', s: 'wrap' } : `备注 ${i + 1}：Mixed 中英文 text 混排`,
            { v: 1000 + i * 321.5, s: 'yuan' },
            { v: 46023 + i * 31, s: 'zhDate' },
          ]),
          [],
          ['字体', '宋体', { v: '这一格指定了“宋体”字体（Windows 有，Mac 上要替换）', s: 'songti' }],
          ['符号', '①②③ ㎡ ℃ ¥ ￥ 〇', '全角标点：，。；：“”‘’（）【】《》'],
          ['繁体', '臺灣 香港', '繁體中文：開啟、儲存、試算表'],
          ['日文和韩文', '日本語 かな カナ', '한국어 텍스트'],
          ['生僻字', '𠀀 𪚥 𰻞', '扩展 B 区以后的字，系统字体里可能没有'],
        ],
      },
      {
        name: '函数',
        colWidths: [36, 20],
        rows: [
          header('公式', '结果'),
          ['=LEN("中文abc")', f('LEN("中文abc")')],
          ['=LENB("中文abc")', f('LENB("中文abc")')],
          ['=LEFT("北京市朝阳区",3)', f('LEFT("北京市朝阳区",3)')],
          ['=MID("北京市朝阳区",4,3)', f('MID("北京市朝阳区",4,3)')],
          ['=FIND("市","北京市朝阳区")', f('FIND("市","北京市朝阳区")')],
          ['=SUBSTITUTE("一二一","一","三")', f('SUBSTITUTE("一二一","一","三")')],
          ['=TEXT(46327,"aaaa")', f('TEXT(46327,"aaaa")')],
          ['=TEXT(46327,"yyyy年m月d日")', f('TEXT(46327,"yyyy年m月d日")')],
          ['=TEXT(1234,"[DBNum1]")', f('TEXT(1234,"[DBNum1]")')],
          ['=TEXT(1234,"[DBNum2]")', f('TEXT(1234,"[DBNum2]")')],
          ['=UNICODE("中")', f('UNICODE("中")')],
          ['=UNICHAR(20013)', f('UNICHAR(20013)')],
          ['=UPPER("ａｂｃ")', f('UPPER("ａｂｃ")')],
          ['=EXACT("中","中")', f('EXACT("中","中")')],
          ['="中"<"国"', f('"中"<"国"')],
        ],
      },
    ],
  }),
  'Chinese names, long Chinese text that wraps, mixed Chinese and Latin text, a cell set in 宋体, full-width punctuation, Traditional Chinese, Japanese, Korean, rare characters, a Chinese date format, ¥ accounting and Chinese text functions',
);

// The same table as CSV in two encodings, and ID and phone numbers that must not lose digits.
const zhCsvRows = [['姓名', '地区', '金额', '日期'], ...zhNames.slice(0, 6).map((n, i) => [n, zhRegions[i % 5], (1000 + i * 321.5).toFixed(2), `2026/${i + 1}/15`])];
const csv = (rows) => rows.map((row) => row.join(',')).join('\r\n') + '\r\n';
await emit('ZH-CSV-UTF8-BOM.csv', Buffer.concat([Buffer.from([0xef, 0xbb, 0xbf]), Buffer.from(csv(zhCsvRows), 'utf8')]), 'The Chinese table as UTF-8 CSV with a byte order mark, as Excel saves "CSV UTF-8"');
await emit('ZH-CSV-UTF8.csv', csv(zhCsvRows), 'The Chinese table as UTF-8 CSV without a byte order mark');
await emit('ZH-CSV-GBK.csv', encodeGbk(csv(zhCsvRows)), 'The Chinese table encoded in GBK, as Chinese Excel and WPS save plain "CSV"');
await emit(
  'ZH-IDS.csv',
  csv([
    ['姓名', '身份证号', '手机', '工号', '银行卡号'],
    ['张伟', '110105199003071234', '13800138000', '000123', '6222021234567890123'],
    ['王芳', '31010419851230002X', '015912345678', '000456', '6228480402564890018'],
  ]),
  'Fake 18-digit ID numbers (one ending in X), phone numbers, IDs with leading zeros and 19-digit card numbers: digits past the 15th must not turn into zeros',
);

// GBK for the few characters above, from a small table: Node has no GBK encoder.
function encodeGbk(text) {
  const table = {
    姓: 'd0d5', 名: 'c3fb', 地: 'b5d8', 区: 'c7f8', 金: 'bdf0', 额: 'b6ee', 日: 'c8d5', 期: 'c6da', 张: 'd5c5', 伟: 'ceb0',
    王: 'cdf5', 芳: 'b7bc', 李: 'c0ee', 娜: 'c4c8', 刘: 'c1f5', 洋: 'd1f3', 陈: 'b3c2', 静: 'beb2', 杨: 'd1ee', 磊: 'c0da',
    华: 'bbaa', 北: 'b1b1', 南: 'c4cf', 东: 'b6ab', 西: 'cef7', 中: 'd6d0',
  };
  const bytes = [];
  for (const ch of text) {
    if (ch.charCodeAt(0) < 0x80) bytes.push(ch.charCodeAt(0));
    else if (table[ch]) bytes.push(parseInt(table[ch].slice(0, 2), 16), parseInt(table[ch].slice(2), 16));
    else throw new Error(`no GBK code for ${ch}`);
  }
  return Buffer.from(bytes);
}

// ---- CSV dates ---------------------------------------------------------------------------------

await emit('CSV-DATES-DMY.csv', csv([['Date', 'Amount'], ['03/04/2026', '10'], ['31/05/2026', '20'], ['12/11/2026', '30']]), 'Day-first dates; 31/05 can only be day first, so 03/04/2026 is 3 April');
await emit('CSV-DATES-MDY.csv', csv([['Date', 'Amount'], ['03/04/2026', '10'], ['12/31/2026', '20'], ['11/12/2026', '30']]), 'Month-first dates; 12/31 can only be month first, so 03/04/2026 is 4 March');
await emit('CSV-DATES-AMBIG.csv', csv([['Date', 'Amount'], ['03/04/2026', '10'], ['05/06/2026', '20'], ['11/12/2026', '30']]), 'Dates that read either way; GridCraft 0.15.0 says it guessed month first');
await emit('CSV-DATES-MIXED.csv', csv([['Date', 'Amount'], ['31/05/2026', '10'], ['12/31/2026', '20']]), 'One day-first and one month-first date in a column: should stay text, with a warning');
await emit('CSV-DATES-ISO.csv', csv([['Date', 'Amount'], ['2026-04-03', '10'], ['2026/05/31', '20'], ['2026年6月1日', '30']]), 'Year-first and Chinese dates, which have no ambiguity');

// ---- ED: a list to sort, filter and edit --------------------------------------------------------

const r2 = rng(42);
await emit(
  'ED-LIST.xlsx',
  buildXlsx({
    title: 'ED-LIST',
    sheets: [
      {
        name: 'List',
        colWidths: [8, 10, 10, 12, 12, 10],
        freeze: 'A2',
        rows: [
          header('ID', 'Name', 'Region', 'Date', 'Amount', 'Units'),
          ...Array.from({ length: 200 }, (_, i) => [
            i + 1,
            zhNames[Math.floor(r2() * zhNames.length)],
            regions[Math.floor(r2() * 5)],
            { v: 46023 + Math.floor(r2() * 365), s: 'date' },
            { v: Math.round(r2() * 1000000) / 100, s: 'money' },
            Math.floor(r2() * 50),
          ]),
          ['Total', null, null, null, f('SUBTOTAL(109,E2:E201)', 'money'), f('SUBTOTAL(109,F2:F201)')],
        ],
      },
    ],
  }),
  '200 rows (IDs, Chinese names, regions, dates, amounts, units) with a SUBTOTAL row, for sorting, filtering, filling, PivotTables and charts',
);

// ---- PR: pages to print -------------------------------------------------------------------------

await emit(
  'PR-PAGES.xlsx',
  buildXlsx({
    title: 'PR-PAGES',
    sheets: [
      {
        name: 'Report',
        colWidths: [8, 14, 30, 14, 14, 14, 14],
        rows: [
          header('Row', 'Region', 'Description 说明', 'Q1', 'Q2', 'Q3', 'Q4'),
          ...Array.from({ length: 150 }, (_, i) => [
            i + 1,
            `${regions[i % 5]} ${zhRegions[i % 5]}`,
            i % 10 === 0 ? '中文说明：这一行用来检查 PDF 里的中文' : `Line ${i + 1}`,
            { v: 1000 + i * 3.5, s: 'money' },
            { v: 1100 + i * 2.25, s: 'money' },
            { v: 900 + i * 4.75, s: 'money' },
            { v: 1200 + i * 1.5, s: 'money' },
          ]),
        ],
      },
    ],
  }),
  '150 rows across several printed pages, with a header row to repeat and Chinese text every tenth row',
);

// ---- SC: links and functions that could reach outside -------------------------------------------

await emit(
  'SC-LINKS.xlsx',
  buildXlsx({
    title: 'SC-LINKS',
    sheets: [
      {
        name: 'Links',
        colWidths: [24, 46, 50],
        rows: [
          header('Kind', 'Cell', 'Expected'),
          ['https', 'https://example.com/', 'Asks first, then opens the browser'],
          ['http', 'http://example.com/', 'Asks first, then opens the browser'],
          ['mailto', 'mailto:test@example.com', 'Asks first, then opens the mail app'],
          ['file', 'file:///etc/hosts', 'Refused'],
          ['UNC path', '\\\\server\\share\\file.xlsx', 'Refused'],
          ['custom scheme', 'ms-settings:privacy', 'Refused'],
          ['javascript', 'javascript:alert(1)', 'Refused'],
          ['HYPERLINK()', f('HYPERLINK("https://example.com/hyperlink-function","HYPERLINK function")'), 'Asks first'],
          ['WEBSERVICE()', f('WEBSERVICE("https://example.com/")'), '#VALUE!, no network request'],
          ['FILTERXML()', f('FILTERXML("<a><b>1</b></a>","//b")'), '#VALUE!'],
        ],
        links: [
          { ref: 'B2', url: 'https://example.com/' },
          { ref: 'B3', url: 'http://example.com/' },
          { ref: 'B4', url: 'mailto:test@example.com' },
          { ref: 'B5', url: 'file:///etc/hosts' },
          { ref: 'B6', url: '\\\\server\\share\\file.xlsx' },
          { ref: 'B7', url: 'ms-settings:privacy' },
          { ref: 'B8', url: 'javascript:alert(1)' },
        ],
      },
    ],
  }),
  'Cell hyperlinks of every kind (https, http, mailto, file, UNC, a custom scheme, javascript), HYPERLINK, WEBSERVICE and FILTERXML',
);

// ---- PF: large files ---------------------------------------------------------------------------

function bigList(rows, cols, seed) {
  const g = rng(seed);
  return [
    header(...Array.from({ length: cols }, (_, c) => `Col ${c + 1}`)),
    ...Array.from({ length: rows }, (_, i) =>
      Array.from({ length: cols }, (_, c) => (c === 0 ? i + 1 : c === 1 ? regions[i % 5] : Math.round(g() * 100000) / 100)),
    ),
  ];
}
await emit('PF-ROWS-100K.xlsx', buildXlsx({ title: 'PF-ROWS-100K', sheets: [{ name: 'Data', rows: bigList(100000, 8, 1) }] }), '100,000 rows × 8 columns of values');
await emit(
  'PF-CALC-20K.xlsx',
  buildXlsx({
    title: 'PF-CALC-20K',
    sheets: [
      {
        name: 'Calc',
        rows: [
          header('Input', 'Running total', 'Share', 'Region', 'Region total'),
          ...Array.from({ length: 20000 }, (_, i) => [
            (i % 97) + 1,
            f(i === 0 ? 'A2' : `B${i + 1}+A${i + 2}`),
            f(`A${i + 2}/SUM($A$2:$A$20001)`),
            regions[i % 5],
            f(`SUMIFS($A$2:$A$20001,$D$2:$D$20001,D${i + 2})`),
          ]),
        ],
      },
    ],
  }),
  '20,000 rows of chained running totals, shares of a total and SUMIFS: a recalculation load. Change A2 and time the update',
);
await emit('PF-WIDE.xlsx', buildXlsx({ title: 'PF-WIDE', sheets: [{ name: 'Wide', rows: bigList(200, 1000, 3) }] }), '200 rows × 1,000 columns of values');
if (big) await emit('PF-ROWS-1M.xlsx', buildXlsx({ title: 'PF-ROWS-1M', sheets: [{ name: 'Data', rows: bigList(1048575, 5, 9) }] }), '1,048,575 rows × 5 columns: a full sheet');

// ---- ST: damaged and hostile files --------------------------------------------------------------

const basic = buildXlsx({ sheets: [{ name: 'Sheet1', rows: [['ST probe'], [1], [2], [f('A2+A3')]] }] });
await emit('ST-BAD-TRUNC.xlsx', basic.subarray(0, Math.floor(basic.length * 0.6)), 'The first 60% of a valid workbook');
await emit('ST-BAD-EMPTY.xlsx', Buffer.alloc(0), 'An empty file named .xlsx');
await emit('ST-BAD-NOTZIP.xlsx', 'This is a text file with an .xlsx name.\n', 'A text file named .xlsx');
await emit(
  'ST-BAD-XML.xlsx',
  writeZip([
    { name: '[Content_Types].xml', data: '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>' },
    { name: '_rels/.rels', data: '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>' },
    { name: 'xl/workbook.xml', data: '<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>' },
    { name: 'xl/_rels/workbook.xml.rels', data: '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>' },
    { name: 'xl/worksheets/sheet1.xml', data: '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>unclosed' },
  ]),
  'A package whose sheet XML stops in the middle of a cell',
);
// About 400 MB of row XML that packs into well under a megabyte.
const bombRow = Buffer.from('<row r="1"><c r="A1"><v>1</v></c></row>'.padEnd(4096, ' '));
const bombRaw = Buffer.concat([Buffer.from('<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>'), ...Array(100000).fill(bombRow), Buffer.from('</sheetData></worksheet>')]);
await emit('ST-BAD-BOMB.xlsx', bombXlsx(deflateRawSync(bombRaw, { level: 9 }), crc32(bombRaw) >>> 0, bombRaw.length), 'A sheet part that unpacks to about 400 MB of whitespace');
await emit(
  'ST-BAD-DEEP.xlsx',
  buildXlsx({ sheets: [{ name: 'Deep', rows: [[f(`${'('.repeat(3000)}1${')'.repeat(3000)}`)], [f(`${'SUM('.repeat(1000)}1${')'.repeat(1000)}`)], [f('SUM(A:XFD)')], [f('ROWS(SEQUENCE(1048576,16384))')], [f('LEN(REPT("x",32767))*LEN(REPT("y",32767))')]] }] }),
  'Formulas nested 3,000 and 1,000 levels deep, a sum over the whole sheet and a 17-billion-cell SEQUENCE: should give errors or results, not a crash or a hang',
);

function bombXlsx(packedSheet, crc, rawSize) {
  // The writer deflates entries itself, so build this package by hand around the pre-packed sheet.
  const parts = [
    ['[Content_Types].xml', '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'],
    ['_rels/.rels', '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'],
    ['xl/workbook.xml', '<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Bomb" sheetId="1" r:id="rId1"/></sheets></workbook>'],
    ['xl/_rels/workbook.xml.rels', '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>'],
  ];
  const zip = writeZip(parts.map(([name, data]) => ({ name, data })));
  // Append the sheet as a fifth entry: rebuild the central directory with it.
  return appendPacked(zip, 'xl/worksheets/sheet1.xml', packedSheet, crc, rawSize);
}

function appendPacked(zip, name, packed, crc, rawSize) {
  const eocd = zip.length - 22;
  const count = zip.readUInt16LE(eocd + 10);
  const dirSize = zip.readUInt32LE(eocd + 12);
  const dirOffset = zip.readUInt32LE(eocd + 16);
  const nameBuf = Buffer.from(name);
  const head = Buffer.alloc(30);
  head.writeUInt32LE(0x04034b50, 0);
  head.writeUInt16LE(20, 4);
  head.writeUInt16LE(8, 8);
  head.writeUInt32LE(crc, 14);
  head.writeUInt32LE(packed.length, 18);
  head.writeUInt32LE(rawSize, 22);
  head.writeUInt16LE(nameBuf.length, 26);
  const dir = Buffer.alloc(46);
  dir.writeUInt32LE(0x02014b50, 0);
  dir.writeUInt16LE(20, 4);
  dir.writeUInt16LE(20, 6);
  dir.writeUInt16LE(8, 10);
  dir.writeUInt32LE(crc, 16);
  dir.writeUInt32LE(packed.length, 20);
  dir.writeUInt32LE(rawSize, 24);
  dir.writeUInt16LE(nameBuf.length, 28);
  dir.writeUInt32LE(dirOffset, 42);
  const local = zip.subarray(0, dirOffset);
  const oldDir = zip.subarray(dirOffset, dirOffset + dirSize);
  const newDirOffset = dirOffset + head.length + nameBuf.length + packed.length;
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0);
  end.writeUInt16LE(count + 1, 8);
  end.writeUInt16LE(count + 1, 10);
  end.writeUInt32LE(dirSize + dir.length + nameBuf.length, 12);
  end.writeUInt32LE(newDirOffset, 16);
  return Buffer.concat([local, head, nameBuf, packed, oldDir, dir, nameBuf, end]);
}

await writeFile(join(out, 'manifest.json'), JSON.stringify({ generator: 'tools/make-corpus.mjs', files: manifest }, null, 1) + '\n');
console.log(`${manifest.length} files in corpus/generated/`);
for (const m of manifest) console.log(`  ${m.name.padEnd(22)} ${String(m.bytes).padStart(10)} bytes`);
