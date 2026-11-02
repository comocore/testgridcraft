// Writes small XLSX workbooks from a plain description, and reads back what a comparison needs:
// sheets, cell values, formulas, number formats, merges, names, parts and document properties.
// Only the parts of SpreadsheetML these tests use; no dependencies.

import { readZip, writeZip } from './zip.mjs';

const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const unesc = (s) =>
  s
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#x([0-9a-f]+);/gi, (_, h) => String.fromCodePoint(parseInt(h, 16)))
    .replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(Number(d)))
    .replace(/&amp;/g, '&');

// ---- addresses -------------------------------------------------------------------------------

export function colName(n) {
  let s = '';
  for (n += 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
}
export const ref = (row, col) => `${colName(col)}${row + 1}`;
export function parseRef(r) {
  const m = /^([A-Z]+)(\d+)$/.exec(r);
  if (!m) throw new Error(`bad cell reference ${r}`);
  let col = 0;
  for (const ch of m[1]) col = col * 26 + ch.charCodeAt(0) - 64;
  return { row: Number(m[2]) - 1, col: col - 1 };
}

// ---- styles ----------------------------------------------------------------------------------

// Named styles the generator can use. Each is [numFmtId, fontId, fillId, borderId, alignment XML].
const CUSTOM_FORMATS = {
  164: 'yyyy-mm-dd',
  165: '#,##0.00;[Red]-#,##0.00',
  166: 'yyyy"年"m"月"d"日"',
  167: '_-¥* #,##0.00_-;-¥* #,##0.00_-;_-¥* "-"??_-;_-@_-',
  168: '0.000%',
  169: '[$-x-sysdate]dddd, mmmm dd, yyyy',
  170: 'hh:mm:ss',
  171: '0000000000',
};
export const STYLES = {
  default: [0, 0, 0, 0, ''],
  bold: [0, 1, 0, 0, ''],
  header: [0, 1, 2, 1, '<alignment horizontal="center"/>'],
  int: [3, 0, 0, 0, ''], // #,##0
  money: [4, 0, 0, 0, ''], // #,##0.00
  pct: [10, 0, 0, 0, ''], // 0.00%
  date: [164, 0, 0, 0, ''],
  red: [165, 0, 0, 0, ''],
  zhDate: [166, 0, 0, 0, ''],
  yuan: [167, 0, 0, 0, ''],
  pct3: [168, 0, 0, 0, ''],
  time: [170, 0, 0, 0, ''],
  zeros: [171, 0, 0, 0, ''],
  text: [49, 0, 0, 0, ''], // @
  wrap: [0, 0, 0, 1, '<alignment wrapText="1" vertical="top"/>'],
  yellow: [0, 0, 3, 1, ''],
  center: [0, 0, 0, 0, '<alignment horizontal="center" vertical="center"/>'],
  rotated: [0, 0, 0, 0, '<alignment textRotation="45"/>'],
  italicBlue: [0, 2, 0, 0, ''],
  songti: [0, 3, 0, 0, ''],
};
const STYLE_NAMES = Object.keys(STYLES);

function stylesXml() {
  const numFmts = Object.entries(CUSTOM_FORMATS)
    .map(([id, code]) => `<numFmt numFmtId="${id}" formatCode="${esc(code)}"/>`)
    .join('');
  const xfs = STYLE_NAMES.map((name) => {
    const [numFmtId, fontId, fillId, borderId, align] = STYLES[name];
    const apply = `${numFmtId ? ' applyNumberFormat="1"' : ''}${fontId ? ' applyFont="1"' : ''}${fillId ? ' applyFill="1"' : ''}${borderId ? ' applyBorder="1"' : ''}${align ? ' applyAlignment="1"' : ''}`;
    const open = `<xf numFmtId="${numFmtId}" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}" xfId="0"${apply}`;
    return align ? `${open}>${align}</xf>` : `${open}/>`;
  }).join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="${Object.keys(CUSTOM_FORMATS).length}">${numFmts}</numFmts>
<fonts count="4"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font><font><b/><sz val="11"/><name val="Calibri"/><family val="2"/></font><font><i/><sz val="11"/><color rgb="FF1F4E9E"/><name val="Calibri"/><family val="2"/></font><font><sz val="12"/><name val="宋体"/><charset val="134"/></font></fonts>
<fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFD9E1F2"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFF2A8"/></patternFill></fill></fills>
<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"/><right style="thin"/><top style="thin"/><bottom style="thin"/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="${STYLE_NAMES.length}">${xfs}</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
<dxfs count="1"><dxf><font><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf></dxfs>
</styleSheet>`;
}

// ---- writing ---------------------------------------------------------------------------------

// Functions added to Excel after 2007 are stored with an _xlfn. prefix; without it Excel reads #NAME?.
const NEWER_FUNCTIONS = [
  'CONCAT', 'TEXTJOIN', 'IFS', 'SWITCH', 'MAXIFS', 'MINIFS', 'IFNA', 'DAYS', 'ISOWEEKNUM', 'UNICODE', 'UNICHAR',
  'STDEV.S', 'STDEV.P', 'VAR.S', 'VAR.P', 'PERCENTILE.INC', 'PERCENTILE.EXC', 'QUARTILE.INC', 'RANK.EQ', 'RANK.AVG',
  'NORM.S.DIST', 'NORM.DIST', 'NORM.INV', 'NORM.S.INV', 'CEILING.MATH', 'FLOOR.MATH', 'AGGREGATE',
  'SEQUENCE', 'XLOOKUP', 'XMATCH', 'FILTER', 'SORT', 'SORTBY', 'UNIQUE', 'LET', 'LAMBDA', 'WEBSERVICE', 'FILTERXML',
];
const NEWER_RE = new RegExp(`(?<![A-Za-z0-9_.])(${NEWER_FUNCTIONS.map((n) => n.replace('.', '\\.')).join('|')})\\(`, 'g');
/** Adds _xlfn. to newer functions outside string literals. */
export function fileFormula(formula) {
  return formula
    .replace(/^=/, '')
    .split(/("(?:[^"]|"")*")/)
    .map((part, i) => (i % 2 ? part : part.replace(NEWER_RE, '_xlfn.$1(')))
    .join('');
}

/**
 * @typedef {{ v?: string | number | boolean, f?: string, s?: string, t?: 'str' | 'n' | 'b' | 'e' }} CellSpec
 * @typedef {{ name: string, rows: (CellSpec | string | number | boolean | null)[][], colWidths?: number[],
 *   merges?: string[], freeze?: string, links?: { ref: string, url: string }[], validations?: string[],
 *   condFormats?: string[], autoFilter?: string, rightToLeft?: boolean, hidden?: boolean }} SheetSpec
 * @typedef {{ sheets: SheetSpec[], names?: { name: string, ref: string }[], creator?: string, title?: string,
 *   fullCalcOnLoad?: boolean }} WorkbookSpec
 */

/** Builds an XLSX file. Plain values in `rows` become cells; `{ f }` becomes a formula with no cached value. */
export function buildXlsx(spec) {
  const strings = [];
  const stringIndex = new Map();
  const sst = (s) => {
    if (!stringIndex.has(s)) {
      stringIndex.set(s, strings.length);
      strings.push(s);
    }
    return stringIndex.get(s);
  };
  const files = [];
  const sheetEntries = [];
  spec.sheets.forEach((sheet, i) => {
    const n = i + 1;
    const rels = [];
    const rowsXml = sheet.rows
      .map((row, r) => {
        const cells = row
          .map((cell, c) => {
            if (cell === null || cell === undefined || cell === '') return '';
            const spec = typeof cell === 'object' ? cell : { v: cell };
            const style = spec.s ? ` s="${STYLE_NAMES.indexOf(spec.s)}"` : '';
            if (spec.s && STYLE_NAMES.indexOf(spec.s) < 0) throw new Error(`unknown style ${spec.s}`);
            const at = ref(r, c);
            if (spec.f !== undefined) {
              const t = spec.t ? ` t="${spec.t}"` : '';
              const v = spec.v !== undefined ? `<v>${esc(spec.v)}</v>` : '';
              return `<c r="${at}"${style}${t}><f>${esc(fileFormula(spec.f))}</f>${v}</c>`;
            }
            if (typeof spec.v === 'number') return `<c r="${at}"${style}><v>${spec.v}</v></c>`;
            if (typeof spec.v === 'boolean') return `<c r="${at}"${style} t="b"><v>${spec.v ? 1 : 0}</v></c>`;
            return `<c r="${at}"${style} t="s"><v>${sst(String(spec.v))}</v></c>`;
          })
          .join('');
        return cells ? `<row r="${r + 1}">${cells}</row>` : '';
      })
      .join('');
    const cols = sheet.colWidths?.length
      ? `<cols>${sheet.colWidths.map((w, c) => `<col min="${c + 1}" max="${c + 1}" width="${w}" customWidth="1"/>`).join('')}</cols>`
      : '';
    const pane = sheet.freeze
      ? (() => {
          const { row, col } = parseRef(sheet.freeze);
          const split = `${col ? ` xSplit="${col}"` : ''}${row ? ` ySplit="${row}"` : ''}`;
          const active = row && col ? 'bottomRight' : row ? 'bottomLeft' : 'topRight';
          return `<pane${split} topLeftCell="${sheet.freeze}" activePane="${active}" state="frozen"/>`;
        })()
      : '';
    const view = `<sheetViews><sheetView workbookViewId="0"${i === 0 ? ' tabSelected="1"' : ''}${sheet.rightToLeft ? ' rightToLeft="1"' : ''}>${pane}</sheetView></sheetViews>`;
    const autoFilter = sheet.autoFilter ? `<autoFilter ref="${sheet.autoFilter}"/>` : '';
    const merges = sheet.merges?.length ? `<mergeCells count="${sheet.merges.length}">${sheet.merges.map((m) => `<mergeCell ref="${m}"/>`).join('')}</mergeCells>` : '';
    const cf = (sheet.condFormats ?? []).join('');
    const dv = sheet.validations?.length ? `<dataValidations count="${sheet.validations.length}">${sheet.validations.join('')}</dataValidations>` : '';
    const links = sheet.links?.length
      ? `<hyperlinks>${sheet.links
          .map((l, k) => {
            rels.push(`<Relationship Id="rIdL${k + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${esc(l.url)}" TargetMode="External"/>`);
            return `<hyperlink ref="${l.ref}" r:id="rIdL${k + 1}"/>`;
          })
          .join('')}</hyperlinks>`
      : '';
    files.push({
      name: `xl/worksheets/sheet${n}.xml`,
      data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${view}<sheetFormatPr defaultRowHeight="15"/>${cols}<sheetData>${rowsXml}</sheetData>${autoFilter}${merges}${cf}${dv}${links}<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>`,
    });
    if (rels.length) {
      files.push({
        name: `xl/worksheets/_rels/sheet${n}.xml.rels`,
        data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join('')}</Relationships>`,
      });
    }
    sheetEntries.push(`<sheet name="${esc(sheet.name)}" sheetId="${n}"${sheet.hidden ? ' state="hidden"' : ''} r:id="rId${n}"/>`);
  });

  const n = spec.sheets.length;
  const names = spec.names?.length ? `<definedNames>${spec.names.map((d) => `<definedName name="${esc(d.name)}">${esc(d.ref)}</definedName>`).join('')}</definedNames>` : '';
  const calc = `<calcPr calcId="191029"${spec.fullCalcOnLoad === false ? '' : ' fullCalcOnLoad="1"'}/>`;
  files.unshift(
    {
      name: '[Content_Types].xml',
      data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>${spec.sheets
        .map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`)
        .join('')}<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>`,
    },
    {
      name: '_rels/.rels',
      data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>`,
    },
    {
      name: 'xl/workbook.xml',
      data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView activeTab="0"/></bookViews><sheets>${sheetEntries.join('')}</sheets>${names}${calc}</workbook>`,
    },
    {
      name: 'xl/_rels/workbook.xml.rels',
      data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${spec.sheets
        .map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`)
        .join('')}<Relationship Id="rId${n + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId${n + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>`,
    },
    { name: 'xl/styles.xml', data: stylesXml() },
  );
  files.push(
    {
      name: 'xl/sharedStrings.xml',
      data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${strings.length}" uniqueCount="${strings.length}">${strings
        .map((s) => `<si><t${/^\s|\s$|\n/.test(s) ? ' xml:space="preserve"' : ''}>${esc(s)}</t></si>`)
        .join('')}</sst>`,
    },
    {
      name: 'docProps/core.xml',
      data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${esc(spec.title ?? '')}</dc:title><dc:creator>${esc(spec.creator ?? 'GetGridCraft test corpus')}</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">2026-01-01T00:00:00Z</dcterms:created></cp:coreProperties>`,
    },
  );
  return writeZip(files);
}

// ---- reading ---------------------------------------------------------------------------------

const BUILTIN_FORMATS = {
  0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00', 9: '0%', 10: '0.00%', 11: '0.00E+00', 12: '# ?/?',
  13: '# ??/??', 14: 'mm-dd-yy', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy', 18: 'h:mm AM/PM', 19: 'h:mm:ss AM/PM',
  20: 'h:mm', 21: 'h:mm:ss', 22: 'm/d/yy h:mm', 37: '#,##0 ;(#,##0)', 38: '#,##0 ;[Red](#,##0)', 39: '#,##0.00;(#,##0.00)',
  40: '#,##0.00;[Red](#,##0.00)', 45: 'mm:ss', 46: '[h]:mm:ss', 47: 'mmss.0', 48: '##0.0E+0', 49: '@',
};

const attr = (tag, name) => {
  const m = new RegExp(`\\s${name}="([^"]*)"`).exec(tag);
  return m ? unesc(m[1]) : undefined;
};
const textOf = (xml) => [...xml.matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g)].map((m) => unesc(m[1])).join('');

/** Encrypted workbooks and .xls files are OLE compound files, not ZIPs. */
export const isCompoundFile = (buf) => buf.length >= 8 && buf.readUInt32BE(0) === 0xd0cf11e0 && buf.readUInt32BE(4) === 0xa1b11ae1;

/** Reads an XLSX into a plain structure. Throws on a file that isn't a readable XLSX. */
export function readXlsx(buf) {
  if (isCompoundFile(buf)) throw new Error('an OLE compound file: an encrypted workbook or an .xls, not a plain XLSX');
  const zip = readZip(buf);
  const part = (name) => {
    const get = zip.get(name);
    return get ? get().toString('utf8') : undefined;
  };
  const relsOf = (xml) =>
    new Map([...(xml ?? '').matchAll(/<Relationship\b[^>]*>/g)].map((m) => [attr(m[0], 'Id'), { target: attr(m[0], 'Target'), type: attr(m[0], 'Type'), mode: attr(m[0], 'TargetMode') }]));
  const resolve = (base, target) => {
    if (target.startsWith('/')) return target.slice(1);
    const parts = base.split('/').slice(0, -1);
    for (const seg of target.split('/')) {
      if (seg === '..') parts.pop();
      else if (seg !== '.') parts.push(seg);
    }
    return parts.join('/');
  };

  const rootRels = relsOf(part('_rels/.rels'));
  const wbPath = [...rootRels.values()].find((r) => r.type?.endsWith('/officeDocument'))?.target?.replace(/^\//, '') ?? 'xl/workbook.xml';
  const wb = part(wbPath);
  if (!wb) throw new Error('no workbook part');
  const wbRels = relsOf(part(wbPath.replace(/([^/]+)$/, '_rels/$1.rels')));

  const shared = [...(part('xl/sharedStrings.xml') ?? '').matchAll(/<si>([\s\S]*?)<\/si>/g)].map((m) => textOf(m[1].replace(/<rPh\b[\s\S]*?<\/rPh>/g, '')));

  const stylesXml = part('xl/styles.xml') ?? '';
  const customFormats = new Map([...stylesXml.matchAll(/<numFmt\b[^>]*>/g)].map((m) => [Number(attr(m[0], 'numFmtId')), attr(m[0], 'formatCode')]));
  const cellXfs = /<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(stylesXml)?.[1] ?? '';
  const xfFormats = [...cellXfs.matchAll(/<xf\b[^>]*>/g)].map((m) => {
    const id = Number(attr(m[0], 'numFmtId') ?? 0);
    return customFormats.get(id) ?? BUILTIN_FORMATS[id] ?? `numFmtId ${id}`;
  });

  const sheets = [...wb.matchAll(/<sheet\b[^>]*>/g)].map((m) => {
    const rid = attr(m[0], 'r:id');
    const path = resolve(wbPath, wbRels.get(rid)?.target ?? '');
    const xml = part(path) ?? '';
    const cells = new Map();
    for (const c of xml.matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
      const tag = `<c${c[1]}>`;
      const body = c[2] ?? '';
      const r = attr(tag, 'r');
      const t = attr(tag, 't') ?? 'n';
      const s = Number(attr(tag, 's') ?? 0);
      const fm = /<f\b([^>]*?)(?:\/>|>([\s\S]*?)<\/f>)/.exec(body);
      const vm = /<v>([\s\S]*?)<\/v>/.exec(body);
      let v = vm ? unesc(vm[1]) : undefined;
      if (t === 's' && v !== undefined) v = shared[Number(v)];
      else if (t === 'inlineStr') v = textOf(body);
      else if (t === 'b' && v !== undefined) v = v === '1';
      else if (t === 'n' && v !== undefined) v = Number(v);
      const formula = fm ? (fm[2] !== undefined ? unesc(fm[2]) : `(shared ${attr(`<f${fm[1]}>`, 'si') ?? '?'})`) : undefined;
      cells.set(r, { v, t, f: formula, fmt: xfFormats[s] ?? 'General' });
    }
    const merges = [...xml.matchAll(/<mergeCell\b[^>]*>/g)].map((mm) => attr(mm[0], 'ref'));
    const count = (re) => (xml.match(re) ?? []).length;
    const sheetRels = relsOf(part(path.replace(/([^/]+)$/, '_rels/$1.rels')));
    return {
      name: attr(m[0], 'name'),
      state: attr(m[0], 'state') ?? 'visible',
      path,
      cells,
      merges,
      condFormats: count(/<conditionalFormatting\b/g),
      validations: count(/<dataValidation\b/g),
      hyperlinks: [...xml.matchAll(/<hyperlink\b[^>]*>/g)].map((h) => {
        const id = attr(h[0], 'r:id');
        return { ref: attr(h[0], 'ref'), target: id ? sheetRels.get(id)?.target : `#${attr(h[0], 'location') ?? ''}` };
      }),
      protected: /<sheetProtection\b/.test(xml),
      frozen: /state="frozen"/.test(xml),
      rightToLeft: /rightToLeft="1"/.test(xml),
    };
  });

  const names = [...wb.matchAll(/<definedName\b([^>]*)>([\s\S]*?)<\/definedName>/g)].map((m) => ({
    name: attr(`<d${m[1]}>`, 'name'),
    ref: unesc(m[2]),
    hidden: attr(`<d${m[1]}>`, 'hidden') === '1',
  }));
  const core = part('docProps/core.xml') ?? '';
  const prop = (tag) => {
    const m = new RegExp(`<${tag}[^>]*>([\\s\\S]*?)</${tag}>`).exec(core);
    return m ? unesc(m[1]) : undefined;
  };
  return {
    parts: [...zip.keys()].sort(),
    sheets,
    names,
    workbookProtected: /<workbookProtection\b/.test(wb),
    fullCalcOnLoad: /fullCalcOnLoad="(1|true)"/.test(wb),
    props: { creator: prop('dc:creator'), lastModifiedBy: prop('cp:lastModifiedBy'), title: prop('dc:title') },
  };
}
