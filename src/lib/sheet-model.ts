/**
 * Spreadsheet document model + pure helpers shared by the Sheets screen,
 * the chart renderer and the file I/O layer.
 */
import { colIdx, colName, evalCell, serialToDate, type CellGetter, type EvalOptions } from './formulas';

export { colIdx, colName };

export const MAX_ROWS = 1000;
export const MAX_COLS = 100; // A..CV
export const DEFAULT_COL_W = 92;
export const DEFAULT_ROW_H = 34;

export type NumFmt = 'gen' | 'num' | 'cur' | 'pct' | 'date' | 'time' | 'sci' | 'text';
export type HAlign = 'left' | 'center' | 'right';
export type VAlign = 'top' | 'middle' | 'bottom';
export type BorderKind = 'all' | 'outline' | 'bottom' | 'top' | 'left' | 'right' | 'none';

export interface CellStyle {
  b?: boolean;
  i?: boolean;
  u?: boolean;
  s?: boolean;
  color?: string;
  fill?: string;
  align?: HAlign;
  valign?: VAlign;
  fmt?: NumFmt;
  dec?: number;
  wrap?: boolean;
  size?: number; // px font size
  bt?: boolean; // borders
  bb?: boolean;
  bl?: boolean;
  br?: boolean;
}

export type ChartType = 'column' | 'bar' | 'line' | 'pie' | 'area';

export interface Chart {
  id: string;
  type: ChartType;
  range: string;
  title: string;
  /** first column of the range holds category labels */
  labelsInFirstCol: boolean;
  /** first row holds series names */
  headerRow: boolean;
}

export interface Page {
  cells: Record<string, string>;
  styles: Record<string, CellStyle>;
  notes?: Record<string, string>;
  colW?: Record<number, number>;
  rowH?: Record<number, number>;
  merges?: string[];
  freeze?: { r: number; c: number };
  filter?: { range: string; criteria: Record<number, string[]> };
  charts?: Chart[];
}

export interface Book {
  order: string[];
  active: string;
  sheets: Record<string, Page>;
}

export const FMT_OPTIONS: { v: NumFmt; t: string }[] = [
  { v: 'gen', t: 'General' },
  { v: 'num', t: 'Number' },
  { v: 'cur', t: 'Currency' },
  { v: 'pct', t: 'Percent' },
  { v: 'date', t: 'Date' },
  { v: 'time', t: 'Time' },
  { v: 'sci', t: 'Scientific' },
  { v: 'text', t: 'Text' },
];

// ---------------------------------------------------------------------------
// references
// ---------------------------------------------------------------------------

const REF_RE = /^([A-Za-z]{1,3})([0-9]{1,5})$/;

export function parseRef(ref: string): [number, number] {
  const m = REF_RE.exec(ref);
  if (!m) return [0, 0];
  return [colIdx(m[1]), parseInt(m[2], 10) - 1];
}

export function refOf(c: number, r: number): string {
  return `${colName(c)}${r + 1}`;
}

export function isRef(s: string): boolean {
  return REF_RE.test(s.trim());
}

export interface Rect {
  c1: number;
  r1: number;
  c2: number;
  r2: number;
}

export function rectOf(a: string, b?: string): Rect {
  const [c1, r1] = parseRef(a);
  const [c2, r2] = parseRef(b ?? a);
  return { c1: Math.min(c1, c2), r1: Math.min(r1, r2), c2: Math.max(c1, c2), r2: Math.max(r1, r2) };
}

export function rectToRange(r: Rect): string {
  return r.c1 === r.c2 && r.r1 === r.r2 ? refOf(r.c1, r.r1) : `${refOf(r.c1, r.r1)}:${refOf(r.c2, r.r2)}`;
}

export function rangeToRect(range: string): Rect | null {
  const parts = range.split(':');
  if (parts.length === 1) return isRef(parts[0]) ? rectOf(parts[0]) : null;
  if (parts.length === 2 && isRef(parts[0]) && isRef(parts[1])) return rectOf(parts[0], parts[1]);
  return null;
}

export function refsInRect(r: Rect): string[] {
  const out: string[] = [];
  for (let row = r.r1; row <= r.r2; row++) for (let c = r.c1; c <= r.c2; c++) out.push(refOf(c, row));
  return out;
}

export function inRect(c: number, r: number, rect: Rect): boolean {
  return c >= rect.c1 && c <= rect.c2 && r >= rect.r1 && r <= rect.r2;
}

/** Used extent of a page (max col/row index with content or style). */
export function usedExtent(pg: Page): { cols: number; rows: number } {
  let cols = 0;
  let rows = 0;
  const scan = (k: string) => {
    const [c, r] = parseRef(k);
    if (c + 1 > cols) cols = c + 1;
    if (r + 1 > rows) rows = r + 1;
  };
  Object.keys(pg.cells).forEach(scan);
  Object.keys(pg.styles).forEach(scan);
  return { cols, rows };
}

/** Contiguous block of data around a cell (like Excel's Ctrl+A / current region). */
export function currentRegion(cells: Record<string, string>, c: number, r: number): Rect {
  const has = (cc: number, rr: number) => (cells[refOf(cc, rr)] ?? '') !== '';
  if (!has(c, r)) return { c1: c, r1: r, c2: c, r2: r };
  let rect: Rect = { c1: c, r1: r, c2: c, r2: r };
  let grown = true;
  const rowHas = (rr: number, a: number, b: number) => { for (let cc = a; cc <= b; cc++) if (has(cc, rr)) return true; return false; };
  const colHas = (cc: number, a: number, b: number) => { for (let rr = a; rr <= b; rr++) if (has(cc, rr)) return true; return false; };
  while (grown) {
    grown = false;
    if (rect.r1 > 0 && rowHas(rect.r1 - 1, Math.max(0, rect.c1 - 1), rect.c2 + 1)) { rect = { ...rect, r1: rect.r1 - 1 }; grown = true; }
    if (rect.r2 < MAX_ROWS - 1 && rowHas(rect.r2 + 1, Math.max(0, rect.c1 - 1), rect.c2 + 1)) { rect = { ...rect, r2: rect.r2 + 1 }; grown = true; }
    if (rect.c1 > 0 && colHas(rect.c1 - 1, Math.max(0, rect.r1 - 1), rect.r2 + 1)) { rect = { ...rect, c1: rect.c1 - 1 }; grown = true; }
    if (rect.c2 < MAX_COLS - 1 && colHas(rect.c2 + 1, Math.max(0, rect.r1 - 1), rect.r2 + 1)) { rect = { ...rect, c2: rect.c2 + 1 }; grown = true; }
  }
  return rect;
}

/** Heuristic: does the first row of the rect look like a header? */
export function hasHeaderRow(cells: Record<string, string>, rect: Rect): boolean {
  if (rect.r2 === rect.r1) return false;
  let textHead = 0;
  let numBody = 0;
  for (let c = rect.c1; c <= rect.c2; c++) {
    const h = cells[refOf(c, rect.r1)] ?? '';
    const b = cells[refOf(c, rect.r1 + 1)] ?? '';
    if (h && Number.isNaN(Number(h)) && !h.startsWith('=')) textHead++;
    if (b && (!Number.isNaN(Number(b)) || b.startsWith('='))) numBody++;
  }
  return textHead > 0 && (numBody > 0 || textHead === rect.c2 - rect.c1 + 1);
}

// ---------------------------------------------------------------------------
// formulas: remapping when rows/cols move
// ---------------------------------------------------------------------------

const TOKEN_RE = /(?<![A-Za-z0-9_.])(\$?)([A-Za-z]{1,3})(\$?)([0-9]{1,5})(?![A-Za-z0-9_(])/g;
/** Same, but skips sheet-qualified refs (`Other!A1`) — structural edits on this sheet must not move refs into other sheets. */
const LOCAL_TOKEN_RE = /(?<![A-Za-z0-9_.!'])(\$?)([A-Za-z]{1,3})(\$?)([0-9]{1,5})(?![A-Za-z0-9_(])/g;

/** Rewrite every cell reference in a formula through `fn` (null → #REF!). Strings are left alone. */
export function remapFormula(f: string, fn: (c: number, r: number) => [number, number] | null): string {
  if (!f.startsWith('=')) return f;
  return f.replace(/"[^"]*"|[^"]+/g, (seg) => {
    if (seg.startsWith('"')) return seg;
    return seg.replace(LOCAL_TOKEN_RE, (m, d1: string, col: string, d2: string, row: string) => {
      const res = fn(colIdx(col), parseInt(row, 10) - 1);
      if (!res) return '#REF!';
      return `${d1}${colName(res[0])}${d2}${res[1] + 1}`;
    });
  });
}

/** Shift relative refs by (dc, dr) — used when copying / filling formulas. Absolute ($) parts stay. */
export function offsetFormula(f: string, dc: number, dr: number): string {
  if (!f.startsWith('=')) return f;
  return f.replace(/"[^"]*"|[^"]+/g, (seg) => {
    if (seg.startsWith('"')) return seg;
    return seg.replace(TOKEN_RE, (m, d1: string, col: string, d2: string, row: string) => {
      const c = colIdx(col) + (d1 ? 0 : dc);
      const r = parseInt(row, 10) - 1 + (d2 ? 0 : dr);
      if (c < 0 || r < 0) return '#REF!';
      return `${d1}${colName(c)}${d2}${r + 1}`;
    });
  });
}

/** Apply a coordinate mapping to every keyed record + formulas + structural extras. */
export function transformPage(pg: Page, map: (c: number, r: number) => [number, number] | null): Page {
  const move = <T,>(rec: Record<string, T> | undefined, fx?: (v: T) => T): Record<string, T> => {
    const out: Record<string, T> = {};
    if (!rec) return out;
    for (const [k, v] of Object.entries(rec)) {
      const [c, r] = parseRef(k);
      const n = map(c, r);
      if (!n) continue;
      out[refOf(n[0], n[1])] = fx ? fx(v) : v;
    }
    return out;
  };
  const cells = move(pg.cells, (v) => remapFormula(v, map));
  const styles = move(pg.styles);
  const notes = pg.notes ? move(pg.notes) : undefined;
  const merges = (pg.merges ?? [])
    .map((m) => {
      const rect = rangeToRect(m);
      if (!rect) return null;
      const a = map(rect.c1, rect.r1);
      const b = map(rect.c2, rect.r2);
      if (!a || !b) return null;
      const nr = { c1: a[0], r1: a[1], c2: b[0], r2: b[1] };
      return nr.c1 === nr.c2 && nr.r1 === nr.r2 ? null : rectToRange(nr);
    })
    .filter((x): x is string => !!x);
  const charts = (pg.charts ?? [])
    .map((ch) => {
      const rect = rangeToRect(ch.range);
      if (!rect) return ch;
      const a = map(rect.c1, rect.r1);
      const b = map(rect.c2, rect.r2);
      if (!a || !b) return ch;
      return { ...ch, range: rectToRange({ c1: a[0], r1: a[1], c2: b[0], r2: b[1] }) };
    });
  return { ...pg, cells, styles, notes, merges, charts };
}

function shiftIndexRecord(rec: Record<number, number> | undefined, at: number, delta: number): Record<number, number> | undefined {
  if (!rec) return rec;
  const out: Record<number, number> = {};
  for (const [k, v] of Object.entries(rec)) {
    const i = Number(k);
    if (delta < 0 && i === at) continue;
    out[i >= at ? i + delta : i] = v;
  }
  return out;
}

export function insertRows(pg: Page, at: number, n = 1): Page {
  const out = transformPage(pg, (c, r) => (r >= at ? (r + n < MAX_ROWS ? [c, r + n] : null) : [c, r]));
  out.rowH = shiftIndexRecord(pg.rowH, at, n);
  if (pg.freeze && pg.freeze.r > at) out.freeze = { ...pg.freeze, r: pg.freeze.r + n };
  return out;
}
export function deleteRows(pg: Page, at: number, n = 1): Page {
  const out = transformPage(pg, (c, r) => (r >= at && r < at + n ? null : r >= at + n ? [c, r - n] : [c, r]));
  let rowH = pg.rowH;
  for (let i = 0; i < n; i++) rowH = shiftIndexRecord(rowH, at, -1);
  out.rowH = rowH;
  if (pg.freeze && pg.freeze.r > at) out.freeze = { ...pg.freeze, r: Math.max(0, pg.freeze.r - n) };
  return out;
}
export function insertCols(pg: Page, at: number, n = 1): Page {
  const out = transformPage(pg, (c, r) => (c >= at ? (c + n < MAX_COLS ? [c + n, r] : null) : [c, r]));
  out.colW = shiftIndexRecord(pg.colW, at, n);
  if (pg.freeze && pg.freeze.c > at) out.freeze = { ...pg.freeze, c: pg.freeze.c + n };
  return out;
}
export function deleteCols(pg: Page, at: number, n = 1): Page {
  const out = transformPage(pg, (c, r) => (c >= at && c < at + n ? null : c >= at + n ? [c - n, r] : [c, r]));
  let colW = pg.colW;
  for (let i = 0; i < n; i++) colW = shiftIndexRecord(colW, at, -1);
  out.colW = colW;
  if (pg.freeze && pg.freeze.c > at) out.freeze = { ...pg.freeze, c: Math.max(0, pg.freeze.c - n) };
  return out;
}

// ---------------------------------------------------------------------------
// display formatting
// ---------------------------------------------------------------------------

const numFmtCache = new Map<string, Intl.NumberFormat>();
function nf(min: number, max: number): Intl.NumberFormat {
  const k = `${min}:${max}`;
  let f = numFmtCache.get(k);
  if (!f) {
    f = new Intl.NumberFormat(undefined, { minimumFractionDigits: min, maximumFractionDigits: max });
    numFmtCache.set(k, f);
  }
  return f;
}

export function parseNumberish(s: string): number | null {
  if (s === '' || s === undefined) return null;
  const t = s.trim();
  if (/^-?\d+(\.\d+)?%$/.test(t)) return parseFloat(t) / 100;
  const n = Number(t.replace(/[,$€£¥₹\s]/g, ''));
  return t !== '' && Number.isFinite(n) ? n : null;
}

/** Turn an evaluated display string into its formatted form according to the style. */
export function fmtDisplay(val: string, st: CellStyle | undefined, currency = '$'): string {
  if (!val) return '';
  const fmt = st?.fmt ?? 'gen';
  if (fmt === 'text') return val;
  if (val.startsWith('#')) return val;
  const n = parseNumberish(val);
  if (n === null) {
    if (fmt === 'date') {
      const d = new Date(val);
      return Number.isNaN(d.getTime()) ? val : d.toLocaleDateString();
    }
    return val;
  }
  const dec = st?.dec;
  switch (fmt) {
    case 'num':
      return nf(dec ?? 2, dec ?? 2).format(n);
    case 'cur': {
      const s = nf(dec ?? 2, dec ?? 2).format(Math.abs(n));
      return n < 0 ? `-${currency}${s}` : `${currency}${s}`;
    }
    case 'pct':
      return `${nf(dec ?? 0, dec ?? 2).format(n * 100)}%`;
    case 'date':
      return n > 0 && n < 2958466 ? serialToDate(n).toLocaleDateString() : val;
    case 'time': {
      const frac = n - Math.floor(n);
      const mins = Math.round(frac * 24 * 60);
      const h = Math.floor(mins / 60) % 24;
      const m = mins % 60;
      const d = new Date();
      d.setHours(h, m, 0, 0);
      return d.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
    }
    case 'sci':
      return n.toExponential(dec ?? 2);
    default:
      if (dec !== undefined) return nf(dec, dec).format(n);
      // general: trim float noise, group big integers lightly
      if (Number.isInteger(n)) return String(n);
      return String(parseFloat(n.toPrecision(10)));
  }
}

/** Evaluate every cell once (shared cache) → display strings keyed by ref. */
/** Prefix every unqualified ref in a formula with a sheet name (`=A1+B2` → `='Data'!A1+'Data'!B2`). */
export function qualifyFormula(f: string, sheetName: string): string {
  if (!f.startsWith('=')) return f;
  const prefix = /^[A-Za-z_][A-Za-z0-9_.]*$/.test(sheetName) ? `${sheetName}!` : `'${sheetName.replace(/'/g, "''")}'!`;
  return f.replace(/"[^"]*"|[^"]+/g, (seg) =>
    seg.startsWith('"') ? seg : seg.replace(LOCAL_TOKEN_RE, (m, _d1, _c, _d2, _r, offset: number, whole: string) => (/:\s*$/.test(whole.slice(0, offset)) ? m : prefix + m)),
  );
}

/**
 * Evaluate every cell of a page. Pass the whole `book` to resolve cross-sheet
 * references (`Sheet2!A1`, `'My Sheet'!B2:B9`). Formulas that live on another
 * sheet are handed to the engine with their bare refs qualified, so they
 * evaluate against their own sheet while sharing one cache / cycle stack.
 */
export function computeDisplay(cells: Record<string, string>, book?: Book, active?: string): Record<string, string> {
  const get: CellGetter = (r) => {
    const bang = r.indexOf('!');
    if (bang < 0) return cells[r] ?? '';
    const sheet = r.slice(0, bang);
    const ref = r.slice(bang + 1);
    if (!book) return '#REF!';
    const name = sheet === active ? active : book.order.find((n) => n === sheet || n.toLowerCase() === sheet.toLowerCase());
    if (!name) return '#REF!';
    if (name === active) return cells[ref] ?? '';
    const raw = book.sheets[name].cells[ref] ?? '';
    return raw.startsWith('=') ? qualifyFormula(raw, name) : raw;
  };
  const opts: EvalOptions = { cache: new Map() };
  const out: Record<string, string> = {};
  for (const k of Object.keys(cells)) out[k] = evalCell(cells[k], get, opts);
  return out;
}

// ---------------------------------------------------------------------------
// fill / series
// ---------------------------------------------------------------------------

/** Values for a fill-series operation given the seed values (1..2 numbers → step, text → copy, formulas → offset). */
export function seriesFill(seeds: string[], count: number, dir: 'down' | 'right'): string[] {
  const out: string[] = [];
  const nums = seeds.map(parseNumberish);
  const allNum = nums.every((n) => n !== null) && seeds.every((s) => !s.startsWith('='));
  if (allNum && seeds.length >= 1) {
    const step = seeds.length >= 2 ? (nums[nums.length - 1] as number) - (nums[nums.length - 2] as number) : 1;
    let last = nums[nums.length - 1] as number;
    for (let i = 0; i < count; i++) {
      last += step;
      out.push(String(parseFloat(last.toPrecision(12))));
    }
    return out;
  }
  // text ending with a number: "Item 1" → "Item 2"
  const m = /^(.*?)(\d+)$/.exec(seeds[seeds.length - 1]);
  if (m && !seeds[seeds.length - 1].startsWith('=')) {
    let n = parseInt(m[2], 10);
    for (let i = 0; i < count; i++) { n++; out.push(`${m[1]}${n}`); }
    return out;
  }
  // month / weekday names
  const months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  const days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
  for (const list of [months, days, months.map((x) => x.slice(0, 3)), days.map((x) => x.slice(0, 3))]) {
    const idx = list.findIndex((x) => x.toLowerCase() === seeds[seeds.length - 1].toLowerCase());
    if (idx >= 0) {
      for (let i = 1; i <= count; i++) out.push(list[(idx + i) % list.length]);
      return out;
    }
  }
  // formulas / plain copy → repeat pattern with offset
  for (let i = 0; i < count; i++) {
    const src = seeds[i % seeds.length];
    const k = Math.floor(i / seeds.length) + 1;
    out.push(src.startsWith('=') ? offsetFormula(src, dir === 'right' ? k * seeds.length : 0, dir === 'down' ? k * seeds.length : 0) : src);
  }
  return out;
}

// ---------------------------------------------------------------------------
// CSV / TSV
// ---------------------------------------------------------------------------

export function parseDelimited(text: string, delim?: string): string[][] {
  const d = delim ?? (text.includes('\t') ? '\t' : ',');
  const rows: string[][] = [];
  let row: string[] = [];
  let cur = '';
  let q = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (q) {
      if (ch === '"') {
        if (text[i + 1] === '"') { cur += '"'; i++; } else q = false;
      } else cur += ch;
    } else if (ch === '"') q = true;
    else if (ch === d) { row.push(cur); cur = ''; }
    else if (ch === '\n' || ch === '\r') {
      if (ch === '\r' && text[i + 1] === '\n') i++;
      row.push(cur); cur = '';
      rows.push(row); row = [];
    } else cur += ch;
  }
  if (cur !== '' || row.length) { row.push(cur); rows.push(row); }
  return rows.filter((r) => r.some((c) => c !== ''));
}

export function toCsv(rows: string[][]): string {
  return rows.map((r) => r.map((v) => (/[",\n]/.test(v) ? `"${v.replace(/"/g, '""')}"` : v)).join(',')).join('\n');
}

// ---------------------------------------------------------------------------
// book helpers
// ---------------------------------------------------------------------------

export function emptyPage(): Page {
  return { cells: {}, styles: {} };
}

export function freshBook(): Book {
  return { order: ['Sheet1'], active: 'Sheet1', sheets: { Sheet1: emptyPage() } };
}

export function cloneBook(b: Book): Book {
  return JSON.parse(JSON.stringify(b)) as Book;
}

export function uniqueSheetName(order: string[], base: string): string {
  if (!order.includes(base)) return base;
  let i = 2;
  while (order.includes(`${base}${i}`)) i++;
  return `${base}${i}`;
}

export function cleanSheetName(name: string): string {
  return name.replace(/[\\/?*[\]:]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 31);
}

/** Starter workbooks (Excel mobile shows these on "New"). */
export const SHEET_TEMPLATES: { id: string; label: string; desc: string; build: () => Page }[] = [
  { id: 'blank', label: 'Blank workbook', desc: 'Empty grid', build: emptyPage },
  {
    id: 'budget',
    label: 'Monthly budget',
    desc: 'Income, expenses, balance',
    build: () => {
      const cells: Record<string, string> = {
        A1: 'Category', B1: 'Planned', C1: 'Actual', D1: 'Difference',
        A2: 'Salary', B2: '3000', C2: '3000', D2: '=C2-B2',
        A3: 'Rent', B3: '-900', C3: '-900', D3: '=C3-B3',
        A4: 'Groceries', B4: '-350', C4: '-410', D4: '=C4-B4',
        A5: 'Transport', B5: '-120', C5: '-95', D5: '=C5-B5',
        A6: 'Utilities', B6: '-150', C6: '-162', D6: '=C6-B6',
        A7: 'Fun', B7: '-200', C7: '-260', D7: '=C7-B7',
        A9: 'Balance', B9: '=SUM(B2:B7)', C9: '=SUM(C2:C7)', D9: '=SUM(D2:D7)',
        A11: 'Savings rate', B11: '=B9/B2', C11: '=C9/C2',
      };
      const styles: Record<string, CellStyle> = {};
      ['A1', 'B1', 'C1', 'D1'].forEach((k) => (styles[k] = { b: true, fill: '#DEEBF7', bb: true }));
      ['A9', 'B9', 'C9', 'D9'].forEach((k) => (styles[k] = { b: true, bt: true }));
      for (let r = 2; r <= 9; r++) for (const c of ['B', 'C', 'D']) styles[`${c}${r}`] = { ...styles[`${c}${r}`], fmt: 'cur' };
      styles.B11 = { fmt: 'pct' }; styles.C11 = { fmt: 'pct' }; styles.A11 = { b: true };
      return { cells, styles, colW: { 0: 120 }, freeze: { r: 1, c: 0 } };
    },
  },
  {
    id: 'invoice',
    label: 'Invoice',
    desc: 'Items, quantities, tax, total',
    build: () => {
      const cells: Record<string, string> = {
        A1: 'INVOICE', A2: 'Bill to:', B2: 'Customer name', A3: 'Date:', B3: '=TODAY()', A4: 'Invoice #:', B4: '1001',
        A6: 'Item', B6: 'Qty', C6: 'Unit price', D6: 'Amount',
        A7: 'Design work', B7: '10', C7: '45', D7: '=B7*C7',
        A8: 'Hosting (12 mo)', B8: '1', C8: '120', D8: '=B8*C8',
        A9: 'Support', B9: '3', C9: '60', D9: '=B9*C9',
        C11: 'Subtotal', D11: '=SUM(D7:D10)',
        C12: 'Tax rate', D12: '0.18',
        C13: 'Tax', D13: '=D11*D12',
        C14: 'TOTAL', D14: '=D11+D13',
      };
      const styles: Record<string, CellStyle> = { A1: { b: true, size: 22 }, B3: { fmt: 'date' }, D12: { fmt: 'pct' }, C14: { b: true }, D14: { b: true, fmt: 'cur', bt: true } };
      ['A6', 'B6', 'C6', 'D6'].forEach((k) => (styles[k] = { b: true, fill: '#E2F0D9', bb: true }));
      ['C7', 'C8', 'C9', 'D7', 'D8', 'D9', 'D11', 'D13'].forEach((k) => (styles[k] = { ...styles[k], fmt: 'cur' }));
      ['C11', 'C12', 'C13'].forEach((k) => (styles[k] = { align: 'right', color: '#555555' }));
      return { cells, styles, colW: { 0: 140 }, merges: ['A1:D1'] };
    },
  },
  {
    id: 'tracker',
    label: 'Task tracker',
    desc: 'Tasks, owners, status, due dates',
    build: () => {
      const cells: Record<string, string> = {
        A1: 'Task', B1: 'Owner', C1: 'Status', D1: 'Due', E1: 'Days left',
        A2: 'Write proposal', B2: 'Ana', C2: 'In progress', D2: '=TODAY()+3', E2: '=D2-TODAY()',
        A3: 'Review budget', B3: 'Ravi', C3: 'Not started', D3: '=TODAY()+7', E3: '=D3-TODAY()',
        A4: 'Send invoices', B4: 'Mei', C4: 'Done', D4: '=TODAY()-1', E4: '=D4-TODAY()',
        A6: 'Done', B6: '=COUNTIF(C2:C4,"Done")', A7: 'Open', B7: '=COUNTA(A2:A4)-B6', A8: 'Progress', B8: '=B6/COUNTA(A2:A4)',
      };
      const styles: Record<string, CellStyle> = { D2: { fmt: 'date' }, D3: { fmt: 'date' }, D4: { fmt: 'date' }, B8: { fmt: 'pct' }, A6: { b: true }, A7: { b: true }, A8: { b: true } };
      ['A1', 'B1', 'C1', 'D1', 'E1'].forEach((k) => (styles[k] = { b: true, fill: '#FFF2CC', bb: true }));
      styles.C4 = { fill: '#E2F0D9' };
      return { cells, styles, colW: { 0: 150, 2: 110 }, freeze: { r: 1, c: 0 }, filter: { range: 'A1:E4', criteria: {} } };
    },
  },
  {
    id: 'grades',
    label: 'Grade book',
    desc: 'Scores, averages, letter grades',
    build: () => {
      const cells: Record<string, string> = {
        A1: 'Student', B1: 'Test 1', C1: 'Test 2', D1: 'Test 3', E1: 'Average', F1: 'Grade',
        A2: 'Aarav', B2: '78', C2: '85', D2: '90', A3: 'Diya', B3: '92', C3: '88', D3: '95', A4: 'Kabir', B4: '60', C4: '72', D4: '68', A5: 'Sara', B5: '85', C5: '79', D5: '91',
        A7: 'Class average', B7: '=AVERAGE(B2:B5)', C7: '=AVERAGE(C2:C5)', D7: '=AVERAGE(D2:D5)', E7: '=AVERAGE(E2:E5)',
        A8: 'Highest', E8: '=MAX(E2:E5)', A9: 'Lowest', E9: '=MIN(E2:E5)',
      };
      for (let r = 2; r <= 5; r++) {
        cells[`E${r}`] = `=AVERAGE(B${r}:D${r})`;
        cells[`F${r}`] = `=IF(E${r}>=90,"A",IF(E${r}>=80,"B",IF(E${r}>=70,"C",IF(E${r}>=60,"D","F"))))`;
      }
      const styles: Record<string, CellStyle> = {};
      ['A1', 'B1', 'C1', 'D1', 'E1', 'F1'].forEach((k) => (styles[k] = { b: true, fill: '#E9D7F1', bb: true, align: 'center' }));
      for (let r = 2; r <= 9; r++) styles[`E${r}`] = { fmt: 'num', dec: 1, b: r > 6 };
      ['A7', 'A8', 'A9'].forEach((k) => (styles[k] = { b: true }));
      for (let r = 2; r <= 5; r++) styles[`F${r}`] = { align: 'center', b: true };
      return { cells, styles, colW: { 0: 110 }, freeze: { r: 1, c: 1 } };
    },
  },
  {
    id: 'loan',
    label: 'Loan calculator',
    desc: 'PMT, total interest, schedule',
    build: () => {
      const cells: Record<string, string> = {
        A1: 'Loan amount', B1: '250000', A2: 'Annual rate', B2: '0.085', A3: 'Years', B3: '5',
        A5: 'Monthly payment', B5: '=-PMT(B2/12,B3*12,B1)', A6: 'Total paid', B6: '=B5*B3*12', A7: 'Total interest', B7: '=B6-B1',
        A9: 'Month', B9: 'Payment', C9: 'Interest', D9: 'Principal', E9: 'Balance',
      };
      for (let i = 1; i <= 12; i++) {
        const r = 9 + i;
        cells[`A${r}`] = String(i);
        cells[`B${r}`] = '=$B$5';
        cells[`C${r}`] = i === 1 ? '=$B$1*$B$2/12' : `=E${r - 1}*$B$2/12`;
        cells[`D${r}`] = `=B${r}-C${r}`;
        cells[`E${r}`] = i === 1 ? `=$B$1-D${r}` : `=E${r - 1}-D${r}`;
      }
      const styles: Record<string, CellStyle> = { B1: { fmt: 'cur' }, B2: { fmt: 'pct' }, B5: { fmt: 'cur', b: true }, B6: { fmt: 'cur' }, B7: { fmt: 'cur' } };
      ['A1', 'A2', 'A3', 'A5', 'A6', 'A7'].forEach((k) => (styles[k] = { b: true }));
      ['A9', 'B9', 'C9', 'D9', 'E9'].forEach((k) => (styles[k] = { b: true, fill: '#DEEBF7', bb: true }));
      for (let r = 10; r <= 21; r++) for (const c of ['B', 'C', 'D', 'E']) styles[`${c}${r}`] = { fmt: 'cur' };
      return { cells, styles, colW: { 0: 130 } };
    },
  },
];
