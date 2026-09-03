/**
 * Spreadsheet formula engine (no dependencies).
 *
 * Values: numbers, strings, booleans, errors. Cell refs A1..Z999 (+ absolute
 * $A$1), ranges A1:B9, arithmetic + - * / ^ %, comparison = <> < > <= >=,
 * string concat &, and ~60 Excel-compatible functions (math, stats, text,
 * logic, lookup, date). Errors follow Excel: #DIV/0!, #VALUE!, #REF!, #NAME?,
 * #N/A, #NUM!, #CIRC!.
 */

export type CellGetter = (ref: string) => string;

export type Val = number | string | boolean | FormulaError;

export class FormulaError extends Error {
  code: string;
  constructor(code: string) {
    super(code);
    this.code = code;
  }
}

const ERR = {
  div0: () => new FormulaError('#DIV/0!'),
  value: () => new FormulaError('#VALUE!'),
  ref: () => new FormulaError('#REF!'),
  name: () => new FormulaError('#NAME?'),
  na: () => new FormulaError('#N/A'),
  num: () => new FormulaError('#NUM!'),
  circ: () => new FormulaError('#CIRC!'),
};

export function colIdx(col: string): number {
  let n = 0;
  for (const c of col.toUpperCase()) n = n * 26 + (c.charCodeAt(0) - 64);
  return n - 1;
}

export function colName(i: number): string {
  let s = '';
  let n = i + 1;
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/**
 * Optional sheet qualifier in front of a reference: `Sheet1!A1`, `'My Sheet'!A1:B2`.
 * Qualified refs are passed to the CellGetter as `Sheet!A1` (sheet name unquoted, cell uppercased).
 */
const SHEET_PREFIX = "(?:(?:'((?:[^']|'')+)'|([A-Za-z_][A-Za-z0-9_.]*))!)?";
const sheetOf = (quoted: string | undefined, bare: string | undefined): string => (quoted ? quoted.replace(/''/g, "'") : bare ?? '');
const qual = (sheet: string, ref: string): string => (sheet ? `${sheet}!${ref}` : ref);

/** All A1-style refs in a formula (uppercased, without $; sheet-qualified as `Sheet!A1`), for dependency scans. */
export function refsIn(formula: string): string[] {
  const out: string[] = [];
  const re = new RegExp(`${SHEET_PREFIX}\\$?([A-Za-z]{1,2})\\$?(\\d{1,4})(?::\\$?([A-Za-z]{1,2})\\$?(\\d{1,4}))?`, 'g');
  let m: RegExpExecArray | null;
  const src = formula.replace(/"[^"]*"/g, '""');
  while ((m = re.exec(src))) {
    const sheet = sheetOf(m[1], m[2]);
    if (m[5]) {
      const c1 = colIdx(m[3]);
      const c2 = colIdx(m[5]);
      const r1 = parseInt(m[4], 10);
      const r2 = parseInt(m[6], 10);
      for (let c = Math.min(c1, c2); c <= Math.max(c1, c2); c++)
        for (let r = Math.min(r1, r2); r <= Math.max(r1, r2); r++) out.push(qual(sheet, colName(c) + r));
    } else out.push(qual(sheet, m[3].toUpperCase() + m[4]));
  }
  return out;
}

// ---------------------------------------------------------------------------
// coercion helpers
// ---------------------------------------------------------------------------

function isErr(v: Val | Val[]): v is FormulaError {
  return v instanceof FormulaError;
}

function toNum(v: Val): number {
  if (isErr(v)) throw v;
  if (typeof v === 'number') return v;
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (v.trim() === '') return 0;
  const cleaned = v.replace(/[,$%\s]/g, '');
  const n = Number(cleaned);
  if (!Number.isFinite(n)) throw ERR.value();
  return v.trim().endsWith('%') ? n / 100 : n;
}

function toStr(v: Val): string {
  if (isErr(v)) throw v;
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (typeof v === 'number') return fmtNum(v);
  return v;
}

function toBool(v: Val): boolean {
  if (isErr(v)) throw v;
  if (typeof v === 'boolean') return v;
  if (typeof v === 'number') return v !== 0;
  const u = v.trim().toUpperCase();
  if (u === 'TRUE') return true;
  if (u === 'FALSE' || u === '') return false;
  const n = Number(u);
  if (Number.isFinite(n)) return n !== 0;
  throw ERR.value();
}

/** Parse a raw cell string into a typed value (numbers/booleans recognised). */
const ERROR_CODES = new Set(['#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#N/A', '#NUM!', '#CIRC!', '#NULL!']);

function parseRaw(raw: string): Val {
  if (raw === '') return '';
  const t = raw.trim();
  // a referenced cell that already *holds* an error (e.g. an evaluated value from another sheet) propagates it
  if (ERROR_CODES.has(t)) return new FormulaError(t);
  if (/^-?\d+(\.\d+)?$/.test(t) || /^-?\d*\.\d+$/.test(t)) return parseFloat(t);
  if (/^-?\d+(\.\d+)?%$/.test(t)) return parseFloat(t) / 100;
  if (/^-?\$\d[\d,]*(\.\d+)?$/.test(t)) return parseFloat(t.replace(/[$,]/g, ''));
  const u = t.toUpperCase();
  if (u === 'TRUE') return true;
  if (u === 'FALSE') return false;
  return raw;
}

/** Flatten function args: ranges are arrays; numbers extracted for math funcs. */
function numsOf(args: Val[][]): number[] {
  const out: number[] = [];
  for (const a of args) {
    for (const v of a) {
      if (isErr(v)) throw v;
      if (typeof v === 'number') out.push(v);
      else if (typeof v === 'boolean') out.push(v ? 1 : 0);
      else if (typeof v === 'string' && v.trim() !== '') {
        const n = Number(v.replace(/[,$\s]/g, ''));
        if (Number.isFinite(n)) out.push(n);
      }
    }
  }
  return out;
}

function fmtNum(v: number): string {
  if (!Number.isFinite(v)) return '#NUM!';
  if (Number.isInteger(v)) return String(v);
  return String(parseFloat(v.toPrecision(12)));
}

// ---------------------------------------------------------------------------
// dates (Excel serial numbers, 1900 system)
// ---------------------------------------------------------------------------

const DAY_MS = 86400000;
const EPOCH = Date.UTC(1899, 11, 30);

export function dateToSerial(d: Date): number {
  return (Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()) - EPOCH) / DAY_MS;
}

export function serialToDate(n: number): Date {
  return new Date(EPOCH + Math.floor(n) * DAY_MS);
}

function serialOf(v: Val): number {
  if (typeof v === 'string' && v.trim() !== '' && !/^-?[\d.]+$/.test(v.trim())) {
    const d = new Date(v);
    if (!Number.isNaN(d.getTime())) return dateToSerial(d);
    throw ERR.value();
  }
  return toNum(v);
}

// ---------------------------------------------------------------------------
// criteria matching for COUNTIF / SUMIF / AVERAGEIF
// ---------------------------------------------------------------------------

function matcher(crit: Val): (v: Val) => boolean {
  if (typeof crit === 'number' || typeof crit === 'boolean') {
    return (v) => !isErr(v) && v !== '' && (v === crit || (typeof v === 'string' && Number(v) === crit));
  }
  if (isErr(crit)) throw crit;
  const m = /^(<>|>=|<=|=|>|<)?(.*)$/.exec(crit.trim());
  const op = m?.[1] ?? '=';
  const rhs = m?.[2] ?? '';
  const rhsNum = rhs.trim() === '' ? NaN : Number(rhs);
  if (Number.isFinite(rhsNum)) {
    return (v) => {
      if (isErr(v) || v === '') return false;
      const n = typeof v === 'number' ? v : Number(v);
      if (!Number.isFinite(n)) return op === '<>';
      switch (op) {
        case '>': return n > rhsNum;
        case '<': return n < rhsNum;
        case '>=': return n >= rhsNum;
        case '<=': return n <= rhsNum;
        case '<>': return n !== rhsNum;
        default: return n === rhsNum;
      }
    };
  }
  const re = new RegExp(`^${rhs.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.')}$`, 'i');
  return (v) => {
    const s = isErr(v) ? '' : toStr(v);
    const hit = re.test(s);
    return op === '<>' ? !hit : hit;
  };
}

// ---------------------------------------------------------------------------
// function library
// ---------------------------------------------------------------------------

type RawArg = Val | Val[];
type Fn = (args: Val[][], raw: RawArg[]) => Val;

const single = (a: Val[][], i: number): Val => {
  const arr = a[i];
  if (!arr || arr.length === 0) return '';
  return arr[0];
};

const FUNCS: Record<string, Fn> = {
  // ---- math ----
  SUM: (a) => numsOf(a).reduce((x, y) => x + y, 0),
  PRODUCT: (a) => numsOf(a).reduce((x, y) => x * y, 1),
  AVERAGE: (a) => {
    const n = numsOf(a);
    if (!n.length) throw ERR.div0();
    return n.reduce((x, y) => x + y, 0) / n.length;
  },
  AVG: (a) => FUNCS.AVERAGE(a, []),
  MEDIAN: (a) => {
    const n = numsOf(a).sort((x, y) => x - y);
    if (!n.length) throw ERR.num();
    const mid = Math.floor(n.length / 2);
    return n.length % 2 ? n[mid] : (n[mid - 1] + n[mid]) / 2;
  },
  MODE: (a) => {
    const n = numsOf(a);
    const c = new Map<number, number>();
    let best: number | null = null;
    let bc = 0;
    for (const v of n) {
      const k = (c.get(v) ?? 0) + 1;
      c.set(v, k);
      if (k > bc) { bc = k; best = v; }
    }
    if (best === null || bc < 2) throw ERR.na();
    return best;
  },
  COUNT: (a) => numsOf(a).length,
  COUNTA: (a) => a.flat().filter((v) => !isErr(v) && v !== '').length,
  COUNTBLANK: (a) => a.flat().filter((v) => v === '').length,
  MIN: (a) => { const n = numsOf(a); return n.length ? Math.min(...n) : 0; },
  MAX: (a) => { const n = numsOf(a); return n.length ? Math.max(...n) : 0; },
  ABS: (a) => Math.abs(toNum(single(a, 0))),
  SQRT: (a) => { const n = toNum(single(a, 0)); if (n < 0) throw ERR.num(); return Math.sqrt(n); },
  POWER: (a) => Math.pow(toNum(single(a, 0)), toNum(single(a, 1))),
  EXP: (a) => Math.exp(toNum(single(a, 0))),
  LN: (a) => { const n = toNum(single(a, 0)); if (n <= 0) throw ERR.num(); return Math.log(n); },
  LOG: (a) => {
    const n = toNum(single(a, 0));
    const b = a[1] ? toNum(single(a, 1)) : 10;
    if (n <= 0 || b <= 0) throw ERR.num();
    return Math.log(n) / Math.log(b);
  },
  LOG10: (a) => { const n = toNum(single(a, 0)); if (n <= 0) throw ERR.num(); return Math.log10(n); },
  MOD: (a) => {
    const n = toNum(single(a, 0));
    const d = toNum(single(a, 1));
    if (d === 0) throw ERR.div0();
    return n - d * Math.floor(n / d);
  },
  INT: (a) => Math.floor(toNum(single(a, 0))),
  TRUNC: (a) => {
    const n = toNum(single(a, 0));
    const d = a[1] ? toNum(single(a, 1)) : 0;
    const f = 10 ** d;
    return Math.trunc(n * f) / f;
  },
  ROUND: (a) => {
    const n = toNum(single(a, 0));
    const d = a[1] ? toNum(single(a, 1)) : 0;
    const f = 10 ** d;
    return Math.round((n + Number.EPSILON * Math.sign(n)) * f) / f;
  },
  ROUNDUP: (a) => {
    const n = toNum(single(a, 0));
    const d = a[1] ? toNum(single(a, 1)) : 0;
    const f = 10 ** d;
    return (n < 0 ? -1 : 1) * Math.ceil(Math.abs(n) * f - 1e-9) / f;
  },
  ROUNDDOWN: (a) => {
    const n = toNum(single(a, 0));
    const d = a[1] ? toNum(single(a, 1)) : 0;
    const f = 10 ** d;
    return (n < 0 ? -1 : 1) * Math.floor(Math.abs(n) * f + 1e-9) / f;
  },
  CEILING: (a) => {
    const n = toNum(single(a, 0));
    const s = a[1] ? toNum(single(a, 1)) : 1;
    if (s === 0) return 0;
    return Math.ceil(n / s) * s;
  },
  FLOOR: (a) => {
    const n = toNum(single(a, 0));
    const s = a[1] ? toNum(single(a, 1)) : 1;
    if (s === 0) return 0;
    return Math.floor(n / s) * s;
  },
  SIGN: (a) => Math.sign(toNum(single(a, 0))),
  PI: () => Math.PI,
  RAND: () => Math.random(),
  RANDBETWEEN: (a) => {
    const lo = Math.ceil(toNum(single(a, 0)));
    const hi = Math.floor(toNum(single(a, 1)));
    return lo + Math.floor(Math.random() * (hi - lo + 1));
  },
  SIN: (a) => Math.sin(toNum(single(a, 0))),
  COS: (a) => Math.cos(toNum(single(a, 0))),
  TAN: (a) => Math.tan(toNum(single(a, 0))),
  RADIANS: (a) => (toNum(single(a, 0)) * Math.PI) / 180,
  DEGREES: (a) => (toNum(single(a, 0)) * 180) / Math.PI,
  SUMPRODUCT: (a) => {
    const cols = a.map((c) => c.map((v) => (v === '' ? 0 : toNum(v))));
    const len = cols[0]?.length ?? 0;
    if (cols.some((c) => c.length !== len)) throw ERR.value();
    let s = 0;
    for (let i = 0; i < len; i++) s += cols.reduce((p, c) => p * c[i], 1);
    return s;
  },
  SUMSQ: (a) => numsOf(a).reduce((x, y) => x + y * y, 0),
  STDEV: (a) => {
    const n = numsOf(a);
    if (n.length < 2) throw ERR.div0();
    const m = n.reduce((x, y) => x + y, 0) / n.length;
    return Math.sqrt(n.reduce((x, y) => x + (y - m) ** 2, 0) / (n.length - 1));
  },
  STDEVP: (a) => {
    const n = numsOf(a);
    if (!n.length) throw ERR.div0();
    const m = n.reduce((x, y) => x + y, 0) / n.length;
    return Math.sqrt(n.reduce((x, y) => x + (y - m) ** 2, 0) / n.length);
  },
  VAR: (a) => {
    const n = numsOf(a);
    if (n.length < 2) throw ERR.div0();
    const m = n.reduce((x, y) => x + y, 0) / n.length;
    return n.reduce((x, y) => x + (y - m) ** 2, 0) / (n.length - 1);
  },
  LARGE: (a) => {
    const n = numsOf([a[0] ?? []]).sort((x, y) => y - x);
    const k = toNum(single(a, 1));
    if (k < 1 || k > n.length) throw ERR.num();
    return n[k - 1];
  },
  SMALL: (a) => {
    const n = numsOf([a[0] ?? []]).sort((x, y) => x - y);
    const k = toNum(single(a, 1));
    if (k < 1 || k > n.length) throw ERR.num();
    return n[k - 1];
  },
  RANK: (a) => {
    const v = toNum(single(a, 0));
    const n = numsOf([a[1] ?? []]);
    const asc = a[2] ? toBool(single(a, 2)) : false;
    const sorted = [...n].sort((x, y) => (asc ? x - y : y - x));
    const i = sorted.indexOf(v);
    if (i < 0) throw ERR.na();
    return i + 1;
  },
  // ---- conditional aggregates ----
  COUNTIF: (a) => {
    const test = matcher(single(a, 1));
    return (a[0] ?? []).filter(test).length;
  },
  SUMIF: (a) => {
    const test = matcher(single(a, 1));
    const range = a[0] ?? [];
    const sumRange = a[2] ?? range;
    let s = 0;
    range.forEach((v, i) => {
      if (test(v)) {
        const sv = sumRange[i];
        if (typeof sv === 'number') s += sv;
        else if (typeof sv === 'string' && sv.trim() !== '' && Number.isFinite(Number(sv))) s += Number(sv);
      }
    });
    return s;
  },
  AVERAGEIF: (a) => {
    const test = matcher(single(a, 1));
    const range = a[0] ?? [];
    const avgRange = a[2] ?? range;
    let s = 0;
    let c = 0;
    range.forEach((v, i) => {
      if (test(v)) {
        const sv = avgRange[i];
        const n = typeof sv === 'number' ? sv : Number(sv);
        if (sv !== '' && Number.isFinite(n)) { s += n; c++; }
      }
    });
    if (!c) throw ERR.div0();
    return s / c;
  },
  COUNTIFS: (a) => {
    if (a.length < 2 || a.length % 2) throw ERR.value();
    const len = a[0].length;
    let c = 0;
    for (let i = 0; i < len; i++) {
      let ok = true;
      for (let p = 0; p < a.length; p += 2) {
        if (!matcher(single(a, p + 1))(a[p][i] ?? '')) { ok = false; break; }
      }
      if (ok) c++;
    }
    return c;
  },
  SUMIFS: (a) => {
    if (a.length < 3 || a.length % 2 === 0) throw ERR.value();
    const sumRange = a[0];
    let s = 0;
    for (let i = 0; i < sumRange.length; i++) {
      let ok = true;
      for (let p = 1; p < a.length; p += 2) {
        if (!matcher(single(a, p + 1))(a[p][i] ?? '')) { ok = false; break; }
      }
      if (ok) {
        const sv = sumRange[i];
        const n = typeof sv === 'number' ? sv : Number(sv);
        if (sv !== '' && Number.isFinite(n)) s += n;
      }
    }
    return s;
  },
  // ---- logic ----
  IF: (a) => (toBool(single(a, 0)) ? (a[1] ? single(a, 1) : true) : a[2] ? single(a, 2) : false),
  IFS: (a) => {
    for (let i = 0; i + 1 < a.length; i += 2) if (toBool(single(a, i))) return single(a, i + 1);
    throw ERR.na();
  },
  AND: (a) => a.flat().filter((v) => v !== '').every(toBool),
  OR: (a) => a.flat().filter((v) => v !== '').some(toBool),
  XOR: (a) => a.flat().filter((v) => v !== '').filter(toBool).length % 2 === 1,
  NOT: (a) => !toBool(single(a, 0)),
  TRUE: () => true,
  FALSE: () => false,
  IFERROR: (a, raw) => (isErr(raw[0]) ? single(a, 1) : single(a, 0)),
  IFNA: (a, raw) => (isErr(raw[0]) && (raw[0] as FormulaError).code === '#N/A' ? single(a, 1) : single(a, 0)),
  ISBLANK: (a) => single(a, 0) === '',
  ISNUMBER: (a) => typeof single(a, 0) === 'number',
  ISTEXT: (a) => typeof single(a, 0) === 'string' && single(a, 0) !== '',
  ISERROR: (a, raw) => isErr(raw[0]),
  ISEVEN: (a) => Math.floor(toNum(single(a, 0))) % 2 === 0,
  ISODD: (a) => Math.abs(Math.floor(toNum(single(a, 0)))) % 2 === 1,
  SWITCH: (a) => {
    const v = single(a, 0);
    for (let i = 1; i + 1 < a.length; i += 2) if (eq(v, single(a, i))) return single(a, i + 1);
    if ((a.length - 1) % 2 === 1) return single(a, a.length - 1);
    throw ERR.na();
  },
  CHOOSE: (a) => {
    const i = Math.floor(toNum(single(a, 0)));
    if (i < 1 || i >= a.length) throw ERR.value();
    return single(a, i);
  },
  // ---- text ----
  CONCATENATE: (a) => a.flat().map(toStr).join(''),
  CONCAT: (a) => a.flat().map(toStr).join(''),
  TEXTJOIN: (a) => {
    const sep = toStr(single(a, 0));
    const skip = toBool(single(a, 1));
    return a.slice(2).flat().filter((v) => !(skip && v === '')).map(toStr).join(sep);
  },
  LEN: (a) => toStr(single(a, 0)).length,
  UPPER: (a) => toStr(single(a, 0)).toUpperCase(),
  LOWER: (a) => toStr(single(a, 0)).toLowerCase(),
  PROPER: (a) => toStr(single(a, 0)).toLowerCase().replace(/(^|[^a-z'])([a-z])/g, (_, p, c: string) => p + c.toUpperCase()),
  TRIM: (a) => toStr(single(a, 0)).trim().replace(/\s+/g, ' '),
  CLEAN: (a) => toStr(single(a, 0)).replace(/[\x00-\x1f]/g, ''),
  LEFT: (a) => toStr(single(a, 0)).slice(0, a[1] ? toNum(single(a, 1)) : 1),
  RIGHT: (a) => {
    const s = toStr(single(a, 0));
    const n = a[1] ? toNum(single(a, 1)) : 1;
    return n <= 0 ? '' : s.slice(-n);
  },
  MID: (a) => {
    const s = toStr(single(a, 0));
    const start = toNum(single(a, 1));
    const n = toNum(single(a, 2));
    if (start < 1 || n < 0) throw ERR.value();
    return s.substr(start - 1, n);
  },
  FIND: (a) => {
    const needle = toStr(single(a, 0));
    const hay = toStr(single(a, 1));
    const start = a[2] ? toNum(single(a, 2)) : 1;
    const i = hay.indexOf(needle, start - 1);
    if (i < 0) throw ERR.value();
    return i + 1;
  },
  SEARCH: (a) => {
    const needle = toStr(single(a, 0)).toLowerCase();
    const hay = toStr(single(a, 1)).toLowerCase();
    const start = a[2] ? toNum(single(a, 2)) : 1;
    const i = hay.indexOf(needle, start - 1);
    if (i < 0) throw ERR.value();
    return i + 1;
  },
  SUBSTITUTE: (a) => {
    const s = toStr(single(a, 0));
    const from = toStr(single(a, 1));
    const to = toStr(single(a, 2));
    if (!from) return s;
    if (a[3]) {
      const nth = toNum(single(a, 3));
      let idx = -1;
      for (let k = 0; k < nth; k++) {
        idx = s.indexOf(from, idx + 1);
        if (idx < 0) return s;
      }
      return s.slice(0, idx) + to + s.slice(idx + from.length);
    }
    return s.split(from).join(to);
  },
  REPLACE: (a) => {
    const s = toStr(single(a, 0));
    const start = toNum(single(a, 1));
    const n = toNum(single(a, 2));
    return s.slice(0, start - 1) + toStr(single(a, 3)) + s.slice(start - 1 + n);
  },
  REPT: (a) => toStr(single(a, 0)).repeat(Math.max(0, Math.floor(toNum(single(a, 1))))),
  EXACT: (a) => toStr(single(a, 0)) === toStr(single(a, 1)),
  VALUE: (a) => toNum(single(a, 0)),
  N: (a) => { const v = single(a, 0); return typeof v === 'number' ? v : typeof v === 'boolean' ? (v ? 1 : 0) : 0; },
  T: (a) => { const v = single(a, 0); return typeof v === 'string' ? v : ''; },
  TEXT: (a) => {
    const v = single(a, 0);
    const fmt = toStr(single(a, 1));
    return formatText(v, fmt);
  },
  CHAR: (a) => String.fromCharCode(toNum(single(a, 0))),
  CODE: (a) => { const s = toStr(single(a, 0)); if (!s) throw ERR.value(); return s.charCodeAt(0); },
  // ---- lookup ----
  VLOOKUP: (a, _raw) => {
    const key = single(a, 0);
    const table = a[1] ?? [];
    const colN = Math.floor(toNum(single(a, 2)));
    const approx = a[3] ? toBool(single(a, 3)) : true;
    const meta = rangeMeta.get(table);
    if (!meta) throw ERR.value();
    if (colN < 1 || colN > meta.cols) throw ERR.ref();
    const rowOf = (r: number) => table[r * meta.cols + (colN - 1)];
    let lastLE = -1;
    for (let r = 0; r < meta.rows; r++) {
      const v = table[r * meta.cols];
      if (eq(v, key)) return rowOf(r);
      if (approx && cmp(v, key) <= 0) lastLE = r;
    }
    if (approx && lastLE >= 0) return rowOf(lastLE);
    throw ERR.na();
  },
  HLOOKUP: (a) => {
    const key = single(a, 0);
    const table = a[1] ?? [];
    const rowN = Math.floor(toNum(single(a, 2)));
    const approx = a[3] ? toBool(single(a, 3)) : true;
    const meta = rangeMeta.get(table);
    if (!meta) throw ERR.value();
    if (rowN < 1 || rowN > meta.rows) throw ERR.ref();
    let lastLE = -1;
    for (let c = 0; c < meta.cols; c++) {
      const v = table[c];
      if (eq(v, key)) return table[(rowN - 1) * meta.cols + c];
      if (approx && cmp(v, key) <= 0) lastLE = c;
    }
    if (approx && lastLE >= 0) return table[(rowN - 1) * meta.cols + lastLE];
    throw ERR.na();
  },
  MATCH: (a) => {
    const key = single(a, 0);
    const arr = a[1] ?? [];
    const type = a[2] ? toNum(single(a, 2)) : 1;
    for (let i = 0; i < arr.length; i++) {
      if (type === 0 && eq(arr[i], key)) return i + 1;
    }
    if (type === 0) throw ERR.na();
    let best = -1;
    for (let i = 0; i < arr.length; i++) {
      const c = cmp(arr[i], key);
      if (type > 0 ? c <= 0 : c >= 0) best = i;
      else break;
    }
    if (best < 0) throw ERR.na();
    return best + 1;
  },
  INDEX: (a) => {
    const arr = a[0] ?? [];
    const meta = rangeMeta.get(arr) ?? { rows: arr.length, cols: 1 };
    const r = Math.floor(toNum(single(a, 1)));
    const c = a[2] ? Math.floor(toNum(single(a, 2))) : 1;
    if (meta.rows === 1 && !a[2]) {
      if (r < 1 || r > meta.cols) throw ERR.ref();
      return arr[r - 1];
    }
    if (r < 1 || r > meta.rows || c < 1 || c > meta.cols) throw ERR.ref();
    return arr[(r - 1) * meta.cols + (c - 1)];
  },
  XLOOKUP: (a) => {
    const key = single(a, 0);
    const look = a[1] ?? [];
    const ret = a[2] ?? [];
    for (let i = 0; i < look.length; i++) if (eq(look[i], key)) return ret[i] ?? '';
    if (a[3]) return single(a, 3);
    throw ERR.na();
  },
  ROW: (a, raw) => { const m = Array.isArray(raw[0]) ? refMeta.get(raw[0]) : undefined; return m ? m.row : 0; },
  COLUMN: (a, raw) => { const m = Array.isArray(raw[0]) ? refMeta.get(raw[0]) : undefined; return m ? m.col : 0; },
  ROWS: (a) => rangeMeta.get(a[0] ?? [])?.rows ?? 1,
  COLUMNS: (a) => rangeMeta.get(a[0] ?? [])?.cols ?? 1,
  // ---- date & time ----
  TODAY: () => dateToSerial(new Date()),
  NOW: () => { const d = new Date(); return dateToSerial(d) + (d.getHours() * 3600 + d.getMinutes() * 60 + d.getSeconds()) / 86400; },
  DATE: (a) => {
    const y = toNum(single(a, 0));
    const m = toNum(single(a, 1));
    const d = toNum(single(a, 2));
    return (Date.UTC(y < 1900 ? y + 1900 : y, m - 1, d) - EPOCH) / DAY_MS;
  },
  YEAR: (a) => serialToDate(serialOf(single(a, 0))).getUTCFullYear(),
  MONTH: (a) => serialToDate(serialOf(single(a, 0))).getUTCMonth() + 1,
  DAY: (a) => serialToDate(serialOf(single(a, 0))).getUTCDate(),
  WEEKDAY: (a) => {
    const d = serialToDate(serialOf(single(a, 0))).getUTCDay();
    const type = a[1] ? toNum(single(a, 1)) : 1;
    if (type === 2) return d === 0 ? 7 : d;
    if (type === 3) return d === 0 ? 6 : d - 1;
    return d + 1;
  },
  DAYS: (a) => Math.floor(serialOf(single(a, 0))) - Math.floor(serialOf(single(a, 1))),
  DATEDIF: (a) => {
    const s = serialToDate(serialOf(single(a, 0)));
    const e = serialToDate(serialOf(single(a, 1)));
    const unit = toStr(single(a, 2)).toUpperCase();
    if (e < s) throw ERR.num();
    if (unit === 'D') return Math.floor((e.getTime() - s.getTime()) / DAY_MS);
    let months = (e.getUTCFullYear() - s.getUTCFullYear()) * 12 + (e.getUTCMonth() - s.getUTCMonth());
    if (e.getUTCDate() < s.getUTCDate()) months--;
    if (unit === 'M') return months;
    if (unit === 'Y') return Math.floor(months / 12);
    throw ERR.value();
  },
  EDATE: (a) => {
    const d = serialToDate(serialOf(single(a, 0)));
    const m = toNum(single(a, 1));
    d.setUTCMonth(d.getUTCMonth() + m);
    return (d.getTime() - EPOCH) / DAY_MS;
  },
  EOMONTH: (a) => {
    const d = serialToDate(serialOf(single(a, 0)));
    const m = toNum(single(a, 1));
    const last = new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth() + m + 1, 0));
    return (last.getTime() - EPOCH) / DAY_MS;
  },
  NETWORKDAYS: (a) => {
    let s = Math.floor(serialOf(single(a, 0)));
    const e = Math.floor(serialOf(single(a, 1)));
    let n = 0;
    const sign = e >= s ? 1 : -1;
    for (; sign > 0 ? s <= e : s >= e; s += sign) {
      const wd = serialToDate(s).getUTCDay();
      if (wd !== 0 && wd !== 6) n++;
    }
    return n * sign;
  },
  HOUR: (a) => Math.floor(((toNum(single(a, 0)) % 1) * 24) % 24),
  MINUTE: (a) => Math.floor(((toNum(single(a, 0)) % 1) * 1440) % 60),
  TIME: (a) => (toNum(single(a, 0)) * 3600 + toNum(single(a, 1)) * 60 + toNum(single(a, 2))) / 86400,
  // ---- financial ----
  PMT: (a) => {
    const r = toNum(single(a, 0));
    const n = toNum(single(a, 1));
    const pv = toNum(single(a, 2));
    const fv = a[3] ? toNum(single(a, 3)) : 0;
    const type = a[4] ? toNum(single(a, 4)) : 0;
    if (r === 0) return -(pv + fv) / n;
    const f = Math.pow(1 + r, n);
    return -(r * (pv * f + fv)) / ((1 + r * type) * (f - 1));
  },
  FV: (a) => {
    const r = toNum(single(a, 0));
    const n = toNum(single(a, 1));
    const pmt = toNum(single(a, 2));
    const pv = a[3] ? toNum(single(a, 3)) : 0;
    if (r === 0) return -(pv + pmt * n);
    const f = Math.pow(1 + r, n);
    return -(pv * f + (pmt * (f - 1)) / r);
  },
  PV: (a) => {
    const r = toNum(single(a, 0));
    const n = toNum(single(a, 1));
    const pmt = toNum(single(a, 2));
    if (r === 0) return -pmt * n;
    return -(pmt * (1 - Math.pow(1 + r, -n))) / r;
  },
  NPV: (a) => {
    const r = toNum(single(a, 0));
    return numsOf(a.slice(1)).reduce((s, v, i) => s + v / Math.pow(1 + r, i + 1), 0);
  },
};

// range metadata (shape) for lookup functions
const rangeMeta = new WeakMap<Val[], { rows: number; cols: number }>();
const refMeta = new WeakMap<Val[], { row: number; col: number }>();

function eq(a: Val, b: Val): boolean {
  if (isErr(a) || isErr(b)) return false;
  if (typeof a === 'string' && typeof b === 'string') return a.toLowerCase() === b.toLowerCase();
  if (typeof a === 'number' && typeof b === 'string') return Number(b) === a && b.trim() !== '';
  if (typeof a === 'string' && typeof b === 'number') return Number(a) === b && a.trim() !== '';
  return a === b;
}

function cmp(a: Val, b: Val): number {
  if (isErr(a)) throw a;
  if (isErr(b)) throw b;
  if (typeof a === 'number' && typeof b === 'number') return a - b;
  if (typeof a === 'boolean' || typeof b === 'boolean') return Number(a) - Number(b);
  const na = typeof a === 'number' ? a : Number(a);
  const nb = typeof b === 'number' ? b : Number(b);
  if (Number.isFinite(na) && Number.isFinite(nb) && String(a).trim() !== '' && String(b).trim() !== '') return na - nb;
  return String(a).toLowerCase().localeCompare(String(b).toLowerCase());
}

/** Small subset of Excel TEXT() formats. */
function formatText(v: Val, fmt: string): string {
  if (isErr(v)) throw v;
  const f = fmt.trim();
  if (typeof v === 'number') {
    if (/^0+(\.0+)?%$/.test(f)) {
      const d = (f.split('.')[1] ?? '').length - 1;
      return `${(v * 100).toFixed(Math.max(0, d))}%`;
    }
    if (/^#,##0(\.0+)?$/.test(f) || /^0(\.0+)?$/.test(f)) {
      const d = (f.split('.')[1] ?? '').length;
      const n = v.toFixed(d);
      return f.startsWith('#') ? Number(n).toLocaleString('en-US', { minimumFractionDigits: d, maximumFractionDigits: d }) : n;
    }
    if (/^\$#,##0(\.0+)?$/.test(f)) {
      const d = (f.split('.')[1] ?? '').length;
      return `$${Number(v.toFixed(d)).toLocaleString('en-US', { minimumFractionDigits: d, maximumFractionDigits: d })}`;
    }
    if (/[ymd]/i.test(f) && !/[#0]/.test(f)) {
      const d = serialToDate(v);
      const yyyy = String(d.getUTCFullYear());
      const mm = String(d.getUTCMonth() + 1).padStart(2, '0');
      const dd = String(d.getUTCDate()).padStart(2, '0');
      const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
      return f
        .replace(/yyyy/i, yyyy)
        .replace(/yy/i, yyyy.slice(2))
        .replace(/mmmm/i, months[d.getUTCMonth()] + '')
        .replace(/mmm/i, months[d.getUTCMonth()])
        .replace(/mm/i, mm)
        .replace(/dd/i, dd)
        .replace(/\bd\b/i, String(d.getUTCDate()))
        .replace(/\bm\b/i, String(d.getUTCMonth() + 1));
    }
  }
  return toStr(v);
}

export const FUNCTION_NAMES = Object.keys(FUNCS).sort();

/** Short help used by the formula picker (name, signature, description). */
export const FUNCTION_HELP: { name: string; sig: string; desc: string; cat: string }[] = [
  { name: 'SUM', sig: 'SUM(A1:A9)', desc: 'Adds numbers', cat: 'Math' },
  { name: 'AVERAGE', sig: 'AVERAGE(A1:A9)', desc: 'Arithmetic mean', cat: 'Math' },
  { name: 'COUNT', sig: 'COUNT(A1:A9)', desc: 'Counts numeric cells', cat: 'Math' },
  { name: 'COUNTA', sig: 'COUNTA(A1:A9)', desc: 'Counts non-empty cells', cat: 'Math' },
  { name: 'MIN', sig: 'MIN(A1:A9)', desc: 'Smallest value', cat: 'Math' },
  { name: 'MAX', sig: 'MAX(A1:A9)', desc: 'Largest value', cat: 'Math' },
  { name: 'MEDIAN', sig: 'MEDIAN(A1:A9)', desc: 'Middle value', cat: 'Math' },
  { name: 'ROUND', sig: 'ROUND(A1, 2)', desc: 'Round to digits', cat: 'Math' },
  { name: 'ROUNDUP', sig: 'ROUNDUP(A1, 0)', desc: 'Round away from zero', cat: 'Math' },
  { name: 'ROUNDDOWN', sig: 'ROUNDDOWN(A1, 0)', desc: 'Round toward zero', cat: 'Math' },
  { name: 'ABS', sig: 'ABS(A1)', desc: 'Absolute value', cat: 'Math' },
  { name: 'SQRT', sig: 'SQRT(A1)', desc: 'Square root', cat: 'Math' },
  { name: 'POWER', sig: 'POWER(A1, 2)', desc: 'Raise to a power', cat: 'Math' },
  { name: 'MOD', sig: 'MOD(A1, 3)', desc: 'Remainder', cat: 'Math' },
  { name: 'INT', sig: 'INT(A1)', desc: 'Round down to integer', cat: 'Math' },
  { name: 'PRODUCT', sig: 'PRODUCT(A1:A5)', desc: 'Multiplies numbers', cat: 'Math' },
  { name: 'SUMPRODUCT', sig: 'SUMPRODUCT(A1:A5, B1:B5)', desc: 'Sum of products', cat: 'Math' },
  { name: 'STDEV', sig: 'STDEV(A1:A9)', desc: 'Sample standard deviation', cat: 'Math' },
  { name: 'LARGE', sig: 'LARGE(A1:A9, 2)', desc: 'k-th largest', cat: 'Math' },
  { name: 'SMALL', sig: 'SMALL(A1:A9, 2)', desc: 'k-th smallest', cat: 'Math' },
  { name: 'RANK', sig: 'RANK(A1, A1:A9)', desc: 'Rank in a list', cat: 'Math' },
  { name: 'SUMIF', sig: 'SUMIF(A1:A9, ">5")', desc: 'Sum if condition', cat: 'Conditional' },
  { name: 'SUMIFS', sig: 'SUMIFS(C1:C9, A1:A9, "x", B1:B9, ">1")', desc: 'Sum with several conditions', cat: 'Conditional' },
  { name: 'COUNTIF', sig: 'COUNTIF(A1:A9, "yes")', desc: 'Count if condition', cat: 'Conditional' },
  { name: 'COUNTIFS', sig: 'COUNTIFS(A1:A9, "x", B1:B9, ">1")', desc: 'Count with several conditions', cat: 'Conditional' },
  { name: 'AVERAGEIF', sig: 'AVERAGEIF(A1:A9, ">0")', desc: 'Average if condition', cat: 'Conditional' },
  { name: 'IF', sig: 'IF(A1>10, "High", "Low")', desc: 'Conditional value', cat: 'Logic' },
  { name: 'IFS', sig: 'IFS(A1>90,"A", A1>80,"B", TRUE,"C")', desc: 'Multiple conditions', cat: 'Logic' },
  { name: 'IFERROR', sig: 'IFERROR(A1/B1, 0)', desc: 'Fallback on error', cat: 'Logic' },
  { name: 'AND', sig: 'AND(A1>0, B1>0)', desc: 'All true', cat: 'Logic' },
  { name: 'OR', sig: 'OR(A1>0, B1>0)', desc: 'Any true', cat: 'Logic' },
  { name: 'NOT', sig: 'NOT(A1)', desc: 'Invert', cat: 'Logic' },
  { name: 'SWITCH', sig: 'SWITCH(A1, 1,"One", 2,"Two", "Other")', desc: 'Match against values', cat: 'Logic' },
  { name: 'ISBLANK', sig: 'ISBLANK(A1)', desc: 'Is empty', cat: 'Logic' },
  { name: 'ISNUMBER', sig: 'ISNUMBER(A1)', desc: 'Is a number', cat: 'Logic' },
  { name: 'CONCAT', sig: 'CONCAT(A1, " ", B1)', desc: 'Join text', cat: 'Text' },
  { name: 'TEXTJOIN', sig: 'TEXTJOIN(", ", TRUE, A1:A5)', desc: 'Join with delimiter', cat: 'Text' },
  { name: 'LEN', sig: 'LEN(A1)', desc: 'Text length', cat: 'Text' },
  { name: 'UPPER', sig: 'UPPER(A1)', desc: 'Uppercase', cat: 'Text' },
  { name: 'LOWER', sig: 'LOWER(A1)', desc: 'Lowercase', cat: 'Text' },
  { name: 'PROPER', sig: 'PROPER(A1)', desc: 'Capitalize words', cat: 'Text' },
  { name: 'TRIM', sig: 'TRIM(A1)', desc: 'Remove extra spaces', cat: 'Text' },
  { name: 'LEFT', sig: 'LEFT(A1, 3)', desc: 'First characters', cat: 'Text' },
  { name: 'RIGHT', sig: 'RIGHT(A1, 3)', desc: 'Last characters', cat: 'Text' },
  { name: 'MID', sig: 'MID(A1, 2, 3)', desc: 'Middle characters', cat: 'Text' },
  { name: 'FIND', sig: 'FIND("x", A1)', desc: 'Position of text (case-sensitive)', cat: 'Text' },
  { name: 'SEARCH', sig: 'SEARCH("x", A1)', desc: 'Position of text', cat: 'Text' },
  { name: 'SUBSTITUTE', sig: 'SUBSTITUTE(A1, "a", "b")', desc: 'Replace text', cat: 'Text' },
  { name: 'REPT', sig: 'REPT("*", 5)', desc: 'Repeat text', cat: 'Text' },
  { name: 'TEXT', sig: 'TEXT(A1, "0.00")', desc: 'Format number as text', cat: 'Text' },
  { name: 'VALUE', sig: 'VALUE(A1)', desc: 'Text to number', cat: 'Text' },
  { name: 'VLOOKUP', sig: 'VLOOKUP(A1, D1:F9, 2, FALSE)', desc: 'Vertical lookup', cat: 'Lookup' },
  { name: 'HLOOKUP', sig: 'HLOOKUP(A1, A1:F2, 2, FALSE)', desc: 'Horizontal lookup', cat: 'Lookup' },
  { name: 'XLOOKUP', sig: 'XLOOKUP(A1, D1:D9, E1:E9, "none")', desc: 'Lookup with return range', cat: 'Lookup' },
  { name: 'INDEX', sig: 'INDEX(A1:C9, 2, 3)', desc: 'Value at row/column', cat: 'Lookup' },
  { name: 'MATCH', sig: 'MATCH(A1, B1:B9, 0)', desc: 'Position in a range', cat: 'Lookup' },
  { name: 'CHOOSE', sig: 'CHOOSE(2, "a", "b", "c")', desc: 'Pick by index', cat: 'Lookup' },
  { name: 'ROW', sig: 'ROW(A5)', desc: 'Row number', cat: 'Lookup' },
  { name: 'COLUMN', sig: 'COLUMN(C1)', desc: 'Column number', cat: 'Lookup' },
  { name: 'TODAY', sig: 'TODAY()', desc: 'Current date', cat: 'Date' },
  { name: 'NOW', sig: 'NOW()', desc: 'Current date and time', cat: 'Date' },
  { name: 'DATE', sig: 'DATE(2026, 9, 3)', desc: 'Build a date', cat: 'Date' },
  { name: 'YEAR', sig: 'YEAR(A1)', desc: 'Year of a date', cat: 'Date' },
  { name: 'MONTH', sig: 'MONTH(A1)', desc: 'Month of a date', cat: 'Date' },
  { name: 'DAY', sig: 'DAY(A1)', desc: 'Day of a date', cat: 'Date' },
  { name: 'WEEKDAY', sig: 'WEEKDAY(A1)', desc: 'Day of week (1=Sun)', cat: 'Date' },
  { name: 'DAYS', sig: 'DAYS(B1, A1)', desc: 'Days between dates', cat: 'Date' },
  { name: 'DATEDIF', sig: 'DATEDIF(A1, B1, "M")', desc: 'Difference in D/M/Y', cat: 'Date' },
  { name: 'EDATE', sig: 'EDATE(A1, 3)', desc: 'Shift by months', cat: 'Date' },
  { name: 'EOMONTH', sig: 'EOMONTH(A1, 0)', desc: 'End of month', cat: 'Date' },
  { name: 'NETWORKDAYS', sig: 'NETWORKDAYS(A1, B1)', desc: 'Working days between', cat: 'Date' },
  { name: 'PMT', sig: 'PMT(0.05/12, 60, 10000)', desc: 'Loan payment', cat: 'Finance' },
  { name: 'FV', sig: 'FV(0.05/12, 120, -100)', desc: 'Future value', cat: 'Finance' },
  { name: 'PV', sig: 'PV(0.05/12, 60, -200)', desc: 'Present value', cat: 'Finance' },
  { name: 'NPV', sig: 'NPV(0.1, B1:B5)', desc: 'Net present value', cat: 'Finance' },
];

// ---------------------------------------------------------------------------
// parser
// ---------------------------------------------------------------------------

const REF_RE = new RegExp(`^${SHEET_PREFIX}\\$?([A-Za-z]{1,2})\\$?([0-9]{1,4})(?![A-Za-z0-9(])`);
const REF_ARG_RE = new RegExp(`^${SHEET_PREFIX}\\$?([A-Za-z]{1,2})\\$?([0-9]{1,4})(?![A-Za-z0-9(:])`);
const RANGE_RE = new RegExp(`^${SHEET_PREFIX}\\$?([A-Za-z]{1,2})\\$?([0-9]{1,4})\\s*:\\s*\\$?([A-Za-z]{1,2})\\$?([0-9]{1,4})`);

interface Ctx {
  get: CellGetter;
  /** evaluates another cell (with cycle detection) */
  cell: (ref: string) => Val;
}

class Parser {
  private i = 0;

  constructor(
    private src: string,
    private ctx: Ctx,
  ) {}

  private ws(): void {
    while (this.i < this.src.length && /\s/.test(this.src[this.i])) this.i++;
  }

  private peek(n = 1): string {
    this.ws();
    return this.src.slice(this.i, this.i + n);
  }

  parseAll(): Val {
    const v = this.parseCompare();
    this.ws();
    if (this.i < this.src.length) throw ERR.value();
    return v;
  }

  private parseCompare(): Val {
    let v = this.parseConcat();
    for (;;) {
      const two = this.peek(2);
      const one = two[0];
      let op = '';
      if (two === '<>' || two === '<=' || two === '>=') op = two;
      else if (one === '=' || one === '<' || one === '>') op = one;
      if (!op) return v;
      this.i += op.length;
      const r = this.parseConcat();
      const c = op === '=' || op === '<>' ? (eq(v, r) ? 0 : 1) : cmp(v, r);
      switch (op) {
        case '=': v = c === 0; break;
        case '<>': v = c !== 0; break;
        case '<': v = c < 0; break;
        case '>': v = c > 0; break;
        case '<=': v = c <= 0; break;
        default: v = c >= 0;
      }
    }
  }

  private parseConcat(): Val {
    let v = this.parseExpr();
    while (this.peek() === '&') {
      this.i++;
      const r = this.parseExpr();
      v = toStr(v) + toStr(r);
    }
    return v;
  }

  private parseExpr(): Val {
    let v = this.parseTerm();
    for (;;) {
      const c = this.peek();
      if (c === '+' || c === '-') {
        this.i++;
        const r = this.parseTerm();
        v = c === '+' ? toNum(v) + toNum(r) : toNum(v) - toNum(r);
      } else return v;
    }
  }

  private parseTerm(): Val {
    let v = this.parseFactor();
    for (;;) {
      const c = this.peek();
      if (c === '*' || c === '/') {
        this.i++;
        const r = this.parseFactor();
        if (c === '*') v = toNum(v) * toNum(r);
        else {
          const d = toNum(r);
          if (d === 0) throw ERR.div0();
          v = toNum(v) / d;
        }
      } else return v;
    }
  }

  private parseFactor(): Val {
    let v = this.parseUnary();
    if (this.peek() === '^') {
      this.i++;
      v = Math.pow(toNum(v), toNum(this.parseFactor()));
    }
    return v;
  }

  private parseUnary(): Val {
    const c = this.peek();
    if (c === '-') {
      this.i++;
      return -toNum(this.parseUnary());
    }
    if (c === '+') {
      this.i++;
      return this.parseUnary();
    }
    let v = this.parsePostfix();
    return v;
  }

  private parsePostfix(): Val {
    let v = this.parseBase();
    while (this.peek() === '%') {
      this.i++;
      v = toNum(v) / 100;
    }
    return v;
  }

  private parseBase(): Val {
    const c = this.peek();
    if (c === '(') {
      this.i++;
      const v = this.parseCompare();
      if (this.peek() !== ')') throw ERR.value();
      this.i++;
      return v;
    }
    if (c === '"') {
      this.i++;
      let s = '';
      for (;;) {
        if (this.i >= this.src.length) throw ERR.value();
        const ch = this.src[this.i];
        if (ch === '"') {
          if (this.src[this.i + 1] === '"') {
            s += '"';
            this.i += 2;
            continue;
          }
          this.i++;
          break;
        }
        s += ch;
        this.i++;
      }
      return s;
    }
    this.ws();
    const rest = this.src.slice(this.i);
    let m = /^([A-Za-z][A-Za-z0-9._]*)\s*\(/.exec(rest);
    if (m) {
      const name = m[1].toUpperCase();
      this.i += m[0].length;
      const args: Val[][] = [];
      const raws: RawArg[] = [];
      if (this.peek() === ')') this.i++;
      else {
        for (;;) {
          const { arr, raw } = this.parseArg(name);
          args.push(arr);
          raws.push(raw);
          if (this.peek() === ',' || this.peek() === ';') {
            this.i++;
            continue;
          }
          break;
        }
        if (this.peek() !== ')') throw ERR.value();
        this.i++;
      }
      const fn = FUNCS[name];
      if (!fn) throw ERR.name();
      return fn(args, raws);
    }
    // range as a bare value (e.g. inside SUMPRODUCT) -> handled in parseArg; here treat as error
    m = REF_RE.exec(rest);
    if (m) {
      this.i += m[0].length;
      const ref = qual(sheetOf(m[1], m[2]), m[3].toUpperCase() + m[4]);
      return this.ctx.cell(ref);
    }
    m = /^(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?/.exec(rest);
    if (m) {
      this.i += m[0].length;
      return parseFloat(m[0]);
    }
    m = /^(TRUE|FALSE)(?![A-Za-z0-9_])/i.exec(rest);
    if (m) {
      this.i += m[0].length;
      return m[1].toUpperCase() === 'TRUE';
    }
    m = /^#(DIV\/0!|VALUE!|REF!|NAME\?|N\/A|NUM!)/.exec(rest);
    if (m) {
      this.i += m[0].length;
      throw new FormulaError(`#${m[1]}`);
    }
    if (/^[A-Za-z_][A-Za-z0-9_.]*/.test(rest)) throw ERR.name();
    throw ERR.value();
  }

  /** An argument is either a range (A1:B3) or an expression. Errors are caught per-arg for IFERROR-style functions. */
  private parseArg(fnName: string): { arr: Val[]; raw: RawArg } {
    this.ws();
    const rm = RANGE_RE.exec(this.src.slice(this.i));
    if (rm) {
      this.i += rm[0].length;
      const sheet = sheetOf(rm[1], rm[2]);
      const c1 = colIdx(rm[3]);
      const c2 = colIdx(rm[5]);
      const r1 = parseInt(rm[4], 10);
      const r2 = parseInt(rm[6], 10);
      const out: Val[] = [];
      const rows = Math.abs(r2 - r1) + 1;
      const cols = Math.abs(c2 - c1) + 1;
      for (let rr = Math.min(r1, r2); rr <= Math.max(r1, r2); rr++) {
        for (let cc = Math.min(c1, c2); cc <= Math.max(c1, c2); cc++) {
          const ref = qual(sheet, colName(cc) + rr);
          try {
            out.push(this.ctx.cell(ref));
          } catch (e) {
            out.push(e instanceof FormulaError ? e : ERR.value());
          }
        }
      }
      rangeMeta.set(out, { rows, cols });
      return { arr: out, raw: out };
    }
    const lenient = fnName === 'IFERROR' || fnName === 'IFNA' || fnName === 'ISERROR' || fnName === 'ISERR';
    const start = this.i;
    const refM = REF_ARG_RE.exec(this.src.slice(this.i));
    try {
      const v = this.parseCompare();
      const arr = [v];
      if (refM && this.i === start + refM[0].length) {
        const meta = { row: parseInt(refM[4], 10), col: colIdx(refM[3]) + 1 };
        refMeta.set(arr, meta);
        rangeMeta.set(arr, { rows: 1, cols: 1 });
        return { arr, raw: arr };
      }
      return { arr, raw: v };
    } catch (e) {
      if (lenient && e instanceof FormulaError) {
        // skip the rest of this argument
        let depth = 0;
        while (this.i < this.src.length) {
          const ch = this.src[this.i];
          if (ch === '"') {
            this.i++;
            while (this.i < this.src.length && this.src[this.i] !== '"') this.i++;
          } else if (ch === '(') depth++;
          else if (ch === ')') {
            if (depth === 0) break;
            depth--;
          } else if ((ch === ',' || ch === ';') && depth === 0) break;
          this.i++;
        }
        return { arr: [e], raw: e };
      }
      throw e;
    }
  }
}

// ROW()/COLUMN() need the raw reference; they receive the args array as "raw"
// only when the argument was a bare reference (see parseArg).

// ---------------------------------------------------------------------------
// public API
// ---------------------------------------------------------------------------

export interface EvalOptions {
  /** cached evaluated values per cell — filled during evaluation */
  cache?: Map<string, Val>;
}

/** Evaluate a whole formula string ("=SUM(A1:A3)*2") to a typed value. */
export function evalValue(raw: string, get: CellGetter, opts: EvalOptions = {}, stack: Set<string> = new Set()): Val {
  if (!raw.startsWith('=')) return parseRaw(raw);
  const cache = opts.cache;
  const ctx: Ctx = {
    get,
    cell: (ref) => {
      if (cache?.has(ref)) return cache.get(ref)!;
      if (stack.has(ref)) throw ERR.circ();
      const r = get(ref);
      if (!r.startsWith('=')) return parseRaw(r);
      stack.add(ref);
      try {
        const v = evalValue(r, get, opts, stack);
        cache?.set(ref, v);
        return v;
      } finally {
        stack.delete(ref);
      }
    },
  };
  const body = raw.slice(1).trim();
  if (!body) return '';
  const p = new Parser(body, ctx);
  return p.parseAll();
}

export function valToString(v: Val): string {
  if (isErr(v)) return v.code;
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (typeof v === 'number') return fmtNum(v);
  return v;
}

/** Display string for a raw cell (formula or literal). Never throws. */
export function evalCell(raw: string, get: CellGetter, opts?: EvalOptions): string {
  if (!raw || !raw.startsWith('=')) return raw;
  try {
    return valToString(evalValue(raw, get, opts));
  } catch (e) {
    return e instanceof FormulaError ? e.code : '#VALUE!';
  }
}

/** Numeric value of a cell if it evaluates to a number (for charts/sorting). */
export function evalNumber(raw: string, get: CellGetter, opts?: EvalOptions): number | null {
  try {
    const v = evalValue(raw, get, opts);
    if (typeof v === 'number') return v;
    if (typeof v === 'string') {
      const n = Number(v.replace(/[,$%\s]/g, ''));
      return v.trim() !== '' && Number.isFinite(n) ? n : null;
    }
    if (typeof v === 'boolean') return v ? 1 : 0;
    return null;
  } catch {
    return null;
  }
}

export function evalFormula(formula: string, get: CellGetter): string {
  return evalCell(formula, get);
}

export function isFormulaError(s: string): boolean {
  return /^#(DIV\/0!|VALUE!|REF!|NAME\?|N\/A|NUM!|CIRC!|ERR)$/.test(s);
}
