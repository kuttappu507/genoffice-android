export type CellGetter = (ref: string) => string;

function colIdx(col: string): number {
  let n = 0;
  for (const c of col.toUpperCase()) n = n * 26 + (c.charCodeAt(0) - 64);
  return n - 1;
}

function colName(i: number): string {
  return String.fromCharCode(65 + i);
}

function colOk(col: string): boolean {
  const u = col.toUpperCase();
  return u.length === 1 && u >= 'A' && u <= 'Z';
}

function refKey(col: string, row: string): string {
  return `${col.toUpperCase()}${row}`;
}

function applyFunc(name: string, a: number[]): number {
  switch (name) {
    case 'SUM':
      return a.reduce((x, y) => x + y, 0);
    case 'AVERAGE':
    case 'AVG':
      return a.length ? a.reduce((x, y) => x + y, 0) / a.length : 0;
    case 'COUNT':
      return a.length;
    case 'MIN':
      return a.length ? Math.min(...a) : 0;
    case 'MAX':
      return a.length ? Math.max(...a) : 0;
    case 'ABS':
      return Math.abs(a[0] ?? 0);
    case 'ROUND': {
      const d = a[1] ?? 0;
      const f = 10 ** d;
      return Math.round((a[0] ?? 0) * f) / f;
    }
    default:
      throw new Error('unknown function');
  }
}

function fmtNum(v: number): string {
  if (!Number.isFinite(v)) return '#ERR';
  if (Number.isInteger(v)) return String(v);
  return String(parseFloat(v.toFixed(6)));
}

/**
 * Small recursive-descent evaluator for spreadsheet formulas.
 * Supports: numbers, cell refs (A1), ranges inside functions (A1:B3),
 * functions SUM/AVG/AVERAGE/COUNT/MIN/MAX/ABS/ROUND, operators + - * / ^ and
 * parentheses. Enough for everyday budget/list math without any dependency.
 */
class Parser {
  private i = 0;

  constructor(
    private src: string,
    private get: CellGetter,
  ) {}

  private ws(): void {
    while (this.i < this.src.length && this.src[this.i] === ' ') this.i++;
  }

  private peek(): string {
    this.ws();
    return this.src[this.i] ?? '';
  }

  parseExpr(): number {
    let v = this.parseTerm();
    for (;;) {
      const c = this.peek();
      if (c === '+' || c === '-') {
        this.i++;
        const r = this.parseTerm();
        v = c === '+' ? v + r : v - r;
      } else {
        return v;
      }
    }
  }

  private parseTerm(): number {
    let v = this.parseFactor();
    for (;;) {
      const c = this.peek();
      if (c === '*' || c === '/') {
        this.i++;
        const r = this.parseFactor();
        v = c === '*' ? v * r : v / r;
      } else {
        return v;
      }
    }
  }

  private parseFactor(): number {
    let v = this.parseBase();
    if (this.peek() === '^') {
      this.i++;
      v = Math.pow(v, this.parseFactor());
    }
    return v;
  }

  private parseBase(): number {
    const c = this.peek();
    if (c === '(') {
      this.i++;
      const v = this.parseExpr();
      if (this.peek() !== ')') throw new Error('expected )');
      this.i++;
      return v;
    }
    if (c === '-') {
      this.i++;
      return -this.parseBase();
    }
    this.ws();
    const rest = this.src.slice(this.i);
    let m = /^([A-Za-z]{1,8})\s*\(/.exec(rest);
    if (m) {
      const name = m[1].toUpperCase();
      this.i += m[0].length;
      const args: number[] = [];
      if (this.peek() === ')') {
        this.i++;
      } else {
        for (;;) {
          args.push(...this.parseArg());
          if (this.peek() === ',') {
            this.i++;
            continue;
          }
          break;
        }
        if (this.peek() !== ')') throw new Error('expected )');
        this.i++;
      }
      return applyFunc(name, args);
    }
    m = /^([A-Za-z])([0-9]{1,3})/.exec(rest);
    if (m && colOk(m[1])) {
      this.i += m[0].length;
      const n = parseFloat(this.get(refKey(m[1], m[2])));
      if (!Number.isFinite(n)) throw new Error('bad cell');
      return n;
    }
    m = /^\d+(\.\d+)?/.exec(rest);
    if (m) {
      this.i += m[0].length;
      return parseFloat(m[0]);
    }
    throw new Error('unexpected token');
  }

  private parseArg(): number[] {
    const save = this.i;
    const rm = /^([A-Za-z])([0-9]{1,3})\s*:\s*([A-Za-z])([0-9]{1,3})/.exec(this.src.slice(this.i));
    if (rm && colOk(rm[1]) && colOk(rm[3])) {
      this.i += rm[0].length;
      const c1 = colIdx(rm[1]);
      const c2 = colIdx(rm[3]);
      const r1 = parseInt(rm[2], 10);
      const r2 = parseInt(rm[4], 10);
      const out: number[] = [];
      for (let cc = Math.min(c1, c2); cc <= Math.max(c1, c2); cc++) {
        for (let rr = Math.min(r1, r2); rr <= Math.max(r1, r2); rr++) {
          const n = parseFloat(this.get(colName(cc) + rr));
          if (Number.isFinite(n)) out.push(n);
        }
      }
      return out;
    }
    this.i = save;
    return [this.parseExpr()];
  }
}

export function evalFormula(formula: string, get: CellGetter): string {
  const p = new Parser(formula.slice(1), get);
  const v = p.parseExpr();
  return fmtNum(v);
}

export function evalCell(raw: string, get: CellGetter): string {
  if (!raw || !raw.startsWith('=')) return raw;
  try {
    return evalFormula(raw, get);
  } catch {
    return '#ERR';
  }
}
