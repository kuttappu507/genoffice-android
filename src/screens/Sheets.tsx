import { useMemo, useRef, useState } from 'react';
import { FileTypeIcon } from '../components/Icon';
import { Palette, RBtn, RGroup, RSelect, RWide } from '../components/Ribbon';
import { evalCell } from '../lib/formulas';
import { chatStream, errMsg } from '../lib/ai-client';
import { exportXlsx, importSpreadsheet, openFilePicker, saveBinary, sanitizeName } from '../lib/fileio';
import { debounce, downloadText, getDoc, getSettings, putDoc, uid } from '../lib/storage';

const COLS = 26;
const ROWS = 60;

interface CellStyle {
  b?: boolean;
  i?: boolean;
  u?: boolean;
  color?: string;
  fill?: string;
  align?: 'left' | 'center' | 'right';
  fmt?: 'gen' | 'num' | 'cur' | 'pct';
}

interface Page {
  cells: Record<string, string>;
  styles: Record<string, CellStyle>;
}

interface Book {
  order: string[];
  active: string;
  sheets: Record<string, Page>;
}

type RibbonTab = 'home' | 'insert' | 'data' | 'ai';

const FMTS = [
  { v: 'gen', t: 'General' },
  { v: 'num', t: 'Number' },
  { v: 'cur', t: 'Currency' },
  { v: 'pct', t: 'Percent' },
];

function colName(i: number): string {
  return String.fromCharCode(65 + i);
}

function parseRef(ref: string): [number, number] {
  const m = /^([A-Z])(\d{1,3})$/.exec(ref);
  return m ? [m[1].charCodeAt(0) - 65, parseInt(m[2], 10) - 1] : [0, 0];
}

function refOf(c: number, r: number): string {
  return `${colName(c)}${r + 1}`;
}

function clone<T>(v: T): T {
  return JSON.parse(JSON.stringify(v)) as T;
}

function freshBook(): Book {
  return { order: ['Sheet1'], active: 'Sheet1', sheets: { Sheet1: { cells: {}, styles: {} } } };
}

function loadBook(id: string | undefined): Book {
  if (!id) return freshBook();
  const d = getDoc<Partial<Book> & { cells?: Record<string, string> }>(id);
  if (!d) return freshBook();
  if (d.sheets && d.order && d.active) return clone(d as Book);
  // migrate the old single-sheet format
  return { order: ['Sheet1'], active: 'Sheet1', sheets: { Sheet1: { cells: d.cells ?? {}, styles: {} } } };
}

/** Re-map A1-style refs inside a formula string. fn returns the new [col,row] (0-based) or null to keep. */
function remapFormula(f: string, fn: (c: number, r: number) => [number, number] | null): string {
  return f.replace(/([A-Z])(\d{1,3})/g, (m, c: string, r: string) => {
    // refs inside formulas are 1-based ("B3") — convert to the 0-based map domain
    const t = fn(c.charCodeAt(0) - 65, parseInt(r, 10) - 1);
    return t ? refOf(t[0], t[1]) : m;
  });
}

/** Move a whole page (cells + styles) with a mapping; null drops the cell. */
function movePage(page: Page, map: (c: number, r: number) => [number, number] | null): Page {
  const cells: Record<string, string> = {};
  const styles: Record<string, CellStyle> = {};
  for (const [k, v] of Object.entries(page.cells)) {
    const [c, r] = parseRef(k);
    const t = map(c, r);
    if (!t) continue;
    let val = v;
    if (val.startsWith('=')) val = remapFormula(val, map);
    cells[refOf(t[0], t[1])] = val;
  }
  for (const [k, v] of Object.entries(page.styles)) {
    const [c, r] = parseRef(k);
    const t = map(c, r);
    if (t) styles[refOf(t[0], t[1])] = v;
  }
  return { cells, styles };
}

function fmtDisplay(val: string, st?: CellStyle): string {
  if (!st?.fmt || st.fmt === 'gen' || !val.trim()) return val;
  const n = parseFloat(val.replace(/[$,]/g, ''));
  if (!Number.isFinite(n)) return val;
  if (st.fmt === 'num') return n.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  if (st.fmt === 'cur') {
    const abs = Math.abs(n).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    return n < 0 ? `-$${abs}` : `$${abs}`;
  }
  return `${(n * 100).toFixed(1)}%`;
}

export default function Sheets({ initialId, onExit }: { initialId?: string; onExit?: () => void }) {
  const sheetId = useRef(initialId ?? uid()).current;
  const [title, setTitle] = useState(initialId ? 'Spreadsheet' : 'Untitled sheet');
  const [book, setBook] = useState<Book>(() => loadBook(initialId));
  const [, bumpUndo] = useState(0);
  const [sel, setSel] = useState('A1');
  const [rangeEnd, setRangeEnd] = useState<string | null>(null);
  const [rangeMode, setRangeMode] = useState(false);
  const [editing, setEditing] = useState<string | null>(null);
  const [rTab, setRTab] = useState<RibbonTab>('home');
  const [palette, setPalette] = useState<'text' | 'fill' | null>(null);
  const [menu, setMenu] = useState(false);
  const [clip, setClip] = useState<{ cells: Record<string, string>; styles: Record<string, CellStyle>; w: number; h: number } | null>(null);
  const [aiOpen, setAiOpen] = useState(false);
  const [aiPrompt, setAiPrompt] = useState('');
  const [aiBusy, setAiBusy] = useState(false);
  const [toast, setToast] = useState('');

  const past = useRef<Book[]>([]);
  const future = useRef<Book[]>([]);
  const press = useRef<{ timer: ReturnType<typeof setTimeout> | null; ref: string } | null>(null);

  const page = book.sheets[book.active] ?? { cells: {}, styles: {} };
  const cells = page.cells;
  const styles = page.styles;

  const save = useMemo(
    () =>
      debounce((b: Book, t: string) => {
        putDoc<Book>('sheet', sheetId, t, b);
      }, 700),
    [sheetId],
  );

  const pushU = (b: Book) => {
    past.current.push(clone(b));
    if (past.current.length > 40) past.current.shift();
    future.current = [];
    bumpUndo((v) => v + 1);
  };

  const mutate = (fn: (b: Book) => Book) => {
    pushU(book);
    const nb = fn(book);
    setBook(nb);
    save(nb, title);
    bumpUndo((v) => v + 1);
  };

  const undo = () => {
    const p = past.current.pop();
    if (!p) return;
    future.current.push(clone(book));
    setBook(p);
    save(p, title);
    bumpUndo((v) => v + 1);
  };

  const redo = () => {
    const f = future.current.pop();
    if (!f) return;
    past.current.push(clone(book));
    setBook(f);
    save(f, title);
    bumpUndo((v) => v + 1);
  };

  const display = (ref: string): string => evalCell(cells[ref] ?? '', (r) => cells[r] ?? '');

  // ------------------------------------------------------------------ selection
  const rg = useMemo(() => {
    const [c1a, r1a] = parseRef(sel);
    const [c2a, r2a] = parseRef(rangeEnd ?? sel);
    return {
      c1: Math.min(c1a, c2a), r1: Math.min(r1a, r2a),
      c2: Math.max(c1a, c2a), r2: Math.max(r1a, r2a),
    };
  }, [sel, rangeEnd]);

  const rangeRefs = (): string[] => {
    const out: string[] = [];
    for (let c = rg.c1; c <= rg.c2; c++) for (let r = rg.r1; r <= rg.r2; r++) out.push(refOf(c, r));
    return out;
  };

  const cellDown = (ref: string) => {
    press.current = {
      ref,
      timer: setTimeout(() => {
        press.current = null;
        if (!rangeMode) {
          setRangeMode(true);
          setRangeEnd(null);
          flash('Range select: now tap the last cell');
          try { navigator.vibrate?.(25); } catch { /* noop */ }
        }
      }, 550),
    };
  };

  const cellUp = (ref: string, shift: boolean) => {
    const p = press.current;
    if (p?.timer) clearTimeout(p.timer);
    press.current = null;
    if (!p) return; // long press already handled
    if (shift) {
      setRangeEnd(ref);
      return;
    }
    if (rangeMode) {
      if (!rangeEnd) {
        if (ref !== sel) setRangeEnd(ref);
      } else {
        setRangeMode(false);
        setRangeEnd(null);
        setSel(ref);
        setEditing(cells[ref] ?? '');
      }
      return;
    }
    setSel(ref);
    setRangeEnd(null);
    setEditing(cells[ref] ?? '');
  };

  // ------------------------------------------------------------------ edits
  const commit = (ref: string, raw: string) => {
    if ((cells[ref] ?? '') === raw) return;
    mutate((b) => {
      const pg = { ...b.sheets[b.active] };
      pg.cells = { ...pg.cells };
      if (raw === '') delete pg.cells[ref];
      else pg.cells[ref] = raw;
      return { ...b, sheets: { ...b.sheets, [b.active]: pg } };
    });
  };

  const commitAndMove = (dc: number, dr: number, live?: string) => {
    if (live !== undefined) commit(sel, live);
    else if (editing !== null) commit(sel, editing);
    const [c, r] = parseRef(sel);
    const nc = c + dc;
    const nr = r + dr;
    if (nc >= 0 && nc < COLS && nr >= 0 && nr < ROWS) {
      const next = refOf(nc, nr);
      setSel(next);
      setEditing(cells[next] ?? '');
    }
  };

  const setEditingRaw = (v: string | null) => setEditing(v);

  // ------------------------------------------------------------------ styling
  const applyStyle = (patch: CellStyle) => {
    mutate((b) => {
      const pg = { ...b.sheets[b.active] };
      pg.styles = { ...pg.styles };
      for (const ref of rangeRefs()) {
        const next: CellStyle = { ...pg.styles[ref], ...patch };
        for (const k of Object.keys(next) as (keyof CellStyle)[]) {
          if (next[k] === undefined || next[k] === null) delete next[k];
        }
        if (Object.keys(next).length === 0) delete pg.styles[ref];
        else pg.styles[ref] = next;
      }
      return { ...b, sheets: { ...b.sheets, [b.active]: pg } };
    });
  };

  const toggleStyle = (key: 'b' | 'i' | 'u') => {
    const cur = styles[sel]?.[key];
    applyStyle({ [key]: !cur } as CellStyle);
  };

  const clearContents = () => {
    mutate((b) => {
      const pg = { ...b.sheets[b.active] };
      pg.cells = { ...pg.cells };
      for (const ref of rangeRefs()) delete pg.cells[ref];
      return { ...b, sheets: { ...b.sheets, [b.active]: pg } };
    });
    flash('Contents cleared.');
  };

  const clearFormatting = () => {
    mutate((b) => {
      const pg = { ...b.sheets[b.active] };
      pg.styles = { ...pg.styles };
      for (const ref of rangeRefs()) delete pg.styles[ref];
      return { ...b, sheets: { ...b.sheets, [b.active]: pg } };
    });
    flash('Formatting cleared.');
  };

  // ------------------------------------------------------------------ clipboard
  const copySel = () => {
    const c: Record<string, string> = {};
    const s: Record<string, CellStyle> = {};
    for (const ref of rangeRefs()) {
      if (cells[ref] !== undefined) c[ref] = cells[ref];
      if (styles[ref]) s[ref] = { ...styles[ref] };
    }
    setClip({ cells: c, styles: s, w: rg.c2 - rg.c1 + 1, h: rg.r2 - rg.r1 + 1 });
    flash(`Copied ${Object.keys(c).length} cell(s).`);
  };

  const pasteClip = () => {
    if (!clip) {
      flash('Nothing to paste. Copy a cell or range first.');
      return;
    }
    const [c0, r0] = parseRef(sel);
    mutate((b) => {
      const pg = { ...b.sheets[b.active] };
      pg.cells = { ...pg.cells };
      pg.styles = { ...pg.styles };
      for (const [k, v] of Object.entries(clip.cells)) {
        const [c, r] = parseRef(k);
        pg.cells[refOf(c0 + c, r0 + r)] = v;
      }
      for (const [k, v] of Object.entries(clip.styles)) {
        const [c, r] = parseRef(k);
        pg.styles[refOf(c0 + c, r0 + r)] = { ...v };
      }
      return { ...b, sheets: { ...b.sheets, [b.active]: pg } };
    });
  };

  // ------------------------------------------------------------------ structure ops
  const insertRow = (at: number) => {
    mutate((b) => {
      const pg = movePage(b.sheets[b.active], (c, r) => (r >= at ? (r + 1 < ROWS ? [c, r + 1] : null) : [c, r]));
      return { ...b, sheets: { ...b.sheets, [b.active]: pg } };
    });
  };

  const deleteRow = (at: number) => {
    mutate((b) => {
      const pg = movePage(b.sheets[b.active], (c, r) => (r === at ? null : r > at ? [c, r - 1] : [c, r]));
      return { ...b, sheets: { ...b.sheets, [b.active]: pg } };
    });
  };

  const insertCol = (at: number) => {
    mutate((b) => {
      const pg = movePage(b.sheets[b.active], (c, r) => (c >= at ? (c + 1 < COLS ? [c + 1, r] : null) : [c, r]));
      return { ...b, sheets: { ...b.sheets, [b.active]: pg } };
    });
  };

  const deleteCol = (at: number) => {
    mutate((b) => {
      const pg = movePage(b.sheets[b.active], (c, r) => (c === at ? null : c > at ? [c - 1, r] : [c, r]));
      return { ...b, sheets: { ...b.sheets, [b.active]: pg } };
    });
  };

  const [sc, sr] = parseRef(sel);

  const sortRange = (dir: 1 | -1) => {
    let maxR = 0;
    for (const k of Object.keys(cells)) {
      const r = parseInt(k.slice(1), 10);
      if (Number.isFinite(r)) maxR = Math.max(maxR, r);
    }
    if (maxR < 1) {
      flash('Nothing to sort yet.');
      return;
    }
    const startRow = sr === 0 && maxR >= 1 ? 1 : 0; // row 1 acts as header when sorting from it
    const col = sc;
    const rows: { r: number; raw: string; disp: string }[] = [];
    for (let r = startRow; r <= maxR - 1; r++) {
      const ref = refOf(col, r);
      const raw = cells[ref];
      if (raw === undefined) continue;
      rows.push({ r, raw, disp: evalCell(raw, (x) => cells[x] ?? '') });
    }
    if (rows.length < 2) {
      flash('Need at least two values in this column to sort.');
      return;
    }
    const num = (d: string) => {
      const n = parseFloat(d.replace(/[$,]/g, ''));
      return Number.isFinite(n) && d.trim() !== '' ? n : null;
    };
    rows.sort((a, b) => {
      const na = num(a.disp);
      const nb = num(b.disp);
      if (na !== null && nb !== null) return (na - nb) * dir;
      if (na !== null) return -1;
      if (nb !== null) return 1;
      return a.disp.localeCompare(b.disp) * dir;
    });
    mutate((b) => {
      const pg = { ...b.sheets[b.active] };
      pg.cells = { ...pg.cells };
      for (let r = startRow; r <= maxR - 1; r++) delete pg.cells[refOf(col, r)];
      rows.forEach((row, i) => {
        pg.cells[refOf(col, startRow + i)] = row.raw;
      });
      return { ...b, sheets: { ...b.sheets, [b.active]: pg } };
    });
    flash(`Sorted ${rows.length} rows ${dir === 1 ? 'A to Z' : 'Z to A'}.`);
  };

  // ------------------------------------------------------------------ formulas
  const autoSumRange = (): string => {
    if (rangeEnd) return `${sel}:${rangeEnd}`;
    const [c, r0] = parseRef(sel);
    let top = r0;
    while (top - 1 >= 0 && cells[refOf(c, top - 1)] !== undefined) top--;
    if (top === r0) return refOf(c, Math.max(0, r0 - 1));
    return `${refOf(c, top)}:${refOf(c, r0 - 1 > top - 1 ? r0 - 1 : top)}`;
  };

  const insertFormula = (name: string) => {
    setEditingRaw(`=${name}(${autoSumRange()})`);
  };

  const autoSum = () => {
    setEditingRaw(`=SUM(${autoSumRange()})`);
  };

  // ------------------------------------------------------------------ sheet ops
  const uniqueName = (base: string): string => {
    if (!book.order.includes(base)) return base;
    let i = 2;
    while (book.order.includes(`${base}${i}`)) i++;
    return `${base}${i}`;
  };

  const addSheet = () => {
    const name = uniqueName(`Sheet${book.order.length + 1}`);
    mutate((b) => ({
      ...b,
      order: [...b.order, name],
      active: name,
      sheets: { ...b.sheets, [name]: { cells: {}, styles: {} } },
    }));
    setSel('A1');
    setEditing('');
    setRangeEnd(null);
    setRangeMode(false);
  };

  const switchSheet = (name: string) => {
    if (name === book.active) return;
    pushU(book);
    const nb = { ...book, active: name };
    setBook(nb);
    save(nb, title);
    bumpUndo((v) => v + 1);
    setSel('A1');
    setEditing(nb.sheets[name].cells['A1'] ?? '');
    setRangeEnd(null);
    setRangeMode(false);
  };

  const renameSheet = () => {
    setMenu(false);
    const name = window.prompt('Rename sheet', book.active);
    if (!name || name === book.active) return;
    const clean = name.replace(/[\\/?*[\]:]/g, ' ').slice(0, 31).trim();
    if (!clean || book.order.includes(clean)) {
      flash('That name is empty or already in use.');
      return;
    }
    mutate((b) => {
      const sheets = { ...b.sheets };
      sheets[clean] = sheets[b.active];
      delete sheets[b.active];
      return { order: b.order.map((n) => (n === b.active ? clean : n)), active: clean, sheets };
    });
  };

  const duplicateSheet = () => {
    setMenu(false);
    const name = uniqueName(`${book.active} copy`);
    mutate((b) => ({
      ...b,
      order: [...b.order, name],
      active: name,
      sheets: { ...b.sheets, [name]: clone(b.sheets[b.active]) },
    }));
  };

  const deleteSheet = () => {
    setMenu(false);
    if (book.order.length <= 1) {
      flash('A workbook needs at least one sheet.');
      return;
    }
    if (!window.confirm(`Delete sheet "${book.active}"?`)) return;
    mutate((b) => {
      const order = b.order.filter((n) => n !== b.active);
      const sheets = { ...b.sheets };
      delete sheets[b.active];
      return { order, active: order[0], sheets };
    });
    setSel('A1');
    setEditing('');
  };

  // ------------------------------------------------------------------ AI
  const runAi = async () => {
    const s = getSettings();
    if (!s.apiKey) {
      flash('Add your API key in Settings first.');
      setAiOpen(false);
      return;
    }
    const topic = aiPrompt.trim();
    if (!topic) return;
    setAiBusy(true);
    try {
      const out = await chatStream(s, [
        {
          role: 'system',
          content:
            'You generate spreadsheet data. Return ONLY a tab-separated table. First row is a header row. No markdown fences, no commentary, max 8 columns and 20 rows. Same language as the request.',
        },
        { role: 'user', content: topic },
      ]);
      const clean = out.replace(/```[a-z]*\n?/gi, '').trim();
      const lines = clean.split('\n').filter((l) => l.trim());
      mutate((b) => {
        const pg = { cells: {} as Record<string, string>, styles: {} as Record<string, CellStyle> };
        let r = 0;
        for (const line of lines.slice(0, ROWS)) {
          const cols = line.split('\t');
          for (let c = 0; c < Math.min(cols.length, COLS); c++) {
            const v = cols[c].trim();
            if (v) pg.cells[refOf(c, r)] = v;
          }
          r++;
        }
        return { ...b, sheets: { ...b.sheets, [b.active]: pg } };
      });
      setAiOpen(false);
      flash(`Imported ${Math.min(lines.length, ROWS)} rows starting at A1.`);
    } catch (e) {
      flash(`Error: ${errMsg(e)}`);
    } finally {
      setAiBusy(false);
    }
  };

  // ------------------------------------------------------------------ files
  const openFile = async () => {
    setMenu(false);
    const pick = await openFilePicker('.xlsx,.xls,.csv');
    if (!pick) return;
    try {
      const { cells: loaded } = await importSpreadsheet(pick.buf);
      pushU(book);
      const nb = freshBook();
      nb.sheets.Sheet1.cells = loaded;
      setBook(nb);
      save(nb, pick.name.replace(/\.[^.]+$/, ''));
      setTitle(pick.name.replace(/\.[^.]+$/, ''));
      setSel('A1');
      setEditing(loaded['A1'] ?? '');
      setRangeEnd(null);
      setRangeMode(false);
      flash(`Opened ${pick.name} (${Object.keys(loaded).length} cells)`);
    } catch (e) {
      flash(`Could not open: ${errMsg(e)}`);
    }
  };

  const saveXlsx = async () => {
    try {
      const bytes = await exportXlsx(title, cells);
      await saveBinary(sanitizeName(title, 'xlsx'), bytes, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      flash('Workbook saved (active sheet)');
    } catch (e) {
      flash(`Save failed: ${errMsg(e)}`);
    }
  };

  const exportCsv = () => {
    setMenu(false);
    let maxRow = 0;
    let maxCol = 0;
    for (const k of Object.keys(cells)) {
      const r = parseInt(k.slice(1), 10);
      const c = k.charCodeAt(0) - 65;
      if (Number.isFinite(r) && cells[k] !== '') {
        maxRow = Math.max(maxRow, r);
        maxCol = Math.max(maxCol, c);
      }
    }
    const rows: string[] = [];
    for (let r = 1; r <= maxRow; r++) {
      const cols: string[] = [];
      for (let c = 0; c <= maxCol; c++) {
        const ref = `${colName(c)}${r}`;
        const v = fmtDisplay(display(ref), styles[ref]);
        cols.push(/[",\n]/.test(v) ? `"${v.replace(/"/g, '""')}"` : v);
      }
      rows.push(cols.join(','));
    }
    void downloadText(`${title.replace(/[^\w-]+/g, '_') || 'sheet'}.csv`, rows.join('\n'), 'text/csv').then(flash);
  };

  const goRef = () => {
    const r = window.prompt('Go to cell', sel);
    if (!r) return;
    const m = /^([A-Za-z])([0-9]{1,3})$/.exec(r.trim());
    if (!m) {
      flash('Use a cell reference like B4.');
      return;
    }
    const ref = `${m[1].toUpperCase()}${m[2]}`;
    setSel(ref);
    setRangeEnd(null);
    setRangeMode(false);
    setEditing(cells[ref] ?? '');
  };

  const flash = (msg: string) => {
    setToast(msg);
    setTimeout(() => setToast(''), 3000);
  };

  const curStyle = styles[sel];
  const rows = [];
  for (let r = 0; r < ROWS; r++) {
    const cols = [];
    for (let c = 0; c < COLS; c++) {
      const ref = refOf(c, r);
      const st = styles[ref];
      const inRange = c >= rg.c1 && c <= rg.c2 && r >= rg.r1 && r <= rg.r2;
      const isAnchor = ref === sel;
      cols.push(
        <button
          key={ref}
          data-ref={ref}
          className={`cell${inRange ? (isAnchor ? ' selected' : ' range') : ''}`}
          style={{
            fontWeight: st?.b ? 700 : undefined,
            fontStyle: st?.i ? 'italic' : undefined,
            textDecoration: st?.u ? 'underline' : undefined,
            color: st?.color,
            background: st?.fill,
            textAlign: st?.align,
          }}
          onPointerDown={() => cellDown(ref)}
          onPointerUp={(e) => cellUp(ref, e.shiftKey)}
        >
          {fmtDisplay(display(ref), st)}
        </button>,
      );
    }
    rows.push(
      <div className="sheet-row" key={r}>
        <div className={`cell head rowhead${r === sr ? ' hitsel' : ''}`}>{r + 1}</div>
        {cols}
      </div>,
    );
  }

  const hasUndo = past.current.length > 0;
  const hasRedo = future.current.length > 0;

  return (
    <div className="edscreen" style={{ ['--app' as string]: 'var(--excel)' }}>
      <header className="appbar">
        <button className="icon-btn light" aria-label="Back to Home" onClick={onExit}>
          <FileTypeIcon kind="sheet" size={22} />
        </button>
        <input
          className="appbar-title"
          value={title}
          onChange={(e) => {
            setTitle(e.target.value);
            save(book, e.target.value);
          }}
          placeholder="Sheet title"
        />
        <button className="icon-btn light" aria-label="Save as .xlsx" onClick={() => void saveXlsx()}>
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z" /><polyline points="17 21 17 13 7 13 7 21" /><polyline points="7 3 7 8 15 8" /></svg>
        </button>
        <div className="menu-wrap">
          <button className="icon-btn light" aria-label="More actions" onClick={() => setMenu(!menu)}>
            <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" aria-hidden="true"><circle cx="12" cy="5" r="1.7" fill="currentColor" /><circle cx="12" cy="12" r="1.7" fill="currentColor" /><circle cx="12" cy="19" r="1.7" fill="currentColor" /></svg>
          </button>
          {menu && (
            <>
              <div className="menu-backdrop" onClick={() => setMenu(false)} />
              <div className="menu">
                <button className="menu-item" onClick={() => void openFile()}>
                  Open .xlsx / .csv
                </button>
                <button className="menu-item" onClick={() => { setMenu(false); setAiOpen(true); }}>
                  AI Fill (new table)
                </button>
                <button className="menu-item" onClick={exportCsv}>
                  Export CSV
                </button>
              </div>
            </>
          )}
        </div>
      </header>

      <div className="formula-bar">
        <button className="namebox" onClick={goRef} title="Go to cell">
          {rangeEnd ? `${sel}:${rangeEnd}` : sel}
        </button>
        <span className="fx">fx</span>
        <input
          key={sel + (rangeEnd ?? '') + book.active}
          className="fx-input"
          value={editing ?? ''}
          placeholder="Value or =A1+B2, =SUM(A1:A9)"
          onChange={(e) => setEditingRaw(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === 'Enter') commitAndMove(0, 1, e.currentTarget.value);
            if (e.key === 'Tab') {
              e.preventDefault();
              commitAndMove(1, 0, e.currentTarget.value);
            }
          }}
          onBlur={() => {
            if (editing !== null) commit(sel, editing);
          }}
        />
      </div>

      <div className="sheet-grid">
        <div className="sheet-row">
          <div className="cell head corner" />
          {Array.from({ length: COLS }, (_, c) => (
            <div key={c} className={`cell head colhead${c === sc ? ' hitsel' : ''}`}>
              {colName(c)}
            </div>
          ))}
        </div>
        {rows}
      </div>

      <div className="sheet-tabs">
        <div className="stab-scroll">
          {book.order.map((name) => (
            <button key={name} className={`stab${name === book.active ? ' active' : ''}`} onClick={() => switchSheet(name)}>
              {name}
            </button>
          ))}
        </div>
        <button className="stab-add" aria-label="Add sheet" onClick={addSheet}>
          +
        </button>
        <div className="menu-wrap">
          <button className="stab-add" aria-label="Sheet options" onClick={() => setMenu(!menu)}>
            ⋯
          </button>
          {menu && (
            <>
              <div className="menu-backdrop" onClick={() => setMenu(false)} />
              <div className="menu up">
                <button className="menu-item" onClick={renameSheet}>
                  Rename sheet
                </button>
                <button className="menu-item" onClick={duplicateSheet}>
                  Duplicate sheet
                </button>
                <button className="menu-item" onClick={deleteSheet}>
                  Delete sheet
                </button>
              </div>
            </>
          )}
        </div>
      </div>

      <div className="ribbon">
        <div className="ribbon-tabs">
          {(['home', 'insert', 'data', 'ai'] as RibbonTab[]).map((t) => (
            <button key={t} className={`ribbon-tab${rTab === t ? ' active' : ''}`} onClick={() => { setRTab(t); setPalette(null); }}>
              {t === 'home' ? 'Home' : t === 'insert' ? 'Insert' : t === 'data' ? 'Data' : 'AI'}
            </button>
          ))}
        </div>

        {palette && (
          <Palette
            onPick={(c) => {
              applyStyle(palette === 'text' ? { color: c } : { fill: c });
              setPalette(null);
            }}
            auto={() => {
              applyStyle(palette === 'text' ? { color: undefined } : { fill: undefined });
              setPalette(null);
            }}
          />
        )}

        {rTab === 'home' && (
          <div className="ribbon-row">
            <RGroup label="Edit">
              <RBtn icon="undo" label="Undo" disabled={!hasUndo} onRun={undo} />
              <RBtn icon="redo" label="Redo" disabled={!hasRedo} onRun={redo} />
              <RBtn icon="copy" label="Copy" onRun={copySel} />
              <RBtn icon="paste" label="Paste" onRun={pasteClip} />
            </RGroup>
            <RGroup label="Font">
              <RBtn icon="bold" label="Bold" active={!!curStyle?.b} onRun={() => toggleStyle('b')} />
              <RBtn icon="italic" label="Italic" active={!!curStyle?.i} onRun={() => toggleStyle('i')} />
              <RBtn icon="underline" label="Underline" active={!!curStyle?.u} onRun={() => toggleStyle('u')} />
              <RBtn icon="fontColor" label="Color" colorBar={curStyle?.color ?? '#C00000'} onRun={() => setPalette(palette === 'text' ? null : 'text')} />
              <RBtn icon="fill" label="Fill" colorBar={curStyle?.fill ?? '#FFFF00'} onRun={() => setPalette(palette === 'fill' ? null : 'fill')} />
            </RGroup>
            <RGroup label="Alignment">
              <RBtn icon="alignLeft" label="Left" active={curStyle?.align === 'left' || !curStyle?.align} onRun={() => applyStyle({ align: 'left' })} />
              <RBtn icon="alignCenter" label="Center" active={curStyle?.align === 'center'} onRun={() => applyStyle({ align: 'center' })} />
              <RBtn icon="alignRight" label="Right" active={curStyle?.align === 'right'} onRun={() => applyStyle({ align: 'right' })} />
            </RGroup>
            <RGroup label="Number">
              <RSelect
                value={curStyle?.fmt ?? 'gen'}
                options={FMTS}
                onChange={(v) => applyStyle({ fmt: v === 'gen' ? undefined : (v as CellStyle['fmt']) })}
                width={96}
                title="Number format"
              />
            </RGroup>
            <RGroup label="Clear">
              <RBtn icon="minus" label="Contents" onRun={clearContents} />
              <RBtn icon="clearFormat" label="Formats" onRun={clearFormatting} />
            </RGroup>
          </div>
        )}

        {rTab === 'insert' && (
          <div className="ribbon-row">
            <RGroup label="Rows">
              <RBtn icon="rowAbove" label="Above" onRun={() => insertRow(sr)} />
              <RBtn icon="rowBelow" label="Below" onRun={() => insertRow(sr + 1)} />
              <RBtn icon="rowDelete" label="Delete row" onRun={() => deleteRow(sr)} />
            </RGroup>
            <RGroup label="Columns">
              <RBtn icon="colLeft" label="Left" onRun={() => insertCol(sc)} />
              <RBtn icon="colRight" label="Right" onRun={() => insertCol(sc + 1)} />
              <RBtn icon="colDelete" label="Delete col" onRun={() => deleteCol(sc)} />
            </RGroup>
            <RGroup label="Functions">
              <RBtn icon="sigma" label="AutoSum" onRun={autoSum} />
              <RBtn icon="fx" label="AVERAGE" onRun={() => insertFormula('AVERAGE')} />
              <RBtn icon="fx" label="COUNT" onRun={() => insertFormula('COUNT')} />
              <RBtn icon="fx" label="MIN" onRun={() => insertFormula('MIN')} />
              <RBtn icon="fx" label="MAX" onRun={() => insertFormula('MAX')} />
              <RBtn icon="fx" label="ROUND" onRun={() => insertFormula('ROUND')} />
            </RGroup>
          </div>
        )}

        {rTab === 'data' && (
          <div className="ribbon-row">
            <RGroup label="Sort column">
              <RBtn icon="sortAZ" label="A to Z" onRun={() => sortRange(1)} />
              <RBtn icon="sortZA" label="Z to A" onRun={() => sortRange(-1)} />
            </RGroup>
            <RGroup label="Selection">
              <RBtn icon="selectRange" label={rangeMode ? 'Picking range…' : 'Range pick'} active={rangeMode} onRun={() => { setRangeMode(!rangeMode); setRangeEnd(null); flash(rangeMode ? 'Range mode off.' : 'Tap the first cell, then the last cell.'); }} />
              <RBtn icon="copy" label="Copy" onRun={copySel} />
              <RBtn icon="paste" label="Paste" onRun={pasteClip} />
            </RGroup>
            <RGroup label="Clear">
              <RBtn icon="minus" label="Contents" onRun={clearContents} />
              <RBtn icon="clearFormat" label="Formats" onRun={clearFormatting} />
            </RGroup>
          </div>
        )}

        {rTab === 'ai' && (
          <div className="ribbon-row">
            <RGroup label="AI tools">
              <RWide icon="sparkle" label="AI Fill (generate a table)" disabled={aiBusy} onRun={() => setAiOpen(true)} />
            </RGroup>
          </div>
        )}
      </div>

      {aiOpen && (
        <div className="modal" onClick={() => !aiBusy && setAiOpen(false)}>
          <div className="modal-body" onClick={(e) => e.stopPropagation()}>
            <h3>Generate a table</h3>
            <p className="hint">Describe the table. It replaces the current sheet.</p>
            <textarea
              className="input"
              rows={3}
              value={aiPrompt}
              placeholder="e.g. Monthly budget for a student with rent, food, transport"
              onChange={(e) => setAiPrompt(e.target.value)}
            />
            <div className="btn-row">
              <button className="btn primary" disabled={aiBusy || !aiPrompt.trim()} onClick={() => void runAi()}>
                {aiBusy ? 'Generating…' : 'Generate'}
              </button>
              <button className="btn" onClick={() => setAiOpen(false)}>
                Cancel
              </button>
            </div>
          </div>
        </div>
      )}

      {toast && <div className="toast">{toast}</div>}
    </div>
  );
}
