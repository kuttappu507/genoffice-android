import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { FileTypeIcon, Icon } from '../components/Icon';
import { AppBar, Palette, RBtn, RGroup, RSeg, RSelect, RStepper, RWide, RibbonPanel, RibbonTabs } from '../components/Ribbon';
import { BottomSheet, ConfirmSheet, PromptSheet, SheetMenu, Toast, useToast } from '../components/Sheet';
import { FunctionPicker } from '../components/FunctionPicker';
import { ChartView, chartData } from '../components/ChartView';
import { chatStream, errMsg } from '../lib/ai-client';
import { debounce, downloadText, getDoc, getMeta, getPrefs, getSettings, putDoc, uid } from '../lib/storage';
import { exportXlsx, importWorkbook, openFilePicker, saveBinary, sanitizeName, svgToPng } from '../lib/fileio';
import { evalCell, refsIn } from '../lib/formulas';
import { onBack, tap } from '../lib/native';
import {
  Book, CellStyle, Chart, ChartType, Page, Rect,
  DEFAULT_COL_W, DEFAULT_ROW_H, FMT_OPTIONS, MAX_COLS, MAX_ROWS, SHEET_TEMPLATES,
  cleanSheetName, cloneBook, colName, computeDisplay, currentRegion, deleteCols, deleteRows, emptyPage, fmtDisplay, freshBook, hasHeaderRow,
  inRect, insertCols, insertRows, offsetFormula, parseDelimited, parseNumberish, parseRef, rangeToRect, rectOf, rectToRange, refOf, refsInRect, seriesFill, toCsv, uniqueSheetName, usedExtent,
} from '../lib/sheet-model';

type RibbonTab = 'home' | 'insert' | 'formulas' | 'data' | 'view';
type PanelKind = 'text' | 'fill' | 'borders' | 'chart' | 'numfmt' | null;

interface Clip { cells: Record<string, string>; styles: Record<string, CellStyle>; rect: Rect; cut?: boolean }

const ROW_HEAD_W = 40;
const OVERSCAN = 4;

function loadBook(id: string | undefined): Book {
  if (!id) return freshBook();
  const b = getDoc<Book>(id);
  if (b && b.order && b.sheets) {
    for (const n of b.order) if (!b.sheets[n]) b.sheets[n] = emptyPage();
    if (!b.sheets[b.active]) b.active = b.order[0];
    return b;
  }
  // legacy single-sheet format
  const legacy = b as unknown as { cells?: Record<string, string> } | null;
  if (legacy?.cells) return { order: ['Sheet1'], active: 'Sheet1', sheets: { Sheet1: { cells: legacy.cells, styles: {} } } };
  return freshBook();
}

export default function Sheets({ initialId, onExit }: { initialId?: string; onExit?: () => void }) {
  const sheetId = useRef(initialId ?? uid()).current;
  const prefs = useMemo(() => getPrefs(), []);
  const [title, setTitle] = useState(() => (initialId ? getMeta(initialId)?.title ?? 'Spreadsheet' : 'Untitled workbook'));
  const [book, setBook] = useState<Book>(() => loadBook(initialId));
  const [, bump] = useState(0);
  const [sel, setSel] = useState('A1');
  const [rangeEnd, setRangeEnd] = useState<string | null>(null);
  const [editing, setEditing] = useState<string | null>(null);
  const [inlineEdit, setInlineEdit] = useState(false);
  const [rTab, setRTab] = useState<RibbonTab>('home');
  const [panel, setPanel] = useState<PanelKind>(null);
  const [menu, setMenu] = useState(false);
  const [tabMenu, setTabMenu] = useState(false);
  const [clip, setClip] = useState<Clip | null>(null);
  const [aiOpen, setAiOpen] = useState(false);
  const [aiMode, setAiMode] = useState<'table' | 'formula' | 'explain' | 'analyze'>('table');
  const [aiPrompt, setAiPrompt] = useState('');
  const [aiBusy, setAiBusy] = useState(false);
  const [aiOut, setAiOut] = useState('');
  const [toast, flash] = useToast();
  const [fnOpen, setFnOpen] = useState(false);
  const [goOpen, setGoOpen] = useState(false);
  const [renameOpen, setRenameOpen] = useState(false);
  const [confirmDelSheet, setConfirmDelSheet] = useState(false);
  const [noteOpen, setNoteOpen] = useState(false);
  const [findOpen, setFindOpen] = useState(false);
  const [fq, setFq] = useState('');
  const [fr, setFr] = useState('');
  const [hits, setHits] = useState<string[]>([]);
  const [hitIdx, setHitIdx] = useState(0);
  const [filterCol, setFilterCol] = useState<number | null>(null);
  const [chartEdit, setChartEdit] = useState<Chart | null>(null);
  const [chartView, setChartView] = useState<Chart | null>(null);
  const [templatesOpen, setTemplatesOpen] = useState(!initialId);
  const [zoom, setZoom] = useState(100);
  const [showGrid, setShowGrid] = useState(true);
  const [showFormulas, setShowFormulas] = useState(false);
  const [sortOpen, setSortOpen] = useState(false);
  const [statsOpen, setStatsOpen] = useState(false);
  const [saved, setSaved] = useState<'saved' | 'saving' | 'dirty'>('saved');
  const [viewport, setViewport] = useState({ x: 0, y: 0, w: 360, h: 400 });

  const past = useRef<Book[]>([]);
  const future = useRef<Book[]>([]);
  const gridRef = useRef<HTMLDivElement>(null);
  const fxRef = useRef<HTMLInputElement>(null);
  const inlineRef = useRef<HTMLInputElement>(null);
  const press = useRef<{ timer: ReturnType<typeof setTimeout> | null; ref: string; moved: boolean; x: number; y: number } | null>(null);
  const dragging = useRef<null | 'range' | 'fill'>(null);
  const colResize = useRef<{ c: number; startX: number; startW: number } | null>(null);
  const rowResize = useRef<{ r: number; startY: number; startH: number } | null>(null);
  const [fillPreview, setFillPreview] = useState<Rect | null>(null);

  const page: Page = book.sheets[book.active] ?? emptyPage();
  const cells = page.cells;
  const styles = page.styles;
  const currency = prefs.currency;

  // ------------------------------------------------------------------ persistence / history
  const save = useMemo(() => debounce((b: Book, t: string) => { putDoc<Book>('sheet', sheetId, t, b); setSaved('saved'); }, 600), [sheetId]);
  useEffect(() => { save(book, title); }, [title]); // eslint-disable-line react-hooks/exhaustive-deps

  const pushU = (b: Book) => {
    past.current.push(cloneBook(b));
    if (past.current.length > 60) past.current.shift();
    future.current = [];
  };
  const mutate = (fn: (b: Book) => Book, opts: { silent?: boolean } = {}) => {
    if (!opts.silent) pushU(book);
    const nb = fn(book);
    setBook(nb);
    setSaved('dirty');
    save(nb, title);
    bump((v) => v + 1);
  };
  const mutatePage = (fn: (pg: Page) => Page) => mutate((b) => ({ ...b, sheets: { ...b.sheets, [b.active]: fn(b.sheets[b.active] ?? emptyPage()) } }));
  const undo = () => {
    const p = past.current.pop();
    if (!p) return;
    future.current.push(cloneBook(book));
    setBook(p); save(p, title); bump((v) => v + 1);
  };
  const redo = () => {
    const f = future.current.pop();
    if (!f) return;
    past.current.push(cloneBook(book));
    setBook(f); save(f, title); bump((v) => v + 1);
  };

  // ------------------------------------------------------------------ evaluation (one pass per change)
  const display = useMemo(() => computeDisplay(cells, book, book.active), [cells, book]);
  const shown = useCallback((ref: string) => fmtDisplay(display[ref] ?? '', styles[ref], currency), [display, styles, currency]);

  // ------------------------------------------------------------------ selection
  const rg = useMemo(() => rectOf(sel, rangeEnd ?? undefined), [sel, rangeEnd]);
  const [sc, sr] = parseRef(sel);
  const selRefs = () => refsInRect(rg);
  const curStyle = styles[sel];
  const mergeMap = useMemo(() => {
    const anchors = new Map<string, Rect>();
    const covered = new Map<string, string>();
    for (const m of page.merges ?? []) {
      const r = rangeToRect(m);
      if (!r) continue;
      anchors.set(refOf(r.c1, r.r1), r);
      for (const ref of refsInRect(r)) if (ref !== refOf(r.c1, r.r1)) covered.set(ref, refOf(r.c1, r.r1));
    }
    return { anchors, covered };
  }, [page.merges]);

  const select = (ref: string, end: string | null = null) => {
    const anchor = mergeMap.covered.get(ref) ?? ref;
    setSel(anchor);
    setRangeEnd(end);
    setEditing(cells[anchor] ?? '');
    setInlineEdit(false);
  };

  // hidden rows from filter
  const hiddenRows = useMemo(() => {
    const f = page.filter;
    const hidden = new Set<number>();
    if (!f) return hidden;
    const rect = rangeToRect(f.range);
    if (!rect) return hidden;
    for (let r = rect.r1 + 1; r <= rect.r2; r++) {
      for (const [cStr, allowed] of Object.entries(f.criteria)) {
        if (!allowed.length) continue;
        const v = shown(refOf(Number(cStr), r));
        if (!allowed.includes(v)) { hidden.add(r); break; }
      }
    }
    return hidden;
  }, [page.filter, shown]);

  // ------------------------------------------------------------------ geometry (virtualisation)
  const z = zoom / 100;
  const colW = (c: number) => Math.round((page.colW?.[c] ?? DEFAULT_COL_W) * z);
  const rowH = (r: number) => (hiddenRows.has(r) ? 0 : Math.round((page.rowH?.[r] ?? DEFAULT_ROW_H) * z));
  const extent = useMemo(() => usedExtent(page), [page]);
  const nCols = Math.min(MAX_COLS, Math.max(extent.cols + 6, 12, sc + 4));
  const nRows = Math.min(MAX_ROWS, Math.max(extent.rows + 20, 40, sr + 10));
  const colX = useMemo(() => { const xs = [0]; for (let c = 0; c < nCols; c++) xs.push(xs[c] + colW(c)); return xs; }, [nCols, page.colW, z]); // eslint-disable-line react-hooks/exhaustive-deps
  const rowY = useMemo(() => { const ys = [0]; for (let r = 0; r < nRows; r++) ys.push(ys[r] + rowH(r)); return ys; }, [nRows, page.rowH, z, hiddenRows]); // eslint-disable-line react-hooks/exhaustive-deps
  const totalW = colX[nCols];
  const totalH = rowY[nRows];
  const headH = Math.round(30 * z);
  const freeze = page.freeze ?? { r: 0, c: 0 };
  const frozenW = colX[Math.min(freeze.c, nCols)];
  const frozenH = rowY[Math.min(freeze.r, nRows)];

  const findIdx = (arr: number[], v: number) => { let lo = 0, hi = arr.length - 2; while (lo < hi) { const mid = (lo + hi + 1) >> 1; if (arr[mid] <= v) lo = mid; else hi = mid - 1; } return lo; };
  const firstCol = Math.max(freeze.c, findIdx(colX, viewport.x + frozenW) - OVERSCAN);
  const lastCol = Math.min(nCols - 1, findIdx(colX, viewport.x + viewport.w) + OVERSCAN);
  const firstRow = Math.max(freeze.r, findIdx(rowY, viewport.y + frozenH) - OVERSCAN);
  const lastRow = Math.min(nRows - 1, findIdx(rowY, viewport.y + viewport.h) + OVERSCAN);

  useEffect(() => {
    const el = gridRef.current;
    if (!el) return;
    // Native scroll must yield to range/fill/resize drags; touch-action can't change mid-gesture,
    // so cancel touchmove explicitly (non-passive) while a drag is active.
    const onTouchMove = (ev: TouchEvent) => { if (dragging.current || colResize.current || rowResize.current) ev.preventDefault(); };
    el.addEventListener('touchmove', onTouchMove, { passive: false });
    const onScroll = () => setViewport({ x: el.scrollLeft, y: el.scrollTop, w: el.clientWidth, h: el.clientHeight });
    onScroll();
    el.addEventListener('scroll', onScroll, { passive: true });
    const ro = new ResizeObserver(onScroll);
    ro.observe(el);
    return () => { el.removeEventListener('scroll', onScroll); el.removeEventListener('touchmove', onTouchMove); ro.disconnect(); };
  }, []);

  const scrollToCell = (c: number, r: number) => {
    const el = gridRef.current;
    if (!el) return;
    const x = colX[c] ?? 0, y = rowY[r] ?? 0;
    const right = x + colW(c), bottom = y + rowH(r);
    let sx = el.scrollLeft, sy = el.scrollTop;
    if (x < sx + frozenW && c >= freeze.c) sx = x - frozenW;
    else if (right > sx + el.clientWidth - ROW_HEAD_W) sx = right - el.clientWidth + ROW_HEAD_W + 8;
    if (y < sy + frozenH && r >= freeze.r) sy = y - frozenH;
    else if (bottom > sy + el.clientHeight - headH) sy = bottom - el.clientHeight + headH + 8;
    if (sx !== el.scrollLeft || sy !== el.scrollTop) el.scrollTo({ left: Math.max(0, sx), top: Math.max(0, sy) });
  };

  // Android back
  useEffect(
    () =>
      onBack(() => {
        if (panel) { setPanel(null); return true; }
        if (inlineEdit) { setInlineEdit(false); return true; }
        if (findOpen) { setFindOpen(false); setHits([]); return true; }
        if (rangeEnd) { setRangeEnd(null); return true; }
        onExit?.();
        return true;
      }),
    [panel, inlineEdit, findOpen, rangeEnd, onExit],
  );

  // ------------------------------------------------------------------ pointer handling on the grid
  const cellAt = (clientX: number, clientY: number): string | null => {
    const el = gridRef.current;
    if (!el) return null;
    const b = el.getBoundingClientRect();
    const lx = clientX - b.left, ly = clientY - b.top;
    if (lx < ROW_HEAD_W || ly < headH) return null;
    const gx = lx - ROW_HEAD_W, gy = ly - headH;
    const x = gx < frozenW ? gx : gx + el.scrollLeft;
    const y = gy < frozenH ? gy : gy + el.scrollTop;
    if (x > totalW || y > totalH) return null;
    const c = findIdx(colX, x), r = findIdx(rowY, y);
    return refOf(c, r);
  };

  const onGridPointerDown = (e: React.PointerEvent) => {
    const t = e.target as HTMLElement;
    if (t.closest('.col-resize') || t.closest('.row-resize') || t.closest('.fill-handle') || t.closest('.chart-card') || t.closest('.filter-btn')) return;
    const head = t.closest('[data-col]') as HTMLElement | null;
    const rhead = t.closest('[data-row]') as HTMLElement | null;
    if (head) { const c = Number(head.dataset.col); select(refOf(c, 0), refOf(c, nRows - 1)); return; }
    if (rhead) { const r = Number(rhead.dataset.row); select(refOf(0, r), refOf(nCols - 1, r)); return; }
    if (t.closest('.corner')) { select('A1', refOf(nCols - 1, nRows - 1)); return; }
    const ref = cellAt(e.clientX, e.clientY);
    if (!ref) return;
    (e.currentTarget as HTMLElement).setPointerCapture?.(e.pointerId);
    press.current = { ref, moved: false, x: e.clientX, y: e.clientY, timer: setTimeout(() => {
      // long-press → start range drag from this cell
      if (press.current && !press.current.moved) {
        dragging.current = 'range';
        select(ref, ref);
        void tap('medium');
      }
    }, 420) };
    if (e.shiftKey) { setRangeEnd(ref); if (press.current.timer) clearTimeout(press.current.timer); press.current = null; }
  };
  const onGridPointerMove = (e: React.PointerEvent) => {
    if (colResize.current) {
      const { c, startX, startW } = colResize.current;
      const w = Math.max(28, Math.round(startW + (e.clientX - startX) / z));
      setBook((b) => ({ ...b, sheets: { ...b.sheets, [b.active]: { ...b.sheets[b.active], colW: { ...b.sheets[b.active].colW, [c]: w } } } }));
      return;
    }
    if (rowResize.current) {
      const { r, startY, startH } = rowResize.current;
      const h = Math.max(20, Math.round(startH + (e.clientY - startY) / z));
      setBook((b) => ({ ...b, sheets: { ...b.sheets, [b.active]: { ...b.sheets[b.active], rowH: { ...b.sheets[b.active].rowH, [r]: h } } } }));
      return;
    }
    const p = press.current;
    if (p && !p.moved && Math.hypot(e.clientX - p.x, e.clientY - p.y) > 8) {
      p.moved = true;
      if (p.timer) clearTimeout(p.timer);
      if (dragging.current !== 'range' && dragging.current !== 'fill') { press.current = null; return; } // plain scroll
    }
    if (dragging.current === 'range' || dragging.current === 'fill') {
      e.preventDefault();
      const ref = cellAt(e.clientX, e.clientY);
      if (!ref) return;
      if (dragging.current === 'range') setRangeEnd(ref === sel ? null : ref);
      else {
        const [c, r] = parseRef(ref);
        // fill only along one axis from the source rect
        const down = Math.abs(r - rg.r2) >= Math.abs(c - rg.c2) || (r > rg.r2 && c >= rg.c1 && c <= rg.c2);
        setFillPreview(down ? { c1: rg.c1, c2: rg.c2, r1: Math.min(rg.r1, r), r2: Math.max(rg.r2, r) } : { r1: rg.r1, r2: rg.r2, c1: Math.min(rg.c1, c), c2: Math.max(rg.c2, c) });
      }
    }
  };
  const onGridPointerUp = (e: React.PointerEvent) => {
    if (colResize.current || rowResize.current) { colResize.current = null; rowResize.current = null; save(book, title); pushU(book); return; }
    const p = press.current;
    press.current = null;
    if (p?.timer) clearTimeout(p.timer);
    if (dragging.current === 'fill') { dragging.current = null; applyFill(); return; }
    if (dragging.current === 'range') { dragging.current = null; void tap('light'); return; }
    if (!p || p.moved) return;
    const ref = p.ref;
    if (ref === sel && !rangeEnd && editing !== null) {
      // second tap on the selected cell → inline edit
      setInlineEdit(true);
      setTimeout(() => inlineRef.current?.focus(), 30);
      return;
    }
    if (rangeEnd && e.shiftKey) { setRangeEnd(ref); return; }
    select(ref);
  };

  useEffect(() => { scrollToCell(sc, sr); }, [sel]); // eslint-disable-line react-hooks/exhaustive-deps

  // ------------------------------------------------------------------ edits
  const commit = (ref: string, raw: string) => {
    if ((cells[ref] ?? '') === raw) return;
    mutatePage((pg) => {
      const c = { ...pg.cells };
      if (raw === '') delete c[ref]; else c[ref] = raw;
      return { ...pg, cells: c };
    });
  };
  const commitAndMove = (dc: number, dr: number, live?: string) => {
    const v = live !== undefined ? live : editing;
    if (v !== null && v !== undefined) commit(sel, v);
    setInlineEdit(false);
    let [c, r] = parseRef(sel);
    c = Math.max(0, Math.min(MAX_COLS - 1, c + dc));
    r = Math.max(0, Math.min(MAX_ROWS - 1, r + dr));
    while (hiddenRows.has(r) && r < MAX_ROWS - 1 && dr > 0) r++;
    const next = refOf(c, r);
    setSel(next); setRangeEnd(null); setEditing(cells[next] ?? '');
  };
  const setRange = (fn: (pg: Page, refs: string[]) => Page) => mutatePage((pg) => fn(pg, selRefs()));

  const applyStyle = (patch: Partial<CellStyle>) =>
    setRange((pg, refs) => {
      const st = { ...pg.styles };
      for (const ref of refs) {
        const next: CellStyle = { ...st[ref], ...patch };
        (Object.keys(next) as (keyof CellStyle)[]).forEach((k) => { if (next[k] === undefined || next[k] === null || next[k] === false) delete next[k]; });
        if (Object.keys(next).length === 0) delete st[ref]; else st[ref] = next;
      }
      return { ...pg, styles: st };
    });
  const toggle = (k: 'b' | 'i' | 'u' | 's' | 'wrap') => applyStyle({ [k]: !curStyle?.[k] });
  const setBorders = (kind: 'all' | 'outline' | 'top' | 'bottom' | 'left' | 'right' | 'none') =>
    setRange((pg) => {
      const st = { ...pg.styles };
      const put = (ref: string, p: Partial<CellStyle>) => { st[ref] = { ...st[ref], ...p }; };
      for (let r = rg.r1; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) {
        const ref = refOf(c, r);
        if (kind === 'none') { const s = { ...st[ref] }; delete s.bt; delete s.bb; delete s.bl; delete s.br; if (Object.keys(s).length) st[ref] = s; else delete st[ref]; continue; }
        if (kind === 'all') put(ref, { bt: true, bb: true, bl: true, br: true });
        if (kind === 'outline' || kind === 'top') if (r === rg.r1) put(ref, { bt: true });
        if (kind === 'outline' || kind === 'bottom') if (r === rg.r2) put(ref, { bb: true });
        if (kind === 'outline' || kind === 'left') if (c === rg.c1) put(ref, { bl: true });
        if (kind === 'outline' || kind === 'right') if (c === rg.c2) put(ref, { br: true });
      }
      return { ...pg, styles: st };
    });
  const clearContents = () => { setRange((pg, refs) => { const c = { ...pg.cells }; refs.forEach((r) => delete c[r]); return { ...pg, cells: c }; }); };
  const clearFormats = () => { setRange((pg, refs) => { const s = { ...pg.styles }; refs.forEach((r) => delete s[r]); return { ...pg, styles: s }; }); };
  const clearAll = () => { setRange((pg, refs) => { const c = { ...pg.cells }; const s = { ...pg.styles }; const n = { ...pg.notes }; refs.forEach((r) => { delete c[r]; delete s[r]; delete n[r]; }); return { ...pg, cells: c, styles: s, notes: n }; }); };
  const changeDec = (d: 1 | -1) => {
    const cur = curStyle?.dec ?? (curStyle?.fmt === 'pct' ? 0 : curStyle?.fmt === 'num' || curStyle?.fmt === 'cur' ? 2 : 0);
    applyStyle({ dec: Math.max(0, Math.min(8, cur + d)), fmt: curStyle?.fmt && curStyle.fmt !== 'gen' ? curStyle.fmt : 'num' });
  };

  const mergeCells = () => {
    if (rg.c1 === rg.c2 && rg.r1 === rg.r2) { flash('Select a range of cells to merge.'); return; }
    const range = rectToRange(rg);
    mutatePage((pg) => {
      const existing = (pg.merges ?? []).filter((m) => { const r = rangeToRect(m); return !r || !(r.c1 <= rg.c2 && r.c2 >= rg.c1 && r.r1 <= rg.r2 && r.r2 >= rg.r1); });
      const anchor = refOf(rg.c1, rg.r1);
      const c = { ...pg.cells };
      // keep top-left value; drop the rest (Excel behaviour)
      for (const ref of refsInRect(rg)) if (ref !== anchor) delete c[ref];
      return { ...pg, cells: c, merges: [...existing, range], styles: { ...pg.styles, [anchor]: { ...pg.styles[anchor], align: pg.styles[anchor]?.align ?? 'center' } } };
    });
    setRangeEnd(null);
  };
  const unmergeCells = () => {
    mutatePage((pg) => ({ ...pg, merges: (pg.merges ?? []).filter((m) => { const r = rangeToRect(m); return !r || !inRect(sc, sr, r); }) }));
  };
  const isMerged = mergeMap.anchors.has(sel);

  // ------------------------------------------------------------------ clipboard
  const copySel = (cut = false) => {
    const c: Record<string, string> = {}; const s: Record<string, CellStyle> = {};
    for (const ref of selRefs()) { if (cells[ref] !== undefined) c[ref] = cells[ref]; if (styles[ref]) s[ref] = { ...styles[ref] }; }
    setClip({ cells: c, styles: s, rect: rg, cut });
    // also to system clipboard as TSV
    const rows: string[] = [];
    for (let r = rg.r1; r <= rg.r2; r++) { const cols: string[] = []; for (let cc = rg.c1; cc <= rg.c2; cc++) cols.push(shown(refOf(cc, r))); rows.push(cols.join('\t')); }
    void navigator.clipboard?.writeText(rows.join('\n')).catch(() => undefined);
    flash(`${cut ? 'Cut' : 'Copied'} ${rectToRange(rg)}.`);
  };
  const pasteClip = async (mode: 'all' | 'values' | 'formats' = 'all') => {
    let source = clip;
    if (!source) {
      // try text from the system clipboard (TSV / CSV from other apps)
      try {
        const txt = await navigator.clipboard?.readText();
        if (txt) { pasteText(txt); return; }
      } catch { /* denied */ }
      flash('Nothing to paste. Copy a cell or range first.');
      return;
    }
    const [c0, r0] = parseRef(sel);
    const dc = c0 - source.rect.c1, dr = r0 - source.rect.r1;
    // tile the clip over the selection if the selection is a multiple of the clip size
    const cw = source.rect.c2 - source.rect.c1 + 1, ch = source.rect.r2 - source.rect.r1 + 1;
    const tilesX = rangeEnd && (rg.c2 - rg.c1 + 1) % cw === 0 ? (rg.c2 - rg.c1 + 1) / cw : 1;
    const tilesY = rangeEnd && (rg.r2 - rg.r1 + 1) % ch === 0 ? (rg.r2 - rg.r1 + 1) / ch : 1;
    mutatePage((pg) => {
      const c = { ...pg.cells }; const s = { ...pg.styles };
      if (source!.cut) for (const k of Object.keys(source!.cells)) delete c[k];
      for (let tx = 0; tx < tilesX; tx++) for (let ty = 0; ty < tilesY; ty++) {
        const odc = dc + tx * cw, odr = dr + ty * ch;
        if (mode !== 'formats') for (const [k, v] of Object.entries(source!.cells)) {
          const [cc, rr] = parseRef(k); const dst = refOf(cc + odc, rr + odr);
          if (cc + odc >= MAX_COLS || rr + odr >= MAX_ROWS) continue;
          c[dst] = mode === 'values' ? display[k] ?? v : source!.cut ? v : offsetFormula(v, odc, odr);
        }
        if (mode !== 'values') for (const [k, v] of Object.entries(source!.styles)) { const [cc, rr] = parseRef(k); s[refOf(cc + odc, rr + odr)] = { ...v }; }
      }
      return { ...pg, cells: c, styles: s };
    });
    if (source.cut) setClip(null);
    setRangeEnd(rangeEnd ? rangeEnd : refOf(c0 + cw * tilesX - 1, r0 + ch * tilesY - 1));
  };
  const pasteText = (txt: string) => {
    const rows = parseDelimited(txt);
    if (!rows.length) return;
    const [c0, r0] = parseRef(sel);
    mutatePage((pg) => {
      const c = { ...pg.cells };
      rows.forEach((row, ri) => row.forEach((v, ci) => { if (c0 + ci < MAX_COLS && r0 + ri < MAX_ROWS) { if (v === '') delete c[refOf(c0 + ci, r0 + ri)]; else c[refOf(c0 + ci, r0 + ri)] = v; } }));
      return { ...pg, cells: c };
    });
    setRangeEnd(refOf(c0 + Math.max(...rows.map((r) => r.length)) - 1, r0 + rows.length - 1));
    flash(`Pasted ${rows.length} row(s).`);
  };

  // ------------------------------------------------------------------ fill
  const fill = (dir: 'down' | 'right', target?: Rect) => {
    const t = target ?? rg;
    mutatePage((pg) => {
      const c = { ...pg.cells }; const s = { ...pg.styles };
      if (dir === 'down') {
        for (let col = t.c1; col <= t.c2; col++) {
          const seedRows: number[] = []; for (let r = rg.r1; r <= rg.r2; r++) seedRows.push(r);
          const seeds = seedRows.map((r) => pg.cells[refOf(col, r)] ?? '');
          if (seeds.every((x) => x === '')) continue;
          const count = t.r2 - rg.r2;
          const vals = seriesFill(seeds, count, 'down');
          vals.forEach((v, i) => { const dst = refOf(col, rg.r2 + 1 + i); if (v === '') delete c[dst]; else c[dst] = v; const st = pg.styles[refOf(col, rg.r1 + (i % seeds.length))]; if (st) s[dst] = { ...st }; });
        }
      } else {
        for (let r = t.r1; r <= t.r2; r++) {
          const seeds: string[] = []; for (let col = rg.c1; col <= rg.c2; col++) seeds.push(pg.cells[refOf(col, r)] ?? '');
          if (seeds.every((x) => x === '')) continue;
          const vals = seriesFill(seeds, t.c2 - rg.c2, 'right');
          vals.forEach((v, i) => { const dst = refOf(rg.c2 + 1 + i, r); if (v === '') delete c[dst]; else c[dst] = v; const st = pg.styles[refOf(rg.c1 + (i % seeds.length), r)]; if (st) s[dst] = { ...st }; });
        }
      }
      return { ...pg, cells: c, styles: s };
    });
    setRangeEnd(refOf(t.c2, t.r2));
    setSel(refOf(t.c1, t.r1));
  };
  const applyFill = () => {
    const t = fillPreview;
    setFillPreview(null);
    if (!t) return;
    if (t.r2 > rg.r2) fill('down', t);
    else if (t.c2 > rg.c2) fill('right', t);
  };
  const fillDownQuick = () => {
    // Excel Ctrl+D behaviour: with a single cell, extend to the bottom of the neighbouring column's data
    if (rangeEnd) { fill('down'); return; }
    let end = sr;
    const nb = sc > 0 ? sc - 1 : sc + 1;
    while (end + 1 < MAX_ROWS && (cells[refOf(nb, end + 1)] ?? '') !== '') end++;
    if (end === sr) { flash('Select the range to fill (long-press and drag), or fill next to a data column.'); return; }
    fill('down', { c1: sc, c2: sc, r1: sr, r2: end });
  };

  // ------------------------------------------------------------------ structure
  const rowsInSel = rg.r2 - rg.r1 + 1, colsInSel = rg.c2 - rg.c1 + 1;
  const doInsertRows = (at: number) => mutatePage((pg) => insertRows(pg, at, rowsInSel));
  const doDeleteRows = () => mutatePage((pg) => deleteRows(pg, rg.r1, rowsInSel));
  const doInsertCols = (at: number) => mutatePage((pg) => insertCols(pg, at, colsInSel));
  const doDeleteCols = () => mutatePage((pg) => deleteCols(pg, rg.c1, colsInSel));
  const autoFitCol = () => {
    mutatePage((pg) => {
      const cw = { ...pg.colW };
      for (let c = rg.c1; c <= rg.c2; c++) {
        let max = 4;
        for (let r = 0; r < nRows; r++) { const v = shown(refOf(c, r)); if (v.length > max) max = v.length; }
        cw[c] = Math.min(400, Math.max(48, Math.round(max * 7.2 * ((pg.styles[refOf(c, 0)]?.size ?? 14) / 14) + 18)));
      }
      return { ...pg, colW: cw };
    });
  };
  const setFreeze = (kind: 'none' | 'top' | 'first' | 'here') =>
    mutatePage((pg) => ({ ...pg, freeze: kind === 'none' ? undefined : kind === 'top' ? { r: 1, c: 0 } : kind === 'first' ? { r: 0, c: 1 } : { r: sr, c: sc } }));

  // ------------------------------------------------------------------ sort / filter
  const dataRect = (): Rect => (rangeEnd ? rg : currentRegion(cells, sc, sr));
  const sortBy = (col: number, dir: 1 | -1, rect = dataRect()) => {
    const header = hasHeaderRow(cells, rect);
    const start = header ? rect.r1 + 1 : rect.r1;
    if (rect.r2 - start < 1) { flash('Need at least two rows to sort.'); return; }
    const rows: { r: number; key: string; num: number | null }[] = [];
    for (let r = start; r <= rect.r2; r++) { const d = display[refOf(col, r)] ?? cells[refOf(col, r)] ?? ''; rows.push({ r, key: d, num: parseNumberish(d) }); }
    rows.sort((a, b) => {
      if (a.num !== null && b.num !== null) return (a.num - b.num) * dir;
      if (a.num !== null) return -1 * dir;
      if (b.num !== null) return 1 * dir;
      if (a.key === '' && b.key !== '') return 1; if (b.key === '' && a.key !== '') return -1;
      return a.key.localeCompare(b.key, undefined, { numeric: true, sensitivity: 'base' }) * dir;
    });
    mutatePage((pg) => {
      const c = { ...pg.cells }; const s = { ...pg.styles }; const n = { ...pg.notes };
      const srcRows = rows.map((x) => x.r);
      const snapshot: Record<string, string> = {}; const snapS: Record<string, CellStyle> = {}; const snapN: Record<string, string> = {};
      for (const r of srcRows) for (let cc = rect.c1; cc <= rect.c2; cc++) { const k = refOf(cc, r); snapshot[k] = pg.cells[k]; if (pg.styles[k]) snapS[k] = pg.styles[k]; if (pg.notes?.[k]) snapN[k] = pg.notes[k]; }
      rows.forEach((row, i) => {
        const dstR = start + i;
        for (let cc = rect.c1; cc <= rect.c2; cc++) {
          const src = refOf(cc, row.r); const dst = refOf(cc, dstR);
          const v = snapshot[src];
          if (v === undefined) delete c[dst]; else c[dst] = offsetFormula(v, 0, dstR - row.r);
          if (snapS[src]) s[dst] = snapS[src]; else delete s[dst];
          if (snapN[src]) n[dst] = snapN[src]; else delete n[dst];
        }
      });
      return { ...pg, cells: c, styles: s, notes: n };
    });
    setSel(refOf(rect.c1, rect.r1)); setRangeEnd(refOf(rect.c2, rect.r2));
    flash(`Sorted by column ${colName(col)} (${dir === 1 ? 'A→Z' : 'Z→A'}).`);
  };
  const toggleFilter = () => {
    if (page.filter) { mutatePage((pg) => ({ ...pg, filter: undefined })); flash('Filter removed.'); return; }
    const rect = dataRect();
    if (rect.r1 === rect.r2) { flash('Tap inside your table first.'); return; }
    mutatePage((pg) => ({ ...pg, filter: { range: rectToRange(rect), criteria: {} } }));
    flash('Filter buttons added to the header row.');
  };
  const filterValues = useMemo(() => {
    if (filterCol === null || !page.filter) return [] as { v: string; n: number }[];
    const rect = rangeToRect(page.filter.range); if (!rect) return [];
    const counts = new Map<string, number>();
    for (let r = rect.r1 + 1; r <= rect.r2; r++) { const v = shown(refOf(filterCol, r)); counts.set(v, (counts.get(v) ?? 0) + 1); }
    return Array.from(counts.entries()).map(([v, n]) => ({ v, n })).sort((a, b) => a.v.localeCompare(b.v, undefined, { numeric: true }));
  }, [filterCol, page.filter, shown]);
  const setCriteria = (col: number, allowed: string[] | null) =>
    mutatePage((pg) => { if (!pg.filter) return pg; const crit = { ...pg.filter.criteria }; if (!allowed) delete crit[col]; else crit[col] = allowed; return { ...pg, filter: { ...pg.filter, criteria: crit } }; });
  const removeDuplicates = () => {
    const rect = dataRect();
    const header = hasHeaderRow(cells, rect);
    const start = header ? rect.r1 + 1 : rect.r1;
    const seen = new Set<string>(); const keep: number[] = [];
    for (let r = start; r <= rect.r2; r++) { const key = Array.from({ length: rect.c2 - rect.c1 + 1 }, (_, i) => cells[refOf(rect.c1 + i, r)] ?? '').join('\u0001'); if (!seen.has(key)) { seen.add(key); keep.push(r); } }
    const removed = rect.r2 - start + 1 - keep.length;
    if (!removed) { flash('No duplicate rows found.'); return; }
    mutatePage((pg) => {
      const c = { ...pg.cells };
      const vals = keep.map((r) => Array.from({ length: rect.c2 - rect.c1 + 1 }, (_, i) => pg.cells[refOf(rect.c1 + i, r)]));
      for (let r = start; r <= rect.r2; r++) for (let cc = rect.c1; cc <= rect.c2; cc++) delete c[refOf(cc, r)];
      vals.forEach((row, i) => row.forEach((v, j) => { if (v !== undefined) c[refOf(rect.c1 + j, start + i)] = v; }));
      return { ...pg, cells: c };
    });
    flash(`Removed ${removed} duplicate row(s).`);
  };
  const textToColumns = () => {
    const rect = rangeEnd ? rg : { c1: sc, c2: sc, r1: sr, r2: sr };
    mutatePage((pg) => {
      const c = { ...pg.cells };
      for (let r = rect.r1; r <= rect.r2; r++) {
        const v = pg.cells[refOf(rect.c1, r)] ?? '';
        const parts = v.includes('\t') ? v.split('\t') : v.includes(',') ? v.split(',') : v.split(/\s+/);
        parts.forEach((p, i) => { if (rect.c1 + i < MAX_COLS) c[refOf(rect.c1 + i, r)] = p.trim(); });
      }
      return { ...pg, cells: c };
    });
    flash('Split text into columns.');
  };

  // ------------------------------------------------------------------ formulas
  const autoRange = (): string => {
    if (rangeEnd) return rectToRange(rg);
    let top = sr;
    while (top - 1 >= 0 && (cells[refOf(sc, top - 1)] ?? '') !== '') top--;
    if (top < sr) return `${refOf(sc, top)}:${refOf(sc, sr - 1)}`;
    let left = sc;
    while (left - 1 >= 0 && (cells[refOf(left - 1, sr)] ?? '') !== '') left--;
    if (left < sc) return `${refOf(left, sr)}:${refOf(sc - 1, sr)}`;
    return refOf(sc, Math.max(0, sr - 1));
  };
  const insertFn = (name: string, sig?: string) => {
    // Quick aggregates go straight in; others open the editor with the signature template.
    const quick = ['SUM', 'AVERAGE', 'COUNT', 'MIN', 'MAX', 'COUNTA', 'MEDIAN', 'PRODUCT'];
    let formula: string;
    if (rangeEnd && quick.includes(name)) {
      // aggregate below the selection
      formula = `=${name}(${rectToRange(rg)})`;
      const dst = refOf(rg.c1, rg.r2 + 1);
      commit(dst, formula); select(dst); flash(`${name} added in ${dst}.`); return;
    }
    if (quick.includes(name)) formula = `=${name}(${autoRange()})`;
    else formula = `=${(sig ?? `${name}()`)}`;
    setEditing(formula);
    setInlineEdit(false);
    setTimeout(() => { const el = fxRef.current; if (!el) return; el.focus(); const p = formula.indexOf('('); el.setSelectionRange(p + 1, formula.length - 1); }, 30);
  };
  const toggleAbs = () => {
    // F4 behaviour on the formula editor: cycle A1 → $A$1 → A$1 → $A1
    const el = fxRef.current; const v = editing ?? '';
    if (!el || !v.startsWith('=')) return;
    const pos = el.selectionStart ?? v.length;
    const re = /\$?[A-Za-z]{1,3}\$?\d{1,5}/g; let m: RegExpExecArray | null;
    while ((m = re.exec(v))) {
      if (pos >= m.index && pos <= m.index + m[0].length) {
        const t = m[0]; const col = t.replace(/\$/g, '').match(/[A-Za-z]+/)![0]; const row = t.replace(/\$/g, '').match(/\d+/)![0];
        const next = t.startsWith('$') && t.includes(`$${row}`) ? `${col}$${row}` : t.startsWith(col) && t.includes(`$${row}`) ? `$${col}${row}` : t.startsWith('$') ? `${col}${row}` : `$${col}$${row}`;
        const nv = v.slice(0, m.index) + next + v.slice(m.index + t.length);
        setEditing(nv); setTimeout(() => { el.focus(); el.setSelectionRange(m!.index + next.length, m!.index + next.length); }, 0);
        return;
      }
    }
  };
  const showPrecedents = useMemo(() => (editing?.startsWith('=') ? new Set(refsIn(editing)) : new Set<string>()), [editing]);

  // ------------------------------------------------------------------ find & replace
  const runFind = () => {
    const q = fq.trim().toLowerCase(); if (!q) { setHits([]); return; }
    const found = Object.keys(cells).filter((k) => (cells[k] ?? '').toLowerCase().includes(q) || (display[k] ?? '').toLowerCase().includes(q))
      .sort((a, b) => { const [ca, ra] = parseRef(a); const [cb, rb] = parseRef(b); return ra - rb || ca - cb; });
    setHits(found); setHitIdx(0);
    if (found.length) select(found[0]); else flash('No matches.');
  };
  const nextHit = (d: 1 | -1) => { if (!hits.length) return; const i = (hitIdx + d + hits.length) % hits.length; setHitIdx(i); select(hits[i]); };
  const replaceOne = () => { if (!hits.length) return; const ref = hits[hitIdx]; const re = new RegExp(fq.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'ig'); commit(ref, (cells[ref] ?? '').replace(re, fr)); const rest = hits.filter((h) => h !== ref); setHits(rest); if (rest.length) select(rest[Math.min(hitIdx, rest.length - 1)]); };
  const replaceAll = () => { if (!hits.length) return; const re = new RegExp(fq.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'ig'); mutatePage((pg) => { const c = { ...pg.cells }; hits.forEach((h) => { c[h] = (c[h] ?? '').replace(re, fr); if (c[h] === '') delete c[h]; }); return { ...pg, cells: c }; }); flash(`Replaced in ${hits.length} cell(s).`); setHits([]); };

  // ------------------------------------------------------------------ sheets
  const addSheet = () => { const name = uniqueSheetName(book.order, `Sheet${book.order.length + 1}`); mutate((b) => ({ ...b, order: [...b.order, name], active: name, sheets: { ...b.sheets, [name]: emptyPage() } })); setSel('A1'); setRangeEnd(null); setInlineEdit(false); setEditing(''); };
  const switchSheet = (name: string) => { if (name === book.active) return; mutate((b) => ({ ...b, active: name }), { silent: true }); const c = book.sheets[name]?.cells ?? {}; setSel('A1'); setRangeEnd(null); setEditing(c.A1 ?? ''); setInlineEdit(false); };
  const renameSheet = (name: string) => {
    const clean = cleanSheetName(name);
    if (!clean || clean === book.active) return;
    if (book.order.includes(clean)) { flash('A sheet with that name already exists.'); return; }
    mutate((b) => { const sheets = { ...b.sheets }; sheets[clean] = sheets[b.active]; delete sheets[b.active]; return { order: b.order.map((n) => (n === b.active ? clean : n)), active: clean, sheets }; });
  };
  const duplicateSheet = () => { const name = uniqueSheetName(book.order, `${book.active} (2)`); mutate((b) => ({ ...b, order: [...b.order, name], active: name, sheets: { ...b.sheets, [name]: JSON.parse(JSON.stringify(b.sheets[b.active])) as Page } })); };
  const deleteSheet = () => { mutate((b) => { const order = b.order.filter((n) => n !== b.active); const sheets = { ...b.sheets }; delete sheets[b.active]; return { order, active: order[0], sheets }; }); select('A1'); };
  const moveSheet = (d: 1 | -1) => mutate((b) => { const i = b.order.indexOf(b.active); const j = i + d; if (j < 0 || j >= b.order.length) return b; const order = [...b.order]; [order[i], order[j]] = [order[j], order[i]]; return { ...b, order }; });

  // ------------------------------------------------------------------ charts
  const addChart = (type: ChartType) => {
    const rect = rangeEnd ? rg : currentRegion(cells, sc, sr);
    if (rect.r1 === rect.r2 && rect.c1 === rect.c2) { flash('Select the data range for the chart first.'); return; }
    const header = hasHeaderRow(cells, rect);
    const firstColText = Array.from({ length: rect.r2 - rect.r1 }, (_, i) => display[refOf(rect.c1, rect.r1 + 1 + i)] ?? '').some((v) => v !== '' && parseNumberish(v) === null);
    const ch: Chart = { id: uid(), type, range: rectToRange(rect), title: header ? display[refOf(rect.c1 + (firstColText ? 1 : 0), rect.r1)] ?? 'Chart' : 'Chart', labelsInFirstCol: firstColText, headerRow: header };
    mutatePage((pg) => ({ ...pg, charts: [...(pg.charts ?? []), ch] }));
    setPanel(null);
    setChartView(ch);
  };
  const updateChart = (ch: Chart) => { mutatePage((pg) => ({ ...pg, charts: (pg.charts ?? []).map((x) => (x.id === ch.id ? ch : x)) })); setChartView(ch); };
  const deleteChart = (id: string) => { mutatePage((pg) => ({ ...pg, charts: (pg.charts ?? []).filter((x) => x.id !== id) })); setChartView(null); setChartEdit(null); };
  const exportChartPng = async (ch: Chart) => {
    const svg = document.querySelector<SVGSVGElement>(`#chart-${ch.id} svg`);
    if (!svg) return;
    try { const png = await svgToPng(svg, 2); flash(await saveBinary(sanitizeName(ch.title || 'chart', 'png'), png, 'image/png')); } catch (e) { flash(`Export failed: ${errMsg(e)}`); }
  };

  // ------------------------------------------------------------------ AI
  const runAi = async (mode = aiMode, prompt = aiPrompt) => {
    const s = getSettings();
    if (!s.apiKey) { flash('Add your API key in Settings first.'); setAiOpen(false); return; }
    const topic = prompt.trim();
    if (!topic && mode !== 'explain' && mode !== 'analyze') return;
    setAiMode(mode);
    setAiBusy(true); setAiOut('');
    try {
      if (mode === 'table') {
        const out = await chatStream(s, [
          { role: 'system', content: 'You generate spreadsheet data. Return ONLY a tab-separated table. First row is a header row. No markdown fences, no commentary, max 10 columns and 40 rows. Use plain numbers (no thousands separators, no currency symbols). Formulas are allowed in Excel syntax starting with = when they add value (totals, averages). Same language as the request.' },
          { role: 'user', content: topic },
        ]);
        const clean = out.replace(/```[a-z]*\n?/gi, '').trim();
        const rows = parseDelimited(clean, '\t');
        const [c0, r0] = parseRef(sel);
        mutatePage((pg) => {
          const c = { ...pg.cells }; const st = { ...pg.styles };
          rows.forEach((row, ri) => row.forEach((v, ci) => { if (v.trim()) c[refOf(c0 + ci, r0 + ri)] = v.trim(); }));
          if (rows[0]) rows[0].forEach((_, ci) => { st[refOf(c0 + ci, r0)] = { ...st[refOf(c0 + ci, r0)], b: true, fill: '#DEEBF7', bb: true }; });
          return { ...pg, cells: c, styles: st };
        });
        setAiOpen(false); flash(`Inserted ${rows.length} rows at ${sel}.`);
      } else if (mode === 'formula') {
        const ctxRect = currentRegion(cells, sc, sr);
        const sample = refsInRect({ ...ctxRect, r2: Math.min(ctxRect.r2, ctxRect.r1 + 5) }).map((r) => `${r}=${JSON.stringify(cells[r] ?? '')}`).join(', ');
        const out = await chatStream(s, [
          { role: 'system', content: `You write spreadsheet formulas. Supported functions: SUM AVERAGE COUNT COUNTA MIN MAX MEDIAN ROUND IF IFS AND OR NOT IFERROR SUMIF SUMIFS COUNTIF COUNTIFS AVERAGEIF VLOOKUP HLOOKUP INDEX MATCH XLOOKUP CONCAT TEXT LEFT RIGHT MID LEN UPPER LOWER TRIM TODAY NOW DATE YEAR MONTH DAY PMT FV PV NPV RANK ABS SQRT POWER MOD. Output ONLY the formula starting with "=" — no explanation, no code fences.` },
          { role: 'user', content: `Selected cell: ${sel}. Nearby data: ${sample || '(empty)'}. Task: ${topic}` },
        ]);
        const f = out.trim().replace(/^```[a-z]*|```$/g, '').trim().split('\n')[0];
        if (!f.startsWith('=')) throw new Error(`Model did not return a formula: ${f.slice(0, 80)}`);
        setEditing(f); setAiOpen(false); flash('Formula ready — press Enter to apply.');
        setTimeout(() => fxRef.current?.focus(), 50);
      } else {
        const rect = rangeEnd ? rg : currentRegion(cells, sc, sr);
        const rows: string[] = [];
        for (let r = rect.r1; r <= Math.min(rect.r2, rect.r1 + 60); r++) rows.push(Array.from({ length: rect.c2 - rect.c1 + 1 }, (_, i) => shown(refOf(rect.c1 + i, r))).join('\t'));
        const content = mode === 'explain'
          ? `Explain what this formula does, step by step, in plain language (max 6 short lines): ${cells[sel] ?? '(empty cell)'}\nReferenced cells: ${refsIn(cells[sel] ?? '').slice(0, 12).map((r) => `${r}=${shown(r)}`).join(', ')}`
          : `Analyze this table (${rectToRange(rect)}). Give 3-6 short bullet insights (trends, outliers, totals) and one suggested next step. Data (TSV):\n${rows.join('\n')}${topic ? `\nFocus: ${topic}` : ''}`;
        let acc = '';
        const out = await chatStream(s, [{ role: 'system', content: 'You are a concise spreadsheet analyst. Plain text, short bullets, no markdown headings.' }, { role: 'user', content }], { onDelta: (d) => { acc += d; setAiOut(acc); } });
        setAiOut(out);
      }
    } catch (e) { setAiOut(`Error: ${errMsg(e)}`); } finally { setAiBusy(false); }
  };

  // ------------------------------------------------------------------ files
  const openFile = async () => {
    const pick = await openFilePicker('.xlsx,.xls,.csv,.tsv,.txt,.ods');
    if (!pick) return;
    try {
      const wb = await importWorkbook(pick.buf, pick.name);
      pushU(book);
      setBook(wb); save(wb, pick.name.replace(/\.[^.]+$/, ''));
      setTitle(pick.name.replace(/\.[^.]+$/, ''));
      // `select` reads the previous workbook's cells from its closure; seed the formula bar from the new one
      setSel('A1'); setRangeEnd(null); setInlineEdit(false);
      setEditing(wb.sheets[wb.active]?.cells.A1 ?? '');
      setTemplatesOpen(false);
      flash(`Opened ${pick.name} (${wb.order.length} sheet${wb.order.length > 1 ? 's' : ''})`);
    } catch (e) { flash(`Could not open: ${errMsg(e)}`); }
  };
  const saveXlsx = async () => {
    try { const bytes = await exportXlsx(title, book); flash(await saveBinary(sanitizeName(title, 'xlsx'), bytes, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')); } catch (e) { flash(`Save failed: ${errMsg(e)}`); }
  };
  const exportCsv = () => {
    const ext = usedExtent(page);
    const rows: string[][] = [];
    for (let r = 0; r < ext.rows; r++) { const row: string[] = []; for (let c = 0; c < ext.cols; c++) row.push(shown(refOf(c, r))); rows.push(row); }
    void downloadText(sanitizeName(`${title}-${book.active}`, 'csv'), toCsv(rows), 'text/csv').then(flash);
  };
  const applyTemplate = (id: string) => {
    const t = SHEET_TEMPLATES.find((x) => x.id === id); if (!t) return;
    const built = t.build();
    mutatePage(() => built);
    if (id !== 'blank' && title.startsWith('Untitled')) setTitle(t.label);
    setTemplatesOpen(false);
    setSel('A1'); setRangeEnd(null); setInlineEdit(false); setEditing(built.cells.A1 ?? '');
  };

  // ------------------------------------------------------------------ status bar stats
  const selStats = useMemo(() => {
    if (!rangeEnd) return null;
    let sum = 0, count = 0, n = 0, min = Infinity, max = -Infinity;
    for (const ref of refsInRect(rg)) { const d = display[ref]; if (d === undefined || d === '') continue; n++; const v = parseNumberish(d); if (v !== null) { count++; sum += v; if (v < min) min = v; if (v > max) max = v; } }
    return { sum, count, n, min, max, avg: count ? sum / count : 0 };
  }, [rg, rangeEnd, display]);

  const startInline = () => { setInlineEdit(true); setTimeout(() => inlineRef.current?.focus(), 30); };

  // ------------------------------------------------------------------ render helpers
  const cellStyleCss = (st: CellStyle | undefined, ref: string): React.CSSProperties => ({
    fontWeight: st?.b ? 700 : undefined,
    fontStyle: st?.i ? 'italic' : undefined,
    textDecoration: st?.u && st?.s ? 'underline line-through' : st?.u ? 'underline' : st?.s ? 'line-through' : undefined,
    color: st?.color,
    background: st?.fill,
    textAlign: st?.align ?? (parseNumberish(display[ref] ?? '') !== null && !(cells[ref] ?? '').startsWith("'") && st?.fmt !== 'text' ? 'right' : undefined),
    alignItems: st?.valign === 'top' ? 'flex-start' : st?.valign === 'bottom' ? 'flex-end' : undefined,
    whiteSpace: st?.wrap ? 'normal' : undefined,
    fontSize: st?.size ? `${st.size * z}px` : undefined,
    borderTop: st?.bt ? '1.5px solid var(--cell-border)' : undefined,
    borderBottom: st?.bb ? '1.5px solid var(--cell-border)' : undefined,
    borderLeft: st?.bl ? '1.5px solid var(--cell-border)' : undefined,
    borderRight: st?.br ? '1.5px solid var(--cell-border)' : undefined,
  });

  const renderCells = (cFrom: number, cTo: number, rFrom: number, rTo: number, offX: number, offY: number) => {
    const out: React.ReactNode[] = [];
    for (let r = rFrom; r <= rTo; r++) {
      if (hiddenRows.has(r)) continue;
      for (let c = cFrom; c <= cTo; c++) {
        const ref = refOf(c, r);
        if (mergeMap.covered.has(ref)) continue;
        const st = styles[ref];
        const merged = mergeMap.anchors.get(ref);
        const w = merged ? colX[Math.min(merged.c2 + 1, nCols)] - colX[c] : colW(c);
        const h = merged ? rowY[Math.min(merged.r2 + 1, nRows)] - rowY[r] : rowH(r);
        const raw = cells[ref] ?? '';
        const txt = showFormulas && raw.startsWith('=') ? raw : shown(ref);
        const isErr = txt.startsWith('#') && /!|\?|N\/A/.test(txt);
        const isHit = hits.includes(ref);
        const prec = showPrecedents.has(ref);
        const hasNote = !!page.notes?.[ref];
        const filterHead = page.filter && (() => { const fr = rangeToRect(page.filter!.range); return fr && r === fr.r1 && c >= fr.c1 && c <= fr.c2; })();
        out.push(
          <div
            key={ref}
            data-ref={ref}
            className={`cell${isErr ? ' err' : ''}${isHit ? ' hit' : ''}${prec ? ' prec' : ''}${raw.startsWith('=') && !showFormulas ? ' formula' : ''}`}
            style={{ left: colX[c] - offX, top: rowY[r] - offY, width: w, height: h, ...cellStyleCss(st, ref) }}
            title={hasNote ? page.notes![ref] : undefined}
          >
            <span className="cell-text">{txt}</span>
            {hasNote && <i className="note-dot" />}
            {filterHead && (
              <button className={`filter-btn${page.filter!.criteria[c]?.length ? ' on' : ''}`} aria-label={`Filter column ${colName(c)}`} onPointerDown={(e) => e.stopPropagation()} onClick={(e) => { e.stopPropagation(); setFilterCol(c); }}>
                <Icon name="filter" size={11} />
              </button>
            )}
          </div>,
        );
      }
    }
    return out;
  };

  const selectionBox = (offX: number, offY: number, clipRect?: Rect) => {
    const a = rg;
    if (clipRect && !(a.c1 <= clipRect.c2 && a.c2 >= clipRect.c1 && a.r1 <= clipRect.r2 && a.r2 >= clipRect.r1)) return null;
    const mergedAnchor = mergeMap.anchors.get(sel);
    const c2 = mergedAnchor && !rangeEnd ? mergedAnchor.c2 : a.c2;
    const r2 = mergedAnchor && !rangeEnd ? mergedAnchor.r2 : a.r2;
    const x = colX[a.c1] - offX, y = rowY[a.r1] - offY;
    const w = colX[Math.min(c2 + 1, nCols)] - colX[a.c1], h = rowY[Math.min(r2 + 1, nRows)] - rowY[a.r1];
    return (
      <>
        <div className="sel-box" style={{ left: x, top: y, width: w, height: h }} />
        <div className="fill-handle" style={{ left: x + w - 6, top: y + h - 6 }} onPointerDown={(e) => { e.stopPropagation(); dragging.current = 'fill'; (gridRef.current as HTMLElement).setPointerCapture?.(e.pointerId); }} />
        {fillPreview && <div className="fill-preview" style={{ left: colX[fillPreview.c1] - offX, top: rowY[fillPreview.r1] - offY, width: colX[fillPreview.c2 + 1] - colX[fillPreview.c1], height: rowY[fillPreview.r2 + 1] - rowY[fillPreview.r1] }} />}
        {clip && !clip.cut && clip.rect && <div className="clip-box" style={{ left: colX[clip.rect.c1] - offX, top: rowY[clip.rect.r1] - offY, width: colX[Math.min(clip.rect.c2 + 1, nCols)] - colX[clip.rect.c1], height: rowY[Math.min(clip.rect.r2 + 1, nRows)] - rowY[clip.rect.r1] }} />}
      </>
    );
  };

  const hasUndo = past.current.length > 0;
  const hasRedo = future.current.length > 0;
  const chartsHere = page.charts ?? [];
  const inlineCellRect = inlineEdit ? { x: colX[sc], y: rowY[sr], w: mergeMap.anchors.get(sel) ? colX[mergeMap.anchors.get(sel)!.c2 + 1] - colX[sc] : colW(sc), h: rowH(sr) } : null;

  return (
    <div className={`edscreen${showGrid ? '' : ' nogrid'}`} style={{ ['--app' as string]: 'var(--excel)' }}>
      <AppBar kindIcon={<FileTypeIcon kind="sheet" size={24} light />} title={title} onTitle={setTitle} placeholder="Workbook title" onBack={onExit} saved={saved}>
        <button className="icon-btn light" aria-label="Undo" disabled={!hasUndo} onClick={undo}><Icon name="undo" size={20} /></button>
        <button className="icon-btn light" aria-label="Save as .xlsx" onClick={() => void saveXlsx()}><Icon name="save" size={20} /></button>
        <button className="icon-btn light" aria-label="More actions" onClick={() => setMenu(true)}><Icon name="more" size={20} /></button>
      </AppBar>

      {/* formula bar */}
      <div className="formula-bar">
        <button className="namebox" onClick={() => setGoOpen(true)} title="Go to cell / range">{rangeEnd ? rectToRange(rg) : sel}</button>
        <button className="fx" aria-label="Insert function" onClick={() => setFnOpen(true)}>fx</button>
        <input
          ref={fxRef}
          className="fx-input"
          value={editing ?? ''}
          placeholder="Value or formula, e.g. =SUM(A1:A9)"
          inputMode={(editing ?? '').startsWith('=') ? 'text' : undefined}
          onChange={(e) => setEditing(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === 'Enter') { e.preventDefault(); commitAndMove(0, e.shiftKey ? -1 : 1, e.currentTarget.value); }
            else if (e.key === 'Tab') { e.preventDefault(); commitAndMove(e.shiftKey ? -1 : 1, 0, e.currentTarget.value); }
            else if (e.key === 'Escape') { setEditing(cells[sel] ?? ''); e.currentTarget.blur(); }
            else if (e.key === 'F4') { e.preventDefault(); toggleAbs(); }
          }}
          onBlur={() => { if (editing !== null && !inlineEdit) commit(sel, editing); }}
        />
        {(editing ?? '').startsWith('=') ? (
          <button className="icon-btn" aria-label="Toggle absolute reference" title="$ toggle (F4)" onPointerDown={(e) => e.preventDefault()} onClick={toggleAbs}><b>$</b></button>
        ) : null}
        <button className="icon-btn accent" aria-label="Confirm" onPointerDown={(e) => e.preventDefault()} onClick={() => commitAndMove(0, 1)}><Icon name="check" size={20} /></button>
      </div>

      {findOpen && (
        <div className="findbar">
          <div className="find-row">
            <Icon name="search" size={16} className="dim" />
            <input className="input find-input" value={fq} placeholder="Find in sheet" onChange={(e) => setFq(e.target.value)} onKeyDown={(e) => e.key === 'Enter' && runFind()} />
            <span className="find-count">{hits.length ? `${hitIdx + 1}/${hits.length}` : ''}</span>
            <button className="icon-btn" disabled={!hits.length} onClick={() => nextHit(-1)} aria-label="Previous"><Icon name="chevronUp" size={18} /></button>
            <button className="icon-btn" disabled={!hits.length} onClick={() => nextHit(1)} aria-label="Next"><Icon name="chevronDown" size={18} /></button>
            <button className="icon-btn" onClick={() => { setFindOpen(false); setHits([]); }} aria-label="Close"><Icon name="close" size={18} /></button>
          </div>
          <div className="find-row">
            <Icon name="replace" size={16} className="dim" />
            <input className="input find-input" value={fr} placeholder="Replace with" onChange={(e) => setFr(e.target.value)} />
            <button className="btn small" onClick={runFind}>Find</button>
            <button className="btn small" disabled={!hits.length} onClick={replaceOne}>Replace</button>
            <button className="btn small" disabled={!hits.length} onClick={replaceAll}>All</button>
          </div>
        </div>
      )}

      {/* grid */}
      <div
        ref={gridRef}
        className={`sheet-grid${dragging.current ? ' dragging' : ''}`}
        style={{ ['--head-h' as string]: `${headH}px`, ['--rowhead-w' as string]: `${ROW_HEAD_W}px`, fontSize: `${13 * z}px` }}
        onPointerDown={onGridPointerDown}
        onPointerMove={onGridPointerMove}
        onPointerUp={onGridPointerUp}
        onPointerCancel={() => { press.current = null; dragging.current = null; setFillPreview(null); }}
        onDoubleClick={(e) => { if (cellAt(e.clientX, e.clientY) === sel) startInline(); }}
      >
        <div className="grid-inner" style={{ width: totalW + ROW_HEAD_W, height: totalH + headH }}>
        {/* header strips ride in a 0×0 sticky anchor pinned to the viewport's top-left corner */}
        <div className="grid-sticky">
        <div className="col-heads" style={{ height: headH, width: viewport.w }}>
          <div className="cell head corner" style={{ width: ROW_HEAD_W, height: headH, left: 0 }} />
          {Array.from({ length: lastCol - firstCol + 1 }, (_, i) => firstCol + i).concat(Array.from({ length: freeze.c }, (_, i) => i)).map((c) => (
            <div key={c} data-col={c} className={`cell head colhead${c >= rg.c1 && c <= rg.c2 ? ' hitsel' : ''}${c < freeze.c ? ' frozen' : ''}`} style={{ left: (c < freeze.c ? colX[c] : colX[c] - viewport.x) + ROW_HEAD_W, width: colW(c), height: headH }}>
              {colName(c)}
              <span className="col-resize" onPointerDown={(e) => { e.stopPropagation(); colResize.current = { c, startX: e.clientX, startW: page.colW?.[c] ?? DEFAULT_COL_W }; (gridRef.current as HTMLElement).setPointerCapture?.(e.pointerId); }} onDoubleClick={(e) => { e.stopPropagation(); select(refOf(c, sr)); setTimeout(autoFitCol, 0); }} />
            </div>
          ))}
        </div>

        {/* row headers */}
        <div className="row-heads" style={{ width: ROW_HEAD_W, height: viewport.h }}>
          {Array.from({ length: lastRow - firstRow + 1 }, (_, i) => firstRow + i).concat(Array.from({ length: freeze.r }, (_, i) => i)).map((r) =>
            hiddenRows.has(r) ? null : (
              <div key={r} data-row={r} className={`cell head rowhead${r >= rg.r1 && r <= rg.r2 ? ' hitsel' : ''}${r < freeze.r ? ' frozen' : ''}`} style={{ top: (r < freeze.r ? rowY[r] : rowY[r] - viewport.y) + headH, height: rowH(r), width: ROW_HEAD_W }}>
                {r + 1}
                <span className="row-resize" onPointerDown={(e) => { e.stopPropagation(); rowResize.current = { r, startY: e.clientY, startH: page.rowH?.[r] ?? DEFAULT_ROW_H }; (gridRef.current as HTMLElement).setPointerCapture?.(e.pointerId); }} />
              </div>
            ),
          )}
        </div>
        </div>

        {/* scrolling cells */}
        <div className="cells-layer" style={{ left: ROW_HEAD_W, top: headH, width: totalW, height: totalH }}>
          {renderCells(firstCol, lastCol, firstRow, lastRow, 0, 0)}
          {selectionBox(0, 0)}
          {inlineEdit && inlineCellRect && (
            <input
              ref={inlineRef}
              className="inline-edit"
              style={{ left: inlineCellRect.x, top: inlineCellRect.y, width: Math.max(inlineCellRect.w, 120), height: inlineCellRect.h, textAlign: styles[sel]?.align }}
              value={editing ?? ''}
              onChange={(e) => setEditing(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === 'Enter') { e.preventDefault(); commitAndMove(0, e.shiftKey ? -1 : 1, e.currentTarget.value); }
                else if (e.key === 'Tab') { e.preventDefault(); commitAndMove(e.shiftKey ? -1 : 1, 0, e.currentTarget.value); }
                else if (e.key === 'Escape') { setEditing(cells[sel] ?? ''); setInlineEdit(false); }
              }}
              onBlur={() => { if (editing !== null) commit(sel, editing); setInlineEdit(false); }}
            />
          )}
        </div>

        {/* frozen panes */}
        {freeze.c > 0 && (
          <div className="cells-layer frozen-cols" style={{ left: ROW_HEAD_W, top: headH, width: frozenW, height: totalH, transform: `translateX(${viewport.x}px)` }}>
            {renderCells(0, freeze.c - 1, firstRow, lastRow, 0, 0)}
            {selectionBox(0, 0, { c1: 0, c2: freeze.c - 1, r1: 0, r2: nRows })}
          </div>
        )}
        {freeze.r > 0 && (
          <div className="cells-layer frozen-rows" style={{ left: ROW_HEAD_W, top: headH, width: totalW, height: frozenH, transform: `translateY(${viewport.y}px)` }}>
            {renderCells(firstCol, lastCol, 0, freeze.r - 1, 0, 0)}
            {selectionBox(0, 0, { c1: 0, c2: nCols, r1: 0, r2: freeze.r - 1 })}
          </div>
        )}
        {freeze.c > 0 && freeze.r > 0 && (
          <div className="cells-layer frozen-corner" style={{ left: ROW_HEAD_W, top: headH, width: frozenW, height: frozenH, transform: `translate(${viewport.x}px, ${viewport.y}px)` }}>
            {renderCells(0, freeze.c - 1, 0, freeze.r - 1, 0, 0)}
            {selectionBox(0, 0, { c1: 0, c2: freeze.c - 1, r1: 0, r2: freeze.r - 1 })}
          </div>
        )}

        {/* charts float bottom-right of their data range */}
        {chartsHere.map((ch, i) => {
          const rect = rangeToRect(ch.range);
          const x = rect ? colX[Math.min(rect.c2 + 1, nCols)] + 12 : 20;
          const y = rect ? rowY[rect.r1] : 20 + i * 30;
          return (
            <div key={ch.id} id={`chart-${ch.id}`} className="chart-card" style={{ left: ROW_HEAD_W + Math.min(x, Math.max(0, totalW - 280)), top: headH + y }} onClick={() => setChartView(ch)}>
              <ChartView chart={ch} data={chartData(ch, cells, display)} width={260} height={170} />
            </div>
          );
        })}
        </div>
      </div>

      {/* sheet tabs */}
      <div className="sheet-tabs">
        <button className="stab-add" aria-label="Sheet options" onClick={() => setTabMenu(true)}><Icon name="more" size={18} /></button>
        <div className="stab-scroll">
          {book.order.map((name) => (
            <button key={name} className={`stab${name === book.active ? ' active' : ''}`} onClick={() => switchSheet(name)} onDoubleClick={() => setRenameOpen(true)}>{name}</button>
          ))}
        </div>
        <button className="stab-add" aria-label="Add sheet" onClick={addSheet}><Icon name="plus" size={18} /></button>
      </div>

      {/* status bar */}
      <div className="sheet-status">
        {selStats && selStats.count > 0 ? (
          <button className="status-btn" onClick={() => setStatsOpen(true)}>
            Sum {fmtDisplay(String(selStats.sum), { fmt: 'gen' })} · Avg {fmtDisplay(String(selStats.avg), { fmt: 'num' })} · Count {selStats.n}
          </button>
        ) : selStats ? <span>{selStats.n} cell(s) selected</span> : <span>{rangeEnd ? rectToRange(rg) : sel}{cells[sel]?.startsWith('=') ? ` = ${display[sel]}` : ''}</span>}
        <span className="status-right">
          {page.filter && hiddenRows.size > 0 && <span>{hiddenRows.size} hidden</span>}
          <button className="status-btn" onClick={() => setZoom((v) => (v >= 150 ? 75 : v + 25))}>{zoom}%</button>
        </span>
      </div>

      {(aiBusy && aiMode !== 'table' && aiMode !== 'formula') || aiOut ? (
        <div className="ai-sheet">
          <div className="ai-sheet-head"><Icon name="ai" size={14} /> {aiMode === 'explain' ? 'Explain formula' : 'Analyze data'}{aiBusy && <span className="ai-dot" />}
            <button className="icon-btn" aria-label="Dismiss" onClick={() => { setAiOut(''); }}><Icon name="close" size={16} /></button>
          </div>
          <div className="ai-sheet-body">{aiOut || 'Thinking…'}</div>
          {!aiBusy && aiOut && (
            <div className="btn-row">
              <button className="btn small" onClick={() => void navigator.clipboard?.writeText(aiOut).then(() => flash('Copied.'))}>Copy</button>
              <button className="btn small" onClick={() => { const ext = usedExtent(page); const r0 = ext.rows + 1; mutatePage((pg) => { const c = { ...pg.cells }; aiOut.split('\n').filter(Boolean).forEach((l, i) => { c[refOf(0, r0 + i)] = l.replace(/^[-•*]\s*/, ''); }); return { ...pg, cells: c }; }); setAiOut(''); flash('Insights added below your data.'); }}>Insert below data</button>
            </div>
          )}
        </div>
      ) : null}

      {/* ribbon */}
      <div className="ribbon">
        {panel === 'text' && <Palette current={curStyle?.color} onPick={(c) => { applyStyle({ color: c }); setPanel(null); }} auto={() => { applyStyle({ color: undefined }); setPanel(null); }} autoLabel="Automatic" />}
        {panel === 'fill' && <Palette current={curStyle?.fill} onPick={(c) => { applyStyle({ fill: c }); setPanel(null); }} auto={() => { applyStyle({ fill: undefined }); setPanel(null); }} autoLabel="No fill" />}
        {panel === 'borders' && (
          <RibbonPanel title="Borders" onClose={() => setPanel(null)}>
            <div className="rpanel-actions">
              {(['all', 'outline', 'top', 'bottom', 'left', 'right', 'none'] as const).map((k) => (
                <button key={k} className="btn small" onClick={() => { setBorders(k); setPanel(null); }}>{k === 'none' ? 'No border' : k[0].toUpperCase() + k.slice(1)}</button>
              ))}
            </div>
          </RibbonPanel>
        )}
        {panel === 'chart' && (
          <RibbonPanel title="Insert chart (uses the selected range or the table around the cursor)" onClose={() => setPanel(null)}>
            <div className="rpanel-actions">
              {([['column', 'chart', 'Column'], ['bar', 'chart', 'Bar'], ['line', 'chartLine', 'Line'], ['area', 'chartLine', 'Area'], ['pie', 'chartPie', 'Pie']] as const).map(([t, ic, l]) => (
                <button key={t} className="btn small" onClick={() => addChart(t)}><Icon name={ic} size={16} /> {l}</button>
              ))}
            </div>
          </RibbonPanel>
        )}
        {panel === 'numfmt' && (
          <RibbonPanel title="Number format" onClose={() => setPanel(null)}>
            <div className="rpanel-actions">
              {FMT_OPTIONS.map((o) => (
                <button key={o.v} className={`btn small${(curStyle?.fmt ?? 'gen') === o.v ? ' primary' : ''}`} onClick={() => { applyStyle({ fmt: o.v === 'gen' ? undefined : o.v }); setPanel(null); }}>
                  {o.t}
                  <small className="dim"> {o.v === 'cur' ? `${currency}1,234.50` : o.v === 'pct' ? '12.5%' : o.v === 'num' ? '1,234.50' : o.v === 'date' ? new Date().toLocaleDateString() : o.v === 'time' ? '09:30' : o.v === 'sci' ? '1.23E+3' : ''}</small>
                </button>
              ))}
            </div>
          </RibbonPanel>
        )}

        <RibbonTabs tabs={[{ id: 'home', label: 'Home' }, { id: 'insert', label: 'Insert' }, { id: 'formulas', label: 'Formulas' }, { id: 'data', label: 'Data' }, { id: 'view', label: 'View' }]} value={rTab} onChange={(t) => { setRTab(t); setPanel(null); void tap(); }} />

        {rTab === 'home' && (
          <div className="ribbon-row">
            <RGroup label="Clipboard">
              <RBtn icon="undo" label="Undo" disabled={!hasUndo} onRun={undo} />
              <RBtn icon="redo" label="Redo" disabled={!hasRedo} onRun={redo} />
              <RBtn icon="cut" label="Cut" onRun={() => copySel(true)} />
              <RBtn icon="copy" label="Copy" onRun={() => copySel(false)} />
              <RBtn icon="paste" label="Paste" onRun={() => void pasteClip('all')} />
            </RGroup>
            <RGroup label="Font">
              <RStepper value={String(curStyle?.size ?? 13)} title="Font size" onDec={() => applyStyle({ size: Math.max(8, (curStyle?.size ?? 13) - 1) })} onInc={() => applyStyle({ size: Math.min(36, (curStyle?.size ?? 13) + 1) })} />
              <RBtn icon="bold" label="Bold" active={!!curStyle?.b} onRun={() => toggle('b')} />
              <RBtn icon="italic" label="Italic" active={!!curStyle?.i} onRun={() => toggle('i')} />
              <RBtn icon="underline" label="Underline" active={!!curStyle?.u} onRun={() => toggle('u')} />
              <RBtn icon="strike" label="Strike" active={!!curStyle?.s} onRun={() => toggle('s')} />
              <RBtn icon="fontColor" label="Color" colorBar={curStyle?.color ?? '#C00000'} menu active={panel === 'text'} onRun={() => setPanel(panel === 'text' ? null : 'text')} />
              <RBtn icon="fill" label="Fill" colorBar={curStyle?.fill ?? '#FFFF00'} menu active={panel === 'fill'} onRun={() => setPanel(panel === 'fill' ? null : 'fill')} />
              <RBtn icon="border" label="Borders" menu active={panel === 'borders'} onRun={() => setPanel(panel === 'borders' ? null : 'borders')} />
            </RGroup>
            <RGroup label="Alignment">
              <RBtn icon="alignLeft" label="Left" active={curStyle?.align === 'left'} onRun={() => applyStyle({ align: curStyle?.align === 'left' ? undefined : 'left' })} />
              <RBtn icon="alignCenter" label="Center" active={curStyle?.align === 'center'} onRun={() => applyStyle({ align: curStyle?.align === 'center' ? undefined : 'center' })} />
              <RBtn icon="alignRight" label="Right" active={curStyle?.align === 'right'} onRun={() => applyStyle({ align: curStyle?.align === 'right' ? undefined : 'right' })} />
              <RBtn icon="alignTop" label="Top" active={curStyle?.valign === 'top'} onRun={() => applyStyle({ valign: curStyle?.valign === 'top' ? undefined : 'top' })} />
              <RBtn icon="alignBottom" label="Bottom" active={curStyle?.valign === 'bottom'} onRun={() => applyStyle({ valign: curStyle?.valign === 'bottom' ? undefined : 'bottom' })} />
              <RBtn icon="wrap" label="Wrap" active={!!curStyle?.wrap} onRun={() => toggle('wrap')} />
              <RBtn icon="merge" label={isMerged ? 'Unmerge' : 'Merge'} active={isMerged} onRun={() => (isMerged ? unmergeCells() : mergeCells())} />
            </RGroup>
            <RGroup label="Number">
              <RSelect value={curStyle?.fmt ?? 'gen'} options={FMT_OPTIONS} onChange={(v) => applyStyle({ fmt: v === 'gen' ? undefined : (v as CellStyle['fmt']) })} width={96} title="Number format" />
              <RBtn icon="currency" label="Currency" active={curStyle?.fmt === 'cur'} onRun={() => applyStyle({ fmt: curStyle?.fmt === 'cur' ? undefined : 'cur' })} />
              <RBtn icon="percent" label="Percent" active={curStyle?.fmt === 'pct'} onRun={() => applyStyle({ fmt: curStyle?.fmt === 'pct' ? undefined : 'pct' })} />
              <RBtn icon="decimalMore" label="+.0" onRun={() => changeDec(1)} />
              <RBtn icon="decimalLess" label="−.0" onRun={() => changeDec(-1)} />
            </RGroup>
            <RGroup label="Cells">
              <RBtn icon="rowBelow" label="Insert" onRun={() => doInsertRows(sr)} />
              <RBtn icon="rowDelete" label="Delete" onRun={doDeleteRows} />
              <RBtn icon="clearFormat" label="Clear" menu onRun={clearAll} />
              <RBtn icon="fillDown" label="Fill" onRun={fillDownQuick} />
            </RGroup>
            <RGroup label="Editing">
              <RBtn icon="sigma" label="AutoSum" onRun={() => insertFn('SUM')} />
              <RBtn icon="sortAZ" label="Sort" onRun={() => setSortOpen(true)} />
              <RBtn icon="filter" label="Filter" active={!!page.filter} onRun={toggleFilter} />
              <RBtn icon="search" label="Find" active={findOpen} onRun={() => setFindOpen(!findOpen)} />
            </RGroup>
          </div>
        )}

        {rTab === 'insert' && (
          <div className="ribbon-row">
            <RGroup label="Rows">
              <RBtn icon="rowAbove" label="Above" onRun={() => doInsertRows(rg.r1)} />
              <RBtn icon="rowBelow" label="Below" onRun={() => doInsertRows(rg.r2 + 1)} />
              <RBtn icon="rowDelete" label={rowsInSel > 1 ? `Delete ${rowsInSel}` : 'Delete row'} onRun={doDeleteRows} />
            </RGroup>
            <RGroup label="Columns">
              <RBtn icon="colLeft" label="Left" onRun={() => doInsertCols(rg.c1)} />
              <RBtn icon="colRight" label="Right" onRun={() => doInsertCols(rg.c2 + 1)} />
              <RBtn icon="colDelete" label={colsInSel > 1 ? `Delete ${colsInSel}` : 'Delete col'} onRun={doDeleteCols} />
              <RBtn icon="fitWidth" label="AutoFit" onRun={autoFitCol} />
            </RGroup>
            <RGroup label="Charts">
              <RBtn icon="chart" label="Column" onRun={() => addChart('column')} />
              <RBtn icon="chartLine" label="Line" onRun={() => addChart('line')} />
              <RBtn icon="chartPie" label="Pie" onRun={() => addChart('pie')} />
              <RBtn icon="chart" label="More" menu active={panel === 'chart'} onRun={() => setPanel(panel === 'chart' ? null : 'chart')} />
            </RGroup>
            <RGroup label="Insert">
              <RBtn icon="fx" label="Function" onRun={() => setFnOpen(true)} />
              <RBtn icon="comment" label="Note" active={!!page.notes?.[sel]} onRun={() => setNoteOpen(true)} />
              <RBtn icon="calendar" label="Date" onRun={() => commit(sel, new Date().toLocaleDateString())} />
              <RBtn icon="check" label="✓" onRun={() => commit(sel, '✓')} />
              <RBtn icon="table" label="Template" onRun={() => setTemplatesOpen(true)} />
            </RGroup>
          </div>
        )}

        {rTab === 'formulas' && (
          <div className="ribbon-row">
            <RGroup label="Library">
              <RBtn icon="fx" label="Insert" onRun={() => setFnOpen(true)} />
              <RBtn icon="sigma" label="AutoSum" onRun={() => insertFn('SUM')} />
              <RBtn icon="fx" label="Average" onRun={() => insertFn('AVERAGE')} />
              <RBtn icon="fx" label="Count" onRun={() => insertFn('COUNT')} />
              <RBtn icon="fx" label="Max" onRun={() => insertFn('MAX')} />
              <RBtn icon="fx" label="Min" onRun={() => insertFn('MIN')} />
            </RGroup>
            <RGroup label="Common">
              <RBtn icon="fx" label="IF" onRun={() => insertFn('IF', 'IF(A1>0, "yes", "no")')} />
              <RBtn icon="fx" label="SUMIF" onRun={() => insertFn('SUMIF', 'SUMIF(A1:A9, ">0")')} />
              <RBtn icon="fx" label="VLOOKUP" onRun={() => insertFn('VLOOKUP', 'VLOOKUP(A1, A1:D9, 2, FALSE)')} />
              <RBtn icon="fx" label="TODAY" onRun={() => insertFn('TODAY', 'TODAY()')} />
              <RBtn icon="fx" label="PMT" onRun={() => insertFn('PMT', 'PMT(rate/12, months, -principal)')} />
              <RBtn icon="fx" label="CONCAT" onRun={() => insertFn('CONCAT', 'CONCAT(A1, " ", B1)')} />
            </RGroup>
            <RGroup label="Auditing">
              <RBtn icon="eye" label={showFormulas ? 'Values' : 'Formulas'} active={showFormulas} onRun={() => setShowFormulas(!showFormulas)} />
              <RBtn icon="ai" label="Explain" disabled={!cells[sel]?.startsWith('=')} onRun={() => void runAi('explain', '')} />
              <RBtn icon="sparkle" label="AI formula" onRun={() => { setAiMode('formula'); setAiPrompt(''); setAiOpen(true); }} />
            </RGroup>
          </div>
        )}

        {rTab === 'data' && (
          <div className="ribbon-row">
            <RGroup label="Sort & filter">
              <RBtn icon="sortAZ" label="A→Z" onRun={() => sortBy(sc, 1)} />
              <RBtn icon="sortZA" label="Z→A" onRun={() => sortBy(sc, -1)} />
              <RBtn icon="sortNum" label="Custom" onRun={() => setSortOpen(true)} />
              <RBtn icon="filter" label={page.filter ? 'Clear filter' : 'Filter'} active={!!page.filter} onRun={toggleFilter} />
            </RGroup>
            <RGroup label="Data tools">
              <RBtn icon="duplicate" label="Dedupe" onRun={removeDuplicates} />
              <RBtn icon="columns" label="Split text" onRun={textToColumns} />
              <RBtn icon="fillDown" label="Fill down" onRun={fillDownQuick} />
              <RBtn icon="fillRight" label="Fill right" onRun={() => (rangeEnd ? fill('right') : flash('Select the range to fill first.'))} />
              <RBtn icon="selectRange" label="Select table" onRun={() => { const r = currentRegion(cells, sc, sr); setSel(refOf(r.c1, r.r1)); setRangeEnd(refOf(r.c2, r.r2)); }} />
            </RGroup>
            <RGroup label="Get data">
              <RBtn icon="fileOpen" label="Import" onRun={() => void openFile()} />
              <RBtn icon="paste" label="Paste CSV" onRun={() => void pasteClip()} />
              <RBtn icon="ai" label="AI table" onRun={() => { setAiMode('table'); setAiOpen(true); }} />
              <RBtn icon="sparkle" label="Analyze" onRun={() => { setAiMode('analyze'); setAiPrompt(''); setAiOpen(true); }} />
            </RGroup>
          </div>
        )}

        {rTab === 'view' && (
          <div className="ribbon-row">
            <RGroup label="Freeze">
              <RSeg value={!page.freeze ? 'none' : page.freeze.r === 1 && page.freeze.c === 0 ? 'top' : page.freeze.r === 0 && page.freeze.c === 1 ? 'first' : 'here'} options={[{ v: 'none', t: 'Off' }, { v: 'top', t: 'Top row' }, { v: 'first', t: 'First col' }, { v: 'here', t: 'At cell' }]} onChange={(v) => setFreeze(v)} />
            </RGroup>
            <RGroup label="Zoom">
              <RStepper value={`${zoom}%`} title="Zoom" width={44} onDec={() => setZoom((v) => Math.max(50, v - 10))} onInc={() => setZoom((v) => Math.min(200, v + 10))} />
              <RBtn icon="fitWidth" label="100%" onRun={() => setZoom(100)} />
            </RGroup>
            <RGroup label="Show">
              <RBtn icon="grid" label="Gridlines" active={showGrid} onRun={() => setShowGrid(!showGrid)} />
              <RBtn icon="eye" label="Formulas" active={showFormulas} onRun={() => setShowFormulas(!showFormulas)} />
              <RBtn icon="wordCount" label="Stats" onRun={() => setStatsOpen(true)} />
            </RGroup>
            <RGroup label="Sheet">
              <RBtn icon="rename" label="Rename" onRun={() => setRenameOpen(true)} />
              <RBtn icon="duplicate" label="Duplicate" onRun={duplicateSheet} />
              <RBtn icon="arrowLeft" label="Move left" onRun={() => moveSheet(-1)} />
              <RBtn icon="arrowRight" label="Move right" onRun={() => moveSheet(1)} />
            </RGroup>
          </div>
        )}
      </div>

      {/* ---------------- sheets & dialogs ---------------- */}
      <BottomSheet open={menu} onClose={() => setMenu(false)} title={title}>
        <SheetMenu onClose={() => setMenu(false)} items={[
          { icon: 'fileOpen', label: 'Open file', hint: '.xlsx · .xls · .csv · .ods', onRun: () => void openFile() },
          { icon: 'template', label: 'Templates', hint: 'Budget, invoice, tracker, loan…', onRun: () => setTemplatesOpen(true) },
          'divider',
          { icon: 'save', label: 'Save as Excel (.xlsx)', hint: 'All sheets, formulas, formats', onRun: () => void saveXlsx() },
          { icon: 'fileText', label: `Export "${book.active}" as CSV`, onRun: exportCsv },
          { icon: 'share', label: 'Share selection as text', onRun: () => { copySel(false); flash('Copied as tab-separated text — paste anywhere.'); } },
          'divider',
          { icon: 'search', label: 'Find & replace', onRun: () => setFindOpen(true) },
          { icon: 'wordCount', label: 'Workbook statistics', onRun: () => setStatsOpen(true) },
        ]} />
      </BottomSheet>

      <BottomSheet open={tabMenu} onClose={() => setTabMenu(false)} title={`Sheet: ${book.active}`}>
        <SheetMenu onClose={() => setTabMenu(false)} items={[
          { icon: 'rename', label: 'Rename', onRun: () => setRenameOpen(true) },
          { icon: 'duplicate', label: 'Duplicate', onRun: duplicateSheet },
          { icon: 'plus', label: 'New sheet', onRun: addSheet },
          { icon: 'arrowLeft', label: 'Move left', disabled: book.order.indexOf(book.active) === 0, onRun: () => moveSheet(-1) },
          { icon: 'arrowRight', label: 'Move right', disabled: book.order.indexOf(book.active) === book.order.length - 1, onRun: () => moveSheet(1) },
          'divider',
          { icon: 'trash', label: 'Delete sheet', danger: true, disabled: book.order.length <= 1, onRun: () => setConfirmDelSheet(true) },
        ]} />
      </BottomSheet>

      <BottomSheet open={templatesOpen} onClose={() => setTemplatesOpen(false)} title="New workbook">
        <div className="template-grid">
          {SHEET_TEMPLATES.map((t) => (
            <button key={t.id} className="template-card" onClick={() => applyTemplate(t.id)}>
              <span className="template-thumb sheet" data-kind={t.id}><i /><i /><i /></span>
              <strong>{t.label}</strong><small>{t.desc}</small>
            </button>
          ))}
          <button className="template-card ai" onClick={() => { setTemplatesOpen(false); setAiMode('table'); setAiOpen(true); }}>
            <span className="template-thumb ai"><Icon name="ai" size={22} /></span>
            <strong>AI table</strong><small>Describe the data you need</small>
          </button>
        </div>
      </BottomSheet>

      <BottomSheet open={sortOpen} onClose={() => setSortOpen(false)} title="Sort">
        <p className="hint">Sorts the table around the cursor (or the selected range). Header rows are detected automatically.</p>
        <div className="sort-grid">
          {Array.from({ length: dataRect().c2 - dataRect().c1 + 1 }, (_, i) => dataRect().c1 + i).slice(0, 12).map((c) => (
            <div key={c} className="sort-row">
              <span className="sort-col">{colName(c)}{hasHeaderRow(cells, dataRect()) ? ` · ${display[refOf(c, dataRect().r1)] ?? cells[refOf(c, dataRect().r1)] ?? ''}` : ''}</span>
              <button className="btn small" onClick={() => { sortBy(c, 1); setSortOpen(false); }}><Icon name="sortAZ" size={14} /> A→Z</button>
              <button className="btn small" onClick={() => { sortBy(c, -1); setSortOpen(false); }}><Icon name="sortZA" size={14} /> Z→A</button>
            </div>
          ))}
        </div>
      </BottomSheet>

      <BottomSheet open={filterCol !== null} onClose={() => setFilterCol(null)} title={filterCol !== null ? `Filter column ${colName(filterCol)}` : ''} tall>
        {filterCol !== null && (
          <FilterList values={filterValues} selected={page.filter?.criteria[filterCol] ?? null} onApply={(sel2) => { setCriteria(filterCol, sel2); setFilterCol(null); }} onSort={(d) => { const fr = rangeToRect(page.filter!.range)!; sortBy(filterCol, d, fr); setFilterCol(null); }} />
        )}
      </BottomSheet>

      <BottomSheet open={!!chartView} onClose={() => setChartView(null)} title={chartView?.title || 'Chart'} tall>
        {chartView && (() => { const ch = chartsHere.find((x) => x.id === chartView.id) ?? chartView; return (
          <div className="chart-editor">
            <div id={`chart-${ch.id}-big`} className="chart-big"><ChartView chart={ch} data={chartData(ch, cells, display)} width={Math.min(window.innerWidth - 48, 480)} height={240} /></div>
            <label className="field"><span>Title</span><input className="input" value={ch.title} onChange={(e) => updateChart({ ...ch, title: e.target.value })} /></label>
            <label className="field"><span>Data range</span><input className="input" value={ch.range} onChange={(e) => rangeToRect(e.target.value) && updateChart({ ...ch, range: e.target.value.toUpperCase() })} /></label>
            <RSeg value={ch.type} options={[{ v: 'column', t: 'Column' }, { v: 'bar', t: 'Bar' }, { v: 'line', t: 'Line' }, { v: 'area', t: 'Area' }, { v: 'pie', t: 'Pie' }]} onChange={(v) => updateChart({ ...ch, type: v })} />
            <div className="btn-row">
              <button className={`btn small${ch.headerRow ? ' primary' : ''}`} onClick={() => updateChart({ ...ch, headerRow: !ch.headerRow })}>First row = names</button>
              <button className={`btn small${ch.labelsInFirstCol ? ' primary' : ''}`} onClick={() => updateChart({ ...ch, labelsInFirstCol: !ch.labelsInFirstCol })}>First column = labels</button>
            </div>
            <div className="btn-row end">
              <button className="btn small danger" onClick={() => deleteChart(ch.id)}>Delete</button>
              <button className="btn small" onClick={() => void exportChartPng(ch)}>Save PNG</button>
              <button className="btn small primary" onClick={() => setChartView(null)}>Done</button>
            </div>
          </div>
        ); })()}
      </BottomSheet>

      <BottomSheet open={statsOpen} onClose={() => setStatsOpen(false)} title="Statistics">
        <div className="stats-grid">
          {selStats ? (<>
            <div><b>{fmtDisplay(String(selStats.sum), { fmt: 'gen' })}</b><span>Sum</span></div>
            <div><b>{fmtDisplay(String(selStats.avg), { fmt: 'num' })}</b><span>Average</span></div>
            <div><b>{selStats.count}</b><span>Numbers</span></div>
            <div><b>{selStats.n}</b><span>Non-empty</span></div>
            <div><b>{selStats.count ? fmtDisplay(String(selStats.min), { fmt: 'gen' }) : '–'}</b><span>Min</span></div>
            <div><b>{selStats.count ? fmtDisplay(String(selStats.max), { fmt: 'gen' }) : '–'}</b><span>Max</span></div>
          </>) : (<>
            <div><b>{book.order.length}</b><span>Sheets</span></div>
            <div><b>{Object.keys(cells).length}</b><span>Cells used</span></div>
            <div><b>{Object.values(cells).filter((v) => v.startsWith('=')).length}</b><span>Formulas</span></div>
            <div><b>{extent.cols} × {extent.rows}</b><span>Used range</span></div>
            <div><b>{chartsHere.length}</b><span>Charts</span></div>
            <div><b>{(page.merges ?? []).length}</b><span>Merged ranges</span></div>
          </>)}
        </div>
        {!selStats && <p className="hint">Select a range (long-press and drag) to see Sum, Average, Min and Max here and in the status bar.</p>}
      </BottomSheet>

      <BottomSheet open={aiOpen} onClose={() => !aiBusy && setAiOpen(false)} title={aiMode === 'table' ? 'AI: generate a table' : aiMode === 'formula' ? 'AI: write a formula' : 'AI: analyze'}>
        <p className="hint">{aiMode === 'table' ? `The table is inserted at ${sel} with a header row.` : aiMode === 'formula' ? `Describe the calculation for ${sel}. Nearby data is sent as context.` : 'Optional: what should the analysis focus on?'}</p>
        <textarea className="input" rows={3} value={aiPrompt} placeholder={aiMode === 'table' ? 'e.g. 12-month savings plan with monthly deposit, interest and running balance' : aiMode === 'formula' ? 'e.g. total of column B where column A says "Food"' : 'e.g. which month had the biggest drop?'} onChange={(e) => setAiPrompt(e.target.value)} />
        <div className="btn-row end">
          <button className="btn" onClick={() => setAiOpen(false)} disabled={aiBusy}>Cancel</button>
          <button className="btn primary" disabled={aiBusy || (aiMode !== 'analyze' && !aiPrompt.trim())} onClick={() => void runAi(aiMode, aiPrompt)}>{aiBusy ? 'Working…' : 'Generate'}</button>
        </div>
      </BottomSheet>

      <FunctionPicker open={fnOpen} onClose={() => setFnOpen(false)} onPick={(n, sig) => insertFn(n, sig)} />
      <PromptSheet open={goOpen} title="Go to" label="Cell or range (e.g. C12 or A1:D20)" initial={sel} confirmLabel="Go" onSubmit={(v) => { const r = rangeToRect(v.trim().toUpperCase()); if (!r) { flash('Use a reference like B4 or A1:C9.'); return; } setSel(refOf(r.c1, r.r1)); setRangeEnd(r.c1 === r.c2 && r.r1 === r.r2 ? null : refOf(r.c2, r.r2)); setEditing(cells[refOf(r.c1, r.r1)] ?? ''); }} onClose={() => setGoOpen(false)} />
      <PromptSheet open={renameOpen} title="Rename sheet" initial={book.active} confirmLabel="Rename" onSubmit={renameSheet} onClose={() => setRenameOpen(false)} validate={(v) => (cleanSheetName(v) ? null : 'Enter a name.')} />
      <PromptSheet open={noteOpen} title={`Note on ${sel}`} initial={page.notes?.[sel] ?? ''} placeholder="Add a note (shown on long-press / hover)" multiline confirmLabel="Save" onSubmit={(v) => mutatePage((pg) => { const n = { ...pg.notes }; if (v.trim()) n[sel] = v.trim(); else delete n[sel]; return { ...pg, notes: n }; })} onClose={() => setNoteOpen(false)} />
      <ConfirmSheet open={confirmDelSheet} title={`Delete sheet "${book.active}"?`} message="All cells, formats and charts on this sheet will be removed. Undo is available right after." onConfirm={deleteSheet} onClose={() => setConfirmDelSheet(false)} />

      <Toast msg={toast} />
    </div>
  );
}

/** Checkbox list for the AutoFilter dropdown. */
function FilterList({ values, selected, onApply, onSort }: { values: { v: string; n: number }[]; selected: string[] | null; onApply: (sel: string[] | null) => void; onSort: (d: 1 | -1) => void }) {
  const [q, setQ] = useState('');
  const [picked, setPicked] = useState<Set<string>>(() => new Set(selected ?? values.map((x) => x.v)));
  const list = values.filter((x) => !q || x.v.toLowerCase().includes(q.toLowerCase()));
  const all = picked.size === values.length;
  return (
    <div className="filter-list">
      <div className="btn-row">
        <button className="btn small" onClick={() => onSort(1)}><Icon name="sortAZ" size={14} /> Sort A→Z</button>
        <button className="btn small" onClick={() => onSort(-1)}><Icon name="sortZA" size={14} /> Sort Z→A</button>
      </div>
      <input className="input" placeholder="Search values" value={q} onChange={(e) => setQ(e.target.value)} />
      <label className="check-row"><input type="checkbox" checked={all} onChange={() => setPicked(all ? new Set() : new Set(values.map((x) => x.v)))} /> <span>(Select all)</span><small>{values.length}</small></label>
      <div className="filter-values">
        {list.map((x) => (
          <label key={x.v} className="check-row">
            <input type="checkbox" checked={picked.has(x.v)} onChange={() => { const n = new Set(picked); if (n.has(x.v)) n.delete(x.v); else n.add(x.v); setPicked(n); }} />
            <span>{x.v === '' ? '(Blanks)' : x.v}</span><small>{x.n}</small>
          </label>
        ))}
      </div>
      <div className="btn-row end">
        <button className="btn small" onClick={() => onApply(null)}>Clear filter</button>
        <button className="btn small primary" onClick={() => onApply(all ? null : Array.from(picked))}>Apply</button>
      </div>
    </div>
  );
}
