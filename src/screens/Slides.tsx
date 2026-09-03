import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import type { PointerEvent as RPointerEvent } from 'react';
import { FileTypeIcon, Icon } from '../components/Icon';
import { AppBar, Palette, RBtn, RGroup, RSeg, RSelect, RStepper, RibbonPanel, RibbonTabs } from '../components/Ribbon';
import { BottomSheet, ConfirmSheet, SheetMenu, Toast, useToast } from '../components/Sheet';
import { SlideView } from '../components/SlideView';
import { chatStream, errMsg } from '../lib/ai-client';
import { exportPptx, exportSlidesPdf, importPptx, openFilePicker, pickImage, saveBinary, sanitizeName, slideToPng } from '../lib/fileio';
import { debounce, getDoc, getMeta, getSettings, putDoc, uid } from '../lib/storage';
import { keepAwake, onBack, tap } from '../lib/native';
import {
  CH, CW, DeckSlide, LAYOUTS, LayoutId, SlideShape, THEMES, TRANSITIONS, Theme, Transition,
  applyThemeToSlide, cloneSlide, deckFromOutline, firstRun, formatParas, formatShape, geomShape, isDark, newSlide, normalizeDeck, para, parseOutline, setShapeText, shapeId, shapeText, slideText, textBox, themeOf,
} from '../lib/deck-model';

type RibbonTab = 'home' | 'insert' | 'design' | 'transitions' | 'view';
type Panel = 'textColor' | 'fillColor' | 'bgColor' | 'lineColor' | 'shapes' | 'layouts' | null;

interface DeckData { slides: DeckSlide[]; }

const FONTS = ['Calibri', 'Segoe UI', 'Arial', 'Georgia', 'Cambria', 'Verdana', 'Trebuchet MS', 'Garamond', 'Courier New', 'Impact'].map((f) => ({ v: f, t: f }));
const SHAPE_PRESETS: { geom: NonNullable<SlideShape['geom']>; icon: string; label: string }[] = [
  { geom: 'rect', icon: 'square', label: 'Rectangle' },
  { geom: 'roundRect', icon: 'square', label: 'Rounded' },
  { geom: 'ellipse', icon: 'circle', label: 'Ellipse' },
  { geom: 'triangle', icon: 'triangle', label: 'Triangle' },
  { geom: 'diamond', icon: 'shapes', label: 'Diamond' },
  { geom: 'rightArrow', icon: 'arrowRight', label: 'Arrow' },
  { geom: 'chevron', icon: 'chevronRight', label: 'Chevron' },
  { geom: 'star', icon: 'sparkle', label: 'Star' },
  { geom: 'hexagon', icon: 'shapes', label: 'Hexagon' },
  { geom: 'line', icon: 'minus', label: 'Line' },
];

function loadDeck(id: string | undefined): DeckSlide[] {
  if (!id) return [];
  const d = getDoc<DeckData>(id);
  return d?.slides ? normalizeDeck(d.slides) : [];
}

export default function Slides({ initialId, onExit }: { initialId?: string; onExit?: () => void }) {
  const deckId = useRef(initialId ?? uid()).current;
  const [title, setTitle] = useState(() => (initialId ? getMeta(initialId)?.title ?? 'Presentation' : 'Untitled presentation'));
  const [slides, setSlides] = useState<DeckSlide[]>(() => loadDeck(initialId));
  const [sel, setSel] = useState(0);
  const [shapeSel, setShapeSel] = useState<string | null>(null);
  const [rTab, setRTab] = useState<RibbonTab>('home');
  const [panel, setPanel] = useState<Panel>(null);
  const [menu, setMenu] = useState(false);
  const [slideMenu, setSlideMenu] = useState(false);
  const [aiOpen, setAiOpen] = useState<null | 'outline' | 'rewrite' | 'notes' | 'image'>(null);
  const [aiTopic, setAiTopic] = useState('');
  const [aiBusy, setAiBusy] = useState(false);
  const [aiCount, setAiCount] = useState(6);
  const [toast, flash] = useToast();
  const [textEdit, setTextEdit] = useState<string | null>(null);
  const [textDraft, setTextDraft] = useState('');
  const [notesOpen, setNotesOpen] = useState(false);
  const [notesDraft, setNotesDraft] = useState('');
  const [sorter, setSorter] = useState(false);
  const [outlineOpen, setOutlineOpen] = useState(false);
  const [presenting, setPresenting] = useState<null | { idx: number; step: number; start: number; notes: boolean; laser: boolean; timerOn: boolean }>(null);
  const [presW, setPresW] = useState(0);
  const [presAnim, setPresAnim] = useState<{ dir: 1 | -1; t: Transition; key: number } | null>(null);
  const [newOpen, setNewOpen] = useState(!initialId);
  const [confirmDel, setConfirmDel] = useState(false);
  const [canvasW, setCanvasW] = useState(0);
  const [saved, setSaved] = useState<'saved' | 'saving' | 'dirty'>('saved');
  const [zoom, setZoom] = useState(100);
  const [clip, setClip] = useState<SlideShape | null>(null);
  const [, bump] = useState(0);
  const [tick, setTick] = useState(0);

  const holderRef = useRef<HTMLDivElement>(null);
  const filmRef = useRef<HTMLDivElement>(null);
  const past = useRef<DeckSlide[][]>([]);
  const future = useRef<DeckSlide[][]>([]);
  const drag = useRef<{ id: string; mode: 'move' | 'resize'; x: number; y: number; sx: number; sy: number; sw: number; sh: number; moved: boolean } | null>(null);
  const swipe = useRef<{ x: number; y: number; t: number } | null>(null);
  const laserPos = useRef<{ x: number; y: number } | null>(null);
  const [laser, setLaser] = useState<{ x: number; y: number } | null>(null);

  const cur = slides[sel];
  const theme = cur ? themeOf(cur) : THEMES[0];
  const curShape = cur?.shapes?.find((s) => s.id === shapeSel) ?? null;
  const curRun = firstRun(curShape ?? undefined);
  const visibleSlides = useMemo(() => slides.filter((s) => !s.hidden), [slides]);

  // ------------------------------------------------------------------ persistence
  const save = useMemo(() => debounce((sl: DeckSlide[], t: string) => { putDoc<DeckData>('deck', deckId, t, { slides: sl }); setSaved('saved'); }, 600), [deckId]);
  useEffect(() => { if (slides.length) { setSaved('saving'); save(slides, title); } }, [title]); // eslint-disable-line react-hooks/exhaustive-deps

  const update = (next: DeckSlide[], opts: { history?: boolean } = { history: true }) => {
    if (opts.history !== false) { past.current.push(slides); if (past.current.length > 60) past.current.shift(); future.current = []; }
    setSlides(next);
    setSaved('dirty');
    save(next, title);
    bump((v) => v + 1);
  };
  const undo = () => { const p = past.current.pop(); if (!p) return; future.current.push(slides); setSlides(p); save(p, title); setSel((i) => Math.min(i, p.length - 1)); bump((v) => v + 1); };
  const redo = () => { const f = future.current.pop(); if (!f) return; past.current.push(slides); setSlides(f); save(f, title); bump((v) => v + 1); };

  const editSlide = (patch: Partial<DeckSlide> | ((s: DeckSlide) => DeckSlide), idx = sel, history = true) =>
    update(slides.map((s, i) => (i === idx ? (typeof patch === 'function' ? patch(s) : { ...s, ...patch }) : s)), { history });
  const editShape = (id: string, fn: (sh: SlideShape) => SlideShape, history = true) =>
    editSlide((s) => ({ ...s, shapes: (s.shapes ?? []).map((sh) => (sh.id === id ? fn(sh) : sh)) }), sel, history);
  const editCurShape = (fn: (sh: SlideShape) => SlideShape) => { if (!shapeSel) { flash('Tap a text box or shape first.'); return; } editShape(shapeSel, fn); };

  // ------------------------------------------------------------------ layout / sizing
  useEffect(() => {
    const el = holderRef.current;
    if (!el) return;
    const inner = () => {
      const cs = getComputedStyle(el);
      return Math.max(0, el.clientWidth - parseFloat(cs.paddingLeft) - parseFloat(cs.paddingRight));
    };
    const ro = new ResizeObserver(() => setCanvasW(inner()));
    ro.observe(el);
    setCanvasW(inner());
    return () => ro.disconnect();
  }, [sel, presenting, slides.length, sorter]);

  useEffect(() => {
    if (!presenting) return;
    const calc = () => {
      const s = slides[presenting.idx];
      const cw = s?.cw ?? CW, ch = s?.ch ?? CH;
      const availH = window.innerHeight - (presenting.notes ? 190 : 56);
      setPresW(Math.max(120, Math.min(window.innerWidth - 8, (availH * cw) / ch)));
    };
    calc();
    window.addEventListener('resize', calc);
    return () => window.removeEventListener('resize', calc);
  }, [presenting, slides]);

  useEffect(() => { if (!presenting?.timerOn) return; const t = setInterval(() => setTick((x) => x + 1), 1000); return () => clearInterval(t); }, [presenting?.timerOn]);
  useEffect(() => { void keepAwake(!!presenting); return () => { void keepAwake(false); }; }, [!!presenting]); // eslint-disable-line react-hooks/exhaustive-deps

  // keep the selected thumbnail visible
  useEffect(() => { filmRef.current?.querySelector<HTMLElement>(`[data-idx="${sel}"]`)?.scrollIntoView({ inline: 'center', block: 'nearest', behavior: 'smooth' }); }, [sel]);

  // Android back button
  useEffect(
    () =>
      onBack(() => {
        if (presenting) { setPresenting(null); return true; }
        if (textEdit) { setTextEdit(null); return true; }
        if (panel) { setPanel(null); return true; }
        if (sorter) { setSorter(false); return true; }
        if (shapeSel) { setShapeSel(null); return true; }
        onExit?.();
        return true;
      }),
    [presenting, textEdit, panel, sorter, shapeSel, onExit],
  );

  // keyboard shortcuts while presenting (hardware keyboards / remotes)
  useEffect(() => {
    if (!presenting) return;
    const h = (e: KeyboardEvent) => {
      if (e.key === 'ArrowRight' || e.key === ' ' || e.key === 'PageDown' || e.key === 'Enter') { e.preventDefault(); presNext(); }
      else if (e.key === 'ArrowLeft' || e.key === 'PageUp' || e.key === 'Backspace') { e.preventDefault(); presPrev(); }
      else if (e.key === 'Escape') setPresenting(null);
      else if (e.key.toLowerCase() === 'n') setPresenting((p) => (p ? { ...p, notes: !p.notes } : p));
      else if (e.key === 'Home') setPresenting((p) => (p ? { ...p, idx: 0, step: 0 } : p));
      else if (e.key === 'End') setPresenting((p) => (p ? { ...p, idx: slides.length - 1, step: 99 } : p));
    };
    window.addEventListener('keydown', h);
    return () => window.removeEventListener('keydown', h);
  });

  // ------------------------------------------------------------------ slide ops
  const addSlide = (layout: LayoutId = 'content') => {
    const t = cur ? themeOf(cur) : THEMES[0];
    const s = newSlide(layout, t, {});
    const next = [...slides]; next.splice(sel + 1, 0, s);
    update(next); setSel(slides.length ? sel + 1 : 0); setShapeSel(null); setPanel(null);
    void tap();
  };
  const changeLayout = (layout: LayoutId) => {
    if (!cur) { addSlide(layout); return; }
    const { title: t, body } = slideText(cur);
    const img = cur.shapes?.find((s) => s.kind === 'image')?.img;
    const s = newSlide(layout, themeOf(cur), { title: t, bullets: body, image: img });
    editSlide({ ...s, id: cur.id, notes: cur.notes, transition: cur.transition });
    setShapeSel(null); setPanel(null);
  };
  const duplicateSlide = (idx = sel) => { const next = [...slides]; next.splice(idx + 1, 0, cloneSlide(slides[idx])); update(next); setSel(idx + 1); };
  const removeSlide = (idx = sel) => { const next = slides.filter((_, i) => i !== idx); update(next); setSel(Math.max(0, Math.min(idx, next.length - 1))); setShapeSel(null); };
  const moveSlide = (from: number, to: number) => { if (to < 0 || to >= slides.length || from === to) return; const next = [...slides]; const [s] = next.splice(from, 1); next.splice(to, 0, s); update(next); setSel(to); };
  const toggleHidden = (idx = sel) => editSlide((s) => ({ ...s, hidden: !s.hidden }), idx);

  const applyTheme = (t: Theme, all: boolean) => {
    update(slides.map((s, i) => (all || i === sel ? applyThemeToSlide(s, t) : s)));
    flash(all ? `${t.name} applied to all slides.` : `${t.name} applied.`);
  };

  // ------------------------------------------------------------------ shapes
  const addShape = (sh: SlideShape) => { editSlide((s) => ({ ...s, shapes: [...(s.shapes ?? []), sh] })); setShapeSel(sh.id!); setPanel(null); };
  const insertTextBox = () => addShape(textBox(2, 2, 6, 1, [para('New text', { sz: 20, color: theme.text, align: 'center' })], { valign: 'middle' }));
  const insertShape = (geom: NonNullable<SlideShape['geom']>) => addShape(geom === 'line' ? geomShape('line', 2, 2.8, 6, 0.05, theme.accent, { lineW: 3 }) : geomShape(geom, 3.25, 1.6, 3.5, 2.4, theme.accent));
  const insertImage = async () => {
    try {
      const data = await pickImage(1400);
      if (!data) return;
      const img = new Image();
      img.onload = () => {
        const ratio = img.naturalWidth / Math.max(1, img.naturalHeight);
        let w = 4.5, h = w / ratio;
        if (h > 3.6) { h = 3.6; w = h * ratio; }
        const placeholder = cur?.shapes?.find((s) => s.name === 'picture-placeholder');
        if (placeholder) editShape(placeholder.id!, () => ({ id: placeholder.id, x: placeholder.x, y: placeholder.y, w: placeholder.w, h: placeholder.h, kind: 'image', paras: [], img: data, name: 'picture' }));
        else addShape({ id: shapeId(), x: (CW - w) / 2, y: (CH - h) / 2, w, h, kind: 'image', paras: [], img: data, name: 'picture' });
      };
      img.src = data;
    } catch (e) { flash(`Could not insert image: ${errMsg(e)}`); }
  };
  const deleteShape = () => { if (!shapeSel) return; editSlide((s) => ({ ...s, shapes: (s.shapes ?? []).filter((sh) => sh.id !== shapeSel) })); setShapeSel(null); };
  const duplicateShape = () => { if (!curShape) return; addShape({ ...JSON.parse(JSON.stringify(curShape)), id: shapeId(), x: Math.min(CW - curShape.w, curShape.x + 0.3), y: Math.min(CH - curShape.h, curShape.y + 0.3), name: undefined, locked: false }); };
  const orderShape = (dir: 'front' | 'back' | 'up' | 'down') => {
    if (!shapeSel) return;
    editSlide((s) => {
      const arr = [...(s.shapes ?? [])]; const i = arr.findIndex((x) => x.id === shapeSel); if (i < 0) return s;
      const [sh] = arr.splice(i, 1);
      const j = dir === 'front' ? arr.length : dir === 'back' ? 0 : dir === 'up' ? Math.min(arr.length, i + 1) : Math.max(0, i - 1);
      arr.splice(j, 0, sh); return { ...s, shapes: arr };
    });
  };
  const alignShape = (how: 'left' | 'center' | 'right' | 'top' | 'middle' | 'bottom') =>
    editCurShape((sh) => ({ ...sh, x: how === 'left' ? 0.3 : how === 'center' ? (CW - sh.w) / 2 : how === 'right' ? CW - sh.w - 0.3 : sh.x, y: how === 'top' ? 0.3 : how === 'middle' ? (CH - sh.h) / 2 : how === 'bottom' ? CH - sh.h - 0.3 : sh.y }));
  const copyShape = () => { if (curShape) { setClip(JSON.parse(JSON.stringify(curShape))); flash('Shape copied.'); } };
  const pasteShape = () => { if (!clip) { flash('Nothing to paste.'); return; } addShape({ ...JSON.parse(JSON.stringify(clip)), id: shapeId(), locked: false }); };

  const openTextEdit = (id: string) => { const sh = cur?.shapes?.find((s) => s.id === id); if (!sh || sh.kind === 'image') return; setShapeSel(id); setTextDraft(shapeText(sh)); setTextEdit(id); };
  const saveTextEdit = () => { if (textEdit) editShape(textEdit, (sh) => setShapeText(sh, textDraft)); setTextEdit(null); };

  const stepFont = (d: 1 | -1) => editCurShape((sh) => formatShape(sh, (r) => ({ ...r, sz: Math.max(6, Math.min(96, Math.round((r.sz ?? 18) * (d === 1 ? 1.12 : 0.9)))) })));
  const toggleRun = (k: 'b' | 'i' | 'u' | 's') => editCurShape((sh) => { const on = !!firstRun(sh)?.[k]; return formatShape(sh, (r) => ({ ...r, [k]: !on })); });
  const setAlign = (a: 'left' | 'center' | 'right' | 'justify') => editCurShape((sh) => formatParas(sh, { align: a }));
  const toggleBullets = () => editCurShape((sh) => formatParas(sh, { bullet: !sh.paras[0]?.bullet }));
  const indent = (d: 1 | -1) => editCurShape((sh) => ({ ...sh, paras: sh.paras.map((p) => ({ ...p, level: Math.max(0, Math.min(4, (p.level ?? 0) + d)) })) }));

  // ------------------------------------------------------------------ drag / resize on canvas
  const onDragStart = (e: RPointerEvent, id: string, mode: 'move' | 'resize') => {
    const sh = cur?.shapes?.find((s) => s.id === id); if (!sh) return;
    (e.currentTarget as HTMLElement).setPointerCapture?.(e.pointerId);
    drag.current = { id, mode, x: e.clientX, y: e.clientY, sx: sh.x, sy: sh.y, sw: sh.w, sh: sh.h, moved: false };
  };
  const onCanvasMove = (e: RPointerEvent) => {
    const d = drag.current; if (!d || !canvasW) return;
    const scale = (cur?.cw ?? CW) / Math.min(canvasW, 720) / (zoom / 100); // inches per px
    const dx = (e.clientX - d.x) * scale, dy = (e.clientY - d.y) * scale;
    if (!d.moved && Math.hypot(e.clientX - d.x, e.clientY - d.y) < 4) return;
    d.moved = true;
    e.preventDefault();
    const snap = (v: number) => Math.round(v * 20) / 20;
    editShape(d.id, (sh) => (d.mode === 'move'
      ? { ...sh, x: snap(Math.max(-sh.w + 0.2, Math.min(CW - 0.2, d.sx + dx))), y: snap(Math.max(-sh.h + 0.2, Math.min(CH - 0.2, d.sy + dy))) }
      : { ...sh, w: snap(Math.max(0.3, d.sw + dx)), h: snap(Math.max(0.15, d.sh + dy)) }), false);
  };
  const onCanvasUp = () => { if (drag.current?.moved) { past.current.push(slides); future.current = []; } drag.current = null; };

  // ------------------------------------------------------------------ presenting
  const startPresent = (from = sel) => { setPresenting({ idx: from, step: 0, start: Date.now(), notes: false, laser: false, timerOn: true }); setPanel(null); void tap('medium'); };
  const animCount = (s: DeckSlide | undefined) => (s?.shapes ?? []).filter((sh) => sh.anim && sh.anim !== 'none').length;
  const presNext = () => setPresenting((p) => {
    if (!p) return p;
    const s = slides[p.idx];
    if (p.step < animCount(s)) return { ...p, step: p.step + 1 };
    let i = p.idx + 1; while (i < slides.length && slides[i].hidden) i++;
    if (i >= slides.length) return p;
    setPresAnim({ dir: 1, t: slides[i].transition ?? 'none', key: Date.now() });
    return { ...p, idx: i, step: 0 };
  });
  const presPrev = () => setPresenting((p) => {
    if (!p) return p;
    if (p.step > 0) return { ...p, step: p.step - 1 };
    let i = p.idx - 1; while (i >= 0 && slides[i].hidden) i--;
    if (i < 0) return p;
    setPresAnim({ dir: -1, t: slides[p.idx].transition ?? 'none', key: Date.now() });
    return { ...p, idx: i, step: animCount(slides[i]) };
  });

  // ------------------------------------------------------------------ AI
  const runAi = async () => {
    const s = getSettings();
    if (!s.apiKey) { flash('Add your API key in Settings first.'); setAiOpen(null); return; }
    const topic = aiTopic.trim();
    setAiBusy(true);
    try {
      if (aiOpen === 'outline') {
        if (!topic) return;
        const out = await chatStream(s, [
          { role: 'system', content: `You create presentation outlines. Produce exactly ${aiCount} slides. Format each slide as: a line "### <slide title>", then 3-5 lines starting with "- " (bullets, max 12 words each), then optionally one line "Notes: <speaker notes, 1-2 sentences>". First slide is the title slide with one bullet as subtitle. No extra commentary. Same language as the request.` },
          { role: 'user', content: topic },
        ]);
        const outline = parseOutline(out);
        if (!outline.length) throw new Error('Could not parse the outline; please try again.');
        const t = cur ? themeOf(cur) : THEMES[3];
        const deck = deckFromOutline(outline, t, topic);
        update(slides.length && !newOpen ? [...slides, ...deck] : deck);
        setSel(slides.length && !newOpen ? slides.length : 0);
        if (title.startsWith('Untitled')) setTitle(outline[0]?.title.slice(0, 60) || title);
        setNewOpen(false);
        flash(`Created ${deck.length} slides.`);
      } else if (aiOpen === 'rewrite') {
        if (!curShape) throw new Error('Select a text box first.');
        const out = await chatStream(s, [
          { role: 'system', content: 'You improve presentation slide text. Keep it as concise bullet-style lines, same language, same number of lines unless asked otherwise. Output only the lines, one per line, no bullets or numbering characters.' },
          { role: 'user', content: `${topic ? `Instruction: ${topic}\n` : 'Make it clearer and punchier.\n'}Text:\n${shapeText(curShape)}` },
        ]);
        editShape(curShape.id!, (sh) => setShapeText(sh, out.trim().split('\n').map((l) => l.replace(/^[-*•]\s+/, '')).join('\n')));
        flash('Text updated.');
      } else if (aiOpen === 'notes') {
        if (!cur) return;
        const { title: t, body } = slideText(cur);
        const out = await chatStream(s, [
          { role: 'system', content: 'Write concise speaker notes (60-110 words) for a presentation slide: what to say, in a natural spoken tone. Same language as the slide. Plain text only.' },
          { role: 'user', content: `Slide title: ${t}\nBullets:\n${body.join('\n')}${topic ? `\nAudience/context: ${topic}` : ''}` },
        ]);
        editSlide({ notes: out.trim() });
        setNotesDraft(out.trim());
        flash('Speaker notes added.');
      }
      setAiOpen(null);
      setAiTopic('');
    } catch (e) { flash(`Error: ${errMsg(e)}`); } finally { setAiBusy(false); }
  };

  // ------------------------------------------------------------------ files
  const openFile = async () => {
    const pick = await openFilePicker('.pptx');
    if (!pick) return;
    try {
      const loaded = normalizeDeck(await importPptx(pick.buf));
      if (!loaded.length) throw new Error('No slides found in the file');
      update(loaded); setSel(0); setShapeSel(null); setNewOpen(false);
      setTitle(pick.name.replace(/\.[^.]+$/, ''));
      flash(`Opened ${pick.name} (${loaded.length} slides)`);
    } catch (e) { flash(`Could not open: ${errMsg(e)}`); }
  };
  const savePptx = async () => {
    if (!slides.length) return;
    try { const bytes = await exportPptx(title, slides); flash(await saveBinary(sanitizeName(title, 'pptx'), bytes, 'application/vnd.openxmlformats-officedocument.presentationml.presentation')); } catch (e) { flash(`Save failed: ${errMsg(e)}`); }
  };
  const savePdf = async () => {
    if (!slides.length) return;
    try { flash('Rendering PDF…', 10000); const bytes = await exportSlidesPdf(title, visibleSlides, { notes: false }); flash(await saveBinary(sanitizeName(title, 'pdf'), bytes, 'application/pdf')); } catch (e) { flash(`PDF failed: ${errMsg(e)}`); }
  };
  const saveHandout = async () => {
    if (!slides.length) return;
    try { flash('Rendering handout…', 10000); const bytes = await exportSlidesPdf(title, visibleSlides, { notes: true }); flash(await saveBinary(sanitizeName(`${title}-notes`, 'pdf'), bytes, 'application/pdf')); } catch (e) { flash(`PDF failed: ${errMsg(e)}`); }
  };
  const saveSlideImage = async () => {
    if (!cur) return;
    try { const png = await slideToPng(cur, 1600); flash(await saveBinary(sanitizeName(`${title}-slide${sel + 1}`, 'png'), png, 'image/png')); } catch (e) { flash(`Image export failed: ${errMsg(e)}`); }
  };
  const exportOutlineText = () => {
    const txt = slides.map((s, i) => { const { title: t, body } = slideText(s); return `${i + 1}. ${t}\n${body.map((b) => `   - ${b}`).join('\n')}${s.notes ? `\n   Notes: ${s.notes}` : ''}`; }).join('\n\n');
    void navigator.clipboard?.writeText(txt).then(() => flash('Outline copied to clipboard.')).catch(() => flash('Could not copy.'));
  };
  const startFromTemplate = (t: Theme, kind: 'blank' | 'pitch' | 'lesson' | 'report') => {
    let deck: DeckSlide[];
    if (kind === 'blank') deck = [newSlide('title', t, { title: 'Presentation title', bullets: ['Subtitle'] })];
    else if (kind === 'pitch') deck = deckFromOutline([
      { title: 'Company name', bullets: ['One-line pitch'] }, { title: 'The problem', bullets: ['Who has it', 'Why it matters now', 'What people do today'] },
      { title: 'Our solution', bullets: ['What it is', 'How it works', 'Why it is 10× better'] }, { title: 'Market', bullets: ['Size', 'Growth', 'Who we start with'] },
      { title: 'Business model', bullets: ['How we make money', 'Pricing', 'Unit economics'] }, { title: 'Traction', bullets: ['Users / revenue', 'Key partners', 'Milestones'] },
      { title: 'Team', bullets: ['Founders', 'Advisors'] }, { title: 'The ask', bullets: ['How much', 'What for', 'Contact'] },
    ], t);
    else if (kind === 'lesson') deck = deckFromOutline([
      { title: 'Lesson title', bullets: ['Subject · Class · Date'] }, { title: 'Learning objectives', bullets: ['Objective 1', 'Objective 2', 'Objective 3'] },
      { title: 'Warm-up question', bullets: [] }, { title: 'Key concept', bullets: ['Definition', 'Example', 'Why it matters'] },
      { title: 'Worked example', bullets: ['Step 1', 'Step 2', 'Step 3'] }, { title: 'Practice', bullets: ['Task A', 'Task B'] }, { title: 'Summary & homework', bullets: ['Recap', 'Homework due …'] },
    ], t);
    else deck = deckFromOutline([
      { title: 'Project status report', bullets: [new Date().toLocaleDateString()] }, { title: 'Summary', bullets: ['Overall status: on track', 'Key wins', 'Key risks'] },
      { title: 'Progress this period', bullets: ['Done', 'In progress', 'Blocked'] }, { title: 'Metrics', bullets: ['KPI 1', 'KPI 2', 'KPI 3'] },
      { title: 'Risks & mitigations', bullets: ['Risk → mitigation'] }, { title: 'Next steps', bullets: ['Action · owner · date'] },
    ], t);
    update(deck); setSel(0); setNewOpen(false); setShapeSel(null);
  };

  // ------------------------------------------------------------------ render: presenter
  if (presenting && slides.length > 0) {
    const p = presenting;
    const s = slides[p.idx];
    const elapsed = Math.floor((Date.now() - p.start) / 1000);
    const mm = String(Math.floor(elapsed / 60)).padStart(2, '0'), ss = String(elapsed % 60).padStart(2, '0');
    const visIdx = visibleSlides.indexOf(s);
    const nextSlide = slides.slice(p.idx + 1).find((x) => !x.hidden);
    void tick;
    return (
      <div
        className={`present${p.notes ? ' with-notes' : ''}`}
        onPointerDown={(e) => { swipe.current = { x: e.clientX, y: e.clientY, t: Date.now() }; if (p.laser) { laserPos.current = { x: e.clientX, y: e.clientY }; setLaser(laserPos.current); } }}
        onPointerMove={(e) => { if (p.laser && laserPos.current) setLaser({ x: e.clientX, y: e.clientY }); }}
        onPointerUp={(e) => {
          const sw = swipe.current; swipe.current = null;
          if (p.laser) { setLaser(null); laserPos.current = null; return; }
          if (!sw) return;
          const dx = e.clientX - sw.x, dy = e.clientY - sw.y, dt = Date.now() - sw.t;
          if (Math.abs(dx) > 50 && Math.abs(dx) > Math.abs(dy) * 1.3 && dt < 700) { if (dx < 0) presNext(); else presPrev(); return; }
          if (Math.abs(dx) < 10 && Math.abs(dy) < 10) {
            const target = e.target as HTMLElement;
            if (target.closest('.present-bar') || target.closest('.present-notes')) return;
            if (e.clientX < window.innerWidth * 0.3) presPrev(); else presNext();
          }
        }}
      >
        <div className={`present-stage${presAnim ? ` anim-${presAnim.t} dir${presAnim.dir}` : ''}`} key={presAnim?.key ?? 0} onAnimationEnd={() => setPresAnim(null)}>
          <SlideView slide={s} width={presW} animStep={p.step} />
        </div>
        {laser && <div className="laser-dot" style={{ left: laser.x, top: laser.y }} />}
        {p.notes && (
          <div className="present-notes">
            <div className="present-notes-text">{s.notes?.trim() ? s.notes : <i>No notes for this slide.</i>}</div>
            {nextSlide && <div className="present-next"><span>Next</span><SlideView slide={nextSlide} width={96} /></div>}
          </div>
        )}
        <div className="present-bar" onPointerDown={(e) => e.stopPropagation()} onPointerUp={(e) => e.stopPropagation()}>
          <button className="icon-btn light" aria-label="Previous slide" onClick={presPrev}><Icon name="chevronLeft" size={22} /></button>
          <span className="present-count">{visIdx + 1} / {visibleSlides.length}</span>
          <button className="icon-btn light" aria-label="Next slide" onClick={presNext}><Icon name="chevronRight" size={22} /></button>
          <span className="present-spacer" />
          <button className={`present-timer${p.timerOn ? '' : ' paused'}`} onClick={() => setPresenting({ ...p, timerOn: !p.timerOn })} aria-label="Toggle timer"><Icon name="timer" size={16} /> {mm}:{ss}</button>
          <button className={`icon-btn light${p.laser ? ' on' : ''}`} aria-label="Laser pointer" onClick={() => setPresenting({ ...p, laser: !p.laser })}><Icon name="laser" size={20} /></button>
          <button className={`icon-btn light${p.notes ? ' on' : ''}`} aria-label="Speaker notes" onClick={() => setPresenting({ ...p, notes: !p.notes })}><Icon name="notes" size={20} /></button>
          <button className="icon-btn light" aria-label="Exit presentation" onClick={() => setPresenting(null)}><Icon name="close" size={22} /></button>
        </div>
      </div>
    );
  }

  // ------------------------------------------------------------------ render: sorter
  const slideCard = (s: DeckSlide, i: number, w: number, extra?: React.ReactNode) => (
    <button
      key={s.id ?? i}
      data-idx={i}
      className={`film-thumb${i === sel ? ' active' : ''}${s.hidden ? ' hidden-slide' : ''}`}
      onClick={() => { setSel(i); setShapeSel(null); }}
      onContextMenu={(e) => { e.preventDefault(); setSel(i); setSlideMenu(true); }}
    >
      <SlideView slide={s} width={w} />
      <span className="film-num">{i + 1}</span>
      {s.hidden && <span className="film-hidden"><Icon name="eye" size={12} /></span>}
      {s.transition && s.transition !== 'none' && <span className="film-trans"><Icon name="transition" size={11} /></span>}
      {s.notes && <span className="film-notes"><Icon name="notes" size={11} /></span>}
      {extra}
    </button>
  );

  const canvasWidth = Math.min(canvasW, 720) * (zoom / 100);

  return (
    <div className="edscreen" style={{ ['--app' as string]: 'var(--ppt)' }}>
      <AppBar kindIcon={<FileTypeIcon kind="deck" size={24} light />} title={title} onTitle={setTitle} placeholder="Presentation title" onBack={onExit} saved={saved}>
        <button className="icon-btn light" aria-label="Undo" disabled={past.current.length === 0} onClick={undo}><Icon name="undo" size={20} /></button>
        <button className="icon-btn light" aria-label="Present from current slide" disabled={!slides.length} onClick={() => startPresent(sel)}><Icon name="play" size={20} /></button>
        <button className="icon-btn light" aria-label="More actions" onClick={() => setMenu(true)}><Icon name="more" size={20} /></button>
      </AppBar>

      {sorter ? (
        <div className="edbody sorter" ref={holderRef}>
          <p className="hint" style={{ margin: '4px 8px' }}>Slide sorter — tap to select, use the arrows to reorder, long-press for options.</p>
          <div className="sorter-grid">
            {slides.map((s, i) => (
              <div key={s.id ?? i} className="sorter-item">
                {slideCard(s, i, Math.max(120, (Math.min(canvasW, 720) - 36) / 2))}
                <div className="sorter-actions">
                  <button className="icon-btn" aria-label="Move earlier" disabled={i === 0} onClick={() => moveSlide(i, i - 1)}><Icon name="chevronLeft" size={16} /></button>
                  <button className="icon-btn" aria-label="Move later" disabled={i === slides.length - 1} onClick={() => moveSlide(i, i + 1)}><Icon name="chevronRight" size={16} /></button>
                  <button className="icon-btn" aria-label="Duplicate" onClick={() => duplicateSlide(i)}><Icon name="duplicate" size={16} /></button>
                  <button className="icon-btn" aria-label={s.hidden ? 'Unhide' : 'Hide'} onClick={() => toggleHidden(i)}><Icon name="eye" size={16} /></button>
                  <button className="icon-btn danger" aria-label="Delete" onClick={() => { setSel(i); setConfirmDel(true); }}><Icon name="trash" size={16} /></button>
                </div>
              </div>
            ))}
            <button className="sorter-add" onClick={() => addSlide('content')}><Icon name="plus" size={26} /><span>New slide</span></button>
          </div>
        </div>
      ) : (
        <div className="edbody deck-body">
          {slides.length === 0 ? (
            <div className="deck-empty">
              <FileTypeIcon kind="deck" size={54} />
              <p>No slides yet.</p>
              <div className="btn-row">
                <button className="btn primary" onClick={() => setNewOpen(true)}>Start</button>
                <button className="btn" onClick={() => setAiOpen('outline')}><Icon name="ai" size={16} /> AI outline</button>
              </div>
            </div>
          ) : (
            cur && (
              <div className="shape-canvas-holder" ref={holderRef} style={{ borderColor: cur.accent ?? 'var(--ppt)' }} onPointerMove={onCanvasMove} onPointerUp={onCanvasUp} onPointerCancel={onCanvasUp}>
                <SlideView slide={cur} width={canvasWidth} edit selected={shapeSel} onSelect={setShapeSel} onDoubleTap={openTextEdit} onDragStart={onDragStart} />
                <div className="canvas-foot">
                  <span>Slide {sel + 1} of {slides.length}{cur.hidden ? ' · hidden' : ''}</span>
                  <button className={`status-btn${cur.notes ? ' on' : ''}`} onClick={() => { setNotesDraft(cur.notes ?? ''); setNotesOpen(true); }}><Icon name="notes" size={14} /> {cur.notes ? 'Notes' : 'Add notes'}</button>
                </div>
                {shapeSel && curShape && (
                  <div className="shape-bar">
                    {curShape.kind !== 'image' && <button className="btn small" onClick={() => openTextEdit(shapeSel)}><Icon name="edit" size={14} /> Edit text</button>}
                    <button className="icon-btn" aria-label="Duplicate shape" onClick={duplicateShape}><Icon name="duplicate" size={16} /></button>
                    <button className="icon-btn" aria-label="Bring forward" onClick={() => orderShape('up')}><Icon name="moveUp" size={16} /></button>
                    <button className="icon-btn" aria-label="Send backward" onClick={() => orderShape('down')}><Icon name="moveDown" size={16} /></button>
                    <button className="icon-btn danger" aria-label="Delete shape" onClick={deleteShape}><Icon name="trash" size={16} /></button>
                  </div>
                )}
              </div>
            )
          )}

          {slides.length > 0 && (
            <div className="filmstrip" ref={filmRef}>
              {slides.map((s, i) => slideCard(s, i, 112))}
              <button className="film-add" aria-label="New slide" onClick={() => addSlide('content')}><Icon name="plus" size={22} /></button>
            </div>
          )}
        </div>
      )}

      {/* ribbon */}
      <div className="ribbon">
        {panel === 'textColor' && <Palette current={curRun?.color} onPick={(c) => { editCurShape((sh) => formatShape(sh, { color: c })); setPanel(null); }} auto={() => { editCurShape((sh) => formatShape(sh, (r) => ({ ...r, color: undefined }))); setPanel(null); }} autoLabel="Automatic" />}
        {panel === 'fillColor' && <Palette current={curShape?.fill} onPick={(c) => { editCurShape((sh) => ({ ...sh, fill: c })); setPanel(null); }} auto={() => { editCurShape((sh) => ({ ...sh, fill: undefined })); setPanel(null); }} autoLabel="No fill" />}
        {panel === 'lineColor' && <Palette current={curShape?.line} onPick={(c) => { editCurShape((sh) => ({ ...sh, line: c, lineW: sh.lineW ?? 1.5 })); setPanel(null); }} auto={() => { editCurShape((sh) => ({ ...sh, line: undefined })); setPanel(null); }} autoLabel="No outline" />}
        {panel === 'bgColor' && <Palette current={cur?.bg} onPick={(c) => { editSlide({ bg: c }); setPanel(null); }} auto={() => { editSlide({ bg: theme.bg }); setPanel(null); }} autoLabel="Theme color" />}
        {panel === 'shapes' && (
          <RibbonPanel title="Shapes" onClose={() => setPanel(null)}>
            <div className="rpanel-actions">
              {SHAPE_PRESETS.map((p) => <button key={p.geom} className="btn small" onClick={() => insertShape(p.geom)}><Icon name={p.icon} size={16} /> {p.label}</button>)}
            </div>
          </RibbonPanel>
        )}
        {panel === 'layouts' && (
          <RibbonPanel title={cur ? 'Change layout (keeps your text)' : 'New slide layout'} onClose={() => setPanel(null)}>
            <div className="layout-gallery">
              {LAYOUTS.map((l) => (
                <button key={l.id} className="layout-card" onClick={() => (cur ? changeLayout(l.id) : addSlide(l.id))}>
                  <SlideView slide={newSlide(l.id, theme, { title: l.label, bullets: ['Text', 'Text'] })} width={104} />
                  <span>{l.label}</span>
                </button>
              ))}
            </div>
          </RibbonPanel>
        )}

        <RibbonTabs tabs={[{ id: 'home', label: 'Home' }, { id: 'insert', label: 'Insert' }, { id: 'design', label: 'Design' }, { id: 'transitions', label: 'Transitions' }, { id: 'view', label: 'View' }]} value={rTab} onChange={(t) => { setRTab(t); setPanel(null); void tap(); }} />

        {rTab === 'home' && (
          <div className="ribbon-row">
            <RGroup label="Slides">
              <RBtn icon="plus" label="New slide" onRun={() => addSlide('content')} />
              <RBtn icon="layoutContent" label="Layout" menu active={panel === 'layouts'} onRun={() => setPanel(panel === 'layouts' ? null : 'layouts')} />
              <RBtn icon="duplicate" label="Duplicate" disabled={!cur} onRun={() => duplicateSlide()} />
              <RBtn icon="trash" label="Delete" disabled={!cur} onRun={() => setConfirmDel(true)} />
            </RGroup>
            <RGroup label="Font">
              <RSelect value={curRun?.font ?? theme.font ?? 'Calibri'} options={FONTS} onChange={(f) => editCurShape((sh) => formatShape(sh, { font: f }))} width={100} title="Font" />
              <RStepper value={String(curRun?.sz ?? 18)} title="Font size" onDec={() => stepFont(-1)} onInc={() => stepFont(1)} />
              <RBtn icon="bold" label="Bold" active={!!curRun?.b} onRun={() => toggleRun('b')} />
              <RBtn icon="italic" label="Italic" active={!!curRun?.i} onRun={() => toggleRun('i')} />
              <RBtn icon="underline" label="Underline" active={!!curRun?.u} onRun={() => toggleRun('u')} />
              <RBtn icon="strike" label="Strike" active={!!curRun?.s} onRun={() => toggleRun('s')} />
              <RBtn icon="fontColor" label="Color" colorBar={curRun?.color ?? theme.text} menu active={panel === 'textColor'} onRun={() => setPanel(panel === 'textColor' ? null : 'textColor')} />
              <RBtn icon="highlight" label="Highlight" colorBar={curRun?.highlight ?? '#FFFF00'} onRun={() => editCurShape((sh) => formatShape(sh, (r) => ({ ...r, highlight: r.highlight ? undefined : '#FFF176' })))} />
            </RGroup>
            <RGroup label="Paragraph">
              <RBtn icon="listBullet" label="Bullets" active={!!curShape?.paras[0]?.bullet} onRun={toggleBullets} />
              <RBtn icon="outdent" label="Outdent" onRun={() => indent(-1)} />
              <RBtn icon="indent" label="Indent" onRun={() => indent(1)} />
              <RBtn icon="alignLeft" label="Left" active={curShape?.paras[0]?.align === 'left' || (!!curShape && !curShape.paras[0]?.align)} onRun={() => setAlign('left')} />
              <RBtn icon="alignCenter" label="Center" active={curShape?.paras[0]?.align === 'center'} onRun={() => setAlign('center')} />
              <RBtn icon="alignRight" label="Right" active={curShape?.paras[0]?.align === 'right'} onRun={() => setAlign('right')} />
              <RBtn icon="alignTop" label="Top" active={(curShape?.valign ?? 'top') === 'top'} onRun={() => editCurShape((sh) => ({ ...sh, valign: 'top' }))} />
              <RBtn icon="alignMiddle" label="Middle" active={curShape?.valign === 'middle'} onRun={() => editCurShape((sh) => ({ ...sh, valign: 'middle' }))} />
            </RGroup>
            <RGroup label="Shape">
              <RBtn icon="fill" label="Fill" colorBar={curShape?.fill ?? theme.accent} menu active={panel === 'fillColor'} onRun={() => setPanel(panel === 'fillColor' ? null : 'fillColor')} />
              <RBtn icon="border" label="Outline" colorBar={curShape?.line ?? '#404040'} menu active={panel === 'lineColor'} onRun={() => setPanel(panel === 'lineColor' ? null : 'lineColor')} />
              <RBtn icon="theme" label="Shadow" active={!!curShape?.shadow} onRun={() => editCurShape((sh) => ({ ...sh, shadow: !sh.shadow }))} />
              <RBtn icon="history" label="Rotate" onRun={() => editCurShape((sh) => ({ ...sh, rot: ((sh.rot ?? 0) + 15) % 360 }))} />
              <RBtn icon="moveUp" label="Forward" onRun={() => orderShape('front')} />
              <RBtn icon="moveDown" label="Back" onRun={() => orderShape('back')} />
            </RGroup>
            <RGroup label="Arrange">
              <RBtn icon="alignLeft" label="Left" onRun={() => alignShape('left')} />
              <RBtn icon="alignCenter" label="Center" onRun={() => alignShape('center')} />
              <RBtn icon="alignRight" label="Right" onRun={() => alignShape('right')} />
              <RBtn icon="alignMiddle" label="Middle" onRun={() => alignShape('middle')} />
              <RBtn icon="copy" label="Copy" disabled={!curShape} onRun={copyShape} />
              <RBtn icon="paste" label="Paste" disabled={!clip} onRun={pasteShape} />
            </RGroup>
          </div>
        )}

        {rTab === 'insert' && (
          <div className="ribbon-row">
            <RGroup label="Slides">
              <RBtn icon="plus" label="New slide" onRun={() => addSlide('content')} />
              <RBtn icon="layoutContent" label="Layouts" menu active={panel === 'layouts'} onRun={() => setPanel(panel === 'layouts' ? null : 'layouts')} />
            </RGroup>
            <RGroup label="Content">
              <RBtn icon="textBox" label="Text box" disabled={!cur} onRun={insertTextBox} />
              <RBtn icon="image" label="Picture" disabled={!cur} onRun={() => void insertImage()} />
              <RBtn icon="shapes" label="Shapes" disabled={!cur} menu active={panel === 'shapes'} onRun={() => setPanel(panel === 'shapes' ? null : 'shapes')} />
              <RBtn icon="table" label="Table" disabled={!cur} onRun={() => addShape(textBox(1, 1.6, 8, 2.6, [para('Header 1\tHeader 2\tHeader 3', { sz: 16, b: true, color: theme.text }), para('Cell\tCell\tCell', { sz: 16, color: theme.text }), para('Cell\tCell\tCell', { sz: 16, color: theme.text })], { fill: isDark(theme.bg) ? '#2A3140' : '#F3F4F6', line: theme.accent }))} />
            </RGroup>
            <RGroup label="Text">
              <RBtn icon="notes" label="Notes" disabled={!cur} onRun={() => { setNotesDraft(cur?.notes ?? ''); setNotesOpen(true); }} />
              <RBtn icon="calendar" label="Date" disabled={!cur} onRun={() => addShape(textBox(6.8, 5.1, 3, 0.4, [para(new Date().toLocaleDateString(), { sz: 11, color: theme.sub, align: 'right' })]))} />
              <RBtn icon="wordCount" label="Slide #" disabled={!cur} onRun={() => addShape(textBox(8.8, 5.1, 1, 0.4, [para(String(sel + 1), { sz: 11, color: theme.sub, align: 'right' })]))} />
              <RBtn icon="link" label="Footer" disabled={!cur} onRun={() => addShape(textBox(0.4, 5.1, 6, 0.4, [para(title, { sz: 11, color: theme.sub })]))} />
            </RGroup>
            <RGroup label="AI">
              <RBtn icon="ai" label="Outline" onRun={() => setAiOpen('outline')} />
              <RBtn icon="sparkle" label="Rewrite" disabled={!curShape || curShape.kind === 'image'} onRun={() => setAiOpen('rewrite')} />
              <RBtn icon="notes" label="AI notes" disabled={!cur} onRun={() => setAiOpen('notes')} />
            </RGroup>
          </div>
        )}

        {rTab === 'design' && (
          <div className="ribbon-row">
            <RGroup label="Themes (tap: this slide · hold: all)">
              <div className="theme-strip">
                {THEMES.map((t) => (
                  <ThemeChip key={t.name} t={t} active={theme.name === t.name} onTap={() => applyTheme(t, false)} onHold={() => applyTheme(t, true)} />
                ))}
              </div>
            </RGroup>
            <RGroup label="Background">
              <RBtn icon="fill" label="Color" colorBar={cur?.bg ?? '#FFFFFF'} menu active={panel === 'bgColor'} onRun={() => setPanel(panel === 'bgColor' ? null : 'bgColor')} />
              <RBtn icon="theme" label="Apply all" disabled={!cur} onRun={() => applyTheme(theme, true)} />
            </RGroup>
            <RGroup label="Layout">
              <RBtn icon="layoutContent" label="Layouts" menu active={panel === 'layouts'} onRun={() => setPanel(panel === 'layouts' ? null : 'layouts')} />
            </RGroup>
          </div>
        )}

        {rTab === 'transitions' && (
          <div className="ribbon-row">
            <RGroup label="Transition to this slide">
              <RSeg value={cur?.transition ?? 'none'} options={TRANSITIONS.map((t) => ({ v: t.id, t: t.label }))} onChange={(v) => editSlide({ transition: v })} />
              <RBtn icon="transition" label="Apply all" disabled={!cur} onRun={() => { update(slides.map((s) => ({ ...s, transition: cur?.transition ?? 'none' }))); flash('Transition applied to all slides.'); }} />
            </RGroup>
            <RGroup label="Shape animation">
              <RSeg value={curShape?.anim ?? 'none'} options={[{ v: 'none', t: 'None' }, { v: 'fadeIn', t: 'Fade in' }, { v: 'flyIn', t: 'Fly in' }, { v: 'zoomIn', t: 'Zoom in' }]} onChange={(v) => editCurShape((sh) => ({ ...sh, anim: v }))} />
            </RGroup>
            <RGroup label="Preview">
              <RBtn icon="play" label="Preview" disabled={!cur} onRun={() => startPresent(sel)} />
            </RGroup>
          </div>
        )}

        {rTab === 'view' && (
          <div className="ribbon-row">
            <RGroup label="Views">
              <RSeg value={sorter ? 'sorter' : 'normal'} options={[{ v: 'normal', t: 'Normal', icon: 'layoutContent' }, { v: 'sorter', t: 'Sorter', icon: 'grid4' }]} onChange={(v) => setSorter(v === 'sorter')} />
              <RBtn icon="toc" label="Outline" onRun={() => setOutlineOpen(true)} />
            </RGroup>
            <RGroup label="Zoom">
              <RStepper value={`${zoom}%`} title="Zoom" width={44} onDec={() => setZoom((z) => Math.max(50, z - 10))} onInc={() => setZoom((z) => Math.min(150, z + 10))} />
              <RBtn icon="fitWidth" label="Fit" onRun={() => setZoom(100)} />
            </RGroup>
            <RGroup label="Present">
              <RBtn icon="play" label="From start" disabled={!slides.length} onRun={() => startPresent(0)} />
              <RBtn icon="monitor" label="From here" disabled={!slides.length} onRun={() => startPresent(sel)} />
              <RBtn icon="eye" label={cur?.hidden ? 'Unhide' : 'Hide slide'} active={!!cur?.hidden} disabled={!cur} onRun={() => toggleHidden()} />
            </RGroup>
          </div>
        )}
      </div>

      {/* ---------------- sheets & dialogs ---------------- */}
      <BottomSheet open={menu} onClose={() => setMenu(false)} title={title}>
        <SheetMenu onClose={() => setMenu(false)} items={[
          { icon: 'fileOpen', label: 'Open .pptx', onRun: () => void openFile() },
          { icon: 'template', label: 'New from template', onRun: () => setNewOpen(true) },
          'divider',
          { icon: 'save', label: 'Save as PowerPoint (.pptx)', disabled: !slides.length, onRun: () => void savePptx() },
          { icon: 'pdf', label: 'Save as PDF', disabled: !slides.length, onRun: () => void savePdf() },
          { icon: 'notes', label: 'Handout PDF (slides + notes)', disabled: !slides.length, onRun: () => void saveHandout() },
          { icon: 'image', label: 'Save current slide as image', disabled: !cur, onRun: () => void saveSlideImage() },
          { icon: 'copy', label: 'Copy outline as text', disabled: !slides.length, onRun: exportOutlineText },
          'divider',
          { icon: 'play', label: 'Present from start', disabled: !slides.length, onRun: () => startPresent(0) },
        ]} />
      </BottomSheet>

      <BottomSheet open={slideMenu} onClose={() => setSlideMenu(false)} title={`Slide ${sel + 1}`}>
        <SheetMenu onClose={() => setSlideMenu(false)} items={[
          { icon: 'duplicate', label: 'Duplicate', onRun: () => duplicateSlide() },
          { icon: 'layoutContent', label: 'Change layout', onRun: () => setPanel('layouts') },
          { icon: 'notes', label: 'Speaker notes', onRun: () => { setNotesDraft(cur?.notes ?? ''); setNotesOpen(true); } },
          { icon: 'eye', label: cur?.hidden ? 'Unhide slide' : 'Hide slide', onRun: () => toggleHidden() },
          { icon: 'chevronLeft', label: 'Move earlier', disabled: sel === 0, onRun: () => moveSlide(sel, sel - 1) },
          { icon: 'chevronRight', label: 'Move later', disabled: sel >= slides.length - 1, onRun: () => moveSlide(sel, sel + 1) },
          'divider',
          { icon: 'trash', label: 'Delete slide', danger: true, onRun: () => setConfirmDel(true) },
        ]} />
      </BottomSheet>

      <BottomSheet open={newOpen} onClose={() => setNewOpen(false)} title="New presentation" tall>
        <p className="hint">Pick a theme, then a starting point.</p>
        <div className="theme-picker">
          {THEMES.map((t) => <ThemeChip key={t.name} t={t} active={theme.name === t.name} big onTap={() => { if (cur) applyTheme(t, true); else update([newSlide('title', t, { title: 'Presentation title', bullets: ['Subtitle'] })], { history: false }); }} />)}
        </div>
        <div className="template-grid">
          {([['blank', 'Blank', 'Title slide only'], ['pitch', 'Pitch deck', '8 slides'], ['lesson', 'Lesson', '7 slides'], ['report', 'Status report', '6 slides']] as const).map(([k, l, d]) => (
            <button key={k} className="template-card" onClick={() => startFromTemplate(theme, k)}>
              <span className="template-thumb deck" data-kind={k}><i /><i /><i /></span>
              <strong>{l}</strong><small>{d}</small>
            </button>
          ))}
          <button className="template-card ai" onClick={() => { setNewOpen(false); setAiOpen('outline'); }}>
            <span className="template-thumb ai"><Icon name="ai" size={22} /></span>
            <strong>AI outline</strong><small>Describe your talk</small>
          </button>
        </div>
      </BottomSheet>

      <BottomSheet open={outlineOpen} onClose={() => setOutlineOpen(false)} title="Outline" tall>
        <div className="outline-list">
          {slides.map((s, i) => { const { title: t, body } = slideText(s); return (
            <button key={s.id ?? i} className={`outline-item lvl1${i === sel ? ' on' : ''}`} onClick={() => { setSel(i); setOutlineOpen(false); setSorter(false); }}>
              <b>{i + 1}. {t || '(untitled)'}</b>
              {body.slice(0, 4).map((b, j) => <small key={j}>• {b}</small>)}
            </button>
          ); })}
        </div>
      </BottomSheet>

      <BottomSheet open={notesOpen} onClose={() => { editSlide({ notes: notesDraft.trim() || undefined }); setNotesOpen(false); }} title={`Speaker notes · slide ${sel + 1}`}>
        <textarea className="input" rows={6} value={notesDraft} placeholder="What you'll say for this slide…" onChange={(e) => setNotesDraft(e.target.value)} autoFocus />
        <div className="btn-row end">
          <button className="btn small" onClick={() => { setNotesOpen(false); setAiOpen('notes'); }}><Icon name="ai" size={14} /> Write with AI</button>
          <button className="btn primary" onClick={() => { editSlide({ notes: notesDraft.trim() || undefined }); setNotesOpen(false); }}>Done</button>
        </div>
      </BottomSheet>

      <BottomSheet open={textEdit !== null} onClose={saveTextEdit} title="Edit text">
        <textarea className="input slide-text-edit" rows={5} value={textDraft} onChange={(e) => setTextDraft(e.target.value)} autoFocus placeholder="One paragraph per line" />
        <div className="btn-row">
          <button className="btn small" onClick={() => { if (textEdit) { editShape(textEdit, (sh) => formatParas(sh, { bullet: !sh.paras[0]?.bullet })); } }}><Icon name="listBullet" size={14} /> Bullets</button>
          <button className="btn small" onClick={() => setAiOpen('rewrite')}><Icon name="ai" size={14} /> Rewrite</button>
          <span className="spacer" />
          <button className="btn primary" onClick={saveTextEdit}>Done</button>
        </div>
      </BottomSheet>

      <BottomSheet open={aiOpen !== null} onClose={() => !aiBusy && setAiOpen(null)} title={aiOpen === 'outline' ? 'AI: create slides' : aiOpen === 'rewrite' ? 'AI: rewrite text' : 'AI: speaker notes'}>
        {aiOpen === 'outline' && (
          <>
            <textarea className="input" rows={3} value={aiTopic} placeholder="e.g. Quarterly sales review for the leadership team, upbeat tone" onChange={(e) => setAiTopic(e.target.value)} autoFocus />
            <div className="rpanel-row"><span className="hint">Slides</span><RStepper value={String(aiCount)} title="Slide count" onDec={() => setAiCount((c) => Math.max(3, c - 1))} onInc={() => setAiCount((c) => Math.min(15, c + 1))} /></div>
            {slides.length > 0 && <p className="hint">New slides are appended after the current deck.</p>}
          </>
        )}
        {aiOpen === 'rewrite' && <textarea className="input" rows={2} value={aiTopic} placeholder="Optional instruction, e.g. shorter, more formal, translate to Hindi" onChange={(e) => setAiTopic(e.target.value)} autoFocus />}
        {aiOpen === 'notes' && <textarea className="input" rows={2} value={aiTopic} placeholder="Optional: audience or context" onChange={(e) => setAiTopic(e.target.value)} autoFocus />}
        <div className="btn-row end">
          <button className="btn" disabled={aiBusy} onClick={() => setAiOpen(null)}>Cancel</button>
          <button className="btn primary" disabled={aiBusy || (aiOpen === 'outline' && !aiTopic.trim())} onClick={() => void runAi()}>{aiBusy ? 'Working…' : 'Generate'}</button>
        </div>
      </BottomSheet>

      <ConfirmSheet open={confirmDel} title={`Delete slide ${sel + 1}?`} message="You can undo this with the Undo button right after." onConfirm={() => removeSlide(sel)} onClose={() => setConfirmDel(false)} />
      <Toast msg={toast} />
    </div>
  );
}

/** Theme swatch with tap (this slide) / long-press (all slides). */
function ThemeChip({ t, active, big, onTap, onHold }: { t: Theme; active: boolean; big?: boolean; onTap: () => void; onHold?: () => void }) {
  const timer = useRef<ReturnType<typeof setTimeout> | null>(null);
  const held = useRef(false);
  const start = useCallback(() => { held.current = false; if (onHold) timer.current = setTimeout(() => { held.current = true; onHold(); void tap('medium'); }, 550); }, [onHold]);
  const end = useCallback(() => { if (timer.current) clearTimeout(timer.current); }, []);
  return (
    <button
      className={`theme-chip${active ? ' on' : ''}${big ? ' big' : ''}`}
      style={{ background: t.bg, color: t.text, borderColor: active ? t.accent : undefined }}
      onPointerDown={start}
      onPointerUp={end}
      onPointerLeave={end}
      onClick={() => { if (!held.current) onTap(); }}
      aria-label={`${t.name} theme`}
      title={t.name}
    >
      <span className="theme-chip-bar" style={{ background: t.accent }} />
      <span className="theme-chip-name">{t.name}</span>
    </button>
  );
}
