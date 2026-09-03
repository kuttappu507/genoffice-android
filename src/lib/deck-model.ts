/**
 * Presentation model helpers. Every slide is a list of positioned shapes on a
 * 10 × 5.63 inch canvas (16:9) — legacy "title + bullets" slides are converted
 * to shapes on load so the editor, filmstrip, presenter and exporters all
 * share one renderer.
 */
import type { DeckSlide, ShapePara, ShapeRun, SlideShape } from './fileio';

export type { DeckSlide, ShapePara, ShapeRun, SlideShape };

export const CW = 10;
export const CH = 5.63;

export type Transition = 'none' | 'fade' | 'push' | 'zoom' | 'flip';
export type LayoutId = 'title' | 'content' | 'section' | 'two' | 'quote' | 'image' | 'blank';

export interface Theme {
  name: string;
  bg: string;
  accent: string;
  text: string;
  sub: string;
  font?: string;
  /** decorative band style */
  deco?: 'bar' | 'side' | 'corner' | 'none';
}

export const THEMES: Theme[] = [
  { name: 'Classic', bg: '#FFFFFF', accent: '#C43E1C', text: '#1F2430', sub: '#5B6270', deco: 'bar' },
  { name: 'Ivory', bg: '#FDF6EC', accent: '#B7791F', text: '#3B2F1A', sub: '#7A6A4F', font: 'Georgia', deco: 'side' },
  { name: 'Forest', bg: '#F0F7F0', accent: '#107C41', text: '#12301C', sub: '#4E6B57', deco: 'bar' },
  { name: 'Ocean', bg: '#EAF3FB', accent: '#185ABD', text: '#0F2543', sub: '#4B607F', deco: 'corner' },
  { name: 'Plum', bg: '#F6EEF9', accent: '#7030A0', text: '#2E1A3B', sub: '#6A5678', deco: 'side' },
  { name: 'Ink', bg: '#1E2430', accent: '#7EB8DA', text: '#F2F4F7', sub: '#B7C0CE', deco: 'bar' },
  { name: 'Midnight', bg: '#0F172A', accent: '#F59E0B', text: '#F8FAFC', sub: '#CBD5E1', deco: 'corner' },
  { name: 'Coral', bg: '#FFF1EE', accent: '#E0503C', text: '#3A1F1B', sub: '#7D5A54', deco: 'side' },
  { name: 'Slate', bg: '#F3F4F6', accent: '#374151', text: '#111827', sub: '#4B5563', deco: 'none' },
  { name: 'Mint', bg: '#ECFDF5', accent: '#059669', text: '#064E3B', sub: '#3F6F60', deco: 'corner' },
];

export const LAYOUTS: { id: LayoutId; icon: string; label: string }[] = [
  { id: 'title', icon: 'layoutTitle', label: 'Title' },
  { id: 'content', icon: 'layoutContent', label: 'Title & content' },
  { id: 'two', icon: 'columns', label: 'Two columns' },
  { id: 'section', icon: 'layoutSection', label: 'Section header' },
  { id: 'quote', icon: 'quote', label: 'Quote' },
  { id: 'image', icon: 'image', label: 'Picture with caption' },
  { id: 'blank', icon: 'layoutBlank', label: 'Blank' },
];

export const TRANSITIONS: { id: Transition; label: string }[] = [
  { id: 'none', label: 'None' },
  { id: 'fade', label: 'Fade' },
  { id: 'push', label: 'Push' },
  { id: 'zoom', label: 'Zoom' },
  { id: 'flip', label: 'Flip' },
];

export function isDark(hex?: string): boolean {
  const m = /^#([0-9a-fA-F]{6})$/.exec(hex ?? '');
  if (!m) return false;
  const v = parseInt(m[1], 16);
  return 0.299 * ((v >> 16) & 255) + 0.587 * ((v >> 8) & 255) + 0.114 * (v & 255) < 140;
}

export function themeOf(slide: DeckSlide): Theme {
  return THEMES.find((t) => t.bg.toLowerCase() === (slide.bg ?? '').toLowerCase() && t.accent.toLowerCase() === (slide.accent ?? '').toLowerCase()) ?? {
    ...THEMES[0],
    bg: slide.bg ?? THEMES[0].bg,
    accent: slide.accent ?? THEMES[0].accent,
    text: isDark(slide.bg) ? '#F2F4F7' : '#1F2430',
    sub: isDark(slide.bg) ? '#B7C0CE' : '#5B6270',
  };
}

let seq = 0;
export function shapeId(): string {
  return `s${Date.now().toString(36)}${(seq++).toString(36)}`;
}

export function para(text: string, opts: Partial<ShapeRun> & { align?: ShapePara['align']; bullet?: boolean; level?: number } = {}): ShapePara {
  const { align, bullet, level, ...run } = opts;
  return { align, bullet, level, runs: [{ text, ...run }] };
}

export function textBox(x: number, y: number, w: number, h: number, paras: ShapePara[], extra: Partial<SlideShape> = {}): SlideShape {
  return { id: shapeId(), x, y, w, h, kind: 'text', paras, ...extra };
}

export function geomShape(geom: NonNullable<SlideShape['geom']>, x: number, y: number, w: number, h: number, fill: string, extra: Partial<SlideShape> = {}): SlideShape {
  return { id: shapeId(), x, y, w, h, kind: 'shape', geom, paras: [], fill, ...extra };
}

/** Decorative theme element (accent bar / side stripe / corner block). */
function decoration(t: Theme): SlideShape[] {
  switch (t.deco) {
    case 'side':
      return [geomShape('rect', 0, 0, 0.22, CH, t.accent, { locked: true, name: 'deco' })];
    case 'corner':
      return [geomShape('rect', CW - 1.6, 0, 1.6, 0.22, t.accent, { locked: true, name: 'deco' }), geomShape('rect', CW - 0.22, 0, 0.22, 1.6, t.accent, { locked: true, name: 'deco' })];
    case 'bar':
      return [geomShape('rect', 0, CH - 0.16, CW, 0.16, t.accent, { locked: true, name: 'deco' })];
    default:
      return [];
  }
}

/** Build the shapes for a layout. */
export function layoutShapes(layout: LayoutId, t: Theme, content: { title?: string; bullets?: string[]; image?: string } = {}): SlideShape[] {
  const title = content.title ?? '';
  const bullets = (content.bullets ?? []).filter((b) => b.trim());
  const font = t.font;
  const bulletParas = (list: string[], sz = 18) => (list.length ? list : ['']).map((b) => para(b, { sz, color: t.text, bullet: true, font }));
  const deco = decoration(t);
  switch (layout) {
    case 'title':
      return [
        ...deco,
        textBox(0.7, 1.55, 8.6, 1.4, [para(title || 'Presentation title', { sz: 40, b: true, color: t.text, align: 'center', font })], { name: 'title', valign: 'middle' }),
        textBox(1.2, 3.05, 7.6, 0.8, [para(bullets[0] ?? 'Subtitle', { sz: 20, color: t.accent, align: 'center', font })], { name: 'subtitle', valign: 'top' }),
      ];
    case 'section':
      return [
        ...deco,
        geomShape('rect', 0, 2.25, CW, 1.15, t.accent, { name: 'band' }),
        textBox(0.6, 2.3, 8.8, 1.05, [para(title || 'Section title', { sz: 30, b: true, color: '#FFFFFF', align: 'center', font })], { name: 'title', valign: 'middle' }),
        textBox(1.2, 3.55, 7.6, 0.6, [para(bullets[0] ?? '', { sz: 16, color: t.sub, align: 'center', font })], { name: 'subtitle' }),
      ];
    case 'two':
      return [
        ...deco,
        textBox(0.6, 0.35, 8.8, 0.95, [para(title || 'Slide title', { sz: 28, b: true, color: t.text, font })], { name: 'title', valign: 'middle' }),
        geomShape('rect', 0.62, 1.32, 1.4, 0.06, t.accent, { name: 'rule' }),
        textBox(0.6, 1.6, 4.25, 3.6, bulletParas(bullets.slice(0, Math.ceil(bullets.length / 2)), 16), { name: 'body' }),
        textBox(5.15, 1.6, 4.25, 3.6, bulletParas(bullets.slice(Math.ceil(bullets.length / 2)), 16), { name: 'body2' }),
      ];
    case 'quote':
      return [
        ...deco,
        textBox(0.9, 0.9, 8.2, 2.6, [para(`“${title || 'A memorable quote goes here.'}”`, { sz: 28, i: true, color: t.text, align: 'center', font: font ?? 'Georgia' })], { name: 'quote', valign: 'middle' }),
        geomShape('rect', 4.4, 3.6, 1.2, 0.05, t.accent, { name: 'rule' }),
        textBox(1.5, 3.8, 7, 0.6, [para(bullets[0] ?? '— Attribution', { sz: 16, color: t.sub, align: 'center', font })], { name: 'attribution' }),
      ];
    case 'image':
      return [
        ...deco,
        content.image
          ? { id: shapeId(), x: 0.6, y: 0.45, w: 8.8, h: 3.7, kind: 'image', paras: [], img: content.image, name: 'picture' }
          : geomShape('rect', 0.6, 0.45, 8.8, 3.7, isDark(t.bg) ? '#2A3140' : '#E9ECF1', { name: 'picture-placeholder', paras: [para('Tap Insert → Picture', { sz: 16, color: t.sub, align: 'center' })], kind: 'text', valign: 'middle' }),
        textBox(0.6, 4.3, 8.8, 0.9, [para(title || 'Caption', { sz: 18, color: t.text, align: 'center', font })], { name: 'caption', valign: 'top' }),
      ];
    case 'blank':
      return [...deco];
    case 'content':
    default:
      return [
        ...deco,
        textBox(0.6, 0.35, 8.8, 0.95, [para(title || 'Slide title', { sz: 28, b: true, color: t.text, font })], { name: 'title', valign: 'middle' }),
        geomShape('rect', 0.62, 1.32, 1.6, 0.06, t.accent, { name: 'rule' }),
        textBox(0.6, 1.6, content.image ? 4.9 : 8.8, 3.6, bulletParas(bullets), { name: 'body' }),
        ...(content.image ? [{ id: shapeId(), x: 5.85, y: 1.75, w: 3.6, h: 2.9, kind: 'image' as const, paras: [], img: content.image, name: 'picture' }] : []),
      ];
  }
}

export function newSlide(layout: LayoutId, t: Theme, content: { title?: string; bullets?: string[]; image?: string } = {}): DeckSlide {
  const shapes = layoutShapes(layout, t, content);
  return {
    id: shapeId(),
    title: content.title ?? '',
    bullets: content.bullets ?? [],
    layout: layout === 'two' || layout === 'quote' || layout === 'image' ? 'content' : layout,
    bg: t.bg,
    accent: t.accent,
    shapes,
    cw: CW,
    ch: CH,
    transition: 'none',
  };
}

/** Upgrade legacy slides (title + bullets, no shapes) so everything is shape-based. */
export function normalizeSlide(s: DeckSlide): DeckSlide {
  const withId = s.id ? s : { ...s, id: shapeId() };
  if (withId.shapes && withId.shapes.length > 0) {
    const shapes = withId.shapes.map((sh) => (sh.id ? sh : { ...sh, id: shapeId() }));
    return { ...withId, shapes, cw: withId.cw ?? CW, ch: withId.ch ?? CH };
  }
  const t = themeOf(withId);
  const layout: LayoutId = (withId.layout as LayoutId) ?? 'content';
  return { ...withId, shapes: layoutShapes(layout, t, { title: withId.title, bullets: withId.bullets, image: withId.image }), image: undefined, cw: CW, ch: CH };
}

export function normalizeDeck(slides: DeckSlide[]): DeckSlide[] {
  return slides.map(normalizeSlide);
}

export function shapeText(sh: SlideShape): string {
  return sh.paras.map((p) => p.runs.map((r) => r.text).join('')).join('\n');
}

/** Plain-text view of a slide (title line + body lines) — used for AI context, outline and search. */
export function slideText(s: DeckSlide): { title: string; body: string[] } {
  const texts = (s.shapes ?? []).filter((sh) => sh.kind === 'text' && shapeText(sh).trim());
  const titleShape = texts.find((sh) => sh.name === 'title' || sh.name === 'quote') ?? texts[0];
  const title = titleShape ? shapeText(titleShape).split('\n')[0] : s.title;
  const body = texts.filter((sh) => sh !== titleShape).flatMap((sh) => shapeText(sh).split('\n')).filter((l) => l.trim());
  return { title: title || s.title || '', body };
}

/** Replace the text of a shape from plain lines, keeping paragraph/run formatting of the first paragraph. */
export function setShapeText(sh: SlideShape, text: string): SlideShape {
  const lines = text.replace(/\r/g, '').split('\n');
  const template = sh.paras[0];
  const paras: ShapePara[] = lines.map((ln, i) => {
    const old = sh.paras[i] ?? template;
    const base = old?.runs.find((r) => r.text.trim()) ?? old?.runs[0] ?? { sz: 18 };
    return { align: old?.align, bullet: old?.bullet, level: old?.level, runs: [{ ...base, text: ln }] };
  });
  return { ...sh, paras: paras.length ? paras : [para('', { sz: 18 })] };
}

/** Apply run formatting to every run of a shape (font size step, bold, color…). */
export function formatShape(sh: SlideShape, patch: Partial<ShapeRun> | ((r: ShapeRun) => ShapeRun)): SlideShape {
  const fn = typeof patch === 'function' ? patch : (r: ShapeRun) => ({ ...r, ...patch });
  return { ...sh, paras: sh.paras.map((p) => ({ ...p, runs: p.runs.map(fn) })) };
}

export function formatParas(sh: SlideShape, patch: Partial<Omit<ShapePara, 'runs'>>): SlideShape {
  return { ...sh, paras: sh.paras.map((p) => ({ ...p, ...patch })) };
}

export function firstRun(sh: SlideShape | undefined): ShapeRun | undefined {
  if (!sh) return undefined;
  for (const p of sh.paras) for (const r of p.runs) if (r.text.trim()) return r;
  return sh.paras[0]?.runs[0];
}

export function cloneSlide(s: DeckSlide): DeckSlide {
  const c = JSON.parse(JSON.stringify(s)) as DeckSlide;
  c.id = shapeId();
  c.shapes = (c.shapes ?? []).map((sh) => ({ ...sh, id: shapeId() }));
  return c;
}

export function applyThemeToSlide(s: DeckSlide, t: Theme): DeckSlide {
  const old = themeOf(s);
  const shapes = (s.shapes ?? [])
    .filter((sh) => sh.name !== 'deco')
    .map((sh) => {
      let next = sh;
      if (sh.fill && sh.fill.toLowerCase() === old.accent.toLowerCase()) next = { ...next, fill: t.accent };
      next = formatShape(next, (r) => {
        const c = (r.color ?? '').toLowerCase();
        if (c === old.accent.toLowerCase()) return { ...r, color: t.accent };
        if (c === old.text.toLowerCase() || (!r.color && sh.name !== 'title' && sh.name !== 'band')) return { ...r, color: t.text };
        if (c === old.sub.toLowerCase()) return { ...r, color: t.sub };
        return r;
      });
      return next;
    });
  return { ...s, bg: t.bg, accent: t.accent, shapes: [...decoration(t), ...shapes] };
}

/** Parse an AI outline ("### Title" + "- bullet" lines; also accepts "Slide N:" / markdown headings). */
export function parseOutline(out: string): { title: string; bullets: string[]; notes?: string }[] {
  const slides: { title: string; bullets: string[]; notes?: string }[] = [];
  let cur: { title: string; bullets: string[]; notes?: string } | null = null;
  for (const raw of out.split('\n')) {
    const t = raw.trim();
    if (!t) continue;
    const h = /^(?:#{1,3}\s+|slide\s*\d+\s*[:.-]\s*)(.+)$/i.exec(t);
    if (h) { if (cur) slides.push(cur); cur = { title: h[1].replace(/\*\*/g, '').trim(), bullets: [] }; continue; }
    const n = /^notes?\s*:\s*(.+)$/i.exec(t);
    if (n && cur) { cur.notes = n[1]; continue; }
    const b = /^(?:[-*•]|\d+[.)])\s+(.+)$/.exec(t);
    if (b && cur) { cur.bullets.push(b[1].replace(/\*\*/g, '').trim()); continue; }
    if (cur && cur.bullets.length === 0 && !cur.title) cur.title = t;
    else if (cur) cur.bullets.push(t.replace(/\*\*/g, ''));
  }
  if (cur) slides.push(cur);
  return slides.filter((s) => s.title || s.bullets.length);
}

export function deckFromOutline(outline: { title: string; bullets: string[]; notes?: string }[], t: Theme, deckTitle?: string): DeckSlide[] {
  const slides: DeckSlide[] = [];
  outline.forEach((o, i) => {
    const layout: LayoutId = i === 0 && o.bullets.length <= 1 ? 'title' : o.bullets.length === 0 ? 'section' : o.bullets.length > 6 ? 'two' : 'content';
    const s = newSlide(layout, t, { title: o.title || deckTitle, bullets: o.bullets });
    s.notes = o.notes;
    slides.push(s);
  });
  return slides;
}
