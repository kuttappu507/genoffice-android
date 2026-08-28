import { Capacitor } from '@capacitor/core';
import { Directory, Filesystem } from '@capacitor/filesystem';
import type { CellObject, WorkSheet } from 'xlsx';
import type * as DocxNs from 'docx';

export interface PickedFile {
  name: string;
  buf: ArrayBuffer;
}

/** System file picker. Works on Android WebView (Capacitor) and desktop browsers. */
export function openFilePicker(accept: string): Promise<PickedFile | null> {
  return new Promise((resolve) => {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = accept;
    input.style.display = 'none';
    document.body.appendChild(input);
    input.onchange = () => {
      const f = input.files?.[0];
      if (!f) {
        resolve(null);
        return;
      }
      f.arrayBuffer()
        .then((buf) => resolve({ name: f.name, buf }))
        .catch(() => resolve(null))
        .finally(() => input.remove());
    };
    input.oncancel = () => {
      input.remove();
      resolve(null);
    };
    input.click();
  });
}

function bytesToBase64(bytes: Uint8Array): string {
  let bin = '';
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) {
    bin += String.fromCharCode(...bytes.subarray(i, i + chunk));
  }
  return btoa(bin);
}

/**
 * Save raw bytes on the device.
 * Native: writes to the app cache, then opens the Android share sheet so the
 * user can save it anywhere (Files, Drive) or send it to another app.
 * Web: triggers a browser download.
 */
export async function saveBinary(
  name: string,
  data: Blob | ArrayBuffer | Uint8Array,
  mime: string,
  opts: { share?: boolean } = {},
): Promise<string> {
  let bytes: Uint8Array;
  if (data instanceof Blob) {
    bytes = new Uint8Array(await data.arrayBuffer());
  } else if (data instanceof Uint8Array) {
    bytes = data;
  } else {
    bytes = new Uint8Array(data);
  }
  if (Capacitor.isNativePlatform()) {
    const res = await Filesystem.writeFile({
      path: name,
      data: bytesToBase64(bytes),
      directory: Directory.Cache,
    });
    if (opts.share !== false) {
      try {
        const { Share } = await import('@capacitor/share');
        await Share.share({ title: name, url: res.uri, dialogTitle: 'Save or send file' });
        return `Share sheet opened for ${name}`;
      } catch {
        return `Saved to app storage: ${name}`;
      }
    }
    return `Saved to app storage: ${name}`;
  }
  const url = URL.createObjectURL(new Blob([bytes as BlobPart], { type: mime }));
  const a = document.createElement('a');
  a.href = url;
  a.download = name;
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 5000);
  return `Downloaded ${name}`;
}

export function sanitizeName(title: string, ext: string): string {
  const base = title.replace(/[^\w-]+/g, '_').replace(/^_+|_+$/g, '').slice(0, 60);
  return `${base || 'file'}.${ext}`;
}

export function baseName(fileName: string): string {
  return fileName.replace(/\.[^.]+$/, '') || fileName;
}

// ---------------------------------------------------------------------------
// Word (.docx)
// ---------------------------------------------------------------------------

/** .docx -> HTML for the Docs editor (mammoth). */
export async function importDocx(buf: ArrayBuffer): Promise<string> {
  const mammoth = await import('mammoth');
  const res = await mammoth.convertToHtml({ arrayBuffer: buf });
  return res.value;
}

/** Plain text / markdown -> simple HTML paragraphs for the Docs editor. */
export function textToHtml(text: string): string {
  const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  return text
    .split(/\n\s*\n/)
    .map((p) => {
      const trimmed = p.trim();
      if (!trimmed) return '';
      const h = /^(#{1,3})\s+(.*)$/.exec(trimmed.split('\n')[0]);
      if (h) {
        const level = Math.min(h[1].length + 1, 4); // # -> h2, ## -> h3
        const rest = trimmed.split('\n').slice(1).join('\n');
        const body = rest ? `<p>${esc(rest).replace(/\n/g, '<br>')}</p>` : '';
        return `<h${level}>${esc(h[2])}</h${level}>${body}`;
      }
      return `<p>${esc(trimmed).replace(/\n/g, '<br>')}</p>`;
    })
    .join('');
}

interface RunProps {
  bold?: boolean;
  italics?: boolean;
  underline?: boolean;
}

function docxRuns(node: Node, props: RunProps, D: typeof DocxNs): DocxNs.TextRun[] {
  if (node.nodeType === 3) {
    const text = node.textContent ?? '';
    if (!text) return [];
    return [
      new D.TextRun({
        text,
        bold: props.bold,
        italics: props.italics,
        underline: props.underline ? {} : undefined,
      }),
    ];
  }
  const el = node as Element;
  const tag = el.tagName?.toLowerCase() ?? '';
  const p: RunProps = { ...props };
  if (tag === 'b' || tag === 'strong') p.bold = true;
  if (tag === 'i' || tag === 'em') p.italics = true;
  if (tag === 'u') p.underline = true;
  if (tag === 'br') return [new D.TextRun({ text: '', break: 1 })];
  return Array.from(el.childNodes).flatMap((c) => docxRuns(c, p, D));
}

/** Docs editor HTML -> real .docx file (docx library). */
export async function exportDocx(title: string, html: string): Promise<Blob> {
  const D = await import('docx');
  const doc = new DOMParser().parseFromString(`<body>${html}</body>`, 'text/html');
  const body = doc.body;
  const paras: DocxNs.Paragraph[] = [];
  paras.push(new D.Paragraph({ text: title, heading: D.HeadingLevel.TITLE }));

  const blocks = Array.from(body.children);
  const emitBlock = (b: Element) => {
    const tag = b.tagName.toLowerCase();
    if (tag === 'ul' || tag === 'ol') {
      Array.from(b.children).forEach((li) => {
        paras.push(new D.Paragraph({ children: docxRuns(li, {}, D), bullet: { level: 0 } }));
      });
    } else if (tag === 'h1') {
      paras.push(new D.Paragraph({ heading: D.HeadingLevel.HEADING_1, children: docxRuns(b, {}, D) }));
    } else if (tag === 'h2') {
      paras.push(new D.Paragraph({ heading: D.HeadingLevel.HEADING_2, children: docxRuns(b, {}, D) }));
    } else if (tag === 'h3') {
      paras.push(new D.Paragraph({ heading: D.HeadingLevel.HEADING_3, children: docxRuns(b, {}, D) }));
    } else if (b.textContent?.trim()) {
      paras.push(new D.Paragraph({ children: docxRuns(b, {}, D) }));
    }
  };
  if (blocks.length === 0) {
    if (body.textContent?.trim()) {
      paras.push(new D.Paragraph({ children: docxRuns(body, {}, D) }));
    }
  } else {
    blocks.forEach(emitBlock);
  }
  return D.Packer.toBlob(new D.Document({ sections: [{ children: paras }] }));
}

// ---------------------------------------------------------------------------
// Spreadsheets (.xlsx / .xls / .csv)
// ---------------------------------------------------------------------------

const MAX_ROWS = 60;
const MAX_COLS = 26;

/** First sheet of an Excel/CSV file -> cell map for the Sheets grid. Formulas kept. */
export async function importSpreadsheet(buf: ArrayBuffer): Promise<{ cells: Record<string, string> }> {
  const X = await import('xlsx');
  const wb = X.read(buf, { type: 'array', cellFormula: true, cellText: false });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const cells: Record<string, string> = {};
  if (!ws || !ws['!ref']) return { cells };
  const range = X.utils.decode_range(ws['!ref']);
  for (let r = range.s.r; r <= Math.min(range.e.r, MAX_ROWS - 1); r++) {
    for (let c = range.s.c; c <= Math.min(range.e.c, MAX_COLS - 1); c++) {
      const addr = X.utils.encode_cell({ r, c });
      const cell = ws[addr] as CellObject | undefined;
      if (!cell) continue;
      let raw: string;
      if (cell.f) raw = `=${cell.f}`;
      else if (cell.v !== undefined && cell.v !== null) raw = String(cell.v);
      else continue;
      if (raw !== '') cells[addr] = raw;
    }
  }
  return { cells };
}

/** Cells grid -> .xlsx with live formulas (values cached via the local engine). */
export async function exportXlsx(title: string, cells: Record<string, string>): Promise<Uint8Array> {
  const X = await import('xlsx');
  const { evalCell } = await import('./formulas');
  const ws: WorkSheet = {};
  let maxR = 0;
  let maxC = 0;
  for (const key of Object.keys(cells)) {
    const raw = cells[key];
    if (!raw) continue;
    const m = /^([A-Z])([0-9]{1,3})$/.exec(key);
    if (!m) continue;
    const c = m[1].charCodeAt(0) - 65;
    const r = parseInt(m[2], 10) - 1;
    maxR = Math.max(maxR, r);
    maxC = Math.max(maxC, c);
    if (raw.startsWith('=')) {
      const shown = evalCell(raw, (ref) => cells[ref] ?? '');
      const num = parseFloat(shown);
      const isNum = shown !== '#ERR' && Number.isFinite(num) && shown.trim() !== '';
      ws[key] = isNum
        ? { t: 'n', v: num, f: raw.slice(1) }
        : { t: 's', v: shown, f: raw.slice(1) };
    } else {
      const num = parseFloat(raw);
      ws[key] = Number.isFinite(num) && raw.trim() !== '' ? { t: 'n', v: num } : { t: 's', v: raw };
    }
  }
  ws['!ref'] = X.utils.encode_range({ s: { c: 0, r: 0 }, e: { c: Math.max(maxC, 0), r: Math.max(maxR, 0) } });
  const wb = X.utils.book_new();
  const sheetName = (title || 'Sheet').replace(/[\\/?*[\]:]/g, ' ').slice(0, 31) || 'Sheet';
  X.utils.book_append_sheet(wb, ws, sheetName);
  const out = X.write(wb, { bookType: 'xlsx', type: 'array' }) as ArrayBuffer;
  return new Uint8Array(out);
}

// ---------------------------------------------------------------------------
// Images
// ---------------------------------------------------------------------------

/** Pick an image from the device and return a downscaled JPEG data URL. */
export async function pickImage(maxDim = 1280): Promise<string | null> {
  const pick = await openFilePicker('image/*');
  if (!pick) return null;
  const url = URL.createObjectURL(new Blob([pick.buf as BlobPart]));
  try {
    const img = await new Promise<HTMLImageElement>((res, rej) => {
      const i = new Image();
      i.onload = () => res(i);
      i.onerror = () => rej(new Error('Not a valid image'));
      i.src = url;
    });
    const scale = Math.min(1, maxDim / Math.max(img.width, img.height));
    const w = Math.max(1, Math.round(img.width * scale));
    const h = Math.max(1, Math.round(img.height * scale));
    const cv = document.createElement('canvas');
    cv.width = w;
    cv.height = h;
    const ctx = cv.getContext('2d');
    if (!ctx) throw new Error('Canvas is unavailable');
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, w, h);
    ctx.drawImage(img, 0, 0, w, h);
    return cv.toDataURL('image/jpeg', 0.85);
  } finally {
    URL.revokeObjectURL(url);
  }
}

/** Rough luminance test so text stays readable on colored slide backgrounds. */
export function isDarkColor(hex?: string): boolean {
  if (!hex || !/^#[0-9a-fA-F]{6}$/.test(hex)) return false;
  const r = parseInt(hex.slice(1, 3), 16);
  const g = parseInt(hex.slice(3, 5), 16);
  const b = parseInt(hex.slice(5, 7), 16);
  return 0.299 * r + 0.587 * g + 0.114 * b < 140;
}

// ---------------------------------------------------------------------------
// PowerPoint (.pptx)
// ---------------------------------------------------------------------------

export interface DeckSlide {
  title: string;
  bullets: string[];
  /** slide background color (#rrggbb) */
  bg?: string;
  /** accent color for bars / highlights (#rrggbb) */
  accent?: string;
  /** image placed on the slide (data URL) */
  image?: string;
  layout?: 'title' | 'content' | 'section' | 'blank';
  /** faithful rendering of an imported slide: positioned shapes */
  shapes?: SlideShape[];
  /** slide canvas size in inches (defaults 10 x 5.63) */
  cw?: number;
  ch?: number;
}

/** One styled text run inside a shape paragraph. */
export interface ShapeRun {
  text: string;
  /** font size in points */
  sz?: number;
  b?: boolean;
  i?: boolean;
  u?: boolean;
  color?: string;
}

/** One paragraph inside a text shape. */
export interface ShapePara {
  runs: ShapeRun[];
  align?: 'left' | 'center' | 'right' | 'justify';
  bullet?: boolean;
  /** 0-based indent level */
  level?: number;
}

/** A positioned shape on the slide canvas (geometry in inches). */
export interface SlideShape {
  x: number;
  y: number;
  w: number;
  h: number;
  kind: 'text' | 'image' | 'shape';
  paras: ShapePara[];
  fill?: string;
  line?: string;
  img?: string;
}

function decodeXmlEntities(s: string): string {
  return s
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#x([0-9a-fA-F]+);/g, (_, h) => String.fromCharCode(parseInt(h, 16)))
    .replace(/&#(\d+);/g, (_, d) => String.fromCharCode(parseInt(d, 10)))
    .replace(/&amp;/g, '&');
}

// ---------------------------------------------------------------------------
// PowerPoint (.pptx) — shape-level import/export for PC-grade fidelity
// ---------------------------------------------------------------------------

const SCHEME_COLORS: Record<string, string> = {
  dk1: '000000', tx1: '000000', lt1: 'FFFFFF', bg1: 'FFFFFF',
  dk2: '44546A', tx2: '44546A', lt2: 'E7E6E6', bg2: 'E7E6E6',
  accent1: '4472C4', accent2: 'ED7D31', accent3: 'A5A5A5', accent4: 'FFC000',
  accent5: '5B9BD5', accent6: '70AD47', hlink: '0563C1', folHlink: '954F72',
  phClr: '808080',
};

const IMG_MIME: Record<string, string> = {
  png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif',
  bmp: 'image/bmp', webp: 'image/webp', svg: 'image/svg+xml',
};

const EMU_IN = 914400;
/** Our standard canvas is 10 x 5.63 in (WIDE layout). */
const CANVAS_W = 10;
const CANVAS_H = 5.63;

function attr(tag: string, name: string): string | undefined {
  const m = new RegExp(`\\b${name}="([^"]*)"`).exec(tag);
  return m ? m[1] : undefined;
}

/** First explicit color inside an XML fragment (srgbClr, schemeClr or sysClr). */
function firstColorIn(xml: string): string | undefined {
  const m =
    /<a:solidFill>\s*<a:srgbClr\s+val="([0-9A-Fa-f]{6})"/.exec(xml) ??
    /<a:solidFill>\s*<a:schemeClr\s+val="([a-zA-Z0-9]+)"/.exec(xml);
  if (m) {
    const v = SCHEME_COLORS[m[1]] ?? m[1];
    return `#${v}`;
  }
  const sys = /<a:sysClr[^>]*lastClr="([0-9A-Fa-f]{6})"/.exec(xml);
  return sys ? `#${sys[1]}` : undefined;
}

interface Geom {
  x: number;
  y: number;
  w: number;
  h: number;
}

/** a:xfrm -> inches on our canvas, scaled from the source slide size. */
function parseXfrm(block: string, scale: number): Geom | null {
  const xfrm = /<a:xfrm[^>]*>([\s\S]*?)<\/a:xfrm>/.exec(block);
  if (!xfrm) return null;
  const offTag = /<a:off\b[^>]*\/?>/.exec(xfrm[1])?.[0] ?? '';
  const extTag = /<a:ext\b[^>]*\/?>/.exec(xfrm[1])?.[0] ?? '';
  const x = attr(offTag, 'x');
  const y = attr(offTag, 'y');
  const cx = attr(extTag, 'cx');
  const cy = attr(extTag, 'cy');
  if (x === undefined || y === undefined || cx === undefined || cy === undefined) return null;
  return {
    x: (parseInt(x, 10) / EMU_IN) * scale,
    y: (parseInt(y, 10) / EMU_IN) * scale,
    w: (parseInt(cx, 10) / EMU_IN) * scale,
    h: (parseInt(cy, 10) / EMU_IN) * scale,
  };
}

/** Fallback geometry for placeholders that inherit position from the layout. */
function placeholderGeom(spXml: string): Geom {
  const ph = /<p:ph\b([^>]*)\/?>/.exec(spXml)?.[1] ?? '';
  const type = attr(ph, 'type') ?? 'body';
  const idx = attr(ph, 'idx');
  if (type === 'title' || type === 'ctrTitle') return { x: 0.5, y: 0.35, w: 9, h: 1.15 };
  if (type === 'subTitle' || (type === 'body' && idx === '1')) return { x: 0.8, y: 1.75, w: 8.4, h: 3.4 };
  if (type === 'sldNum') return { x: 8.9, y: 5.2, w: 0.9, h: 0.35 };
  return { x: 0.8, y: 1.55, w: 8.4, h: 3.6 };
}

/** p:txBody -> styled paragraphs. */
function parseParas(txBody: string): ShapePara[] {
  const paras: ShapePara[] = [];
  const pBlocks = txBody.match(/<a:p(?:\s[^>]*)?>[\s\S]*?<\/a:p>/g) ?? [];
  for (const pb of pBlocks) {
    const pPrTag = /<a:pPr\b[^>]*?(?:\/>|>)/.exec(pb)?.[0] ?? '';
    const algn = attr(pPrTag, 'algn');
    const lvlS = attr(pPrTag, 'lvl');
    const bullet = /<a:buChar\b/.test(pb) || /<a:buAutoNum\b/.test(pb) ? true : /<a:buNone\b/.test(pb) ? false : undefined;
    const runs: ShapeRun[] = [];
    const rBlocks = pb.match(/<a:(?:r|fld)>[\s\S]*?<\/a:(?:r|fld)>/g) ?? [];
    for (const rb of rBlocks) {
      const t = /<a:t(?:\s[^>]*)?>([\s\S]*?)<\/a:t>/.exec(rb);
      const text = decodeXmlEntities(t ? t[1] : '');
      if (!text) continue;
      const rPrTag = /<a:rPr\b[^>]*?(?:\/>|>)/.exec(rb)?.[0] ?? '';
      const szS = attr(rPrTag, 'sz');
      const color =
        /<a:solidFill>\s*<a:srgbClr\s+val="([0-9A-Fa-f]{6})"/.exec(rb)?.[1] ??
        (() => {
          const sc = /<a:solidFill>\s*<a:schemeClr\s+val="([a-zA-Z0-9]+)"/.exec(rb)?.[1];
          return sc ? SCHEME_COLORS[sc] : undefined;
        })();
      runs.push({
        text,
        sz: szS ? parseInt(szS, 10) / 100 : undefined,
        b: /\bb="1"/.test(rPrTag) || undefined,
        i: /\bi="1"/.test(rPrTag) || undefined,
        u: /\bu="(?:sng|dbl)"/.test(rPrTag) || undefined,
        color: color ? `#${color}` : undefined,
      });
    }
    if (/<a:br\s*\/>/.test(pb)) runs.push({ text: '\n' });
    const hasText = runs.some((r) => r.text.trim());
    if (hasText || (runs.length === 0 && paras.length > 0)) {
      paras.push({
        runs,
        align:
          algn === 'ctr' ? 'center' : algn === 'r' ? 'right' : algn === 'just' ? 'justify' : algn === 'l' ? 'left' : undefined,
        bullet,
        level: lvlS ? Math.min(4, parseInt(lvlS, 10)) : 0,
      });
    }
  }
  while (paras.length && !paras[paras.length - 1].runs.some((r) => r.text.trim())) paras.pop();
  return paras;
}

function parseRels(relXml: string): Record<string, string> {
  const map: Record<string, string> = {};
  for (const tag of relXml.match(/<Relationship\b[^>]*>/g) ?? []) {
    const id = attr(tag, 'Id');
    const target = attr(tag, 'Target');
    if (id && target) map[id] = target;
  }
  return map;
}

/** Resolve an OPC part path relative to a base directory inside the package. */
function resolvePptPath(baseDir: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1);
  const parts = `${baseDir}/${target}`.split('/');
  const out: string[] = [];
  for (const p of parts) {
    if (p === '..') out.pop();
    else if (p && p !== '.') out.push(p);
  }
  return out.join('/');
}

type PptxZip = { files: Record<string, { async: (t: 'string' | 'base64') => Promise<string> }> };

async function mediaData(zip: PptxZip, target: string): Promise<string | undefined> {
  const path = resolvePptPath('ppt/slides', target);
  const f = zip.files[path];
  if (!f) return undefined;
  const ext = path.split('.').pop()?.toLowerCase() ?? '';
  const mime = IMG_MIME[ext];
  if (!mime) return undefined;
  const b64 = await f.async('base64');
  return `data:${mime};base64,${b64}`;
}

/** Background color of a slide XML, or null when it inherits. */
function slideBg(xml: string): string | null {
  const bg = /<p:bg>([\s\S]*?)<\/p:bg>/.exec(xml);
  return bg ? (firstColorIn(bg[1]) ?? null) : null;
}

/**
 * .pptx -> deck slides with faithful shape geometry: positioned text boxes,
 * fills, per-run formatting, pictures and slide backgrounds.
 * Falls back to plain title/bullets extraction when nothing parses.
 */
export async function importPptx(buf: ArrayBuffer): Promise<DeckSlide[]> {
  const JSZip = (await import('jszip')).default;
  const zip = await JSZip.loadAsync(buf);
  const names = Object.keys(zip.files)
    .filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))
    .sort((a, b) => {
      const na = parseInt(a.replace(/\D+/g, ''), 10);
      const nb = parseInt(b.replace(/\D+/g, ''), 10);
      return na - nb;
    });

  // Slide size -> uniform scale onto our 10 x 5.63 canvas.
  let srcW = 12192000;
  let srcH = 6858000;
  try {
    const presXml = await zip.files['ppt/presentation.xml']?.async('text');
    const sz = presXml ? /<p:sldSz\b[^>]*\/?>/.exec(presXml)?.[0] : undefined;
    if (sz) {
      srcW = parseInt(attr(sz, 'cx') ?? '', 10) || srcW;
      srcH = parseInt(attr(sz, 'cy') ?? '', 10) || srcH;
    }
  } catch {
    /* defaults fine */
  }
  const scale = Math.min(CANVAS_W / (srcW / EMU_IN), CANVAS_H / (srcH / EMU_IN));
  const cw = (srcW / EMU_IN) * scale;
  const ch = (srcH / EMU_IN) * scale;

  const slides: DeckSlide[] = [];
  for (const n of names) {
    const xml = await zip.files[n].async('text');
    const relPath = `ppt/slides/_rels/${n.split('/').pop()}.rels`;
    let rels: Record<string, string> = {};
    if (zip.files[relPath]) rels = parseRels(await zip.files[relPath].async('text'));

    let bg = slideBg(xml);
    if (!bg) {
      // inherit from slide layout, then its master
      const layoutTarget = Object.entries(rels).find(([, t]) => t.includes('slideLayout'))?.[1];
      if (layoutTarget) {
        const lp = resolvePptPath('ppt/slides', layoutTarget);
        const lxml = zip.files[lp] ? await zip.files[lp].async('text') : '';
        bg = lxml ? slideBg(lxml) : null;
        if (!bg && lxml) {
          const lRelPath = `${lp.replace('ppt/', 'ppt/_rels/')}.rels`;
          const lRels = zip.files[lRelPath] ? parseRels(await zip.files[lRelPath].async('text')) : {};
          const mTarget = Object.entries(lRels).find(([, t]) => t.includes('slideMaster'))?.[1];
          if (mTarget) {
            const mp = resolvePptPath('ppt/slideLayouts', mTarget);
            const mxml = zip.files[mp] ? await zip.files[mp].async('text') : '';
            bg = mxml ? slideBg(mxml) : null;
          }
        }
      }
    }

    const shapes: SlideShape[] = [];
    for (const sp of xml.match(/<p:sp>[\s\S]*?<\/p:sp>/g) ?? []) {
      const geom = parseXfrm(sp, scale) ?? placeholderGeom(sp);
      const spPr = /<p:spPr\b[^>]*>([\s\S]*?)<\/p:spPr>/.exec(sp)?.[1] ?? '';
      const fill = firstColorIn(spPr.replace(/<a:ln\b[\s\S]*$/, '')); // fill only, not outline
      const lineM = /<a:ln\b[^>]*>[\s\S]*?<a:solidFill>\s*<a:srgbClr\s+val="([0-9A-Fa-f]{6})"/.exec(spPr);
      const paras = parseParas(/<p:txBody\b[^>]*>([\s\S]*?)<\/p:txBody>/.exec(sp)?.[1] ?? '');
      const isText = paras.length > 0;
      const isPlaceholder = /<p:ph\b/.test(sp);
      if (!isText && !fill && !isPlaceholder) continue; // empty decorative box
      shapes.push({
        ...geom,
        kind: isText ? 'text' : 'shape',
        paras,
        fill,
        line: lineM ? `#${lineM[1]}` : undefined,
      });
    }
    for (const pic of xml.match(/<p:pic>[\s\S]*?<\/p:pic>/g) ?? []) {
      const geom = parseXfrm(pic, scale);
      const embed = /\br:(?:embed|link)="([^"]+)"/.exec(pic)?.[1];
      if (!geom || !embed) continue;
      const img = await mediaData(zip, rels[embed] ?? '');
      if (img) shapes.push({ ...geom, kind: 'image', paras: [], img });
    }

    // Title: first paragraph of the first text shape; bullets follow it.
    const firstText = shapes.find((s) => s.kind === 'text' && s.paras.some((p) => p.runs.some((r) => r.text.trim())));
    const firstLine = firstText?.paras.find((p) => p.runs.some((r) => r.text.trim()));
    const title = firstLine?.runs.map((r) => r.text).join('').trim() ?? '';
    const bullets = (firstText?.paras.slice(1) ?? [])
      .map((p) => p.runs.map((r) => r.text).join('').trim())
      .filter(Boolean)
      .slice(0, 12);

    slides.push({
      title: title || 'Slide',
      bullets: bullets.length ? bullets : [''],
      bg: bg ?? undefined,
      shapes,
      cw,
      ch,
    });
  }
  return slides;
}

/** Deck slides -> real .pptx file (pptxgenjs), honoring layouts, theme, shapes and images. */
export async function exportPptx(title: string, slides: DeckSlide[]): Promise<Uint8Array> {
  const PptxGenJS = (await import('pptxgenjs')).default;
  const pptx = new PptxGenJS();
  pptx.title = title;
  pptx.defineLayout({ name: 'WIDE', width: 10, height: 5.63 });
  pptx.layout = 'WIDE';
  for (const s of slides) {
    const slide = pptx.addSlide();
    const dark = isDarkColor(s.bg);
    if (s.bg) slide.background = { color: s.bg.replace('#', '') };

    // Faithful path: slides that carry imported/positioned shapes.
    if (s.shapes && s.shapes.length > 0) {
      const offX = Math.max(0, (10 - (s.cw ?? 10)) / 2);
      for (const sh of s.shapes) {
        const x = sh.x + offX;
        if (sh.kind === 'image' && sh.img) {
          slide.addImage({ data: sh.img, x, y: sh.y, w: sh.w, h: sh.h });
          continue;
        }
        const fill = sh.fill ? { color: sh.fill.replace('#', '') } : undefined;
        const items = sh.paras.flatMap((p) => {
          if (p.runs.length === 0) {
            return [{ text: ' ', options: { breakLine: true, fontSize: 8 } }];
          }
          return p.runs.map((r, ri) => ({
            text: r.text,
            options: {
              fontSize: r.sz ?? 14,
              bold: r.b,
              italic: r.i,
              underline: r.u ? { style: 'sng' as const } : undefined,
              color: (r.color ?? (dark ? 'FFFFFF' : '333333')).replace('#', ''),
              align: p.align,
              bullet: p.bullet ? { characterCode: '2022', indent: 12 + (p.level ?? 0) * 18 } : p.bullet === false ? false : undefined,
              indentLevel: p.level,
              breakLine: ri === p.runs.length - 1,
            },
          }));
        });
        if (items.length === 0) {
          if (fill) {
            slide.addShape(pptx.ShapeType.rect, {
              x, y: sh.y, w: sh.w, h: sh.h, fill,
              line: sh.line ? { color: sh.line.replace('#', '') } : undefined,
            });
          }
        } else {
          slide.addText(items as never, {
            x, y: sh.y, w: sh.w, h: sh.h,
            valign: 'top',
            fill,
            fontSize: 14,
            color: dark ? 'FFFFFF' : '333333',
          });
        }
      }
      continue;
    }

    const accent = (s.accent ?? '#C43E1C').replace('#', '');
    const main = dark ? 'FFFFFF' : '1F2430';
    const body = dark ? 'E8E8E8' : '333333';
    if (s.layout === 'title') {
      slide.addText(s.title || title, {
        x: 0.5, y: 2.1, w: 9, h: 1.3, fontSize: 34, bold: true, color: main, align: 'center',
      });
      const sub = s.bullets.find((b) => b.trim());
      if (sub) {
        slide.addText(sub, { x: 0.5, y: 3.45, w: 9, h: 0.7, fontSize: 17, color: accent, align: 'center' });
      }
    } else if (s.layout === 'section') {
      slide.addShape(pptx.ShapeType.rect, { x: 0, y: 2.3, w: 10, h: 1.05, fill: { color: accent } });
      slide.addText(s.title || 'Section', {
        x: 0.5, y: 2.34, w: 9, h: 0.97, fontSize: 27, bold: true, color: 'FFFFFF', align: 'center',
      });
    } else if (s.layout === 'blank') {
      if (s.image) {
        slide.addImage({ data: s.image, x: 0.5, y: 0.4, w: 9, h: 4.85, sizing: { type: 'contain', w: 9, h: 4.85 } });
      }
    } else {
      slide.addText(s.title || 'Slide', {
        x: 0.6, y: 0.35, w: 8.8, h: 1.1, fontSize: 27, bold: true, color: main,
      });
      slide.addShape(pptx.ShapeType.rect, { x: 0.62, y: 1.42, w: 1.6, h: 0.06, fill: { color: accent } });
      const wide = s.image ? 4.9 : 8.3;
      const lines = s.bullets.filter((b) => b.trim());
      if (lines.length) {
        slide.addText(
          lines.map((t) => ({ text: t, options: { bullet: true, fontSize: 15 } })),
          { x: 0.8, y: 1.75, w: wide, h: 3.6, color: body },
        );
      }
    }
    if (s.image && s.layout !== 'blank') {
      slide.addImage({ data: s.image, x: 5.85, y: 1.8, w: 3.6, h: 2.7, sizing: { type: 'contain', w: 3.6, h: 2.7 } });
    }
  }
  const out = (await pptx.write({ outputType: 'arraybuffer' })) as ArrayBuffer;
  return new Uint8Array(out);
}
