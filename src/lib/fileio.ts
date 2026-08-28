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
// PowerPoint (.pptx)
// ---------------------------------------------------------------------------

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

/** .pptx -> deck slides (title = first text box, bullets = the rest). */
export async function importPptx(buf: ArrayBuffer): Promise<{ title: string; bullets: string[] }[]> {
  const JSZip = (await import('jszip')).default;
  const zip = await JSZip.loadAsync(buf);
  const names = Object.keys(zip.files)
    .filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))
    .sort((a, b) => {
      const na = parseInt(a.replace(/\D+/g, ''), 10);
      const nb = parseInt(b.replace(/\D+/g, ''), 10);
      return na - nb;
    });
  const slides: { title: string; bullets: string[] }[] = [];
  for (const n of names) {
    const xml = await zip.files[n].async('text');
    const texts = Array.from(xml.matchAll(/<a:t>([^<]*)<\/a:t>/g))
      .map((m) => decodeXmlEntities(m[1]).trim())
      .filter(Boolean);
    slides.push({ title: texts[0] ?? 'Slide', bullets: texts.slice(1) });
  }
  return slides;
}

/** Deck slides -> real .pptx file (pptxgenjs). */
export async function exportPptx(
  title: string,
  slides: { title: string; bullets: string[] }[],
): Promise<Uint8Array> {
  const PptxGenJS = (await import('pptxgenjs')).default;
  const pptx = new PptxGenJS();
  pptx.title = title;
  for (const s of slides) {
    const slide = pptx.addSlide();
    slide.addText(s.title || 'Slide', {
      x: 0.6,
      y: 0.4,
      w: 8.8,
      h: 1.2,
      fontSize: 28,
      bold: true,
      color: '1F2430',
    });
    const lines = s.bullets.filter((b) => b.trim());
    if (lines.length) {
      slide.addText(
        lines.map((t) => ({ text: t, options: { bullet: true, fontSize: 16 } })),
        { x: 0.8, y: 1.9, w: 8.3, h: 4.6, color: '333333' },
      );
    }
  }
  const out = (await pptx.write({ outputType: 'arraybuffer' })) as ArrayBuffer;
  return new Uint8Array(out);
}
