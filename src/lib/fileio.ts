import { Capacitor } from '@capacitor/core';
import { Directory, Filesystem } from '@capacitor/filesystem';
import type { CellObject, WorkSheet } from 'xlsx';
import type * as DocxNs from 'docx';
import type { Book, CellStyle, Page } from './sheet-model';
import { applyXlsxExtras, readXlsxExtras } from './xlsx-styles';

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
  const res = await mammoth.convertToHtml(
    // browser build reads `arrayBuffer`; the Node build (used by the test scripts) wants `buffer`
    { arrayBuffer: buf, ...(typeof Buffer !== 'undefined' ? { buffer: Buffer.from(buf) } : {}) } as { arrayBuffer: ArrayBuffer },
    {
      styleMap: [
        "p[style-name='Title'] => h4:fresh",
        "p[style-name='Subtitle'] => h3:fresh",
        "p[style-name='Quote'] => blockquote:fresh",
        "p[style-name='Intense Quote'] => blockquote:fresh",
        "r[style-name='Strong'] => strong",
        "r[style-name='Emphasis'] => em",
        "br[type='page'] => hr.page-break", // mammoth's style-map grammar only accepts single-quoted strings
        'highlight => mark',
        'strike => s',
        'u => u',
      ],
      includeDefaultStyleMap: true,
      convertImage: mammoth.images.imgElement(async (img) => ({ src: `data:${img.contentType};base64,${await img.read('base64')}` })),
    },
  );
  const host = document.createElement('div');
  host.innerHTML = res.value;
  // mammoth emits the page break inline (<p>x<hr class="page-break">y</p>). The HTML parser already hoists
  // the <hr> out of the paragraph; tidy the empty shells it leaves behind and re-wrap any trailing text.
  const isEmptyBlock = (n: Node | null) => !!n && n.nodeType === 1 && (n as Element).tagName === 'P' && !(n as Element).textContent?.trim() && !(n as Element).querySelector('img, br');
  for (const hr of Array.from(host.querySelectorAll('hr.page-break'))) {
    while (isEmptyBlock(hr.previousSibling)) hr.previousSibling!.remove();
    while (isEmptyBlock(hr.nextSibling)) hr.nextSibling!.remove();
    if (hr.nextSibling && (hr.nextSibling.nodeType === 3 || !/^(P|H[1-6]|UL|OL|TABLE|BLOCKQUOTE|HR|DIV)$/.test((hr.nextSibling as Element).tagName))) {
      const wrap = document.createElement('p');
      while (hr.nextSibling && (hr.nextSibling.nodeType === 3 || !/^(P|H[1-6]|UL|OL|TABLE|BLOCKQUOTE|HR|DIV)$/.test((hr.nextSibling as Element).tagName))) wrap.appendChild(hr.nextSibling);
      if (wrap.textContent?.trim() || wrap.querySelector('img')) hr.after(wrap);
    }
  }
  // paragraph alignment is not in mammoth's default map; recover it from the raw XML when the document is small enough
  try {
    const JSZip = (await import('jszip')).default;
    const zip = await JSZip.loadAsync(buf);
    const xml = await zip.file('word/document.xml')?.async('string');
    if (xml && xml.length < 2_000_000) {
      const doc = new DOMParser().parseFromString(xml, 'application/xml');
      const body = doc.getElementsByTagNameNS('*', 'body')[0] ?? doc.documentElement;
      // only paragraphs that carry text or drawings produce blocks in mammoth's output (pure page-break paragraphs don't)
      const paras = Array.from(body.getElementsByTagNameNS('*', 'p')).filter((p) => (p.textContent ?? '').trim() !== '' || p.getElementsByTagNameNS('*', 'drawing').length > 0);
      const aligns = paras.map((p) => {
        const jc = p.getElementsByTagNameNS('*', 'jc')[0];
        return jc ? jc.getAttribute('w:val') ?? jc.getAttribute('val') ?? '' : '';
      });
      if (aligns.some((a) => a && a !== 'left' && a !== 'start')) {
        const blocks = Array.from(host.querySelectorAll('p, h1, h2, h3, h4, h5, h6, li')).filter((b) => !b.querySelector('p, li'));
        if (blocks.length === paras.length) {
          blocks.forEach((b, i) => {
            const a = aligns[i];
            if (a === 'center' || a === 'right' || a === 'both') (b as HTMLElement).style.textAlign = a === 'both' ? 'justify' : a;
          });
        }
      }
    }
  } catch {
    /* alignment recovery is best effort */
  }
  const html = host.innerHTML;
  return html;
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
  strike?: boolean;
  superScript?: boolean;
  subScript?: boolean;
  color?: string;
  highlight?: string;
  font?: string;
  size?: number; // half-points
  link?: string;
}

export interface DocExportOptions {
  page?: { size: 'A4' | 'Letter' | 'Legal'; orientation: 'portrait' | 'landscape'; margins: 'normal' | 'narrow' | 'wide' };
  header?: string;
  footer?: string;
  pageNumbers?: boolean;
  font?: string;
  lineSpacing?: number;
}

const PAGE_INCH: Record<string, { w: number; h: number }> = { A4: { w: 8.27, h: 11.69 }, Letter: { w: 8.5, h: 11 }, Legal: { w: 8.5, h: 14 } };
const MARGIN_INCH: Record<string, number> = { normal: 1, narrow: 0.5, wide: 1.5 };

function cssColorToHex(c: string | null | undefined): string | undefined {
  if (!c || c === 'transparent' || c === 'inherit' || c === 'initial') return undefined;
  const m = /^rgba?\((\d+),\s*(\d+),\s*(\d+)/.exec(c);
  if (m) return [m[1], m[2], m[3]].map((x) => parseInt(x, 10).toString(16).padStart(2, '0')).join('').toUpperCase();
  const h = /^#([0-9a-f]{3,8})$/i.exec(c.trim());
  if (h) {
    let hex = h[1];
    if (hex.length === 3) hex = hex.split('').map((ch) => ch + ch).join('');
    return hex.slice(0, 6).toUpperCase();
  }
  const named: Record<string, string> = { black: '000000', white: 'FFFFFF', red: 'FF0000', blue: '0000FF', green: '008000', yellow: 'FFFF00', gray: '808080', grey: '808080', orange: 'FFA500', purple: '800080' };
  return named[c.toLowerCase()];
}

function propsFromElement(el: Element, base: RunProps): RunProps {
  const tag = el.tagName?.toLowerCase() ?? '';
  const p: RunProps = { ...base };
  if (tag === 'b' || tag === 'strong') p.bold = true;
  if (tag === 'i' || tag === 'em') p.italics = true;
  if (tag === 'u') p.underline = true;
  if (tag === 's' || tag === 'strike' || tag === 'del') p.strike = true;
  if (tag === 'sup') p.superScript = true;
  if (tag === 'sub') p.subScript = true;
  if (tag === 'a') p.link = (el as HTMLAnchorElement).getAttribute('href') ?? undefined;
  if (tag === 'mark') p.highlight = p.highlight ?? 'yellow';
  if (tag === 'code') p.font = 'Courier New';
  const st = (el as HTMLElement).style;
  if (st) {
    if (st.fontWeight === 'bold' || parseInt(st.fontWeight, 10) >= 600) p.bold = true;
    if (st.fontStyle === 'italic') p.italics = true;
    if (st.textDecoration?.includes('underline') || st.textDecorationLine?.includes('underline')) p.underline = true;
    if (st.textDecoration?.includes('line-through') || st.textDecorationLine?.includes('line-through')) p.strike = true;
    const col = cssColorToHex(st.color);
    if (col) p.color = col;
    const bg = cssColorToHex(st.backgroundColor);
    if (bg) p.highlight = bg;
    if (st.fontFamily) p.font = st.fontFamily.split(',')[0].replace(/["']/g, '').trim();
    if (st.fontSize) {
      const m = /^([\d.]+)(pt|px)$/.exec(st.fontSize);
      if (m) p.size = Math.round(parseFloat(m[1]) * (m[2] === 'pt' ? 2 : 1.5));
    }
    if (st.verticalAlign === 'super') p.superScript = true;
    if (st.verticalAlign === 'sub') p.subScript = true;
  }
  return p;
}

const HIGHLIGHT_NAMES = ['yellow', 'green', 'cyan', 'magenta', 'blue', 'red', 'darkBlue', 'darkCyan', 'darkGreen', 'darkMagenta', 'darkRed', 'darkYellow', 'darkGray', 'lightGray', 'black'];

function docxRuns(node: Node, props: RunProps, D: typeof DocxNs): DocxNs.ParagraphChild[] {
  if (node.nodeType === 3) {
    const text = node.textContent ?? '';
    if (!text) return [];
    const run = new D.TextRun({
      text,
      bold: props.bold,
      italics: props.italics,
      underline: props.underline ? {} : undefined,
      strike: props.strike,
      superScript: props.superScript,
      subScript: props.subScript,
      color: props.color,
      font: props.font,
      size: props.size,
      highlight: props.highlight && HIGHLIGHT_NAMES.includes(props.highlight) ? (props.highlight as 'yellow') : undefined,
      shading: props.highlight && !HIGHLIGHT_NAMES.includes(props.highlight) ? { type: D.ShadingType.CLEAR, fill: props.highlight } : undefined,
    });
    if (props.link) return [new D.ExternalHyperlink({ link: props.link, children: [new D.TextRun({ text, style: 'Hyperlink', bold: props.bold, italics: props.italics })] })];
    return [run];
  }
  if (node.nodeType !== 1) return [];
  const el = node as Element;
  const tag = el.tagName.toLowerCase();
  if (tag === 'br') return [new D.TextRun({ text: '', break: 1 })];
  if (tag === 'img') {
    const img = imageRun(el as HTMLImageElement, D);
    return img ? [img] : [];
  }
  const p = propsFromElement(el, props);
  return Array.from(el.childNodes).flatMap((c) => docxRuns(c, p, D));
}

function decodeImageDataUrl(src: string): { bytes: Uint8Array; type: 'png' | 'jpg' | 'gif' | 'bmp' } | null {
  const m = /^data:image\/(png|jpe?g|gif|bmp);base64,(.+)$/i.exec(src);
  if (!m) return null;
  const bin = atob(m[2]);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  const t = m[1].toLowerCase();
  return { bytes, type: t === 'jpeg' ? 'jpg' : (t as 'png' | 'gif' | 'bmp' | 'jpg') };
}

function imageRun(img: HTMLImageElement, D: typeof DocxNs): DocxNs.ImageRun | null {
  const data = decodeImageDataUrl(img.getAttribute('src') ?? '');
  if (!data) return null;
  let w = img.naturalWidth || parseInt(img.getAttribute('width') ?? '', 10) || 480;
  let h = img.naturalHeight || parseInt(img.getAttribute('height') ?? '', 10) || 320;
  const maxW = 600;
  if (w > maxW) { h = Math.round((h * maxW) / w); w = maxW; }
  try {
    return new D.ImageRun({ type: data.type, data: data.bytes, transformation: { width: w, height: h } });
  } catch {
    return null;
  }
}

function alignmentOf(el: Element, D: typeof DocxNs) {
  const a = (el as HTMLElement).style?.textAlign || el.getAttribute('align') || '';
  if (a === 'center') return D.AlignmentType.CENTER;
  if (a === 'right') return D.AlignmentType.RIGHT;
  if (a === 'justify') return D.AlignmentType.JUSTIFIED;
  return undefined;
}

/** Docs editor HTML -> real .docx file (docx library). */
export async function exportDocx(title: string, html: string, opts: DocExportOptions = {}): Promise<Blob> {
  const D = await import('docx');
  const doc = new DOMParser().parseFromString(`<body>${html}</body>`, 'text/html');
  const body = doc.body;
  const children: (DocxNs.Paragraph | DocxNs.Table)[] = [];
  const lineSpacing = opts.lineSpacing ? { line: Math.round(opts.lineSpacing * 240) } : undefined;

  const para = (el: Element, extra: Partial<DocxNs.IParagraphOptions> = {}) =>
    new D.Paragraph({ children: docxRuns(el, {}, D), alignment: alignmentOf(el, D), spacing: lineSpacing, ...extra });

  const emitList = (list: Element, level: number, ordered: boolean) => {
    Array.from(list.children).forEach((li) => {
      if (li.tagName.toLowerCase() !== 'li') return;
      const nested = Array.from(li.children).filter((c) => /^(ul|ol)$/i.test(c.tagName));
      const clone = li.cloneNode(true) as Element;
      clone.querySelectorAll('ul,ol').forEach((n) => n.remove());
      children.push(
        new D.Paragraph({
          children: docxRuns(clone, {}, D),
          alignment: alignmentOf(li, D),
          spacing: lineSpacing,
          ...(ordered ? { numbering: { reference: 'numbers', level: Math.min(level, 2) } } : { bullet: { level: Math.min(level, 2) } }),
        }),
      );
      nested.forEach((n) => emitList(n, level + 1, n.tagName.toLowerCase() === 'ol'));
    });
  };

  const emitTable = (table: HTMLTableElement) => {
    const rows = Array.from(table.rows);
    if (rows.length === 0) return;
    const cols = Math.max(...rows.map((r) => r.cells.length));
    children.push(
      new D.Table({
        width: { size: 100, type: D.WidthType.PERCENTAGE },
        rows: rows.map(
          (r) =>
            new D.TableRow({
              children: Array.from({ length: cols }, (_, i) => {
                const cell = r.cells[i];
                const isHead = cell?.tagName === 'TH';
                return new D.TableCell({
                  shading: isHead ? { type: D.ShadingType.CLEAR, fill: 'E7E6E6' } : undefined,
                  children: [new D.Paragraph({ children: cell ? docxRuns(cell, { bold: isHead || undefined }, D) : [] })],
                });
              }),
            }),
        ),
      }),
    );
    children.push(new D.Paragraph({ text: '' }));
  };

  const emitBlock = (b: Element) => {
    const tag = b.tagName.toLowerCase();
    if (tag === 'ul' || tag === 'ol') emitList(b, 0, tag === 'ol');
    else if (tag === 'table') emitTable(b as HTMLTableElement);
    else if (tag === 'hr') {
      if (b.classList.contains('page-break')) children.push(new D.Paragraph({ children: [new D.PageBreak()] }));
      else children.push(new D.Paragraph({ border: { bottom: { color: '999999', space: 1, style: D.BorderStyle.SINGLE, size: 6 } }, text: '' }));
    } else if (tag === 'h1') children.push(para(b, { heading: D.HeadingLevel.HEADING_1 }));
    else if (tag === 'h2') children.push(para(b, { heading: D.HeadingLevel.HEADING_2 }));
    else if (tag === 'h3') children.push(para(b, { heading: D.HeadingLevel.HEADING_3 }));
    else if (tag === 'h4') children.push(para(b, { heading: D.HeadingLevel.TITLE }));
    else if (tag === 'blockquote') children.push(para(b, { indent: { left: 720 }, border: { left: { color: 'BBBBBB', space: 8, style: D.BorderStyle.SINGLE, size: 18 } } }));
    else if (tag === 'pre') children.push(new D.Paragraph({ children: docxRuns(b, { font: 'Courier New', size: 20 }, D), shading: { type: D.ShadingType.CLEAR, fill: 'F3F3F3' } }));
    else if (tag === 'div' || tag === 'section') {
      const kids = Array.from(b.children);
      if (kids.length && kids.every((k) => /^(p|h[1-6]|ul|ol|table|hr|div|blockquote|pre)$/i.test(k.tagName))) kids.forEach(emitBlock);
      else if (b.textContent?.trim() || b.querySelector('img')) children.push(para(b));
    } else if (b.textContent?.trim() || b.querySelector('img')) children.push(para(b));
    else if (tag === 'p') children.push(new D.Paragraph({ text: '' }));
  };

  const blocks = Array.from(body.children);
  if (blocks.length === 0) {
    if (body.textContent?.trim()) children.push(new D.Paragraph({ children: docxRuns(body, {}, D) }));
  } else {
    // stray text nodes directly in body become paragraphs
    Array.from(body.childNodes).forEach((n) => {
      if (n.nodeType === 1) emitBlock(n as Element);
      else if (n.textContent?.trim()) children.push(new D.Paragraph({ text: n.textContent.trim() }));
    });
  }

  const pg = opts.page ?? { size: 'A4' as const, orientation: 'portrait' as const, margins: 'normal' as const };
  const inch = PAGE_INCH[pg.size] ?? PAGE_INCH.A4;
  const m = Math.round((MARGIN_INCH[pg.margins] ?? 1) * 1440);
  const landscape = pg.orientation === 'landscape';
  const footerRuns: DocxNs.ParagraphChild[] = [];
  if (opts.footer) footerRuns.push(new D.TextRun({ text: opts.footer }));
  if (opts.pageNumbers) footerRuns.push(new D.TextRun({ children: [opts.footer ? '    Page ' : 'Page ', D.PageNumber.CURRENT, ' of ', D.PageNumber.TOTAL_PAGES] }));

  const document = new D.Document({
    title,
    creator: 'GenOffice',
    styles: { default: { document: { run: { font: opts.font ?? 'Calibri', size: 22 } } } },
    numbering: {
      config: [
        {
          reference: 'numbers',
          levels: [0, 1, 2].map((lvl) => ({
            level: lvl,
            format: D.LevelFormat.DECIMAL,
            text: `%${lvl + 1}.`,
            alignment: D.AlignmentType.START,
            style: { paragraph: { indent: { left: 720 * (lvl + 1), hanging: 360 } } },
          })),
        },
      ],
    },
    sections: [
      {
        properties: {
          page: {
            size: { width: Math.round((landscape ? inch.h : inch.w) * 1440), height: Math.round((landscape ? inch.w : inch.h) * 1440), orientation: landscape ? D.PageOrientation.LANDSCAPE : D.PageOrientation.PORTRAIT },
            margin: { top: m, bottom: m, left: m, right: m },
          },
        },
        headers: opts.header ? { default: new D.Header({ children: [new D.Paragraph({ children: [new D.TextRun({ text: opts.header, color: '666666' })], alignment: D.AlignmentType.RIGHT })] }) } : undefined,
        footers: footerRuns.length ? { default: new D.Footer({ children: [new D.Paragraph({ children: footerRuns, alignment: D.AlignmentType.CENTER })] }) } : undefined,
        children,
      },
    ],
  });
  return D.Packer.toBlob(document);
}

/** Editor HTML -> paginated PDF (rendered offscreen with html2canvas via jsPDF.html). */
export async function exportPdfFromHtml(title: string, html: string, opts: DocExportOptions = {}): Promise<Uint8Array> {
  const { jsPDF } = await import('jspdf');
  const pg = opts.page ?? { size: 'A4' as const, orientation: 'portrait' as const, margins: 'normal' as const };
  const pdf = new jsPDF({ orientation: pg.orientation, unit: 'pt', format: pg.size.toLowerCase() as 'a4' | 'letter' | 'legal', compress: true });
  const pageW = pdf.internal.pageSize.getWidth();
  const pageH = pdf.internal.pageSize.getHeight();
  const margin = (MARGIN_INCH[pg.margins] ?? 1) * 72;
  const contentW = pageW - margin * 2;

  // Offscreen container that mirrors the page width so line breaks match.
  const host = document.createElement('div');
  host.style.cssText = `position:fixed;left:-10000px;top:0;width:${contentW}pt;background:#fff;color:#111;font-family:${opts.font ?? 'Calibri'},'Segoe UI',Arial,sans-serif;font-size:11pt;line-height:${opts.lineSpacing ?? 1.15};`;
  host.innerHTML = `<style>
    .pdfroot p{margin:0 0 8pt}.pdfroot h1{font-size:20pt;margin:14pt 0 6pt}.pdfroot h2{font-size:16pt;margin:12pt 0 5pt}.pdfroot h3{font-size:13pt;margin:10pt 0 4pt}
    .pdfroot h4{font-size:24pt;text-align:center;margin:0 0 12pt}.pdfroot table{border-collapse:collapse;width:100%;margin:6pt 0}.pdfroot td,.pdfroot th{border:1px solid #888;padding:3pt 5pt;vertical-align:top}
    .pdfroot th{background:#e7e6e6}.pdfroot img{max-width:100%;height:auto}.pdfroot blockquote{border-left:3px solid #bbb;margin:6pt 0;padding:2pt 10pt;color:#444}
    .pdfroot pre{background:#f3f3f3;padding:6pt;font-family:'Courier New',monospace;font-size:10pt;white-space:pre-wrap}.pdfroot hr{border:0;border-top:1px solid #999;margin:8pt 0}
    .pdfroot hr.page-break{page-break-after:always;border:0;margin:0}.pdfroot mark.cmt{background:#fff3b0}.pdfroot mark.find-hit{background:transparent}.pdfroot a{color:#1155cc}
  </style><div class="pdfroot">${html}</div>`;
  document.body.appendChild(host);
  try {
    await pdf.html(host, {
      x: margin,
      y: margin,
      width: contentW,
      windowWidth: Math.round(contentW * (96 / 72)),
      margin: [margin, margin, margin, margin],
      autoPaging: 'text',
      html2canvas: { scale: 2, useCORS: true, logging: false, backgroundColor: '#ffffff' },
    });
    const pages = pdf.getNumberOfPages();
    for (let i = 1; i <= pages; i++) {
      pdf.setPage(i);
      pdf.setFontSize(9);
      pdf.setTextColor(110);
      if (opts.header) pdf.text(opts.header, pageW - margin, margin * 0.6, { align: 'right' });
      const footParts: string[] = [];
      if (opts.footer) footParts.push(opts.footer);
      if (opts.pageNumbers) footParts.push(`Page ${i} of ${pages}`);
      if (footParts.length) pdf.text(footParts.join('    '), pageW / 2, pageH - margin * 0.55, { align: 'center' });
    }
    pdf.setProperties({ title, creator: 'GenOffice' });
    return new Uint8Array(pdf.output('arraybuffer'));
  } finally {
    host.remove();
  }
}

/** Editor HTML -> readable Markdown (headings, lists, emphasis, links, tables, images). */
export function htmlToMarkdown(html: string): string {
  const doc = new DOMParser().parseFromString(`<body>${html}</body>`, 'text/html');
  const inline = (n: Node): string => {
    if (n.nodeType === 3) return (n.textContent ?? '').replace(/\s+/g, ' ');
    if (n.nodeType !== 1) return '';
    const el = n as Element;
    const tag = el.tagName.toLowerCase();
    const inner = () => Array.from(el.childNodes).map(inline).join('');
    switch (tag) {
      case 'b': case 'strong': return `**${inner()}**`;
      case 'i': case 'em': return `*${inner()}*`;
      case 's': case 'strike': case 'del': return `~~${inner()}~~`;
      case 'code': return `\`${inner()}\``;
      case 'br': return '  \n';
      case 'a': return `[${inner()}](${el.getAttribute('href') ?? ''})`;
      case 'img': return `![${el.getAttribute('alt') ?? ''}](${(el.getAttribute('src') ?? '').startsWith('data:') ? 'image' : el.getAttribute('src')})`;
      case 'sup': return `^${inner()}^`;
      case 'sub': return `~${inner()}~`;
      default: return inner();
    }
  };
  const out: string[] = [];
  const block = (el: Element, depth = 0) => {
    const tag = el.tagName.toLowerCase();
    const pad = '  '.repeat(depth);
    if (/^h[1-6]$/.test(tag)) out.push(`${'#'.repeat(Math.min(6, parseInt(tag[1], 10)))} ${inline(el).trim()}\n`);
    else if (tag === 'ul' || tag === 'ol') {
      Array.from(el.children).forEach((li, i) => {
        const clone = li.cloneNode(true) as Element;
        const nested = Array.from(li.children).filter((c) => /^(ul|ol)$/i.test(c.tagName));
        clone.querySelectorAll('ul,ol').forEach((x) => x.remove());
        out.push(`${pad}${tag === 'ol' ? `${i + 1}.` : '-'} ${inline(clone).trim()}`);
        nested.forEach((n) => block(n, depth + 1));
      });
      if (depth === 0) out.push('');
    } else if (tag === 'table') {
      const rows = Array.from((el as HTMLTableElement).rows);
      rows.forEach((r, i) => {
        const cells = Array.from(r.cells).map((c) => inline(c).trim().replace(/\|/g, '\\|'));
        out.push(`| ${cells.join(' | ')} |`);
        if (i === 0) out.push(`| ${cells.map(() => '---').join(' | ')} |`);
      });
      out.push('');
    } else if (tag === 'blockquote') out.push(`> ${inline(el).trim()}\n`);
    else if (tag === 'pre') out.push(`\`\`\`\n${el.textContent ?? ''}\n\`\`\`\n`);
    else if (tag === 'hr') out.push(el.classList.contains('page-break') ? '\n<div style="page-break-after:always"></div>\n' : '---\n');
    else if (tag === 'div' && Array.from(el.children).some((k) => /^(p|h[1-6]|ul|ol|table)$/i.test(k.tagName))) Array.from(el.children).forEach((k) => block(k, depth));
    else {
      const t = inline(el).trim();
      out.push(t ? `${t}\n` : '');
    }
  };
  Array.from(doc.body.childNodes).forEach((n) => {
    if (n.nodeType === 1) block(n as Element);
    else if (n.textContent?.trim()) out.push(`${n.textContent.trim()}\n`);
  });
  return out.join('\n').replace(/\n{3,}/g, '\n\n').trim() + '\n';
}

// ---------------------------------------------------------------------------
// Spreadsheets (.xlsx / .xls / .csv / .ods)
// ---------------------------------------------------------------------------


const IMPORT_MAX_ROWS = 1000;
const IMPORT_MAX_COLS = 100;

function argbToHex(rgb?: string): string | undefined {
  if (!rgb) return undefined;
  const h = rgb.length === 8 ? rgb.slice(2) : rgb;
  return /^[0-9A-Fa-f]{6}$/.test(h) ? `#${h.toUpperCase()}` : undefined;
}

/** Every sheet of an Excel/CSV/ODS file -> workbook model. Formulas, basic styles, widths and merges kept. */
export async function importWorkbook(buf: ArrayBuffer, fileName = ''): Promise<Book> {
  const X = await import('xlsx');
  const isText = /\.(csv|tsv|txt)$/i.test(fileName);
  const wb = isText
    ? X.read(new TextDecoder().decode(buf), { type: 'string', cellFormula: true, cellText: false, cellStyles: true, cellDates: false, raw: true })
    : X.read(buf, { type: 'array', cellFormula: true, cellText: false, cellStyles: true, cellDates: false, cellNF: true });
  const order: string[] = [];
  const sheets: Record<string, Page> = {};
  // SheetJS (community build) only surfaces number formats and fill colours; read fonts, borders,
  // alignment and frozen panes straight from the package parts
  const extras = isText ? null : await readXlsxExtras(buf, IMPORT_MAX_ROWS, IMPORT_MAX_COLS);
  for (const name of wb.SheetNames) {
    const ws = wb.Sheets[name];
    const page: Page = { cells: {}, styles: {} };
    order.push(name);
    sheets[name] = page;
    const extra = extras?.[name];
    if (extra?.freeze) page.freeze = extra.freeze;
    if (!ws || !ws['!ref']) continue;
    const range = X.utils.decode_range(ws['!ref']);
    for (let r = range.s.r; r <= Math.min(range.e.r, IMPORT_MAX_ROWS - 1); r++) {
      for (let c = range.s.c; c <= Math.min(range.e.c, IMPORT_MAX_COLS - 1); c++) {
        const addr = X.utils.encode_cell({ r, c });
        const cell = ws[addr] as (CellObject & { s?: { font?: { bold?: boolean; italic?: boolean; underline?: boolean; color?: { rgb?: string } }; fill?: { fgColor?: { rgb?: string } }; alignment?: { horizontal?: string; wrapText?: boolean } } }) | undefined;
        if (!cell) continue;
        let raw: string | undefined;
        if (cell.f) raw = `=${cell.f}`;
        else if (cell.v !== undefined && cell.v !== null) raw = cell.t === 'b' ? (cell.v ? 'TRUE' : 'FALSE') : String(cell.v);
        if (raw !== undefined && raw !== '') page.cells[addr] = raw;
        const st: CellStyle = {};
        const nf = (cell.z ?? '').toString();
        if (nf) {
          if (/%/.test(nf)) st.fmt = 'pct';
          else if (/[$€£¥₹]|\[\$/.test(nf)) st.fmt = 'cur';
          else if (/[dmy]{1,4}[/\-.][dmy]/i.test(nf) || /yyyy|mmm/i.test(nf)) st.fmt = 'date';
          else if (/h+:mm/i.test(nf)) st.fmt = 'time';
          else if (/E\+/.test(nf)) st.fmt = 'sci';
          else if (/#,##0\.00|0\.00/.test(nf)) st.fmt = 'num';
          else if (nf === '@') st.fmt = 'text';
          const dm = /0\.(0+)/.exec(nf);
          if (dm && st.fmt && st.fmt !== 'pct') st.dec = dm[1].length;
        }
        const s = cell.s;
        if (s?.font?.bold) st.b = true;
        if (s?.font?.italic) st.i = true;
        if (s?.font?.underline) st.u = true;
        const fc = argbToHex(s?.font?.color?.rgb);
        if (fc && fc !== '#000000') st.color = fc;
        const bg = argbToHex(s?.fill?.fgColor?.rgb);
        if (bg && bg !== '#FFFFFF') st.fill = bg;
        if (s?.alignment?.horizontal === 'center' || s?.alignment?.horizontal === 'right' || s?.alignment?.horizontal === 'left') st.align = s.alignment.horizontal;
        if (s?.alignment?.wrapText) st.wrap = true;
        const ex = extra?.styles[addr];
        if (ex) Object.assign(st, ex);
        if (Object.keys(st).length) page.styles[addr] = st;
      }
    }
    // styled-but-empty cells (e.g. a filled header band) have no CellObject to iterate above
    if (extra) for (const [addr, st] of Object.entries(extra.styles)) if (!page.styles[addr]) page.styles[addr] = { ...st };
    const cols = ws['!cols'] as { wpx?: number; wch?: number }[] | undefined;
    if (cols) {
      page.colW = {};
      cols.forEach((col, i) => { const px = col?.wpx ?? (col?.wch ? Math.round(col.wch * 7 + 5) : undefined); if (px && i < IMPORT_MAX_COLS) page.colW![i] = Math.max(30, Math.min(500, px)); });
    }
    const rows = ws['!rows'] as { hpx?: number; hpt?: number }[] | undefined;
    if (rows) {
      page.rowH = {};
      rows.forEach((row, i) => { const px = row?.hpx ?? (row?.hpt ? Math.round(row.hpt * 1.33) : undefined); if (px && i < IMPORT_MAX_ROWS) page.rowH![i] = Math.max(18, Math.min(300, px)); });
    }
    const merges = ws['!merges'] as { s: { r: number; c: number }; e: { r: number; c: number } }[] | undefined;
    if (merges?.length) page.merges = merges.map((m) => `${X.utils.encode_cell(m.s)}:${X.utils.encode_cell(m.e)}`);
    const af = (ws as unknown as { '!autofilter'?: { ref: string } })['!autofilter'];
    if (af?.ref) page.filter = { range: af.ref, criteria: {} };
  }
  if (order.length === 0) { order.push('Sheet1'); sheets.Sheet1 = { cells: {}, styles: {} }; }
  return { order, active: order[0], sheets };
}

/** Back-compat: first sheet only. */
export async function importSpreadsheet(buf: ArrayBuffer): Promise<{ cells: Record<string, string> }> {
  const wb = await importWorkbook(buf);
  return { cells: wb.sheets[wb.order[0]].cells };
}

/** Whole workbook -> .xlsx with live formulas, cached values, number formats, widths and merges. */
export async function exportXlsx(title: string, book: Book | Record<string, string>): Promise<Uint8Array> {
  const X = await import('xlsx');
  const { evalCell } = await import('./formulas');
  const { parseRef } = await import('./sheet-model');
  const bk: Book = 'order' in book && 'sheets' in book ? (book as Book) : { order: ['Sheet1'], active: 'Sheet1', sheets: { Sheet1: { cells: book as Record<string, string>, styles: {} } } };
  const wb = X.utils.book_new();
  const used = new Set<string>();
  const written: { name: string; page: Page }[] = [];
  for (const name of bk.order) {
    const pg = bk.sheets[name];
    if (!pg) continue;
    const ws: WorkSheet = {};
    let maxR = 0, maxC = 0;
    const cache = new Map();
    const get = (ref: string) => pg.cells[ref] ?? '';
    const keys = new Set([...Object.keys(pg.cells), ...Object.keys(pg.styles)]);
    for (const key of keys) {
      const raw = pg.cells[key];
      const st = pg.styles[key];
      const [c, r] = parseRef(key);
      if (r >= IMPORT_MAX_ROWS || c >= IMPORT_MAX_COLS) continue;
      maxR = Math.max(maxR, r); maxC = Math.max(maxC, c);
      let cell: CellObject | null = null;
      if (raw !== undefined && raw !== '') {
        if (raw.startsWith('=')) {
          const shown = evalCell(raw, get, { cache });
          const num = parseFloat(shown.replace(/[,$%]/g, ''));
          const isNum = !shown.startsWith('#') && Number.isFinite(num) && /^-?[\d,.$%]+$/.test(shown.trim());
          cell = isNum ? { t: 'n', v: num, f: raw.slice(1) } : shown === 'TRUE' || shown === 'FALSE' ? { t: 'b', v: shown === 'TRUE', f: raw.slice(1) } : { t: 's', v: shown.startsWith('#') ? '' : shown, f: raw.slice(1) };
        } else {
          const num = Number(raw.trim());
          const up = raw.trim().toUpperCase();
          if (raw.trim() !== '' && Number.isFinite(num) && st?.fmt !== 'text' && !raw.startsWith("'")) cell = { t: 'n', v: num };
          else if (up === 'TRUE' || up === 'FALSE') cell = { t: 'b', v: up === 'TRUE' };
          else cell = { t: 's', v: raw.startsWith("'") ? raw.slice(1) : raw };
        }
      } else if (st) cell = { t: 'z' } as CellObject;
      if (!cell) continue;
      if (st?.fmt) {
        const d = st.dec ?? 2;
        const zeros = d > 0 ? `.${'0'.repeat(d)}` : '';
        cell.z = st.fmt === 'num' ? `#,##0${zeros}` : st.fmt === 'cur' ? `"$"#,##0${zeros}` : st.fmt === 'pct' ? `0${st.dec ? `.${'0'.repeat(st.dec)}` : ''}%` : st.fmt === 'date' ? 'yyyy-mm-dd' : st.fmt === 'time' ? 'hh:mm' : st.fmt === 'sci' ? `0${zeros}E+00` : st.fmt === 'text' ? '@' : undefined;
      } else if (st?.dec !== undefined) cell.z = `0${st.dec > 0 ? `.${'0'.repeat(st.dec)}` : ''}`;
      if (st && (st.b || st.i || st.u || st.color || st.fill || st.align || st.wrap)) {
        // xlsx community build ignores styles on write, but keep them so style-aware builds pick them up
        (cell as CellObject & { s?: unknown }).s = {
          font: { bold: st.b, italic: st.i, underline: st.u, color: st.color ? { rgb: st.color.replace('#', '') } : undefined },
          fill: st.fill ? { patternType: 'solid', fgColor: { rgb: st.fill.replace('#', '') } } : undefined,
          alignment: { horizontal: st.align, wrapText: st.wrap },
        };
      }
      ws[key] = cell;
    }
    ws['!ref'] = X.utils.encode_range({ s: { c: 0, r: 0 }, e: { c: Math.max(maxC, 0), r: Math.max(maxR, 0) } });
    if (pg.colW) ws['!cols'] = Array.from({ length: maxC + 1 }, (_, i) => (pg.colW![i] ? { wpx: pg.colW![i] } : {}));
    if (pg.rowH) ws['!rows'] = Array.from({ length: maxR + 1 }, (_, i) => (pg.rowH![i] ? { hpx: pg.rowH![i] } : {}));
    if (pg.merges?.length) ws['!merges'] = pg.merges.map((m) => X.utils.decode_range(m));
    if (pg.filter) (ws as unknown as { '!autofilter': { ref: string } })['!autofilter'] = { ref: pg.filter.range };
    let sheetName = (name || 'Sheet').replace(/[\\/?*[\]:]/g, ' ').slice(0, 31) || 'Sheet';
    let i = 2;
    while (used.has(sheetName)) sheetName = `${sheetName.slice(0, 28)}(${i++})`;
    used.add(sheetName);
    written.push({ name: sheetName, page: pg });
    X.utils.book_append_sheet(wb, ws, sheetName);
  }
  wb.Props = { Title: title, Application: 'GenOffice' };
  const out = X.write(wb, { bookType: 'xlsx', type: 'array', cellStyles: true }) as ArrayBuffer;
  // SheetJS ignores fonts/fills/borders/alignment/freeze on write; patch them into the package
  return applyXlsxExtras(new Uint8Array(out), written);
}

/** Rasterise an inline <svg> element to PNG bytes (charts → image). */
export async function svgToPng(svg: SVGSVGElement, scale = 2): Promise<Uint8Array> {
  const w = svg.clientWidth || parseInt(svg.getAttribute('width') ?? '300', 10);
  const h = svg.clientHeight || parseInt(svg.getAttribute('height') ?? '200', 10);
  const clone = svg.cloneNode(true) as SVGSVGElement;
  clone.setAttribute('xmlns', 'http://www.w3.org/2000/svg');
  clone.setAttribute('width', String(w));
  clone.setAttribute('height', String(h));
  const xml = new XMLSerializer().serializeToString(clone);
  const url = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(xml)}`;
  const img = new Image();
  await new Promise<void>((res, rej) => { img.onload = () => res(); img.onerror = () => rej(new Error('SVG render failed')); img.src = url; });
  const canvas = document.createElement('canvas');
  canvas.width = Math.round(w * scale);
  canvas.height = Math.round(h * scale);
  const ctx = canvas.getContext('2d')!;
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
  const blob = await new Promise<Blob | null>((res) => canvas.toBlob(res, 'image/png'));
  if (!blob) throw new Error('PNG encode failed');
  return new Uint8Array(await blob.arrayBuffer());
}

/** Data URL (or Blob) → PNG bytes; used by slide/image exports. */
export async function dataUrlToBytes(dataUrl: string): Promise<Uint8Array> {
  const res = await fetch(dataUrl);
  return new Uint8Array(await res.arrayBuffer());
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
  /** stable id (added on load for older decks) */
  id?: string;
  title: string;
  bullets: string[];
  /** slide background color (#rrggbb) */
  bg?: string;
  /** accent color for bars / highlights (#rrggbb) */
  accent?: string;
  /** image placed on the slide (data URL) — legacy, converted to a shape on load */
  image?: string;
  layout?: 'title' | 'content' | 'section' | 'blank';
  /** faithful rendering of an imported slide: positioned shapes */
  shapes?: SlideShape[];
  /** slide canvas size in inches (defaults 10 x 5.63) */
  cw?: number;
  ch?: number;
  /** speaker notes */
  notes?: string;
  /** transition used when presenting */
  transition?: 'none' | 'fade' | 'push' | 'zoom' | 'flip';
  /** hidden slides are skipped when presenting/exporting */
  hidden?: boolean;
}

/** One styled text run inside a shape paragraph. */
export interface ShapeRun {
  text: string;
  /** font size in points */
  sz?: number;
  b?: boolean;
  i?: boolean;
  u?: boolean;
  s?: boolean;
  color?: string;
  font?: string;
  highlight?: string;
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
  id?: string;
  x: number;
  y: number;
  w: number;
  h: number;
  kind: 'text' | 'image' | 'shape';
  paras: ShapePara[];
  fill?: string;
  line?: string;
  lineW?: number;
  img?: string;
  /** preset geometry for kind 'shape' (and text boxes with a background) */
  geom?: 'rect' | 'roundRect' | 'ellipse' | 'triangle' | 'diamond' | 'rightArrow' | 'line' | 'star' | 'hexagon' | 'chevron';
  /** rotation in degrees */
  rot?: number;
  /** vertical text alignment */
  valign?: 'top' | 'middle' | 'bottom';
  /** semantic role (title/body/…): lets layouts and AI target the right box */
  name?: string;
  /** locked shapes can't be moved/deleted from the canvas (theme decorations) */
  locked?: boolean;
  /** entrance animation when presenting */
  anim?: 'none' | 'fadeIn' | 'flyIn' | 'zoomIn';
  opacity?: number;
  shadow?: boolean;
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
      const prst = /<a:prstGeom\s+prst="([a-zA-Z0-9]+)"/.exec(spPr)?.[1];
      const geomMap: Record<string, SlideShape['geom']> = { rect: 'rect', roundRect: 'roundRect', ellipse: 'ellipse', triangle: 'triangle', diamond: 'diamond', rightArrow: 'rightArrow', line: 'line', straightConnector1: 'line', star5: 'star', hexagon: 'hexagon', chevron: 'chevron', homePlate: 'chevron' };
      const rotRaw = /<a:xfrm[^>]*\brot="(-?\d+)"/.exec(sp)?.[1];
      const anchor = /<a:bodyPr\b[^>]*\banchor="(t|ctr|b)"/.exec(sp)?.[1];
      const phType = /<p:ph\b[^>]*\btype="([a-zA-Z]+)"/.exec(sp)?.[1];
      const lineWRaw = /<a:ln\b[^>]*\bw="(\d+)"/.exec(spPr)?.[1];
      shapes.push({
        ...geom,
        kind: isText ? 'text' : 'shape',
        paras,
        fill,
        line: lineM ? `#${lineM[1]}` : undefined,
        lineW: lineWRaw ? Math.max(0.5, parseInt(lineWRaw, 10) / 12700) : undefined,
        geom: prst ? geomMap[prst] ?? 'rect' : undefined,
        rot: rotRaw ? parseInt(rotRaw, 10) / 60000 : undefined,
        valign: anchor === 'ctr' ? 'middle' : anchor === 'b' ? 'bottom' : undefined,
        name: phType === 'title' || phType === 'ctrTitle' ? 'title' : phType === 'subTitle' ? 'subtitle' : phType === 'body' ? 'body' : undefined,
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

    // speaker notes
    let notes: string | undefined;
    const notesTarget = Object.entries(rels).find(([, t]) => t.includes('notesSlide'))?.[1];
    if (notesTarget) {
      const np = resolvePptPath('ppt/slides', notesTarget);
      const nxml = zip.files[np] ? await zip.files[np].async('text') : '';
      if (nxml) {
        const bodies = nxml.match(/<p:sp>[\s\S]*?<\/p:sp>/g) ?? [];
        const bodySp = bodies.find((b) => /type="body"/.test(b)) ?? bodies[bodies.length - 1];
        if (bodySp) {
          const nparas = parseParas(/<p:txBody\b[^>]*>([\s\S]*?)<\/p:txBody>/.exec(bodySp)?.[1] ?? '');
          const txt = nparas.map((p) => p.runs.map((r) => r.text).join('')).join('\n').trim();
          if (txt) notes = txt;
        }
      }
    }
    const hidden = /<p:sld\b[^>]*\bshow="0"/.test(xml);

    slides.push({
      title: title || 'Slide',
      bullets: bullets.length ? bullets : [''],
      bg: bg ?? undefined,
      shapes,
      cw,
      ch,
      notes,
      hidden: hidden || undefined,
    });
  }
  return slides;
}

/** Deck slides -> real .pptx file (pptxgenjs): shapes, text runs, images, notes, backgrounds. */
export async function exportPptx(title: string, slides: DeckSlide[]): Promise<Uint8Array> {
  const PptxGenJS = (await import('pptxgenjs')).default;
  const pptx = new PptxGenJS();
  pptx.title = title;
  pptx.defineLayout({ name: 'WIDE', width: 10, height: 5.63 });
  pptx.layout = 'WIDE';
  const geomMap: Record<NonNullable<SlideShape['geom']>, keyof typeof pptx.ShapeType> = {
    rect: 'rect', roundRect: 'roundRect', ellipse: 'ellipse', triangle: 'triangle', diamond: 'diamond', rightArrow: 'rightArrow', line: 'line', star: 'star5', hexagon: 'hexagon', chevron: 'chevron',
  };
  for (const s of slides) {
    if (s.hidden) continue;
    const slide = pptx.addSlide();
    const dark = isDarkColor(s.bg);
    if (s.bg) slide.background = { color: s.bg.replace('#', '') };
    if (s.notes) slide.addNotes(s.notes);

    const shapes = s.shapes && s.shapes.length > 0 ? s.shapes : legacyShapes(s, title);
    const offX = Math.max(0, (10 - (s.cw ?? 10)) / 2);
    for (const sh of shapes) {
      const x = sh.x + offX;
      const common = { x, y: sh.y, w: sh.w, h: sh.h, rotate: sh.rot };
      if (sh.kind === 'image' && sh.img) {
        slide.addImage({ data: sh.img, ...common, sizing: { type: 'contain', w: sh.w, h: sh.h } });
        continue;
      }
      const fill = sh.fill ? { color: sh.fill.replace('#', ''), transparency: sh.opacity !== undefined ? Math.round((1 - sh.opacity) * 100) : undefined } : undefined;
      const line = sh.line ? { color: sh.line.replace('#', ''), width: sh.lineW ?? 1 } : undefined;
      const hasText = sh.paras.some((p) => p.runs.some((r) => r.text.trim()));
      const shapeType = pptx.ShapeType[geomMap[sh.geom ?? 'rect'] ?? 'rect'];
      if (!hasText) {
        if (fill || line) slide.addShape(shapeType, { ...common, fill, line, shadow: sh.shadow ? { type: 'outer', blur: 6, offset: 2, angle: 45, opacity: 0.3 } : undefined });
        continue;
      }
      const items = sh.paras.flatMap((p) => {
        if (p.runs.length === 0) return [{ text: ' ', options: { breakLine: true, fontSize: 8 } }];
        return p.runs.map((r, ri) => ({
          text: r.text,
          options: {
            fontSize: r.sz ?? 14,
            fontFace: r.font,
            bold: r.b,
            italic: r.i,
            underline: r.u ? { style: 'sng' as const } : undefined,
            strike: r.s ? ('sngStrike' as const) : undefined,
            color: (r.color ?? (dark ? 'FFFFFF' : '333333')).replace('#', ''),
            highlight: r.highlight?.replace('#', ''),
            align: p.align,
            bullet: p.bullet ? { characterCode: '2022', indent: 12 + (p.level ?? 0) * 18 } : p.bullet === false ? false : undefined,
            indentLevel: p.level,
            breakLine: ri === p.runs.length - 1,
          },
        }));
      });
      slide.addText(items as never, {
        ...common,
        shape: sh.geom && sh.geom !== 'rect' ? shapeType : undefined,
        valign: sh.valign === 'middle' ? 'middle' : sh.valign === 'bottom' ? 'bottom' : 'top',
        fill,
        line,
        fontSize: 14,
        color: dark ? 'FFFFFF' : '333333',
        margin: 4,
        shadow: sh.shadow ? { type: 'outer', blur: 6, offset: 2, angle: 45, opacity: 0.3 } : undefined,
      });
    }
  }
  const out = (await pptx.write({ outputType: 'arraybuffer' })) as ArrayBuffer;
  return new Uint8Array(out);
}

/** Shapes for a legacy slide that only has title/bullets (kept for old saved decks). */
function legacyShapes(s: DeckSlide, deckTitle: string): SlideShape[] {
  const dark = isDarkColor(s.bg);
  const accent = s.accent ?? '#C43E1C';
  const main = dark ? '#FFFFFF' : '#1F2430';
  const body = dark ? '#E8E8E8' : '#333333';
  const P = (text: string, o: Partial<ShapeRun> & { align?: ShapePara['align']; bullet?: boolean }): ShapePara => ({ align: o.align, bullet: o.bullet, runs: [{ text, sz: o.sz, b: o.b, color: o.color }] });
  const out: SlideShape[] = [];
  const lines = s.bullets.filter((b) => b.trim());
  if (s.layout === 'title') {
    out.push({ x: 0.5, y: 2.1, w: 9, h: 1.3, kind: 'text', paras: [P(s.title || deckTitle, { sz: 34, b: true, color: main, align: 'center' })] });
    if (lines[0]) out.push({ x: 0.5, y: 3.45, w: 9, h: 0.7, kind: 'text', paras: [P(lines[0], { sz: 17, color: accent, align: 'center' })] });
  } else if (s.layout === 'section') {
    out.push({ x: 0, y: 2.3, w: 10, h: 1.05, kind: 'shape', paras: [], fill: accent });
    out.push({ x: 0.5, y: 2.34, w: 9, h: 0.97, kind: 'text', paras: [P(s.title || 'Section', { sz: 27, b: true, color: '#FFFFFF', align: 'center' })] });
  } else if (s.layout !== 'blank') {
    out.push({ x: 0.6, y: 0.35, w: 8.8, h: 1.1, kind: 'text', paras: [P(s.title || 'Slide', { sz: 27, b: true, color: main })] });
    out.push({ x: 0.62, y: 1.42, w: 1.6, h: 0.06, kind: 'shape', paras: [], fill: accent });
    if (lines.length) out.push({ x: 0.8, y: 1.75, w: s.image ? 4.9 : 8.3, h: 3.6, kind: 'text', paras: lines.map((t) => P(t, { sz: 15, color: body, bullet: true })) });
  }
  if (s.image) out.push(s.layout === 'blank' ? { x: 0.5, y: 0.4, w: 9, h: 4.85, kind: 'image', paras: [], img: s.image } : { x: 5.85, y: 1.8, w: 3.6, h: 2.7, kind: 'image', paras: [], img: s.image });
  return out;
}

// ---------------------------------------------------------------------------
// Slides → raster / PDF (renders the shared SlideView offscreen)
// ---------------------------------------------------------------------------

async function renderSlideCanvas(slide: DeckSlide, widthPx: number): Promise<HTMLCanvasElement> {
  const [{ createRoot }, React, { SlideView }, { default: html2canvas }] = await Promise.all([
    import('react-dom/client'),
    import('react'),
    import('../components/SlideView'),
    import('html2canvas'),
  ]);
  const host = document.createElement('div');
  host.style.cssText = 'position:fixed;left:-20000px;top:0;';
  document.body.appendChild(host);
  const root = createRoot(host);
  try {
    await new Promise<void>((resolve) => {
      root.render(React.createElement(SlideView, { slide, width: widthPx }));
      requestAnimationFrame(() => requestAnimationFrame(() => resolve()));
    });
    // wait for images inside the slide
    await Promise.all(Array.from(host.querySelectorAll('img')).map((im) => (im.complete ? Promise.resolve() : new Promise<void>((r) => { im.onload = () => r(); im.onerror = () => r(); }))));
    const el = host.firstElementChild as HTMLElement;
    return await html2canvas(el, { scale: 1, useCORS: true, logging: false, backgroundColor: slide.bg ?? '#ffffff' });
  } finally {
    root.unmount();
    host.remove();
  }
}

/** One slide → PNG bytes. */
export async function slideToPng(slide: DeckSlide, widthPx = 1600): Promise<Uint8Array> {
  const canvas = await renderSlideCanvas(slide, widthPx);
  const blob = await new Promise<Blob | null>((res) => canvas.toBlob(res, 'image/png'));
  if (!blob) throw new Error('PNG encode failed');
  return new Uint8Array(await blob.arrayBuffer());
}

/** Whole deck → landscape PDF (one slide per page; optional notes handout layout). */
export async function exportSlidesPdf(title: string, slides: DeckSlide[], opts: { notes?: boolean } = {}): Promise<Uint8Array> {
  const { jsPDF } = await import('jspdf');
  const cw = slides[0]?.cw ?? 10;
  const ch = slides[0]?.ch ?? 5.63;
  const pdf = opts.notes ? new jsPDF({ orientation: 'portrait', unit: 'in', format: 'a4', compress: true }) : new jsPDF({ orientation: 'landscape', unit: 'in', format: [cw, ch], compress: true });
  const pageW = pdf.internal.pageSize.getWidth();
  const pageH = pdf.internal.pageSize.getHeight();
  for (let i = 0; i < slides.length; i++) {
    if (i > 0) pdf.addPage();
    const canvas = await renderSlideCanvas(slides[i], 1400);
    const img = canvas.toDataURL('image/jpeg', 0.9);
    if (opts.notes) {
      const w = pageW - 1.2;
      const h = (w * ch) / cw;
      pdf.addImage(img, 'JPEG', 0.6, 0.6, w, h);
      pdf.setDrawColor(180);
      pdf.rect(0.6, 0.6, w, h);
      pdf.setFontSize(11);
      pdf.setTextColor(40);
      const notes = slides[i].notes?.trim() || '(no notes)';
      const lines = pdf.splitTextToSize(notes, w) as string[];
      pdf.text(lines.slice(0, 40), 0.6, 0.6 + h + 0.45);
      pdf.setFontSize(9);
      pdf.setTextColor(120);
      pdf.text(`${title} — slide ${i + 1} of ${slides.length}`, pageW / 2, pageH - 0.4, { align: 'center' });
    } else {
      pdf.addImage(img, 'JPEG', 0, 0, pageW, pageH);
    }
  }
  pdf.setProperties({ title, creator: 'GenOffice' });
  return new Uint8Array(pdf.output('arraybuffer'));
}
