/**
 * Cell-style round trip for .xlsx files.
 *
 * The community build of SheetJS neither writes fonts/fills/borders/alignment
 * nor freeze panes, and on read only exposes fill colours. These helpers patch
 * the OOXML parts directly (JSZip + regex/DOMParser) so bold, italic, colours,
 * fills, borders, alignment, wrapping, font size and frozen panes survive a
 * save → open cycle, in Excel as well as in GenOffice.
 */
import type { CellStyle, HAlign, Page, VAlign } from './sheet-model';

const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const NS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** Legacy indexed palette (0–63). 64/65 are system colours and intentionally absent. */
const INDEXED = (
  '000000 FFFFFF FF0000 00FF00 0000FF FFFF00 FF00FF 00FFFF 000000 FFFFFF FF0000 00FF00 0000FF FFFF00 FF00FF 00FFFF ' +
  '800000 008000 000080 808000 800080 008080 C0C0C0 808080 9999FF 993366 FFFFCC CCFFFF 660066 FF8080 0066CC CCCCFF ' +
  '000080 FF00FF FFFF00 00FFFF 800080 800000 008080 0000FF 00CCFF CCFFFF CCFFCC FFFF99 99CCFF FF99CC CC99FF FFCC99 ' +
  '3366FF 33CCCC 99CC00 FFCC00 FF9900 FF6600 666699 969696 003366 339966 003300 333300 993300 993366 333399 333333'
).split(' ');

function esc(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function colName(i: number): string {
  let s = '';
  for (let n = i + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
}

function colIdx(col: string): number {
  let n = 0;
  for (const c of col) n = n * 26 + (c.charCodeAt(0) - 64);
  return n - 1;
}

// ---------------------------------------------------------------------------
// colour helpers
// ---------------------------------------------------------------------------

function hexToRgb(hex: string): [number, number, number] {
  const h = hex.replace('#', '').slice(-6);
  return [parseInt(h.slice(0, 2), 16), parseInt(h.slice(2, 4), 16), parseInt(h.slice(4, 6), 16)];
}

function rgbToHex([r, g, b]: [number, number, number]): string {
  return `#${[r, g, b].map((v) => Math.max(0, Math.min(255, Math.round(v))).toString(16).padStart(2, '0')).join('').toUpperCase()}`;
}

/** Apply an OOXML theme tint (-1..1) to a colour, ECMA-376 §18.8.19. */
function applyTint(hex: string, tint: number): string {
  if (!tint) return hex;
  const [r, g, b] = hexToRgb(hex).map((v) => v / 255) as [number, number, number];
  const max = Math.max(r, g, b), min = Math.min(r, g, b);
  let h = 0, s = 0;
  let l = (max + min) / 2;
  if (max !== min) {
    const d = max - min;
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
    h /= 6;
  }
  l = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;
  const hue2rgb = (p: number, q: number, t: number) => {
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  if (s === 0) return rgbToHex([l * 255, l * 255, l * 255]);
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  return rgbToHex([hue2rgb(p, q, h + 1 / 3) * 255, hue2rgb(p, q, h) * 255, hue2rgb(p, q, h - 1 / 3) * 255]);
}

function readColor(el: Element | null, theme: string[]): string | undefined {
  if (!el) return undefined;
  if (el.getAttribute('auto') === '1') return undefined;
  const tint = parseFloat(el.getAttribute('tint') ?? '0') || 0;
  const rgb = el.getAttribute('rgb');
  if (rgb) return applyTint(`#${rgb.slice(-6).toUpperCase()}`, tint);
  const th = el.getAttribute('theme');
  if (th !== null && theme[+th]) return applyTint(theme[+th], tint);
  const idx = el.getAttribute('indexed');
  if (idx !== null && INDEXED[+idx]) return applyTint(`#${INDEXED[+idx]}`, tint);
  return undefined;
}

function firstNS(el: Element | Document, tag: string): Element | null {
  return el.getElementsByTagNameNS('*', tag)[0] ?? null;
}

function childrenNS(el: Element, tag: string): Element[] {
  return Array.from(el.getElementsByTagNameNS('*', tag)).filter((c) => c.parentNode === el);
}

// ---------------------------------------------------------------------------
// package helpers
// ---------------------------------------------------------------------------

interface ZipLike {
  file(path: string): { async(type: 'string'): Promise<string> } | null;
  file(path: string, data: string): unknown;
}

/** sheet display name -> zip path of its worksheet part */
async function sheetParts(zip: ZipLike): Promise<Map<string, string>> {
  const out = new Map<string, string>();
  const wbXml = await zip.file('xl/workbook.xml')?.async('string');
  const relXml = await zip.file('xl/_rels/workbook.xml.rels')?.async('string');
  if (!wbXml || !relXml) return out;
  const rels = new Map<string, string>();
  for (const m of relXml.matchAll(/<Relationship\b([^>]*)\/?>/g)) {
    const id = /\bId="([^"]+)"/.exec(m[1])?.[1];
    const target = /\bTarget="([^"]+)"/.exec(m[1])?.[1];
    if (id && target) rels.set(id, target.startsWith('/') ? target.slice(1) : `xl/${target.replace(/^\.\//, '')}`);
  }
  const wb = new DOMParser().parseFromString(wbXml, 'application/xml');
  for (const sh of Array.from(wb.getElementsByTagNameNS('*', 'sheet'))) {
    const name = sh.getAttribute('name') ?? '';
    const rid = sh.getAttributeNS(NS_REL, 'id') ?? sh.getAttribute('r:id') ?? '';
    const path = rels.get(rid);
    if (name && path) out.set(name, path);
  }
  return out;
}

async function themeColors(zip: ZipLike): Promise<string[]> {
  try {
    const xml = await zip.file('xl/theme/theme1.xml')?.async('string');
    if (!xml) return [];
    const doc = new DOMParser().parseFromString(xml, 'application/xml');
    const scheme = firstNS(doc, 'clrScheme');
    if (!scheme) return [];
    const order = ['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
    const raw = order.map((tag) => {
      const el = firstNS(scheme, tag);
      const c = el?.getElementsByTagNameNS('*', 'srgbClr')[0]?.getAttribute('val') ?? el?.getElementsByTagNameNS('*', 'sysClr')[0]?.getAttribute('lastClr') ?? '000000';
      return `#${c.toUpperCase()}`;
    });
    // theme indices swap the dk/lt pairs: 0=lt1 1=dk1 2=lt2 3=dk2
    return [raw[1], raw[0], raw[3], raw[2], ...raw.slice(4)];
  } catch {
    return [];
  }
}

// ---------------------------------------------------------------------------
// import: styles.xml + sheet xml -> CellStyle map, freeze panes
// ---------------------------------------------------------------------------

export interface SheetExtras {
  styles: Record<string, CellStyle>;
  freeze?: { r: number; c: number };
}

/** Read cell formatting and frozen panes straight from the OOXML parts. Returns null for non-xlsx input. */
export async function readXlsxExtras(buf: ArrayBuffer, maxRows = 1000, maxCols = 100): Promise<Record<string, SheetExtras> | null> {
  const bytes = new Uint8Array(buf.slice(0, 4));
  if (bytes[0] !== 0x50 || bytes[1] !== 0x4b) return null; // not a zip
  try {
    const JSZip = (await import('jszip')).default;
    const zip = (await JSZip.loadAsync(buf)) as unknown as ZipLike;
    const stylesXml = await zip.file('xl/styles.xml')?.async('string');
    const parts = await sheetParts(zip);
    if (!parts.size) return null;
    const theme = await themeColors(zip);

    // ---- styles.xml -> xf index -> CellStyle
    const xfStyles: CellStyle[] = [];
    if (stylesXml) {
      const doc = new DOMParser().parseFromString(stylesXml, 'application/xml');
      const fontsEl = firstNS(doc, 'fonts');
      const fonts = fontsEl ? childrenNS(fontsEl, 'font') : [];
      const defaultSz = parseFloat(firstNS(fonts[0] ?? doc.createElement('x'), 'sz')?.getAttribute('val') ?? '11') || 11;
      const fontStyles = fonts.map((f) => {
        const st: CellStyle = {};
        if (firstNS(f, 'b') && firstNS(f, 'b')!.getAttribute('val') !== '0') st.b = true;
        if (firstNS(f, 'i') && firstNS(f, 'i')!.getAttribute('val') !== '0') st.i = true;
        if (firstNS(f, 'u') && firstNS(f, 'u')!.getAttribute('val') !== 'none') st.u = true;
        if (firstNS(f, 'strike') && firstNS(f, 'strike')!.getAttribute('val') !== '0') st.s = true;
        const color = readColor(firstNS(f, 'color'), theme);
        if (color && color !== '#000000') st.color = color;
        const sz = parseFloat(firstNS(f, 'sz')?.getAttribute('val') ?? '') || defaultSz;
        if (Math.abs(sz - defaultSz) >= 1) st.size = Math.round((sz * 4) / 3);
        return st;
      });
      const fillsEl = firstNS(doc, 'fills');
      const fills = (fillsEl ? childrenNS(fillsEl, 'fill') : []).map((f) => {
        const pf = firstNS(f, 'patternFill');
        const type = pf?.getAttribute('patternType') ?? 'none';
        if (!pf || type === 'none') return undefined;
        const c = readColor(firstNS(pf, 'fgColor'), theme) ?? readColor(firstNS(pf, 'bgColor'), theme);
        return c && c !== '#FFFFFF' ? c : undefined;
      });
      const bordersEl = firstNS(doc, 'borders');
      const borders = (bordersEl ? childrenNS(bordersEl, 'border') : []).map((b) => {
        const has = (side: string) => {
          const el = firstNS(b, side);
          const style = el?.getAttribute('style');
          return !!style && style !== 'none';
        };
        return { bt: has('top'), bb: has('bottom'), bl: has('left'), br: has('right') };
      });
      const xfsEl = firstNS(doc, 'cellXfs');
      for (const xf of xfsEl ? childrenNS(xfsEl, 'xf') : []) {
        const st: CellStyle = {};
        const fontId = +(xf.getAttribute('fontId') ?? 0);
        if (xf.getAttribute('applyFont') !== '0' && fontStyles[fontId]) Object.assign(st, fontStyles[fontId]);
        const fillId = +(xf.getAttribute('fillId') ?? 0);
        if (fills[fillId]) st.fill = fills[fillId];
        const borderId = +(xf.getAttribute('borderId') ?? 0);
        const bd = borders[borderId];
        if (bd) {
          if (bd.bt) st.bt = true;
          if (bd.bb) st.bb = true;
          if (bd.bl) st.bl = true;
          if (bd.br) st.br = true;
        }
        const al = firstNS(xf, 'alignment');
        if (al) {
          const h = al.getAttribute('horizontal');
          if (h === 'center' || h === 'right' || h === 'left' || h === 'centerContinuous') st.align = (h === 'centerContinuous' ? 'center' : h) as HAlign;
          const v = al.getAttribute('vertical');
          if (v === 'top' || v === 'center') st.valign = (v === 'center' ? 'middle' : v) as VAlign;
          const wrap = al.getAttribute('wrapText');
          if (wrap === '1' || wrap === 'true') st.wrap = true;
        }
        xfStyles.push(st);
      }
    }

    // ---- each worksheet: <c s=".."> and <pane state="frozen">
    const out: Record<string, SheetExtras> = {};
    for (const [name, path] of parts) {
      const xml = await zip.file(path)?.async('string');
      if (!xml) continue;
      const extras: SheetExtras = { styles: {} };
      const pane = /<pane\b([^>]*)\/?>/.exec(xml);
      if (pane && /state="frozen(?:Split)?"/.test(pane[1])) {
        const c = parseInt(/\bxSplit="(\d+)"/.exec(pane[1])?.[1] ?? '0', 10) || 0;
        const r = parseInt(/\bySplit="(\d+)"/.exec(pane[1])?.[1] ?? '0', 10) || 0;
        if (r > 0 || c > 0) extras.freeze = { r, c };
      }
      if (xfStyles.length) {
        for (const m of xml.matchAll(/<c r="([A-Z]+)(\d+)"([^>]*?)\/?>/g)) {
          const s = /\bs="(\d+)"/.exec(m[3]);
          if (!s) continue;
          const st = xfStyles[+s[1]];
          if (!st || !Object.keys(st).length) continue;
          const r = parseInt(m[2], 10) - 1;
          const c = colIdx(m[1]);
          if (r >= maxRows || c >= maxCols) continue;
          extras.styles[`${m[1]}${m[2]}`] = { ...st };
        }
      }
      out[name] = extras;
    }
    return out;
  } catch {
    return null;
  }
}

// ---------------------------------------------------------------------------
// export: patch styles.xml + worksheet xml written by SheetJS
// ---------------------------------------------------------------------------

function fontKey(st: CellStyle): string {
  return [st.b ? 'b' : '', st.i ? 'i' : '', st.u ? 'u' : '', st.s ? 's' : '', st.color ?? '', st.size ?? ''].join('|');
}

function fontXml(st: CellStyle, defaultPt: number): string {
  const pt = st.size ? Math.round((st.size * 3) / 4) : defaultPt;
  return (
    '<font>' +
    (st.b ? '<b/>' : '') +
    (st.i ? '<i/>' : '') +
    (st.u ? '<u/>' : '') +
    (st.s ? '<strike/>' : '') +
    `<sz val="${pt}"/>` +
    (st.color ? `<color rgb="FF${st.color.replace('#', '').toUpperCase()}"/>` : '<color theme="1"/>') +
    '<name val="Calibri"/><family val="2"/><scheme val="minor"/></font>'
  );
}

function borderKey(st: CellStyle): string {
  return `${st.bt ? 1 : 0}${st.bb ? 1 : 0}${st.bl ? 1 : 0}${st.br ? 1 : 0}`;
}

function borderXml(st: CellStyle): string {
  const side = (tag: string, on: boolean | undefined) => (on ? `<${tag} style="thin"><color auto="1"/></${tag}>` : `<${tag}/>`);
  return `<border>${side('left', st.bl)}${side('right', st.br)}${side('top', st.bt)}${side('bottom', st.bb)}<diagonal/></border>`;
}

function hasVisualStyle(st: CellStyle | undefined): st is CellStyle {
  return !!st && !!(st.b || st.i || st.u || st.s || st.color || st.size || st.fill || st.bt || st.bb || st.bl || st.br || st.align || st.valign || st.wrap);
}

/**
 * Inject fonts/fills/borders/alignment and freeze panes into an .xlsx produced by
 * SheetJS. `sheets` maps the *written* sheet names to their pages. Falls back to
 * the untouched bytes if anything about the package looks unexpected.
 */
export async function applyXlsxExtras(bytes: Uint8Array, sheets: { name: string; page: Page }[]): Promise<Uint8Array> {
  const needed = sheets.some(({ page }) => page.freeze || Object.values(page.styles).some(hasVisualStyle));
  if (!needed) return bytes;
  try {
    const JSZip = (await import('jszip')).default;
    const zip = (await JSZip.loadAsync(bytes)) as unknown as ZipLike & { generateAsync(o: unknown): Promise<Uint8Array> };
    const stylesXml = await zip.file('xl/styles.xml')?.async('string');
    const parts = await sheetParts(zip);
    if (!stylesXml || !parts.size) return bytes;

    // existing tables (SheetJS writes one default font/border and two default fills)
    const count = (tag: string) => {
      const m = new RegExp(`<${tag}\\b[^>]*count="(\\d+)"`).exec(stylesXml);
      return m ? parseInt(m[1], 10) : 0;
    };
    const baseFonts = Math.max(1, count('fonts'));
    const defaultPt = parseFloat(/<fonts\b[^>]*>\s*<font>[\s\S]*?<sz val="([\d.]+)"/.exec(stylesXml)?.[1] ?? '11') || 11;
    const baseFills = Math.max(2, count('fills'));
    const baseBorders = Math.max(1, count('borders'));
    const xfsBlock = /<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(stylesXml);
    if (!xfsBlock) return bytes;
    const baseXfs = Array.from(xfsBlock[1].matchAll(/<xf\b([^>]*?)(?:\/>|>[\s\S]*?<\/xf>)/g)).map((m) => m[1]);
    const numFmtOf = (xfIdx: number) => parseInt(/\bnumFmtId="(\d+)"/.exec(baseXfs[xfIdx] ?? '')?.[1] ?? '0', 10) || 0;

    const fonts: string[] = [], fontIds = new Map<string, number>();
    const fills: string[] = [], fillIds = new Map<string, number>();
    const borders: string[] = [], borderIds = new Map<string, number>();
    const xfs: string[] = [], xfIds = new Map<string, number>();

    const xfFor = (st: CellStyle, numFmtId: number): number => {
      const fk = fontKey(st);
      let fontId = 0;
      if (fk !== '|||||') {
        if (!fontIds.has(fk)) { fontIds.set(fk, baseFonts + fonts.length); fonts.push(fontXml(st, defaultPt)); }
        fontId = fontIds.get(fk)!;
      }
      let fillId = 0;
      if (st.fill) {
        const key = st.fill.toUpperCase();
        if (!fillIds.has(key)) { fillIds.set(key, baseFills + fills.length); fills.push(`<fill><patternFill patternType="solid"><fgColor rgb="FF${key.replace('#', '')}"/><bgColor indexed="64"/></patternFill></fill>`); }
        fillId = fillIds.get(key)!;
      }
      let borderId = 0;
      const bk = borderKey(st);
      if (bk !== '0000') {
        if (!borderIds.has(bk)) { borderIds.set(bk, baseBorders + borders.length); borders.push(borderXml(st)); }
        borderId = borderIds.get(bk)!;
      }
      const h = st.align, v = st.valign === 'middle' ? 'center' : st.valign, wrap = st.wrap;
      const key = [numFmtId, fontId, fillId, borderId, h ?? '', v ?? '', wrap ? 1 : 0].join('/');
      if (!xfIds.has(key)) {
        const align = h || v || wrap ? `<alignment${h ? ` horizontal="${h}"` : ''}${v ? ` vertical="${v}"` : ''}${wrap ? ' wrapText="1"' : ''}/>` : '';
        xfIds.set(key, baseXfs.length + xfs.length);
        xfs.push(
          `<xf numFmtId="${numFmtId}" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}" xfId="0"` +
            `${numFmtId ? ' applyNumberFormat="1"' : ''}${fontId ? ' applyFont="1"' : ''}${fillId ? ' applyFill="1"' : ''}${borderId ? ' applyBorder="1"' : ''}${align ? ' applyAlignment="1"' : ''}` +
            (align ? `>${align}</xf>` : '/>'),
        );
      }
      return xfIds.get(key)!;
    };

    for (const { name, page } of sheets) {
      const path = parts.get(name);
      if (!path) continue;
      let xml = await zip.file(path)?.async('string');
      if (!xml) continue;
      const styled = Object.entries(page.styles).filter(([, st]) => hasVisualStyle(st));
      if (styled.length) {
        // rebuild <sheetData> so styled-but-empty cells get a stub <c/> too
        const sd = /<sheetData>([\s\S]*?)<\/sheetData>|<sheetData\/>/.exec(xml);
        if (sd) {
          const rows = new Map<number, { attrs: string; cells: Map<number, string> }>();
          for (const rm of (sd[1] ?? '').matchAll(/<row\b([^>]*)>([\s\S]*?)<\/row>|<row\b([^>]*)\/>/g)) {
            const attrs = rm[1] ?? rm[3] ?? '';
            const r = parseInt(/\br="(\d+)"/.exec(attrs)?.[1] ?? '0', 10);
            const cells = new Map<number, string>();
            for (const cm of (rm[2] ?? '').matchAll(/<c r="([A-Z]+)\d+"[^>]*?(?:\/>|>[\s\S]*?<\/c>)/g)) cells.set(colIdx(cm[1]), cm[0]);
            rows.set(r, { attrs, cells });
          }
          for (const [addr, st] of styled) {
            const m = /^([A-Z]+)(\d+)$/.exec(addr);
            if (!m) continue;
            const r = parseInt(m[2], 10), c = colIdx(m[1]);
            let row = rows.get(r);
            if (!row) { row = { attrs: ` r="${r}"`, cells: new Map() }; rows.set(r, row); }
            const existing = row.cells.get(c);
            const cur = existing ? parseInt(/^<c [^>]*?\bs="(\d+)"/.exec(existing)?.[1] ?? '0', 10) : 0;
            const s = xfFor(st, numFmtOf(cur));
            if (existing) row.cells.set(c, /^<c [^>]*?\bs="\d+"/.test(existing) ? existing.replace(/^(<c [^>]*?\bs=")\d+"/, `$1${s}"`) : existing.replace(/^<c r="([A-Z]+\d+)"/, `<c r="$1" s="${s}"`));
            else row.cells.set(c, `<c r="${addr}" s="${s}"/>`);
          }
          const body = Array.from(rows.keys())
            .sort((a, b) => a - b)
            .map((r) => {
              const row = rows.get(r)!;
              const cells = Array.from(row.cells.keys()).sort((a, b) => a - b).map((c) => row.cells.get(c)!).join('');
              const attrs = /\br="\d+"/.test(row.attrs) ? row.attrs : ` r="${r}"${row.attrs}`;
              return `<row${attrs}>${cells}</row>`;
            })
            .join('');
          xml = xml.replace(sd[0], `<sheetData>${body}</sheetData>`);
        }
      }
      if (page.freeze && (page.freeze.r > 0 || page.freeze.c > 0)) {
        const { r, c } = page.freeze;
        const topLeft = `${colName(c)}${r + 1}`;
        const activePane = r > 0 && c > 0 ? 'bottomRight' : r > 0 ? 'bottomLeft' : 'topRight';
        const pane = `<pane${c > 0 ? ` xSplit="${c}"` : ''}${r > 0 ? ` ySplit="${r}"` : ''} topLeftCell="${topLeft}" activePane="${activePane}" state="frozen"/><selection pane="${activePane}" activeCell="${topLeft}" sqref="${topLeft}"/>`;
        xml = /<sheetView\b[^>]*\/>/.test(xml)
          ? xml.replace(/<sheetView\b([^>]*)\/>/, `<sheetView$1>${pane}</sheetView>`)
          : xml.replace(/<sheetView\b([^>]*)>/, `<sheetView$1>${pane}`);
      }
      zip.file(path, xml);
    }

    if (fonts.length || fills.length || borders.length || xfs.length) {
      let sty = stylesXml;
      const extend = (tag: string, add: string[], base: number, fallback: string) => {
        if (!add.length) return;
        const re = new RegExp(`<${tag}\\b[^>]*>([\\s\\S]*?)</${tag}>`);
        if (re.test(sty)) sty = sty.replace(re, (_m, inner: string) => `<${tag} count="${base + add.length}">${inner}${add.join('')}</${tag}>`);
        else sty = sty.replace(/<cellStyleXfs\b/, `<${tag} count="${base + add.length}">${fallback}${add.join('')}</${tag}><cellStyleXfs`);
      };
      extend('fonts', fonts, baseFonts, '<font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>');
      extend('fills', fills, baseFills, '<fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>');
      extend('borders', borders, baseBorders, '<border><left/><right/><top/><bottom/><diagonal/></border>');
      sty = sty.replace(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/, (_m, inner: string) => `<cellXfs count="${baseXfs.length + xfs.length}">${inner}${xfs.join('')}</cellXfs>`);
      zip.file('xl/styles.xml', sty);
    }
    return await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  } catch {
    return bytes;
  }
}

export { esc as _escXml, NS_MAIN as _XLSX_NS_MAIN };
