/**
 * Validates shape-level PPTX import/export:
 * 1. Builds a synthetic .pptx (raw OOXML) with background, positioned text
 *    boxes (bold/sized/colored runs, bullets, centered paragraph), a filled
 *    rectangle and an embedded picture.
 * 2. Runs importPptx and asserts geometry, styles and image parsing.
 * 3. Round-trips the imported deck through exportPptx and re-imports,
 *    asserting shape count and positions survive.
 */
import { buildPptx } from './synthetic-pptx.mjs';

import { build } from 'esbuild';
import { unlinkSync } from 'node:fs';
import path from 'node:path';

// Bundle src/lib/fileio.ts for Node. The pptx code paths only need DOMParser/XMLSerializer,
// which we provide through a small shim (linkedom) when running outside the browser.
const bundle = path.resolve('scripts/.fileio-bundle.mjs'); // inside the repo so bare imports (pptxgenjs, xlsx…) resolve from node_modules
process.on('exit', () => { try { unlinkSync(bundle); } catch {} });
await build({
  entryPoints: ['src/lib/fileio.ts'],
  bundle: true,
  format: 'esm',
  platform: 'node',
  outfile: bundle,
  logLevel: 'silent',
  external: ['react', 'react-dom/client', '../components/SlideView', 'html2canvas', 'jspdf', 'docx', 'mammoth', 'xlsx', 'pptxgenjs', 'jszip', '@capacitor/*'],
  alias: { '@capacitor/core': path.resolve('scripts/shims/capacitor-core.mjs'), '@capacitor/filesystem': path.resolve('scripts/shims/capacitor-core.mjs') },
});
if (typeof globalThis.DOMParser === 'undefined') {
  const { DOMParser, XMLSerializer } = await import('@xmldom/xmldom').catch(() => ({}));
  if (DOMParser) { globalThis.DOMParser = DOMParser; globalThis.XMLSerializer = XMLSerializer; }
}
const { importPptx, exportPptx } = await import(bundle);

let failures = 0;
function check(name, cond, detail = '') {
  if (cond) console.log(`  PASS ${name}`);
  else {
    failures++;
    console.error(`  FAIL ${name} ${detail}`);
  }
}
const near = (a, b, eps = 0.03) => Math.abs(a - b) <= eps;

// --- 1. parse synthetic pptx -------------------------------------------------
const buf = await buildPptx();
const slides = await importPptx(buf);
console.log('Import results:');
check('one slide parsed', slides.length === 1);
const s = slides[0];
check('background #1E2430', s.bg?.toUpperCase() === '#1E2430', `got ${s.bg}`);
check('canvas size 10x5.63', near(s.cw ?? 10, 10) && near(s.ch ?? 5.63, 5.625, 0.02), `cw=${s.cw} ch=${s.ch}`);
const sh = s.shapes ?? [];
console.log(`  shapes: ${sh.length} -> ${sh.map((x) => x.kind).join(', ')}`);
check('four shapes', sh.length === 4, `got ${sh.length}`);

const title = sh[0];
check('title text parsed', title.paras[0]?.runs[0]?.text === 'Quarterly Review', JSON.stringify(title.paras));
check('title font 44pt', title.paras[0]?.runs[0]?.sz === 44, `${title.paras[0]?.runs[0]?.sz}`);
check('title bold', title.paras[0]?.runs[0]?.b === true);
check('title white', title.paras[0]?.runs[0]?.color?.toUpperCase() === '#FFFFFF');
// 914400 EMU * 0.75 scale = 0.75in
check('title x scaled to 0.75in', near(title.x, 0.75), `${title.x}`);
check('title w scaled to 6.75in', near(title.w, 6.75), `${title.w}`);

const body = sh[1];
check('body has 2 paragraphs', body.paras.length === 2, `${body.paras.length}`);
check('bullet detected', body.paras[0]?.bullet === true);
check('centered para', body.paras[0]?.align === 'center' || body.paras[1]?.align === 'center');
check('body run colored', (body.paras[0]?.runs[0]?.color ?? '').toUpperCase() === '#FFC000');

const rect = sh[2];
check('rect fill #C43E1C', rect.fill?.toUpperCase() === '#C43E1C', `got ${rect.fill}`);
check('rect is shape kind', rect.kind === 'shape');

const pic = sh[3];
check('picture parsed as image', pic.kind === 'image');
check('picture data URL', (pic.img ?? '').startsWith('data:image/png;base64,'));

// --- 2. round-trip export -> reimport ----------------------------------------
console.log('Round-trip export -> reimport:');
const out = await exportPptx('Test deck', slides);
const re = await importPptx(out.buffer.slice(out.byteOffset, out.byteOffset + out.byteLength));
check('reimport 1 slide', re.length === 1);
const rs = re[0];
check('reimport keeps 4 shapes', (rs.shapes ?? []).length === 4, `got ${(rs.shapes ?? []).length}`);
const rt = (rs.shapes ?? [])[0];
if (rt) {
  check('reimport title text', rt.paras[0]?.runs[0]?.text === 'Quarterly Review');
  check('reimport title near-original x', near(rt.x, title.x, 0.06), `${rt.x} vs ${title.x}`);
  check('reimport title bold+size', rt.paras[0]?.runs[0]?.b === true && rt.paras[0]?.runs[0]?.sz === 44);
}

if (failures) {
  console.error(`\n${failures} check(s) FAILED`);
  process.exit(1);
}
console.log('\nAll PPTX shape checks passed.');
