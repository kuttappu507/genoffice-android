/**
 * File I/O round-trip checks (run: node scripts/test-io.mjs)
 *
 *   docx  : export -> import keeps headings, inline formatting, links, lists,
 *           tables, footnotes, page breaks and paragraph alignment
 *   md    : htmlToMarkdown output
 *   xlsx  : export -> import keeps sheet order, formulas, merges, column
 *           widths, number formats, fonts/fills/borders/alignment, freeze panes
 *   xlsx  : an Excel-authored package (theme colours + tints, indexed colours,
 *           shared strings) imports with styles and frozen rows
 *   pptx  : export -> import keeps titles/bodies/notes/backgrounds
 *
 * Runs src/lib/fileio.ts under Node with a jsdom window; no browser needed.
 */
import { build } from 'esbuild';
import { unlinkSync } from 'node:fs';
import path from 'node:path';
import JSZip from 'jszip';
import { JSDOM } from 'jsdom';

const bundle = path.resolve('scripts/.fileio-bundle.mjs');
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
const dom = new JSDOM('<!doctype html><html><body></body></html>');
for (const k of ['window', 'document', 'DOMParser', 'XMLSerializer', 'Node', 'HTMLElement', 'Blob']) globalThis[k] = dom.window[k];
const io = await import(bundle);

let fails = 0;
const check = (name, cond, detail = '') => {
  if (!cond) fails++;
  console.log(`${cond ? 'PASS' : 'FAIL'} ${name}${cond || !detail ? '' : `   (${detail})`}`);
};
const toAB = (u8) => u8.buffer.slice(u8.byteOffset, u8.byteOffset + u8.byteLength);

// ---------------------------------------------------------------------------
console.log('DOCX');
const docHtml =
  '<h1>Title here</h1>' +
  '<p>Plain <b>bold</b> and <i>italic</i> text with <a href="https://x.io">link</a>.</p>' +
  '<ul><li>one</li><li style="text-align:center">two</li></ul>' +
  '<table><tr><td><b>H1</b></td><td><b>H2</b></td></tr><tr><td>c1</td><td>c2</td></tr></table>' +
  '<p style="text-align:center">centered</p>' +
  '<hr class="page-break">' +
  '<p style="text-align:right">Second page<sup class="fn">[1]</sup></p>' +
  '<p style="text-align:justify">justified</p>' +
  '<p class="footnote">[1] a footnote</p>';
const docxBlob = await io.exportDocx('Test', docHtml, { page: { size: 'A4', orient: 'portrait', margins: 'normal' }, font: 'Calibri', lineSpacing: 1.15 });
const docxBytes = new Uint8Array(await docxBlob.arrayBuffer());
check('docx export is a zip', docxBytes[0] === 0x50 && docxBytes[1] === 0x4b && docxBytes.length > 2000, `${docxBytes.length} bytes`);
const back = await io.importDocx(toAB(docxBytes));
check('h1 survives', /<h1>Title here<\/h1>/.test(back), back);
check('bold survives', /<strong>bold<\/strong>/.test(back));
check('italic survives', /<em>italic<\/em>/.test(back));
check('link survives', /<a href="https:\/\/x\.io">link<\/a>/.test(back));
check('list survives', /<ul><li>one<\/li><li[^>]*>two<\/li><\/ul>/.test(back));
check('list item alignment survives', /<li style="text-align: center;">two<\/li>/.test(back), back);
check('table survives', /<table>[\s\S]*c1[\s\S]*c2[\s\S]*<\/table>/.test(back));
check('centered paragraph survives', /<p style="text-align: center;">centered<\/p>/.test(back), back);
check('page break survives', /<hr class="page-break">/.test(back), back);
check('no empty paragraphs around page break', !/<p><\/p>/.test(back), back);
check('right + justify survive', /text-align: right;">Second page/.test(back) && /text-align: justify;">justified/.test(back), back);
check('footnote text survives', /a footnote/.test(back));

const md = io.htmlToMarkdown(docHtml);
check('markdown heading', md.startsWith('# Title here'));
check('markdown emphasis + link', md.includes('**bold**') && md.includes('*italic*') && md.includes('[link](https://x.io)'));
check('markdown list + table', md.includes('- one') && md.includes('| **H1** | **H2** |') && md.includes('| c1 | c2 |'), md);

// ---------------------------------------------------------------------------
console.log('XLSX');
const book = {
  order: ['Sheet1', 'Data'],
  active: 'Sheet1',
  sheets: {
    Sheet1: {
      cells: { A1: 'Item', B1: 'Qty', C1: 'Price', D1: 'Total', A2: 'Pen', B2: '3', C2: '2.5', D2: '=B2*C2', A3: 'Pad', B3: '4', C3: '10', D3: '=B3*C3', D4: '=SUM(D2:D3)', A8: 'note' },
      styles: { A1: { b: true, fill: '#FFEB3B', align: 'center' }, B1: { b: true, i: true, color: '#C62828' }, D4: { fmt: 'cur', dec: 2 }, C2: { bb: true, wrap: true }, E1: { fill: '#E0E0E0' } },
      merges: ['A6:C6'],
      colW: { 0: 140 },
      freeze: { r: 1, c: 1 },
    },
    Data: { cells: { A1: 'Other', A2: '=Sheet1!D4*2' }, styles: {} },
  },
};
const xlsxBytes = await io.exportXlsx('Test', book);
check('xlsx export is a zip', xlsxBytes[0] === 0x50 && xlsxBytes[1] === 0x4b);
const wb = await io.importWorkbook(toAB(xlsxBytes), 'Test.xlsx');
const s1 = wb.sheets.Sheet1;
check('sheet order', JSON.stringify(wb.order) === '["Sheet1","Data"]', JSON.stringify(wb.order));
check('formulas kept', s1.cells.D2 === '=B2*C2' && s1.cells.D4 === '=SUM(D2:D3)', JSON.stringify(s1.cells));
check('cross-sheet formula kept', wb.sheets.Data.cells.A2 === '=Sheet1!D4*2', wb.sheets.Data.cells.A2);
check('far cell kept', s1.cells.A8 === 'note');
check('merges kept', JSON.stringify(s1.merges) === '["A6:C6"]', JSON.stringify(s1.merges));
check('column width kept', s1.colW?.[0] === 140, JSON.stringify(s1.colW));
check('currency format kept', s1.styles.D4?.fmt === 'cur' && s1.styles.D4?.dec === 2, JSON.stringify(s1.styles.D4));
check('bold + fill + align kept', s1.styles.A1?.b === true && s1.styles.A1?.fill === '#FFEB3B' && s1.styles.A1?.align === 'center', JSON.stringify(s1.styles.A1));
check('italic + colour kept', s1.styles.B1?.b === true && s1.styles.B1?.i === true && s1.styles.B1?.color === '#C62828', JSON.stringify(s1.styles.B1));
check('border + wrap kept', s1.styles.C2?.bb === true && s1.styles.C2?.wrap === true, JSON.stringify(s1.styles.C2));
check('fill on empty cell kept', s1.styles.E1?.fill === '#E0E0E0', JSON.stringify(s1.styles.E1));
check('freeze panes kept', s1.freeze?.r === 1 && s1.freeze?.c === 1, JSON.stringify(s1.freeze));
check('no stray size on default font', s1.styles.A1?.size === undefined, JSON.stringify(s1.styles.A1));

// Excel-authored package: theme colour + tint, indexed colour, shared strings, frozen header row.
const theme =
  '<a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>' +
  '<a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme>';
const zip = new JSZip();
zip.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/><Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>');
zip.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
zip.file('xl/workbook.xml', '<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Budget 2026" sheetId="1" r:id="rId1"/></sheets></workbook>');
zip.file('xl/_rels/workbook.xml.rels', '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/></Relationships>');
zip.file('xl/theme/theme1.xml', `<?xml version="1.0" encoding="UTF-8"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme>${theme}</a:themeElements></a:theme>`);
zip.file('xl/sharedStrings.xml', '<?xml version="1.0" encoding="UTF-8"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3" uniqueCount="3"><si><t>Category</t></si><si><t>Amount</t></si><si><t>Rent</t></si></sst>');
zip.file('xl/styles.xml', '<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="164" formatCode="&quot;₹&quot;#,##0.00"/></numFmts><fonts count="3"><font><sz val="11"/><color theme="1"/><name val="Calibri"/></font><font><b/><sz val="11"/><color theme="0"/><name val="Calibri"/></font><font><i/><sz val="14"/><color indexed="10"/><name val="Calibri"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor theme="4" tint="-0.249977111117893"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"/><right style="thin"/><top style="medium"/><bottom style="thin"/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf><xf numFmtId="164" fontId="2" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1"><alignment horizontal="right"/></xf><xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1"/></cellXfs></styleSheet>');
zip.file('xl/worksheets/sheet1.xml', '<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:C3"/><sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews><sheetData><row r="1"><c r="A1" s="1" t="s"><v>0</v></c><c r="B1" s="1" t="s"><v>1</v></c><c r="C1" s="3"/></row><row r="2"><c r="A2" t="s"><v>2</v></c><c r="B2" s="2"><v>1200</v></c></row><row r="3"><c r="B3" s="2"><f>SUM(B2:B2)</f><v>1200</v></c></row></sheetData></worksheet>');
const excelBytes = await zip.generateAsync({ type: 'uint8array' });
const ewb = await io.importWorkbook(toAB(excelBytes), 'budget.xlsx');
const ep = ewb.sheets['Budget 2026'];
check('excel: sheet name with space', !!ep, JSON.stringify(ewb.order));
check('excel: shared strings + formula', ep?.cells.A2 === 'Rent' && ep?.cells.B3 === '=SUM(B2:B2)', JSON.stringify(ep?.cells));
check('excel: theme fill with tint', ep?.styles.A1?.fill === '#2F5597', JSON.stringify(ep?.styles.A1));
check('excel: white bold header (theme 0)', ep?.styles.A1?.b === true && ep?.styles.A1?.color === '#FFFFFF' && ep?.styles.A1?.valign === 'middle' && ep?.styles.A1?.wrap === true, JSON.stringify(ep?.styles.A1));
check('excel: fill on empty cell', ep?.styles.C1?.fill === '#2F5597', JSON.stringify(ep?.styles.C1));
check('excel: indexed red italic 14pt with borders', ep?.styles.B2?.i === true && ep?.styles.B2?.color === '#FF0000' && ep?.styles.B2?.size === 19 && ep?.styles.B2?.bt && ep?.styles.B2?.bl && ep?.styles.B2?.align === 'right', JSON.stringify(ep?.styles.B2));
check('excel: rupee currency format', ep?.styles.B2?.fmt === 'cur' && ep?.styles.B2?.dec === 2, JSON.stringify(ep?.styles.B2));
check('excel: frozen header row', ep?.freeze?.r === 1 && ep?.freeze?.c === 0, JSON.stringify(ep?.freeze));

// ---------------------------------------------------------------------------
console.log('PPTX');
const dm = await import(path.resolve('scripts/.deck-bundle.mjs')).catch(() => null);
if (!dm) {
  await build({ entryPoints: ['src/lib/deck-model.ts'], bundle: true, format: 'esm', platform: 'node', outfile: path.resolve('scripts/.deck-bundle.mjs'), logLevel: 'silent' });
}
const deck = dm ?? (await import(path.resolve('scripts/.deck-bundle.mjs')));
process.on('exit', () => { try { unlinkSync(path.resolve('scripts/.deck-bundle.mjs')); } catch {} });
const slides = deck.deckFromOutline(
  [
    { title: 'Quarterly Review', bullets: ['FY26 Q2'], notes: 'welcome everyone' },
    { title: 'Highlights', bullets: ['Revenue up 12%', 'Churn down', 'New markets: IN, BR'] },
    { title: 'Roadmap', bullets: [] },
  ],
  deck.THEMES[3],
  'Quarterly Review',
);
slides[1].shapes.push(deck.geomShape('ellipse', 7, 3.5, 2, 1.5, '#FF7043'));
slides[1].transition = 'fade';
const pptxBytes = await io.exportPptx('Deck', slides);
check('pptx export is a zip', pptxBytes[0] === 0x50 && pptxBytes[1] === 0x4b && pptxBytes.length > 5000, `${pptxBytes.length} bytes`);
const backSlides = await io.importPptx(toAB(pptxBytes));
check('slide count', backSlides.length === 3, `${backSlides.length}`);
const t = backSlides.map((s) => deck.slideText(s));
check('titles kept', t[0].title === 'Quarterly Review' && t[1].title === 'Highlights' && t[2].title === 'Roadmap', JSON.stringify(t.map((x) => x.title)));
check('bullets kept', t[1].body.join('|').includes('Revenue up 12%') && t[1].body.join('|').includes('New markets: IN, BR'), JSON.stringify(t[1].body));
check('notes kept', backSlides[0].notes === 'welcome everyone', JSON.stringify(backSlides[0].notes));
check('ellipse kept', backSlides[1].shapes.some((sh) => sh.geom === 'ellipse' && (sh.fill ?? '').toUpperCase() === '#FF7043'), JSON.stringify(backSlides[1].shapes.map((s) => [s.kind, s.geom, s.fill])));
check('background kept', typeof backSlides[0].bg === 'string' && backSlides[0].bg.toUpperCase() === (slides[0].bg ?? '').toUpperCase(), `${backSlides[0].bg} vs ${slides[0].bg}`);

console.log(fails ? `\n${fails} FAILED` : '\nALL PASS');
process.exit(fails ? 1 : 0);
