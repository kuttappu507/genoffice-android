// Generates sample .docx/.xlsx/.pptx files and round-trip checks the parsers.
import * as fs from 'node:fs';
import path from 'node:path';

// usage: node scripts/make-samples.mjs [outDir]   (default: ./samples, git-ignored)
const OUT = path.resolve(process.argv[2] ?? 'samples');
fs.mkdirSync(OUT, { recursive: true });

// ---- sample .docx ----
const { Document, Packer, Paragraph, TextRun, HeadingLevel } = await import('docx');
const doc = new Document({
  sections: [
    {
      children: [
        new Paragraph({ text: 'Quarterly Report', heading: HeadingLevel.HEADING_1 }),
        new Paragraph({
          children: [new TextRun('Revenue grew '), new TextRun({ text: '18%', bold: true }), new TextRun(' this quarter.')],
        }),
        new Paragraph({ text: 'Highlights', heading: NewHeading() }),
        new Paragraph({ text: 'Launched mobile app', bullet: { level: 0 } }),
        new Paragraph({ text: 'Cut inference costs by 40 percent', bullet: { level: 0 } }),
      ],
    },
  ],
});
function NewHeading() {
  return HeadingLevel.HEADING_2;
}
const docxBuf = await Packer.toBuffer(doc);
fs.writeFileSync(path.join(OUT, 'sample.docx'), docxBuf);

// ---- sample .xlsx (with formulas) ----
const XLSX = await import('xlsx');
const ws = {};
const rows = [
  ['Item', 'Qty', 'Price'],
  ['Keyboard', 2, 40],
  ['Monitor', 1, 180],
  ['Cable', 3, 8],
];
rows.forEach((r, ri) => {
  r.forEach((v, ci) => {
    const addr = XLSX.utils.encode_cell({ r: ri, c: ci });
    ws[addr] = typeof v === 'number' ? { t: 'n', v } : { t: 's', v };
  });
});
ws['A6'] = { t: 's', v: 'Total qty' };
ws['B6'] = { t: 'n', f: 'SUM(B2:B4)', v: 5 };
ws['A7'] = { t: 's', v: 'Total cost' };
ws['B7'] = { t: 'n', f: 'SUMPRODUCT(B2:B4,C2:C4)', v: 284 };
ws['!ref'] = 'A1:C7';
const wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, ws, 'Order');
XLSX.writeFile(wb, path.join(OUT, 'sample.xlsx'));

// ---- sample .pptx ----
const PptxGenJS = (await import('pptxgenjs')).default;
const pptx = new PptxGenJS();
const s1 = pptx.addSlide();
s1.addText('GenOffice Overview', { x: 0.6, y: 0.5, w: 8.5, h: 1, fontSize: 30, bold: true });
s1.addText(
  [
    { text: 'Full office suite on your phone', options: { bullet: true } },
    { text: 'BYOK AI built in', options: { bullet: true } },
  ],
  { x: 0.8, y: 1.8, w: 8, h: 3 },
);
const s2 = pptx.addSlide();
s2.addText('Next Steps', { x: 0.6, y: 0.5, w: 8.5, h: 1, fontSize: 30, bold: true });
s2.addText([{ text: 'Install the APK', options: { bullet: true } }], { x: 0.8, y: 1.8, w: 8, h: 2 });
await pptx.writeFile({ fileName: path.join(OUT, 'sample.pptx') });

// ---- round-trip verification ----
const mammoth = await import('mammoth');
const html = (await mammoth.convertToHtml({ path: path.join(OUT, 'sample.docx') })).value;
console.log('DOCX round-trip:', html.includes('Revenue grew') && html.includes('18%') ? 'PASS' : 'FAIL', '|', html.slice(0, 80));

const wb2 = XLSX.read(fs.readFileSync(path.join(OUT, 'sample.xlsx')), { type: 'buffer', cellFormula: true });
const ws2 = wb2.Sheets[wb2.SheetNames[0]];
console.log('XLSX formula B6:', ws2['B6'].f, '| B7:', ws2['B7'].f);

const JSZip = (await import('jszip')).default;
const zip = await JSZip.loadAsync(fs.readFileSync(path.join(OUT, 'sample.pptx')));
const slideCount = Object.keys(zip.files).filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n)).length;
const slide1 = await zip.file('ppt/slides/slide1.xml').async('text');
console.log('PPTX slides:', slideCount, '| title found:', slide1.includes('GenOffice Overview') ? 'PASS' : 'FAIL');

console.log('sizes:', fs.readdirSync(OUT).map((f) => `${f}=${fs.statSync(path.join(OUT, f)).size}B`).join(' '));
