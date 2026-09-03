// Cross-sheet references, structural edits and CSV helpers (run: node scripts/test-sheet-model.mjs)
import { build } from 'esbuild';
import { mkdtempSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';

const dir = mkdtempSync(path.join(tmpdir(), 'gof-'));
const out = path.join(dir, 'sheet-model.mjs');
await build({ entryPoints: ['src/lib/sheet-model.ts'], bundle: true, format: 'esm', outfile: out, platform: 'neutral', logLevel: 'silent' });
const sheet = await import(out);

let fails = 0;
const t = (name, got, exp) => {
  const ok = JSON.stringify(got) === JSON.stringify(exp);
  if (!ok) fails++;
  console.log(`${ok ? 'PASS' : 'FAIL'} ${name.padEnd(46)} -> ${JSON.stringify(got)}${ok ? '' : `   (expected ${JSON.stringify(exp)})`}`);
};

const p1 = sheet.emptyPage();
p1.cells = { B2: '3', C2: '2.5', D2: '=B2*C2', B3: '4', C3: '10', D3: '=B3*C3', D4: '=SUM(D2:D3)', E1: '=Data!A2', E2: '=Data!A6', E3: '=IFERROR(Data!A6,"safe")', E4: "='My Sheet'!A1*10", E5: '=Data!A7' };
const p2 = sheet.emptyPage();
p2.cells = { A1: 'Other', A2: '=Sheet1!D4*2', A3: "='Sheet1'!D2", A4: '=SUM(Sheet1!D2:D3)', A5: '=Nope!A1', A6: '=1/0', A7: '=Sheet1!E5' };
const p3 = sheet.emptyPage();
p3.cells = { A1: '=B1+1', B1: '4' };
const book = { order: ['Sheet1', 'Data', 'My Sheet'], active: 'Sheet1', sheets: { Sheet1: p1, Data: p2, 'My Sheet': p3 } };

const d2 = sheet.computeDisplay(p2.cells, book, 'Data');
t('cross-sheet scalar', d2.A2, '95');
t('cross-sheet quoted name', d2.A3, '7.5');
t('cross-sheet range', d2.A4, '47.5');
t('unknown sheet', d2.A5, '#REF!');
const d1 = sheet.computeDisplay(p1.cells, book, 'Sheet1');
t('nested foreign formula', d1.E1, '95');
t('foreign error propagates', d1.E2, '#DIV/0!');
t('IFERROR over foreign error', d1.E3, 'safe');
t('sheet name with space', d1.E4, '50');
t('cross-sheet cycle', d1.E5, '#CIRC!');

t('qualifyFormula keeps range end + strings', sheet.qualifyFormula('=SUM(A1:B2)+Sheet1!C3*IF(A1>2,1,0)&"A1"', 'My Sheet'), `=SUM('My Sheet'!A1:B2)+Sheet1!C3*IF('My Sheet'!A1>2,1,0)&"A1"`);
t('offsetFormula shifts relative refs (incl. foreign, like Excel)', sheet.offsetFormula('=A1+Data!A1+$B$2', 1, 1), '=B2+Data!B2+$B$2');
t('remapFormula leaves foreign refs', sheet.remapFormula('=A1+Data!A1', (c, r) => [c + 2, r]), '=C1+Data!A1');

t('parseDelimited quotes', sheet.parseDelimited('a,b\n1,"x,y"\n'), [['a', 'b'], ['1', 'x,y']]);
t('parseDelimited tabs', sheet.parseDelimited('a\tb\n1\t2'), [['a', 'b'], ['1', '2']]);
t('toCsv escapes', sheet.toCsv([['a', 'b,c'], ['1', 'say "hi"']]), 'a,"b,c"\n1,"say ""hi"""');

const pi = sheet.emptyPage();
pi.cells = { A1: '1', B1: '=A1*2', C1: '=SUM(A1:B1)' };
const ins = sheet.insertCols(pi, 1, 1);
t('insertCols shifts refs', [ins.cells.C1, ins.cells.D1], ['=A1*2', '=SUM(A1:C1)']);
const del = sheet.deleteCols(pi, 0, 1);
t('deleteCols → #REF!', [del.cells.A1, del.cells.B1], ['=#REF!*2', '=SUM(#REF!:A1)']);

console.log(fails ? `\n${fails} FAILED` : '\nALL PASS');
process.exit(fails ? 1 : 0);
