// roundtrip.test.mjs — 数式入り xlsx の往復。
//   ① 式が保たれるか  ② 計算済みの値(キャッシュ値)が保たれるか
//   ③ 書き出した物を実Excelで開いて再計算した値が golden と一致するか
//      → ③はWindows+Excelが要るので tools/roundtrip-excel.ps1 (CIでは走らない)。
//        ここでは ③用のファイル tmp/roundtrip.xlsx を書き出すところまでやる。
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createRequire } from 'node:module';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(__dirname, '..', '..');
const require = createRequire(pathToFileURL(path.join(ROOT, 'package.json')));
const XlsxIO = require(path.join(__dirname, '..', '..', 'lib', 'xlsx-io.js'));

let pass = 0, fail = 0;
const T = (name, fn) => { try { fn(); pass++; console.log('  ok   ' + name); } catch (e) { fail++; console.log('  NG   ' + name + '\n       ' + e.message); } };
const ok = (c, m) => { if (!c) throw new Error(m || 'expected truthy'); };
const eq = (a, b, m) => { if (a !== b) throw new Error((m || '') + ` 期待=${JSON.stringify(b)} 実際=${JSON.stringify(a)}`); };

console.log('=== xlsx 往復(SheetJS ' + XlsxIO.sheetjsVersion + ') ===');

/* 中小の実務で出る形をひと通り: 数値・文字列・日付シリアル・各種数式 */
const src = {
  sheets: [{
    name: '請求',
    cells: {
      A1: { v: '品名', t: 's' }, B1: { v: '数量', t: 's' }, C1: { v: '単価', t: 's' }, D1: { v: '金額', t: 's' },
      A2: { v: '作業A', t: 's' }, B2: { v: 2, t: 'n' }, C2: { v: 15000, t: 'n' }, D2: { f: '=B2*C2', v: 30000, t: 'n' },
      A3: { v: '作業B', t: 's' }, B3: { v: 3, t: 'n' }, C3: { v: 8000, t: 'n' }, D3: { f: '=B3*C3', v: 24000, t: 'n' },
      D4: { f: '=SUM(D2:D3)', v: 54000, t: 'n' },
      D5: { f: '=ROUNDDOWN(D4*0.1,0)', v: 5400, t: 'n' },
      D6: { f: '=D4+D5', v: 59400, t: 'n' },
      A7: { v: '締日', t: 's' }, B7: { v: 46234, t: 'n' },
      A8: { f: '=TEXT(B7,"yyyy/mm/dd")', v: '2026/07/31', t: 's' },
      A9: { f: '=IF(D6>50000,"要確認","")', v: '要確認', t: 's' },
      A10: { f: '=COUNTIF(B2:B3,">=3")', v: 1, t: 'n' },
      A11: { f: '=VLOOKUP("作業B",A2:D3,4,FALSE)', v: 24000, t: 'n' },
      // ★新しい関数(_xlfn. を付けないと Excel がファイルごと開けなくなる物)
      A12: { f: '=XLOOKUP("作業B",A2:A3,D2:D3,0)', v: 24000, t: 'n' },
      A13: { f: '=IFS(D6>50000,"大",TRUE,"小")', v: '大', t: 's' },
      A14: { f: '=TEXTJOIN(",",TRUE,A2:A3)', v: '作業A,作業B', t: 's' },
      A15: { f: '=COUNTA(UNIQUE(A2:A3))', v: 2, t: 'n' },
      A16: { f: '=SUM(FILTER(D2:D3,B2:B3>=3))', v: 24000, t: 'n' },
      A17: { f: '=INDEX(SORT(D2:D3,1,-1),1)', v: 30000, t: 'n' }
    }
  }]
};

const buf = XlsxIO.writeBook(src);
T('書き出せる(バイト列が返る)', () => ok(buf && buf.length > 1000, 'size=' + (buf && buf.length)));

const back = XlsxIO.readBook(buf);
T('シートが1枚・名前が保たれる', () => { eq(back.sheets.length, 1); eq(back.sheets[0].name, '請求'); });

const A = back.sheets[0].cells;
const srcCells = src.sheets[0].cells;
const formulaAddrs = Object.keys(srcCells).filter(a => srcCells[a].f);

T(`①式が保たれる(${formulaAddrs.length}本)`, () => {
  for (const a of formulaAddrs) {
    ok(A[a], `${a} が読み戻せていない`);
    eq(A[a].f, srcCells[a].f, `${a} の式:`);
  }
});

T(`②計算済みの値が保たれる(${formulaAddrs.length}本)`, () => {
  for (const a of formulaAddrs) eq(A[a].v, srcCells[a].v, `${a} の値:`);
});

T('値だけのセルも型ごと保たれる', () => {
  for (const a of Object.keys(srcCells)) {
    if (srcCells[a].f) continue;
    eq(A[a].v, srcCells[a].v, `${a} の値:`);
    eq(A[a].t, srcCells[a].t, `${a} の型:`);
  }
});

T('日本語のシート名・文字列が壊れない', () => { eq(A.A2.v, '作業A'); eq(A.A9.v, '要確認'); });

T('2回往復しても変わらない(冪等)', () => {
  const twice = XlsxIO.readBook(XlsxIO.writeBook(back));
  const B = twice.sheets[0].cells;
  for (const a of formulaAddrs) { eq(B[a].f, srcCells[a].f, `${a} の式(2回目):`); eq(B[a].v, srcCells[a].v, `${a} の値(2回目):`); }
});

/* ★新しい関数の接頭辞(_xlfn.)。これが無いと Excel はファイルごと開けない(実測)。
 *  SheetJSは読み戻す時に _xlfn. を外して正規化するので、読み戻した式では確かめられない。
 *  ファイル(zip)の中の sheet1.xml を直接見る。 */
import zlib from 'node:zlib';
function readZipEntry(raw, wantName) {
  const zip = Buffer.isBuffer(raw) ? raw : Buffer.from(raw);   // SheetJSは Uint8Array を返す事がある
  for (let i = 0; i + 30 < zip.length; i++) {
    if (zip.readUInt32LE(i) !== 0x04034b50) continue;          // local file header
    const method = zip.readUInt16LE(i + 8);
    const compSize = zip.readUInt32LE(i + 18);
    const nameLen = zip.readUInt16LE(i + 26);
    const extraLen = zip.readUInt16LE(i + 28);
    const name = zip.slice(i + 30, i + 30 + nameLen).toString('utf8');
    const start = i + 30 + nameLen + extraLen;
    if (name !== wantName) continue;
    const data = zip.slice(start, start + compSize);
    return (method === 0 ? data : zlib.inflateRawSync(data)).toString('utf8');
  }
  return null;
}
const sheetXml = readZipEntry(buf, 'xl/worksheets/sheet1.xml');

T('書き出したxlsxの中身(sheet1.xml)で新関数に _xlfn. が付いている', () => {
  ok(sheetXml, 'sheet1.xml を取り出せない');
  for (const want of ['_xlfn.XLOOKUP(', '_xlfn.IFS(', '_xlfn.TEXTJOIN(', '_xlfn.UNIQUE(', '_xlfn.FILTER(', '_xlfn.SORT(']) {
    ok(sheetXml.includes(want), want + ' がファイルの中に無い');
  }
});
T('古い関数には付かない', () => {
  ok(sheetXml.includes('SUM(D2:D3)'), 'SUM が書き換わっている');
  ok(!sheetXml.includes('_xlfn.SUM'), 'SUM に余計な接頭辞が付いた');
  ok(!sheetXml.includes('_xlfn.VLOOKUP'), 'VLOOKUP に余計な接頭辞が付いた');
});
T('読み戻すと元の名前に戻る(グリッドは _xlfn. を知らなくてよい)', () => {
  eq(A.A12.f, '=XLOOKUP("作業B",A2:A3,D2:D3,0)');
  eq(A.A17.f, '=INDEX(SORT(D2:D3,1,-1),1)');
});
T('文字列の中の関数名は書き換えない', () => {
  const s = XlsxIO.applyXlfn('CONCATENATE("SORT(",A1,")")');
  eq(s, 'CONCATENATE("SORT(",A1,")")');
});
T('すでに付いている物を二重に付けない', () => {
  eq(XlsxIO.applyXlfn('_xlfn.XLOOKUP(1,A1:A2,B1:B2)'), '_xlfn.XLOOKUP(1,A1:A2,B1:B2)');
});
T('関数呼び出しでない同名の語は触らない', () => {
  eq(XlsxIO.applyXlfn('A1&"-"&SORT'), 'A1&"-"&SORT');
});
T('LET/LAMBDA は黙って壊れたファイルを作らず、書き出しを止める', () => {
  let threw = '';
  try { XlsxIO.writeBook({ sheets: [{ name: 'S', cells: { A1: { f: '=LET(x,2,x*3)' } } }] }); }
  catch (e) { threw = e.message; }
  ok(threw.includes('LET'), '例外が出ていない: ' + threw);
});

/* ③用: 実Excelで開いて再計算するためのファイルを置く */
const tmpDir = path.join(__dirname, 'tmp');
fs.mkdirSync(tmpDir, { recursive: true });
const outPath = path.join(tmpDir, 'roundtrip.xlsx');
fs.writeFileSync(outPath, buf);
fs.writeFileSync(path.join(tmpDir, 'roundtrip-expected.json'), JSON.stringify(
  Object.fromEntries(formulaAddrs.map(a => [a, { f: srcCells[a].f, v: srcCells[a].v }])), null, 1));
console.log(`  (③実Excel確認用: tests/xlsx-harness/tmp/roundtrip.xlsx を書き出した。pwsh -File tools/roundtrip-excel.ps1 で確認)`);

console.log(`\nxlsx往復: ${pass}/${pass + fail} passed`);
process.exit(fail ? 1 : 0);
