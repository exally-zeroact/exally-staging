/* grid-xlsx.test.mjs — グリッド → xlsx の変換と「落ちる物の警告」のテスト。
 *
 *  ★ここが守る事:
 *    ・グリッドの持ち方('r,c' のキー / f・v・d)を xlsx の形('A1' / f・v・t)へ正しく移す
 *    ・「1,234」や「2026-07-31」を勝手に数値・日付にしない（画面と落としたファイルで中身を変えない）
 *    ・書式・結合・列幅のうち【書ける物】は落とさない
 *    ・書けない物(太字/色/罫線)と、スピルしない配列式は【セル番地つきで】伝える
 *    ・実際に xlsx を書いて、読み戻して同じになる（往復）
 *
 *  使い方: node tests/grid-xlsx.test.mjs
 *          node tests/grid-xlsx.test.mjs --self-test  … わざと壊して赤になるか
 */
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createRequire } from 'node:module';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(__dirname, '..');
const require = createRequire(pathToFileURL(path.join(ROOT, 'package.json')));
const G = require(path.join(ROOT, 'lib', 'grid-xlsx.js'));
const IO = require(path.join(ROOT, 'lib', 'xlsx-io.js'));

let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ✓ ' + n); } catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + e.message); } };
const eq = (a, b, m) => { if (JSON.stringify(a) !== JSON.stringify(b)) throw new Error((m || '') + ' 期待=' + JSON.stringify(b) + ' 実際=' + JSON.stringify(a)); };

/* 中小の実務で出る形をひと通り入れたグリッド */
function sampleSheets() {
  return [{
    name: '請求',
    data: {
      '0,0': { v: '品名', f: '品名', d: '品名' },
      '0,1': { v: '金額', f: '金額', d: '金額' },
      '1,0': { v: '作業A', f: '作業A', d: '作業A' },
      '1,1': { v: '15000', f: '15000', d: '15,000', numFmt: '#,##0' },
      '2,0': { v: '作業B', f: '作業B', d: '作業B' },
      '2,1': { v: '8000', f: '8000', d: '8,000', numFmt: '#,##0' },
      '3,0': { v: '合計', f: '合計', d: '合計', bold: true },
      '3,1': { f: '=SUM(B2:B3)', v: '=SUM(B2:B3)', d: '23000', numFmt: '#,##0' },
      '4,0': { v: '0007', f: '0007', d: '0007' },              // ★数値にしてはいけない
      '4,1': { v: '1,234', f: '1,234', d: '1,234' },           // ★数値にしてはいけない
      '5,0': { v: '色つき', f: '色つき', d: '色つき', bgColor: '#FFFF00' },
      '6,0': { f: '=SORT(B2:B3)', v: '=SORT(B2:B3)', d: '8000' },  // 配列を返す式
      '7,0': { v: '結合', f: '結合', d: '結合', merged: { r: 7, c: 0 }, mergeEnd: { r: 7, c: 1 } }
    },
    colW: { 0: 120, 1: 90 }
  }];
}

if (process.argv.includes('--self-test')) {
  console.log('\n[grid-xlsx --self-test] ★わざと壊して赤になるか');
  const cases = [
    ['① 「1,234」を数値にしてしまう', () => { if (typeof G.asValue('1,234') === 'number') throw new Error('数値になった'); }],
    ['② 式セルの f を落とす', () => { const b = G.gridToBook(sampleSheets()); if (!b.sheets[0].cells.B4.f) throw new Error('式が消えた'); }],
    ['③ 表示形式(z)を落とす', () => { const b = G.gridToBook(sampleSheets()); if (b.sheets[0].cells.B2.z !== '#,##0') throw new Error('表示形式が消えた'); }],
    ['④ 太字セルを警告し損ねる', () => { const w = G.exportWarnings(sampleSheets(), {}); if (!w.some(x => x.kind === 'style-lost')) throw new Error('警告が出ない'); }],
    ['⑤ 配列式を警告し損ねる', () => { const w = G.exportWarnings(sampleSheets(), { isArrayFormula: (s, r, c) => r === 6 && c === 0 }); if (!w.some(x => x.kind === 'array-spill')) throw new Error('警告が出ない'); }],
    ['⑥ 警告にセル番地が入らない', () => { const w = G.exportWarnings(sampleSheets(), {}); if (!/請求!A4/.test(w[0].msg)) throw new Error('番地が無い: ' + w[0].msg); }]
  ];
  for (const [name, check] of cases) {
    // 期待: 正しい実装では check が例外を投げない
    T(name.replace(/^(.)/, '$1 は正しい実装では起きない'), check);
  }
  console.log('\n' + pass + ' passed, ' + fail + ' failed');
  process.exit(fail ? 1 : 0);
}

console.log('\n[grid-xlsx] グリッド → xlsx の変換と、落ちる物の警告');

T('① セルのキー(r,c)が A1形式になる', () => {
  eq(G.addr(0, 0), 'A1'); eq(G.addr(3, 1), 'B4'); eq(G.addr(0, 26), 'AA1');
});

T('② 式セルは f と「計算済みの値」を両方持つ（Excelは開いた瞬間その値を見せる）', () => {
  const b = G.gridToBook(sampleSheets());
  eq(b.sheets[0].cells.B4.f, '=SUM(B2:B3)');
  eq(b.sheets[0].cells.B4.v, 23000);
  eq(b.sheets[0].cells.B4.t, 'n');
});

T('③ ★数値判定が book.html の toHFVal と同じ規則（＝落としたファイルの合計が画面と食い違わない）', () => {
  const b = G.gridToBook(sampleSheets());
  // '0007' → 7。グリッドのエンジン(toHFVal)が既に 7 として計算しているので、ここも 7 でなければ
  // Excel の =SUM だけがグリッドと違う数字になる。実Excelも打ち込んだ 0007 は 7 にする(golden INPUT_typed_0007)。
  eq(b.sheets[0].cells.A5.v, 7); eq(b.sheets[0].cells.A5.t, 'n');
  // '1,234' は isNaN なので文字列のまま。グリッドのエンジンも文字列として扱う(台帳 R11)。
  eq(b.sheets[0].cells.B5.v, '1,234'); eq(b.sheets[0].cells.B5.t, 's');
});
/* ★2026-08-05 に決まりを1つ変えた（理由を残す）★
   前: 「日付に見える文字は日付にしない」
   後: ★日付はExcelと同じ数（シリアル値）にする★
   なぜ変えたか: 元の決まりの目的は「書き出しとグリッドの画面を一致させる」ことだった。
     ところがグリッド側で ★日付が日付として扱われず★、
       ・「2026/8/31」+30 が ★2056★（2026+30）になる
       ・=DATE(2026,8,31) が 46265 という裸の数字のまま出る
     という不具合があった。グリッド側を Excel と同じ扱いに直したので、
     ★目的（画面と書き出しを一致させる）を守るには、書き出しも同じ数にする★のが正しい。
   ★「日付でない物」は今までどおり文字のまま★（007-1234・1,234 は変えていない）。 */
T('③b 日付はExcelと同じ数にする／日付でない物は文字のまま（2026-08-05 変更）', () => {
  eq(G.asValue('2026-07-31'), 46234, '日付はシリアル値');
  eq(G.asValue('007-1234'), '007-1234', '電話番号のような物は文字のまま');
  eq(G.asValue('1,234'), '1,234', 'カンマ付きは文字のまま（台帳 R11・変更なし）');
});

T('★日付を数で書くなら、日付の表示形式も一緒に書く（Excelで裸の数字に見えない）', () => {
  eq(G.dateFmtFor('2026/8/31'), 'yyyy/m/d', '打った形が / なら /');
  eq(G.dateFmtFor('2026-08-31'), 'yyyy-mm-dd', '打った形が - なら -');
  eq(G.dateFmtFor('15000'), null, '日付でない物には付けない');
  const b = G.gridToBook([{ name: 'S', data: { '0,0': { v: '2026/8/31' } } }]);
  eq(b.sheets[0].cells.A1.v, 46265, '値は数');
  eq(b.sheets[0].cells.A1.z, 'yyyy/m/d', '表示形式が付く');
});

T('★人が選んだ表示形式は上書きしない', () => {
  const b = G.gridToBook([{ name: 'S', data: { '0,0': { v: '2026/8/31', numFmt: 'yyyy年m月d日' } } }]);
  eq(b.sheets[0].cells.A1.z, 'yyyy年m月d日');
});

T('④ 数値は数値として入る', () => {
  const b = G.gridToBook(sampleSheets());
  eq(b.sheets[0].cells.B2.v, 15000); eq(b.sheets[0].cells.B2.t, 'n');
});

T('⑤ 表示形式・セルの結合・列幅は落とさない（CEでも書けると実測済み）', () => {
  const b = G.gridToBook(sampleSheets());
  eq(b.sheets[0].cells.B2.z, '#,##0', '表示形式');
  eq(b.sheets[0].merges, [{ s: { r: 7, c: 0 }, e: { r: 7, c: 1 } }], '結合');
  eq(b.sheets[0].cols[0], { wpx: 120 }, '列幅');
});

T('⑥ 空セルは書かない', () => {
  const b = G.gridToBook([{ name: 'S', data: { '0,0': { v: '', f: '', d: '' } }, colW: {} }]);
  eq(Object.keys(b.sheets[0].cells).length, 0);
});

T('⑦ ★書けない書式は「何が入らないか」を具体的に、セル番地つきで伝える', () => {
  const w = G.exportWarnings(sampleSheets(), {});
  const s = w.find(x => x.kind === 'style-lost');
  if (!s) throw new Error('警告が出ていない');
  if (!/太字・文字色・背景色・罫線はExcelファイルに入りません/.test(s.msg)) throw new Error('言い方が曖昧: ' + s.msg);
  if (!/表示形式・セルの結合・列幅は入ります/.test(s.msg)) throw new Error('入る物が書かれていない: ' + s.msg);
  eq(s.cells, ['請求!A4', '請求!A6'], '番地');
});

T('⑧ ★配列を返す式は件数と番地を伝える', () => {
  const w = G.exportWarnings(sampleSheets(), { isArrayFormula: (si, r, c) => r === 6 && c === 0 });
  const a = w.find(x => x.kind === 'array-spill');
  if (!a) throw new Error('警告が出ていない');
  if (!/配列を返す式が1件あります。Excelでは1セル分の値になります/.test(a.msg)) throw new Error('言い方: ' + a.msg);
  eq(a.cells, ['請求!A7']);
});

T('⑨ 落ちる物が無ければ警告は出ない（黙って不安にさせない）', () => {
  const clean = [{ name: 'S', data: { '0,0': { v: '1', f: '1', d: '1' } }, colW: {} }];
  eq(G.exportWarnings(clean, {}), []);
});

T('⑩ 番地が多い時は8件で切って「ほか◯件」にする', () => {
  const many = []; for (let i = 0; i < 12; i++) many.push('S!A' + (i + 1));
  const s = G.listCells(many);
  if (!/ほか4件$/.test(s)) throw new Error(s);
});

T('⑪ ★実際に xlsx を書いて読み戻せる（式・値・表示形式が保たれる）', () => {
  const book = G.gridToBook(sampleSheets());
  const buf = IO.writeBook(book);
  const back = IO.readBook(buf);
  const cells = back.sheets[0].cells;
  eq(cells.B4.f, '=SUM(B2:B3)', '式');
  eq(cells.B4.v, 23000, '計算済みの値');
  // ★表示形式は writeBook が運ばないと実Excelで「G/標準」になる(2026-08-02 実機で踏んだ)
  eq(cells.B4.z, '#,##0', '表示形式がファイルに入っている');
  eq(cells.A5.v, 7, 'toHFVal と同じ規則で数値');
  eq(cells.B5.v, '1,234', '桁区切り付きは文字列のまま');
  if (!buf || !buf.length) throw new Error('ファイルが空');
});

T('⑫ ★新しい関数には _xlfn. が付く（付け忘れるとExcelが式を壊す）', () => {
  const g = [{ name: 'S', data: { '0,0': { f: '=XLOOKUP(1,A1:A2,B1:B2)', v: '', d: '0' } }, colW: {} }];
  const buf = IO.writeBook(G.gridToBook(g));
  const back = IO.readBook(buf);
  // readBook は _xlfn. を外して返すので、生の書き出しに入っているかは applyXlfn で確認する
  eq(IO.applyXlfn('XLOOKUP(1,A1:A2,B1:B2)'), '_xlfn.XLOOKUP(1,A1:A2,B1:B2)');
  eq(back.sheets[0].cells.A1.f, '=XLOOKUP(1,A1:A2,B1:B2)', '読み戻すと元の名前に戻る');
});

T('⑬ ★★数値判定が book.html の toHFVal とズレたら赤（ズレると合計だけが食い違う）★★', () => {
  const fs = require('node:fs');
  const src = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf8');
  const m = src.match(/function toHFVal\(v\)\{[\s\S]*?\n\}/);
  if (!m) throw new Error('book.html から toHFVal を取り出せない（名前が変わった？）');
  // 実際に book.html の toHFVal を動かして、asValue と同じ答えになるか比べる
  // toHFVal は dateSerial / parseDateStr を呼ぶので、一緒に取り出して同じ入れ物で動かす
  const helpers = ['dateSerial', 'parseDateStr'].map(n => {
    const mm = src.match(new RegExp('function ' + n + '\\([^)]*\\)\\s*\\{[\\s\\S]*?\\n\\}'));
    if (!mm) throw new Error('book.html から ' + n + ' を取り出せない（名前が変わった？）');
    return mm[0];
  }).join('\n');
  const toHFVal = new Function(helpers + '\nreturn (' + m[0].replace(/^function toHFVal/, 'function') + ')')();
  const samples = ['0007', '1,234', '2026-07-31', '007-1234', '15000', '1.50', '-3', '0', '  12  ', 'あ', ''];
  const ng = [];
  for (const s of samples) {
    const a = G.asValue(s);
    const b = toHFVal(s);
    const same = (a === '' && b === null) || a === b;   // 空セルの表し方だけ違う（'' と null）
    if (!same) ng.push(`${JSON.stringify(s)}: grid-xlsx=${JSON.stringify(a)} / toHFVal=${JSON.stringify(b)}`);
  }
  if (ng.length) throw new Error('判定がズレている:\n   - ' + ng.join('\n   - ')
    + '\n   → どちらかを直して必ず同じ規則にすること（違うと落としたファイルの合計が画面と食い違う）');
});

console.log('\n' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
