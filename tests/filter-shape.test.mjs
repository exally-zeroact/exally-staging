/* filter-shape.test.mjs — ★実物 19,323本を 100% にした 3つの直し★ 2026-08-29
 *
 *  ★なぜ在るか（実測の履歴）★
 *    司さんの実物（代行計算表2026.xlsb）を 本番の道で 突き合わせた数字:
 *      84.6%  … 空セル参照 2,918本 が 食い違い（別の試験 empty-ref-zero で 直した）
 *      99.73% … .xlsb の 表参照 52本 が #ERROR!
 *      99.95% … #ERROR! は 消えたが 10本が ★金額 → 0★
 *      ★100%★ … FILTER の形を 直して 全部 一致
 *
 *  ★ここで 見る 3つ★
 *    ①FILTER は ★形を保つ★（縦の条件＝行を選ぶ／横の条件＝列を選ぶ）
 *      … 真値は ★実Excel 16.0 を COM で 動かして 測った★（下の表の「実Excel」列）
 *    ②`_xlws.` など Excelの内部の印を 外す（.xlsb は ★裸の _xlws.★ を 出してくる）
 *    ③範囲を「:」で つないだ物を ★囲む四角★に まとめる（HyperFormula は 鎖を 読めない）
 *
 *  走らせ方: node tests/filter-shape.test.mjs [--self-test]
 */
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createRequire } from 'node:module';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(pathToFileURL(path.join(ROOT, 'package.json')));
const 壊す = process.argv.includes('--self-test');
let 緑 = 0, 赤 = 0;
const ok = (名, 条件, 添え) => {
  if (条件) { 緑++; console.log('  ok   ' + 名); }
  else { 赤++; console.log('  ★NG★ ' + 名 + (添え !== undefined ? '  … ' + 添え : '')); }
};

const HFns = require_(path.join(ROOT, 'hyperformula.full.min.js'));
const EF = require_(path.join(ROOT, 'exally-formula.js'));
ok('★プラグインを 積めた（積み忘れると 素のエンジンが 答えて 嘘の緑になる）', EF.registerExallyFunctions(HFns) === true);

/* ══ ① FILTER の形 ═══════════════════════════════════════════
 *  ★真値は 実Excel 16.0 を COM で 動かして 測った物★（2026-08-29）
 *    A1:B4 = [[100,1],[200,2],[300,3],[400,4]] */
console.log('\n[① FILTER は 形を保つ（真値＝実Excel 16.0 実測）]');
const hf = HFns.HyperFormula.buildEmpty({ licenseKey: 'gpl-v3', useArrayArithmetic: true, smartRounding: false });
hf.addSheet('S'); hf.addSheet('D');
const s = hf.getSheetId('S'), d = hf.getSheetId('D');
hf.setSheetContent(d, [[100, 1], [200, 2], [300, 3], [400, 4]]);
const 表 = [
  ['縦の条件＝★行★を選ぶ', '=SUM(FILTER(D!A1:B4,{1;0;0;1}))', 505],
  ['その時の 列の数', '=COLUMNS(FILTER(D!A1:B4,{1;0;0;1}))', 2],
  ['その時の 行の数', '=ROWS(FILTER(D!A1:B4,{1;0;0;1}))', 2],
  ['横の条件＝★列★を選ぶ', '=SUM(FILTER(D!A1:B4,{1,0}))', 1000],
  ['その時の 列の数', '=COLUMNS(FILTER(D!A1:B4,{1,0}))', 1],
  ['★0以外の数は 真★', '=SUM(FILTER(D!A1:A4,{2;0;0;0}))', 100],
  ['★負の数も 真★', '=SUM(FILTER(D!A1:A4,{-1;0;0;0}))', 100],
  ['1列は 前と同じ', '=SUM(FILTER(D!A1:A4,{1;0;0;1}))', 500],
];
hf.setSheetContent(s, 表.map((x) => [x[1]]));
表.forEach((x, i) => {
  let v = hf.getCellValue({ sheet: s, row: i, col: 0 });
  if (v && v.value !== undefined) v = v.value;
  ok(x[0] + '  ' + x[1], v === x[2], '出た=' + v + ' ／ 実Excel=' + x[2]);
});
/* 空の時の 3つ目（★下に 何も無い所で 測る★＝隣に式が在ると #SPILL! に なる） */
{
  const hf2 = HFns.HyperFormula.buildEmpty({ licenseKey: 'gpl-v3', useArrayArithmetic: true, smartRounding: false });
  hf2.addSheet('S'); hf2.addSheet('D');
  hf2.setSheetContent(hf2.getSheetId('D'), [[100], [200], [300], [400]]);
  hf2.setSheetContent(hf2.getSheetId('S'), [['=FILTER(D!A1:A4,{0;0;0;0},"なし")']]);
  let v = hf2.getCellValue({ sheet: hf2.getSheetId('S'), row: 0, col: 0 });
  if (v && v.value !== undefined) v = v.value;
  ok('空の時 3つ目を返す（実Excel＝なし）', v === 'なし', '出た=' + v);
  hf2.destroy();
}
/* ★2列以上を 平らにしていないか★＝直す前の 姿を 名指しで 弾く */
{
  const v = hf.getCellValue({ sheet: s, row: 0, col: 0 });
  ok('★直す前の 姿（102）に なっていない★', (v && v.value !== undefined ? v.value : v) !== 102);
}

/* ══ ①-b SORT / UNIQUE も 形を保つ ═══════════════════════════
 *  ★FILTER と 同じ作りだったので 同じ壊れ方を していた★
 *  （実測 2026-08-29＝10項目中 6件が 実Excelと 違った）。
 *  真値は ★実Excel 16.0 を COM で 動かして 測った物★。
 *  ★TRUE / FALSE は 本番の道（convertFormula）で TRUE() に なる★ので、
 *  ここも ★convertFormula を 通してから★ エンジンに 渡す（素で渡すと #NAME? に なる）。 */
console.log('\n[①-b SORT / UNIQUE も 形を保つ（真値＝実Excel 16.0 実測）]');
{
  const hf4 = HFns.HyperFormula.buildEmpty({ licenseKey: 'gpl-v3', useArrayArithmetic: true, smartRounding: false });
  hf4.addSheet('S'); hf4.addSheet('D');
  const s4 = hf4.getSheetId('S');
  hf4.setSheetContent(hf4.getSheetId('D'), [[300, 1], [100, 2], [200, 3], [100, 2]]);
  const 表2 = [
    ['SORT 2列 の 列の数', '=COLUMNS(SORT(D!A1:B4))', 2],
    ['SORT 2列 の 行の数', '=ROWS(SORT(D!A1:B4))', 4],
    ['SORT 2列 の 左上', '=INDEX(SORT(D!A1:B4),1,1)', 100],
    ['SORT 2列 の 1行2列', '=INDEX(SORT(D!A1:B4),1,2)', 2],
    ['★2列目で 並べ替え★', '=INDEX(SORT(D!A1:B4,2),1,1)', 300],
    ['★2列目で 降順★', '=INDEX(SORT(D!A1:B4,2,-1),1,1)', 200],
    ['1列目 降順', '=INDEX(SORT(D!A1:B4,1,-1),1,1)', 300],
    ['★by_col＝列ごと 並べ替え★', '=INDEX(SORT(D!A1:B4,1,1,TRUE),1,1)', 1],
    ['UNIQUE 2列 の 行の数', '=ROWS(UNIQUE(D!A1:B4))', 3],
    ['UNIQUE 2列 の 列の数', '=COLUMNS(UNIQUE(D!A1:B4))', 2],
    ['★UNIQUE by_col★', '=COLUMNS(UNIQUE(D!A1:B4,TRUE))', 2],
    ['★UNIQUE 1度だけ★', '=ROWS(UNIQUE(D!A1:B4,FALSE,TRUE))', 2],
    ['SORT 1列（前から在る形）', '=SUM(SORT(D!A1:A4))', 700],
    ['UNIQUE 1列（前から在る形）', '=SUM(UNIQUE(D!A1:A4))', 600],
  ];
  hf4.setSheetContent(s4, 表2.map((x) => [EF.convertFormula(x[1])]));
  表2.forEach((x, i) => {
    let v = hf4.getCellValue({ sheet: s4, row: i, col: 0 });
    if (v && v.value !== undefined) v = v.value;
    ok(x[0] + '  ' + x[1], v === x[2], '出た=' + v + ' ／ 実Excel=' + x[2]);
  });
  hf4.destroy();
}

/* ══ ② Excelの内部の印を 外す ══════════════════════════════ */
console.log('\n[② _xlfn. / _xlws. / _xludf. / _xlpm. を 外す]');
const 印 = [
  ['裸の _xlws.（★.xlsb が これを 出す★）', '=_xlws.FILTER(A1:A2,B1:B2)', '=FILTER(A1:A2,B1:B2)'],
  ['_xlfn._xlws. の 二段', '=_xlfn._xlws.SORT(A1:A2)', '=SORT(A1:A2)'],
  ['_xlfn. だけ', '=_xlfn.XLOOKUP(A1,B:B,C:C)', '=XLOOKUP(A1,B:B,C:C)'],
  ['_xludf.', '=_xludf.MYFUNC(A1)', '=MYFUNC(A1)'],
  ['印が 無ければ そのまま', '=SUM(A1:A2)', '=SUM(A1:A2)'],
];
for (const [名, 入, 期待] of 印) ok(名 + '  ' + 入, EF.convertFormula(入) === 期待, EF.convertFormula(入));
/* ★LET は 印を外した後 さらに 展開される★ので convertFormula の 出口では 比べられない。
   ⇒ ★印を外す部品だけ★ を 直に 見る（2026-08-29＝私が 期待値を 間違えて 赤を出した所）。
   `=_xlfn.LET(_xlpm.x,2,_xlpm.x*3)` は 印を外すと `=LET(x,2,x*3)`、
   convertFormula を 通ると 展開されて `=(2)*3`（＝6）に なる。どちらも 正しい。 */
ok('_xlpm.（LET の引数名）を 外す＝印の部品だけ 見る',
  EF._stripXlPrefix('=_xlfn.LET(_xlpm.x,2,_xlpm.x*3)') === '=LET(x,2,x*3)',
  EF._stripXlPrefix('=_xlfn.LET(_xlpm.x,2,_xlpm.x*3)'));
/* ★もう1つの持ち主（lib/xlsx-io.js の stripXlfn）と 同じ物を 外すか★ */
const IO = require_(path.join(ROOT, 'lib/xlsx-io.js'));
if (IO && typeof IO.stripXlfn === 'function') {
  const 種 = '=_xlfn._xlws.SORT(_xlpm.x)+_xlws.FILTER(A1)+_xludf.F(1)+_xlfn.G(2)';
  ok('★2つの持ち主が 同じ物を 外す★（片方だけ直す 事故を 止める）',
    IO.stripXlfn(種) === EF._stripXlPrefix(種), IO.stripXlfn(種) + ' / ' + EF._stripXlPrefix(種));
} else {
  ok('★lib/xlsx-io.js が stripXlfn を 出している（突き合わせが 空振りしていない）', false);
}

/* ══ ③ 範囲の鎖を 囲む四角に ═══════════════════════════════ */
console.log('\n[③ 「:」で つないだ範囲を 囲む四角に まとめる]');
const 鎖 = [
  ['ふつうの範囲は 触らない', '=SUM(A1:C3)', '=SUM(A1:C3)'],
  ['2つを まとめる', '=SUM(A1:A3:C1:C3)', '=SUM(A1:C3)'],
  ['シート名つき', '=SUM(歩合!V2:V32:歩合!AJ2:AJ32)', '=SUM(歩合!V2:AJ32)'],
  ['引用符つき', "=SUM('歩合'!A1:A3:'歩合'!C1:C3)", "=SUM('歩合'!A1:C3)"],
  ['★司さんの実物の形（8本）★',
    '=F(歩合!V2:V32:歩合!X2:X32:歩合!Z2:Z32:歩合!AB2:AB32:歩合!AD2:AD32:歩合!AF2:AF32:歩合!AH2:AH32:歩合!AJ2:AJ32)',
    '=F(歩合!V2:AJ32)'],
  ['★別シートは まとめない★', '=SUM(A!A1:A3:B!C1:C3)', '=SUM(A!A1:A3:B!C1:C3)'],
  ['ただの2セル参照は 触らない', '=A1:B2', '=A1:B2'],
];
for (const [名, 入, 期待] of 鎖) ok(名, EF.convertFormula(入) === 期待, EF.convertFormula(入));

/* ══ ④ 3つを つないで 通す（本番と 同じ順）══════════════════ */
console.log('\n[④ 3つ まとめて＝本番と 同じ形の式が 通る]');
{
  const hf3 = HFns.HyperFormula.buildEmpty({ licenseKey: 'gpl-v3', useArrayArithmetic: true, smartRounding: false });
  hf3.addSheet('S'); hf3.addSheet('歩合');
  const b3 = hf3.getSheetId('歩合');
  /* A列＝日付(シリアル) ／ B,D＝金額 ／ C,E＝空（実物の「時間」列と同じ形） */
  hf3.setSheetContent(b3, [
    [46027, 10, null, 1, null],
    [46037, 20, null, 2, null],
    [46047, 30, null, 3, null],
  ]);
  const 生 = "=SUM(_xlws.FILTER('歩合'!B1:B3:'歩合'!D1:D3,(DAY('歩合'!A1:A3)>=1)*(DAY('歩合'!A1:A3)<=10)))";
  const 直 = EF.convertFormula(生);
  ok('印が 外れ 鎖が まとまった', 直 === "=SUM(FILTER('歩合'!B1:D3,(DAY('歩合'!A1:A3)>=1)*(DAY('歩合'!A1:A3)<=10)))", 直);
  hf3.setSheetContent(hf3.getSheetId('S'), [[直]]);
  let v = hf3.getCellValue({ sheet: hf3.getSheetId('S'), row: 0, col: 0 });
  if (v && v.value !== undefined) v = v.value;
  /* 46027=1月5日 / 46037=1月15日 / 46047=1月25日 ⇒ 1〜10日は 1行目だけ ⇒ 10+1=11 */
  ok('★答えが 出る（#ERROR! でも 0 でもない）★', v === 11, '出た=' + v);
  hf3.destroy();
}

console.log('\nfilter-shape: ' + 緑 + '/' + (緑 + 赤) + ' passed');

if (壊す) {
  console.log('\n★--self-test＝わざと 壊した物を 食わせて 赤に なるか★');
  let 素通り = 0;
  /* 壊し① 直す前の FILTER（平らにして 縦1列） */
  const 壊れFilter = (A, C) => {
    const out = [];
    const flatA = [].concat.apply([], A), flatC = [].concat.apply([], C);
    for (let i = 0; i < flatA.length; i++) if (flatC[i] === 1) out.push(flatA[i]);
    return out.reduce((a, b) => a + b, 0);
  };
  if (壊れFilter([[100, 1], [200, 2], [300, 3], [400, 4]], [[1], [0], [0], [1]]) === 505) {
    素通り++; console.log('  ★素通り★ 直す前の FILTER が 505 を 出してしまった');
  }
  /* 壊し② 印を 短い方から 外す（_xlfn. だけ 消えて _xlws. が 残る） */
  const 壊れ印 = (f) => String(f).replace(/_xlfn\./g, '').replace(/_xlfn\._xlws\./g, '');
  if (壊れ印('=_xlfn._xlws.SORT(A1)') === '=SORT(A1)') { 素通り++; console.log('  ★素通り★ 短い方から 外して 通ってしまった'); }
  /* 壊し③ 鎖を 端どうしで つなぐ（最小・最大を 取らない） */
  const 壊れ鎖 = (f) => f.replace(/(\w+!)?([A-Z]+\d+):[^,)]*?:(\w+!)?([A-Z]+\d+)/, '$1$2:$4');
  if (壊れ鎖('=F(歩合!V2:V32:歩合!X2:X32)') === '=F(歩合!V2:X32)') {
    /* これは たまたま 合う形。合わない形で 見る */
    if (壊れ鎖('=F(歩合!AJ2:AJ32:歩合!V2:V32)') === '=F(歩合!V2:AJ32)') { 素通り++; console.log('  ★素通り★ 逆順の鎖を 正しく まとめてしまった'); }
  }
  if (素通り) { hf.destroy(); console.log('★抜け道 ' + 素通り + '件★'); process.exit(1); }
  console.log('  ok   抜け道 0件');
}
hf.destroy();
process.exit(赤 ? 1 : 0);
