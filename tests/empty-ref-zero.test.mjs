/* empty-ref-zero.test.mjs — ★式が「空のセル」を指したら Excelは 0★ 2026-08-29
 *
 *  ★なぜ在るか（実測）★
 *    司さんの実物 19,323本を 本番の道で 突き合わせたら、
 *    ★2,918本が「Excelは 0・うちは 空」★で 食い違っていた（Excelとの一致 84.6%）。
 *    ★実Excel 16.0 で 裁定した（新しい空ブック・司さんの実物は 開いていない）★:
 *
 *      =C2（C2は空）        → ★0★（数）  画面「0」
 *      =INDEX(C1:C5,2)（空） → ★0★（数）  画面「0」
 *      =IF(C2="","",1)      → ""（字）    画面 空
 *      =C2&""               → ""（字）    画面 空
 *
 *    ⇒ ★決まり＝「数の場所なら 0・字の場所なら 空」★
 *      HyperFormula は 前者で null・後者で "" を返すので ★区別できる★。
 *
 *  ★この試験で 必ず 見る事★
 *    ①空セル参照 → 0 に なる
 *    ②★"" は 空のまま★（0 に しない）＝ここを 間違えると 表が 0 だらけになる
 *    ③★ただの空セル（式でない）は 空のまま★＝式のセルだけ 0 にする
 *
 *  走らせ方: node tests/empty-ref-zero.test.mjs [--self-test]
 */
import fs from 'node:fs';
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

/* ── 本物の エンジンと 本物の _hfGetDisplay を 使う ─────────── */
const HFns = require_(path.join(ROOT, 'hyperformula.full.min.js'));
const EF = require_(path.join(ROOT, 'exally-formula.js'));
ok('★exally-formula.js が _hfGetDisplay を 出している（試験が 空振りしていない）', typeof EF._hfGetDisplay === 'function');
if (typeof EF.registerExallyFunctions === 'function') EF.registerExallyFunctions(HFns);

const { HyperFormula } = HFns;
const hf = HyperFormula.buildEmpty({ licenseKey: 'gpl-v3', useArrayArithmetic: true, smartRounding: false });
hf.addSheet('S');
const sid = hf.getSheetId('S');
/*  A列＝式 ／ C列＝空のまま
      A1 = =C2            （空セルを 指す）           → Excel 0
      A2 = =INDEX(C1:C5,2)（空セルを 指す）           → Excel 0
      A3 = =IF(C2="","",1)（字の場所）                → Excel ""（空）
      A4 = =C2&""         （字の場所）                → Excel ""（空）
      A5 = =1+1           （ふつうの数）              → 2
      B1 = （式でない 空セル）                        → 空のまま           */
hf.setSheetContent(sid, [
  ['=C2', null, null],
  ['=INDEX(C1:C5,2)', null, null],
  ['=IF(C2="","",1)', null, null],
  ['=C2&""', null, null],
  ['=1+1', null, null],
]);

/* ★本番と 同じ入口で エンジンを 差し込む★（book.html も これを 呼ぶ） */
if (typeof EF.initExallyFormula !== 'function') { console.log('  ★NG★ initExallyFormula が 無い'); process.exit(1); }
EF.initExallyFormula(hf);

/* ★決まり そのものを 別に 書いた物★（⑤で 本物と 突き合わせる相手。
   これだけで 緑に しない＝⑤が 本物を 呼ぶ） */
function 画面の字(r, c, 式のセルか) {
  const v = hf.getCellValue({ sheet: sid, row: r, col: c });
  if (v === null || v === undefined) return 式のセルか ? '0' : '';
  if (typeof v === 'object' && v.type) return '#' + v.type;
  return String(v);
}

console.log('\n[① 空のセルを 指した式 → 0]');
ok('=C2（C2は空） → 0', 画面の字(0, 0, true) === '0', 画面の字(0, 0, true));
ok('=INDEX(C1:C5,2)（空） → 0', 画面の字(1, 0, true) === '0', 画面の字(1, 0, true));

console.log('\n[② ★"" は 空のまま★（0 に しない）]');
ok('=IF(C2="","",1) → 空', 画面の字(2, 0, true) === '', JSON.stringify(画面の字(2, 0, true)));
ok('=C2&"" → 空', 画面の字(3, 0, true) === '', JSON.stringify(画面の字(3, 0, true)));

console.log('\n[③ 式でない 空セルは 空のまま]');
ok('式でない空セル → 空', 画面の字(0, 1, false) === '', JSON.stringify(画面の字(0, 1, false)));

console.log('\n[④ ふつうの数は そのまま]');
ok('=1+1 → 2', 画面の字(4, 0, true) === '2', 画面の字(4, 0, true));

/* ── ⑤ ★本物の _hfGetDisplay そのものを 呼ぶ★（写しで 緑に しない）───
   ★決まり＝検査は ソースを読むな・実際に 走らせろ★ */
console.log('\n[⑤ ★本物の _hfGetDisplay を 呼ぶ★]');
ok('本物: =C2（空を指す） → 0', EF._hfGetDisplay('S', 0, 0, true) === '0', JSON.stringify(EF._hfGetDisplay('S', 0, 0, true)));
ok('本物: =INDEX(C1:C5,2)（空） → 0', EF._hfGetDisplay('S', 1, 0, true) === '0', JSON.stringify(EF._hfGetDisplay('S', 1, 0, true)));
ok('本物: =IF(C2="","",1) → ★空のまま★', EF._hfGetDisplay('S', 2, 0, true) === '', JSON.stringify(EF._hfGetDisplay('S', 2, 0, true)));
ok('本物: =C2&"" → ★空のまま★', EF._hfGetDisplay('S', 3, 0, true) === '', JSON.stringify(EF._hfGetDisplay('S', 3, 0, true)));
ok('本物: 式でない空セル → ★空のまま★', EF._hfGetDisplay('S', 0, 1, false) === '', JSON.stringify(EF._hfGetDisplay('S', 0, 1, false)));
ok('本物: =1+1 → 2', EF._hfGetDisplay('S', 4, 0, true) === '2', JSON.stringify(EF._hfGetDisplay('S', 4, 0, true)));

/* ── ⑥ 呼び手が 渡し忘れていないか（book.html）───────────── */
console.log('\n[⑥ 呼び手が 式かどうかを 渡している]');
const book = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf8');
const 呼び出し = book.match(/_hfGetDisplay\([^)]*\)/g) || [];
/* ★★数を 焼き込まない（2026-09-10 に 踏んだ）★★
   前は `呼び出し.length === 2` と 書いて いました。
   ⇒ 溢れ（1つの 式が 何マスにも 広がる）を 直して ★3か所目が 増えた★だけで
     ★中身は 正しいのに 赤★に なりました。
   ⇒★ここで 守りたいのは「呼び手が ★全員★ 4つ目を 渡して いる」事★
     ＝★数では ありません★。数は ★空振りして いない★事の 確かめだけ。
   ⇒ 同じ型を 今日 3回 見て います（帯の 数／丸めの 一覧の 数／ここ）。 */
ok('★呼び手を 見つけた（空振りしていない）', 呼び出し.length >= 1, 呼び出し.length + 'か所 … ' + 呼び出し.join(' / '));
ok('★どの呼び手も 4つ目を 渡している★', 呼び出し.every((s) => s.split(',').length >= 4), 呼び出し.join(' / '));

console.log('\nempty-ref-zero: ' + 緑 + '/' + (緑 + 赤) + ' passed');

/* ★エンジンを 閉じるのは 一番最後★（自己確認も エンジンを 使う。
   先に 閉じると 例外が 出て、その中身に ライブラリ丸ごとが 乗って 画面が 埋まる
   ＝2026-08-29 実測で 873KB 吐いた） */
if (壊す) {
  console.log('\n★--self-test＝わざと 壊した物を 食わせて 赤に なるか★');
  let 素通り = 0;
  /* 壊し① 式のセルでも '' を返す（＝直す前の 姿） */
  const 壊れ1 = (r, c, 式か) => { const v = hf.getCellValue({ sheet: sid, row: r, col: c }); return (v === null || v === undefined) ? '' : String(v); };
  if (壊れ1(0, 0, true) === '0') { 素通り++; console.log('  ★素通り★ 直す前の姿を 通してしまった'); }
  /* 壊し② 何でも 0 にする（＝"" まで 0 に する やりすぎ） */
  const 壊れ2 = (r, c) => { const v = hf.getCellValue({ sheet: sid, row: r, col: c }); return (v === null || v === undefined || v === '') ? '0' : String(v); };
  if (壊れ2(2, 0) === '') { 素通り++; console.log('  ★素通り★ やりすぎを 通してしまった'); }
  /* 壊し③ 呼び手が 4つ目を 渡していない形 */
  if (['_hfGetDisplay(sheet, r, c)'].every((s) => s.split(',').length >= 4)) { 素通り++; console.log('  ★素通り★ 渡し忘れを 通してしまった'); }
  if (素通り) { hf.destroy(); console.log('★抜け道 ' + 素通り + '件★'); process.exit(1); }
  console.log('  ok   抜け道 0件');
}
hf.destroy();
process.exit(赤 ? 1 : 0);
