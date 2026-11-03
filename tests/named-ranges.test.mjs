/* named-ranges.test.mjs — ★名前の定義（名前付き範囲）★ 2026-08-29
 *
 *  ★なぜ 自分で 作ったか（実測）★
 *    HyperFormula の 名前の仕組みは ★日本語の名前を 受け付けない★:
 *      uriage … ok ／ 売上・うりあげ・_売上・売上_2026・Sales1 … ★全部 invalid★
 *    司さんの実物は 日本語の名前で 書かれている ⇒ ★式に 入る前に 開く★。
 *
 *  ★実Excelの 決まり（COMで 実測 2026-08-29）★
 *    =SUM(うりあげ) → 60（日本語の名前は 使える）
 *    ★A1 のような 番地の形は 断られる★ ／ ★数で 始まる名前も 断られる★
 *    ブック全体の名前と シートだけの名前（Sheet1!しーとない）が ある
 *
 *  走らせ方: node tests/named-ranges.test.mjs [--self-test]
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

const N = require_(path.join(ROOT, 'lib/named-ranges.js'));

console.log('\n[① 名前の 決まり（実Excelと 同じ）]');
for (const [名, 使える, なぜ] of [
  ['うりあげ', true, ''], ['売上_2026', true, ''], ['Sales', true, ''],
  ['A1', false, '番地の形'], ['ZZ9999', false, '番地の形'],
  ['1あ', false, '数で 始まる'], ['', false, '空'], ['R', false, 'R/C'],
  ['あ い', false, '使えない字'],
]) {
  const r = N.名前を確かめる(名);
  ok((名 || '（空）') + ' … ' + (使える ? '使える' : '断る（' + なぜ + '）'),
    使える ? r === null : r !== null, String(r));
}

console.log('\n[② 名前を 実際の範囲に 開く]');
const 箱 = N.作る();
箱.足す('うりあげ', "='歩合'!A1:A10");
箱.足す('うりあげ合計', "='歩合'!B1");
箱.足す('しーとだけ', "='月別'!C1", '月別');
for (const [式, シート, 期待] of [
  ['=SUM(うりあげ)', null, "=SUM(('歩合'!A1:A10))"],
  ['★長い名前が 先★ =うりあげ合計+1', null, null],
  ['="うりあげ"&1', null, '="うりあげ"&1'],
  ['=しーとだけ', '月別', "=('月別'!C1)"],
  ['=しーとだけ', '歩合', '=しーとだけ'],
  ['=SUM(A1:A3)', null, '=SUM(A1:A3)'],
]) {
  if (期待 === null) continue;
  ok(式 + '（' + (シート || 'ブック') + '）', 箱.開く(式, シート) === 期待, 箱.開く(式, シート));
}
ok('★長い名前から 先に 開く★（うりあげ合計 が うりあげ に 食われない）',
  箱.開く('=うりあげ合計+1') === "=('歩合'!B1)+1", 箱.開く('=うりあげ合計+1'));
ok('★字の中は 触らない★', 箱.開く('="うりあげ"&1') === '="うりあげ"&1', 箱.開く('="うりあげ"&1'));
ok('★関数の名前は 触らない★', 箱.開く('=SUM(1)') === '=SUM(1)', 箱.開く('=SUM(1)'));

console.log('\n[③ ★実際に 計算させる（実Excel＝60）★]');
{
  const HFns = require_(path.join(ROOT, 'hyperformula.full.min.js'));
  const EF = require_(path.join(ROOT, 'exally-formula.js'));
  EF.registerExallyFunctions(HFns);
  const g = globalThis, 前w = g.window;
  const 箱2 = N.作る();
  箱2.足す('うりあげ', "='Sheet1'!A1:A3");
  g.window = { 名前の箱: 箱2 };
  const hf = HFns.HyperFormula.buildEmpty({ licenseKey: 'gpl-v3', useArrayArithmetic: true, smartRounding: false });
  hf.addSheet('Sheet1');
  const sid = hf.getSheetId('Sheet1');
  hf.setSheetContent(sid, [[10, null], [20, null], [30, null]]);
  const 式 = EF.convertFormula('=SUM(うりあげ)');
  hf.setCellContents({ sheet: sid, row: 0, col: 1 }, 式);
  let v = hf.getCellValue({ sheet: sid, row: 0, col: 1 });
  if (v && v.value !== undefined) v = v.value;
  g.window = 前w;
  hf.destroy();
  ok('★=SUM(うりあげ) → 60（実Excelと 同じ）★', v === 60, String(v) + '  開いた式=' + 式);
}
{
  /* ★エンジンに そのまま 渡すと 落ちる事★も 見せる（なぜ 開くのかの 証拠） */
  const HFns = require_(path.join(ROOT, 'hyperformula.full.min.js'));
  const hf = HFns.HyperFormula.buildEmpty({ licenseKey: 'gpl-v3' });
  hf.addSheet('Sheet1');
  let 落ちた = false;
  try { hf.addNamedExpression('うりあげ', "='Sheet1'!A1:A3"); } catch (e) { 落ちた = true; }
  hf.destroy();
  ok('★エンジンは 日本語の名前を 受け付けない（だから 開く）★', 落ちた);
}

console.log('\n[④ 画面に 繋がっている]');
const book = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf8');
ok('部品を 読み込んでいる', /lib\/named-ranges\.js/.test(book));
ok('★window.名前の箱 を 置いている（convertFormula が 見る）★', /window\.名前の箱\s*=/.test(book));
ok('名前の窓が 在る', /id="nameOverlay"/.test(book));
ok('★alert / prompt を 使っていない★（画面が 止まる）', !/\balert\(|\bprompt\(/.test(book));
const EFsrc = fs.readFileSync(path.join(ROOT, 'exally-formula.js'), 'utf8');
ok('★convertFormula が 開いている★', /window\.名前の箱[\s\S]{0,80}開く\(f/.test(EFsrc));

console.log('\n[⑤ リボンから 押せる]');
const ACT = require_(path.join(ROOT, 'lib/ribbon-actions.js'));
for (const 名 of ['名前の定義', '名前の管理']) {
  const g = globalThis, 前w = g.window;
  let 受け = null;
  g.window = { 名前の窓を開く: function () { 受け = 'ok'; } };
  ACT[名]();
  g.window = 前w;
  ok('「' + 名 + '」→ 名前の窓を開く', 受け === 'ok', String(受け));
}

console.log('\nnamed-ranges: ' + 緑 + '/' + (緑 + 赤) + ' passed');

if (壊す) {
  console.log('\n★--self-test＝わざと 壊して 赤に なるか★');
  let 素通り = 0;
  /* 壊し① 短い名前から 開く（長い名前が 壊れる） */
  const 箱3 = N.作る();
  箱3.足す('あ', '=Z1');
  箱3.足す('ああ', '=Z2');
  if (箱3.開く('=ああ') !== '=(Z2)') { 素通り++; console.log('  ★素通り★ 長い名前を 先に 開けていない: ' + 箱3.開く('=ああ')); }
  /* 壊し② 字の中を 開いてしまう形 */
  const 素朴 = (f, 名, 参照) => f.split(名).join(参照);
  if (素朴('="あ"', 'あ', '(Z1)') === '="あ"') { 素通り++; console.log('  ★素通り★ 壊し方が おかしい'); }
  /* 壊し③ A1 を 名前に できてしまう */
  if (N.名前を確かめる('A1') === null) { 素通り++; console.log('  ★素通り★ A1 を 通した'); }
  if (素通り) { console.log('★抜け道 ' + 素通り + '件★'); process.exit(1); }
  console.log('  ok   抜け道 0件');
}
process.exit(赤 ? 1 : 0);
