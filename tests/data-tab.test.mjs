/* data-tab.test.mjs — ★データタブ（フラッシュ フィル・再適用・ゴール シーク）★ 2026-08-30
 *
 *  ★真値（実Excel 16.0 で 実測）★ … tools/measure-data-tab.ps1
 *    ・フラッシュ フィル … A列「やまだ たろう」…／B1 に「やまだ」だけ 打って 呼ぶと
 *        B1='やまだ' B2='すずき' B3='さとう'（★1つの 見本から 覚えて 下まで 埋める★）
 *    ・ゴール シーク … D2=D1*3+5 で 50 を 目指すと ★D1=15★
 *        反復の 決め＝★多くて 100回／変化の下限 0.001★
 *    ・再適用 … 絞った後に 値を 直しても 数は 変わらず（2行）、
 *        ★ApplyFilter を 呼ぶと 3行に なる★
 *
 *  ★未測定★ 実Excel が どんな 決まりを どの順で 当てるかは 外から 読めない
 *    ⇒ うちの 決まりの 並びは lib/flash-fill.js に 理由つきで 書いた
 *
 *  走らせ方: node tests/data-tab.test.mjs [--self-test]
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

const F = require_(path.join(ROOT, 'lib/flash-fill.js'));
const book = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf8');
const lib = fs.readFileSync(path.join(ROOT, 'lib/flash-fill.js'), 'utf8');

function 抜く(名) {
  const i = book.indexOf('function ' + 名 + '(');
  if (i < 0) return null;
  let d = 0; const j = book.indexOf('{', i);
  for (let k = j; k < book.length; k++) {
    if (book[k] === '{') d++;
    else if (book[k] === '}') { d--; if (d === 0) return book.slice(i, k + 1); }
  }
  return null;
}

console.log('\n[① 測った 道具が 残っている]');
ok('tools/measure-data-tab.ps1 が 在る', fs.existsSync(path.join(ROOT, 'tools/measure-data-tab.ps1')));
ok('★決まりの 並びは うちの物★ と 書いてある', /未測定/.test(lib) && /うちの物/.test(lib));

console.log('\n[② フラッシュ フィル（実測と 同じ 出方）]');
{
  const やる = (例, 元) => { const r = F.覚える(例); return r ? r.やる(元) : null; };
  /* ★実測そのもの★ */
  ok('★「やまだ」1つで 姓を 覚える（実測）★',
    やる([{ 元: ['やまだ たろう'], 答え: 'やまだ' }], ['すずき はなこ']) === 'すずき',
    String(やる([{ 元: ['やまだ たろう'], 答え: 'やまだ' }], ['すずき はなこ'])));
  ok('  3つ目も 合う',
    やる([{ 元: ['やまだ たろう'], 答え: 'やまだ' }], ['さとう じろう']) === 'さとう');
  ok('名だけ（2番目）', やる([{ 元: ['やまだ たろう'], 答え: 'たろう' }], ['すずき はなこ']) === 'はなこ');
  ok('全角の 空白でも 分ける',
    やる([{ 元: ['山田　太郎'], 答え: '山田' }], ['鈴木　花子']) === '鈴木');
  ok('カンマでも 分ける', やる([{ 元: ['a,b,c'], 答え: 'b' }], ['x,y,z']) === 'y');
  ok('前から 3文字', やる([{ 元: ['ABCDEF'], 答え: 'ABC' }], ['XYZ123']) === 'XYZ');
  ok('後ろから 2文字', やる([{ 元: ['ABCDEF'], 答え: 'EF' }], ['XYZ123']) === '23');
  ok('隣を つなぐ（間の 字も 覚える）',
    やる([{ 元: ['山田', '太郎'], 答え: '山田 太郎' }], ['鈴木', '花子']) === '鈴木 花子');
  ok('前に 付ける', やる([{ 元: ['1234'], 答え: '〒1234' }], ['5678']) === '〒5678');
  ok('後ろに 付ける', やる([{ 元: ['5'], 答え: '5円' }], ['9']) === '9円');
  ok('★分からない時は null（何も しない）★',
    F.覚える([{ 元: ['あ'], 答え: 'ぜんぜん違う' }]) === null);
  ok('★答えが 空なら null★', F.覚える([{ 元: ['あ'], 答え: '' }]) === null);
  ok('★例が 無ければ null★', F.覚える([]) === null);
  /* ★見せた 例で 合うか 確かめてから 使う★ */
  ok('★2つの 例が 食い違えば null★',
    F.覚える([{ 元: ['a b'], 答え: 'a' }, { 元: ['x y'], 答え: 'y' }]) === null,
    JSON.stringify(F.覚える([{ 元: ['a b'], 答え: 'a' }, { 元: ['x y'], 答え: 'y' }])));
  ok('  2つとも 同じ 決まりなら 覚える',
    やる([{ 元: ['a b'], 答え: 'a' }, { 元: ['x y'], 答え: 'x' }], ['p q']) === 'p');
  /* 名前が 人に 分かる形か */
  const r = F.覚える([{ 元: ['やまだ たろう'], 答え: 'やまだ' }]);
  ok('★何を 覚えたか 名前で 分かる★', /空白/.test(r.名) && /1番目/.test(r.名), r.名);
}

console.log('\n[③ 画面に つながっている]');
for (const n of ['フラッシュフィル', '絞りを再適用', 'ゴールシークを開く', 'ゴールシークをやる']) {
  ok(n + ' が 在る', !!抜く(n));
}
ok('部品を 読み込んでいる', /src="lib\/flash-fill\.js/.test(book));
ok('★Ctrl+E で 出る（実Excelと 同じ）★', /if\(ek==='e'\)\{[^}]*フラッシュフィル\(\)/.test(book));
ok('★見本が 無ければ 断る★', /先に 1つ 見本を 打ってください/.test(book));
ok('★分からなければ 何も 直さないと 言う★', /どう すればよいか 分かりませんでした（何も 直していません）/.test(book));
ok('★左に 元の 列が 無ければ 断る★', /左に 元の 列が ありません/.test(book));
ok('★絞った時に 覚えて 再適用できる★', /前の絞り = \{ 表: _絞る表, 列: _絞る列, 残す: 残す \}/.test(book));
ok('★まだ 絞っていなければ 断る★', /まだ 絞っていません/.test(book));

console.log('\n[④ ゴール シーク（実測＝100回・0.001）]');
ok('★回数 100（実測）★', /var ゴールの回数 = 100/.test(book));
ok('★変化の下限 0.001（実測）★', /ゴールの下限 = 0\.001/.test(book));
ok('窓が 在る', /id="goalOverlay"/.test(book));
ok('★式で なければ 断る★', /式の セルは 式で なければ なりません/.test(book));
ok('★同じ セルは 選べない★', /同じ セルは 選べません/.test(book));
ok('★見つからなければ 元に 戻す★', /元に 戻しました/.test(book));
ok('★入れたら その場で 計算し直す（後回しでは 進まない）★',
  /setCell\(変\.r, 変\.c, String\(x\), true\);\s*recalcSheet\(activeSheet, d\);/.test(book));
{
  /* はさみうちの 中身だけ 取り出して 実測の 例（D2=D1*3+5 → 50 なら D1=15）を 解かせる */
  let x = 10;
  const 目標 = 50;
  const f = (v) => v * 3 + 5;
  let x0 = x, x1 = x * 1.1, f0 = f(x0) - 目標, f1 = f(x1) - 目標, 答え = null;
  for (let i = 0; i < 100; i++) {
    if (Math.abs(f1) <= 0.001) { 答え = x1; break; }
    if (f1 === f0) break;
    const x2 = x1 - f1 * (x1 - x0) / (f1 - f0);
    const f2 = f(x2) - 目標;
    x0 = x1; f0 = f1; x1 = x2; f1 = f2;
  }
  ok('★実測の 例が 解ける（D1=15）★', 答え !== null && Math.abs(答え - 15) < 0.001, String(答え));
}
ok('★alert / prompt / confirm を 使っていない★', !/\balert\(|\bprompt\(|\bconfirm\(/.test(book));

console.log('\n[⑤ リボンから 押せる]');
{
  const ACT = require_(path.join(ROOT, 'lib/ribbon-actions.js'));
  const g = globalThis, 前w = g.window;
  for (const [ボタン, 呼ぶ名] of [
    ['フラッシュフィル', 'フラッシュフィル'], ['再適用', '絞りを再適用'], ['WhatIf分析', 'ゴールシークを開く'],
  ]) {
    let 受け = null;
    const w = {}; w[呼ぶ名] = function () { 受け = 'ok'; };
    g.window = w;
    ACT[ボタン]();
    g.window = 前w;
    ok('「' + ボタン + '」→ ' + 呼ぶ名, 受け === 'ok', String(受け));
  }
}

console.log('\ndata-tab: ' + 緑 + '/' + (緑 + 赤) + ' passed');

if (壊す) {
  console.log('\n★--self-test＝この 見張りが 見ている物を 直に 確かめる（★壊して 赤を 見るのは tools/break-check.mjs★）★');
  let 素通り = 0;
  /* 壊し① 例に 合わない 決まりを 返していないか */
  const r = F.覚える([{ 元: ['a b'], 答え: 'a' }, { 元: ['x y'], 答え: 'y' }]);
  if (r) { 素通り++; console.log('  ★素通り★ 食い違う 例で 決まりを 返した … ' + r.名); }
  else console.log('  ok   食い違えば 返さない');
  /* 壊し② 何でも「前から n文字」に なっていないか */
  const r2 = F.覚える([{ 元: ['やまだ たろう'], 答え: 'たろう' }]);
  if (!r2 || /前から/.test(r2.名)) { 素通り++; console.log('  ★素通り★ 名を「前から」と 覚えた'); }
  else console.log('  ok   名は 区切りで 覚える（' + r2.名 + '）');
  /* 壊し③ 空の 答えを 覚えていないか */
  if (F.覚える([{ 元: ['あ'], 答え: '' }]) !== null) { 素通り++; console.log('  ★素通り★ 空を 覚えた'); }
  else console.log('  ok   空は 覚えない');
  if (素通り) { console.log('★抜け道 ' + 素通り + '件★'); process.exit(1); }
  console.log('  ok   抜け道 0件');
}
process.exit(赤 ? 1 : 0);
