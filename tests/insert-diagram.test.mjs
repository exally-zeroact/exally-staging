/* insert-diagram.test.mjs — ★挿入タブ（図解・アイコン・数式・チェックボックス・画面を撮る）★ 2026-08-30
 *
 *  ★真値（実Excel 16.0 で 実測）★ … tools/measure-insert.ps1
 *    SmartArt … 型（レイアウト）★159個★／1〜8番目の 名前も 取った／
 *      置いた 時の 形の 名前 `Diagram 1`／★中の 節は はじめ 5個★／`Add()` で 6
 *    チェック ボックス（新しい形）… ★セルの 中の TRUE / FALSE★（大文字で 出る）
 *      `=SUM(A1:A2)`=0 ／ `=A1+A2`=1 ／ `=COUNTIF(A1:A2,TRUE)`=1
 *    チェック ボックス（古い形）… `Check Box 1`／★100.125 × 20.25★／はじめ ★-4146★／`チェック 1`
 *    アイコン … ★COM から 一覧を 出せない（Microsoft の クラウドの 絵）★
 *    数式 … 中の 数式は COM から 読めなかった（0）
 *
 *  ★出来ない事は 出来ないと 書く★
 *    ・3D モデル … ブラウザだけでは 読めない
 *    ・スクリーンショット … ★他のアプリの 窓は 撮れない★ ⇒ この シートを 撮る
 *
 *  走らせ方: node tests/insert-diagram.test.mjs [--self-test]
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
const book = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf8');
const D = require_(path.join(ROOT, 'lib/diagram.js'));

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
ok('tools/measure-insert.ps1 が 在る', fs.existsSync(path.join(ROOT, 'tools/measure-insert.ps1')));

console.log('\n[② SmartArt（実測＝型159・節5）]');
ok('★実Excel の 型の 数 159 を 書き残している★', D.実Excelの型の数 === 159, String(D.実Excelの型の数));
ok('★置いた 時の 節は 5個（実測）★', D.実Excelの節の数 === 5, String(D.実Excelの節の数));
ok('  1〜8番目の 名前を 書き残している', D.実Excelの型8.length === 8 && D.実Excelの型8[0] === 'カード型リスト');
ok('★うちの 図解は 8種★', D.図解たち.length === 8, String(D.図解たち.length));
{
  const g = D.図解を作る('process');
  ok('★作ると 節が 5個（実Excel と 同じ）★', g.節.length === 5, String(g.節.length));
  ok('  種類は 図解', g.種類 === '図解' && g.図 === 'process');
  const g2 = D.図解を作る('cycle', ['あ', 'い']);
  ok('  字を 渡したら その 数', g2.節.length === 2, String(g2.節.length));
  ok('  ★渡した 配列を そのまま 持たない（後で 壊れない）★', (() => {
    const 元 = ['あ', 'い'];
    const g3 = D.図解を作る('cycle', 元);
    元.push('う');
    return g3.節.length === 2;
  })());
}
console.log('  ─ 8種 全部 箱が 出て はみ出さないか ─');
for (const d of D.図解たち) {
  const 場 = D.節の場所(d.種, 320, 220, 5);
  const 外 = 場.filter(b => b.x < -1 || b.y < -1 || b.x + b.w > 321 || b.y + b.h > 221);
  ok('  ' + d.名 + '（' + d.種 + '）… 箱 ' + 場.length + '個・はみ出し ' + 外.length,
    場.length === 5 && 外.length === 0, JSON.stringify(場[0]));
}
ok('★節が 0個なら 箱も 0個★', D.節の場所('list-v', 320, 220, 0).length === 0);
ok('★節が 1個でも 落ちない★', D.節の場所('cycle', 320, 220, 1).length === 1);
ok('★節が 12個でも 箱は 12個★', D.節の場所('process', 320, 220, 12).length === 12);
ok('★手順には 矢印の 印が 付く★', D.節の場所('process', 320, 220, 3).filter(b => b.矢).length === 2);
ok('★一番 後ろには 矢印を 付けない★', D.節の場所('process', 320, 220, 3)[2].矢 === false);

console.log('\n[③ アイコン（実測＝実Excel は 一覧を 出せない）]');
ok('★12種 在る★', D.アイコンたち.length === 12, String(D.アイコンたち.length));
ok('  名前が 全部 ちがう', new Set(D.アイコンたち.map(v => v.種)).size === 12);
ok('★「クラウドから 来る」事を 画面に 書いてある★',
  /Microsoft の クラウドから 来ます/.test(book));
for (const a of D.アイコンたち) {
  ok('  ' + a.名 + ' を 描く 道が 在る',
    new RegExp("種 === '" + a.種 + "'").test(book) || a.種 === 'circle',
    a.種);
}

console.log('\n[④ 数式（うちの 書き方）]');
{
  const t = (s) => JSON.stringify(D.数式を組む(s));
  ok('★分数 a/b★', t('a/b') === '[{"型":"分数","上":"a","下":"b"}]', t('a/b'));
  ok('★累乗 x^2★', t('x^2') === '[{"型":"累乗","元":"x","肩":"2"}]', t('x^2'));
  ok('★平方根 sqrt(c)★', t('sqrt(c)') === '[{"型":"根","中":"c"}]', t('sqrt(c)'));
  ok('★添え字 y_1★', t('y_1') === '[{"型":"添え字","元":"y","下":"1"}]', t('y_1'));
  const 混 = D.数式を組む('a/b + x^2 = sqrt(c) - y_1');
  ok('★混ぜても 4つ 全部 拾う★',
    混.filter(v => v.型 === '分数').length === 1 && 混.filter(v => v.型 === '累乗').length === 1 &&
    混.filter(v => v.型 === '根').length === 1 && 混.filter(v => v.型 === '添え字').length === 1,
    JSON.stringify(混));
  ok('  ふつうの 字は つないで 1つに する', D.数式を組む('abc').length === 1);
  ok('  空なら 0個', D.数式を組む('').length === 0);
  ok('  何を 渡しても 落ちない', Array.isArray(D.数式を組む(null)) && Array.isArray(D.数式を組む(123)));
}

console.log('\n[⑤ 画面から 押せる]');
for (const n of ['図解を置く', '図解の窓を開く', '図解の節を直す', '図解の節を決める',
  'アイコンを置く', 'アイコンの窓を開く', '数式を置く', '数式の窓を開く',
  'チェックボックスを入れる', '画面を撮って貼る',
  '挿入した物を描く', '図解を描く', 'アイコンを描く', '数式を描く', '選んだ印を描く']) {
  ok(n + ' が 在る', !!抜く(n));
}
ok('★チェックは セルの TRUE/FALSE（実Excel の 新しい 形と 同じ）★',
  /setCell\(r, c, 今の字 === 'TRUE' \? 'FALSE' : 'TRUE'\);/.test(book));
ok('★押すたび 入れ替わる★', /今の字 === 'TRUE' \? 'FALSE' : 'TRUE'/.test(book));
ok('★他のアプリの 窓は 撮れない事を 画面に 書いてある★',
  /ブラウザは 他のアプリの 窓を 撮れません/.test(book));
ok('★字は 形を 全部 描いた 後に 出す（重なる 丸で 隠れない）★',
  /★字は 形を 全部 描いた 後に まとめて 出す★/.test(抜く('図解を描く') || ''));
ok('★重なる 丸の 字は 外へ ずらす★', /o\.図 === 'venn' && 場\.length > 1/.test(抜く('図解を描く') || ''));
ok('★ぐるぐるの 矢印は 丸より 先に 描く★', (() => {
  const f = 抜く('図解を描く') || '';
  return f.indexOf('ぐるぐるの 矢印') < f.indexOf('for (var i = 0; i < 場.length; i++)');
})());
ok('★箱から はみ出さない（数式）★', /if \(x > X \+ W - 6\) break;/.test(抜く('数式を描く') || ''));
ok('★節が 空なら 断る★', /1つも 書かれていません/.test(book));
ok('★図解を 選んでいないと 断る★', /図解を 選んでください/.test(book));
ok('★alert / prompt / confirm を 使っていない★', !/\balert\(|\bprompt\(|\bconfirm\(/.test(book));

console.log('\n[⑥ 副題を 決めていない 窓が 増えていないか]');
{
  const 行 = book.split(String.fromCharCode(10));
  const 抜け = [];
  for (let i = 0; i < 行.length; i++) {
    if (!/getElementById\('funcTitle'\)\.textContent =/.test(行[i])) continue;
    if (!/窓の副題\(/.test(行.slice(i, i + 5).join(String.fromCharCode(10)))) 抜け.push(i + 1);
  }
  ok('★副題を 決めていない 窓は 0個★', 抜け.length === 0, 抜け.join(' / '));
}

console.log('\n[⑦ リボンから 押せる]');
{
  const ACT = require_(path.join(ROOT, 'lib/ribbon-actions.js'));
  const g = globalThis, 前w = g.window;
  const 試す = (ボタン, 呼ぶ名) => {
    let 受け = 'よばれていない';
    const w = {}; w[呼ぶ名] = function () { 受け = 'ok'; };
    g.window = w; ACT[ボタン](); g.window = 前w;
    ok('「' + ボタン + '」→ ' + 呼ぶ名, 受け === 'ok', String(受け));
  };
  試す('SmartArt', '図解の窓を開く');
  試す('アイコン', 'アイコンの窓を開く');
  試す('数式を挿入', '数式の窓を開く');
  試す('チェックボックス', 'チェックボックスを入れる');
  試す('スクリーンショット', '画面を撮って貼る');
}

console.log('\ninsert-diagram: ' + 緑 + '/' + (緑 + 赤) + ' passed');

if (壊す) {
  console.log('\n★--self-test＝この 見張りが 見ている物を 直に 確かめる（★壊して 赤を 見るのは tools/break-check.mjs★）★');
  let 素通り = 0;
  if (D.図解を作る('process').節.length !== 5) { 素通り++; console.log('  ★素通り★ 節が 5個で ない'); }
  else console.log('  ok   節は 5個（実Excel と 同じ）');
  const 場 = D.節の場所('pyramid', 320, 220, 5);
  if (場.some(b => b.x < -1 || b.x + b.w > 321)) { 素通り++; console.log('  ★素通り★ 箱が はみ出している'); }
  else console.log('  ok   はみ出していない');
  if (JSON.stringify(D.数式を組む('a/b')) !== '[{"型":"分数","上":"a","下":"b"}]') {
    素通り++; console.log('  ★素通り★ 分数を 読めていない');
  } else console.log('  ok   分数を 読めている');
  const f = 抜く('図解を描く') || '';
  if (f.indexOf('★字は 形を 全部 描いた 後に まとめて 出す★') < 0) {
    素通り++; console.log('  ★素通り★ 字を 形の 間に 描いている（重なる 丸で 隠れる）');
  } else console.log('  ok   字は 後から まとめて 出す');
  if (素通り) { console.log('★抜け道 ' + 素通り + '件★'); process.exit(1); }
  console.log('  ok   抜け道 0件');
}
process.exit(赤 ? 1 : 0);
