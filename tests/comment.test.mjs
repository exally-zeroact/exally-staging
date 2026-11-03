/* comment.test.mjs — ★コメント（校閲タブ）★ 2026-08-29
 *
 *  ★真値（実Excel 16.0 で 実測）★
 *    ・セルに 付く（ref="A1"）／★作者が 付く★／★既定では 見えない★（Visible=False）
 *    ・ファイルの 中では xl/comments1.xml（<authors> と <comment ref="A1">）
 *  ⇒ うちも ★赤い印だけ 出す★（実Excelと 同じ 見せ方）
 *
 *  走らせ方: node tests/comment.test.mjs [--self-test]
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
function 抜く(名) {
  const i = book.indexOf('function ' + 名 + '(');
  if (i < 0) return null;
  let d = 0;
  const j = book.indexOf('{', i);
  for (let k = j; k < book.length; k++) {
    if (book[k] === '{') d++;
    else if (book[k] === '}') { d--; if (d === 0) return book.slice(i, k + 1); }
  }
  return null;
}

console.log('\n[① 画面に 在る]');
for (const n of ['コメントの箱', 'コメントを取る', 'コメントの窓を開く', 'コメントを決める',
  'コメントを消す', 'コメントへ移る', 'コメントの一覧']) ok(n + ' が 在る', !!抜く(n));
ok('★窓が 在る（prompt を 使わない）★', /id="commentOverlay"/.test(book));
ok('★alert / prompt / confirm を 使っていない★', !/\balert\(|\bprompt\(|\bconfirm\(/.test(book));

console.log('\n[② ★赤い印を 描く（実Excelと 同じ 見せ方）★]');
ok('印を 描く所が 在る', /sheets\[activeSheet\]\.comments/.test(book) && /#E53935/.test(book));
/* ★2026-09-03 に 直した★＝★空白 1つで 死ぬ 形だった★
   `moveTo(x+w-6, y)` と 書き直した だけで ★見つからない＝赤★に なった（実際に なった）。
   ★字の 並びで 探す 見張りは 書き直しで 死ぬ★（今日 何度も 出た 型）
   ⇒★空白は 幾つ 在っても よい形★に する。 */
ok('★右上に 描く★（moveTo(x+w-6, y)）',
  /moveTo\(\s*x\s*\+\s*w\s*-\s*6\s*,\s*y\s*\)/.test(book));

console.log('\n[③ 実際に 走らせる]');
function 台(選 = [0, 0]) {
  const s = { data: {}, comments: {} };
  const 出た = [];
  const 動いた = [];
  const 本文 = ['コメントの箱', 'コメントを取る', 'コメントを決める', 'コメントを消す',
    'コメントへ移る', 'コメントの一覧'].map(抜く).join('\n');
  const f = new Function('sheets', 'activeSheet', 'selR1', 'selC1', 'document', 'render', 'notify',
    'hideCtxMenu', 'cellAddr', 'sel', 'updateBar', 'コメントの窓を開く', 'コメントの窓を閉じる',
    本文 + '\nreturn { 箱: コメントの箱, 取る: コメントを取る, 決める: コメントを決める, 消す: コメントを消す, 移る: コメントへ移る, 一覧: コメントの一覧 };');
  const 欄 = { value: '' };
  const にせDOM = { getElementById: (id) => (id === 'commentText' ? 欄 : { style: {}, textContent: '', value: '' }) };
  const api = f([s], 0, 選[0], 選[1], にせDOM, function () {}, function (m) { 出た.push(m); },
    function () {}, function (r, c) { return 'R' + r + 'C' + c; },
    function (r1, c1) { 動いた.push([r1, c1]); }, function () {}, function () {}, function () {});
  return { s, api, 欄, 出た, 動いた };
}

{
  const t = 台([0, 0]);
  t.欄.value = 'ここ 見て';
  t.api.決める();
  ok('★コメントが 付く★', !!t.s.comments['0,0'], JSON.stringify(t.s.comments));
  ok('中身が 入る', t.s.comments['0,0'].文 === 'ここ 見て', t.s.comments['0,0'].文);
  ok('★作者が 付く（実Excelと 同じ）★', !!t.s.comments['0,0'].誰, String(t.s.comments['0,0'].誰));
  ok('★日付が 付く★', /\d{4}\/\d+\/\d+/.test(t.s.comments['0,0'].いつ), String(t.s.comments['0,0'].いつ));
  ok('★つけたと 言う★', /つけました/.test(t.出た.join('')), JSON.stringify(t.出た));
}
{
  const t = 台([0, 0]);
  t.欄.value = 'あ';
  t.api.決める();
  t.欄.value = '';
  t.api.決める();
  ok('★空にすると 消える★', !t.s.comments['0,0'], JSON.stringify(t.s.comments));
  ok('★消したと 言う★', /消しました/.test(t.出た.join('')), JSON.stringify(t.出た));
}
{
  const t = 台([0, 0]);
  t.s.comments['0,0'] = { 文: 'あ', 誰: 'x', いつ: '2026/8/29' };
  t.api.消す();
  ok('消すボタンで 消える', !t.s.comments['0,0']);
}

console.log('\n[④ 次／前の コメントへ 移る]');
{
  const t = 台([0, 0]);
  t.s.comments['0,0'] = { 文: 'a', 誰: 'x', いつ: '' };
  t.s.comments['2,1'] = { 文: 'b', 誰: 'x', いつ: '' };
  t.s.comments['5,0'] = { 文: 'c', 誰: 'x', いつ: '' };
  t.api.移る(1);
  ok('★次は 行の順で 次★', JSON.stringify(t.動いた[0]) === '[2,1]', JSON.stringify(t.動いた));
  t.動いた.length = 0;
  t.api.移る(-1);
  ok('★前は 最後へ 回る（今が 先頭なので）★', JSON.stringify(t.動いた[0]) === '[5,0]', JSON.stringify(t.動いた));
}
{
  const t = 台([0, 0]);
  ok('★1つも 無い時は 言う★', t.api.移る(1) === false);
  ok('★「まだ ありません」と 出す★', /まだ ありません/.test(t.出た.join('')), JSON.stringify(t.出た));
}

console.log('\n[⑤ 一覧]');
{
  const t = 台([0, 0]);
  t.s.comments['0,0'] = { 文: 'あ', 誰: 'x', いつ: '' };
  t.s.comments['1,0'] = { 文: 'い', 誰: 'y', いつ: '' };
  ok('★2個 あると 言う★', t.api.一覧() === 2, String(t.api.一覧()));
  ok('★中身も 出す★', /あ/.test(t.出た.join('')) && /い/.test(t.出た.join('')), JSON.stringify(t.出た));
}

console.log('\n[⑥ リボンから 押せる]');
const ACT = require_(path.join(ROOT, 'lib/ribbon-actions.js'));
for (const [名, 行き先] of [['新しいコメント', 'コメントの窓を開く'], ['コメントを消す', 'コメントを消す'],
  ['前のコメント', '前のコメントへ'], ['次のコメント', '次のコメントへ'], ['コメントの表示', 'コメントの一覧']]) {
  const g = globalThis, 前w = g.window;
  let 受け = null;
  g.window = {}; g.window[行き先] = function () { 受け = 行き先; };
  ACT[名]();
  g.window = 前w;
  ok('「' + 名 + '」→ ' + 行き先, 受け === 行き先, String(受け));
}

console.log('\ncomment: ' + 緑 + '/' + (緑 + 赤) + ' passed');

if (壊す) {
  console.log('\n★--self-test＝この 見張りが 見ている物を 直に 確かめる（★壊して 赤を 見るのは tools/break-check.mjs★）★');
  let 素通り = 0;
  const t = 台([0, 0]);
  /* 壊し① 空でも 消さない */
  t.欄.value = 'あ'; t.api.決める();
  t.欄.value = ''; t.api.決める();
  if (t.s.comments['0,0']) { 素通り++; console.log('  ★素通り★ 空にしても 消えない'); }
  /* 壊し② 作者を 付けない */
  t.欄.value = 'い'; t.api.決める();
  if (!t.s.comments['0,0'].誰) { 素通り++; console.log('  ★素通り★ 作者が 付いていない'); }
  /* 壊し③ 印の 色が 赤で ない */
  if (!/#E53935/.test(book)) { 素通り++; console.log('  ★素通り★ 赤い印が 無い'); }
  if (素通り) { console.log('★抜け道 ' + 素通り + '件★'); process.exit(1); }
  console.log('  ok   抜け道 0件');
}
process.exit(赤 ? 1 : 0);
