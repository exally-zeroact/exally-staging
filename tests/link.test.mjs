/* link.test.mjs — ★ハイパーリンク（挿入→リンク）★ 2026-08-29
 *
 *  ★真値（実Excel 16.0 で 実測）★
 *    行き先（Address）／画面の字（TextToDisplay）／説明（ScreenTip）を 持つ
 *    ★字の色 BGR 867846 ＝ RGB #467886★／★下線 あり★
 *
 *  ★一番 大事なのは 安全★
 *    `javascript:` と `data:` は ★入れさせない★。
 *    リンクは ★他の人が 作った書類から 来る★＝押した人の 画面で 何かを 走らせる 道に なる。
 *    開く時も ★別の窓＋noopener★（元の画面を 乗っ取らせない）。
 *
 *  走らせ方: node tests/link.test.mjs [--self-test]
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
for (const n of ['リンクの箱', 'リンクの窓を開く', 'リンクの先を確かめる', 'リンクを決める',
  'リンクを消す', 'リンクを開く']) ok(n + ' が 在る', !!抜く(n));
ok('窓が 在る', /id="linkOverlay"/.test(book));

console.log('\n[② ★危ない行き先を 弾く★]');
{
  const 確かめる = new Function(抜く('リンクの先を確かめる') + '\nreturn リンクの先を確かめる;')();
  for (const [先, 通す, なぜ] of [
    ['https://example.com', true, ''],
    ['http://example.com', true, ''],
    ['mailto:a@example.com', true, ''],
    ['#A1', true, 'シートの中'],
    ['javascript:alert(1)', false, '★走らせる道★'],
    ['JavaScript:alert(1)', false, '★大文字でも★'],
    ['  javascript:alert(1)', false, '★前に 空白が 在っても★'],
    ['data:text/html,<script>', false, '★中身を 埋め込む道★'],
    ['', false, '空'],
    ['ただの字', false, '形が 違う'],
  ]) {
    const r = 確かめる(先);
    ok((通す ? '通す' : '★弾く★') + ' … ' + (先 || '（空）') + (なぜ ? '  ' + なぜ : ''),
      通す ? r === null : r !== null, String(r));
  }
}

console.log('\n[③ つける・消す]');
{
  const s = { data: { '0,0': { v: '', f: '', d: '' } }, links: {} };
  const 出た = [];
  const 欄 = { linkUrl: { value: 'https://example.com' }, linkText: { value: '見に行く' }, linkTip: { value: '説明' },
    linkWhy: { textContent: '' }, linkDel: { style: {} } };
  const f = new Function('sheets', 'activeSheet', 'selR1', 'selC1', 'document', 'setCell',
    'render', 'updateBar', 'notify', 'リンクの窓を閉じる',
    抜く('リンクの箱') + '\n' + 抜く('リンクの先を確かめる') + '\n' + 抜く('リンクを決める') + '\n' + 抜く('リンクを消す')
    + '\nreturn { 決める: リンクを決める, 消す: リンクを消す, 箱: リンクの箱 };');
  const api = f([s], 0, 0, 0, { getElementById: (id) => 欄[id] || { value: '', style: {}, textContent: '' } },
    function (r, c, v) { s.data[r + ',' + c] = s.data[r + ',' + c] || {}; s.data[r + ',' + c].v = v; s.data[r + ',' + c].d = v; },
    function () {}, function () {}, function (m) { 出た.push(m); }, function () {});
  api.決める();
  ok('★リンクが 付く★', !!s.links['0,0'], JSON.stringify(s.links));
  ok('行き先が 入る', s.links['0,0'].先 === 'https://example.com', String(s.links['0,0'].先));
  ok('★画面の字も 入る★', s.data['0,0'].v === '見に行く', String(s.data['0,0'].v));
  ok('★実Excelと 同じ 色（#467886）★', s.data['0,0'].color === '#467886', String(s.data['0,0'].color));
  ok('★下線が 付く★', s.data['0,0'].underline === true, String(s.data['0,0'].underline));
  ok('★つけたと 言う★', /つけました/.test(出た.join('')), JSON.stringify(出た));
  api.消す();
  ok('★消えると 色も 下線も 外れる★',
    !s.links['0,0'] && s.data['0,0'].color === undefined && s.data['0,0'].underline === undefined,
    JSON.stringify(s.data['0,0']));
}
{
  /* 危ない物は ★入らない★ */
  const s = { data: {}, links: {} };
  const 欄 = { linkUrl: { value: 'javascript:alert(1)' }, linkText: { value: 'x' }, linkTip: { value: '' },
    linkWhy: { textContent: '' }, linkDel: { style: {} } };
  const f = new Function('sheets', 'activeSheet', 'selR1', 'selC1', 'document', 'setCell',
    'render', 'updateBar', 'notify', 'リンクの窓を閉じる',
    抜く('リンクの箱') + '\n' + 抜く('リンクの先を確かめる') + '\n' + 抜く('リンクを決める')
    + '\nreturn リンクを決める;');
  const n = f([s], 0, 0, 0, { getElementById: (id) => 欄[id] || { value: '', style: {}, textContent: '' } },
    function () {}, function () {}, function () {}, function () {}, function () {})();
  ok('★危ない行き先は 入らない（0）★', n === 0, String(n));
  ok('★箱も 空のまま★', !s.links['0,0'], JSON.stringify(s.links));
  ok('★なぜ ダメかを 画面に 出す★', /javascript/.test(欄.linkWhy.textContent), 欄.linkWhy.textContent);
}

console.log('\n[④ 開く時も 安全]');
ok('★別の窓で 開く★', /window\.open\(l\.先, '_blank'/.test(book));
ok('★noopener を 付ける（元の画面を 乗っ取らせない）★', /noopener,noreferrer/.test(book));
ok('★alert / prompt / confirm を 使っていない★', !/\balert\(|\bprompt\(|\bconfirm\(/.test(book));

console.log('\n[⑤ リボンから 押せる]');
const ACT = require_(path.join(ROOT, 'lib/ribbon-actions.js'));
{
  const g = globalThis, 前w = g.window;
  let 受け = null;
  g.window = { リンクの窓を開く: function () { 受け = 'ok'; } };
  ACT['リンク']();
  g.window = 前w;
  ok('「リンク」→ リンクの窓を開く', 受け === 'ok', String(受け));
}

console.log('\n[⑥ ★押したら 飛ぶ（実Excel と同じ）★]');
{
  /* ★実ブラウザで 実測した（2026-08-30・Playwright）★
       ・リンクのセルを 普通に押す → window.open('https://example.com/','_blank','noopener,noreferrer') が 1回
       ・★Ctrl を押しながら → 0回（選ぶだけ）★
       ・リンクの無いセル → 0回
     ここでは その線が ★外れていないか★ を 見張る。 */
  const md = 抜く('onMD') || '';
  ok('セルを押す所から リンクを開く を 呼んでいる', /リンクを開く\(r,c\)/.test(md), md.slice(-160));
  ok('★Ctrl（⌘）を押しながらなら 飛ばない道が 在る★',
    /!e\.ctrlKey\s*&&\s*!e\.metaKey/.test(md), md.slice(-160));
  ok('★選んでから 呼ぶ（選択が 先）★',
    md.indexOf('sel(r,c,r,c)') < md.indexOf('リンクを開く(r,c)'), '順番');
}

console.log('\nlink: ' + 緑 + '/' + (緑 + 赤) + ' passed');

if (壊す) {
  console.log('\n★--self-test＝この 見張りが 見ている物を 直に 確かめる（★壊して 赤を 見るのは tools/break-check.mjs★）★');
  let 素通り = 0;
  const 確かめる = new Function(抜く('リンクの先を確かめる') + '\nreturn リンクの先を確かめる;')();
  /* 壊し① 前に 空白を 付けて すり抜ける */
  if (確かめる('  javascript:alert(1)') === null) { 素通り++; console.log('  ★素通り★ 空白つきで すり抜けた'); }
  /* 壊し② 大文字で すり抜ける */
  if (確かめる('JAVASCRIPT:alert(1)') === null) { 素通り++; console.log('  ★素通り★ 大文字で すり抜けた'); }
  /* 壊し③ noopener が 無い */
  if (!/noopener/.test(book)) { 素通り++; console.log('  ★素通り★ noopener が 無い'); }
  if (素通り) { console.log('★抜け道 ' + 素通り + '件★'); process.exit(1); }
  console.log('  ok   抜け道 0件');
}
process.exit(赤 ? 1 : 0);
