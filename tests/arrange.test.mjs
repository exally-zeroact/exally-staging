/* arrange.test.mjs — ★配置（そろえる・まとめる・回す）と 拡大縮小印刷★ 2026-08-30
 *
 *  ★真値（実Excel 16.0 で 実測）★ … tools/measure-align.ps1 / tools/measure-fit.ps1
 *    A(左10 上10 60x40)・B(左100 上80 80x20) を そろえると
 *      0 → 両方 左=10        ＝★左ぞろえ★
 *      1 → A左65 / B左55     ＝★左右の 中央★（囲みの 真ん中 95）
 *      2 → A左120 / B左100   ＝★右ぞろえ★（囲みの 右 180）
 *      3 → 両方 上=10        ＝★上ぞろえ★
 *      4 → A上35 / B上45     ＝★上下の 中央★（囲みの 真ん中 55）
 *      5 → A上60 / B上80     ＝★下ぞろえ★（囲みの 下 100）
 *      6 → ★出来ない★
 *    グループ化 … 'Group 3'・中の数 2／解除で 元の 2つに 戻る
 *    回転 … 90 → 90度／さらに 90 足して 180度
 *    拡大縮小印刷 … 既定 Zoom=100・横1・縦1／Zoom を False で「ページに 合わせる」／
 *      Zoom は ★10〜400★（401 も 9 も 入らない）
 *
 *  走らせ方: node tests/arrange.test.mjs [--self-test]
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
const GP = require_(path.join(ROOT, 'lib/grid-print.js'));

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
ok('tools/measure-align.ps1 が 在る', fs.existsSync(path.join(ROOT, 'tools/measure-align.ps1')));
ok('tools/measure-fit.ps1 が 在る', fs.existsSync(path.join(ROOT, 'tools/measure-fit.ps1')));

console.log('\n[② ★そろえ方が 実測と 同じ数に なるか★]');
{
  const f = new Function('notify', 'render', 'hideCtxMenu', '今の物たち', '囲みを取る', 'そろえ方',
    抜く('物をそろえる') + '\nreturn 物をそろえる;');
  const 囲み = new Function(抜く('囲みを取る') + '\nreturn 囲みを取る;')();
  const 走る = (番) => {
    const A = { 名: 'A', x: 10, y: 10, w: 60, h: 40 };
    const B = { 名: 'B', x: 100, y: 80, w: 80, h: 20 };
    f(() => {}, () => {}, () => {}, () => [A, B], 囲み,
      [{ 番: 0, 名: '左' }, { 番: 1, 名: '横中' }, { 番: 2, 名: '右' },
       { 番: 3, 名: '上' }, { 番: 4, 名: '縦中' }, { 番: 5, 名: '下' }])(番);
    return { A: A.x + ',' + A.y, B: B.x + ',' + B.y };
  };
  ok('★0 → 両方 左=10（実測）★', 走る(0).A.split(',')[0] === '10' && 走る(0).B.split(',')[0] === '10',
    JSON.stringify(走る(0)));
  ok('★1 → A左65 / B左55（実測）★',
    走る(1).A.split(',')[0] === '65' && 走る(1).B.split(',')[0] === '55', JSON.stringify(走る(1)));
  ok('★2 → A左120 / B左100（実測）★',
    走る(2).A.split(',')[0] === '120' && 走る(2).B.split(',')[0] === '100', JSON.stringify(走る(2)));
  ok('★3 → 両方 上=10（実測）★',
    走る(3).A.split(',')[1] === '10' && 走る(3).B.split(',')[1] === '10', JSON.stringify(走る(3)));
  ok('★4 → A上35 / B上45（実測）★',
    走る(4).A.split(',')[1] === '35' && 走る(4).B.split(',')[1] === '45', JSON.stringify(走る(4)));
  ok('★5 → A上60 / B上80（実測）★',
    走る(5).A.split(',')[1] === '60' && 走る(5).B.split(',')[1] === '80', JSON.stringify(走る(5)));
}
{
  const 囲み = new Function(抜く('囲みを取る') + '\nreturn 囲みを取る;')();
  const w = 囲み([{ x: 10, y: 10, w: 60, h: 40 }, { x: 100, y: 80, w: 80, h: 20 }]);
  ok('★囲みは 左10 右180 上10 下100★',
    w.左 === 10 && w.右 === 180 && w.上 === 10 && w.下 === 100, JSON.stringify(w));
}
ok('★2つ 選んでいないと 断る★', /2つ以上 選んでください/.test(book));
ok('★Ctrl を 押しながらで 足せる★', /if\(e\.ctrlKey \|\| e\.metaKey\)\{ 物を選び足す\(_o\); render\(\); return; \}/.test(book));

console.log('\n[③ まとめる・解く]');
for (const n of ['グループにする', 'グループを解く', 'グループの窓を開く', '物を選び足す',
  '今の物たち', '囲みを取る', '物を回す', '物の回転を戻す', '回転の窓を開く', '配置の窓を開く']) {
  ok(n + ' が 在る', !!抜く(n));
}
ok('★中の 物と ずれを 覚える（解いたら 元の 場所へ）★', /ずれX: o\.x - w\.左, ずれY: o\.y - w\.上/.test(book));
ok('★まとめた 物を 選んでいないと 断る★', /まとめた 物を 選んでください/.test(book));
ok('★解いた後も 中の 物が 選ばれたまま（実Excelと 同じ）★', /えらんだ物たち = 出した;/.test(book));
ok('★部品が グループを 描ける★',
  /o\.種類 === 'グループ'/.test(fs.readFileSync(path.join(ROOT, 'lib/objects.js'), 'utf8')));
ok('★画面へ 写す 時に 持ち物を 数え上げない（中身・回転が 落ちない）★',
  /for \(var k in o\) if \(Object\.prototype\.hasOwnProperty\.call\(o, k\)\) 写し\[k\] = o\[k\];/.test(book));
ok('★選んだ 印は 選んだ物たち 全部に 付く★', /えらんだ物たち\.indexOf\(o\) >= 0/.test(book));
ok('★回転も 部品が 描く★',
  /ctx\.rotate\(o\.回転 \* Math\.PI \/ 180\)/.test(fs.readFileSync(path.join(ROOT, 'lib/objects.js'), 'utf8')));
{
  const f = new Function('notify', 'render', 'hideCtxMenu', '今の物たち',
    抜く('物を回す') + '\nreturn 物を回す;');
  const o = { x: 0, y: 0, w: 10, h: 10 };
  const 回す = f(() => {}, () => {}, () => {}, () => [o]);
  回す(90);
  ok('★90 で 90度（実測）★', o.回転 === 90, String(o.回転));
  回す(90);
  ok('★さらに 90 で 180度（実測）★', o.回転 === 180, String(o.回転));
  回す(-90);
  ok('マイナスでも 回る', o.回転 === 90, String(o.回転));
  回す(360);
  ok('360 で 一周（90 の まま）', o.回転 === 90, String(o.回転));
}

console.log('\n[④ 拡大縮小印刷（実測＝100／横1／縦1／10〜400）]');
for (const n of ['拡大縮小の初期化', '拡大縮小を開く', '拡大縮小を決める', '印刷倍率を足す',
  '拡大縮小の見せ方']) ok(n + ' が 在る', !!抜く(n));
ok('窓が 在る', /id="fitOverlay"/.test(book));
ok('★はじめは 100％（実測）★', /ページ設定\.印刷倍率 = 100;/.test(book));
ok('★はじめは 横1・縦1（実測）★',
  /ページ設定\.横ページ = 1;/.test(book) && /ページ設定\.縦ページ = 1;/.test(book));
ok('★10〜400 の 外は 断る★', /倍率は 10％から 400％までです/.test(book));
ok('★ページに 合わせる を 選ぶと 倍率は 使わない★', /ページ設定\.ページに合わせる \? null : ページ設定\.印刷倍率/.test(book));
{
  const f = new Function('notify', 'ページ設定', '拡大縮小の初期化',
    抜く('印刷倍率を足す') + '\nreturn 印刷倍率を足す;');
  const ps = { 印刷倍率: 100, ページに合わせる: true };
  const 足す = f(() => {}, ps, () => ps);
  ok('★増やすと 5％ ずつ★', 足す(5) === 105, String(ps.印刷倍率));
  ok('★増やしたら「ページに 合わせる」は やめる★', ps.ページに合わせる === false);
  ps.印刷倍率 = 398;
  ok('★400 で 止まる★', 足す(5) === 400, String(ps.印刷倍率));
  ps.印刷倍率 = 12;
  ok('★10 で 止まる★', 足す(-5) === 10, String(ps.印刷倍率));
}
{
  /* ★刷る 側に 本当に 効くか★ */
  const 中身 = GP.buildHtml({ data: { '0,0': { v: 'あ', d: 'あ' } }, sheetName: 'S', 印刷倍率: 50 });
  ok('★倍率 50 で 紙の CSS に 出る★', /zoom: 0\.5/.test(中身 || ''),
    (String(中身).match(/zoom:[^;]*/) || [''])[0]);
  const 百 = GP.buildHtml({ data: { '0,0': { v: 'あ', d: 'あ' } }, sheetName: 'S', 印刷倍率: 100 });
  ok('★100 の 時は 何も 足さない★', !/zoom:/.test(百 || ''));
  const 無 = GP.buildHtml({ data: { '0,0': { v: 'あ', d: 'あ' } }, sheetName: 'S' });
  ok('★渡さなければ 100 と 同じ★', !/zoom:/.test(無 || ''));
  const 外 = GP.buildHtml({ data: { '0,0': { v: 'あ', d: 'あ' } }, sheetName: 'S', 印刷倍率: 1000 });
  ok('★外れた 数は 400 で 止める★', /zoom: 4/.test(外 || ''), (String(外).match(/zoom:[^;]*/) || [''])[0]);
}
ok('★alert / prompt / confirm を 使っていない★', !/\balert\(|\bprompt\(|\bconfirm\(/.test(book));

console.log('\n[⑤ ★使い回す 窓の 副題★＝前の 用の 字が 残らないか]');
{
  /* 見つけた 事故（08-30 実ブラウザ）… 配置・回転・グループの 窓に
     「押すと 2個 選んでいます に 「=関数名（」まで 入ります」と 出ていた。
     ＝副題が ★固定の HTML★で、関数一覧の 字が そのまま 使い回されていた。 */
  ok('副題に 名札が 付いた', /id="funcSub"/.test(book));
  ok('副題を 決める 部品が 1つ 在る', !!抜く('窓の副題'));
  const 行
    = book.split('\n');
  const 抜け = [];
  for (let i = 0; i < 行.length; i++) {
    if (!/getElementById\('funcTitle'\)\.textContent =/.test(行[i])) continue;
    const 近く = 行.slice(i, i + 5).join('\n');
    if (!/窓の副題\(/.test(近く)) 抜け.push((i + 1) + ': ' + 行[i].trim().slice(0, 60));
  }
  ok('★副題を 決めていない 窓は 0個★', 抜け.length === 0, 抜け.join(' / '));
  /* ※ 説明の行（ * で 始まる）は 数えない＝画面に 出ない字 */
  const 数 = book.split(String.fromCharCode(10)).filter(function (l) {
    return /「=関数名\(」まで 入ります/.test(l) && !/^\s*\*/.test(l);
  }).length;
  ok('★「=関数名\(」の 字は 関数一覧だけ★', 数 === 2, 数 + 'か所');
  ok('★配置の 窓は 配置の 事を 言う★', /個を そろえます/.test(book));
  ok('★回転の 窓は 回転の 事を 言う★', /個を 回します/.test(book));
}

console.log('\n[⑤ リボンから 押せる]');
{
  const ACT = require_(path.join(ROOT, 'lib/ribbon-actions.js'));
  const g = globalThis, 前w = g.window;
  const 試す = (ボタン, 呼ぶ名, 期待) => {
    let 受け = 'よばれていない';
    const w = {}; w[呼ぶ名] = function (a) { 受け = (a === undefined ? 'ok' : a); };
    g.window = w;
    ACT[ボタン]();
    g.window = 前w;
    ok('「' + ボタン + '」→ ' + 呼ぶ名 + (期待 === 'ok' ? '' : '（' + 期待 + '）'), 受け === 期待, String(受け));
  };
  試す('拡大縮小印刷', '拡大縮小を開く', 'ok');
  試す('横ページ', '拡大縮小を開く', 'ok');
  試す('縦ページ', '拡大縮小を開く', 'ok');
  試す('倍率を増やす', '印刷倍率を足す', 5);
  試す('倍率を減らす', '印刷倍率を足す', -5);
  試す('ページ設定ボタン', 'ページ設定を開く', 'ok');
  試す('配置そろえる', '配置の窓を開く', 'ok');
  試す('グループ化', 'グループの窓を開く', 'ok');
  試す('回転', '回転の窓を開く', 'ok');
}

console.log('\narrange: ' + 緑 + '/' + (緑 + 赤) + ' passed');

if (壊す) {
  console.log('\n★--self-test＝この 見張りが 見ている物を 直に 確かめる（★壊して 赤を 見るのは tools/break-check.mjs★）★');
  let 素通り = 0;
  const 中身 = GP.buildHtml({ data: { '0,0': { v: 'あ', d: 'あ' } }, sheetName: 'S', 印刷倍率: 50 });
  if (!/zoom: 0\.5/.test(中身 || '')) { 素通り++; console.log('  ★素通り★ 倍率が 紙に 効いていない'); }
  else console.log('  ok   倍率が 紙に 効く');
  const 囲み = new Function(抜く('囲みを取る') + '\nreturn 囲みを取る;')();
  const w = 囲み([{ x: 10, y: 10, w: 60, h: 40 }, { x: 100, y: 80, w: 80, h: 20 }]);
  if (w.右 !== 180) { 素通り++; console.log('  ★素通り★ 囲みの 右が 180 で ない … ' + w.右); }
  else console.log('  ok   囲みの 右は 180');
  if (!/2つ以上 選んでください/.test(book)) { 素通り++; console.log('  ★素通り★ 1つでも そろえてしまう'); }
  else console.log('  ok   2つ 無いと 断る');
  if (素通り) { console.log('★抜け道 ' + 素通り + '件★'); process.exit(1); }
  console.log('  ok   抜け道 0件');
}
process.exit(赤 ? 1 : 0);
