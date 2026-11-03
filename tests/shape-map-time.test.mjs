/* shape-map-time.test.mjs — ★手書き→図形／マップ／タイムライン★ 2026-08-30
 *
 *  ★実Excel で 測れた 事★ … tools/measure-insert.ps1 ／ measure-review2.ps1
 *    インク … ★形（Shapes）の 1つとして 数える★／★変換の 中の 動きは COM から 見えない★
 *    マップ … ★Microsoft の 地図の サービス★（境目の 形は 向こうの 物）
 *    タイムライン … ★日付の 列を 持つ ピボットが 要る★／スライサーの 仲間／
 *      ★形・大きさは COM から 読めなかった★
 *
 *  ★実Excel の 日付の 通し番号（実測 2026-08-30）★
 *    1→1900-01-01／2→1900-01-02／59→1900-02-28／
 *    ★60→1900-02-29（実際には 無い 日）★／61→1900-03-01／
 *    45900→2025-08-31／46000→2025-12-09
 *
 *  ★真似る 相手の 数字が 無い 所は 自分で 決めて 自分で 測る★
 *    手書きの 形あて … 12通りの 線で ★12/12★
 *
 *  走らせ方: node tests/shape-map-time.test.mjs [--self-test]
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
const S = require_(path.join(ROOT, 'lib/ink-shape.js'));
const M = require_(path.join(ROOT, 'lib/jp-map.js'));
const T = require_(path.join(ROOT, 'lib/timeline.js'));
const Th = require_(path.join(ROOT, 'lib/theme.js'));

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

/* ── 試す 線を 作る（★手の ふるえも 足す★） ── */
function ゆらす(点, 幅) {
  let 種 = 12345;
  const r = () => { 種 = (種 * 1103515245 + 12345) % 2147483648; return 種 / 2147483648 - 0.5; };
  return 点.map(p => ({ x: p.x + r() * 幅, y: p.y + r() * 幅 }));
}
const 円 = (cx, cy, R, n = 60) => Array.from({ length: n + 1 }, (_, i) => {
  const a = i * 2 * Math.PI / n;
  return { x: cx + Math.cos(a) * R, y: cy + Math.sin(a) * R };
});
const 多角 = (点, 一辺 = 12) => {
  const 出 = [];
  for (let i = 0; i < 点.length; i++) {
    const a = 点[i], b = 点[(i + 1) % 点.length];
    for (let t = 0; t < 一辺; t++) 出.push({ x: a.x + (b.x - a.x) * t / 一辺, y: a.y + (b.y - a.y) * t / 一辺 });
  }
  出.push({ x: 点[0].x, y: 点[0].y });
  return 出;
};
const 直線 = (x1, y1, x2, y2, n = 30) => Array.from({ length: n + 1 }, (_, i) => ({
  x: x1 + (x2 - x1) * i / n, y: y1 + (y2 - y1) * i / n,
}));
const 矢印 = () => [
  ...直線(0, 50, 120, 50, 30), ...直線(120, 50, 100, 38, 8),
  ...直線(100, 38, 120, 50, 8), ...直線(120, 50, 100, 62, 8),
];
const 四角点 = [{ x: 0, y: 0 }, { x: 100, y: 0 }, { x: 100, y: 60 }, { x: 0, y: 60 }];
const 三角点 = [{ x: 50, y: 0 }, { x: 100, y: 80 }, { x: 0, y: 80 }];

console.log('\n[① 手書きの 形あて（★12通り★）]');
{
  const 例 = [
    ['丸（きれい）', 円(50, 50, 40), '丸'],
    ['丸（ふるえ3）', ゆらす(円(50, 50, 40), 3), '丸'],
    ['丸（小さい）', 円(20, 20, 12), '丸'],
    ['四角', 多角(四角点), '四角'],
    ['四角（ふるえ3）', ゆらす(多角(四角点), 3), '四角'],
    ['三角', 多角(三角点), '三角'],
    ['三角（ふるえ3）', ゆらす(多角(三角点), 3), '三角'],
    ['線（横）', 直線(0, 30, 150, 30), '線'],
    ['線（ななめ）', 直線(0, 0, 120, 90), '線'],
    ['線（ふるえ3）', ゆらす(直線(0, 30, 150, 30), 3), '線'],
    ['矢印', 矢印(), '右矢印'],
    ['矢印（ふるえ2）', ゆらす(矢印(), 2), '右矢印'],
  ];
  let 合 = 0;
  for (const [名, 点, 期] of 例) {
    const r = S.形を決める(点);
    if (r.形 === 期) 合++;
    ok('  ' + 名 + ' → ' + 期, r.形 === 期,
      r.形 + '（角=' + r.角 + ' 丸さ=' + r.丸さ.toFixed(2) + ' 閉じ=' + r.閉じ + '）');
  }
  ok('★12通り 全部 当たる★', 合 === 12, 合 + '/12');
}
console.log('  ─ 数え方そのもの ─');
ok('★丸さは 真円で 1.0★', Math.abs(S.丸さ(円(50, 50, 40)) - 1) < 0.01, S.丸さ(円(50, 50, 40)).toFixed(3));
ok('★四角の 角は 4個★', S.角たち(S.まびく(多角(四角点), 4), true).length === 4,
  String(S.角たち(S.まびく(多角(四角点), 4), true).length));
ok('★三角の 角は 3個★', S.角たち(S.まびく(多角(三角点), 4), true).length === 3,
  String(S.角たち(S.まびく(多角(三角点), 4), true).length));
ok('★丸の 角は 0個★', S.角たち(S.まびく(円(50, 50, 40), 4), true).length === 0);
ok('★閉じた 線は 端でも 回り込んで 見る（回り込まないと 1個 減る）★',
  /★閉じた 線は 端でも 回り込んで 見る★/.test(fs.readFileSync(path.join(ROOT, 'lib/ink-shape.js'), 'utf8')));
ok('★同じ 角を 2回 数えない（端と 端）★',
  /★閉じた 線は 端と 端も 隣どうし★/.test(fs.readFileSync(path.join(ROOT, 'lib/ink-shape.js'), 'utf8')));
ok('  閉じているか（丸）', S.閉じているか(円(50, 50, 40)) === true);
ok('  閉じているか（線）', S.閉じているか(直線(0, 0, 100, 0)) === false);
ok('  囲みが 合う', (() => { const w = S.囲み(多角(四角点)); return w.w === 100 && w.h === 60; })());
ok('  点が 少なくても 落ちない', S.形を決める([]).形 === '線' && S.形を決める([{ x: 1, y: 1 }]).形 === '線');

console.log('\n[② マップ（升目の 日本地図）]');
ok('★47都道府県★', M.県たち.length === 47, String(M.県たち.length));
ok('★名前が だぶっていない★', new Set(M.県たち.map(v => v.名)).size === 47);
ok('★升目が だぶっていない★', new Set(M.県たち.map(v => v.行 + ',' + v.列)).size === 47);
ok('★北海道は 一番 上★', M.県を探す('北海道').行 === 0);
ok('★沖縄は 一番 下★',
  M.県を探す('沖縄').行 === Math.max(...M.県たち.map(v => v.行)));
ok('★北海道は 沖縄より 右★', M.県を探す('北海道').列 > M.県を探す('沖縄').列);
ok('★東京は 大阪より 右★', M.県を探す('東京').列 > M.県を探す('大阪').列);
ok('★青森は 鹿児島より 上★', M.県を探す('青森').行 < M.県を探す('鹿児島').行);
ok('  「都」「府」「県」が 付いても 分かる',
  M.県を探す('東京都').名 === '東京' && M.県を探す('大阪府').名 === '大阪'
  && M.県を探す('神奈川県').名 === '神奈川');
ok('  北海道は そのまま', M.県を探す('北海道').名 === '北海道');
ok('  無い 名前は null', M.県を探す('ないけん') === null && M.県を探す('') === null);
{
  const r = M.絵にする([['東京都', 100], ['大阪府', 60], ['北海道', 20], ['あやしい', 5]],
    '#3D9E72', Th.濃淡);
  ok('★升は いつも 47個★', r.升の数 === 47, String(r.升の数));
  ok('★値が 在るのは 3つ★', r.升たち.filter(v => v.あるか).length === 3);
  ok('★最小 20・最大 100★', r.最小 === 20 && r.最大 === 100);
  ok('★県で ない 字は 数えて 出す★', r.見つからない.length === 1 && r.見つからない[0] === 'あやしい');
  ok('★大きいほど 濃い★', (() => {
    const a = r.升たち.filter(v => v.あるか).sort((x, y) => x.値 - y.値);
    const 明 = (c) => parseInt(c.slice(1, 3), 16) + parseInt(c.slice(3, 5), 16) + parseInt(c.slice(5, 7), 16);
    for (let i = 1; i < a.length; i++) if (明(a[i].色) >= 明(a[i - 1].色)) return false;
    return true;
  })(), r.升たち.filter(v => v.あるか).map(v => v.名 + '=' + v.色).join(' '));
  ok('  値の 無い 県は 色 なし', r.升たち.filter(v => !v.あるか).every(v => v.色 === null));
  ok('  数に ならない 物は 数えない', M.絵にする([['東京', 'あ']], '#3D9E72', Th.濃淡)
    .升たち.filter(v => v.あるか).length === 0);
  ok('  空でも 落ちない', M.絵にする([], '#3D9E72', Th.濃淡).升の数 === 47);
}

console.log('\n[③ タイムライン（★Excel の 通し番号は 実測と 7/7★）]');
{
  const 実測 = { 1: '1900-01-01', 2: '1900-01-02', 59: '1900-02-28', 60: null,
    61: '1900-03-01', 45900: '2025-08-31', 46000: '2025-12-09' };
  let 合 = 0;
  for (const n of Object.keys(実測)) {
    const d = T.日付にする(String(n), 'yyyy-mm-dd');
    const 出 = d ? d.toISOString().slice(0, 10) : null;
    const 同 = 出 === 実測[n];
    if (同) 合++;
    ok('  通し番号 ' + n + ' → ' + (実測[n] || '読めない'), 同, String(出));
  }
  ok('★7/7 実Excel と 同じ★', 合 === 7, 合 + '/7');
  ok('★60 は 1900-02-29＝無い 日なので 読めないと 言う★', T.日付にする('60', 'yyyy-mm-dd') === null);
}
console.log('  ─ ★ただの 数を 日付に しない★（08-30 実ブラウザで 見つけた）─');
ok('★書式が 無い 100 は 日付に しない★', T.日付にする('100') === null);
ok('  書式が 日付なら 通し番号に する', !!T.日付にする('100', 'yyyy-mm-dd'));
ok('★通貨の 書式は 日付に しない★', T.日付にする('45900', '¥#,##0') === null);
ok('★％の 書式も 日付に しない★', T.日付にする('45900', '0.0%') === null);
ok('  日付の 書式か を 見分ける',
  T.日付の書式か('yyyy-mm-dd') === true && T.日付の書式か('#,##0') === false
  && T.日付の書式か('') === false);
console.log('  ─ 字の 日付 ─');
for (const [s, e] of [['2026-08-30', '2026-08-30'], ['2026/8/30', '2026-08-30'],
  ['2026年8月30日', '2026-08-30'], ['2026-08', '2026-08-01']]) {
  const d = T.日付にする(s);
  ok('  ' + s + ' → ' + e, d && d.toISOString().slice(0, 10) === e, d && d.toISOString().slice(0, 10));
}
ok('  読めない 字は null', T.日付にする('あいう') === null && T.日付にする('') === null);
{
  const 値 = [{ 値: '2026-01-15', 行: 0 }, { 値: '2026-02-20', 行: 1 },
    { 値: '2026-05-01', 行: 2 }, { 値: '2025-12-31', 行: 3 }, { 値: 'よめない', 行: 4 }];
  ok('★区切りは 4つ（年・四半期・月・日）★', T.区切りたち.length === 4);
  const 年 = T.目盛りを作る(値, '年');
  ok('★年で まとめると 2個★', 年.目盛り.length === 2, JSON.stringify(年.目盛り.map(v => v.名)));
  ok('★読めない 物を 黙って 捨てない★', 年.読めない === 1 && 年.読めた === 4);
  ok('  古い順に 並ぶ', 年.目盛り[0].名 === '2025年' && 年.目盛り[1].名 === '2026年');
  ok('  四半期は 3個', T.目盛りを作る(値, '四半期').目盛り.length === 3);
  ok('  月は 4個', T.目盛りを作る(値, '月').目盛り.length === 4);
  ok('  日も 4個', T.目盛りを作る(値, '日').目盛り.length === 4);
  const 月 = T.目盛りを作る(値, '月');
  ok('★1つ 選ぶと その 行だけ★', JSON.stringify(T.選んだ行(月.目盛り, 1, 1)) === '{"0":true}',
    JSON.stringify(T.選んだ行(月.目盛り, 1, 1)));
  ok('★2つ 選ぶと 間も 入る★',
    Object.keys(T.選んだ行(月.目盛り, 0, 2)).length === 3,
    JSON.stringify(T.選んだ行(月.目盛り, 0, 2)));
  ok('★逆に 選んでも 同じ★',
    JSON.stringify(T.選んだ行(月.目盛り, 2, 0)) === JSON.stringify(T.選んだ行(月.目盛り, 0, 2)));
  ok('★null は ぜんぶ 見せる★', T.選んだ行(月.目盛り, null) === null);
}

console.log('\n[④ 画面から 押せる]');
for (const n of ['インクを図形に', 'インクを数式に', 'アクションペン',
  'マップを作る', '地図の窓を開く',
  'タイムラインを作る', 'タイムラインを組む', 'タイムラインの窓を開く',
  'タイムラインの区切り', 'タイムラインを押す', 'タイムラインを外す']) {
  ok(n + ' が 在る', !!抜く(n));
}
ok('★図形に したら 元の 線は 消す★', /箱\.splice\(j, 1\);/.test(抜く('インクを図形に') || ''));
ok('★線が 無ければ 断る★', /手書きの 線が ありません/.test(book));
ok('★字を 読む 仕組みが 無い事を 画面に 書いてある★',
  /手書きの「字」を 読むには 学習した 仕組み/.test(book));
ok('★アクション ペンも 正直に お知らせだけ★', /出来ない物を 出来るように 見せない/.test(book));
ok('★地図は 升目だと 画面に 書いてある★', /うちは 升目の 地図/.test(book));
ok('★境目の 形が 向こうの 物だと 書いてある★', /境目の 形は 向こうの 物なので うちには ありません/.test(book));
ok('★県で ない 字を 画面に 出す★', /県の 名前では ない/.test(book));
ok('★日付の 列は 書式も 見て 決める★', /セル && セル\.numFmt/.test(抜く('タイムラインを作る') || ''));
ok('★読めない 数を 画面に 出す★', /読めない ' \+ 時の決め\.読めない \+ '個（黙って 捨てません）/.test(book));
ok('★絞る 道は 見出しの ▼ と 同じ 1本★', /filterHidden/.test(抜く('タイムラインを押す') || ''));
ok('★alert / prompt / confirm を 使っていない★', !/\balert\(|\bprompt\(|\bconfirm\(/.test(book));

console.log('\n[⑤ 副題を 決めていない 窓が 増えていないか]');
{
  const 行 = book.split(String.fromCharCode(10));
  const 抜け = [];
  for (let i = 0; i < 行.length; i++) {
    if (!/getElementById\('funcTitle'\)\.textContent =/.test(行[i])) continue;
    if (!/窓の副題\(/.test(行.slice(i, i + 5).join(String.fromCharCode(10)))) 抜け.push(i + 1);
  }
  ok('★副題を 決めていない 窓は 0個★', 抜け.length === 0, 抜け.join(' / '));
}

console.log('\n[⑥ リボンから 押せる]');
{
  const ACT = require_(path.join(ROOT, 'lib/ribbon-actions.js'));
  const g = globalThis, 前w = g.window;
  const 試す = (ボタン, 呼ぶ名) => {
    let 受け = 'よばれていない';
    const w = {}; w[呼ぶ名] = function () { 受け = 'ok'; };
    g.window = w; ACT[ボタン](); g.window = 前w;
    ok('「' + ボタン + '」→ ' + 呼ぶ名, 受け === 'ok', String(受け));
  };
  試す('インクを図形に', 'インクを図形に');
  試す('インクを数式に', 'インクを数式に');
  試す('アクションペン', 'アクションペン');
  試す('マップ', 'マップを作る');
  試す('タイムライン', 'タイムラインを作る');
}

console.log('\nshape-map-time: ' + 緑 + '/' + (緑 + 赤) + ' passed');

if (壊す) {
  console.log('\n★--self-test＝この 見張りが 見ている物を 直に 確かめる（★壊して 赤を 見るのは tools/break-check.mjs★）★');
  let 素通り = 0;
  if (S.形を決める(多角(四角点)).形 !== '四角') { 素通り++; console.log('  ★素通り★ 四角を 当てられない'); }
  else console.log('  ok   四角を 当てる');
  if (S.形を決める(多角(三角点)).形 !== '三角') { 素通り++; console.log('  ★素通り★ 三角を 当てられない'); }
  else console.log('  ok   三角を 当てる');
  if (T.日付にする('100') !== null) { 素通り++; console.log('  ★素通り★ ただの 数を 日付に した'); }
  else console.log('  ok   ただの 数は 日付に しない');
  if (T.日付にする('60', 'yyyy-mm-dd') !== null) { 素通り++; console.log('  ★素通り★ 無い 日（1900-02-29）を 通した'); }
  else console.log('  ok   無い 日は 読めないと 言う');
  if (new Set(M.県たち.map(v => v.行 + ',' + v.列)).size !== 47) {
    素通り++; console.log('  ★素通り★ 升目が だぶっている');
  } else console.log('  ok   升目は だぶっていない');
  if (素通り) { console.log('★抜け道 ' + 素通り + '件★'); process.exit(1); }
  console.log('  ok   抜け道 0件');
}
process.exit(赤 ? 1 : 0);
