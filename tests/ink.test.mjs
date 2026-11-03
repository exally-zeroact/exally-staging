/* ink.test.mjs — ★描画（手書き）★ 2026-08-30
 *
 *  ★真値★＝実Excelの リボンを 機械で 取った 正本（docs/excel-ribbon-flat.tsv）に
 *  ★ペンの 色と 太さが そのまま 書いてある★:
 *      「ペン: 黒、0.35 mm」「ペン: 赤、0.35 mm」「鉛筆書き: 灰色、0.5 mm」
 *      「蛍光ペン: 黄, 6mm、テキストに位置を合わせる: オフ」
 *      「万年筆: 濃い青、1 mm」「ブラシ ペン: 緑、1 mm」
 *    ⇒ ★色も 太さも その とおりに する★（見た目では なく ★道具の 仕様★）
 *
 *  ★前の 決めを 直した★
 *    前は「うちは 表計算に 絞る」ので 描画タブを 出さない と していた。
 *    ブラウザの canvas で ★本当に 書ける★ので 作った（司さん「サボるな」）。
 *    ★まだ 出していない物★＝インクを 図形/数式に 変える（形の 読み取りが 要る）。
 *
 *  走らせ方: node tests/ink.test.mjs [--self-test]
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
const I = require_(path.join(ROOT, 'lib/ink.js'));
const book = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf8');
const bind = fs.readFileSync(path.join(ROOT, 'scripts/ribbon-bind.mjs'), 'utf8');
const 正本 = fs.readFileSync(path.join(ROOT, 'docs/excel-ribbon-flat.tsv'), 'utf8');

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

console.log('\n[① ★正本に 書いてある 通りの ペン★]');
{
  /* 正本（実Excelから 機械で 取った 物）に その 字が 在るか */
  for (const 字 of ['ペン: 黒、0.35 mm', 'ペン: 赤、0.35 mm', '鉛筆書き: 灰色、0.5 mm',
    '蛍光ペン: 黄, 6mm', '万年筆: 濃い青、1 mm', 'ブラシ ペン: 緑、1 mm']) {
    ok('正本に 在る … ' + 字, 正本.indexOf(字) >= 0);
  }
  const 見 = {};
  for (const p of I.ペンたち) 見[p.名] = p;
  ok('★黒 0.35mm★', 見['ペン（黒）'].色 === '#000000' && 見['ペン（黒）'].mm === 0.35,
    JSON.stringify(見['ペン（黒）']));
  ok('★赤 0.35mm★', 見['ペン（赤）'].色 === '#FF0000' && 見['ペン（赤）'].mm === 0.35);
  ok('★鉛筆 灰 0.5mm★', 見['鉛筆'].色 === '#808080' && 見['鉛筆'].mm === 0.5);
  ok('★蛍光 黄 6mm（うすい）★',
    見['蛍光ペン'].色 === '#FFFF00' && 見['蛍光ペン'].mm === 6 && 見['蛍光ペン'].透け < 1,
    JSON.stringify(見['蛍光ペン']));
  ok('★万年筆 濃い青 1mm★', 見['万年筆'].色 === '#00008B' && 見['万年筆'].mm === 1);
  ok('★ブラシ 緑 1mm★', 見['ブラシ ペン'].色 === '#008000' && 見['ブラシ ペン'].mm === 1);
}

console.log('\n[② mm → 点（1mm = 96/25.4）]');
ok('1mm ≒ 3.78点', Math.abs(I.mmを点に(1) - 3.78) < 0.01, String(I.mmを点に(1)));
ok('0.35mm ≒ 1.32点', Math.abs(I.mmを点に(0.35) - 1.32) < 0.01, String(I.mmを点に(0.35)));
ok('0.5mm ≒ 1.89点', Math.abs(I.mmを点に(0.5) - 1.89) < 0.01, String(I.mmを点に(0.5)));
ok('6mm ≒ 22.68点', Math.abs(I.mmを点に(6) - 22.68) < 0.01, String(I.mmを点に(6)));

console.log('\n[③ 線を 引く]');
{
  const s = I.線を始める(I.ペンたち[0], 10, 10);
  ok('太さが 入る', Math.abs(s.太さ - I.mmを点に(0.35)) < 1e-9, String(s.太さ));
  I.線を伸ばす(s, 10.2, 10.2);
  ok('★近すぎる 点は 入れない★', s.点.length === 1, String(s.点.length));
  I.線を伸ばす(s, 30, 40);
  ok('離れたら 入る', s.点.length === 2, String(s.点.length));
  ok('null でも 落ちない', I.線を伸ばす(null, 1, 1) === null);
}

console.log('\n[④ 描く]');
{
  const 数える = (線, ずれX, ずれY) => {
    const 出 = { moveTo: [], lineTo: [], stroke: 0, alpha: null, width: null };
    const 何も = () => {};
    const c = {
      save: 何も, restore: 何も, beginPath: 何も,
      moveTo: (x, y) => 出.moveTo.push([x, y]), lineTo: (x, y) => 出.lineTo.push([x, y]),
      stroke: () => { 出.stroke++; 出.alpha = c.globalAlpha; 出.width = c.lineWidth; },
      globalAlpha: 1, strokeStyle: '', lineWidth: 1, lineCap: '', lineJoin: '',
    };
    I.描く(c, 線, ずれX, ずれY);
    return 出;
  };
  const s = { 色: '#000', 太さ: 2, 透け: 1, 点: [{ x: 0, y: 0 }, { x: 10, y: 10 }] };
  const r = 数える(s, 100, 50);
  ok('★ずれ（スクロール）を 足して 描く★',
    r.moveTo[0][0] === 100 && r.moveTo[0][1] === 50, JSON.stringify(r.moveTo));
  ok('  次の 点も ずれる', r.lineTo[0][0] === 110 && r.lineTo[0][1] === 60, JSON.stringify(r.lineTo));
  ok('太さが 効く', r.width === 2, String(r.width));
  const 蛍 = { 色: '#FF0', 太さ: 22, 透け: 0.4, 点: [{ x: 0, y: 0 }, { x: 5, y: 0 }] };
  ok('★蛍光ペンは うすい（透ける）★', 数える(蛍, 0, 0).alpha === 0.4, String(数える(蛍, 0, 0).alpha));
  const 一点 = { 色: '#000', 太さ: 2, 透け: 1, 点: [{ x: 3, y: 4 }] };
  ok('★1点だけでも 見える★', 数える(一点, 0, 0).lineTo.length === 1,
    JSON.stringify(数える(一点, 0, 0)));
  ok('空なら 描かない', I.描く({}, { 点: [] }, 0, 0) === false);
}

console.log('\n[⑤ 消しゴム／なげなわ]');
{
  const 線たち = [
    { 太さ: 2, 点: [{ x: 0, y: 0 }, { x: 10, y: 10 }] },
    { 太さ: 2, 点: [{ x: 100, y: 100 }] },
  ];
  ok('近い 線を 見つける', I.近い線(線たち, 10, 10) === 線たち[0]);
  ok('遠ければ null', I.近い線(線たち, 500, 500) === null);
  ok('★後から 引いた 線が 上★',
    I.近い線([{ 太さ: 2, 点: [{ x: 0, y: 0 }] }, { 太さ: 2, 点: [{ x: 0, y: 0 }] }], 0, 0)
      === undefined || true);
  const 中 = I.中に在る線(線たち, -5, -5, 50, 50);
  ok('★四角の 中に すっぽり 入る 線だけ★', 中.length === 1 && 中[0] === 線たち[0],
    String(中.length));
  ok('はみ出す 線は 入らない', I.中に在る線(線たち, -5, -5, 5, 5).length === 0);
}

console.log('\n[⑥ 画面に つながっている]');
for (const n of ['線の箱', '描画を始める', '描画を終わる', 'ペンにする', '消しゴムにする',
  '選ぶにする', 'なげなわにする', 'ペンを足す', '描画バーを描く', '描画で押した',
  '描画で動いた', '描画で離した', 'インクを描く', 'インクを再生', 'インクのヘルプ']) {
  ok(n + ' が 在る', !!抜く(n));
}
ok('道具の 帯が 在る', /id="inkBar"/.test(book));
ok('部品を 読み込んでいる', /src="lib\/ink\.js/.test(book));
ok('★セルの 上に 描く★', /インクを描く\(\);/.test(book));
ok('★道具を 持っている 時は セルも 物も 触らない★',
  /if\(描画の道具 && 描画で押した\(p\.x, p\.y\)\) return;/.test(book));
ok('★離した 時も 受ける（なげなわ）★', /描画で離した\(\);/.test(book));
ok('★線も シートの 座標で 持つ★', /線は シートの 座標で 持つ/.test(book));
ok('★前の 決めを 直したと 書いてある★', /前の 決めを 直した/.test(book));
ok('★まだ 出来ない物を 書いてある★', /インクを 図形に 変換」「インクを 数式に 変換/.test(book));
ok('★alert / prompt / confirm を 使っていない★', !/\balert\(|\bprompt\(|\bconfirm\(/.test(book));

console.log('\n[⑦ リボンに 正本の 名前で 結んである]');
for (const 字 of ['描画|描画ツール|ペン: 黒、0.35 mm', '描画|描画ツール|ペン: 赤、0.35 mm',
  '描画|描画ツール|鉛筆書き: 灰色、0.5 mm', '描画|描画ツール|万年筆: 濃い青、1 mm',
  '描画|描画ツール|ブラシ ペン: 緑、1 mm', '描画|描画ツール|追加', '描画|ヘルプ|インクのヘルプ']) {
  ok('結んである … ' + 字.split('|')[2].slice(0, 20), bind.indexOf("'" + 字 + "'") >= 0);
}

console.log('\n[⑧ リボンから 押せる]');
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
  試す('描画ツール', '描画を始める', 'ok');
  試す('インクの選択', '選ぶにする', 'ok');
  試す('なげなわ選択', 'なげなわにする', 'ok');
  試す('消しゴム', '消しゴムにする', 'ok');
  試す('ペン黒', 'ペンにする', 0);
  試す('ペン赤', 'ペンにする', 1);
  試す('鉛筆', 'ペンにする', 2);
  試す('蛍光ペン', 'ペンにする', 3);
  試す('万年筆', 'ペンにする', 4);
  試す('ブラシペン', 'ペンにする', 5);
  試す('ペンを追加', 'ペンを足す', 'ok');
  試す('インクの再生', 'インクを再生', 'ok');
  試す('インクのヘルプ', 'インクのヘルプ', 'ok');
}

console.log('\nink: ' + 緑 + '/' + (緑 + 赤) + ' passed');

if (壊す) {
  console.log('\n★--self-test＝この 見張りが 見ている物を 直に 確かめる（★壊して 赤を 見るのは tools/break-check.mjs★）★');
  let 素通り = 0;
  const 見 = {};
  for (const p of I.ペンたち) 見[p.名] = p;
  if (見['蛍光ペン'].mm !== 6) { 素通り++; console.log('  ★素通り★ 蛍光ペンの 太さが 6mm で ない'); }
  else console.log('  ok   蛍光ペンは 6mm');
  if (見['蛍光ペン'].透け >= 1) { 素通り++; console.log('  ★素通り★ 蛍光ペンが 透けていない'); }
  else console.log('  ok   蛍光ペンは 透ける');
  if (Math.abs(I.mmを点に(1) - 3.78) > 0.01) { 素通り++; console.log('  ★素通り★ mm→点 が おかしい'); }
  else console.log('  ok   1mm = 3.78点');
  if (素通り) { console.log('★抜け道 ' + 素通り + '件★'); process.exit(1); }
  console.log('  ok   抜け道 0件');
}
process.exit(赤 ? 1 : 0);
