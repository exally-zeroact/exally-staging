/* sparkline.test.mjs — ★スパークライン（セルの中の 小さな グラフ）★ 2026-08-30
 *
 *  ★真値（実Excel 16.0 で 実測）★ … tools/measure-sparkline.ps1 / tools/measure-spark2.ps1
 *    種類＝折れ線1・縦棒2・勝敗3。どれも:
 *      線の太さ 0.75 ／ 高い点・低い点・最初・最後・印 は ★全部 False（飾りは 出さない）★
 *      ★勝敗だけ マイナス=True★
 *      縦軸の型 ★2★（スパークラインごと＝画面の「自動（各スパークライン）」）
 *      横軸は 出さない ／ 空の扱い ★1（描かない＝隙間）★
 *
 *  走らせ方: node tests/sparkline.test.mjs [--self-test]
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

const S = require_(path.join(ROOT, 'lib/sparkline.js'));
const book = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf8');

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

/* 描いた 命令を 覚える 偽の ctx（jsdom も canvas も 要らない） */
function 偽ctx() {
  const 出 = { 四角: [], 線: [], 色: [], 太さ: null, moveTo: 0, lineTo: 0, stroke: 0 };
  const 何も = () => {};
  const c = {
    _出: 出,
    beginPath: 何も, closePath: 何も,
    moveTo: (x, y) => { 出.moveTo++; 出.線.push(['M', Math.round(x), Math.round(y)]); },
    lineTo: (x, y) => { 出.lineTo++; 出.線.push(['L', Math.round(x), Math.round(y)]); },
    stroke: () => { 出.stroke++; 出.太さ = c.lineWidth; },
    fillRect: (x, y, w, h) => { 出.四角.push({ x: Math.round(x), y: Math.round(y), w: Math.round(w), h: Math.round(h), 色: c.fillStyle }); 出.色.push(c.fillStyle); },
    fillStyle: '', strokeStyle: '', lineWidth: 1,
  };
  return c;
}

console.log('\n[① 実測の 道具が 残っている]');
ok('tools/measure-sparkline.ps1 が 在る', fs.existsSync(path.join(ROOT, 'tools/measure-sparkline.ps1')));
ok('tools/measure-spark2.ps1 が 在る', fs.existsSync(path.join(ROOT, 'tools/measure-spark2.ps1')));
const lib = fs.readFileSync(path.join(ROOT, 'lib/sparkline.js'), 'utf8');
ok('★線の太さ 0.75（実測）★', S.線の太さ === 0.75, String(S.線の太さ));
ok('実測の 表が 部品の 見出しに 在る', /0\.75/.test(lib) && /勝敗/.test(lib));

console.log('\n[② 値の 読み方（空は 隙間＝実測⑤）]');
{
  const v = S.値にする([1, '', 3, null, '2,000', 'あ']);
  ok('空は null に なる', v[1] === null && v[3] === null, JSON.stringify(v));
  ok('けた区切りも 数に なる', v[4] === 2000, String(v[4]));
  ok('字は null（数では ない）', v[5] === null, String(v[5]));
}

console.log('\n[③ 折れ線＝★空の所で 線を 切る（実測⑤）★]');
{
  const c = 偽ctx();
  ok('描けた', S.描く(c, S.値にする([1, 2, 3, 4]), 'line', 0, 0, 40, 20) === true);
  ok('★線の太さは 0.75（実測）★', c._出.太さ === 0.75, String(c._出.太さ));
  ok('4つとも つながる（M1回 + L3回）', c._出.moveTo === 1 && c._出.lineTo === 3,
    'M' + c._出.moveTo + ' L' + c._出.lineTo);
  const c2 = 偽ctx();
  S.描く(c2, S.値にする([1, 2, '', 4, 5]), 'line', 0, 0, 40, 20);
  ok('★真ん中が 空なら 線が 2本に 切れる（M2回）★', c2._出.moveTo === 2,
    'M' + c2._出.moveTo + ' L' + c2._出.lineTo);
  ok('★空の所に 点を 打たない★', c2._出.lineTo === 2, 'L' + c2._出.lineTo);
}

console.log('\n[④ 縦棒＝マイナスは 下・別の色]');
{
  const c = 偽ctx();
  ok('描けた', S.描く(c, S.値にする([10, -5, 30, 20]), 'column', 0, 0, 40, 20) === true);
  ok('4本 出る', c._出.四角.length === 4, String(c._出.四角.length));
  const 赤い = c._出.四角.filter((q) => q.色 === S.色.マイナス);
  ok('★マイナスの1本だけ 別の色★', 赤い.length === 1, JSON.stringify(c._出.色));
  const 上 = c._出.四角[0], 下 = c._出.四角[1];
  ok('★マイナスは 0より 下から 始まる★', 下.y >= (上.y + 上.h) - 1 || 下.y > 上.y,
    JSON.stringify({ 上, 下 }));
  const c2 = 偽ctx();
  S.描く(c2, S.値にする([10, '', 30]), 'column', 0, 0, 40, 20);
  ok('★空の所には 棒を 出さない★', c2._出.四角.length === 2, String(c2._出.四角.length));
}

console.log('\n[⑤ 勝敗＝★0は 描かない・高さは 同じ★]');
{
  const c = 偽ctx();
  ok('描けた', S.描く(c, S.値にする([3, -8, 0, 1]), 'winloss', 0, 0, 40, 20) === true);
  ok('★0は 勝ちでも 負けでも ない＝出さない（3本）★', c._出.四角.length === 3,
    String(c._出.四角.length));
  const 高さたち = c._出.四角.map((q) => q.h);
  ok('★大きさに 関わらず 高さは 同じ★', new Set(高さたち).size === 1, JSON.stringify(高さたち));
  const 赤い = c._出.四角.filter((q) => q.色 === S.色.マイナス);
  ok('★負けだけ 別の色（1本）★', 赤い.length === 1, JSON.stringify(c._出.色));
  const 勝ち = c._出.四角.find((q) => q.色 === S.色.ふつう);
  const 負け = 赤い[0];
  ok('★勝ちは 上・負けは 下★', 勝ち.y < 負け.y, 勝ち.y + ' / ' + 負け.y);
}

console.log('\n[⑥ ★縦の目盛は その1本の 中だけ（実測③）★]');
{
  /* 同じ 形でも 数が 違えば 高さの 割合は 同じ に なる＝1本ごとに 目盛を 取っている印 */
  const 小 = 偽ctx(); S.描く(小, S.値にする([1, 2, 3]), 'column', 0, 0, 30, 20);
  const 大 = 偽ctx(); S.描く(大, S.値にする([100, 200, 300]), 'column', 0, 0, 30, 20);
  ok('1・2・3 と 100・200・300 が 同じ 絵に なる',
    JSON.stringify(小._出.四角.map((q) => q.h)) === JSON.stringify(大._出.四角.map((q) => q.h)),
    JSON.stringify(小._出.四角.map((q) => q.h)) + ' / ' + JSON.stringify(大._出.四角.map((q) => q.h)));
}

console.log('\n[⑦ 出来ない時は ★出来たふりを しない★]');
ok('数が 1つも 無い → false', S.描く(偽ctx(), S.値にする(['', 'あ']), 'line', 0, 0, 40, 20) === false);
ok('空の 並び → false', S.描く(偽ctx(), [], 'line', 0, 0, 40, 20) === false);
ok('セルが 小さすぎる → false', S.描く(偽ctx(), S.値にする([1, 2]), 'line', 0, 0, 3, 2) === false);

console.log('\n[⑧ 画面に つながっている]');
for (const n of ['スパークラインの箱', 'スパークラインの窓を開く', 'スパークラインの元を当てる',
  'スパークラインの元を読む', 'スパークラインを決める', 'スパークラインを消す', 'スパークラインを描く']) {
  ok(n + ' が 在る', !!抜く(n));
}
ok('窓が 在る', /id="sparkOverlay"/.test(book));
ok('部品を 読み込んでいる', /src="lib\/sparkline\.js/.test(book));
ok('★字より 先に 描く（空セルでも 出す）★',
  book.indexOf('スパークラインを描く(r, c, x, y, w, h)') < book.indexOf('if(!raw) return;'),
  '順番');
ok('★alert / prompt / confirm を 使っていない★', !/\balert\(|\bprompt\(|\bconfirm\(/.test(book));

console.log('\n[⑨ 範囲の 読み方]');
{
  const f = new Function('sheets', 'activeSheet', 'addrToRC',
    抜く('スパークラインの元を読む') + '\nreturn スパークラインの元を読む;');
  const データ = { '0,0': { v: 1 }, '0,1': { v: 2 }, '0,2': { v: 3 }, '1,0': { v: 9 } };
  const 読む = f([{ data: データ }], 0, (a) => {
    const m = a.match(/^([A-Z]+)(\d+)$/i);
    if (!m) return null;
    let col = 0;
    for (let i = 0; i < m[1].length; i++) col = col * 26 + (m[1].toUpperCase().charCodeAt(i) - 64);
    return { r: parseInt(m[2]) - 1, c: col - 1 };
  });
  ok('1行なら 読める', JSON.stringify(読む('A1:C1')) === '[1,2,3]', JSON.stringify(読む('A1:C1')));
  ok('1列なら 読める', JSON.stringify(読む('A1:A2')) === '[1,9]', JSON.stringify(読む('A1:A2')));
  ok('★四角い 範囲は 断る（1行 か 1列 だけ）★', 読む('A1:B2') === null, JSON.stringify(読む('A1:B2')));
  ok('形が 違えば 断る', 読む('あいう') === null);
  ok('小文字でも 読める', JSON.stringify(読む('a1:c1')) === '[1,2,3]', JSON.stringify(読む('a1:c1')));
}

console.log('\n[⑨-2 ★元の 範囲を 当てる（空欄を 人に 埋めさせない）★]');
{
  const f = new Function('sheets', 'activeSheet', 'selR1', 'selC1', 'Sparkline', 'cellAddr',
    抜く('スパークラインの元を当てる') + '\nreturn スパークラインの元を当てる;');
  const 名 = (r, c) => String.fromCharCode(65 + c) + (r + 1);
  const 当てる = (データ, r, c) => f([{ data: データ }], 0, r, c, S, 名)();
  const データ = { '0,0': { v: 10 }, '0,1': { v: -5 }, '0,2': { v: 30 }, '0,3': { v: 20 } };
  ok('すぐ 左に 数が 並んでいれば 当てる', 当てる(データ, 0, 4) === 'A1:D1', 当てる(データ, 0, 4));
  ok('★すぐ左が 空でも 当てる（E が 空で F に 置く）★',
    当てる(データ, 0, 5) === 'A1:D1', 当てる(データ, 0, 5));
  ok('数が 1つも 無ければ 空で 返す（★でたらめを 入れない★）', 当てる({}, 0, 5) === '', 当てる({}, 0, 5));
  ok('一番 左に 居る時は 空', 当てる(データ, 0, 0) === '', 当てる(データ, 0, 0));
  const 飛び = { '0,0': { v: 1 }, '0,3': { v: 9 } };
  ok('★間が 空いていたら 続きだけ 取る★', 当てる(飛び, 0, 5) === 'D1:D1', 当てる(飛び, 0, 5));
}

console.log('\n[⑩ リボンから 押せる]');
{
  const ACT = require_(path.join(ROOT, 'lib/ribbon-actions.js'));
  const g = globalThis, 前w = g.window;
  for (const [ボタン, 期待] of [
    ['スパークライン折れ線', 'line'], ['スパークライン縦棒', 'column'], ['スパークライン勝敗', 'winloss'],
  ]) {
    let 受け = null;
    g.window = { スパークラインの窓を開く: function (k) { 受け = k; } };
    ACT[ボタン]();
    g.window = 前w;
    ok('「' + ボタン + '」→ ' + 期待, 受け === 期待, String(受け));
  }
}

console.log('\nsparkline: ' + 緑 + '/' + (緑 + 赤) + ' passed');

if (壊す) {
  console.log('\n★--self-test＝この 見張りが 見ている物を 直に 確かめる（★壊して 赤を 見るのは tools/break-check.mjs★）★');
  let 素通り = 0;
  /* 壊し① 空を 0 として 描いたら（実測は「描かない」）気づけるか */
  const c = 偽ctx();
  S.描く(c, [1, 2, null, 4, 5], 'line', 0, 0, 40, 20);
  if (c._出.moveTo === 1) { 素通り++; console.log('  ★素通り★ 空でも 線が つながっている＝隙間に なっていない'); }
  else console.log('  ok   空で 線が 切れている（M' + c._出.moveTo + '）');
  /* 壊し② 勝敗で 0 を 描いてしまう */
  const c2 = 偽ctx();
  S.描く(c2, [0, 0, 0], 'winloss', 0, 0, 40, 20);
  if (c2._出.四角.length > 0) { 素通り++; console.log('  ★素通り★ 0だけなのに 棒が 出た'); }
  else console.log('  ok   0だけなら 棒は 出ない');
  /* 壊し③ 全部 同じ 色に なっていないか（マイナスの 別色が 消えていないか） */
  const c3 = 偽ctx();
  S.描く(c3, [5, -5], 'winloss', 0, 0, 40, 20);
  if (new Set(c3._出.色).size < 2) { 素通り++; console.log('  ★素通り★ 勝ちと 負けが 同じ色'); }
  else console.log('  ok   勝ちと 負けは 別の色');
  if (素通り) { console.log('★抜け道 ' + 素通り + '件★'); process.exit(1); }
  console.log('  ok   抜け道 0件');
}
process.exit(赤 ? 1 : 0);
