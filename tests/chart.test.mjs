/* chart.test.mjs — ★グラフ（縦棒・折れ線・円）★ 2026-08-29
 *
 *  ★真値（実Excel 16.0 で 実測 2026-08-29）★
 *    ・既定の 種類 … ★51 = 集合縦棒★
 *    ・既定の 大きさ … ★360 × 216★
 *    ・凡例 … ★系列が 1本なら 出さない★（HasLegend=False）
 *    ・タイトル … ★出す★（HasTitle=True）
 *
 *  ★見せ方★＝Excelの 絵・配色は ★1つも 写さない★（うちの緑・自前の canvas・外の部品 0）
 *
 *  ★作っただけで 緑に しない★＝★実際に 描かせて 何が 描かれたかを 数える★。
 *
 *  走らせ方: node tests/chart.test.mjs [--self-test]
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

const C = require_(path.join(ROOT, 'lib/chart.js'));
const book = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf8');

console.log('\n[① 実Excelの 既定と 同じ]');
ok('★大きさ 360 × 216★', C.既定.幅 === 360 && C.既定.高さ === 216, JSON.stringify(C.既定));
ok('★色は うちの緑★（Excelの 配色を 写していない）',
  C.色[0] === '#3D9E72' && C.色.every((c) => /^#[0-9A-F]{6}$/i.test(c)), C.色.join(','));

console.log('\n[② 材料の 取り出し（見出し・系列名の 見立て）]');
{
  /* ★2026-08-30 実測で 直した★
     前は「系列は いつも 列」と していたが、実Excel は ★数の塊で 行数 > 列数 の時だけ 列★。
     ここは 数が 3行2列＝★行数 > 列数★なので 系列は 列（売上・原価）。
     （実測: 見出し付き 数3行2列 → 系列='売上','原価'・横軸=1月,2月,3月） */
  const m = C.材料を作る([['月', '売上', '原価'], ['1月', 100, 60], ['2月', 150, 80], ['3月', 120, 70]]);
  ok('見出しと 系列名を 見分ける', m.ok && m.見出し.join(',') === '1月,2月,3月' && m.系列.length === 2,
    JSON.stringify(m.見出し) + ' / ' + (m.系列 || []).length);
  ok('系列の 名前を 拾う', m.ok && m.系列[0].名 === '売上' && m.系列[1].名 === '原価',
    m.ok ? m.系列.map((s) => s.名).join(',') : '');
  ok('数を 拾う', m.ok && m.系列[0].値.join('/') === '100/150/120', m.ok ? m.系列[0].値.join('/') : '');
}
{
  /* ★行数 ≦ 列数 なら 系列は 行★（実測: 見出し付き 数2行2列 → 系列='1月','2月'・横軸=売上,原価） */
  const m = C.材料を作る([['月', '売上', '原価'], ['1月', 100, 60], ['2月', 150, 80]]);
  ok('★数2行2列は 横に 取る（実測）★',
    m.ok && m.見出し.join(',') === '売上,原価' && m.系列.map((s) => s.名).join(',') === '1月,2月',
    JSON.stringify(m.見出し) + ' / ' + (m.系列 || []).map((s) => s.名).join(','));
  ok('横に 取った印が 立つ', m.横に取った === true, String(m.横に取った));
}
{
  const m = C.材料を作る([[100], [150], [120]]);
  ok('★数だけなら 見出しは 1,2,3★', m.ok && m.見出し.join(',') === '1,2,3', JSON.stringify(m.見出し));
  ok('★名前が 無ければ 系列1★', m.ok && m.系列[0].名 === '系列1', m.ok ? m.系列[0].名 : '');
}
console.log('\n[③ ★出来ない時は 黙らない★]');
ok('数が 1つも 無い → 理由を 返す', C.材料を作る([['あ', 'い'], ['う', 'え']]).なぜ === '数が 1つも ありません');
ok('空 → 理由を 返す', C.材料を作る([]).なぜ === '選んだ所に 何も ありません');
ok('★ok:false を 返す（黙って 空のグラフを 作らない）★', C.材料を作る([]).ok === false);

console.log('\n[④ ★実際に 描かせて 何が 描かれたかを 数える★]');
function にせctx() {
  const 記録 = { fillRect: 0, stroke: 0, fillText: [], arc: 0, clearRect: 0 };
  return {
    記録,
    canvas: {},
    clearRect() { 記録.clearRect++; },
    fillRect() { 記録.fillRect++; },
    strokeRect() {},
    beginPath() {}, moveTo() {}, lineTo() {}, closePath() {},
    stroke() { 記録.stroke++; },
    fill() {},
    arc() { 記録.arc++; },
    fillText(t) { 記録.fillText.push(String(t)); },
    measureText(t) { return { width: String(t).length * 6 }; },
    set fillStyle(v) {}, get fillStyle() { return ''; },
    set strokeStyle(v) {}, get strokeStyle() { return ''; },
    set lineWidth(v) {}, get lineWidth() { return 1; },
    set font(v) {}, get font() { return ''; },
    set textAlign(v) {}, get textAlign() { return ''; },
    set textBaseline(v) {}, get textBaseline() { return ''; },
  };
}
{
  const m = C.材料を作る([['月', '売上'], ['1月', 100], ['2月', 150], ['3月', 120]]);
  const c1 = にせctx();
  ok('縦棒が 描けた', C.描く(c1, m, 'column', 360, 216, 'テスト') === true);
  ok('★棒を 3本 描いた★（＋背景と 目盛）', c1.記録.fillRect >= 3, String(c1.記録.fillRect));
  ok('★下の 見出しを 描いた★', c1.記録.fillText.indexOf('1月') >= 0 && c1.記録.fillText.indexOf('3月') >= 0,
    c1.記録.fillText.join(','));
  ok('★題を 描いた（実Excel＝タイトルは 出す）★', c1.記録.fillText.indexOf('テスト') >= 0);
  ok('★系列1本なら 凡例を 出さない（実Excelと 同じ）★', c1.記録.fillText.indexOf('売上') < 0,
    c1.記録.fillText.join(','));

  const c2 = にせctx();
  const m2 = C.材料を作る([['月', '売上', '原価'], ['1月', 100, 60], ['2月', 150, 80]]);
  C.描く(c2, m2, 'column', 360, 216, 'テスト');
  ok('★系列2本なら 凡例を 出す（実Excelと 同じ）★',
    c2.記録.fillText.indexOf('売上') >= 0 && c2.記録.fillText.indexOf('原価') >= 0, c2.記録.fillText.join(','));

  const c3 = にせctx();
  ok('折れ線が 描けた', C.描く(c3, m, 'line', 360, 216, '') === true);
  ok('★折れ線は 線を 引く★', c3.記録.stroke >= 1, String(c3.記録.stroke));

  const c4 = にせctx();
  ok('円が 描けた', C.描く(c4, m, 'pie', 360, 216, '') === true);
  ok('★円は 扇を 3つ 描く★', c4.記録.arc === 3, String(c4.記録.arc));
}
{
  const c5 = にせctx();
  ok('★材料が 無い時は 描かない★', C.描く(c5, { ok: false }, 'column', 360, 216, '') === false);
}

console.log('\n[⑤ 画面に 繋がっている]');
ok('部品を 読み込んでいる', /lib\/chart\.js/.test(book));
ok('グラフの窓が 在る', /id="chartOverlay"/.test(book));
ok('canvas が 在る', /id="chartCanvas"/.test(book));
ok('★外の部品を 1つも 読み込んでいない★（chart.js は 自前）',
  !/chart\.js.*cdn|chartjs|d3\.js|echarts/i.test(book));
ok('★alert / prompt / confirm を 使っていない★', !/\balert\(|\bprompt\(|\bconfirm\(/.test(book));

console.log('\n[⑥ リボンから 押せる]');
const ACT = require_(path.join(ROOT, 'lib/ribbon-actions.js'));
{
  const g = globalThis, 前w = g.window;
  const 受け = [];
  g.window = { グラフを作る: function (k) { 受け.push(k); } };
  ACT['縦棒グラフ'](); ACT['折れ線グラフ'](); ACT['円グラフ']();
  g.window = 前w;
  ok('★3つとも 別の形を 渡す★', JSON.stringify(受け) === JSON.stringify(['column', 'line', 'pie']),
    JSON.stringify(受け));
}

console.log('\nchart: ' + 緑 + '/' + (緑 + 赤) + ' passed');

if (壊す) {
  console.log('\n★--self-test＝わざと 壊して 赤に なるか★');
  let 素通り = 0;
  /* 壊し① 系列1本でも 凡例を 出す（実Excelと 違う）
     ★見出しの 字と 混ぜて 見ない★＝見出しは 下に 出るので、見るのは ★系列の 名前★だけ。
     （2026-08-30 に 横に取る 直しを 入れた時、ここが「見出しの字」を 拾って 空振りしていた） */
  const c = にせctx();
  const m = C.材料を作る([['月', '売上'], ['1月', 100]]);
  const 系列名 = m.ok ? m.系列[0].名 : '(材料が作れない)';
  const 見出しに無い = m.ok && m.見出し.indexOf(系列名) < 0;
  C.描く(c, m, 'column', 360, 216, '');
  if (!m.ok || !見出しに無い) {
    素通り++; console.log('  ★素通り★ 1系列の 材料が 作れない／系列名が 見出しと かぶって 見分けられない');
  } else if (c.記録.fillText.indexOf(系列名) >= 0) {
    素通り++; console.log('  ★素通り★ 系列1本で 凡例を 出した（' + 系列名 + '）');
  } else {
    console.log('  ok   系列1本（' + 系列名 + '）では 凡例を 出していない');
  }
  /* 壊し② 数が無いのに 描けたと 言う */
  if (C.材料を作る([['あ']]).ok) { 素通り++; console.log('  ★素通り★ 数が 無いのに ok'); }
  /* 壊し③ 大きさが 実Excelと 違う */
  if (C.既定.幅 !== 360 || C.既定.高さ !== 216) { 素通り++; console.log('  ★素通り★ 大きさが 実測と 違う'); }
  if (素通り) { console.log('★抜け道 ' + 素通り + '件★'); process.exit(1); }
  console.log('  ok   抜け道 0件');
}
process.exit(赤 ? 1 : 0);
