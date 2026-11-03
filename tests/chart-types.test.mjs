/* chart-types.test.mjs — ★グラフの 種類を 増やした分★ 2026-08-30
 *
 *  ★真値（実Excel 16.0 で 実測）★
 *    ① `tools/measure-charts.ps1` … 見出し1列＋数2列＋4行 を 選んで AddChart2。
 *       結果は `tests/fixtures/excel-chart-types.txt` に そのまま 置いた。
 *       ★大きさは どの種類も 360x216／題は どの種類も 出す★
 *       ★凡例は「系列が2本以上」だけでは 決まらない★
 *          ツリーマップ … 系列1本でも ★出す★
 *          等高線       … 系列2本でも ★出さない★
 *          ヒストグラム・箱ひげ・サンバースト・じょうご … ★出さない★
 *    ② 箱ひげ図（COMの性質が 読めなかったので ★保存した中身★を 読んだ）
 *       xl/charts/chartEx1.xml
 *         <cx:statistics quartileMethod="exclusive" />
 *         <cx:visibility meanLine="0" meanMarker="1" nonoutliers="0" outliers="1" />
 *       ＋ QUARTILE.EXC(1〜10) = ★2.75 / 5.5 / 8.25★（INC は 3.25 / 5.5 / 7.75）
 *
 *  走らせ方: node tests/chart-types.test.mjs [--self-test]
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
const 近い = (a, b) => Math.abs(a - b) < 1e-9;

const Chart = require_(path.join(ROOT, 'lib/chart.js'));
const book = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf8');

/* ───────── ① 実測の 紙が 在るか（数え直せる形で 残っているか）───────── */
console.log('\n[① 実測の 紙]');
const 紙の道 = path.join(ROOT, 'tests/fixtures/excel-chart-types.txt');
ok('実Excelを 測った 紙が 在る', fs.existsSync(紙の道));
const 紙 = fs.existsSync(紙の道) ? fs.readFileSync(紙の道, 'utf8') : '';
ok('測る道具も 残っている', fs.existsSync(path.join(ROOT, 'tools/measure-charts.ps1')));
ok('★どの種類も 360x216★（紙から 数える）',
  (紙.match(/360x216/g) || []).length >= 15, String((紙.match(/360x216/g) || []).length) + '行');

/* ───────── ② 凡例の 出る/出ない が 実測どおり ───────── */
console.log('\n[② ★凡例は 種類ごと★（系列の数だけでは 決まらない）]');
const 一本 = { ok: true, 見出し: ['あ', 'い', 'う'], 系列: [{ 名: 'A', 値: [1, 2, 3] }] };
const 二本 = { ok: true, 見出し: ['あ', 'い', 'う'], 系列: [{ 名: 'A', 値: [1, 2, 3] }, { 名: 'B', 値: [3, 2, 1] }] };
for (const [種類, 一本で, 二本で, なぜ] of [
  ['column',    false, true,  ''],
  ['line',      false, true,  ''],
  ['pie',       false, true,  ''],
  ['scatter',   false, true,  ''],
  ['radar',     false, true,  ''],
  ['treemap',   true,  true,  '★1本でも 出す★'],
  ['waterfall', true,  true,  '★1本でも 出す★'],
  ['histogram', false, false, '★2本でも 出さない★'],
  ['box',       false, false, '★2本でも 出さない★'],
  ['sunburst',  false, false, '★2本でも 出さない★'],
  ['funnel',    false, false, '★2本でも 出さない★'],
]) {
  ok(種類 + ' 系列1本 → ' + (一本で ? '出す' : '出さない') + (なぜ ? '  ' + なぜ : ''),
    Chart.凡例を出すか(種類, 一本) === 一本で, String(Chart.凡例を出すか(種類, 一本)));
  ok(種類 + ' 系列2本 → ' + (二本で ? '出す' : '出さない'),
    Chart.凡例を出すか(種類, 二本) === 二本で, String(Chart.凡例を出すか(種類, 二本)));
}

/* ───────── ③ 軸の 数が 実測どおり ───────── */
console.log('\n[③ 軸の 数（実測）]');
for (const [種類, 軸] of [
  ['column', 2], ['line', 2], ['area', 2], ['pie', 0], ['doughnut', 0],
  ['scatter', 2], ['bubble', 2], ['radar', 1], ['treemap', 0], ['sunburst', 0],
  ['histogram', 2], ['box', 2], ['waterfall', 2], ['funnel', 1],
]) {
  ok(種類 + ' の 軸 = ' + 軸, Chart.種類の既定[種類] && Chart.種類の既定[種類].軸 === 軸,
    String(Chart.種類の既定[種類] && Chart.種類の既定[種類].軸));
}

/* ───────── ④ 箱ひげ＝実Excel の 四分位（排他）と 同じ 数 ───────── */
console.log('\n[④ ★箱ひげの 四分位＝QUARTILE.EXC（実測 2.75 / 5.5 / 8.25）★]');
{
  const q = Chart.四分位([1, 2, 3, 4, 5, 6, 7, 8, 9, 10]);
  ok('q1 = 2.75（★包含の 3.25 では ない★）', 近い(q.q1, 2.75), String(q.q1));
  ok('中央 = 5.5', 近い(q.中央, 5.5), String(q.中央));
  ok('q3 = 8.25（★包含の 7.75 では ない★）', 近い(q.q3, 8.25), String(q.q3));
  ok('★平均も 持つ（実測 meanMarker=1）★', 近い(q.平均, 5.5), String(q.平均));
  ok('外れ値なし', q.外れ.length === 0, JSON.stringify(q.外れ));
  ok('ひげは 端まで', q.ひげ下 === 1 && q.ひげ上 === 10, q.ひげ下 + '〜' + q.ひげ上);
}
{
  const q = Chart.四分位([1, 2, 3, 4, 5, 6, 7, 8, 9, 100]);
  ok('★外れ値を 見つける（実測 outliers=1）★', q.外れ.length === 1 && q.外れ[0] === 100, JSON.stringify(q.外れ));
  ok('★ひげは 外れ値まで 伸ばさない★', q.ひげ上 === 9, String(q.ひげ上));
}
ok('★4つ 未満は 出せない（排他では 数が 足りない）★', Chart.四分位([1, 2, 3]) === null,
  JSON.stringify(Chart.四分位([1, 2, 3])));

/* ───────── ⑤ ヒストグラム＝系列が 1本に まとまる（実測）───────── */
console.log('\n[⑤ ヒストグラム]');
{
  const m = Chart.ヒストグラムの材料({ 系列: [{ 値: [1, 2, 3] }, { 値: [4, 5, 6] }] });
  ok('★系列が 1本に まとまる（実測 系列=1）★', m.系列.length === 1, String(m.系列.length));
  const 合計 = m.系列[0].値.reduce((a, b) => a + b, 0);
  ok('★数を 1つも 落とさない（6個 入れたら 6個）★', 合計 === 6, String(合計));
  const m2 = Chart.ヒストグラムの材料({ 系列: [{ 値: [1, 1, 1, 1] }] });
  ok('全部 同じ数でも 落ちない', m2.系列[0].値.reduce((a, b) => a + b, 0) === 4,
    JSON.stringify(m2.系列[0].値));
  const m3 = Chart.ヒストグラムの材料({ 系列: [{ 値: [] }] });
  ok('数が 無くても 落ちない', !!m3.ok);
}

/* ───────── ⑥ 散布図の 見立て（実測 tools/measure-scatter.ps1）───────── */
console.log('\n[⑥ ★散布図の X（実Excel を 押して 測った）★]');
{
  /* 実測① 1列目が 字（あ/1/10）→ 系列2本・X=1,2,3 */
  const d1 = Chart.散布の材料(Chart.材料を作る([['あ', 1, 10], ['い', 4, 20], ['う', 9, 30]]));
  ok('①字の1列目 → X は 番号 1,2,3', JSON.stringify(d1.X) === '[1,2,3]', JSON.stringify(d1.X));
  ok('①系列は 2本', d1.系列.length === 2, String(d1.系列.length));
  /* 実測② 数だけ2列 → 系列1本・X=1列目 */
  const d2 = Chart.散布の材料(Chart.材料を作る([[1, 10], [4, 20], [9, 30]]));
  ok('②数だけ2列 → X=1,4,9', JSON.stringify(d2.X) === '[1,4,9]', JSON.stringify(d2.X));
  ok('②系列は 1本（Y=10,20,30）',
    d2.系列.length === 1 && d2.系列[0].値.join(',') === '10,20,30',
    d2.系列.length + ' / ' + (d2.系列[0] ? d2.系列[0].値.join(',') : ''));
  /* 実測③ 数だけ3列 → X=1行目・Y=2行目と3行目 */
  const d3 = Chart.散布の材料(Chart.材料を作る([[1, 10, 100], [4, 20, 200], [9, 30, 300]]));
  ok('③数だけ3列 → X=1,10,100（★1行目★）', JSON.stringify(d3.X) === '[1,10,100]', JSON.stringify(d3.X));
  ok('③系列は 2本（4,20,200 と 9,30,300）',
    d3.系列.length === 2 && d3.系列[0].値.join(',') === '4,20,200' && d3.系列[1].値.join(',') === '9,30,300',
    JSON.stringify(d3.系列.map((s) => s.値.join(','))));
}

/* ───────── ⑥-2 ★系列を 縦に取るか 横に取るか★（実測）───────── */
console.log('\n[⑥-2 ★行数 > 列数 なら 系列は 列・そうでなければ 行★（実測）]');
{
  const 見 = (表) => {
    const m = Chart.材料を作る(表);
    return m.ok ? { 見出し: m.見出し.join(','), 系列: m.系列.map((s) => s.名).join(','), 横: !!m.横に取った } : null;
  };
  const a = 見([['月', '売上', '原価'], ['1月', 100, 60], ['2月', 150, 80]]);
  ok('見出し付 数2行2列 → 系列は 行（1月,2月）', a && a.系列 === '1月,2月' && a.見出し === '売上,原価',
    JSON.stringify(a));
  const b = 見([['月', '売上', '原価'], ['1月', 100, 60], ['2月', 150, 80], ['3月', 120, 70]]);
  ok('見出し付 数3行2列 → 系列は 列（売上,原価）', b && b.系列 === '売上,原価' && b.見出し === '1月,2月,3月',
    JSON.stringify(b));
  const c = 見([['月', '売上', '原価', '粗利'], ['1月', 100, 60, 40], ['2月', 150, 80, 70]]);
  ok('見出し付 数2行3列 → 系列は 行（1月,2月）', c && c.系列 === '1月,2月', JSON.stringify(c));
  const d = 見([[11, 12, 13], [21, 22, 23], [31, 32, 33]]);
  ok('数3行3列 → 系列は 行（3本）', d && d.系列 === '系列1,系列2,系列3' && d.横 === true, JSON.stringify(d));
  const e = 見([[11, 12], [21, 22], [31, 32], [41, 42]]);
  ok('数4行2列 → 系列は 列（2本）', e && e.系列 === '系列1,系列2' && e.横 === false, JSON.stringify(e));
  ok('★1列だけなら ひっくり返さない★', 見([[100], [150], [120]]).横 === false,
    JSON.stringify(見([[100], [150], [120]])));
  ok('測る道具が 残っている（数え直せる）',
    fs.existsSync(path.join(ROOT, 'tools/measure-orient.ps1'))
    && fs.existsSync(path.join(ROOT, 'tools/measure-orient2.ps1'))
    && fs.existsSync(path.join(ROOT, 'tools/measure-scatter.ps1')));
}

/* ───────── ⑦ ★本当に 描けるか★（絵の具の 数を 数える）───────── */
console.log('\n[⑦ ★全部の 種類を 描いて 画に 出るか★]');
{
  /* 小さな canvas の 代わり（描いた 命令を 数える＝jsdom も canvas も 要らない） */
  function 偽ctx() {
    const 出た = { fillRect: 0, arc: 0, stroke: 0, fill: 0, fillText: 0 };
    const 何もしない = () => {};
    return {
      _出た: 出た,
      clearRect: 何もしない, save: 何もしない, restore: 何もしない,
      beginPath: 何もしない, closePath: 何もしない, moveTo: 何もしない, lineTo: 何もしない,
      strokeRect: 何もしない, setTransform: 何もしない, measureText: () => ({ width: 20 }),
      fillRect: () => { 出た.fillRect++; },
      arc: () => { 出た.arc++; },
      stroke: () => { 出た.stroke++; },
      fill: () => { 出た.fill++; },
      fillText: () => { 出た.fillText++; },
      fillStyle: '', strokeStyle: '', lineWidth: 1, font: '', textAlign: '', textBaseline: '',
      globalAlpha: 1,
    };
  }
  const 材料 = Chart.材料を作る([
    ['あ', 1, 10], ['い', 4, 20], ['う', 9, 30], ['え', 16, 40],
    ['お', 25, 50], ['か', 36, 60],
  ]);
  ok('材料が 作れた', 材料.ok, 材料.なぜ);
  for (const 種類 of ['column', 'line', 'area', 'pie', 'doughnut', 'scatter', 'scatterLine',
    'bubble', 'radar', 'treemap', 'sunburst', 'histogram', 'box', 'waterfall', 'funnel', 'combo']) {
    const c = 偽ctx();
    let 落ちた = null, 返り = null;
    try { 返り = Chart.描く(c, 材料, 種類, 360, 216, '題'); } catch (e) { 落ちた = String(e && e.message); }
    const 絵 = c._出た.fillRect + c._出た.arc + c._出た.stroke + c._出た.fill;
    ok(種類 + ' … 落ちない・★画に 何か 出る★', 落ちた === null && 返り === true && 絵 > 3,
      (落ちた || ('返り=' + 返り + ' 命令=' + 絵)));
  }
  /* ★知らない 種類でも 落ちない（縦棒に なる）★ */
  const c2 = 偽ctx();
  ok('知らない 種類でも 落ちない', Chart.描く(c2, 材料, 'しらない', 360, 216, '題') === true);
}

/* ───────── ⑦-2 ★種類ごとに 本当に 違う絵か★ ─────────
   ★これが 無いと「combo と 書いたのに 中では 折れ線」でも 緑に なる（実際 1回 そう なっていた）★ */
console.log('\n[⑦-2 ★種類ごとに 違う絵か★]');
{
  function 数える(種類) {
    const 出た = { fillRect: 0, arc: 0, stroke: 0, fill: 0 };
    const 何も = () => {};
    const c = {
      clearRect: 何も, save: 何も, restore: 何も, beginPath: 何も, closePath: 何も,
      moveTo: 何も, lineTo: 何も, strokeRect: 何も, fillText: 何も,
      measureText: () => ({ width: 10 }),
      fillRect: () => { 出た.fillRect++; }, arc: () => { 出た.arc++; },
      stroke: () => { 出た.stroke++; }, fill: () => { 出た.fill++; },
      fillStyle: '', strokeStyle: '', lineWidth: 1, font: '', textAlign: '', textBaseline: '', globalAlpha: 1,
    };
    Chart.描く(c, Chart.材料を作る([['あ', 1, 10], ['い', 4, 20], ['う', 9, 30]]), 種類, 360, 216, '題');
    return JSON.stringify(出た);
  }
  const 絵 = {};
  for (const k of ['column', 'line', 'area', 'combo', 'pie', 'doughnut', 'scatter', 'bubble',
    'radar', 'treemap', 'sunburst', 'waterfall', 'funnel']) 絵[k] = 数える(k);
  ok('★複合が 折れ線と 同じ絵に なっていない★', 絵.combo !== 絵.line, 絵.combo + ' / ' + 絵.line);
  ok('★複合が 縦棒と 同じ絵に なっていない★', 絵.combo !== 絵.column, 絵.combo + ' / ' + 絵.column);
  ok('★面が 折れ線と 同じ絵に なっていない（下を 塗る）★', 絵.area !== 絵.line, 絵.area + ' / ' + 絵.line);
  ok('★ドーナツが 円と 同じ絵に なっていない（真ん中を 抜く）★', 絵.doughnut !== 絵.pie,
    絵.doughnut + ' / ' + 絵.pie);
  ok('★サンバーストが ツリーマップと 同じ絵に なっていない★', 絵.sunburst !== 絵.treemap,
    絵.sunburst + ' / ' + 絵.treemap);
  ok('★じょうごが ウォーターフォールと 同じ絵に なっていない★', 絵.funnel !== 絵.waterfall,
    絵.funnel + ' / ' + 絵.waterfall);
  ok('★バブルが 散布図と 同じ絵に なっていない（大きさを 使う）★', 絵.bubble !== 絵.scatter,
    絵.bubble + ' / ' + 絵.scatter);
  const c2 = JSON.parse(絵.combo);
  ok('複合は 棒（fillRect）も 線（stroke）も 出る', c2.fillRect > 0 && c2.stroke > 0, 絵.combo);
}

/* ───────── ⑧ 出来ない時は ★出来たふりを しない★ ───────── */
console.log('\n[⑧ 出来ない時]');
{
  function 偽ctx2() {
    const 何も = () => {};
    return { clearRect: 何も, save: 何も, restore: 何も, beginPath: 何も, closePath: 何も,
      moveTo: 何も, lineTo: 何も, strokeRect: 何も, fillRect: 何も, arc: 何も, stroke: 何も,
      fill: 何も, fillText: 何も, measureText: () => ({ width: 10 }),
      fillStyle: '', strokeStyle: '', lineWidth: 1, font: '', textAlign: '', textBaseline: '', globalAlpha: 1 };
  }
  const 少ない = Chart.材料を作る([['あ', 1], ['い', 2]]);
  ok('★箱ひげは 数が 足りなければ false を 返す★',
    Chart.描く(偽ctx2(), 少ない, 'box', 360, 216, '題') === false);
  ok('★レーダーは 3つ 無ければ false を 返す★',
    Chart.描く(偽ctx2(), 少ない, 'radar', 360, 216, '題') === false);
}

/* ───────── ⑨ 画面と リボンに つながっている ───────── */
console.log('\n[⑨ 画面と リボン]');
{
  for (const v of ['scatter', 'radar', 'treemap', 'sunburst', 'histogram', 'box', 'waterfall',
    'funnel', 'combo', 'doughnut', 'area', 'bubble', 'scatterLine']) {
    ok('選び場に 在る … ' + v, book.indexOf('value="' + v + '"') >= 0);
  }
  ok('おすすめグラフ が 画面に 在る', /function おすすめグラフ\(/.test(book));
  ok('おすすめの種類 が 画面に 在る', /function おすすめの種類\(/.test(book));

  const ACT = require_(path.join(ROOT, 'lib/ribbon-actions.js'));
  const g = globalThis, 前w = g.window;
  for (const [ボタン, 期待] of [
    ['散布図', 'scatter'], ['階層構造グラフ', 'treemap'], ['統計グラフ', 'histogram'],
    ['ウォーターフォール', 'waterfall'], ['複合グラフ', 'combo'],
  ]) {
    let 受け = null;
    g.window = { グラフを作る: function (k) { 受け = k; } };
    ACT[ボタン]();
    g.window = 前w;
    ok('「' + ボタン + '」を 押すと ' + 期待 + ' で 呼ばれる', 受け === 期待, String(受け));
  }
  let 受け2 = null;
  g.window = { おすすめグラフ: function () { 受け2 = 'ok'; } };
  ACT['おすすめグラフ']();
  g.window = 前w;
  ok('「おすすめグラフ」→ おすすめグラフ', 受け2 === 'ok', String(受け2));
}

/* ───────── ⑩ おすすめの 決め（理由つき）───────── */
console.log('\n[⑩ おすすめグラフの 決め]');
{
  const f = new Function('Chart', book.slice(book.indexOf('function おすすめの種類('),
    book.indexOf('function おすすめグラフ(')) + '\nreturn おすすめの種類;')(Chart);
  ok('見出しが 字から 来ていない（数だけの塊）→ 散布図',
    f(Chart.材料を作る([[1, 10], [2, 20]])) === 'scatter',
    f(Chart.材料を作る([[1, 10], [2, 20]])));
  ok('1系列で 少ない → 円',
    f(Chart.材料を作る([['あ', 1], ['い', 2], ['う', 3]])) === 'pie');
  ok('1系列で 多い → ヒストグラム',
    f(Chart.材料を作る([['あ', 1], ['い', 2], ['う', 3], ['え', 4], ['お', 5], ['か', 6], ['き', 7]])) === 'histogram');
  ok('2系列（字の見出し付き）→ 縦棒（実Excelの 既定と 同じ）',
    f(Chart.材料を作る([['月', '売上', '原価'], ['1月', 100, 60], ['2月', 150, 80], ['3月', 120, 70]])) === 'column',
    f(Chart.材料を作る([['月', '売上', '原価'], ['1月', 100, 60], ['2月', 150, 80], ['3月', 120, 70]])));
}

console.log('\n[⑪ ★違う所は 黙らない（画面に 断りを 出す）★]');
{
  ok('断りの 欄が 画面に 在る', /id="chartNote"/.test(book));
  const 断り塊 = book.slice(book.indexOf('var 断り = {'), book.indexOf('var 訳 = {'));
  const 訳塊 = book.slice(book.indexOf('var 訳 = {'), book.indexOf("var 欄 = document.getElementById('chartNote')"));
  for (const [種類, 語] of [['combo', '第2軸'], ['histogram', '未測定'],
    ['bubble', '3列目'], ['scatter', '1列目']]) {
    ok('描けた時の 断りに「' + 語 + '」が 在る … ' + 種類, 断り塊.indexOf(語) >= 0);
  }
  /* ★描けているのに「4つ 要ります」と 出さない★（実ブラウザで 出ていた） */
  ok('★「4つ 要ります」は 描けた時の 断りには 無い★', 断り塊.indexOf('4つ') < 0, 断り塊.slice(0, 80));
  ok('★「3つ 要ります」は 描けた時の 断りには 無い★', 断り塊.indexOf('3つ') < 0);
  ok('描けない時の 訳に box の「4つ」が 在る', 訳塊.indexOf('4つ') >= 0);
  ok('描けない時の 訳に radar の「3つ」が 在る', 訳塊.indexOf('3つ') >= 0);
  ok('★描けなかった時は 赤で 言う★', /この 中身では 描けません/.test(book));
}

console.log('\nchart-types: ' + 緑 + '/' + (緑 + 赤) + ' passed');

if (壊す) {
  console.log('\n★--self-test＝この 見張りが 見ている物を 直に 確かめる（★壊して 赤を 見るのは tools/break-check.mjs★）★');
  let 素通り = 0;
  /* 壊し① 「系列2本以上なら 出す」だけに 戻したら 実測と 食い違う */
  const 一本 = { 系列: [{ 名: 'A', 値: [1] }] };
  const 昔のやり方 = (材料) => 材料.系列.length >= 2;
  if (昔のやり方(一本) === Chart.凡例を出すか('treemap', 一本)) {
    素通り++; console.log('  ★素通り★ ツリーマップの 1本でも出す が 見えていない');
  } else {
    console.log('  ok   昔のやり方（系列2本以上）だと ツリーマップで 食い違う＝この試験は 見ている');
  }
  /* 壊し② 包含（QUARTILE.INC）に すると 数が 変わるか */
  const 包含 = (v, p) => { const n = v.length, f = p * (n - 1), i = Math.floor(f); return v[i] + (v[i + 1] - v[i]) * (f - i); };
  const inc = 包含([1, 2, 3, 4, 5, 6, 7, 8, 9, 10], 0.25);
  if (近い(inc, Chart.四分位([1, 2, 3, 4, 5, 6, 7, 8, 9, 10]).q1)) {
    素通り++; console.log('  ★素通り★ 包含と 排他が 同じ数＝四分位の 見張りが 効いていない');
  } else {
    console.log('  ok   包含だと ' + inc + '（実測の 2.75 と 違う）＝この試験は 見ている');
  }
  if (素通り) { console.log('★抜け道 ' + 素通り + '件★'); process.exit(1); }
  console.log('  ok   抜け道 0件');
}
process.exit(赤 ? 1 : 0);
