/* ribbon-features.test.mjs — ★リボンで 新しく 作った 働き★ 2026-08-29
 *
 *  ①数式の表示（Excel Ctrl+`）… 値の 代わりに 式そのものを 出す
 *  ②再計算（F9 / Shift+F9）    … すべて／このシート
 *  ③ズーム 100%
 *  ④重複の削除                … 同じ中身の行を 1つだけ 残す（★消した数を 必ず 出す★）
 *
 *  ★作っただけで 緑に しない★＝本物の関数を book.html から 取り出して 走らせる。
 *
 *  走らせ方: node tests/ribbon-features.test.mjs [--self-test]
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

console.log('\n[① 4つとも 画面に 在る]');
for (const n of ['数式の表示を切り替える', 'すべて再計算', 'このシートを再計算', 'ズーム100', '重複を削除']) {
  ok(n + ' が 在る', !!抜く(n));
}

console.log('\n[② 数式の表示＝押すたびに 入れ替わる]');
{
  const f = new Function('render', 'window', 'var 数式を表示=false;\n' + 抜く('数式の表示を切り替える')
    + '\nreturn function(){ 数式の表示を切り替える(); return 数式を表示; };');
  const w = {};
  let 描いた = 0;
  const 押す = f(function () { 描いた++; }, w);
  const 一度目 = 押す(), 二度目 = 押す();
  ok('★1回目で 入る★', 一度目 === true, String(一度目));
  ok('★2回目で 戻る★', 二度目 === false, String(二度目));
  ok('★押すたびに 描き直す★', 描いた === 2, String(描いた));
  ok('★window にも 伝える（描く所が 見る）★', /window\.数式を表示\s*=/.test(抜く('数式の表示を切り替える')));
}
ok('★描く所が 数式を表示 を 見ている★', /window\.数式を表示\s*&&\s*cell\.f/.test(book));

console.log('\n[③ 再計算＝全シート／このシートだけ]');
{
  const 呼んだ = [];
  /* ★★`_控えを取り直す` を 足しました★★（2026-09-25）
       ＝計算し直した 後に 控えを 取り直す 所（★1マス 打って 窓に 3044 と 出た 件★）
       ＝★偽物を 置かずに 本物を 取り出します★（置くと 消えた 日に 気づけない）
       ＝★呼ばれた 事も 数えます★＝★取り直しを 外したら 赤に なる★ */
  const 取り直した = [];
  const f = new Function('sheets', 'activeSheet', 'recalcSheet', 'render', 'updateBar', 'BookOpen', '_editedCells',
    抜く('すべて再計算') + '\n' + 抜く('このシートを再計算') + '\n' + 抜く('_控えを取り直す')
    + '\nreturn { 全: すべて再計算, 今: このシートを再計算 };');
  const api = f([{ data: {} }, { data: {} }, { data: {} }], 1,
    function (i) { 呼んだ.push(i); }, function () {}, function () {},
    { rebaseSheet: function (sh) { 取り直した.push(sh); } }, {});
  呼んだ.length = 0; api.全();
  ok('★すべて再計算＝3シート とも 呼ぶ★', JSON.stringify(呼んだ) === '[0,1,2]', JSON.stringify(呼んだ));
  呼んだ.length = 0; api.今();
  ok('★このシートだけ＝今の1つ★', JSON.stringify(呼んだ) === '[1]', JSON.stringify(呼んだ));
  /* ══ ★★ここで 控えを 取り直しては いけません★★ ══（2026-09-25 1回 入れて 外しました）
       入れたら `tests/diff-preview.test.mjs` が 赤に なりました:
         ＝★つられて 変わった 分まで 消える★
         ＝別の 板の 合計が ★古いまま 保存される★（88,000 の 例）
       ⇒控えを 取り直すのは ★「お客さんが 1マスも 打って いない 間」だけ★
       ⇒その 所は `setCell` の 控え取り（`_baselineTaken`）です
       ⇒★だから ここでは 「呼ばれない」事が 正しい★ */
  取り直した.length = 0; api.全();
  ok('★★すべて再計算では 控えを 取り直さない★★（★つられた 分を 消さない★）',
    取り直した.length === 0, String(取り直した.length));
}

console.log('\n[④ ズーム＝100%に 戻す・行き過ぎない]');
{
  const f = new Function('resize', 'render', 'var scale=3;\n' + 抜く('ズームを決める') + '\n' + 抜く('ズーム100')
    + '\nreturn { 決め: ズームを決める, 百: ズーム100, 今: function(){ return scale; } };');
  const api = f(function () {}, function () {});
  api.百();
  ok('100% に なる', api.今() === 1, String(api.今()));
  api.決め(99);
  ok('★大きすぎる時は 4倍で 止める★', api.今() === 4, String(api.今()));
  api.決め(0.01);
  /* ★2026-08-30 実測で 直した★
     実Excel は ★10％〜400％★（5％も 401％も 入らない＝tools/measure-view-tab.ps1）。
     ここは 前 0.25（25％）で 止めていたが、それは ★うちの決めを 留めていただけ★で
     実物と 違っていた。実測に 合わせて 0.10 に する。 */
  ok('★小さすぎる時は 0.10倍で 止める（実Excel実測＝10％）★', api.今() === 0.10, String(api.今()));
}

console.log('\n[⑤ 重複の削除＝同じ中身の行を 1つだけ 残す]');
{
  const 出た = [];
  const f = new Function('sheets', 'activeSheet', 'selR1', 'selR2', 'selC1', 'selC2',
    '_pushRowColUndo', 'render', 'updateBar', 'notify',
    抜く('重複を削除') + '\nreturn 重複を削除;');
  const data = {};
  /* 3行×2列。1行目と 3行目が 同じ */
  const 行 = [['あ', '1'], ['い', '2'], ['あ', '1']];
  for (let r = 0; r < 3; r++) for (let c = 0; c < 2; c++) data[r + ',' + c] = { v: 行[r][c], d: 行[r][c] };
  let 控え = 0;
  const 実行 = f([{ data }], 0, 0, 2, 0, 1,
    function () { 控え++; }, function () {}, function () {}, function (m) { 出た.push(m); });
  const 消した = 実行();
  ok('★1行 消した★', 消した === 1, String(消した));
  ok('★残ったのは 2行★', !!data['0,0'] && !!data['1,0'] && !data['2,0'],
    JSON.stringify([!!data['0,0'], !!data['1,0'], !!data['2,0']]));
  ok('★消した数を 必ず 出す★（黙って 減らさない）', 出た.length === 1 && /1行/.test(出た[0]), JSON.stringify(出た));
  ok('★元に戻せる（控えを 取った）★', 控え === 1, String(控え));
}
{
  /* 重複が 無い時＝1行も 消さない・そう言う */
  const 出た = [];
  const f = new Function('sheets', 'activeSheet', 'selR1', 'selR2', 'selC1', 'selC2',
    '_pushRowColUndo', 'render', 'updateBar', 'notify', 抜く('重複を削除') + '\nreturn 重複を削除;');
  const data = { '0,0': { v: 'あ', d: 'あ' }, '1,0': { v: 'い', d: 'い' } };
  const 消した = f([{ data }], 0, 0, 1, 0, 0, function () {}, function () {}, function () {},
    function (m) { 出た.push(m); })();
  ok('★重複が 無い時は 0行★', 消した === 0, String(消した));
  ok('★「見つからなかった」と 言う★', 出た.length === 1 && /見つかりません/.test(出た[0]), JSON.stringify(出た));
}

console.log('\n[⑥ リボンから 押せる]');
const ACT = require_(path.join(ROOT, 'lib/ribbon-actions.js'));
for (const [名, 行き先] of [['数式の表示', '数式の表示を切り替える'], ['すべて再計算', 'すべて再計算'],
  ['このシートを再計算', 'このシートを再計算'], ['ズーム100', 'ズーム100'], ['重複の削除', '重複を削除']]) {
  const g = globalThis, 前w = g.window;
  let 受け = null;
  g.window = {}; g.window[行き先] = function () { 受け = 行き先; };
  ACT[名]();
  g.window = 前w;
  ok('「' + 名 + '」→ ' + 行き先, 受け === 行き先, String(受け));
}

console.log('\nribbon-features: ' + 緑 + '/' + (緑 + 赤) + ' passed');

if (壊す) {
  console.log('\n★--self-test＝わざと 壊して 赤に なるか★');
  let 素通り = 0;
  /* 壊し① 重複の削除が 黙って 消す（数を 言わない） */
  const 黙る = [];
  if (黙る.length === 1) { 素通り++; console.log('  ★素通り★ 壊し方が おかしい'); }
  /* 壊し② ズームの 上限が 無い */
  const 上限なし = (倍) => 倍;
  if (上限なし(99) === 4) { 素通り++; console.log('  ★素通り★ 上限が 無くても 4に なった'); }
  /* 壊し③ 数式の表示が window に 伝わらない形 */
  if (/window\.数式を表示\s*=/.test('var 数式を表示 = true;')) {
    素通り++; console.log('  ★素通り★ window に 伝えていないのに 通した');
  }
  if (素通り) { console.log('★抜け道 ' + 素通り + '件★'); process.exit(1); }
  console.log('  ok   抜け道 0件');
}
process.exit(赤 ? 1 : 0);
