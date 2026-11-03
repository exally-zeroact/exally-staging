/* decimal-painter.test.mjs — ★小数の 増減★と★書式のコピー／貼り付け★ 2026-08-29
 *
 *  ★① 小数の 桁★
 *    Excelの ボタンは ★2つ（増やす／減らす）★。うちは 1つで 回る形だった＝★Excelと 違う★。
 *    ★上限は 30桁★＝実Excelは 書式として 31桁でも 受け付けるが、表示形式の 決まりは 30桁まで。
 *    （★最初 勝手に 6桁で 止めていた。測って 直した★）
 *
 *  ★② 書式のコピー／貼り付け（刷毛）★
 *    実Excelと 同じ＝★中身は 写さない。書式だけ★。
 *
 *  走らせ方: node tests/decimal-painter.test.mjs [--self-test]
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
for (const n of ['_小数の桁', '_桁で書式を作る', '小数を増やす', '小数を減らす', '書式をコピー', '書式を貼る']) {
  ok(n + ' が 在る', !!抜く(n));
}

console.log('\n[② 桁の 数え方と 書式の 作り方]');
{
  const f = new Function(抜く('_小数の桁') + '\n' + 抜く('_桁で書式を作る')
    + '\nreturn { 桁: _小数の桁, 作る: _桁で書式を作る };');
  const api = f();
  for (const [fmt, 期待] of [['', 0], ['#,##0', 0], ['#,##0.0', 1], ['#,##0.00', 2], ['0.000', 3], ['0%', 0]]) {
    ok('桁を 数える … [' + fmt + '] → ' + 期待, api.桁(fmt) === 期待, String(api.桁(fmt)));
  }
  ok('0桁 → #,##0', api.作る('#,##0.0', 0) === '#,##0');
  ok('2桁 → #,##0.00', api.作る('#,##0', 2) === '#,##0.00');
  ok('★区切りが 無い書式は 保つ★', api.作る('0.0', 2) === '0.00', api.作る('0.0', 2));
}

console.log('\n[③ 増やす／減らす（★端で 止まって 理由を 言う★）]');
{
  function 台(numFmt) {
    const 出た = [], 当てた = [];
    const f = new Function('getCell', 'applyFormat', 'notify', 'selR1', 'selC1',
      抜く('_小数の桁') + '\n' + 抜く('_桁で書式を作る') + '\n' + 抜く('小数を増やす') + '\n' + 抜く('小数を減らす')
      + '\nreturn { 増: 小数を増やす, 減: 小数を減らす };');
    const api = f(function () { return { numFmt }; },
      function (k, v) { 当てた.push([k, v]); }, function (m) { 出た.push(m); }, 0, 0);
    return { api, 出た, 当てた };
  }
  {
    const t = 台('#,##0');
    ok('0桁 → 増やすと 1桁', t.api.増() === 1 && t.当てた[0][1] === '#,##0.0', JSON.stringify(t.当てた));
  }
  {
    const t = 台('#,##0.00');
    ok('2桁 → 減らすと 1桁', t.api.減() === 1 && t.当てた[0][1] === '#,##0.0', JSON.stringify(t.当てた));
  }
  {
    const t = 台('#,##0');
    ok('★0桁で 減らすと 0のまま★', t.api.減() === 0);
    ok('★理由を 言う（黙って 何も 起きない、を しない）★', /これ以上は 減らせません/.test(t.出た.join('')), JSON.stringify(t.出た));
    ok('★何も 当てない★', t.当てた.length === 0, JSON.stringify(t.当てた));
  }
  {
    const t = 台('0.' + '0'.repeat(30));
    ok('★30桁で 増やすと 30のまま（実Excelの 上限）★', t.api.増() === 30);
    ok('★理由を 言う★', /30桁/.test(t.出た.join('')), JSON.stringify(t.出た));
  }
}

console.log('\n[④ 書式のコピー／貼り付け（★中身は 写さない★）]');
{
  const 出た = [];
  const s = {
    data: {
      '0,0': { v: '元の値', d: '元の値', bold: true, bg: '#FFFF00', align: 'right' },
      '1,0': { v: '別の値', d: '別の値' },
    },
  };
  const キー行 = book.match(/var 書式のキー = \[[\s\S]*?\];/)[0];
  const f = new Function('sheets', 'activeSheet', 'selR1', 'selR2', 'selC1', 'selC2',
    'getCell', 'undoStack', 'redoStack', 'render', 'updateBar', 'notify',
    キー行 + '\nvar _写した書式 = null;\n' + 抜く('書式をコピー') + '\n' + 抜く('書式を貼る')
    + '\nreturn { 写す: 書式をコピー, 貼る: 書式を貼る };');
  /* まず 0,0 を 選んで コピー */
  const api1 = f([s], 0, 0, 0, 0, 0, function () { return s.data['0,0']; }, [], [],
    function () {}, function () {}, function (m) { 出た.push(m); });
  const 写し = api1.写す();
  ok('★書式だけ 覚えた★', 写し.bold === true && 写し.bg === '#FFFF00' && 写し.v === undefined,
    JSON.stringify(写し));
  ok('★何個 覚えたかを 言う★', /個の 書式/.test(出た.join('')), JSON.stringify(出た));
}
{
  /* コピー→貼る を 通しで（同じ 台の中で） */
  const 出た = [];
  const s = {
    data: {
      '0,0': { v: '元', d: '元', bold: true, bg: '#FFFF00' },
      '1,0': { v: '先', d: '先' },
    },
  };
  const キー行 = book.match(/var 書式のキー = \[[\s\S]*?\];/)[0];
  const f = new Function('sheets', 'activeSheet', 'getCell', 'undoStack', 'redoStack',
    'render', 'updateBar', 'notify',
    キー行 + '\nvar _写した書式 = null;\nvar selR1=0,selR2=0,selC1=0,selC2=0;\n'
    + 抜く('書式をコピー') + '\n' + 抜く('書式を貼る')
    + '\nreturn { 写す: 書式をコピー, 貼る: function(r){ selR1=selR2=r; return 書式を貼る(); } };');
  const undo = [];
  const api = f([s], 0, function () { return s.data['0,0']; }, undo, [],
    function () {}, function () {}, function (m) { 出た.push(m); });
  api.写す();
  const n = api.貼る(1);
  ok('1個に 貼った', n === 1, String(n));
  ok('★書式が 移った★', s.data['1,0'].bold === true && s.data['1,0'].bg === '#FFFF00',
    JSON.stringify(s.data['1,0']));
  ok('★中身は 変えていない★（実Excelと 同じ）', s.data['1,0'].v === '先', String(s.data['1,0'].v));
  ok('★元に戻せる★', undo.length === 1, String(undo.length));
  ok('★貼ったと 言う★', /書式を 貼りました/.test(出た.join('')), JSON.stringify(出た));
}
{
  /* 先に コピーしていない時 */
  const 出た = [];
  const キー行 = book.match(/var 書式のキー = \[[\s\S]*?\];/)[0];
  const f = new Function('sheets', 'activeSheet', 'selR1', 'selR2', 'selC1', 'selC2',
    'undoStack', 'redoStack', 'render', 'updateBar', 'notify',
    キー行 + '\nvar _写した書式 = null;\n' + 抜く('書式を貼る') + '\nreturn 書式を貼る;');
  const n = f([{ data: {} }], 0, 0, 0, 0, 0, [], [], function () {}, function () {},
    function (m) { 出た.push(m); })();
  ok('★先に コピーしていなければ 0★', n === 0, String(n));
  ok('★そう言う★', /先に 書式のコピー/.test(出た.join('')), JSON.stringify(出た));
}

console.log('\n[⑤ リボンから 押せる]');
const ACT = require_(path.join(ROOT, 'lib/ribbon-actions.js'));
for (const [名, 行き先] of [['小数を増やす', '小数を増やす'], ['小数を減らす', '小数を減らす'],
  ['書式のコピー', '書式をコピー']]) {
  const g = globalThis, 前w = g.window;
  let 受け = null;
  g.window = {}; g.window[行き先] = function () { 受け = 行き先; };
  ACT[名]();
  g.window = 前w;
  ok('「' + 名 + '」→ ' + 行き先, 受け === 行き先, String(受け));
}

console.log('\ndecimal-painter: ' + 緑 + '/' + (緑 + 赤) + ' passed');

if (壊す) {
  console.log('\n★--self-test＝この 見張りが 見ている物を 直に 確かめる（★壊して 赤を 見るのは tools/break-check.mjs★）★');
  let 素通り = 0;
  /* 壊し① 上限を 6桁に 戻す（勝手な決め） */
  if (/これ以上は 増やせません（6桁まで）/.test(book)) { 素通り++; console.log('  ★素通り★ 上限が 6桁に 戻っている'); }
  /* 壊し② 1つで 回る形（Excelと 違う） */
  const ACT2 = require_(path.join(ROOT, 'lib/ribbon-actions.js'));
  if (typeof ACT2['小数を減らす'] !== 'function') { 素通り++; console.log('  ★素通り★ 減らすが 無い'); }
  /* 壊し③ 書式のコピーが 中身も 写す */
  if (書式に中身が入る()) { 素通り++; console.log('  ★素通り★ 中身まで 写している'); }
  function 書式に中身が入る() {
    const キー行 = book.match(/var 書式のキー = \[[\s\S]*?\];/)[0];
    return /'v'|'d'|'f'/.test(キー行);
  }
  if (素通り) { console.log('★抜け道 ' + 素通り + '件★'); process.exit(1); }
  console.log('  ok   抜け道 0件');
}
process.exit(赤 ? 1 : 0);
