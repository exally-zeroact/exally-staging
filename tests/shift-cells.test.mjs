/* shift-cells.test.mjs — ★セルの挿入／削除（1つだけ ずらす）★ 2026-08-29
 *
 *  ★真値（実Excel 16.0 で 実測）★
 *    A1..A3 / B1..B3 が 在る所で A2 を「下へ ずらして 挿入」:
 *      行1 A1 B1 ／ 行2 （空） B2 ／ 行3 A2 B3 ／ 行4 A3 （空）
 *      ⇒ ★選んだ セルの 列だけ 動く。B列は 動かない★
 *    A2 を「上へ ずらして 削除」:
 *      行1 A1 B1 ／ 行2 A3 B2 ／ 行3 （空） B3
 *
 *  走らせ方: node tests/shift-cells.test.mjs [--self-test]
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
for (const n of ['セルを挿入', 'セルを削除', 'セルを下へ挿入', 'セルを右へ挿入',
  'セルを上へ削除', 'セルを左へ削除']) ok(n + ' が 在る', !!抜く(n));

function 台(選) {
  const s = { data: {} };
  for (let r = 0; r < 3; r++) {
    s.data[r + ',0'] = { v: 'A' + (r + 1), d: 'A' + (r + 1) };
    s.data[r + ',1'] = { v: 'B' + (r + 1), d: 'B' + (r + 1) };
  }
  const 出た = [];
  let 控え = 0;
  const f = new Function('sheets', 'activeSheet', 'selR1', 'selR2', 'selC1', 'selC2',
    '_pushRowColUndo', 'render', 'updateBar', 'notify',
    抜く('セルを挿入') + '\n' + 抜く('セルを削除') + '\nreturn { 入: セルを挿入, 消: セルを削除 };');
  const api = f([s], 0, 選[0], 選[1], 選[2], 選[3],
    function () { 控え++; }, function () {}, function () {}, function (m) { 出た.push(m); });
  return { s, api, 出た, 控え: () => 控え, 見る: (r, c) => (s.data[r + ',' + c] ? s.data[r + ',' + c].v : '') };
}

console.log('\n[② ★挿入（下へ）＝実Excelと 1マスも 違わない★]');
{
  const t = 台([1, 1, 0, 0]);
  t.api.入('down');
  const 期待 = [['A1', 'B1'], ['', 'B2'], ['A2', 'B3'], ['A3', '']];
  for (let r = 0; r < 4; r++) {
    ok('行' + (r + 1) + ' … A=[' + 期待[r][0] + '] B=[' + 期待[r][1] + ']',
      t.見る(r, 0) === 期待[r][0] && t.見る(r, 1) === 期待[r][1],
      '出た A=[' + t.見る(r, 0) + '] B=[' + t.見る(r, 1) + ']');
  }
  ok('★B列は 動いていない★', t.見る(0, 1) === 'B1' && t.見る(1, 1) === 'B2' && t.見る(2, 1) === 'B3');
  ok('★元に戻せる★', t.控え() === 1, String(t.控え()));
  ok('★何を したかを 言う★', /下へ ずらしました/.test(t.出た.join('')), JSON.stringify(t.出た));
}

console.log('\n[③ ★削除（上へ）＝実Excelと 1マスも 違わない★]');
{
  const t = 台([1, 1, 0, 0]);
  t.api.消('up');
  const 期待 = [['A1', 'B1'], ['A3', 'B2'], ['', 'B3']];
  for (let r = 0; r < 3; r++) {
    ok('行' + (r + 1) + ' … A=[' + 期待[r][0] + '] B=[' + 期待[r][1] + ']',
      t.見る(r, 0) === 期待[r][0] && t.見る(r, 1) === 期待[r][1],
      '出た A=[' + t.見る(r, 0) + '] B=[' + t.見る(r, 1) + ']');
  }
  ok('★B列は 動いていない★', t.見る(0, 1) === 'B1' && t.見る(1, 1) === 'B2' && t.見る(2, 1) === 'B3');
}

console.log('\n[④ 右へ／左へ]');
{
  const t = 台([0, 0, 0, 0]);
  t.api.入('right');
  ok('★右へ ずらすと A1 が B1 へ★', t.見る(0, 1) === 'A1', t.見る(0, 1));
  ok('★元の所は 空★', t.見る(0, 0) === '', t.見る(0, 0));
  ok('★下の行は 動かない★', t.見る(1, 0) === 'A2', t.見る(1, 0));
}
{
  const t = 台([0, 0, 0, 0]);
  t.api.消('left');
  ok('★左へ 詰めると B1 が A1 へ★', t.見る(0, 0) === 'B1', t.見る(0, 0));
  ok('★下の行は 動かない★', t.見る(1, 0) === 'A2', t.見る(1, 0));
}

console.log('\n[⑤ リボンから 押せる]');
const ACT = require_(path.join(ROOT, 'lib/ribbon-actions.js'));
for (const [名, 行き先] of [['セルを挿入', 'セルを下へ挿入'], ['セルを削除', 'セルを上へ削除']]) {
  const g = globalThis, 前w = g.window;
  let 受け = null;
  g.window = {}; g.window[行き先] = function () { 受け = 行き先; };
  ACT[名]();
  g.window = 前w;
  ok('「' + 名 + '」→ ' + 行き先, 受け === 行き先, String(受け));
}

console.log('\nshift-cells: ' + 緑 + '/' + (緑 + 赤) + ' passed');

if (壊す) {
  console.log('\n★--self-test＝この 見張りが 見ている物を 直に 確かめる（★壊して 赤を 見るのは tools/break-check.mjs★）★');
  let 素通り = 0;
  /* 壊し① 隣の列まで 動かす（実Excelと 違う） */
  const t = 台([1, 1, 0, 0]);
  t.api.入('down');
  if (t.見る(1, 1) !== 'B2') { 素通り++; console.log('  ★素通り★ B列が 動いてしまった'); }
  /* 壊し② 上書きして 消える（下から 動かしていない） */
  if (t.見る(3, 0) !== 'A3') { 素通り++; console.log('  ★素通り★ A3 が 消えた'); }
  /* 壊し③ 元に戻す控えを 取らない */
  if (t.控え() === 0) { 素通り++; console.log('  ★素通り★ 控えを 取っていない'); }
  if (素通り) { console.log('★抜け道 ' + 素通り + '件★'); process.exit(1); }
  console.log('  ok   抜け道 0件');
}
process.exit(赤 ? 1 : 0);
