/* data-conn.test.mjs — ★データタブ（詳細設定・クエリと接続）★ 2026-08-30
 *
 *  ★真値（実Excel 16.0 で 実測）★ … tools/measure-data3.ps1
 *    元の 表 A1:A5 ＝ 見出し「名前」／あ・い・あ・う
 *      ① 条件(名前=あ)＋重複残す … ★3行★（見出し・あ・あ）
 *      ② 条件(名前=あ)＋重複除く … ★2行★（見出し・あ）
 *      ③ 条件なし＋重複除く      … ★4行★（見出し・あ・い・う）
 *      ④ 条件なし＋重複残す      … ★5行★
 *      ★見出しは いつも 付く★
 *    つなぎ … `Connections.Count` も `Queries.Count` も はじめ ★0★／
 *      `RefreshAll()` は つなぎが 0でも ★通る（何も 起きない）★
 *    外の ブック … `='C:\ない\[ないブック.xlsx]Sheet1'!A1` と 書くと
 *      `LinkSources(1)` が ★1件★＝★`C:\ない\ないブック.xlsx`★（シート名は 入らない）／
 *      その セルの 答えは ★`#REF!`★
 *
 *  ★測り直した★＝③を はじめ 2行と 読んだが、それは ★前の 条件の 表が 残っていた★せい。
 *    条件の 表を 消して 測り直したら ★4行★だった。
 *    （「前の 物が 残ったまま 測る」と 嘘の 数が 出る）
 *
 *  走らせ方: node tests/data-conn.test.mjs [--self-test]
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
const A = require_(path.join(ROOT, 'lib/adv-filter.js'));
const C = require_(path.join(ROOT, 'lib/connections.js'));

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
ok('tools/measure-data3.ps1 が 在る', fs.existsSync(path.join(ROOT, 'tools/measure-data3.ps1')));

console.log('\n[② ★詳細設定が 実測の 行数に なるか★（4通り）]');
{
  const 元 = [['名前'], ['あ'], ['い'], ['あ'], ['う']];
  const 条 = [['名前'], ['あ']];
  const 数 = (v) => v.length;
  const 中 = (v) => v.map(r => r[0]).join(',');
  ok('★① 条件あり＋重複残す = 3行（実測）★', 数(A.絞る(元, 条, false)) === 3, 中(A.絞る(元, 条, false)));
  ok('  中身は 名前,あ,あ', 中(A.絞る(元, 条, false)) === '名前,あ,あ');
  ok('★② 条件あり＋重複除く = 2行（実測）★', 数(A.絞る(元, 条, true)) === 2, 中(A.絞る(元, 条, true)));
  ok('★③ 条件なし＋重複除く = 4行（実測・測り直した）★',
    数(A.絞る(元, null, true)) === 4, 中(A.絞る(元, null, true)));
  ok('  中身は 名前,あ,い,う', 中(A.絞る(元, null, true)) === '名前,あ,い,う');
  ok('★④ 条件なし＋重複残す = 5行（実測）★', 数(A.絞る(元, null, false)) === 5, 中(A.絞る(元, null, false)));
  ok('★見出しは いつも 付く★', A.絞る(元, [['名前'], ['ない値']], false)[0][0] === '名前');
  ok('★1つも 合わなければ 見出しだけ★', 数(A.絞る(元, [['名前'], ['ない値']], false)) === 1);
  ok('★空の 表なら 空★', 数(A.絞る([], null, false)) === 0);
}

console.log('\n[③ 条件の 書き方（実Excel と 同じ）]');
ok('★>100 は 150 で 通り 50 で 通らない★', A.合うか(150, '>100') === true && A.合うか(50, '>100') === false);
ok('★<=5 は 5 で 通る（等号も 入る）★', A.合うか(5, '<=5') === true && A.合うか(6, '<=5') === false);
ok('★>=5 は 5 で 通る★', A.合うか(5, '>=5') === true && A.合うか(4, '>=5') === false);
ok('★<>あ は い で 通り あ で 通らない★', A.合うか('い', '<>あ') === true && A.合うか('あ', '<>あ') === false);
ok('★=あ は ぴったり だけ★', A.合うか('あ', '=あ') === true && A.合うか('あい', '=あ') === false);
ok('★ふつうの 字は 前が 合えば 通る★', A.合うか('あい', 'あ') === true && A.合うか('いあ', 'あ') === false);
ok('★* が 使える★', A.合うか('あいう', 'あ*') === true && A.合うか('いあう', 'あ*') === false);
ok('★? は 1文字★', A.合うか('あい', 'あ?') === true && A.合うか('あいう', 'あ?') === false);
ok('★空の 条件は 何でも 通る★', A.合うか('なんでも', '') === true);
ok('★同じ 行は「かつ」★', (() => {
  const 元 = [['名', '数'], ['あ', 10], ['あ', 200], ['い', 200]];
  const 出 = A.絞る(元, [['名', '数'], ['あ', '>100']], false);
  return 出.length === 2 && 出[1][1] === 200;
})());
ok('★行を 変えたら「または」★', (() => {
  const 元 = [['名'], ['あ'], ['い'], ['う']];
  const 出 = A.絞る(元, [['名'], ['あ'], ['う']], false);
  return 出.length === 3;
})());
ok('★見出しに 無い 名前の 条件は 合わない★',
  A.絞る([['名'], ['あ']], [['ない列'], ['あ']], false).length === 1);
ok('★条件の 表が 見出しだけなら 条件なし★',
  A.絞る([['名'], ['あ'], ['い']], [['名']], false).length === 3);

console.log('\n[④ クエリと接続]');
{
  const t = [];
  const v = C.足す(t, 'csv', '売上', 'uriage.csv', [[1, 2], [3, 4]]);
  ok('★足せる★', v.名 === '売上' && t.length === 1);
  ok('  行数・列数を 覚える', v.行数 === 2 && v.列数 === 2);
  ok('★同じ 名前は 2つ 作らない★', C.足す(t, 'csv', '売上', 'x.csv', [[1]]).名 === '売上 2');
  ok('  探せる', !!C.探す(t, '売上 2') && C.探す(t, 'ない') === null);
  ok('  消せる', C.消す(t, '売上') === true && t.length === 1);
  ok('  無い 名前は false', C.消す(t, 'ない') === false);
  ok('★この ブックの 表は そのまま 更新 出来る★', C.更新できるか('range') === true);
  ok('★ファイルから 来た 物は 押すだけでは 更新 出来ない★', C.更新できるか('csv') === false);
  ok('  その 訳を 字で 持っている', /勝手に 開けません/.test(C.更新の説明('csv')));
  ok('  この ブックの 表の 説明も 在る', /そのまま 数え直せます/.test(C.更新の説明('range')));
}
{
  /* ★実測＝`C:\ない\ないブック.xlsx`（シート名は 入らない）★ */
  const 式 = ["='C:\\ない\\[ないブック.xlsx]Sheet1'!A1", '=SUM(A1:A2)'];
  const 外 = C.外のブックを探す(式);
  ok('★外の ブックを 1件 見つける★', 外.length === 1, JSON.stringify(外));
  ok('★中身は ブックまで（シート名は 入らない）★', 外[0] === 'C:\\ない\\ないブック.xlsx', 外[0]);
  ok('  ふつうの 式は 数えない', C.外のブックを探す(['=SUM(A1:A2)']).length === 0);
  ok('  同じ ブックを 2回 数えない', C.外のブックを探す([式[0], 式[0]]).length === 1);
  ok('  何も 無ければ 0件', C.外のブックを探す([]).length === 0);
}

console.log('\n[⑤ 画面から 押せる]');
for (const n of ['詳細設定を開く', '詳細設定を実行', '_範囲を配列にする', '_字を範囲に',
  '表から接続を作る', '接続の窓を開く', '接続を更新', '接続を消す', 'すべて更新',
  '接続のプロパティ', 'ブックのリンクを見る', '接続の帯を出す']) {
  ok(n + ' が 在る', !!抜く(n));
}
ok('★つなぎが 0でも すべて更新は 通る（実測と 同じ）★',
  /つなぎは 1件も ありません（何も しませんでした）/.test(book));
ok('★出す 所が 変なら 断る★', /出す 所を A1 の ように 書いてください/.test(book));
ok('★条件の 表が 変なら 断る★', /条件の 表を C1:C2 の ように 書いてください/.test(book));
ok('★#REF! に なる事を 画面に 書いてある★', /#REF!/.test(book));
ok('★勝手に 開けない事を 画面に 書いてある★', /前に 選んだ ファイルを 勝手に 開けません/.test(book));
ok('★帯も 積む 一覧に 入っている★', /'paperBar', 'viewBar', 'winBar', 'connBar'/.test(book));
ok('★alert / prompt / confirm を 使っていない★', !/\balert\(|\bprompt\(|\bconfirm\(/.test(book));
{
  const f = new Function('letterToCol', 抜く('_字を範囲に') + '\nreturn _字を範囲に;')(
    (s) => { let n = 0; s = s.toUpperCase(); for (let i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64); return n - 1; });
  ok('★C1:C2 を 読める★', JSON.stringify(f('C1:C2')) === '{"r1":0,"c1":2,"r2":1,"c2":2}', JSON.stringify(f('C1:C2')));
  ok('★E1 だけでも 読める★', JSON.stringify(f('E1')) === '{"r1":0,"c1":4,"r2":0,"c2":4}', JSON.stringify(f('E1')));
  ok('★AA10 も 読める★', f('AA10').c1 === 26 && f('AA10').r1 === 9, JSON.stringify(f('AA10')));
  ok('★変な字は null★', f('あいう') === null && f('') === null);
  ok('  前後の 空白は 気にしない', JSON.stringify(f('  C1:C2 ')) === JSON.stringify(f('C1:C2')));
}

console.log('\n[⑥ 副題を 決めていない 窓が 増えていないか]');
{
  const 行 = book.split(String.fromCharCode(10));
  const 抜け = [];
  for (let i = 0; i < 行.length; i++) {
    if (!/getElementById\('funcTitle'\)\.textContent =/.test(行[i])) continue;
    if (!/窓の副題\(/.test(行.slice(i, i + 5).join(String.fromCharCode(10)))) 抜け.push(i + 1);
  }
  ok('★副題を 決めていない 窓は 0個★', 抜け.length === 0, 抜け.join(' / '));
}

console.log('\n[⑦ リボンから 押せる]');
{
  const ACT = require_(path.join(ROOT, 'lib/ribbon-actions.js'));
  const g = globalThis, 前w = g.window;
  const 試す = (ボタン, 呼ぶ名) => {
    let 受け = 'よばれていない';
    const w = {}; w[呼ぶ名] = function () { 受け = 'ok'; };
    g.window = w; ACT[ボタン](); g.window = 前w;
    ok('「' + ボタン + '」→ ' + 呼ぶ名, 受け === 'ok', String(受け));
  };
  試す('詳細設定', '詳細設定を開く');
  試す('範囲から', '表から接続を作る');
  試す('すべて更新', 'すべて更新');
  試す('クエリと接続', '接続の窓を開く');
  試す('接続のプロパティ', '接続のプロパティ');
  試す('ブックのリンク', 'ブックのリンクを見る');
  試す('既存の接続', '接続の窓を開く');
}

console.log('\ndata-conn: ' + 緑 + '/' + (緑 + 赤) + ' passed');

if (壊す) {
  console.log('\n★--self-test＝この 見張りが 見ている物を 直に 確かめる（★壊して 赤を 見るのは tools/break-check.mjs★）★');
  let 素通り = 0;
  const 元 = [['名前'], ['あ'], ['い'], ['あ'], ['う']];
  if (A.絞る(元, null, true).length !== 4) { 素通り++; console.log('  ★素通り★ 条件なし＋重複除く が 4行で ない'); }
  else console.log('  ok   条件なし＋重複除く は 4行');
  /* ★測り直す前の 嘘の 数（2行）で 通らない事★を 見る */
  if (A.絞る(元, null, true).length === 2) { 素通り++; console.log('  ★素通り★ 前の 条件が 残った 時の 数（2行）に なっている'); }
  else console.log('  ok   2行（測り直す前の 嘘の 数）には ならない');
  if (!/勝手に 開けません/.test(C.更新の説明('csv'))) { 素通り++; console.log('  ★素通り★ 出来ない事を 言っていない'); }
  else console.log('  ok   出来ない事を 言っている');
  if (C.外のブックを探す(["='C:\\ない\\[ないブック.xlsx]Sheet1'!A1"])[0] !== 'C:\\ない\\ないブック.xlsx') {
    素通り++; console.log('  ★素通り★ 外の ブックの 名前が 実測と 違う');
  } else console.log('  ok   外の ブックの 名前が 実測と 同じ');
  if (素通り) { console.log('★抜け道 ' + 素通り + '件★'); process.exit(1); }
  console.log('  ok   抜け道 0件');
}
process.exit(赤 ? 1 : 0);
