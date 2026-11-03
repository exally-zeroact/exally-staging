/* data-tab2.test.mjs — ★統合・アウトライン・スライサー★ 2026-08-30
 *
 *  ★真値（実Excel 16.0 で 実測）★ … tools/measure-data2.ps1
 *    ・アウトライン … まとめた 行は ★段=2★（外は 1）／★1段だけ 出すと 隠れる★／2段で また 出る
 *      まとめの 行は ★下★（SummaryRow=1＝xlSummaryBelow）／まとめの 列は 右（-4152）
 *    ・統合の まとめ方の 番号 … 合計 -4157／個数 -4112／平均 -4106／最大 -4136／最小 -4139
 *    ・★スライサーは COM から 作れなかった★（Value does not fall within the expected range）
 *      ⇒ ★見た目は 未測定＝うちの決め★。働き（押した 値だけ 見せる）だけ 同じに した。
 *
 *  走らせ方: node tests/data-tab2.test.mjs [--self-test]
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
  let d = 0; const j = book.indexOf('{', i);
  for (let k = j; k < book.length; k++) {
    if (book[k] === '{') d++;
    else if (book[k] === '}') { d--; if (d === 0) return book.slice(i, k + 1); }
  }
  return null;
}

console.log('\n[① 測った 道具が 残っている]');
ok('tools/measure-data2.ps1 が 在る', fs.existsSync(path.join(ROOT, 'tools/measure-data2.ps1')));
ok('★スライサーが 測れなかった事を 書いてある★', /★未測定★＝実Excelの スライサーは COM から 作れず/.test(book));

console.log('\n[② 統合]');
for (const n of ['統合の窓を開く', '統合の中身', '統合の見本', '統合を作る']) ok(n + ' が 在る', !!抜く(n));
ok('窓が 在る', /id="consolOverlay"/.test(book));
ok('★実測の 番号を 書き残してある★', /合計 -4157／個数 -4112／平均 -4106／最大 -4136／最小 -4139/.test(book));
ok('★まとめ方は ピボットの 部品を 使う（同じ物を 2つ 持たない）★',
  /Pivot\.集計たち\[まとめ方\]/.test(book));
ok('★新しい シートに 出す★', /統合の中身[\s\S]{0,900}addSheet\(\)/.test(book));
ok('★作る前に 見本を 見せる★', /統合の見本\(\)/.test(book));
ok('★何行が 何行に なるか 出す★', /行 → ' \+ r\.出来た行 \+ '行/.test(book));
{
  const f = new Function('ピボットの元の表', 'document', 'Pivot',
    抜く('統合の中身') + '\nreturn 統合の中身;');
  const 表 = [['店', '金', '数'], ['東', 100, 1], ['西', 200, 2], ['東', 300, 3]];
  const Pivot = require_(path.join(ROOT, 'lib/pivot.js'));
  const 欄 = { csFrom: { value: 'A1:C4' }, csHeader: { checked: true }, csAgg: { value: '合計' } };
  const r = f(() => 表, { getElementById: (id) => 欄[id] }, Pivot)();
  ok('★同じ 名前を 1行に まとめる★', r.出来た行 === 2, String(r.出来た行));
  ok('★東 = 100+300 = 400★', r.表[1][1] === 400, JSON.stringify(r.表[1]));
  ok('★西 = 200★', r.表[2][1] === 200, JSON.stringify(r.表[2]));
  ok('★見出しは そのまま 残る★', r.表[0].join(',') === '店,金,数', r.表[0].join(','));
  欄.csAgg.value = '平均';
  const r2 = f(() => 表, { getElementById: (id) => 欄[id] }, Pivot)();
  ok('平均に すると 東 = 200', r2.表[1][1] === 200, JSON.stringify(r2.表[1]));
}

console.log('\n[③ アウトライン（実測＝段2／1段で 隠れる）]');
for (const n of ['段の箱', 'グループ化', 'グループ解除', '段を出す', 'アウトラインの窓を開く']) {
  ok(n + ' が 在る', !!抜く(n));
}
ok('★段は 1つずつ 上がる（外は 1・まとめると 2）★', /Math\.min\(8, \(箱\[r\] \|\| 1\) \+ 1\)/.test(book));
ok('★1段だけ 出すと それより 深い 行は 隠れる★', /\(箱\[鍵\[i\]\] \|\| 1\) > 段/.test(book));
ok('★外したら 隠れたままに しない★', /解除したら 隠れたままに しない/.test(book));
{
  const f = new Function('sheets', 'activeSheet', 'hideCtxMenu', 'render', 'notify',
    'selR1', 'selR2', 抜く('段の箱') + '\n' + 抜く('グループ化') + '\n' + 抜く('段を出す')
    + '\nreturn { まとめる: グループ化, 出す: 段を出す, 箱: 段の箱 };');
  const s0 = { outline: {}, hiddenRows: {} };
  const api = f([s0], 0, () => {}, () => {}, () => {}, 4, 6);
  api.まとめる();
  ok('★まとめた 行は 段2（実測）★', s0.outline[4] === 2 && s0.outline[6] === 2, JSON.stringify(s0.outline));
  api.出す(1);
  ok('★1段だけ 出すと 隠れる（実測）★', s0.hiddenRows[4] === true, JSON.stringify(s0.hiddenRows));
  api.出す(2);
  ok('★2段で また 出る（実測）★', !s0.hiddenRows[4], JSON.stringify(s0.hiddenRows));
}

console.log('\n[④ スライサー（働きは 同じ・見た目は うちの決め）]');
for (const n of ['スライサーを開く', 'スライサーを描く', 'スライサーを押す',
  'スライサーを全部に', 'スライサーを効かせる', 'スライサーを消す']) ok(n + ' が 在る', !!抜く(n));
ok('板が 在る', /id="slicerPanel"/.test(book));
ok('★テーブルで なければ 断る★', /先に テーブルに してください/.test(book));
ok('★絞りは テーブルの ▼ と 同じ 道（GridFilter.byValues）★',
  /スライサーを効かせる[\s\S]{0,600}GridFilter\.byValues/.test(book));
ok('★1つも 選ばない を 止める★', /1つも 選ばないと 全部 消えてしまいます/.test(book));
ok('★再適用も 効くように 覚える★', /前の絞り = \{ 表: d\.表, 列: d\.列, 残す: d\.選び \}/.test(book));
ok('★alert / prompt / confirm を 使っていない★', !/\balert\(|\bprompt\(|\bconfirm\(/.test(book));

console.log('\n[⑤ リボンから 押せる]');
{
  const ACT = require_(path.join(ROOT, 'lib/ribbon-actions.js'));
  const g = globalThis, 前w = g.window;
  for (const [ボタン, 呼ぶ名] of [
    ['統合', '統合の窓を開く'], ['アウトライン', 'アウトラインの窓を開く'], ['スライサー', 'スライサーを開く'],
  ]) {
    let 受け = null;
    const w = {}; w[呼ぶ名] = function () { 受け = 'ok'; };
    g.window = w;
    ACT[ボタン]();
    g.window = 前w;
    ok('「' + ボタン + '」→ ' + 呼ぶ名, 受け === 'ok', String(受け));
  }
}

console.log('\ndata-tab2: ' + 緑 + '/' + (緑 + 赤) + ' passed');

if (壊す) {
  console.log('\n★--self-test＝この 見張りが 見ている物を 直に 確かめる（★壊して 赤を 見るのは tools/break-check.mjs★）★');
  let 素通り = 0;
  if (!/\(箱\[鍵\[i\]\] \|\| 1\) > 段/.test(book)) { 素通り++; console.log('  ★素通り★ 段で 隠す所が 無い'); }
  else console.log('  ok   段で 隠している');
  if (!/1つも 選ばないと 全部 消えてしまいます/.test(book)) { 素通り++; console.log('  ★素通り★ 全部 消える道を 止めていない'); }
  else console.log('  ok   全部 消える道を 止めている');
  if (!/Pivot\.集計たち\[まとめ方\]/.test(book)) { 素通り++; console.log('  ★素通り★ まとめ方を 写している'); }
  else console.log('  ok   まとめ方は 部品を 使っている');
  if (素通り) { console.log('★抜け道 ' + 素通り + '件★'); process.exit(1); }
  console.log('  ok   抜け道 0件');
}
process.exit(赤 ? 1 : 0);
