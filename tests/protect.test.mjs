/* protect.test.mjs — ★シートの保護（校閲タブ）★ 2026-08-29
 *
 *  ★真値（実Excel 16.0 で 実測）★
 *    ・新しいセルは ★全部 ロック（Locked=True）★
 *    ・★ロックは 保護して 初めて 効く★（保護していない間は どこでも 書ける）
 *    ・保護中は ロックされたセルに ★書けない★／ロックを 外したセルには ★書ける★
 *
 *  ★合言葉は 掛けない★＝忘れたら 二度と 開けない。
 *  Excelの 合言葉も 守りとしては 弱い（外す道具が 出回っている）＝
 *  「うっかり 直すのを 防ぐ」ための 物と 割り切る。★紙にも そう書く★。
 *
 *  走らせ方: node tests/protect.test.mjs [--self-test]
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
for (const n of ['このシートは保護中', 'セルはロックされているか', 'ここに書けるか',
  'シートを保護する', 'シートの保護をやめる', '選んだ所のロックを切り替える']) ok(n + ' が 在る', !!抜く(n));

console.log('\n[② 実Excelと 同じ 決まり]');
function 台(データ, 保護, 選 = [0, 0, 0, 0]) {
  const s = { data: {}, protect: !!保護 };
  for (const k of Object.keys(データ)) s.data[k] = データ[k];
  const 出た = [];
  const f = new Function('sheets', 'activeSheet', 'selR1', 'selR2', 'selC1', 'selC2', 'render', 'updateBar', 'notify',
    ['このシートは保護中', 'セルはロックされているか', 'ここに書けるか', 'シートを保護する',
      'シートの保護をやめる', 'シートの保護を切り替える', '選んだ所のロックを切り替える'].map(抜く).join('\n')
    + '\nreturn { 保護中: このシートは保護中, ロック: セルはロックされているか, 書ける: ここに書けるか,'
    + ' 掛ける: シートを保護する, やめる: シートの保護をやめる, 切替: シートの保護を切り替える,'
    + ' ロック切替: 選んだ所のロックを切り替える };');
  const api = f([s], 0, 選[0], 選[1], 選[2], 選[3], function () {}, function () {}, function (m) { 出た.push(m); });
  return { s, api, 出た };
}
{
  const t = 台({}, false);
  ok('★新しいセルは ロック（実Excelと 同じ）★', t.api.ロック(0, 0) === true);
  ok('★保護していなければ どこでも 書ける★', t.api.書ける(0, 0) === true);
}
{
  const t = 台({}, true);
  ok('★保護中は ロックされたセルに 書けない★', t.api.書ける(0, 0) === false);
}
{
  const t = 台({ '0,0': { v: '', locked: false } }, true);
  ok('★ロックを 外したセルには 保護中でも 書ける★', t.api.書ける(0, 0) === true);
}

console.log('\n[③ 掛ける／やめる］');
{
  const t = 台({}, false);
  ok('掛けると 保護中', t.api.掛ける() === true && t.s.protect === true);
  ok('★掛けたと 言う★', /保護しました/.test(t.出た.join('')), JSON.stringify(t.出た));
  ok('やめると 外れる', t.api.やめる() === false && t.s.protect === false);
  ok('★やめたと 言う★', /やめました/.test(t.出た.join('')), JSON.stringify(t.出た));
  ok('切り替えで 入る', t.api.切替() === true);
  ok('切り替えで 戻る', t.api.切替() === false);
}

console.log('\n[④ セルの ロックを 外す／掛ける]');
{
  const t = 台({}, false, [0, 1, 0, 1]);   /* 2行×2列 */
  const 外した = t.api.ロック切替();
  ok('★今 ロックなら 外す★', 外した === true);
  ok('4個 全部 外れた',
    ['0,0', '0,1', '1,0', '1,1'].every((k) => t.s.data[k] && t.s.data[k].locked === false),
    JSON.stringify(Object.keys(t.s.data)));
  ok('★何個 変えたかを 言う★', /4個/.test(t.出た.join('')), JSON.stringify(t.出た));
  const 掛けた = t.api.ロック切替();
  ok('★もう1回で 掛け直す★', 掛けた === false && t.s.data['0,0'].locked === true);
}

console.log('\n[⑤ ★書き込みが 本当に 止まる★]');
ok('★setCell の 入口で 断っている★', /if\(typeof ここに書けるか === 'function' && !ここに書けるか\(r,c\)\)/.test(book));
ok('★黙って 止めない（なぜかを 出す）★', /このシートは 保護中です/.test(book));
ok('★知らせが 出っぱなしに ならない（1.5秒で 戻す）★', /_保護の知らせ = false/.test(book));

console.log('\n[⑥ リボンから 押せる]');
const ACT = require_(path.join(ROOT, 'lib/ribbon-actions.js'));
for (const [名, 行き先] of [['シートの保護', 'シートの保護を切り替える'], ['セルのロック', '選んだ所のロックを切り替える']]) {
  const g = globalThis, 前w = g.window;
  let 受け = null;
  g.window = {}; g.window[行き先] = function () { 受け = 行き先; };
  ACT[名]();
  g.window = 前w;
  ok('「' + 名 + '」→ ' + 行き先, 受け === 行き先, String(受け));
}

console.log('\nprotect: ' + 緑 + '/' + (緑 + 赤) + ' passed');

if (壊す) {
  console.log('\n★--self-test＝この 見張りが 見ている物を 直に 確かめる（★壊して 赤を 見るのは tools/break-check.mjs★）★');
  let 素通り = 0;
  /* 壊し① 既定を「ロックされていない」に する（実Excelと 違う） */
  const t = 台({}, true);
  if (t.api.ロック(9, 9) === false) { 素通り++; console.log('  ★素通り★ 既定が ロックでない'); }
  /* 壊し② 保護中でも 書けてしまう */
  if (t.api.書ける(9, 9) === true) { 素通り++; console.log('  ★素通り★ 保護中なのに 書ける'); }
  /* 壊し③ setCell が 見ていない */
  if (!/ここに書けるか\(r,c\)/.test(book)) { 素通り++; console.log('  ★素通り★ setCell が 見ていない'); }
  if (素通り) { console.log('★抜け道 ' + 素通り + '件★'); process.exit(1); }
  console.log('  ok   抜け道 0件');
}
process.exit(赤 ? 1 : 0);
