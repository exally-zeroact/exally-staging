/* view-toggle.test.mjs — ★表示の 切り替え（見出し・数式バー・枠線）★ 2026-08-29
 *
 *  ★真値（実Excel 16.0 で 実測）★
 *    ブックの表示 … 1=標準（既定）／ズーム 100
 *    ★画面の枠線 True・見出し True・数式バー True・数式そのまま False★
 *
 *  ★実ブラウザ（Playwright）で 押して 確かめた★
 *    見出し … HDR_W/HDR_H が 46/22 ↔ ★0/0★
 *    数式バー … 出ている ↔ 消えている
 *
 *  走らせ方: node tests/view-toggle.test.mjs [--self-test]
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

console.log('\n[① 既定が 実Excelと 同じ]');
ok('★見出しは 出す★（実Excel True）', /window\.見出しを出す = true;/.test(book));
ok('★枠線は 出す★（実Excel True）', /var 枠線を出す = true;/.test(book));
ok('★数式の そのまま表示は 消す★（実Excel False）', /var 数式を表示 = false;/.test(book));

console.log('\n[② 画面に 在る]');
for (const n of ['見出しを出すか', '数式バーを出すか', '枠線の表示を切り替える']) ok(n + ' が 在る', !!抜く(n));
ok('★元の数を 残している（戻せる）★', /var HDR_W0 = HDR_W, HDR_H0 = HDR_H;/.test(book));
ok('★resize が 見出しを 0 に する★', /HDR_W = \(window\.見出しを出す === false\) \? 0 : HDR_W0;/.test(book));

console.log('\n[③ 実際に 走らせる]');
{
  const 出た = [];
  const w = { 見出しを出す: true };
  const f = new Function('window', 'resize', 'render', 'notify',
    抜く('見出しを出すか') + '\nreturn 見出しを出すか;');
  const 押す = f(w, function () {}, function () {}, function (m) { 出た.push(m); });
  ok('★1回目で 消す★', 押す() === false && w.見出しを出す === false, String(w.見出しを出す));
  ok('★2回目で 戻る★', 押す() === true && w.見出しを出す === true, String(w.見出しを出す));
  ok('★どちらも 言う★', 出た.length === 2 && /消す/.test(出た[0]) && /出す/.test(出た[1]), JSON.stringify(出た));
}
{
  const 出た = [];
  const el = { style: { display: '' } };
  const f = new Function('document', 'resize', 'render', 'notify',
    抜く('数式バーを出すか') + '\nreturn 数式バーを出すか;');
  const 押す = f({ getElementById: () => el }, function () {}, function () {}, function (m) { 出た.push(m); });
  ok('★1回目で 消える★', 押す() === false && el.style.display === 'none', el.style.display);
  ok('★2回目で 出る★', 押す() === true && el.style.display === '', JSON.stringify(el.style.display));
}
{
  /* 数式バーが 無い時＝★黙らずに 言う★ */
  const 出た = [];
  const f = new Function('document', 'resize', 'render', 'notify',
    抜く('数式バーを出すか') + '\nreturn 数式バーを出すか;');
  const r = f({ getElementById: () => null }, function () {}, function () {}, function (m) { 出た.push(m); })();
  ok('★無い時は null★', r === null, String(r));
  ok('★そう言う★', /見つかりません/.test(出た.join('')), JSON.stringify(出た));
}

console.log('\n[④ リボンから 押せる]');
const ACT = require_(path.join(ROOT, 'lib/ribbon-actions.js'));
for (const [名, 行き先] of [['見出しを表示', '見出しを出すか'], ['数式バー', '数式バーを出すか'],
  ['枠線を表示', '枠線の表示を切り替える']]) {
  const g = globalThis, 前w = g.window;
  let 受け = null;
  g.window = {}; g.window[行き先] = function () { 受け = 行き先; };
  ACT[名]();
  g.window = 前w;
  ok('「' + 名 + '」→ ' + 行き先, 受け === 行き先, String(受け));
}

console.log('\nview-toggle: ' + 緑 + '/' + (緑 + 赤) + ' passed');

if (壊す) {
  console.log('\n★--self-test＝この 見張りが 見ている物を 直に 確かめる（★壊して 赤を 見るのは tools/break-check.mjs★）★');
  let 素通り = 0;
  /* 壊し① 元の数を 残さない（戻せなくなる） */
  if (!/var HDR_W0 = HDR_W/.test(book)) { 素通り++; console.log('  ★素通り★ 元の数を 残していない'); }
  /* 壊し② 既定が 実Excelと 違う */
  if (/window\.見出しを出す = false;/.test(book)) { 素通り++; console.log('  ★素通り★ 既定で 見出しを 消している'); }
  /* 壊し③ 黙って 何も しない（無い時に 言わない） */
  const f = new Function('document', 'resize', 'render', 'notify',
    抜く('数式バーを出すか') + '\nreturn 数式バーを出すか;');
  let 言った = 0;
  f({ getElementById: () => null }, function () {}, function () {}, function () { 言った++; })();
  if (!言った) { 素通り++; console.log('  ★素通り★ 無い時に 黙っている'); }
  if (素通り) { console.log('★抜け道 ' + 素通り + '件★'); process.exit(1); }
  console.log('  ok   抜け道 0件');
}
process.exit(赤 ? 1 : 0);
