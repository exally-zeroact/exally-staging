/* page-setup.test.mjs — ★ページ設定（紙に 刷る時の 決まり）★ 2026-08-29
 *
 *  ★真値（実Excel 16.0 を COM で 実測）★
 *    向き … 1（縦）／用紙 … 9（A4）／倍率 … 100
 *    余白 … 上下 54pt（★19.1mm★）／左右 50.4pt（★17.8mm★）
 *    ★枠線を 印刷 … False★／★見出しを 印刷 … False★
 *    中央に 寄せる … 横も 縦も False
 *
 *  走らせ方: node tests/page-setup.test.mjs [--self-test]
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
const GP = require_(path.join(ROOT, 'lib/grid-print.js'));
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
ok('★既定は 縦★', /var ページ設定 = \{ 向き: 'portrait'/.test(book));
ok('★既定は 枠線を 刷らない★', /枠線: false/.test(book));
ok('★既定は 見出しを 刷らない★', /行列番号: false/.test(book));

console.log('\n[② 紙の 組み立てが 実測の 余白]');
{
  const html = GP.buildHtml({ data: { '0,0': { v: 'あ', d: 'あ' } }, sheetName: 'S' });
  ok('紙が 出来た', !!html);
  ok('★A4★', /size: A4/.test(html || ''));
  /* ★2026-08-30★ 余白を 4方向べつべつに 出せるように したので 4つ書きに なった。
     数は ★丸めない 実測値★（54pt÷28.3464566929134=1.905 ／ 50.4pt÷…=1.778）。 */
  ok('★余白 1.905cm / 1.778cm（＝54pt / 50.4pt）★',
    /margin: 1\.905cm 1\.778cm 1\.905cm 1\.778cm/.test(html || ''),
    (String(html).match(/margin:[^;]*/) || [''])[0]);
  ok('★既定は 縦★', /A4 portrait/.test(html || ''));
}
{
  const html = GP.buildHtml({ data: { '0,0': { v: 'あ', d: 'あ' } }, sheetName: 'S', 向き: 'landscape' });
  ok('★横に できる★', /A4 landscape/.test(html || ''));
}

console.log('\n[③ 切り替えが 効く]');
{
  const f = new Function('notify',
    "var ページ設定 = { 向き: 'portrait', 枠線: false, 行列番号: false };\n"
    + 抜く('向きを切り替える') + '\n' + 抜く('枠線を刷るか') + '\n' + 抜く('見出しを刷るか')
    + '\nreturn { 向き: 向きを切り替える, 枠: 枠線を刷るか, 見出し: 見出しを刷るか, 今: function(){ return ページ設定; } };');
  const 出た = [];
  const api = f(function (m) { 出た.push(m); });
  ok('★1回目で 横★', api.向き() === 'landscape', String(api.向き));
  ok('★2回目で 縦に 戻る★', api.向き() === 'portrait');
  ok('★向きを 言う★', /横|縦/.test(出た.join('')), JSON.stringify(出た.slice(0, 2)));
  ok('枠線 … 押すと 刷る', api.枠() === true);
  ok('枠線 … もう1回で 戻る', api.枠() === false);
  ok('見出し … 押すと 刷る', api.見出し() === true);
  /* 押した回数＝向き2 + 枠線2 + 見出し1 ＝ ★5回★（数え間違いを 直した 2026-08-29） */
  ok('★変えたら 必ず 言う（黙って 変えない）★', 出た.length === 5, String(出た.length));
}

console.log('\n[④ 印刷が ページ設定を 見ている]');
ok('★printSheet が 向きを 渡している★', /向き: ページ設定\.向き/.test(book));
ok('★枠線も 渡している★', /枠線: ページ設定\.枠線/.test(book));
ok('★見出しも 渡している★', /行列番号: ページ設定\.行列番号/.test(book));
ok('窓が 在る', /id="pageOverlay"/.test(book));
ok('★alert / prompt / confirm を 使っていない★', !/\balert\(|\bprompt\(|\bconfirm\(/.test(book));

console.log('\n[⑤ リボンから 押せる]');
const ACT = require_(path.join(ROOT, 'lib/ribbon-actions.js'));
for (const [名, 行き先] of [['ページ設定', 'ページ設定を開く'], ['印刷の向き', '向きを切り替える'],
  ['枠線を刷る', '枠線を刷るか'], ['見出しを刷る', '見出しを刷るか'], ['印刷', 'printSheet']]) {
  const g = globalThis, 前w = g.window;
  let 受け = null;
  g.window = {}; g.window[行き先] = function () { 受け = 行き先; };
  ACT[名]();
  g.window = 前w;
  ok('「' + 名 + '」→ ' + 行き先, 受け === 行き先, String(受け));
}

console.log('\npage-setup: ' + 緑 + '/' + (緑 + 赤) + ' passed');

if (壊す) {
  console.log('\n★--self-test＝わざと 壊して 赤に なるか★');
  let 素通り = 0;
  /* 壊し① 余白が 実測と 違う */
  const にせ = '@page { size: A4 portrait; margin: 2cm 2cm; }';
  if (/margin: 1\.9cm 1\.78cm/.test(にせ)) { 素通り++; console.log('  ★素通り★ 壊し方が おかしい'); }
  /* 壊し② 既定で 枠線を 刷る（実Excelと 違う） */
  if (/var ページ設定 = \{ 向き: 'portrait', 枠線: true/.test(book)) {
    素通り++; console.log('  ★素通り★ 既定で 枠線を 刷る形に なっている');
  }
  /* 壊し③ 向きを 渡していない */
  if (!/向き: ページ設定\.向き/.test(book)) { 素通り++; console.log('  ★素通り★ 向きを 渡していない'); }
  if (素通り) { console.log('★抜け道 ' + 素通り + '件★'); process.exit(1); }
  console.log('  ok   抜け道 0件');
}
process.exit(赤 ? 1 : 0);
