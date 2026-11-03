/* forecast.test.mjs — ★予測シート・メモ・パフォーマンス点検★ 2026-08-30
 *
 *  ★真値（実Excel 16.0 で 実測）★ … tools/measure-forecast.ps1
 *    1..6 に 100,120,140,160,180,200
 *      FORECAST.LINEAR(7) = ★220★ ／ SLOPE = ★20★ ／ INTERCEPT = ★80★
 *    でこぼこ 100,90,130,120,160,150
 *      予測(7) = ★170★ ／ 傾き = ★12.857142857142858★ ／ 切片 = ★80★
 *    メモ … 作者つき（Excelの ユーザー名）／★既定は 見えない★／形 93 x 54.8
 *
 *  ★未測定★ 実Excelの「予測シート」は 季節の くり返し（ETS）も 見る。うちは まだ 直線だけ。
 *  ★未測定★ 「パフォーマンスをチェック」の 窓の 中身は COM から 読めない。
 *
 *  走らせ方: node tests/forecast.test.mjs [--self-test]
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
const Pivot = require_(path.join(ROOT, 'lib/pivot.js'));

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
ok('tools/measure-forecast.ps1 が 在る', fs.existsSync(path.join(ROOT, 'tools/measure-forecast.ps1')));
ok('★季節の くり返しは まだ と 書いてある（画面にも）★',
  /季節の くり返しも 見ます（ETS）。うちは まだ 直線だけです/.test(book)
  && /季節の くり返し（ETS）も 見る★＝うちは まだ 直線だけ/.test(book));
ok('★点検の 中身が 未測定と 書いてある★', /実Excelの 窓の 中身は 外から 読めないため/.test(book));

console.log('\n[② ★実測の 数と 同じ 予測が 出るか★]');
{
  const f = new Function('Pivot', 抜く('直線を当てる') + '\nreturn 直線を当てる;')(Pivot);
  const 真 = f([100, 120, 140, 160, 180, 200]);
  ok('★傾き = 20（実測）★', Math.abs(真.傾き - 20) < 1e-9, String(真.傾き));
  ok('★切片 = 80（実測）★', Math.abs(真.切片 - 80) < 1e-9, String(真.切片));
  ok('★7つ目 = 220（実測 FORECAST.LINEAR）★',
    Math.abs(真.傾き * 7 + 真.切片 - 220) < 1e-9, String(真.傾き * 7 + 真.切片));
  const 凸 = f([100, 90, 130, 120, 160, 150]);
  ok('★でこぼこの 傾き = 12.857142857142858（実測）★',
    Math.abs(凸.傾き - 12.857142857142858) < 1e-9, String(凸.傾き));
  ok('★でこぼこの 切片 = 80（実測）★', Math.abs(凸.切片 - 80) < 1e-9, String(凸.切片));
  ok('★でこぼこの 7つ目 = 170（実測）★',
    Math.abs(凸.傾き * 7 + 凸.切片 - 170) < 1e-9, String(凸.傾き * 7 + 凸.切片));
  ok('★空を とばす（数だけ 見る）★', f([100, '', 140]).個数 === 2, JSON.stringify(f([100, '', 140])));
  ok('★数が 1つなら null（当てずっぽうを 言わない）★', f([100]) === null);
  ok('★数が 無ければ null★', f(['あ', 'い']) === null);
  ok('★けた区切りも 読む★', Math.abs(f(['1,000', '2,000']).傾き - 1000) < 1e-9,
    JSON.stringify(f(['1,000', '2,000'])));
}

console.log('\n[③ 画面に つながっている（予測）]');
for (const n of ['直線を当てる', '予測の元', '予測の窓を開く', '予測の中身', '予測の見本', '予測を作る']) {
  ok(n + ' が 在る', !!抜く(n));
}
ok('窓が 在る', /id="forecastOverlay"/.test(book));
ok('★新しい シートに 出す★', /予測を作る[\s\S]{0,400}addSheet\(\)/.test(book));
ok('★作る前に 見本を 見せる★', /予測の見本\(\)/.test(book));
ok('★数が 足りない時は 断る★', /数が 2つ 以上 要ります/.test(book));
ok('★この先の 数は 1〜100★', /Math\.max\(1, Math\.min\(100,/.test(book));

console.log('\n[④ パフォーマンス点検]');
ok('重さを調べる が 在る', !!抜く('重さを調べる'));
ok('★書式だけの セルを 数える★', /中身が 無いのに 書式だけ 付いたセル/.test(book));
ok('★書式の 一覧を 使い回す（写していない）★', /書式のキー\[k\]/.test(抜く('重さを調べる')));
ok('★見つからない時の 言い方も 在る★', /重くする 物は 見つかりませんでした/.test(book));
ok('★何を 数えたか 出す★', /数えたのは 上の 5つだけ/.test(book));

console.log('\n[⑤ メモ（実Excelの 旧コメント）]');
ok('★コメントと 同じ 窓を 使う（同じ物を 2つ 持たない）★',
  /メモ:\s*function \(\) \{ return 呼ぶ\('コメントの窓を開く'\); \}/
    .test(fs.readFileSync(path.join(ROOT, 'lib/ribbon-actions.js'), 'utf8')));
ok('★alert / prompt / confirm を 使っていない★', !/\balert\(|\bprompt\(|\bconfirm\(/.test(book));

console.log('\n[⑥ リボンから 押せる]');
{
  const ACT = require_(path.join(ROOT, 'lib/ribbon-actions.js'));
  const g = globalThis, 前w = g.window;
  for (const [ボタン, 呼ぶ名] of [
    ['予測シート', '予測の窓を開く'], ['メモ', 'コメントの窓を開く'],
    ['パフォーマンスをチェック', '重さを調べる'],
  ]) {
    let 受け = null;
    const w = {}; w[呼ぶ名] = function () { 受け = 'ok'; };
    g.window = w;
    ACT[ボタン]();
    g.window = 前w;
    ok('「' + ボタン + '」→ ' + 呼ぶ名, 受け === 'ok', String(受け));
  }
}

console.log('\nforecast: ' + 緑 + '/' + (緑 + 赤) + ' passed');

if (壊す) {
  console.log('\n★--self-test＝この 見張りが 見ている物を 直に 確かめる（★壊して 赤を 見るのは tools/break-check.mjs★）★');
  let 素通り = 0;
  const f = new Function('Pivot', 抜く('直線を当てる') + '\nreturn 直線を当てる;')(Pivot);
  /* 壊し① 平均で ごまかしていないか（傾きが 0 に なる） */
  const 真 = f([100, 120, 140, 160, 180, 200]);
  if (Math.abs(真.傾き) < 1e-9) { 素通り++; console.log('  ★素通り★ 傾きが 0＝直線に なっていない'); }
  else console.log('  ok   傾き ' + 真.傾き);
  /* 壊し② 1つでも 答えを 出していないか */
  if (f([100]) !== null) { 素通り++; console.log('  ★素通り★ 数1つで 予測を 出した'); }
  else console.log('  ok   数1つでは 出さない');
  /* 壊し③ 実測の 数と ずれていないか */
  if (Math.abs(真.傾き * 7 + 真.切片 - 220) > 1e-9) { 素通り++; console.log('  ★素通り★ 7つ目が 220 で ない'); }
  else console.log('  ok   7つ目は 220');
  if (素通り) { console.log('★抜け道 ' + 素通り + '件★'); process.exit(1); }
  console.log('  ok   抜け道 0件');
}
process.exit(赤 ? 1 : 0);
