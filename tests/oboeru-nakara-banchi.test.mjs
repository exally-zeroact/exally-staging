/* oboeru-nakara-banchi.test.mjs — ★覚えた 番地を 誰も 書き換えないか★ 2026-09-25
 *
 *  ★★なぜ 在るか★★
 *    `lib/shiki-hyou.js` の `名から番地(名)` は 2026-09-25 から ★覚えて おく★ 形に なりました。
 *    ＝★同じ 字なら 同じ 物を 返します★（前は 毎回 新しく 作って いた）
 *    ⇒★もし 誰かが 返り物を 書き換えると★
 *      ★覚えた 物が 汚れて、次に その 字を 使った 別の マスが 狂います★
 *      ＝★狂うのは 書き換えた 所では ありません★＝★一番 見つけにくい 形★
 *
 *  ★★なぜ そこまでして 覚えるのか（経営者1 の 実測）★★
 *    司さんの 実物（板 15枚・マス 35,760・式 15,799）を 開くと
 *      `名から番地` ... ★141,920,112 回★ 呼ばれて ★別々の 字は 53,365 種類★
 *      ＝★同じ 字を 平均 2,659 回 読み直して いた★
 *    覚えた 結果 ... ★59,471 ms ⇒ 32,078 ms（54%）★（各 3回・中ほど・揺れ幅 20,940 ms）
 *    ★答えは 9回 とも 完全一致★（数の 足し算 54241061.144673）
 *
 *  ★★この 見張りが 見る 物★★
 *    ⑴★返り物を 書き換えて いる 所が 0件★（字で 数える）
 *    ⑵★同じ 字は 同じ 物が 返る★（覚えが 効いて いる）
 *    ⑶★違う 字は 違う 物★（覚えが 混ざって いない）
 *    ⑷★`null` も 覚える★（番地で ない 字を 毎回 読み直さない）
 *    ⑸★上限で 空に しても 答えは 変わらない★
 *
 *  ★字の 形を 見る 門で よい 訳★
 *    普段は 避けますが、ここで 守りたいのは ★「書き換えて いないか」★ そのもの です。
 *    ⇒★字の 形が 中身★（`tests/machi-no-kime.test.mjs` と 同じ 立て方）
 *
 *  走らせ方: node tests/oboeru-nakara-banchi.test.mjs [--self-test]
 */
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath, pathToFileURL } from 'node:url';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(pathToFileURL(path.join(ROOT, 'package.json')));
const H = require_(path.join(ROOT, 'lib/shiki-hyou.js'));
const 壊す = process.argv.includes('--self-test');

let pass = 0, fail = 0;
const T = (n, ok, m) => {
  if (ok) { pass++; console.log('  ok   ' + n); }
  else { fail++; console.log('  NG   ' + n + (m ? '\n       ' + m : '')); }
};

/** ★`名から番地` の 返り物を 書き換えて いる 所を 拾う★
 *  ①`名から番地(･･･).板 = ･･･` の ような 直の 書き換え
 *  ②`var ば = 名から番地(･･･)` と 受けた 名前に 対する `ば.板 = ･･･`
 */
export function 書き換えを拾う(字) {
  const 出 = [];
  /* ① 直に */
  const 直 = /名から番地\([^)]*\)\s*\.\s*(板|行|列)\s*(=[^=]|\+\+|--|\+=|-=)/g;
  let m;
  while ((m = 直.exec(字)) !== null) {
    出.push({ 形: '直に', 印: m[1], 行: 字.slice(0, m.index).split('\n').length });
  }
  /* ② 受けた 名前 */
  const 受け = /(?:var|let|const)\s+([^\s=]+)\s*=\s*名から番地\(/g;
  const 名前 = new Set();
  while ((m = 受け.exec(字)) !== null) 名前.add(m[1]);
  /* `var a = 名から番地(･･･), b = 名から番地(･･･)` の 2つ目も 拾う */
  const 続き = /,\s*([^\s=,]+)\s*=\s*名から番地\(/g;
  while ((m = 続き.exec(字)) !== null) 名前.add(m[1]);
  for (const な of 名前) {
    if (!な || /[^0-9A-Za-z_$぀-ヿ一-鿿]/.test(な)) continue;
    const re = new RegExp(な.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
      + '\\s*\\.\\s*(板|行|列)\\s*(=[^=]|\\+\\+|--|\\+=|-=)', 'g');
    while ((m = re.exec(字)) !== null) {
      出.push({ 形: '受けた名 ' + な, 印: m[1], 行: 字.slice(0, m.index).split('\n').length });
    }
  }
  return 出;
}

console.log('[oboeru-nakara-banchi] ★覚えた 番地を 誰も 書き換えないか★');

/* ══ ⑴ 字で 数える ══ */
const 見る所 = [];
for (const d of ['lib', 'js']) {
  const 所 = path.join(ROOT, d);
  if (!fs.existsSync(所)) continue;
  for (const 名 of fs.readdirSync(所)) if (/\.js$/.test(名)) 見る所.push(path.join(d, 名));
}
for (const f of ['book.html', 'exally-formula.js']) {
  if (fs.existsSync(path.join(ROOT, f))) 見る所.push(f);
}
const 呼ぶ本 = 見る所.filter((f) => fs.readFileSync(path.join(ROOT, f), 'utf8').indexOf('名から番地(') >= 0);
const 悪い = [];
for (const f of 呼ぶ本) {
  for (const x of 書き換えを拾う(fs.readFileSync(path.join(ROOT, f), 'utf8'))) {
    悪い.push(f + ':' + x.行 + '  ' + x.形 + ' → .' + x.印);
  }
}
console.log('      ── 実測 ── 見た ' + 見る所.length + '本 ／ `名から番地` を 呼ぶ ' + 呼ぶ本.length + '本');
T('★★返り物を 書き換えて いる 所が 0件★★（★書き換えると 別の マスが 狂います★）',
  悪い.length === 0, 悪い.join('\n       '));
T('★見る 所が 空に なって いない★（門が 何も 見なく なるのが 一番 こわい）',
  呼ぶ本.length > 0, '呼ぶ本 ' + 呼ぶ本.length + '本');

/* ══ ⑵⑶⑷ 覚えが 効いて いるか ══ */
const a1 = H.名から番地('A1');
const a1b = H.名から番地('A1');
T('★同じ 字は 同じ 物が 返る★（覚えが 効いて いる）', a1 === a1b,
  JSON.stringify(a1) + ' / ' + JSON.stringify(a1b));
const b2 = H.名から番地('B2');
T('★違う 字は 違う 物★（覚えが 混ざって いない）',
  b2 !== a1 && b2.行 === 1 && b2.列 === 1, JSON.stringify(b2));
T('★板つきも 読める★（`\'5月\'!E14`）', (function () {
  const x = H.名から番地("'5月'!E14");
  return x && x.板 === '5月' && x.行 === 13 && x.列 === 4;
}()), JSON.stringify(H.名から番地("'5月'!E14")));
T('★番地で ない 字は `null`★（それも 覚える）',
  H.名から番地('SUM') === null && H.名から番地('SUM') === null);

/* ══ ⑸ わざと 壊して 赤に なるか ══ */
if (壊す) {
  console.log('\n★わざと 壊して 赤に なるか★');
  const W = '名から' + '番地';
  const 組 = [
    ['直に 書き換える（★悪い★）', 'if (' + W + '(n).板 = 1) {}', 1],
    ['受けた 名前を 書き換える（★悪い★）',
      'var ば = ' + W + '(n);\nば.行 = 0;', 1],
    ['受けた 名前の 列を 増やす（★悪い★）',
      'var ば = ' + W + '(n);\nば.列++;', 1],
    ['読むだけ（良い）', 'var ば = ' + W + '(n);\nif (ば.行 === 0) return 1;', 0],
    ['比べるだけ（良い）', 'var a = ' + W + '(x), b = ' + W + '(y);\nif (a.列 === b.列) return 1;', 0],
    ['★別の 物の 同じ 名前（良い）★', 'var 表 = {};\n表.行 = 3;', 0],
    ['2つ 受けて 片方を 書き換える（★悪い★）',
      'var a = ' + W + '(x), b = ' + W + '(y);\nb.板 = "X";', 1],
  ];
  for (const [名, 字, 期待] of 組) {
    const 出 = 書き換えを拾う(字);
    T('  ' + 名 + ' ⇒ ' + 期待 + '件', 出.length === 期待, '出た ' + 出.length + '件');
  }
  /* ★本物の 本でも 1回 壊して みる★＝見本だけでは 私が 思う 形しか 出ない */
  const 本 = path.join(ROOT, 'lib/shiki-hyou.js');
  const 前 = fs.readFileSync(本, 'utf8');
  T('  ★本物の 本は 今 0件★', 書き換えを拾う(前).length === 0);
  const 壊した = 前.replace('      var ば = 名から番地(n);', '      var ば = 名から番地(n);\n      ば.行 = 0;');
  T('  ★本物を 壊したら 赤に なる★', 壊した !== 前 && 書き換えを拾う(壊した).length > 0,
    壊した === 前 ? '★壊せて いません★' : '出た ' + 書き換えを拾う(壊した).length + '件');
  T('  ★手元の ファイルを 触って いない★', fs.readFileSync(本, 'utf8') === 前);
}

console.log('\n' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
