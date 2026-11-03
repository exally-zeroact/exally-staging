/* kansuu46-1taba.test.mjs — ★台に 無かった 49個の 1束目 7 ＋ 2束目 5 ＋ 3束目 2を 紙で 押す★（2026-09-18）
 *
 *  ★★足した 14個★★
 *    ★1束目 7個★ MDETERM ／ KURT ／ TRIMMEAN ／ FORECAST ／ FORECAST.LINEAR
 *                ／ NUMBERVALUE ／ ENCODEURL
 *    ★2束目 5個★ FIXED ／ DOLLAR ／ VALUETOTEXT ／ LOOKUP ／ PHONETIC
 *    ★3束目 2個★ INDIRECT ／ OFFSET
 *      ⇒★この 2つの 為に `lib/shiki-hyou.js` が ★マスを 読む 口★を 渡すように しました★
 *        （`所.四角を取る` ／ `所.字から取る`・★読む 口だけ／書く 口は 渡して いません★）
 *
 *  ★★なぜ 別の 紙（この 試験）が 要るか★★
 *    `shiki-kansuu-kami.test.mjs` は ★紙の `#材料` から★ 材料を 取ります。
 *    ⇒★`golden-oddf-to-46ko-2026-09-16.tsv` には `#材料` が 在りません★
 *    ⇒★★だから この 7個は 1行も 押されて いませんでした★★（実測）
 *    ⇒★紙に 後から `#材料` を 書き足すのは ★測った 記録を 触る★事に なります★
 *      ＝★代わりに ★材料の 出どころ★を 名指しで 書いて ここで 押します★
 *
 *  ★★材料の 出どころ（★私が 決めて いません★）★★
 *    `docs/measured/kansuu46-no-dodai.mjs` の `材料()`
 *      A1:A5 = 1,2,3,4,5 ／ B1:B5 = 2,4,6,8,10
 *      D1 = `=DATE(2024,1,1)` ／ D2 = `=DATE(2026,1,1)`
 *    ★そこに「実Excel を 測った 時と 同じ」と 書いて 在ります★
 *    ★裏を 取りました★ ... `=FORECAST(6,B1:B5,A1:A5)` が 12
 *      ＝B ＝ 2x なので ★この 材料でしか 12に なりません★
 *      ＝`=MDETERM(A1:B2)` が 0 ＝ (1,2／2,4) ＝ ★同じ 材料★
 *
 *  ★★式は 材料の 外（BZ1）★★（`#CYCLE` 避け・`shiki-kansuu-kami` と 同じ）
 *
 *  ★★数どうしは ★数★で 比べます★★
 *    紙は PowerShell の `'R'`／台は JS の 一番 短い 書き方
 *    ⇒★字で 比べると 偽の 負けに なります★（`shiki-kansuu-kami` と 同じ 決め）
 *
 *  ★★この 試験が 見て いない 事★★
 *    ・★1つの 関数に つき 紙は 1行しか 在りません★
 *      ⇒★端（誤り・負・空・長さ違い）は ★測って いません★★
 *      ⇒★`lib/shiki-kansuu.js` の 各関数の 下に ★未測定★と 名指しで 書いて あります★
 *    ・★画面では ありません★（台だけ）
 *
 *  使い方: node tests/kansuu46-1taba.test.mjs
 */
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(path.join(ROOT, 'package.json'));
const H = require_(path.join(ROOT, 'lib/shiki-hyou.js'));
const K = require_(path.join(ROOT, 'lib/shiki-kansuu.js'));

let pass = 0, fail = 0;
const T = (n, よい, 添え) => {
  if (よい) { pass++; console.log('  ok   ' + n); }
  else { fail++; console.log('  NG   ' + n + (添え ? '\n       ' + 添え : '')); }
};

console.log('');
console.log('[kansuu46-1taba] ★台に 無かった 49個の 1束目 7 ＋ 2束目 5 ＋ 3束目 2★');

/* ══ ★足した 分★（★この 名前は 手で 書きます＝「何を 足したか」は 人が 決めた 事★） ══ */
const 足した = ['MDETERM', 'KURT', 'TRIMMEAN', 'FORECAST', 'FORECAST.LINEAR',
  'NUMBERVALUE', 'ENCODEURL',
  /* ★2束目（2026-09-18）★ */
  'FIXED', 'DOLLAR', 'VALUETOTEXT', 'LOOKUP', 'PHONETIC',
  /* ★3束目（2026-09-18）★ 板に ★マスを 読む 口★を 足して 書けた 2個 */
  'INDIRECT', 'OFFSET'];

for (const n of 足した) {
  T('★台が ' + n + ' を 知って いる★', typeof (K.表 || {})[n] === 'function');
}

/* ══ ★紙から 行を 拾う（★手で 写しません★） ══ */
const 紙道 = path.join(ROOT, 'docs/measured/golden-oddf-to-46ko-2026-09-16.tsv');
T('★紙が 在る★', fs.existsSync(紙道), 紙道);
const 行 = fs.readFileSync(紙道, 'utf-8').split(/\r?\n/);
const 柱 = (行.find((l) => l.startsWith('# 種')) || '').replace(/^#\s*/, '').split('\t');
const 式列 = 柱.indexOf('式'), 答列 = 柱.indexOf('答え');
T('★柱が 読めた★（式／答え）', 式列 >= 0 && 答列 >= 0, '柱 ... ' + 柱.join(' / '));

const 組 = [];
for (const l of 行) {
  if (!l || l.startsWith('#')) continue;
  const c = l.split('\t');
  const 式 = c[式列];
  if (!式 || !式.startsWith('=')) continue;
  const m = /^=([A-Z][A-Z0-9._]*)\(/.exec(式);
  if (!m || 足した.indexOf(m[1]) < 0) continue;
  組.push({ 名: m[1], 式: 式, 実: c[答列] });
}
console.log('  ★紙から 拾った 行 ... ' + 組.length + '行★');
/* ★★本数を 決め打ちに する★★＝★紙から 消えたら 赤★（★黙って 0行に なるのを 止める★） */
const 行の本数 = 14;   /* ★1束目 7 ＋ 2束目 5 ＋ 3束目 2★ */
T('★★拾った 行が ' + 行の本数 + '行★★（★増えても 減っても 赤★）',
  組.length === 行の本数, '拾った ' + 組.length + '行 ... ' + 組.map((x) => x.名).join(' '));

/* ★★足した 分 とも 1行ずつ 在るか★★（★合計が 合っても 片寄る 事が 在る★） */
const 名ごと = {};
for (const q of 組) 名ごと[q.名] = (名ごと[q.名] || 0) + 1;
const 抜け = 足した.filter((n) => !名ごと[n]);
T('★★' + 足した.length + '個 とも 紙に 在る★★', 抜け.length === 0, '紙に 無い ... ' + 抜け.join(' '));

/* ══ ★材料（★出どころは 頭に 書いた 通り★） ══ */
function 板を作る() {
  const h = H.表();
  const A = [1, 2, 3, 4, 5], B = [2, 4, 6, 8, 10];
  for (let i = 0; i < 5; i++) {
    h.打つ('A' + (i + 1), String(A[i]));
    h.打つ('B' + (i + 1), String(B[i]));
  }
  h.打つ('D1', '=DATE(2024,1,1)');
  h.打つ('D2', '=DATE(2026,1,1)');
  return h;
}

/* ★材料が 本当に 入ったか 読み返す★（★空の 板を 押して 偽の 負けに した 事が 在る★） */
{
  const h = 板を作る();
  h.打つ('BZ1', '=SUM(A1:A5)');
  const v = h.値 ? h.値('BZ1') : null;
  const s = String(h.字('BZ1'));
  T('★材料が 入った★（=SUM(A1:A5) が 15）', s === '15', '出た「' + s + '」');
}

/* ══ ★1行ずつ 押す★ ══ */
let 合 = 0, 違 = 0;
const 外れ = [];
for (const q of 組) {
  const h = 板を作る();
  h.打つ('BZ1', q.式);
  let 生 = null;
  try { 生 = h.値 ? h.値('BZ1') : null; } catch (e) { /* 字で 見る */ }
  const 字 = String(h.字('BZ1'));
  const 台の数 = (生 && 生.型 === '数') ? 生.値 : (isFinite(Number(字)) && 字 !== '' ? Number(字) : null);
  const 紙の数 = isFinite(Number(q.実)) && q.実 !== '' ? Number(q.実) : null;
  let よい;
  if (台の数 !== null && 紙の数 !== null) {
    /* ★★数どうしは 数で★★（紙は 'R' の 17桁・台は 一番 短い 書き方） */
    よい = (台の数 === 紙の数)
      || Math.abs(台の数 - 紙の数) <= Math.max(1e-9, Math.abs(紙の数) * 1e-9);
  } else {
    よい = (字 === String(q.実));
  }
  if (よい) 合++;
  else { 違++; 外れ.push(q.名 + ' ' + q.式 + ' => 台「' + 字 + '」／紙「' + q.実 + '」'); }
}
console.log('  ★★合った ' + 合 + ' / ' + 組.length + '★★');
for (const s of 外れ) console.log('       ・' + s);
T('★★' + 組.length + '行 とも 紙と 合う★★', 違 === 0, '違った ' + 違 + '行');

console.log('');
console.log('  ★★この 試験が 言える 事／言えない 事★★');
console.log('    ・言える ... ★紙に 在る 1行は 合う★（足した ' + 足した.length + '個 とも）');
console.log('    ・★★言えない ... 端（誤り・負・空・長さ違い）★★');
console.log('      ＝★1つの 関数に つき 紙が 1行しか 在りません★');
console.log('      ＝★`lib/shiki-kansuu.js` の 各関数の 下に ★未測定★と 書いて あります★');
console.log('');
console.log('kansuu46-1taba: ' + pass + ' 緑 / ' + fail + ' 赤');
process.exit(fail ? 1 : 0);
