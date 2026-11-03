/* bessel.test.mjs — ★BESSELI / BESSELJ / BESSELK / BESSELY★（2026-09-16）
 *
 *  ★★合う 幅は 1e-8 です★★（★ふつうは 1e-12 です★）
 *    訳 … ★★実Excel の BESSEL は 8〜9桁しか 正しく ありません★★
 *      `=BESSELK(2,3)` の 実Excel の 答え … ★0.6473854＝7桁★（他の 関数は 17桁）
 *    ⇒★幅を 1e-12 に すると ★うちが 正しくても 8本とも 赤★に なります★
 *    ⇒★但し 1e-3 まで 緩めません★
 *      ＝★1本を 通す 為に 幅を 広げると 本物の 誤りが 隠れます★（経営者1 の 条件）
 *
 *  ★★合わない 3本は ★名指しで 免除★★★（★どれも 実Excel の 方が 間違って います★）
 *    =BESSELI(101,2) 差 5.14e-4 ／ =BESSELK(101,2) 差 9.58e-8 ／ =BESSELY(2,3) 差 1.04e-8
 *    ⇒★3つの 別々の 道で 確かめました★（級数／漸近形／★積分★）
 *    ⇒紙 `docs/measured/bessel-no-kotae.md`
 *    ⇒★逃げ道の 門★ `tests/jitsuexcel-ga-machigai.test.mjs`
 *
 *  使い方: node tests/bessel.test.mjs
 */
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(path.join(ROOT, 'package.json'));
const B = require_(path.join(ROOT, 'lib/bessel.js'));

let pass = 0, fail = 0;
const T = (n, f) => {
  try { f(); pass++; console.log('  ✓ ' + n); }
  catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + e.message); }
};

console.log('\n[bessel] ★BESSELI / BESSELJ / BESSELK / BESSELY★');

const 紙道 = path.join(ROOT, 'docs/measured/kansuu46/golden-346-2026-09-08.tsv');
const 行 = fs.readFileSync(紙道, 'utf-8').replace(/^﻿/, '').split(/\r?\n/)
  .filter((l) => /^BESSEL[IJKY]\t/.test(l)).map((l) => l.split('\t'));

/* ★引数が その場に 書いて ある 式だけ 押します★
     （`=BESSELJ(D1,D2)` は ★材料が 要る★ので ここでは 押しません） */
const 押す = [];
for (const c of 行) {
  const m = /^=BESSEL([IJKY])\("?(-?[\d.]+)"?,\s*"?(-?[\d.]+)"?\)$/.exec(c[1]);
  if (!m) continue;
  押す.push({ 種: m[1], x: Number(m[2]), n: Math.trunc(Number(m[3])), 式: c[1], 実: Number(c[2]) });
}

/* ★★名指しの 免除★★（★実Excel の 方が 間違って いる 物★）
     ★★どれも ★3つ目の 道（積分）★で 確かめてから 足しました★★
     ★門が 狙いどおり 効きました★
       ＝★最初 1件で 置いた門が 赤に なり、★測ってから 名前を 書く★ 事に なった★ */
const 免除 = [
  '=BESSELI("101",2)',   /* 差 5.14e-4 … 積分 2.847013943032507e+42 */
  '=BESSELK("101",2)',   /* 差 9.58e-8 … 積分 1.73851808855756e-45 */
  '=BESSELY(2,3)',       /* 差 1.04e-8 … 積分 -1.127783776725578 */
];

console.log('      … 紙 ' + 行.length + '行 ／ 押す ' + 押す.length + '本 ／ 免除 ' + 免除.length + '本');

T('★紙が 読めて いる（★空振りして いない★）★', () => {
  if (押す.length < 8) throw new Error('★' + 押す.length + '本しか 押せない★（紙が 壊れて いる）');
  for (const t of ['I', 'J', 'K', 'Y']) {
    if (!押す.some((r) => r.種 === t)) throw new Error('★BESSEL' + t + ' が 1本も 無い★');
  }
});

const 出す = (r) => B['BESSEL' + r.種](r.x, r.n);

T('★★紙と 合う（★合う 幅 1e-8★・免除は 名指し）★★', () => {
  const 違 = [];
  let 押した = 0;
  for (const r of 押す) {
    if (免除.indexOf(r.式) >= 0) continue;
    押した++;
    const u = 出す(r);
    /* ★★実Excel が ★何桁 出したか★ で 比べます★★（2026-09-16）
         ★なぜ★ `=BESSELK(2,3)` の 実Excel の 答えは ★`0.6473854`＝7桁★ だけです。
           うちの 値 0.6473853909486349 を ★7桁に 丸めると 0.6473854★
           ⇒★★これは 「ずれ」で は ありません★★＝★実Excelが そこまでしか 出さなかった★
         ★幅を 緩めるのと は 違います★
           ＝★実Excel が 出した 桁数だけ 見る★（★それ以下は 向こうが 言って いません★）
         積分でも 0.6473853909486315 ★うちと 同じ★ */
    const 桁 = (String(r.実).replace(/[-+.]|e[-+]?\d+$/gi, '').replace(/^0+/, '') || '0').length;
    const 幅 = Math.max(1e-8, Math.pow(10, -(桁 - 1)) / 2);
    const 差 = Math.abs(u - r.実) / Math.max(Math.abs(r.実), 1e-300);
    if (!(差 <= 幅)) {
      違.push(r.式 + ' 実' + r.実 + '（' + 桁 + '桁） うち' + u + ' 差' + 差.toExponential(2));
    }
  }
  if (違.length) {
    throw new Error('★' + 違.length + '/' + 押した + '本 違う★ … ' + 違.join(' ／ '));
  }
  console.log('      … ' + 押した + '/' + 押した + '本（★免除を 除く★）');
});

T('★★免除した 3本は「実Excel の 方が 間違って いる」＝★うちは 積分と 合う★★★', () => {
  /* ★免除を「見なかった 事」に しません★
       ＝★★★3つ目の 道（積分）★で 出した 値と 合う 事を 押します★★★
       ★積分は ★級数でも 漸近形でも ない 別の 世界★★ */
  const 見る = [
    ['=BESSELI("101",2)', 2.847013943032507e+42, 1e-13],
    ['=BESSELK("101",2)', 1.73851808855756e-45, 1e-11],
    ['=BESSELY(2,3)', -1.127783776725578, 1e-9],
  ];
  for (const [式, 積分, 幅] of 見る) {
    const r = 押す.find((x) => x.式 === 式);
    if (!r) throw new Error('★免除した 式が 紙に 無い★ … ' + 式);
    const u = 出す(r);
    if (!(Math.abs(u - 積分) <= Math.abs(積分) * 幅)) {
      throw new Error('★' + 式 + ' が 積分と 合わない★ … うち ' + u + ' ／ 積分 ' + 積分);
    }
    /* ★実Excel と 違う 事も 押す★＝★同じに なったら 免除は 要りません★ */
    const 差 = Math.abs(u - r.実) / Math.abs(r.実);
    if (差 <= 1e-8) {
      throw new Error('★' + 式 + ' が 実Excel と 合うように なりました★（差 ' + 差.toExponential(2) + '）'
        + '／★免除を 外して ください★');
    }
    console.log('      … ' + 式 + ' うち ' + u + ' ／ 実Excel ' + r.実 + '（差 ' + 差.toExponential(2) + '）');
  }
});

T('★別の 道でも 同じ 値に なる（★級数と 漸近形★）★', () => {
  /* ★1つの 道だけでは ★自分の 間違いを 自分で 見つけられません★★ */
  const x = 101, nu = 2, mu = 4 * nu * nu;
  let 項 = 1, 和 = 1;
  for (let k = 1; k < 12; k++) {
    項 *= -(mu - (2 * k - 1) * (2 * k - 1)) / (k * 8 * x);
    和 += 項;
    if (Math.abs(項) < 1e-18) break;
  }
  const 漸近 = Math.exp(x) / Math.sqrt(2 * Math.PI * x) * 和;
  const 級数 = B.BESSELI(101, 2);
  if (!(Math.abs(漸近 - 級数) <= 級数 * 1e-13)) {
    throw new Error('★級数 ' + 級数 + ' と 漸近形 ' + 漸近 + ' が 合わない★');
  }
});

T('★名前の 一覧が 4つ★', () => {
  const a = B.名前たち();
  if (a.length !== 4) throw new Error('★' + a.length + '個★（4個の はず）… ' + a.join(' '));
});

console.log('\n' + pass + ' passed, ' + fail + ' failed');
if (fail) process.exit(1);
