/* yearfrac-basis1.test.mjs — ★YEARFRAC の basis 1 が 実Excel と 合うか★（2026-09-16）
 *
 *  ★★なぜ 要るか（★客に 出て いた 欠陥★）★★
 *    `YEARFRAC(始, 終, 1)`（実日数/実日数）が ★実Excel と 違う 答えを 返して いました★。
 *    うちは ★いつでも「年の 平均日数」で 割って いた★（365.5 など）。
 *    実Excel は ★365 と 366 を 使い分けて います★。
 *    ★YEARFRAC は お客さんが 直に 打つ 関数★＝★今 間違った 答えが 出て いた★。
 *
 *  ★★どうやって 見つけたか★★
 *    ODDFPRICE が 合わない ので ★答えを 当てる のを やめて 部品を 実Excel に 聞いた★。
 *    COUPDAYBS/COUPDAYS/COUPDAYSNC/COUPNUM は ★120本 全部 合った★。
 *    ★YEARFRAC basis 1 だけ 30本中 7本 違った★。
 *    （★棚63 で ODDFPRICE の 答えを 6通り 当てて 全部 悪く なった★ので 当てるのは やめた）
 *
 *  ★★この 見張りが 見る 物★★
 *    ①紙が 読めて いる（★空振りして いない★＝分母を 出さない 緑は 嘘）
 *    ②★32組 全部 実Excel と 合う★
 *    ③★分母（日数 ÷ 答え）も 合う★＝★答えが たまたま 合う 形を 弾く★
 *    ④★分かれ道が 紙に 全部 入って いる★（①〜⑥が 1つずつ 在るか）
 *
 *  ★紙の 取り直し方★
 *    pwsh -NoProfile -File docs/measured/toru-yearfrac-basis1.ps1
 *
 *  使い方: node tests/yearfrac-basis1.test.mjs
 */
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(path.join(ROOT, 'package.json'));
const H = require_(path.join(ROOT, 'lib/shiki-hyou.js'));

let pass = 0, fail = 0;
const T = (n, f) => {
  try { f(); pass++; console.log('  ✓ ' + n); }
  catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + e.message); }
};

console.log('\n[yearfrac-basis1] ★YEARFRAC の basis 1 が 実Excel と 合うか★');

const 紙道 = path.join(ROOT, 'docs/measured/golden-yearfrac-basis1-2026-09-16.tsv');
const 紙 = fs.readFileSync(紙道, 'utf-8').replace(/^﻿/, '')
  .split(/\r?\n/).filter((l) => l && !l.startsWith('#'))
  .map((l) => l.split('\t')).filter((c) => c.length >= 8)
  .map((c) => ({ 始: c[0], 終: c[1], 日数: Number(c[2]), 答: c[3], 分母: Number(c[4]), 訳: c[7] }));

console.log('      … 紙 ' + 紙.length + '組');

T('★紙が 読めて いる（★空振りして いない★）★', () => {
  if (紙.length < 30) throw new Error('★' + 紙.length + '組しか 読めない★（紙が 壊れて いる）');
});

T('★分かれ道が 紙に 全部 入って いる（①〜⑥）★', () => {
  /* ★「32組 合った」は 分かれ道が 1本 抜けて いても 出ます★
       ⇒★どの 道が 何組 在るかを 名指しで 数える★ */
  const 数 = {};
  for (const r of 紙) { const k = r.訳.slice(0, 1); 数[k] = (数[k] || 0) + 1; }
  for (const k of ['①', '②', '③', '④', '⑤', '⑥']) {
    if (!数[k]) throw new Error('★分かれ道 ' + k + ' が 紙に 1組も 無い★');
  }
  console.log('      … ' + ['①', '②', '③', '④', '⑤', '⑥'].map((k) => k + 数[k]).join(' '));
});

const 式にする = (s) => {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s);
  return 'DATE(' + Number(m[1]) + ',' + Number(m[2]) + ',' + Number(m[3]) + ')';
};

T('★★32組 全部 実Excel と 合う★★', () => {
  const h = H.表();
  const 違 = [];
  for (const r of 紙) {
    h.打つ('Z1', '=YEARFRAC(' + 式にする(r.始) + ',' + 式にする(r.終) + ',1)');
    const u = Number(h.字('Z1')), e = Number(r.答);
    if (!(Math.abs(u - e) <= Math.max(Math.abs(e), 1) * 1e-12)) 違.push([r.始, r.終, e, u, r.訳]);
  }
  if (違.length) {
    throw new Error('★' + 違.length + '/' + 紙.length + '組 違う★ … '
      + 違.slice(0, 4).map((x) => x[0] + '→' + x[1] + ' 実' + x[2] + ' うち' + x[3]).join(' ／ '));
  }
  console.log('      … ' + 紙.length + '/' + 紙.length + '組');
});

T('★★分母（日数 ÷ 答え）も 合う＝たまたま 合う 形を 弾く★★', () => {
  /* ★答えだけ 見ると ★別の 分母でも 同じ 答え★ に なる 組が 在ります★
       （日数 365・分母 365 と 日数 366・分母 366 は どちらも 1）
     ⇒★分母そのものを 突き合わせる★ */
  const h = H.表();
  const 違 = [];
  for (const r of 紙) {
    h.打つ('Z1', '=YEARFRAC(' + 式にする(r.始) + ',' + 式にする(r.終) + ',1)');
    const u = Number(h.字('Z1'));
    if (u === 0) continue; /* ★0 では 割れません★（今の 紙には 在りません） */
    const 分母 = r.日数 / u;
    if (Math.abs(分母 - r.分母) > 1e-9) {
      違.push([r.始, r.終, r.分母, 分母, r.訳]);
    }
  }
  if (違.length) {
    throw new Error('★' + 違.length + '組 分母が 違う★ … '
      + 違.slice(0, 4).map((x) => x[0] + '→' + x[1] + ' 実' + x[2] + ' うち' + x[3]).join(' ／ '));
  }
});

console.log('\n' + pass + ' passed, ' + fail + ' failed');
if (fail) process.exit(1);
