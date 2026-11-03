/* oddf-hasuu-ga-hasuu-denai.test.mjs — ★端数が「端数で ない」時は #NUM!★（2026-09-16）
 *
 *  ★★測って 決めました（★当てて いません★）★★
 *    紙 `docs/measured/golden-oddf-3kaime-2026-09-16.tsv`
 *    ★同じ 紙の 上に COUPDAYS も 打って あります★
 *    ⇒★紙の 中だけで 説明が 付きます★
 *
 *      f basis 端数の実日数 COUPDAYS 同じか  実Excel
 *      1   1    365          365      同じ    ★#NUM!★
 *      1   2    365          360      違う    数
 *      1   3    365          365      同じ    ★#NUM!★
 *      2   1    181          181      同じ    ★#NUM!★
 *      2   2    181          180      違う    数
 *      2   3    181          182.5    違う    数
 *      4   1     90           90      同じ    ★#NUM!★
 *      4   2     90           90      同じ    ★#NUM!★
 *      4   3     90           91.25   違う    数
 *    ⇒★★9通り とも 合います★★
 *
 *  ★意味★ … ★端数期間が めぐり合せの 1期と 同じ 長さなら それは 端数では ありません★
 *           ＝ふつうの 利付債 ⇒ 実Excel は 受け付けません
 *
 *  ★★前は 数を 返して いました★★＝★お客さんには 誤りと 分かりません★
 *
 *  使い方: node tests/oddf-hasuu-ga-hasuu-denai.test.mjs
 */
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(path.join(ROOT, 'package.json'));
const K = require_(path.join(ROOT, 'lib/formula-kane.js'));

let pass = 0, fail = 0;
const T = (n, f) => {
  try { f(); pass++; console.log('  ✓ ' + n); }
  catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + e.message); }
};

console.log('\n[oddf-hasuu-ga-hasuu-denai] ★端数が「端数で ない」時は #NUM!★');

const 紙道 = path.join(ROOT, 'docs/measured/golden-oddf-3kaime-2026-09-16.tsv');
const 行 = fs.readFileSync(紙道, 'utf-8').replace(/^﻿/, '').split(/\r?\n/)
  .filter((l) => l && !l.startsWith('#')).map((l) => l.split('\t'))
  .filter((c) => c.length >= 6);

const 取 = (訳) => { const r = 行.find((c) => c[0] === 訳); return r || null; };
const 数 = (y, m, d) => K.日から数(y, m, d);

console.log('      … 紙 ' + 行.length + '行');

T('★紙が 読めて いる（★空振りして いない★）★', () => {
  if (行.length < 40) throw new Error('★' + 行.length + '行しか 読めない★');
  for (const n of ['㋑f=1 (basis=1)', 'COUPDAYS f=1 (basis=1)']) {
    if (!取(n)) throw new Error('★紙に「' + n + '」が 無い★');
  }
});

/* ★f ごとの 初回利払日と 端数の 実日数★（★発行は どれも 2009-01-01★） */
const 組 = [
  { f: 1, 初: [2010, 1, 1], 端: 365 },
  { f: 2, 初: [2009, 7, 1], 端: 181 },
  { f: 4, 初: [2009, 4, 1], 端: 90 },
];

T('★★紙の 中だけで 決まりが 説明できる（★COUPDAYS と 突き合わせる★）★★', () => {
  /* ★「端数の 実日数 ＝ COUPDAYS」の 時だけ #NUM! か★
       ⇒★これが 崩れたら 決まりが 違います★ */
  const 違 = [];
  for (const g of 組) {
    for (const b of [1, 2, 3]) {
      const cd = 取('COUPDAYS f=' + g.f + ' (basis=' + b + ')')
        || 取('COUPDAYS 1月始まり (basis=' + b + ')');
      if (!cd) throw new Error('★COUPDAYS f=' + g.f + ' basis=' + b + ' が 紙に 無い★');
      const 同 = (g.端 === Number(cd[2]));
      const r = 取('㋑f=' + g.f + ' (basis=' + b + ')');
      if (!r) throw new Error('★㋑f=' + g.f + ' basis=' + b + ' が 紙に 無い★');
      const 誤 = (r[3] === '#NUM!');
      if (同 !== 誤) {
        違.push('f=' + g.f + ' basis=' + b + '（端数 ' + g.端 + ' ／ COUPDAYS ' + cd[2]
          + ' ／ 実Excel ' + r[3] + '）');
      }
    }
  }
  if (違.length) {
    throw new Error('★' + 違.length + '/9 通り 決まりが 崩れて いる★ … ' + 違.join(' ／ '));
  }
  console.log('      … 9/9 通り（★紙の 中だけで 説明が 付きます★）');
});

T('★★うちも 9通り とも 同じ★★', () => {
  const 違 = [];
  for (const g of 組) {
    for (const b of [1, 2, 3]) {
      const r = 取('㋑f=' + g.f + ' (basis=' + b + ')');
      const 待 = (r[3] === '#NUM!');
      const v = K.初回端数の価格(数(2009, 3, 1), 数(2013, 1, 1), 数(2009, 1, 1),
        数(g.初[0], g.初[1], g.初[2]), 0.06, 0.05, 100, g.f, b);
      const 誤 = !!(v && v.誤り === 'NUM');
      if (誤 !== 待) {
        違.push('f=' + g.f + ' basis=' + b + '（実Excel ' + r[3] + ' ／ うち ' + (誤 ? '#NUM!' : v) + '）');
      }
    }
  }
  if (違.length) throw new Error('★' + 違.length + '/9 通り 違う★ … ' + 違.join(' ／ '));
  console.log('      … 9/9 通り');
});

T('★★数を 返す 分は 値も 合う（★誤りだけ 合わせて いません★）★★', () => {
  /* ★#NUM! を 返すだけなら 中身が 壊れて いても 通ります★
       ⇒★数を 返す 分は ★値まで★ 押します★ */
  const 違 = [];
  let 押した = 0;
  for (const g of 組) {
    for (const b of [1, 2, 3]) {
      const r = 取('㋑f=' + g.f + ' (basis=' + b + ')');
      if (r[3] === '#NUM!') continue;
      押した++;
      const u = K.初回端数の価格(数(2009, 3, 1), 数(2013, 1, 1), 数(2009, 1, 1),
        数(g.初[0], g.初[1], g.初[2]), 0.06, 0.05, 100, g.f, b);
      const e = Number(r[2]);
      if (!(typeof u === 'number' && Math.abs(u - e) <= Math.abs(e) * 1e-12)) {
        違.push('f=' + g.f + ' basis=' + b + ' 実' + e + ' うち' + u);
      }
    }
  }
  if (違.length) throw new Error('★' + 違.length + '/' + 押した + '本 違う★ … ' + 違.join(' ／ '));
  console.log('      … ' + 押した + '/' + 押した + '本');
});

T('★★NC=1 では basis 2・3 も ぴたり 合う（★問題は NC>1 だけ★）★★', () => {
  /* ★2026-09-16 の 実測★
       端数 ちょうど 1期（発行 2009-01-01 ／ 初回 2009-07-01）で 決済を 動かすと
       ★basis 2・3 は 10本とも 1e-14 で 合いました★
     ⇒★★合わないのは NC>1 の 形だけ★★ */
  const 日 = [[1, 2], [1, 3], [4, 1], [6, 29], [6, 30]];
  const 違 = [];
  let 押した = 0;
  for (const [mm, dd] of 日) {
    for (const b of [2, 3]) {
      const r = 取('㋐NC1で1日ずつ DATE(2009,' + mm + ',' + dd + ') (basis=' + b + ')');
      if (!r) throw new Error('★紙に 無い★ … ' + mm + '/' + dd + ' basis=' + b);
      押した++;
      const u = K.初回端数の価格(数(2009, mm, dd), 数(2013, 1, 1), 数(2009, 1, 1),
        数(2009, 7, 1), 0.06, 0.05, 100, 2, b);
      const e = Number(r[2]);
      if (!(typeof u === 'number' && Math.abs(u - e) <= Math.abs(e) * 1e-12)) {
        違.push('2009-' + mm + '-' + dd + ' basis=' + b + ' 実' + e + ' うち' + u);
      }
    }
  }
  if (違.length) throw new Error('★' + 違.length + '/' + 押した + '本 違う★ … ' + 違.join(' ／ '));
  console.log('      … ' + 押した + '/' + 押した + '本（★NC=1 は 全部 合います★）');
});

console.log('\n' + pass + ' passed, ' + fail + ' failed');
if (fail) process.exit(1);
