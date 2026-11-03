/* osu-oddf-buhin.mjs — ★ODDFPRICE の 部品を うちの 台で 押して 実Excel と 突き合わせる★（2026-09-16）
 *
 *  ★★なぜ 部品を 押すか★★
 *    2026-09-16 に ODDFPRICE の ★答えを 当てて 6通り 試し★、★全部 元より 悪く★ なった。
 *    ⇒★答えを 当てるのは 数字合わせ★＝やめた（棚63）
 *    ⇒★★答えでは なく 部品（日数の 数え方）を 突き合わせる★★
 *       ここが 違うなら ★ODDFPRICE が 合わない 訳が 分かります★
 *       ここが 全部 合うなら ★日数の 数え方は 正しい＝式の 組み立てが 違う★
 *
 *  紙: docs/measured/golden-oddf-buhin-2026-09-16.tsv（★実Excel が 書いた 物★）
 *
 *  使い方: node docs/measured/osu-oddf-buhin.mjs
 */
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { ROOT } from './honban-no-michi.mjs';

const require_ = createRequire(path.join(ROOT, 'package.json'));
const H = require_(path.join(ROOT, 'lib/shiki-hyou.js'));

const 紙道 = path.join(ROOT, 'docs/measured/golden-oddf-buhin-2026-09-16.tsv');
const 行 = fs.readFileSync(紙道, 'utf-8').replace(/^﻿/, '')
  .split(/\r?\n/).filter((l) => l && !l.startsWith('#'));

const h = H.表();
let 合 = 0, 違 = 0, 押せず = 0;
const 違う = [];
const 部品ごと = new Map();   /* 部品 → {合, 違} */
const basisごと = new Map();  /* basis → {合, 違} */

const 足す = (m, k, ok) => {
  if (!m.has(k)) m.set(k, { 合: 0, 違: 0 });
  m.get(k)[ok ? '合' : '違']++;
};

for (const l of 行) {
  const c = l.split('\t');
  if (c.length < 5) continue;
  const [組, basis, 部品, 式, 実] = c;
  h.打つ('Z1', 式);
  const 字 = h.字('Z1');
  /* ★押せない（#NAME? 等）は 「違う」と 分けて 数える★ */
  if (/^#/.test(字)) {
    押せず++; 違う.push([組, basis, 部品, 実, 字]);
    足す(部品ごと, 部品, false); 足す(basisごと, basis, false);
    continue;
  }
  const a = Number(字), b = Number(実);
  const 同 = isFinite(a) && isFinite(b) &&
    (a === b || Math.abs(a - b) <= Math.max(Math.abs(b), 1) * 1e-12);
  if (同) { 合++; } else { 違++; 違う.push([組, basis, 部品, 実, 字]); }
  足す(部品ごと, 部品, 同); 足す(basisごと, basis, 同);
}

console.log('');
console.log('★★ODDFPRICE の 部品を 実Excel と 突き合わせた★★ … 押した ' + 行.length + '本');
console.log('');
console.log('  ★合った★ ' + 合 + '本 ／ ★違う★ ' + 違 + '本 ／ ★押せない★ ' + 押せず + '本');
console.log('');
console.log('★部品ごと★');
for (const [k, v] of 部品ごと) {
  console.log('  ' + k.padEnd(12) + '合 ' + String(v.合).padStart(3) + ' ／ 違 ' + String(v.違).padStart(3));
}
console.log('');
console.log('★basis ごと★');
for (const [k, v] of [...basisごと].sort()) {
  console.log('  basis ' + k + ' … 合 ' + String(v.合).padStart(3) + ' ／ 違 ' + String(v.違).padStart(3));
}
if (違う.length) {
  console.log('');
  console.log('★★違う 分（全部）★★');
  console.log('  組 basis 部品          実Excel                うち');
  for (const [組, b, p, e, u] of 違う) {
    console.log('  ' + 組 + '  ' + b + '     ' + p.padEnd(12) + String(e).padEnd(22) + u);
  }
}
