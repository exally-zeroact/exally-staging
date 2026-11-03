/* odd-yonko-no-uchiwake.mjs — ★ODD 4個の 落ちが ★同じ 元か★ を 数える★（2026-09-17）
 *
 *  ★★なぜ 要るか★★
 *    ODDFYIELD は ★ODDFPRICE を 逆に 解く 形★です。
 *    ODDLYIELD も ★ODDLPRICE を 逆に 解く 形★です。
 *    ⇒★★だから 落ちが 同じ 組に 揃って いる はず★★
 *    ⇒★揃って いれば ★ODDFPRICE を 直すと 72通りが 一度に 直ります★★
 *    ⇒★揃って いなければ ★別の 元が もう 1つ 在ります★★
 *
 *  ★★これは 見立てでは なく 数で 出します★★
 *
 *  ★この 道具は 紙を 読むだけ★＝★実Excel を 叩きません★
 *
 *  使い方: node docs/measured/kansuu46/odd-yonko-no-uchiwake.mjs
 */
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '../../..');
const require_ = createRequire(path.join(ROOT, 'package.json'));
const K = require_(path.join(ROOT, 'lib/formula-kane.js'));

const 紙道 = path.join(ROOT, 'docs/measured/kansuu46/golden-kane-2026-09-07.tsv');
const 行 = fs.readFileSync(紙道, 'utf-8').replace(/^﻿/, '').split(/\r?\n/)
  .filter((l) => l && !l.startsWith('#')).map((l) => l.split('\t'))
  .filter((c) => /^ODD[FL](PRICE|YIELD)$/.test(c[0]));

const 引数を取る = (式) => {
  const 中 = /^=ODD[FL](?:PRICE|YIELD)\((.*)\)$/.exec(式.trim())[1];
  const 並 = []; let 深 = 0, 今 = '';
  for (const ch of 中) {
    if (ch === '(') { 深++; 今 += ch; }
    else if (ch === ')') { 深--; 今 += ch; }
    else if (ch === ',' && 深 === 0) { 並.push(今); 今 = ''; }
    else 今 += ch;
  }
  並.push(今);
  return 並;
};
const 日にする = (s) => {
  const m = /^DATE\((\d+),(\d+),(\d+)\)$/.exec(s.trim());
  return m ? { y: +m[1], m: +m[2], d: +m[3] } : null;
};
const 数 = (dt) => K.日から数(dt.y, dt.m, dt.d);

/* ★組（日付の 束）ごとに 4つの 関数の 合否を 並べます★ */
const 組 = new Map();

for (const c of 行) {
  const 名 = c[0], 式 = c[1];
  const 誤りか = /error/.test(c[3] || '') || /^#/.test(c[2]);
  const a = 引数を取る(式);
  const 日 = a.map(日にする);
  const fが初 = /^ODDF/.test(名);
  /* ODDF … 決, 満, 発, 初, 利, （利回り|価格）, 償還, 頻度, basis
     ODDL … 決, 満, 最終, 利, （利回り|価格）, 償還, 頻度, basis */
  const 束 = fが初
    ? [a[0], a[1], a[2], a[3]].join(' ')
    : [a[0], a[1], a[2]].join(' ');
  const f = Number(fが初 ? a[7] : a[6]);
  const basis = (fが初 ? (a.length > 8 ? Number(a[8]) : 0) : (a.length > 7 ? Number(a[7]) : 0));

  let r;
  try {
    if (名 === 'ODDFPRICE') {
      r = K.初回端数の価格(数(日[0]), 数(日[1]), 数(日[2]), 数(日[3]),
        Number(a[4]), Number(a[5]), Number(a[6]), f, basis);
    } else if (名 === 'ODDFYIELD') {
      r = K.初回端数の利回り(数(日[0]), 数(日[1]), 数(日[2]), 数(日[3]),
        Number(a[4]), Number(a[5]), Number(a[6]), f, basis);
    } else if (名 === 'ODDLPRICE') {
      r = K.最終端数の価格(数(日[0]), 数(日[1]), 数(日[2]),
        Number(a[3]), Number(a[4]), Number(a[5]), f, basis);
    } else {
      r = K.最終端数の利回り(数(日[0]), 数(日[1]), 数(日[2]),
        Number(a[3]), Number(a[4]), Number(a[5]), f, basis);
    }
  } catch (e) { r = { 誤り: 'VALUE' }; }

  const u誤 = (r && r.誤り) ? ('#' + r.誤り + '!') : null;
  const 同 = 誤りか ? (u誤 === String(c[2]).trim())
    : (!u誤 && Math.abs(Number(r) - Number(c[2])) <= Math.abs(Number(c[2])) * 1e-9);

  const 鍵 = (fが初 ? 'F' : 'L') + '｜' + 束;
  if (!組.has(鍵)) 組.set(鍵, {});
  const g = 組.get(鍵);
  if (!g[名]) g[名] = { 合: 0, 違: 0 };
  g[名][同 ? '合' : '違']++;
}

console.log('');
console.log('★★ODD 4個の 落ちが 同じ 元か★★（★紙を 読むだけ＝実Excel を 叩いて いません★）');
console.log('');
let 揃 = 0, 揃わず = 0;
let i = 0;
for (const [鍵, g] of 組) {
  i++;
  const 種 = 鍵[0];
  const 価 = 種 === 'F' ? g.ODDFPRICE : g.ODDLPRICE;
  const 利 = 種 === 'F' ? g.ODDFYIELD : g.ODDLYIELD;
  const 出 = (x) => x ? (x.違 === 0 ? '○全部' : (x.合 === 0 ? '★×全部★' : '合' + x.合 + '/違' + x.違)) : '（無）';
  /* ★揃って いるか★ … 価格が 全部 合う 組で 利回りも 全部 合うか（逆も） */
  let 判 = '—';
  if (価 && 利) {
    const 価全合 = 価.違 === 0, 利全合 = 利.違 === 0;
    if (価全合 === 利全合) { 判 = '★揃って いる★'; 揃++; }
    else { 判 = '★★揃って いない★★'; 揃わず++; }
  }
  console.log('  組' + String(i).padStart(2) + ' [' + 種 + '] 価格 ' + 出(価).padEnd(10)
    + ' 利回り ' + 出(利).padEnd(10) + ' ' + 判);
  console.log('        ' + 鍵.slice(2));
}
console.log('');
console.log('  ★揃って いる★ ' + 揃 + '組 ／ ★揃って いない★ ' + 揃わず + '組');
console.log('');
if (揃わず === 0) {
  console.log('★★落ちは 同じ 元です★★');
  console.log('  ⇒★価格を 直すと 利回りも 一度に 直ります★');
} else {
  console.log('★★別の 元が もう 1つ 在ります★★（★揃って いない 組を 見て ください★）');
}
