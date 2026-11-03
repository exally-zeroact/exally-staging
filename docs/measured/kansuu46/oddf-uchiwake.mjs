/* oddf-uchiwake.mjs — ★ODDFPRICE 48本の 内訳を 出す★（2026-09-16・★実Excel は 叩きません★）
 *
 *  ★★なぜ 要るか★★
 *    ODDFPRICE は 24/48 しか 合いません。
 *    ★部品（日数の 数え方）は もう 210/210 合って います★
 *      … COUPDAYBS / COUPDAYS / COUPDAYSNC / COUPNUM / YEARFRAC
 *    ⇒★残るのは ★式の 組み立て★ だけ★
 *
 *  ★★組み立てを 直す 前に 内訳を 出します★★
 *    「24/48」だけでは ★どの 形が 落ちて いるか 分かりません★。
 *    ⇒★組（日付の 束）ごと★／★basis ごと★／★NC（端数期間に 準利払期間が いくつ 入るか）ごと★
 *      に 割って、★落ちて いる 形を 名指し★します。
 *    ⇒記憶「★出す 前に 割れ★＝割る為に 数える事に なり そこで 穴が 出る」
 *
 *  ★★NC とは★★
 *    端数の 初回期間（発行 → 初回利払）に ★準利払期間が いくつ 入るか★。
 *      NC = 1 … 端数が 1期に 収まる  （★今 合って いる 形★の はず）
 *      NC > 1 … 端数が 1期を 越える  （★今 落ちて いる 形★の はず）
 *    ⇒★これを 数で 確かめます（見立てで 書きません）★
 *
 *  ★この 道具は 紙を 読むだけ★＝★実Excel を 叩きません／重く ありません★
 *
 *  使い方: node docs/measured/kansuu46/oddf-uchiwake.mjs
 */
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '../../..');
const require_ = createRequire(path.join(ROOT, 'package.json'));
const K = require_(path.join(ROOT, 'lib/formula-kane.js'));
/* ★★土台（shiki-*）は ODDFPRICE を 知りません★★
     ★24/48 という 数は ★皮（lib/formula-kane.js）の 数★です★
   ⇒★皮の 中身を 直に 呼びます★（★間に 入る 物を 減らす★） */

const 紙道 = path.join(ROOT, 'docs/measured/kansuu46/golden-kane-2026-09-07.tsv');
const 行 = fs.readFileSync(紙道, 'utf-8').replace(/^﻿/, '').split(/\r?\n/)
  .filter((l) => l && !l.startsWith('#'))
  .map((l) => l.split('\t'))
  .filter((c) => c[0] === 'ODDFPRICE');

/* ★式から 引数を 取り出す★（★手で 写しません★） */
const 引数を取る = (式) => {
  const m = /^=ODDFPRICE\((.*)\)$/.exec(式.trim());
  const 中 = m[1];
  /* DATE(y,m,d) を 1つの かたまりに する */
  const 並 = [];
  let 深 = 0, 今 = '';
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
  return { y: +m[1], m: +m[2], d: +m[3] };
};

/* ★NC を 数える★ … 発行 → 初回利払 に 準利払期間が いくつ 入るか
     ★準利払日は 初回利払日から 遡って 置く★（本番の `準の並び前へ` と 同じ 考え）
     ★ここは 日付を 数えるだけ＝実Excel の 部品と 突き合わせ済み★ */
const NCを数える = (発, 初, f) => {
  const 月 = 12 / f;
  let n = 0, cur = { y: 初.y, m: 初.m, d: 初.d };
  const 末日か = (dt) => dt.d === new Date(Date.UTC(dt.y, dt.m, 0)).getUTCDate();
  const 末 = 末日か(初);
  const 引く = (k) => {
    let mm = 初.m - k * 月, yy = 初.y;
    while (mm <= 0) { mm += 12; yy--; }
    while (mm > 12) { mm -= 12; yy++; }
    const 末日 = new Date(Date.UTC(yy, mm, 0)).getUTCDate();
    return { y: yy, m: mm, d: 末 ? 末日 : Math.min(初.d, 末日) };
  };
  const 数に = (dt) => Date.UTC(dt.y, dt.m - 1, dt.d) / 86400000;
  for (let k = 1; k <= 200; k++) {
    n = k;
    cur = 引く(k);
    if (数に(cur) <= 数に(発)) break;
  }
  return n;
};

const 組 = new Map();     /* 日付の束 → {合, 違} */
const basis別 = new Map();
const NC別 = new Map();
const 明細 = [];

const 足す = (m, k, ok) => {
  if (!m.has(k)) m.set(k, { 合: 0, 違: 0 });
  m.get(k)[ok ? '合' : '違']++;
};

for (const c of 行) {
  const 式 = c[1];
  /* ★★紙の 誤りを 数として 読まない★★
       ★前は `Number('#NUM!')` で NaN に なり、★永久に 合わない 6行★に なって いました★
       ⇒★誤りは 誤りと して 突き合わせます★ */
  const 誤りか = /error/.test(c[3] || '') || /^#/.test(c[2]);
  const 実 = 誤りか ? c[2].trim() : Number(c[2]);
  const a = 引数を取る(式);
  const 決 = 日にする(a[0]), 満 = 日にする(a[1]), 発 = 日にする(a[2]), 初 = 日にする(a[3]);
  const f = Number(a[7]);
  const basis = a.length > 8 ? Number(a[8]) : '(既定0)';
  const NC = NCを数える(発, 初, f);

  /* ★皮の 中身を 直に 呼ぶ★（通し番号に するのも 皮の 道具で） */
  const 数 = (dt) => K.日から数(dt.y, dt.m, dt.d);
  const r = K.初回端数の価格(数(決), 数(満), 数(発), 数(初),
    Number(a[4]), Number(a[5]), Number(a[6]), f, a.length > 8 ? Number(a[8]) : 0);
  /* ★うちの 誤りは `{誤り:'NUM'}` の 形★⇒ `#NUM!` に 揃えて 比べる */
  const u誤 = (r && r.誤り) ? ('#' + r.誤り + '!') : null;
  const u = u誤 ? NaN : Number(r);
  const 同 = 誤りか
    ? (u誤 === 実)
    : (!u誤 && isFinite(u) && Math.abs(u - 実) <= Math.abs(実) * 1e-9);

  const 束 = a[0] + '|' + a[1] + '|' + a[2] + '|' + a[3] + '|f' + f;
  足す(組, 束, 同);
  足す(basis別, String(basis), 同);
  足す(NC別, 'NC=' + NC, 同);
  明細.push({ 束: 束, basis: basis, NC: NC, 実: 実, うち: (u誤 || u), 同: 同 });
}

console.log('');
console.log('★★ODDFPRICE 48本の 内訳★★（★紙を 読むだけ＝実Excel を 叩いて いません★）');
console.log('');
const 合計 = 明細.filter((x) => x.同).length;
console.log('  ★合った★ ' + 合計 + ' / ' + 明細.length + '本');
console.log('');

console.log('★★NC ごと（★端数期間に 準利払期間が いくつ 入るか★）★★');
for (const [k, v] of [...NC別].sort()) {
  console.log('  ' + k + ' … 合 ' + String(v.合).padStart(2) + ' ／ 違 ' + String(v.違).padStart(2)
    + (v.違 === 0 ? '  ★全部 合って います★' : (v.合 === 0 ? '  ★★全部 落ちて います★★' : '  ★混ざって います★')));
}
console.log('');

console.log('★basis ごと★');
for (const [k, v] of [...basis別].sort()) {
  console.log('  basis ' + k.padEnd(7) + '… 合 ' + String(v.合).padStart(2) + ' ／ 違 ' + String(v.違).padStart(2));
}
console.log('');

console.log('★組（日付の 束）ごと★');
let i = 0;
for (const [k, v] of 組) {
  i++;
  const ひとつ = 明細.find((x) => x.束 === k);
  console.log('  組' + i + '  NC=' + ひとつ.NC + '  合 ' + String(v.合).padStart(2) + ' ／ 違 ' + String(v.違).padStart(2)
    + (v.違 === 0 ? '  ★全部 合う★' : (v.合 === 0 ? '  ★★全部 落ちる★★' : '')));
  console.log('        ' + k.replace(/\|/g, '  '));
}
console.log('');

console.log('★★落ちて いる 分の 実物（★数だけで 済ませない★）★★');
console.log('  組 basis NC  実Excel                うち                   差');
i = 0;
const 束番号 = new Map();
for (const k of 組.keys()) 束番号.set(k, ++i);
for (const x of 明細) {
  if (x.同) continue;
  console.log('  ' + String(束番号.get(x.束)).padStart(2)
    + '  ' + String(x.basis).padEnd(6)
    + ' ' + x.NC
    + '  ' + String(x.実).padEnd(22)
    + ' ' + String(x.うち).padEnd(22)
    + ' ' + ((typeof x.うち === 'number' && typeof x.実 === 'number')
        ? (x.うち - x.実).toExponential(3) : '★誤りの 出方が 違う★'));
}
