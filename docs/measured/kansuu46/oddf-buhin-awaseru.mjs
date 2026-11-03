/* oddf-buhin-awaseru.mjs — ★ODDF の 部品を 紙と 突き合わせる★（2026-09-18）
 *
 *  ★★なぜ★★
 *    ODDFPRICE は ★39/48★（2026-09-18）。落ちる のは ★組B と 組C★だけです。
 *    ⇒★答えを 当てに 行くのでは なく ★部品（日数の 数え方）★が 合って いるかを 先に 見ます★
 *    ⇒★部品が 合って いれば ★組み立て（式）が 元★★
 *    ⇒★部品が 違えば ★そこが 元★★
 *
 *  ★★紙★★ `docs/measured/golden-oddf-buhin-2026-09-16.tsv`（★実Excel・210行★）
 *    COUPDAYBS ／ COUPDAYS ／ COUPDAYSNC ／ COUPNUM ／ YEARFRAC
 *
 *  ★★この 道具が 見て いない 事★★
 *    ・★ODDFPRICE 自体は 押しません★（★部品だけ★）
 *    ・★紙に 在る 組だけ★（A〜F）
 *
 *  使い方: node docs/measured/kansuu46/oddf-buhin-awaseru.mjs
 */
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '../../..');
const require_ = createRequire(path.join(ROOT, 'package.json'));
const K = require_(path.join(ROOT, 'lib/formula-kane.js'));

const 紙道 = path.join(ROOT, 'docs/measured/golden-oddf-buhin-2026-09-16.tsv');
const 行 = fs.readFileSync(紙道, 'utf-8').split(/\r?\n/).filter((l) => l && !l.startsWith('#'));

/* ★DATE(y,m,d) を 通し番号に★（`formula-kane.js` と 同じ 台を 使う） */
function 日付にする(s) {
  const m = /^DATE\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)$/.exec(s.trim());
  if (!m) return Number(s.trim());
  return K.日から数(Number(m[1]), Number(m[2]), Number(m[3]));
}
function 引数を割る(s) {
  const 出 = []; let 深 = 0, 今 = '';
  for (const ch of s) {
    if (ch === '(') 深++;
    if (ch === ')') 深--;
    if (ch === ',' && 深 === 0) { 出.push(今); 今 = ''; continue; }
    今 += ch;
  }
  出.push(今);
  return 出;
}

/* ★台の 呼び方★（★紙の 関数名 -> `formula-kane.js` の 口★）
     ★★日は 通し番号では なく `{y, m, d}` です★★（★最初 通し番号を 渡して NaN を 96本 出しました★）
     ⇒★★道具が 壊れて いるのを 台の 欠陥と 読む 所でした★★
     ★COUPNUM は 引数 3つ★（basis を 取りません） */
const 日 = (n) => K.数から日(n);
const 呼ぶ = {
  COUPDAYBS: (a) => K.前からの日数(日(a[0]), 日(a[1]), a[2], a[3]),
  COUPDAYS: (a) => K.期間の日数(日(a[0]), 日(a[1]), a[2], a[3]),
  COUPDAYSNC: (a) => K.次までの日数(日(a[0]), 日(a[1]), a[2], a[3]),
  COUPNUM: (a) => K.利払回数(日(a[0]), 日(a[1]), a[2]),
  YEARFRAC: (a) => K.日数(日(a[0]), 日(a[1]), a[2]) / K.年の日数(日(a[0]), 日(a[1]), a[2]),
};

let 合 = 0, 違 = 0, 呼べない = 0;
const 外れ = [];
const 組ごと = {};
for (const l of 行) {
  const c = l.split('\t');
  const 組 = c[0], 部品 = c[2], 式 = c[3], 答 = c[4];
  const 名 = 式.slice(1, 式.indexOf('('));
  const f = 呼ぶ[名];
  if (!f) { 呼べない++; continue; }
  const 引 = 引数を割る(式.slice(式.indexOf('(') + 1, 式.lastIndexOf(')'))).map(日付にする);
  let 出;
  try { 出 = f(引); } catch (e) { 出 = null; }
  const 正 = Number(答);
  if (!組ごと[組]) 組ごと[組] = { 合: 0, 違: 0 };
  const よい = (出 !== null && isFinite(出) && isFinite(正)
    && (出 === 正 || Math.abs(出 - 正) <= Math.max(1e-9, Math.abs(正) * 1e-9)));
  if (よい) { 合++; 組ごと[組].合++; }
  else {
    違++; 組ごと[組].違++;
    外れ.push(組 + '  ' + 部品 + '  ' + 式 + '\n        正 ' + 答 + '\n        出 ' + 出);
  }
}

console.log('');
console.log('★★ODDF の 部品を 紙と 突き合わせる★★（★実Excel 210行★）');
console.log('');
console.log('  ★★合った ' + 合 + ' ／ 違った ' + 違 + ' ／ 呼べない ' + 呼べない + '★★');
console.log('');
console.log('  ★組ごと★');
for (const k of Object.keys(組ごと).sort()) {
  const g = 組ごと[k];
  console.log('    組' + k + '  合 ' + String(g.合).padStart(3) + ' ／ 違 ' + String(g.違).padStart(3)
    + (g.違 ? '  ★ここ★' : ''));
}
if (外れ.length) {
  console.log('');
  console.log('★★違った 物★★');
  for (const s of 外れ) console.log('    ' + s);
}
console.log('');
console.log('★言えない 事★');
console.log('  ・★ODDFPRICE 自体は 押して いません★（★部品だけ★）');
console.log('  ・★部品が 全部 合って いても 組み立てが 合うとは 限りません★');
