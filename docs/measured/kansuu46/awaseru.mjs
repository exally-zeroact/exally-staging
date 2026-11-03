/* awaseru.mjs — ★作った 物を 実Excel の 答えと 突き合わせる★（2026-09-07）
 *
 *  ★★式は 1か所に だけ 置く★★
 *    `cases-kane.txt` の 式を ★実Excel も ここも 同じ 物を 読む★
 *    ⇒ 私が 書き写す 隙が 無い
 *
 *  ★合っているの 決め方★
 *    数 … ★相対の ずれが 1e-9 より 小さい★（Excel も 私も 二進小数）
 *    字 … そのまま 一致
 *
 *  使い方: node docs/measured/kansuu46/awaseru.mjs [関数名]
 */
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';

const ここ = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(ここ, '..', '..', '..');
const require_ = createRequire(import.meta.url);
const K = require_(path.join(ROOT, 'lib/formula-kane.js'));

const 絞り = (process.argv[2] || '').toUpperCase();
const 金 = fs.readFileSync(path.join(ここ, 'golden-kane-2026-09-07.tsv'), 'utf-8')
  .split('\n').filter((l) => l && !l.startsWith('#'))
  .map((l) => { const p = l.split('\t'); return { 名: p[0], 式: p[1], 答: p[2], 型: p[3] }; });

/* ── 式を 読む ── */
function 引数を割る(s) {
  const 出 = []; let 深 = 0, 今 = '';
  for (const ch of s) {
    if (ch === '(') { 深++; 今 += ch; continue; }
    if (ch === ')') { 深--; 今 += ch; continue; }
    if (ch === ',' && 深 === 0) { 出.push(今.trim()); 今 = ''; continue; }
    今 += ch;
  }
  if (今.trim() !== '') 出.push(今.trim());
  return 出;
}
function 値にする(s) {
  s = s.trim();
  const d = s.match(/^DATE\((\d+),(\d+),(\d+)\)$/i);
  if (d) return K.日から数(+d[1], +d[2], +d[3]);
  if (/^TRUE$/i.test(s)) return true;
  if (/^FALSE$/i.test(s)) return false;
  return Number(s);
}

const 呼ぶ = {
  ACCRINT: (a) => K.経過利息(a[0], a[1], a[2], a[3], a[4], a[5], a[6], a[7]),
  ACCRINTM: (a) => K.満期一括の経過利息(a[0], a[1], a[2], a[3], a[4]),
  AMORDEGRC: (a) => K.仏定率(a[0], a[1], a[2], a[3], a[4], a[5], a[6]),
  AMORLINC: (a) => K.仏定額(a[0], a[1], a[2], a[3], a[4], a[5], a[6]),
  COUPDAYBS: (a) => K.前からの日数(K.数から日(a[0]), K.数から日(a[1]), a[2], 基(a[3])),
  COUPDAYS: (a) => K.期間の日数(K.数から日(a[0]), K.数から日(a[1]), a[2], 基(a[3])),
  COUPDAYSNC: (a) => K.次までの日数(K.数から日(a[0]), K.数から日(a[1]), a[2], 基(a[3])),
  COUPNCD: (a) => { const p = K.次の利払日(K.数から日(a[0]), K.数から日(a[1]), a[2]); return K.日から数(p.y, p.m, p.d); },
  COUPPCD: (a) => { const p = K.前の利払日(K.数から日(a[0]), K.数から日(a[1]), a[2]); return K.日から数(p.y, p.m, p.d); },
  COUPNUM: (a) => K.利払回数(K.数から日(a[0]), K.数から日(a[1]), a[2]),
  DISC: (a) => K.割引率(a[0], a[1], a[2], a[3], a[4]),
  DURATION: (a) => K.期間(a[0], a[1], a[2], a[3], a[4], a[5]),
  INTRATE: (a) => K.利率(a[0], a[1], a[2], a[3], a[4]),
  MDURATION: (a) => K.修正期間(a[0], a[1], a[2], a[3], a[4], a[5]),
  ODDFPRICE: (a) => K.初回端数の価格(a[0], a[1], a[2], a[3], a[4], a[5], a[6], a[7], a[8]),
  ODDFYIELD: (a) => K.初回端数の利回り(a[0], a[1], a[2], a[3], a[4], a[5], a[6], a[7], a[8]),
  ODDLPRICE: (a) => K.最終端数の価格(a[0], a[1], a[2], a[3], a[4], a[5], a[6], a[7]),
  ODDLYIELD: (a) => K.最終端数の利回り(a[0], a[1], a[2], a[3], a[4], a[5], a[6], a[7]),
  PRICE: (a) => K.価格(a[0], a[1], a[2], a[3], a[4], a[5], a[6]),
  PRICEDISC: (a) => K.割引債の価格(a[0], a[1], a[2], a[3], a[4]),
  PRICEMAT: (a) => K.満期一括の価格(a[0], a[1], a[2], a[3], a[4], a[5]),
  RECEIVED: (a) => K.受取額(a[0], a[1], a[2], a[3], a[4]),
  VDB: (a) => K.可変定率(a[0], a[1], a[2], a[3], a[4], a[5], a[6]),
  YIELD: (a) => K.利回り(a[0], a[1], a[2], a[3], a[4], a[5], a[6]),
  YIELDDISC: (a) => K.割引債の利回り(a[0], a[1], a[2], a[3], a[4]),
  YIELDMAT: (a) => K.満期一括の利回り(a[0], a[1], a[2], a[3], a[4], a[5]),
};
const 基 = (b) => (b === undefined || b === null || Number.isNaN(b) ? 0 : b);

function 二桁(n) { return (n < 10 ? '0' : '') + n; }
function 年月日(serial) {
  const p = K.数から日(serial);
  return p.y + '-' + 二桁(p.m) + '-' + 二桁(p.d);
}

let 合 = 0, 違 = 0, 落 = 0;
const 束 = new Map();
const 外れ = [];

for (const g of 金) {
  if (絞り && g.名 !== 絞り) continue;
  const 字包み = g.式.match(/^=TEXT\((.+),"yyyy-mm-dd"\)$/);
  const 中 = 字包み ? '=' + 字包み[1] : g.式;
  const m = 中.match(/^=([A-Z.]+)\((.*)\)$/);
  if (!m) { 落++; continue; }
  const fn = 呼ぶ[m[1]];
  if (!fn) { 落++; continue; }
  const 引 = 引数を割る(m[2]).map(値にする);
  let 出;
  try { 出 = fn(引); } catch (e) { 出 = { 誤り: 'EX:' + e.message }; }
  if (字包み && typeof 出 === 'number') 出 = 年月日(出);

  const 状 = 束.get(g.名) || { 合: 0, 違: 0 };
  let よい = false;
  if (出 && 出.誤り) {
    /* ★実Excel の 赤い 値と 突き合わせる★（#NUM! など）
       ⇒ ここが 無いと ★赤で 合っているのに「違う」に 見える★ */
    よい = (String(g.答) === '#' + 出.誤り + '!' || String(g.答) === '#' + 出.誤り
      || (出.誤り === 'NUM' && g.答 === '#NUM!') || (出.誤り === 'VALUE' && g.答 === '#VALUE!'));
  } else if (g.型 === 'Double') {
    const 正 = Number(g.答);
    よい = (Math.abs(出 - 正) <= Math.max(1e-9, Math.abs(正) * 1e-9));
  } else よい = (String(出) === String(g.答));

  if (よい) { 合++; 状.合++; } else {
    違++; 状.違++;
    if (外れ.length < 400) 外れ.push({ 名: g.名, 式: g.式, 正: g.答, 出: (出 && 出.誤り) ? ('★' + 出.誤り + '★') : String(出) });
  }
  束.set(g.名, 状);
}

console.log('\n[awaseru] 実Excel の 答えと 突き合わせ');
console.log('  ★合った ' + 合 + ' ／ 違った ' + 違 + ' ／ 呼べない ' + 落 + '★\n');
const 並 = [...束.entries()].sort((a, b) => b[1].違 - a[1].違);
for (const [名, s] of 並) {
  console.log('  ' + (s.違 ? '✗' : '✓') + ' ' + 名.padEnd(12) + ' 合 ' + String(s.合).padStart(3) + ' ／ 違 ' + String(s.違).padStart(3));
}
if (外れ.length) {
  console.log('\n  ★違った 物（先頭 ' + Math.min(外れ.length, 30) + '本）★');
  for (const x of 外れ.slice(0, 30)) {
    console.log('    ' + x.名);
    console.log('      式  ' + x.式);
    console.log('      正  ' + x.正);
    console.log('      出  ' + x.出);
  }
}
process.exit(違 ? 1 : 0);
