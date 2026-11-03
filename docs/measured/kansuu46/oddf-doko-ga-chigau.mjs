/* oddf-doko-ga-chigau.mjs — ★ODDFPRICE の 落ちる 組は 何が 違うか★（2026-09-18）
 *
 *  ★★先に 分かった 事★★
 *    `oddf-buhin-awaseru.mjs` ... ★部品は 210/210 全部 合って います★
 *      （COUPDAYBS / COUPDAYS / COUPDAYSNC / COUPNUM / YEARFRAC）
 *    ⇒★★日数の 数え方は 合って います★★
 *    ⇒★★落ちて いるのは ★組み立て（式）★です★★
 *
 *  ★★だから ここでは 組の 形を 並べます★★
 *    ・準利払日の 並び（初回から 発行まで さかのぼる）
 *    ・NC（端数の 期間に 準利払日が 何個 入るか）
 *    ・発行が 準利払日の 上に 在るか
 *    ・決済が 何個目の 準期間に 在るか
 *  ⇒★★合う 組と 合わない 組で ★どこが 違うか★を 目で 見る★★
 *
 *  ★★これは 答えを 当てる 道具では ありません★★
 *    ＝★割れ目（どの 形で 落ちるか）を 出すだけ★
 *
 *  使い方: node docs/measured/kansuu46/oddf-doko-ga-chigau.mjs
 */
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '../../..');
const require_ = createRequire(path.join(ROOT, 'package.json'));
const K = require_(path.join(ROOT, 'lib/formula-kane.js'));

const 紙道 = path.join(ROOT, 'docs/measured/kansuu46/golden-kane-2026-09-07.tsv');
const 行 = fs.readFileSync(紙道, 'utf-8').split(/\r?\n/).filter((l) => l && !l.startsWith('#'));

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
function 数に(s) {
  const m = /^DATE\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)$/.exec(s.trim());
  if (!m) return Number(s.trim());
  return K.日から数(Number(m[1]), Number(m[2]), Number(m[3]));
}
const 字 = (n) => { const d = K.数から日(n); return d.y + '-' + String(d.m).padStart(2, '0') + '-' + String(d.d).padStart(2, '0'); };

/* ★準利払日の 並び★（初回から 頻度ぶんずつ さかのぼり 発行を 越えるまで） */
function 準日たち(発行, 初回, 頻度) {
  const 月 = 12 / 頻度;
  const 初 = K.数から日(初回);
  const 末 = (初.d === new Date(Date.UTC(初.y, 初.m, 0)).getUTCDate());
  const 出 = [初回];
  for (let k = 1; k < 200; k++) {
    const y = 初.y, m = 初.m - k * 月;
    let yy = y + Math.floor((m - 1) / 12), mm = ((m - 1) % 12 + 12) % 12 + 1;
    const 末日 = new Date(Date.UTC(yy, mm, 0)).getUTCDate();
    const dd = 末 ? 末日 : Math.min(初.d, 末日);
    const n = K.日から数(yy, mm, dd);
    出.unshift(n);
    if (n <= 発行) break;
  }
  return 出;
}

/* ★紙から ODDFPRICE の 行を 集める★（★組ごとに 1本 だけ 形を 見ます★） */
const 組 = new Map();
for (const l of 行) {
  const c = l.split('\t');
  if (c[0] !== 'ODDFPRICE') continue;
  const 引 = 引数を割る(c[1].slice(c[1].indexOf('(') + 1, c[1].lastIndexOf(')')));
  const v = 引.map(数に);
  const 鍵 = v.slice(0, 4).join(',') + '|' + v[7];
  if (!組.has(鍵)) 組.set(鍵, { v: v, 合: 0, 違: 0, 式: [] });
  const g = 組.get(鍵);
  /* ★台で 押して 紙と 比べる★ */
  /* ★★`初回端数の価格` は ★通し番号★を 受け取ります★★
       （中で `数から日(整える(...))` を します＝★`{y,m,d}` を 渡すと NaN★）
       ★私は さっき 部品の 方で 逆を やりました★＝★口の 形を 毎回 確かめる★ */
  const 出 = K.初回端数の価格(v[0], v[1], v[2], v[3],
    v[4], v[5], v[6], v[7], v.length > 8 ? v[8] : 0);
  const 正 = Number(c[2]);
  const よい = isFinite(出) && isFinite(正) && Math.abs(出 - 正) <= Math.max(1e-9, Math.abs(正) * 1e-9);
  if (よい) g.合++; else { g.違++; g.式.push({ basis: v.length > 8 ? v[8] : 0, 正: 正, 出: 出 }); }
}

console.log('');
console.log('★★ODDFPRICE の 組ごとの 形★★（★部品は 210/210 合って います＝組み立てが 元★）');
console.log('');
let 番 = 0;
for (const [, g] of 組) {
  番++;
  const [決済, 満期, 発行, 初回, , , , 頻度] = g.v;
  const 準 = 準日たち(発行, 初回, 頻度);
  const NC = 準.length - 1;
  const 発行が準日 = 準.includes(発行);
  let 決済の枠 = -1;
  for (let i = 0; i + 1 < 準.length; i++) if (発行 <= 決済 ? true : true) {
    if (決済 > 準[i] && 決済 <= 準[i + 1]) 決済の枠 = i + 1;
  }
  if (決済 <= 準[0]) 決済の枠 = 0;
  console.log('  組' + 番 + '  ' + (g.違 ? '★違 ' + g.違 + ' ／ 合 ' + g.合 + '★' : '合 ' + g.合 + ' ／ 違 0'));
  console.log('        決済 ' + 字(決済) + ' ／ 満期 ' + 字(満期)
    + ' ／ 発行 ' + 字(発行) + ' ／ 初回 ' + 字(初回) + ' ／ f=' + 頻度);
  console.log('        準利払日 ' + 準.map(字).join(' '));
  console.log('        NC=' + NC + ' ／ 発行が準日の上=' + (発行が準日 ? '○' : '★×★')
    + ' ／ 決済は ' + 決済の枠 + '枠目'
    + ' ／ 決済<発行=' + (決済 < 発行 ? '★○★' : '×'));
  for (const x of g.式) {
    console.log('        ★basis ' + x.basis + '★  正 ' + x.正 + ' ／ 出 ' + x.出
      + ' ／ 差 ' + (x.出 - x.正).toExponential(3));
  }
}
console.log('');
console.log('★見る 所★');
console.log('  ・★合う 組と 合わない 組で ★NC／発行が準日／決済の枠★ が どう 違うか★');
console.log('  ・★差が basis で どう 動くか★（★basis 4 だけ 合う 組が 在る★）');
