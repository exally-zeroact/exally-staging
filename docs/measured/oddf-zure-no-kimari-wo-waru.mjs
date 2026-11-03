/* oddf-zure-no-kimari-wo-waru.mjs -- ★Exally1 の 規則を 実測で 割る★（2026-09-19）
 *
 *  ★★何を して いるか★★
 *    ODDFPRICE の ★利率0★（利率 0 ／ 利回り 0.05 ／ 償還 100）は
 *    ★償還を 割り引くだけ★なので ★100 / 1.0125^T★ の 形に なります。
 *    ⇒★T を 逆に 解いて 「T − 次準日までの日数/92」を 見ると ★整数★に なります★
 *    ⇒★その 整数が 24 か 23 か＝★ずれ 0 か −1 か★★
 *
 *  ★★Exally1 の 規則（★測る 前に 書かれました★）★★
 *    ㋐★決済 ≧ EDATE(次の準日, −1)★        ⇒ ★−1★
 *    ㋑★決済 が EDATE(次の準日, −k) に ぴたり乗る★（k≧2）⇒ ★−1★
 *    ★それ以外★ ⇒ 0
 *
 *  ★★なぜ 道具に するか★★
 *    ・★口で「合った」と 言うだけでは 次の 人が 確かめられません★
 *    ・★台を 直した 後に ★同じ 物差し★で もう 一度 割れます★
 *
 *  ★★見て いない 事★★
 *    ・★この 1族だけ★（満 2015-02-28 ／ 発 2008-12-05 ／ 初 2009-08-31 ／ f=4 ／ basis 1）
 *    ・★決済が 次準日（2009-05-31）を ★越えた 側★は 外して います★
 *      ＝★次の 準日が 08-31 に なり 別の 話に なる★
 *    ・★規則が 他の 組でも 立つかは 見て いません★
 *
 *  使い方: node docs/measured/oddf-zure-no-kimari-wo-waru.mjs
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ここ = path.dirname(fileURLToPath(import.meta.url));

/* ★この 1族の 印★（★これに 当たる 行だけ 拾う★） */
const 族 = 'DATE(2015,2,28), DATE(2008,12,5), DATE(2009,8,31), 0, 0.05, 100, 4, 1)';
const 次準日 = new Date(Date.UTC(2009, 4, 31));   /* 2009-05-31 */
const E = 92;
const 率 = Math.log(1.0125);

const 日 = (y, m, d) => new Date(Date.UTC(y, m - 1, d));
const 字 = (dt) => dt.toISOString().slice(0, 10);
const 差日 = (a, b) => Math.round((a - b) / 86400000);
/* ★EDATE と 同じ＝月末は 丸める★ */
function 月を戻す(dt, k) {
  const y0 = dt.getUTCFullYear();
  const m0 = dt.getUTCMonth() - k;
  const y = y0 + Math.floor(m0 / 12);
  const m = ((m0 % 12) + 12) % 12;
  const 末 = new Date(Date.UTC(y, m + 1, 0)).getUTCDate();
  return new Date(Date.UTC(y, m, Math.min(dt.getUTCDate(), 末)));
}

const 一か月前 = 月を戻す(次準日, 1);
const ぴたり = new Set();
for (let k = 2; k <= 40; k += 1) ぴたり.add(字(月を戻す(次準日, k)));

const 見込み = (dt) => (dt >= 一か月前 || ぴたり.has(字(dt)) ? -1 : 0);

/* ══ ★紙を 読む★ ══ */
const 表 = new Map();
for (const f of fs.readdirSync(ここ).filter((f) => /^golden-oddf.*\.tsv$/.test(f))) {
  for (const l of fs.readFileSync(path.join(ここ, f), 'utf-8').split(/\r?\n/)) {
    if (!l || l.startsWith('#') || !l.includes('\t')) continue;
    const c = l.split('\t');
    const 式 = (c[1] || '').trim();
    if (!式.includes(族)) continue;
    const m = /^=ODDFPRICE\(DATE\((\d+),(\d+),(\d+)\)/.exec(式);
    if (!m) continue;
    const d = 日(+m[1], +m[2], +m[3]);
    if (d > 次準日) continue;            /* ★越えた 側は 別の 話★ */
    表.set(字(d), Number(c[2]));
  }
}

const 並び = [...表.keys()].sort();
console.log('');
console.log('[oddf-zure-no-kimari-wo-waru] ★Exally1 の 規則を 実測で 割る★');
console.log('  ★拾った 日 ... ' + 並び.length + '日★（★決済 ≦ 2009-05-31★）');
console.log('  ★次の 準日 ... ' + 字(次準日) + ' ／ その 1か月前 ... ' + 字(一か月前) + '★');

let 合 = 0;
const 外れ = [];
const 引いた = [];
for (const s of 並び) {
  const d = 日(+s.slice(0, 4), +s.slice(5, 7), +s.slice(8, 10));
  const T = -Math.log(表.get(s) / 100) / 率;
  const W = Math.round(T - 差日(次準日, d) / E);
  const ずれ = W - 24;
  if (ずれ === -1) 引いた.push(s);
  if (ずれ === 見込み(d)) 合 += 1;
  else 外れ.push({ s, ずれ, 見: 見込み(d) });
}

console.log('');
console.log('  ★★当たった ' + 合 + ' / ' + 並び.length + '★★ ／ 外れ ' + 外れ.length + '本');
if (外れ.length) {
  for (const x of 外れ) console.log('    ' + x.s + '  実測 ' + x.ずれ + ' ／見込み ' + x.見);
  process.exit(1);
}
console.log('');
console.log('  ★1 引いた 日 ... ' + 引いた.length + '日★');
console.log('    ' + 引いた.join(' / '));
console.log('');
console.log('  ★★この 道具は 「規則が 当たるか」だけを 見ます★★');
console.log('    ＝★台が 直ったかは 見て いません★（`oddf-wo-kami-de-osu.mjs` が 見ます）');
