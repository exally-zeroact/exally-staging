/* ★★ODDFPRICE の 利率0 の 行から ★割り引きの 指数★を 解く★★（2026-09-19）
 *
 *   ★利率 0 の 時★ ... 値 ＝ 償還 ÷ (1 + 利回り/f)^指数
 *   ⇒★指数 ＝ log(償還 / 値) ÷ log(1 + 利回り/f)★（★逆算では なく 定義そのもの★）
 *
 *   ★そして 指数は 2つに 分かれます★（★2026-09-19 に 63点で 測った★）
 *     ★指数 ＝ W ＋ （決済から ★次の 準日★までの 実日数）÷ E★
 *       ・準日 ... 初回利払日から ★1つずつ★ 月を 遡った 日（丸めた 日を そのまま 使う）
 *       ・次の準日 ... ★決済 以上★の 一番 近い 準日（★決済＝準日 なら 0日★）
 *       ・E ...... その 準日で 終わる 準期間の 実日数
 *       ・W ...... ★整数★（★整数に ならなければ この 形では ない★）
 *
 *   ★使い方★ ... node docs/measured/kansuu46/oddf-shisuu-toku.mjs <紙.tsv> [<紙2.tsv> ...]
 *   ★出す物★ ... 決済 ／ 次準日 ／ n ／ E ／ 指数 ／ W ／ 整数か ／ 隣との 差
 */
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';

const ROOT = 'C:/Users/zeroa/exally-prod';
const R = createRequire(path.join(ROOT, 'package.json'));
const K = R(path.join(ROOT, 'lib/formula-kane.js'));

const NL = String.fromCharCode(10), TAB = String.fromCharCode(9);
const 数 = (y, m, d) => K.日から数(y, m, d);
const 日 = (n) => K.数から日(n);
const 月の日数 = (y, m) => new Date(Date.UTC(y, m, 0)).getUTCDate();

function 月足す(p, k) {
  let y = p.y, m = p.m + k;
  while (m > 12) { m -= 12; y++; }
  while (m < 1) { m += 12; y--; }
  return { y, m, d: Math.min(p.d, 月の日数(y, m)) };
}
/** ★初回から 1つずつ 遡って 決済を 跨ぐまで★ */
function 準の並び(初, 決n, f) {
  const 月 = 12 / f, 出 = [初];
  let cur = 初;
  for (let i = 0; i < 400; i++) {
    cur = 月足す(cur, -月);
    出.unshift(cur);
    if (数(cur.y, cur.m, cur.d) <= 決n) break;
  }
  return 出;
}
function 割(s) {
  const o = []; let dep = 0, c = '';
  for (const ch of s) {
    if (ch === '(') { dep++; c += ch; continue; }
    if (ch === ')') { dep--; c += ch; continue; }
    if (ch === ',' && dep === 0) { o.push(c.trim()); c = ''; continue; }
    c += ch;
  }
  if (c.trim() !== '') o.push(c.trim());
  return o;
}
/* ★正規の 字は 正規の 形で 書く★（便りや heredoc で 逆斜線が 1段 落ちる）*/
const DATE正規 = /^DATE\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)$/i;
function vv(s) {
  s = String(s).trim();
  const m = DATE正規.exec(s);
  if (m) return 数(+m[1], +m[2], +m[3]);
  return Number(s);
}
const 字 = (p) => p.y + '-' + String(p.m).padStart(2, '0') + '-' + String(p.d).padStart(2, '0');

const 紙たち = process.argv.slice(2);
if (紙たち.length === 0) { console.log('★紙を 1枚以上 指して ください★'); process.exit(1); }

const 行たち = [];
for (const 紙 of 紙たち) {
  const 中 = fs.readFileSync(path.isAbsolute(紙) ? 紙 : path.join(ROOT, 紙), 'utf-8');
  for (const l of 中.split(NL)) {
    if (!l || l.startsWith('#')) continue;
    const p = l.split(TAB);
    if (p.length < 3) continue;
    const 式 = p[1], 答 = Number(p[2]);
    if (!式 || !/^=ODDFPRICE\(/i.test(式)) continue;
    const a = 割(式.replace(/^=ODDFPRICE\(/i, '').replace(/\)\s*$/, '')).map(vv);
    if (a.length < 8) continue;
    if (a[4] !== 0) continue;                 /* ★利率0 の 行だけ★ */
    if (!Number.isFinite(答)) continue;
    行たち.push({ 紙: path.basename(紙), a, 答 });
  }
}
console.log('★★利率0 の ODDFPRICE ... ' + 行たち.length + '行★★');
console.log('');
行たち.sort((x, y) => (x.a[1] - y.a[1]) || (x.a[3] - y.a[3]) || (x.a[0] - y.a[0]));

let 前 = null;
console.log('決済        次準日      n    E   指数          W     整数か  隣との W差');
for (const r of 行たち) {
  const [決n, 満n, 発n, 初n, , 利回り, 償還, f] = r.a;
  const 並 = 準の並び(日(初n), 決n, f);
  let 次 = null, 前準 = null;
  for (let i = 0; i < 並.length; i++) {
    if (数(並[i].y, 並[i].m, 並[i].d) >= 決n) { 次 = 並[i]; 前準 = 並[i - 1] || null; break; }
  }
  if (!次 || !前準) { console.log('★並びから 外れました★', 決n); continue; }
  const n = 数(次.y, 次.m, 次.d) - 決n;
  const E = 数(次.y, 次.m, 次.d) - 数(前準.y, 前準.m, 前準.d);
  const 指数 = Math.log(償還 / r.答) / Math.log(1 + 利回り / f);
  const W = 指数 - n / E;
  const 丸 = Math.round(W);
  const 整数か = Math.abs(W - 丸) < 1e-9;
  const 同じ組 = 前 && 前.満 === 満n && 前.初 === 初n;
  const 差 = 同じ組 ? (丸 - 前.W) : null;
  console.log(
    字(日(決n)).padEnd(12) + 字(次).padEnd(12) +
    String(n).padStart(3) + ' ' + String(E).padStart(4) + '  ' +
    指数.toFixed(9).padStart(13) + '  ' + String(丸).padStart(4) + '   ' +
    (整数か ? '★整数★' : '★★違う★★ ' + W.toFixed(9)) +
    (差 === null ? '' : ('   ' + (差 === 0 ? '同' : (差 > 0 ? '+' : '') + 差)))
  );
  前 = { 満: 満n, 初: 初n, W: 丸 };
}
