/* ★BESSEL … 大きい x を ハンケルの 漸近形で★（試し 3回目・2026-09-16）
 *
 *  ★2回目で 分かった 事★
 *    ①★級数は 数学として 厳密★なのに ★実Excel と 1e-9 で ずれる★
 *       ⇒★★うちが 間違って いるので は ありません★★
 *       ⇒★実Excel の BESSEL が 8〜9桁しか 正しくない★
 *         （`=BESSELK(2,3)` が `0.6473854`＝★7桁★なのも 同じ）
 *    ②★x が 大きい J で 後退漸化を 使ったのが 間違い★
 *       ⇒後退漸化は ★x < n★ の 時の 手
 *       ⇒★x > n★ は ★J_0・J_1 を 出して 前へ 漸化★
 *
 *  ★★J_0・J_1 を どう 出すか★★
 *    ★覚えた 係数表は 使いません★
 *    ⇒★★ハンケルの 漸近形★★（★係数は 式で 作れます★）
 *        J_ν(x) ≈ √(2/(πx)) [ P cos ω − Q sin ω ]，ω = x − (ν/2 + 1/4)π
 *        P = Σ_k (−1)^k a_{2k}/x^{2k}   Q = Σ_k (−1)^k a_{2k+1}/x^{2k+1}
 *        a_k(ν) = (4ν²−1)(4ν²−9)…(4ν²−(2k−1)²) / (k! 8^k)
 *      ⇒★a_k は ★掛け算で 作れます★＝★覚える 数は 在りません★
 */
import fs from 'node:fs';

function 級数(x, n, s) {
  const h = Math.abs(x) / 2, h2 = h * h;
  let t = 1;
  for (let i = 1; i <= n; i++) t *= h / i;
  let 和 = t;
  for (let k = 1; k < 600; k++) {
    t *= s * h2 / (k * (k + n));
    和 += t;
    if (Math.abs(t) < Math.abs(和) * 1e-17) break;
  }
  return (x < 0 && n % 2) ? -和 : 和;
}
const I = (x, n) => (x === 0 ? (n === 0 ? 1 : 0) : 級数(x, n, 1));

/* ★ハンケルの 漸近形★（★a_k を 式で 作る★） */
function J漸近(x, nu) {
  const a = Math.abs(x);
  const mu = 4 * nu * nu;
  let P = 0, Q = 0;
  let 項 = 1;                     /* a_0 = 1 */
  let 前の大きさ = Infinity;      /* ★漸近は 途中から 悪く なる＝一番 小さい 所で 止める★ */
  /* ★a_k = a_{k−1} × (μ − (2k−1)²) / (k·8)★ */
  for (let k = 0; k < 40; k++) {
    if (k > 0) 項 = 項 * (mu - (2 * k - 1) * (2 * k - 1)) / (k * 8);
    const 大きさ = Math.abs(項) / Math.pow(a, k);
    if (k > 2 && 大きさ > 1e-17 * Math.abs(P || 1) && 大きさ > 前の大きさ) break; /* ★漸近は 途中で 悪く なる★ */
    前の大きさ = 大きさ;
    const 符 = (Math.floor(k / 2) % 2) ? -1 : 1;
    if (k % 2 === 0) P += 符 * 項 / Math.pow(a, k);
    else Q += 符 * 項 / Math.pow(a, k);
    if (大きさ < 1e-18) break;
  }
  const w = a - (nu / 2 + 0.25) * Math.PI;
  return Math.sqrt(2 / (Math.PI * a)) * (P * Math.cos(w) - Q * Math.sin(w));
}

function J(x, n) {
  if (x === 0) return n === 0 ? 1 : 0;
  const a = Math.abs(x);
  let 答;
  if (a < n + 15) {
    答 = 級数(a, n, -1);
  } else {
    /* ★J_0・J_1 を 漸近で 出して 前へ 漸化★
         J_{k+1}(x) = (2k/x) J_k(x) − J_{k−1}(x) */
    let j0 = J漸近(a, 0), j1 = J漸近(a, 1);
    if (n === 0) 答 = j0;
    else if (n === 1) 答 = j1;
    else {
      let 前 = j0, 今 = j1;
      for (let k = 1; k < n; k++) { const 次 = (2 * k / a) * 今 - 前; 前 = 今; 今 = 次; }
      答 = 今;
    }
  }
  return (x < 0 && n % 2) ? -答 : 答;
}

/* ══ ★紙で 押す★ ══ */
const 紙 = fs.readFileSync('C:/Users/zeroa/exally-prod/docs/measured/kansuu46/golden-346-2026-09-08.tsv', 'utf-8')
  .split(/\r?\n/).filter((l) => /^BESSEL[IJ]\t/.test(l)).map((l) => l.split('\t'));

console.log('');
console.log('★★BESSELI と BESSELJ … 級数 ＋ ハンケルの 漸近形★★');
console.log('');
for (const 幅 of [1e-12, 1e-9, 1e-8, 1e-7]) {
  let 合 = 0, 全 = 0;
  for (const c of 紙) {
    const m = /^=BESSEL([IJ])\("?([-\d.]+)"?,\s*"?([-\d.]+)"?\)$/.exec(c[1]);
    if (!m) continue;
    const u = m[1] === 'I' ? I(Number(m[2]), Math.trunc(Number(m[3]))) : J(Number(m[2]), Math.trunc(Number(m[3])));
    const e = Number(c[2]);
    全++;
    if (Math.abs(u - e) <= Math.max(Math.abs(e), 1e-300) * 幅) 合++;
  }
  console.log('  ★合う 幅 ' + 幅.toExponential(0).padEnd(6) + '★ … ' + 合 + ' / ' + 全 + '本');
}
console.log('');
console.log('★1本ずつ★');
for (const c of 紙) {
  const m = /^=BESSEL([IJ])\("?([-\d.]+)"?,\s*"?([-\d.]+)"?\)$/.exec(c[1]);
  if (!m) continue;
  const u = m[1] === 'I' ? I(Number(m[2]), Math.trunc(Number(m[3]))) : J(Number(m[2]), Math.trunc(Number(m[3])));
  const e = Number(c[2]);
  const 差 = Math.abs(u - e) / Math.max(Math.abs(e), 1e-300);
  console.log('  ' + c[1].padEnd(22) + ' 実 ' + String(e).padEnd(24) + ' うち ' + String(u).padEnd(24) + ' 相対差 ' + 差.toExponential(2));
}
