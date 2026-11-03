/* bessel.js — ★BESSELI / BESSELJ / BESSELK / BESSELY★（2026-09-16）
 *
 *  ★★覚えた 係数表は 1つも 使って いません★★
 *    2026-09-16 に J0(2) の 係数を 覚えで 書いて ★0.4116（真の値 0.2239）★と 間違えました。
 *    ⇒★あの 時★ … ★覚えて いる 係数表を 書いた★＝覚え違いが そのまま 答えに なった
 *    ⇒★今回★ … ★級数と 漸化式と 漸近形だけ★＝★覚え違いの 入る 所が 在りません★
 *      （★漸近形の 係数 a_k も 掛け算で 作ります★）
 *
 *  ★★★実Excel の 方が 間違って いる 所が 在ります★★★
 *    `=BESSELI(101,2)` … 実Excel 2.848476711286702e+42
 *                        ★正しい値 2.847013943032507e+42★（相対差 5.14e-4）
 *    ★3つの 別々の 道で 確かめました★（級数／漸近形／★積分★）
 *    ⇒紙 `docs/measured/bessel-no-kotae.md`
 *    ⇒見張り `tests/jitsuexcel-ga-machigai.test.mjs`（★逃げ道を 塞ぐ★）
 *    ★この 判じは BESSEL の 1件だけ★＝★他所の 根拠に しません★
 *
 *  ★★お客さんの 画面には まだ 出て いません★★（2026-09-16 の 実測）
 *    `book.html` の `<script src="lib/shiki-` … ★0件★
 *    `book.html` の 遅れて 読む 道（`_loadScript`）に `shiki-` … ★0件★
 *    ⇒★★今 お客さんが BESSELI を 打つと 答えるのは ★借り物★ です★★
 *    ⇒★これは「これから 出す 分が 1つ 増えた」という 事です★
 *
 *  ★合う 幅は 1e-8★（★実Excel が 8〜9桁しか 正しくない★ため）
 *
 *  ★使う 式★
 *    I_n(x) = Σ_{k≥0} (x/2)^(2k+n) / (k!(k+n)!)            ★全部 正＝桁落ちしません★
 *    J_n(x) = Σ_{k≥0} (−1)^k (x/2)^(2k+n) / (k!(k+n)!)     ★x が 大きいと 桁落ち★
 *    K_0(x) = −(ln(x/2)+γ) I_0(x) + Σ_{k≥1} (x²/4)^k H_k /(k!)²
 *    K_1(x) = 1/x + (ln(x/2)+γ) I_1(x) − (x/4) Σ_{k≥0} (x²/4)^k (H_k+H_{k+1})/(k!(k+1)!)
 *    Y_0(x) = (2/π)[ (ln(x/2)+γ) J_0(x) + Σ_{k≥1} (−1)^{k+1} (x²/4)^k H_k /(k!)² ]
 *    Y_1(x) = (2/π)[ (ln(x/2)+γ) J_1(x) − 1/x
 *                    − (x/4) Σ_{k≥0} (−1)^k (x²/4)^k (H_k+H_{k+1})/(k!(k+1)!) ]
 *    ★漸化★ J_{n+1} = (2n/x)J_n − J_{n−1} ／ Y_{n+1} = (2n/x)Y_n − Y_{n−1}
 *            K_{n+1} = K_{n−1} + (2n/x)K_n
 *    ★漸近★ a_k(ν) = a_{k−1} × (4ν² − (2k−1)²) / (8k)
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.Bessel = factory();
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  /* ★オイラー定数★（★これは 定義された 数＝覚えた 近似の 係数表では ありません★） */
  var γ = 0.5772156649015328606;

  /* ★項の 比で 進める 級数★（s = +1 なら I／−1 なら J） */
  function 級数(x, n, s) {
    var h = Math.abs(x) / 2, h2 = h * h;
    var t = 1, i;
    for (i = 1; i <= n; i++) t *= h / i;
    var 和 = t;
    for (var k = 1; k < 600; k++) {
      t *= s * h2 / (k * (k + n));
      和 += t;
      if (Math.abs(t) < Math.abs(和) * 1e-17) break;
    }
    return (x < 0 && n % 2) ? -和 : 和;
  }

  /* ★ハンケルの 漸近（J と Y の 両方）★ … ★a_k は 掛け算で 作ります★ */
  function PQ(x, nu) {
    var mu = 4 * nu * nu;
    var P = 0, Q = 0, 項 = 1, 前 = Infinity;
    for (var k = 0; k < 40; k++) {
      if (k > 0) 項 = 項 * (mu - (2 * k - 1) * (2 * k - 1)) / (k * 8);
      var 大 = Math.abs(項) / Math.pow(x, k);
      /* ★漸近は 途中から 悪く なります★＝★一番 小さい 所で 止める★ */
      if (k > 2 && 大 > 前) break;
      前 = 大;
      var 符 = (Math.floor(k / 2) % 2) ? -1 : 1;
      if (k % 2 === 0) P += 符 * 項 / Math.pow(x, k);
      else Q += 符 * 項 / Math.pow(x, k);
      if (大 < 1e-18) break;
    }
    return { P: P, Q: Q, w: x - (nu / 2 + 0.25) * Math.PI };
  }

  function J漸近(x, nu) {
    var a = PQ(x, nu);
    return Math.sqrt(2 / (Math.PI * x)) * (a.P * Math.cos(a.w) - a.Q * Math.sin(a.w));
  }
  function Y漸近(x, nu) {
    var a = PQ(x, nu);
    return Math.sqrt(2 / (Math.PI * x)) * (a.P * Math.sin(a.w) + a.Q * Math.cos(a.w));
  }
  function K漸近(x, nu) {
    var mu = 4 * nu * nu;
    var 和 = 1, 項 = 1, 前 = Infinity;
    for (var k = 1; k < 40; k++) {
      項 = 項 * (mu - (2 * k - 1) * (2 * k - 1)) / (k * 8 * x);
      var 大 = Math.abs(項);
      if (大 > 前) break;
      前 = 大;
      和 += 項;
      if (大 < 1e-18) break;
    }
    return Math.sqrt(Math.PI / (2 * x)) * Math.exp(-x) * 和;
  }

  /* ★調和数を 足しながら 進む 級数★ */
  function Y0級数(x) {
    var h2 = x * x / 4, t = 1, H = 0, 和 = 0;
    for (var k = 1; k < 600; k++) {
      t *= h2 / (k * k);
      H += 1 / k;
      var 項 = (k % 2 ? 1 : -1) * t * H;
      和 += 項;
      if (Math.abs(項) < Math.abs(和) * 1e-17) break;
    }
    return (2 / Math.PI) * ((Math.log(x / 2) + γ) * 級数(x, 0, -1) + 和);
  }
  function Y1級数(x) {
    var h2 = x * x / 4, t = 1, Hk = 0, Hk1 = 1, 和 = Hk + Hk1;
    for (var k = 1; k < 600; k++) {
      t *= h2 / (k * (k + 1));
      Hk += 1 / k;
      Hk1 += 1 / (k + 1);
      var 項 = (k % 2 ? -1 : 1) * t * (Hk + Hk1);
      和 += 項;
      if (Math.abs(項) < Math.abs(和) * 1e-17) break;
    }
    return (2 / Math.PI) * ((Math.log(x / 2) + γ) * 級数(x, 1, -1) - 1 / x - (x / 4) * 和);
  }
  function K0級数(x) {
    var h2 = x * x / 4, t = 1, H = 0, 和 = 0;
    for (var k = 1; k < 600; k++) {
      t *= h2 / (k * k);
      H += 1 / k;
      和 += t * H;
      if (t * H < 和 * 1e-17) break;
    }
    return -(Math.log(x / 2) + γ) * 級数(x, 0, 1) + 和;
  }
  function K1級数(x) {
    var h2 = x * x / 4, t = 1, Hk = 0, Hk1 = 1, 和 = Hk + Hk1;
    for (var k = 1; k < 600; k++) {
      t *= h2 / (k * (k + 1));
      Hk += 1 / k;
      Hk1 += 1 / (k + 1);
      和 += t * (Hk + Hk1);
      if (t * (Hk + Hk1) < 和 * 1e-17) break;
    }
    return 1 / x + (Math.log(x / 2) + γ) * 級数(x, 1, 1) - (x / 4) * 和;
  }

  /* ══ ★表に 出す 4つ★ ══ */
  function BESSELI(x, n) {
    if (x === 0) return n === 0 ? 1 : 0;
    return 級数(x, n, 1);
  }
  function BESSELJ(x, n) {
    if (x === 0) return n === 0 ? 1 : 0;
    var a = Math.abs(x), 答;
    if (a < n + 15) {
      答 = 級数(a, n, -1);
    } else {
      /* ★J_0・J_1 を 漸近で 出して 前へ 漸化★ */
      var j0 = J漸近(a, 0), j1 = J漸近(a, 1);
      if (n === 0) 答 = j0;
      else if (n === 1) 答 = j1;
      else {
        var 前 = j0, 今 = j1;
        for (var k = 1; k < n; k++) { var 次 = (2 * k / a) * 今 - 前; 前 = 今; 今 = 次; }
        答 = 今;
      }
    }
    return (x < 0 && n % 2) ? -答 : 答;
  }
  function BESSELK(x, n) {
    if (!(x > 0)) return null;      /* ★x ≤ 0 は #NUM!★（呼ぶ 側が 札を 付けます） */
    var k0, k1;
    /* ★境目 10★ … ★10 より 小さいと 漸近が 足りず、大きいと 級数が 桁落ちします★
         （実測 … 境目を 2 に すると `=BESSELK(2.5,1)` が 1.27e-3 ずれました） */
    if (x < 10) { k0 = K0級数(x); k1 = K1級数(x); }
    else { k0 = K漸近(x, 0); k1 = K漸近(x, 1); }
    if (n === 0) return k0;
    if (n === 1) return k1;
    var 前 = k0, 今 = k1;
    for (var i = 1; i < n; i++) { var 次 = 前 + (2 * i / x) * 今; 前 = 今; 今 = 次; }
    return 今;
  }
  function BESSELY(x, n) {
    if (!(x > 0)) return null;
    var y0, y1;
    if (x < 12) { y0 = Y0級数(x); y1 = Y1級数(x); }
    else { y0 = Y漸近(x, 0); y1 = Y漸近(x, 1); }
    if (n === 0) return y0;
    if (n === 1) return y1;
    var 前 = y0, 今 = y1;
    for (var i = 1; i < n; i++) { var 次 = (2 * i / x) * 今 - 前; 前 = 今; 今 = 次; }
    return 今;
  }

  return {
    BESSELI: BESSELI, BESSELJ: BESSELJ, BESSELK: BESSELK, BESSELY: BESSELY,
    名前たち: function () { return ['BESSELI', 'BESSELJ', 'BESSELK', 'BESSELY']; }
  };
}));
