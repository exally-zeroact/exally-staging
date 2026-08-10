/* seikyu-tax.js — 請求書の消費税（計算の唯一の正）
 * ==============================================================================
 * ★法律で決まっている所（選べない）★  出典: 国税庁 適格請求書等保存方式（インボイス）
 *   ・「一の適格請求書につき、★税率ごとに1回の端数処理★を行う必要があります」
 *   ・「個々の商品ごとに消費税額等を計算し、1円未満の端数処理を行い、その合計額を
 *      消費税額等として記載することは★認められません★」
 *   → ★税率ごとに合計を出してから、そこで1回だけ丸める★
 *   → ★8%が1行でもあれば、10%と8%を別々に集計して別々に丸める★
 *
 * ★会社が選べる所★
 *   ・丸め方（切捨て floor / 切上げ ceil / 四捨五入 round）… 法律は任意としている
 *   ・内税(inclusive) か 外税(exclusive) か
 *
 * ★率の数字はこのファイルに1つも書かない★
 *   kyuyo/lib/shouhizei-ritsu.js（唯一の正）から取る。率が変わってもここは直さない。
 *
 * ★行ごとの税額(taxRef)も返すが、それは「参考」。合計をそこから作らない。★
 *   （紙に行ごとの税額を載せるのは自由。載せても合計はここで出した額を使う）
 *
 * ★画面に依らない（DOMを1つも触らない）★＝素のNodeで全パターン回せる。
 *
 * 【利用】ブラウザ window.SeikyuTax（先に shouhizei-ritsu.js を読む）
 *         Node    require('./seikyu-tax.js')
 */
(function (root, factory) {
  var SR = null;
  if (typeof module === 'object' && module.exports) {
    SR = require('../../kyuyo/lib/shouhizei-ritsu.js');
    module.exports = factory(SR);
  } else {
    root.SeikyuTax = factory(root.ShouhizeiRitsu);
  }
})(typeof self !== 'undefined' ? self : this, function (SR) {
  'use strict';

  var ROUNDINGS = ['floor', 'ceil', 'round'];
  var TAX_MODES = ['inclusive', 'exclusive'];
  var MAX_LINES = 1000; // ★1通の蓋。超えたら赤で止める（黙って切らない）

  /* 率(%)の一覧を「唯一の正」から作る。0=消費税がかからない行（非課税・対象外）。
     ★ここで 10 や 8 と書かない。書いた瞬間、法が変わった日に嘘をつく。 */
  function pctOf(v) { return Math.round(Number(v) * 10000) / 100; }
  function defaultRates() {
    if (!SR || typeof SR.hyojun !== 'number' || typeof SR.keigen !== 'number') return null;
    return [pctOf(SR.hyojun), pctOf(SR.keigen), 0];
  }

  /* 円未満の丸め。★符号対称＝絶対値で丸めて符号を戻す★
     理由: 値引き行(マイナス)で「切捨て」が税を増やす向きに働くと、引いたはずが増える。
     切捨て=0に近づく / 切上げ=0から遠ざかる、で表裏を揃える。 */
  function roundYen(x, mode) {
    var s = x < 0 ? -1 : 1, a = Math.abs(x);
    var r = mode === 'ceil' ? Math.ceil(a) : mode === 'round' ? Math.floor(a + 0.5) : Math.floor(a);
    return s * r;
  }

  /* 行の金額。★金額を直接打った行はその額をそのまま使う（勝手に作り直さない）★
     数量×単価しか無い行だけ、会社の丸め方で円にする。 */
  function amountOf(ln, rounding) {
    if (ln.amount !== undefined && ln.amount !== null && ln.amount !== '') return Number(ln.amount);
    if (ln.qty === undefined || ln.price === undefined) return 0;
    return roundYen(Number(ln.qty) * Number(ln.price), rounding);
  }

  function zero(errors) {
    return {
      ok: false, errors: errors, lines: [], byRate: [],
      exempt: { base: 0 }, hasReduced: false,
      subtotal: 0, taxTotal: 0, grandTotal: 0,
    };
  }

  /**
   * compute({ lines, taxMode, rounding }, opts?)
   *   lines[i] = { name, qty?, price?, amount?, rate }   rate は % の数（10 / 8 / 0）
   *   opts.rates … 率の一覧を差し替える（テスト用。null を渡すと投げる＝出どころ無しで動かさない）
   * 返り = { ok, errors[], lines[], byRate[], exempt, hasReduced, subtotal, taxTotal, grandTotal }
   */
  function compute(input, opts) {
    opts = opts || {};
    var rates = Object.prototype.hasOwnProperty.call(opts, 'rates') ? opts.rates : defaultRates();
    if (!rates || !rates.length) {
      throw new Error('消費税率の出どころがありません（kyuyo/lib/shouhizei-ritsu.js を先に読んでください）');
    }
    var inp = input || {};
    var lines = inp.lines || [];
    var taxMode = inp.taxMode;
    var rounding = inp.rounding;
    var errors = [];

    if (TAX_MODES.indexOf(taxMode) < 0) errors.push('税の入れ方が不明です（' + taxMode + '）。内税か外税かを選んでください');
    if (ROUNDINGS.indexOf(rounding) < 0) errors.push('丸め方が不明です（' + rounding + '）。切捨て・切上げ・四捨五入から選んでください');
    if (!Array.isArray(lines)) errors.push('明細がありません');
    else if (lines.length > MAX_LINES) errors.push('明細が' + lines.length + '行あります。1通あたり' + MAX_LINES + '行までです（分けて出してください）');
    if (errors.length) return zero(errors);

    // ── 行を読む（★税率が空・知らない値なら赤。黙って標準税率に寄せない★）
    var out = [];
    for (var i = 0; i < lines.length; i++) {
      var ln = lines[i] || {};
      var rate = ln.rate;
      if (rate === undefined || rate === null || rate === '') {
        errors.push((i + 1) + '行目「' + (ln.name || '') + '」の税率が空です（税率を選んでください）');
        continue;
      }
      rate = Number(rate);
      if (rates.indexOf(rate) < 0) {
        errors.push((i + 1) + '行目「' + (ln.name || '') + '」の税率 ' + ln.rate + ' は扱えません（使えるのは ' + rates.join('% / ') + '%）');
        continue;
      }
      var amount = amountOf(ln, rounding);
      if (!Number.isFinite(amount)) { errors.push((i + 1) + '行目の金額が数になりません'); continue; }
      if (!Number.isInteger(amount)) { errors.push((i + 1) + '行目の金額に円未満が残っています（' + amount + '）'); continue; }
      out.push({ index: i, name: ln.name || '', qty: ln.qty, unit: ln.unit, price: ln.price, amount: amount, rate: rate, taxRef: 0 });
    }
    if (errors.length) return zero(errors);

    // ── 税率ごとに束ねる（★ここで初めて丸める。1回だけ★）
    var buckets = {};
    var exemptBase = 0;
    for (var j = 0; j < out.length; j++) {
      var r = out[j];
      if (r.rate === 0) { exemptBase += r.amount; continue; }
      if (!buckets[r.rate]) buckets[r.rate] = { pct: r.rate, sum: 0, rows: [] };
      buckets[r.rate].sum += r.amount;
      buckets[r.rate].rows.push(r);
    }

    var byRate = [];
    var taxTotal = 0, taxableBase = 0;
    var pcts = Object.keys(buckets).map(Number).sort(function (a, b) { return b - a; });
    for (var k = 0; k < pcts.length; k++) {
      var b = buckets[pcts[k]];
      var num = Math.round(b.pct * 100); // 10% → 1000（整数で持つ＝小数の誤差を作らない）
      var base, tax, gross;
      if (taxMode === 'inclusive') {
        // ★税込の合計から割り戻す（税率ごとに1回）
        gross = b.sum;
        tax = roundYen(gross * num / (10000 + num), rounding);
        base = gross - tax;
      } else {
        // ★税抜の合計に掛ける（税率ごとに1回）
        base = b.sum;
        tax = b.rows.reduce(function (t, r) { return t + roundYen(r.amount * num / 10000, rounding); }, 0); // わざと壊す
        gross = base + tax;
      }
      byRate.push({ pct: b.pct, base: base, tax: tax, gross: gross });
      taxTotal += tax;
      taxableBase += base;
      // 行ごとの税額＝★参考★（紙に載せてもよいが、合計はここから作らない）
      for (var q = 0; q < b.rows.length; q++) {
        var row = b.rows[q];
        row.taxRef = taxMode === 'inclusive'
          ? roundYen(row.amount * num / (10000 + num), rounding)
          : roundYen(row.amount * num / 10000, rounding);
      }
    }

    var subtotal = taxableBase + exemptBase;
    var redPct = rates.length > 1 ? rates[1] : null;
    return {
      ok: true, errors: [],
      lines: out,
      byRate: byRate,
      exempt: { base: exemptBase },
      hasReduced: redPct !== null && byRate.some(function (x) { return x.pct === redPct; }),
      subtotal: subtotal,
      taxTotal: taxTotal,
      grandTotal: subtotal + taxTotal,
    };
  }

  return {
    compute: compute,
    roundYen: roundYen,
    rates: defaultRates,
    ROUNDINGS: ROUNDINGS,
    TAX_MODES: TAX_MODES,
    MAX_LINES: MAX_LINES,
  };
});
