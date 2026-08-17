/* typed-value.js — ★セルに「打ち込んだ文字」を Excel と同じように解釈する★
 *
 * ═══ なぜ要るのか（E3・期限 2026-09-30・★金が落ちる★）═══════════════
 *   `1,234` と打つと ★文字列のまま★になり、=SUM() が拾わない。
 *   ★#ERROR は出ない。合計が黙って小さくなる★＝いちばん気づけない壊れ方。
 *   実Excelは `1,234` を ★数 1234★ にして、見た目だけ `1,234` のままにする。
 *
 * ═══ 実Excelで測った真値（2026-08-18・Excel 365 16.0.20228・日本語UI）═══
 *   打った物        Excelの型  中の値     画面の字     Excelが当てた書式
 *   1,234           数         1234       1,234        #,##0
 *   1,23            ★文字★    -          1,23         標準     ← 3桁未満の組は文字のまま
 *   12,34           ★文字★    -          12,34        標準
 *   1,2345          数         ★12345★   12,345       #,##0    ← 4桁の組は通る
 *   1,234,567       数         1234567    1,234,567    #,##0
 *   1,234.5         数         1234.5     1,234.50     #,##0.00
 *   -1,234          数         -1234      -1,234       #,##0
 *   +1,234          数         1234       1,234        #,##0
 *   " 1,234 "       数         1234       1,234        #,##0    ← 前後の空白は落とす
 *   1,234-          ★文字★    -                        標準     ← 後ろのマイナスは通らない
 *   ,  /  1,  /  ,234  /  1,,234    ★全部 文字★
 *   ★１，２３４（全角）★  数   1234       1,234        #,##0    ← 日本語入力そのまま
 *   ★1，234（全角カンマ）★数   1234       1,234        #,##0
 *   ★１２３４（全角数字）★数   1234       1234         標準
 *   １，２３４．５（全角）  数   1234.5     1,234.50     #,##0.00
 *   1 234（空白）   ★文字★    -                        標準
 *   1.234           数         1.234      1.234        標準
 *   0007            数         7          7            標準（今までどおり）
 *
 * ═══ ★今回は直さないと決めた物（Excelの真値だけ記録して残す）★ ═══════
 *   (1,234)   Excel=数 -1234（会計のマイナス）… 括弧はカンマ無しの形を測っていないので触らない
 *   1,234%    Excel=数 12.34（0%）           … 値が100分の1になる＝別の話
 *   $1,234    Excel=数 1234（通貨書式）      … ★¥1,234 は Excel も文字のまま（実測）★＝
 *                                              通貨記号は記号ごとに答えが違う。1つずつ測ってから
 *   1,234E2   Excel=数 123400                … 指数はカンマと混ぜない
 *   ★「たぶんこう」で足さない。測ってから足す。★
 *   （上の5件の Excel の真値は tests/fixtures/typed-value-golden.json に残してある）
 *
 * ═══ 使う側 ═══════════════════════════════════════════════════════
 *   book.html の toHFVal（計算に渡す値）と lib/grid-xlsx.js（xlsxへ書き出す値）が
 *   ★同じ規則でなければならない★（ズレると「画面の合計」と「落としたファイルの合計」が食い違う）。
 *   ★だから規則はこの1本だけに置く。両方がここを呼ぶ。★
 *   守り: tests/typed-value.test.mjs（実Excelの真値と1つずつ突き合わせ＋--self-test）
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.TypedValue = factory();
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  /* 全角 → 半角（★日本語入力でそのまま打たれる★。Excelは全角でも数にする＝実測） */
  function toHalf(s) {
    return String(s).replace(/[０-９]/g, function (ch) {
      return String.fromCharCode(ch.charCodeAt(0) - 0xFEE0);
    })
      .replace(/，/g, ',')
      .replace(/．/g, '.')
      .replace(/＋/g, '+')
      .replace(/－|ー|―/g, '-');
  }

  /* 桁区切りの形か。★組は3桁以上★（1,23 と 12,34 は文字のまま＝実測） */
  var GROUPED = /^[+-]?\d+(?:,\d{3,})+(?:\.\d+)?$/;

  /**
   * 打ち込んだ文字を Excel と同じ解釈にする。
   * @param {*} raw 打った物
   * @returns {{num:number, numFmt:(string|null), grouped:boolean}|null}
   *          数として読めなければ null（＝文字のまま。呼ぶ側は今までどおり扱う）
   */
  function parseTyped(raw) {
    if (typeof raw !== 'string') return null;
    var s = toHalf(raw).trim();
    if (s === '') return null;

    if (s.indexOf(',') >= 0) {
      if (!GROUPED.test(s)) return null;                 // ★形が合わなければ文字のまま★
      var n = Number(s.replace(/,/g, ''));
      if (!isFinite(n)) return null;
      /* 書式は実測どおり。小数が付いていれば #,##0.00（Excelが当てる物と同じ） */
      return { num: n, numFmt: (s.indexOf('.') >= 0 ? '#,##0.00' : '#,##0'), grouped: true };
    }

    /* カンマが無い形。★全角の数字だけ★（１２３４）もここで数になる（実測: Excelは 1234・標準書式）。
       半角だけの物は今までと同じ判定に落ちる＝挙動を変えない。 */
    if (s !== String(raw).trim()) {                      // 全角を直した時だけ見る
      if (s !== '' && !isNaN(s)) {
        var m = Number(s);
        if (isFinite(m)) return { num: m, numFmt: null, grouped: false };
      }
    }
    return null;
  }

  return { parseTyped: parseTyped, toHalf: toHalf, GROUPED: GROUPED };
}));
