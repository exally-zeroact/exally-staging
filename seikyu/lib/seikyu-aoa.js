/* seikyu-aoa.js — ★Excelに出す時の中身（行と列の並び・列幅）を作る唯一の場所★
 * ==============================================================================
 * ここは ★表の中身を作るだけ★。SheetJS も Blob も触らない（＝素のNodeで全部測れる）。
 * 実際に .xlsx を組むのは seikyu/js/seikyu-xlsx.js。
 *
 * ★渡した相手の画面で ######## にしない★
 *   日付や金額の列に幅を付けずに出すと、相手のExcelで列が狭く ######## になる（既知の前科）。
 *   だから列幅(cols)をここで必ず返す。
 *
 * ★数は数のまま出す★
 *   金額を "1,234円" のような文字で出すと、相手が足し算できない＝Excelで渡す意味が消える。
 *   桁区切りは書式(z)で付ける。
 *
 * ★税率の数字を1つも書かない★（区分は totals.byRate をそのまま並べる）
 *
 * 【利用】ブラウザ window.SeikyuAoa ／ Node require('./seikyu-aoa.js')
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.SeikyuAoa = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var YEN_FMT = '#,##0';        // 金額（桁区切り・小数なし）
  var NUM_FMT = '#,##0.###';    // 数量（0.5 のような端数も出す）

  function jpDate(ymd) {
    var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(ymd || ''));
    if (!m) return '';
    return +m[1] + '年' + (+m[2]) + '月' + (+m[3]) + '日';
  }
  function honorOf(p) {
    var h = (p && (p.honor || p.keisho)) || '';
    if (!h || h === '（なし）' || h === '(なし)' || h === 'なし') return '';
    return h;
  }

  /**
   * build({ inv, tax, partner, org }) → { aoa, cols, numFmt, name }
   *   aoa    … 行の配列（値は数のまま）
   *   cols   … 列幅（[{wch:…}]）★これが無いと相手の画面で ######## になる★
   *   numFmt … [{r,c,z}] 数の書式（桁区切り）。呼ぶ側がセルに当てる
   *   name   … シート名
   */
  function build(o) {
    o = o || {};
    var inv = o.inv || {};
    var tax = o.tax || {};
    var p = o.partner || {};
    var g = o.org || {};
    var isQuote = (inv.doc_type === 'quote');
    var heading = isQuote ? '見積書' : '請求書';

    var aoa = [];
    var numFmt = [];
    function push(row) { aoa.push(row); return aoa.length - 1; }
    function money(r, c) { numFmt.push({ r: r, c: c, z: YEN_FMT }); }
    function qty(r, c) { numFmt.push({ r: r, c: c, z: NUM_FMT }); }

    push([heading]);
    push([]);
    push([(isQuote ? '見積番号' : '請求番号'), inv.no || '（未採番）']);
    push([(isQuote ? '見積日' : '請求日'), jpDate(inv.issue_ymd) || '（未入力）']);
    if (inv.due_ymd) push(['お支払期限', jpDate(inv.due_ymd)]);
    push([]);
    push(['宛先', (p.name || '（取引先が未選択）') + (honorOf(p) ? ' ' + honorOf(p) : '')]);
    if (p.addr) push(['', p.addr]);
    push(['自社', g.yago || '（自社情報が未入力）']);
    if (g.addr) push(['', g.addr]);
    if (g.tel) push(['', 'TEL ' + g.tel]);
    if (g.invoiceNo) push(['', '登録番号 ' + g.invoiceNo]);
    push([]);
    var gr = push([(isQuote ? 'お見積金額' : 'ご請求金額'), Number(tax.grandTotal) || 0]);
    money(gr, 1);
    if (inv.data && inv.data.subject) push(['件名', inv.data.subject]);
    push([]);

    push(['#', '品名・内容', '数量', '単位', '単価', '金額', '税率']);
    var lines = Array.isArray(tax.lines) ? tax.lines : [];
    for (var i = 0; i < lines.length; i++) {
      var ln = lines[i];
      var r = push([
        i + 1,
        ln.name || '',
        (ln.qty === undefined || ln.qty === null || ln.qty === '') ? '' : Number(ln.qty),
        ln.unit || '',
        (ln.price === undefined || ln.price === null || ln.price === '') ? '' : Number(ln.price),
        Number(ln.amount) || 0,
        Number(ln.rate) === 0 ? '対象外' : (Number(ln.rate) + '%'),
      ]);
      qty(r, 2); money(r, 4); money(r, 5);
    }
    if (!lines.length) push(['', '明細がまだ1行もありません', '', '', '', '', '']);

    push([]);
    var rs = push(['', '', '', '', '小計', Number(tax.subtotal) || 0, '']); money(rs, 5);
    var rt = push(['', '', '', '', '消費税', Number(tax.taxTotal) || 0, '']); money(rt, 5);
    var rg = push(['', '', '', '', '合計', Number(tax.grandTotal) || 0, '']); money(rg, 5);

    push([]);
    push(['区分', '対象額', '消費税']);
    var byRate = Array.isArray(tax.byRate) ? tax.byRate : [];
    for (var k = 0; k < byRate.length; k++) {
      var b = byRate[k];
      var rr = push([Number(b.pct) + '% 対象', Number(b.base) || 0, Number(b.tax) || 0]);
      money(rr, 1); money(rr, 2);
    }
    var ex = (tax.exempt && Number(tax.exempt.base)) || 0;
    if (ex !== 0) { var re = push(['消費税の対象外', ex, '']); money(re, 1); }
    if (!byRate.length && ex === 0) push(['区分はまだありません', '', '']);

    if (g.bank) { push([]); push(['お振込先', g.bank]); }
    if (inv.data && inv.data.memo) { push([]); push(['備考', inv.data.memo]); }

    /* ★列幅★ 品名は広く・金額は桁が入るだけ・日付や長い文字が入る列は余裕を取る */
    var cols = [
      { wch: 6 },   // #
      { wch: 40 },  // 品名・内容
      { wch: 8 },   // 数量
      { wch: 6 },   // 単位
      { wch: 12 },  // 単価
      { wch: 14 },  // 金額
      { wch: 10 },  // 税率
    ];

    return { aoa: aoa, cols: cols, numFmt: numFmt, name: heading, YEN_FMT: YEN_FMT, NUM_FMT: NUM_FMT };
  }

  return { build: build, YEN_FMT: YEN_FMT, NUM_FMT: NUM_FMT, jpDate: jpDate, honorOf: honorOf };
});
