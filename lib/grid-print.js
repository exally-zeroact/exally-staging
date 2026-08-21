/* grid-print.js — ★紙に刷る物を作る（純関数）★
 *
 *  ★実Excel 16.0.20228 を COM で読んで測った既定（2026-08-21）★
 *    向き=縦（xlPortrait）／用紙=A4／倍率=100％
 *    余白 … 左右 50.4pt（1.78cm）／上下 54pt（1.9cm）／ヘッダー・フッター 21.6pt
 *    ★枠線は刷らない★（PrintGridlines=False）／★行と列の番号も刷らない★（PrintHeadings=False）
 *    ページの順＝下へ→右／中央には置かない
 *    120行×20列を A4縦・既定で刷ると ★1ページ 39行 × 8列★
 *      （切れ目が 40・79・118行目／9・17列目）
 *
 *  ★真似られない事（測って分かった）★
 *    ★1ページに何行 入るかは 印刷する機械（プリンタ）で変わる★。
 *    A4の紙の大きさと余白から計算すると 41行 入るはずだが 実Excelは 39行だった＝
 *    プリンタの「刷れない縁」のぶん。★うちはブラウザに刷ってもらうので ページの切れ目は
 *    Excel と同じにはならない★。ここは ★真似ない（真似られない）★ と決めた。
 *
 *  ★うちの方が良い所（違いとして残す）★
 *    Excel は ★列の幅より長い字が 紙で欠ける★（実物で何度も踏んだ）。
 *    うちは ★表で刷るので 折り返して 全部 出る★＝★紙で字が消えない★。
 */
(function (root) {
  'use strict';

  function colName(i) {
    var s = '', n = i;
    do { s = String.fromCharCode(65 + (n % 26)) + s; n = Math.floor(n / 26) - 1; } while (n >= 0);
    return s;
  }

  function esc(s) {
    return String(s === undefined || s === null ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  /** 紙に出す字＝★画面に出ている字★（式なら答え） */
  function shown(cell) {
    if (!cell) return '';
    var v = (cell.d !== undefined && cell.d !== null && cell.d !== '') ? cell.d : cell.v;
    if (v === undefined || v === null) return '';
    var s = String(v);
    /* 式のまま残っているセル（まだ計算していない）は 紙に式を出さない */
    if (s.charAt(0) === '=') return '';
    return s;
  }

  /** 中身が在る一番外（何も無ければ null） */
  function usedRange(data) {
    var maxR = -1, maxC = -1;
    for (var k in data) {
      var cell = data[k];
      if (shown(cell) === '') continue;
      var p = k.split(','), r = +p[0], c = +p[1];
      if (r > maxR) maxR = r;
      if (c > maxC) maxC = c;
    }
    if (maxR < 0) return null;
    return { r1: 0, c1: 0, r2: maxR, c2: maxC };
  }

  /**
   * 紙だけの窓に書く HTML を作る
   * @param {Object} o {sheetName, data, range, fileName, 向き, 枠線, 行列番号}
   * @returns {string|null} 中身が無ければ null（★白紙の印刷ダイアログを出さない★）
   */
  function buildHtml(o) {
    o = o || {};
    var data = o.data || {};
    var rng = o.range || usedRange(data);
    if (!rng) return null;                       /* ★下絵0枚では刷らない★ */
    var 縦 = (o.向き 　!== 'landscape');
    var 枠線 = !!o.枠線;                          /* ★既定は刷らない（実測）★ */
    var 行列番号 = !!o.行列番号;                  /* ★既定は刷らない（実測）★ */

    var 行 = [];
    if (行列番号) {
      var h = ['<th class="hd"></th>'];
      for (var hc = rng.c1; hc <= rng.c2; hc++) h.push('<th class="hd">' + colName(hc) + '</th>');
      行.push('<tr>' + h.join('') + '</tr>');
    }
    var 何か出た = false;
    for (var r = rng.r1; r <= rng.r2; r++) {
      var tds = [];
      if (行列番号) tds.push('<th class="hd">' + (r + 1) + '</th>');
      for (var c = rng.c1; c <= rng.c2; c++) {
        var cell = data[r + ',' + c];
        var t = shown(cell);
        if (t !== '') 何か出た = true;
        var st = [];
        if (cell) {
          if (cell.bold) st.push('font-weight:700');
          if (cell.italic) st.push('font-style:italic');
          if (cell.underline) st.push('text-decoration:underline');
          if (cell.color) st.push('color:' + cell.color);
          if (cell.bgColor) st.push('background:' + cell.bgColor);
          if (cell.align) st.push('text-align:' + cell.align);
          else if (t !== '' && !isNaN(Number(String(t).replace(/,/g, '')))) st.push('text-align:right');
          if (cell.fontSize) st.push('font-size:' + cell.fontSize + 'px');
          if (cell.border) st.push('border:1px solid #333333');
        }
        tds.push('<td' + (st.length ? ' style="' + st.join(';') + '"' : '') + '>' + esc(t) + '</td>');
      }
      行.push('<tr>' + tds.join('') + '</tr>');
    }
    if (!何か出た) return null;                  /* ★字が1つも無いなら 刷らない★ */

    /* ★余白は 実Excel の既定と同じ★（上下 1.90cm／左右 1.78cm） */
    var css = [
      '@page { size: A4 ' + (縦 ? 'portrait' : 'landscape') + '; margin: 1.9cm 1.78cm; }',
      'html,body { margin:0; padding:0; }',
      'body { font-family:"Noto Sans JP","Hiragino Kaku Gothic ProN",sans-serif; color:#1A1A1A; font-size:11px; }',
      'table { border-collapse:collapse; width:100%; table-layout:fixed; }',
      'td,th { padding:2px 4px; vertical-align:middle; word-break:break-all; ' +
        (枠線 ? 'border:1px solid #BBBBBB;' : '') + ' }',
      'th.hd { background:#EFEFEF; color:#1A1A1A; font-weight:700; text-align:center; border:1px solid #BBBBBB; }',
      'tr { page-break-inside:avoid; }',
      '.ttl { font-size:12px; font-weight:700; margin:0 0 6px; }',
    ].join('\n');

    var 題 = esc(o.fileName || '') + (o.sheetName ? '　' + esc(o.sheetName) : '');
    return '<!doctype html><html lang="ja"><head><meta charset="utf-8">'
      + '<title>' + 題 + '</title><style>' + css + '</style></head><body>'
      + (題 ? '<div class="ttl">' + 題 + '</div>' : '')
      + '<table>' + 行.join('') + '</table>'
      + '</body></html>';
  }

  var api = { buildHtml: buildHtml, usedRange: usedRange, shown: shown, colName: colName };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.GridPrint = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
