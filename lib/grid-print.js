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
   * @param {Object} o {sheetName, data, range, fileName, 向き, 枠線, 行列番号, 余白, 用紙}
   *   余白 … {上,下,左,右}（cm）★無ければ 実Excelの 既定（上下1.905 / 左右1.778）★
   *   用紙 … 'A4'（既定）/'A3'/'A5'/'B5'/'Letter'（実測＝A4=9・A3=8・A5=11・B5=13・Letter=1）
   * @returns {string|null} 中身が無ければ null（★白紙の印刷ダイアログを出さない★）
   */
  function buildHtml(o) {
    o = o || {};
    var data = o.data || {};
    var rng = o.range || usedRange(data);
    if (!rng) return null;                       /* ★下絵0枚では刷らない★ */
    var 縦 = (o.向き 　!== 'landscape');
    /* ★余白と 用紙（実Excelから 測った 既定）★ */
    var 余 = o.余白 || { 上: 1.905, 下: 1.905, 左: 1.778, 右: 1.778 };
    var 紙 = ({ A4: 'A4', A3: 'A3', A5: 'A5', B5: 'B5', Letter: 'letter' })[o.用紙 || 'A4'] || 'A4';
    /* ★刷る 倍率★（実Excel実測＝既定 100・10〜400。null＝ページに 合わせる） */
    var 倍率 = (o.印刷倍率 === null || o.印刷倍率 === undefined) ? 100 : Number(o.印刷倍率);
    if (!isFinite(倍率) || 倍率 < 10) 倍率 = 10;
    if (倍率 > 400) 倍率 = 400;
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
      '@page { size: ' + 紙 + ' ' + (縦 ? 'portrait' : 'landscape') + '; margin: '
        + 余.上 + 'cm ' + 余.右 + 'cm ' + 余.下 + 'cm ' + 余.左 + 'cm; }',
      'html,body { margin:0; padding:0; }',
      /* ★倍率は 本当に 効かせる★（100 の 時は 何も しない） */
      (倍率 === 100 ? '' : ('body > * { zoom: ' + (倍率 / 100) + '; }')),
      'body { font-family:"Noto Sans JP","Hiragino Kaku Gothic ProN",sans-serif; color:#1A1A1A; font-size:11px; }',
      'table { border-collapse:collapse; width:100%; table-layout:fixed; }',
      'td,th { padding:2px 4px; vertical-align:middle; word-break:break-all; ' +
        (枠線 ? 'border:1px solid #BBBBBB;' : '') + ' }',
      'th.hd { background:#EFEFEF; color:#1A1A1A; font-weight:700; text-align:center; border:1px solid #BBBBBB; }',
      'tr { page-break-inside:avoid; }',
      '.ttl { font-size:12px; font-weight:700; margin:0 0 6px; }',
      /* ★紙の 上と 下に 入れる 字（ヘッダー／フッター）★ 2026-08-31 */
      '.pgbox { position:relative; overflow:hidden; }',
      '.pgbox + .pgbox { page-break-before: always; }',
      '.hf { display:flex; font-size:10px; color:#333333; }',
      '.hf > div { flex:1; white-space:pre; overflow:hidden; }',
      '.hf .c { text-align:center; }',
      '.hf .r { text-align:right; }',
      '.hfH { margin:0 0 4px; }',
      '.hfF { margin:4px 0 0; }',
    ].join('\n');

    var 題 = esc(o.fileName || '') + (o.sheetName ? '　' + esc(o.sheetName) : '');
    return '<!doctype html><html lang="ja"><head><meta charset="utf-8">'
      + '<title>' + 題 + '</title><style>' + css + '</style></head><body>'
      + (題 ? '<div class="ttl">' + 題 + '</div>' : '')
      + '<table>' + 行.join('') + '</table>'
      + 上下の字(o, 余, 縦)
      + '</body></html>';
  }

  /** ★紙の 上と 下に 入れる 字を 割り付ける★ 2026-08-31
   *
   *  ★なぜ 刷る 窓の 中で やるか★
   *    ★何ページに なるかは 中身を 並べてみないと 分からない★。
   *    そして ★実ブラウザで 測った通り★：
   *      ・position:fixed は 毎ページ 出る（全ページ 上端 y=11）
   *      ・★CSS の counter(page) は 増えない★（10ページ目も「1」のまま）
   *    ⇒ ★自分で 行を 分けて 1枚ずつ 番号を 書く★。
   *
   *  ★決め打ちしない★＝帯の 高さも 行の 高さも ★出してから 測る★。
   *
   *  見張り: tests/header-footer.test.mjs
   */
  function 上下の字(o, 余, 縦) {
    var h = o.ヘッダー, f = o.フッター;
    var 有 = function (v) { return v && (v.左 || v.中 || v.右); };
    if (!有(h) && !有(f)) return '';               /* ★書いていなければ 何もしない★ */
    var 紙 = { A4: [21.0, 29.7], A3: [29.7, 42.0], A5: [14.8, 21.0],
      B5: [18.2, 25.7], Letter: [21.59, 27.94] }[o.用紙 || 'A4'] || [21.0, 29.7];
    var 高さcm = 縦 ? 紙[1] : 紙[0];
    var 設定 = {
      h: 有(h) ? h : null, f: 有(f) ? f : null,
      P: o.頁印 || 'P', N: o.総印 || 'N',
      使える: (高さcm - 余.上 - 余.下) * (96 / 2.54)
    };
    return '<' + 'script>(' + String(割り付ける) + ')(' + JSON.stringify(設定) + ');<' + '/script>';
  }

  /** ★刷る 窓の 中で 走る★（文字にして 送り込むので ★外の 物を 使わない★） */
  function 割り付ける(S) {
    var tbl = document.querySelector('table');
    if (!tbl) return;
    var 帯 = function (o, cls) {
      var d = document.createElement('div');
      d.className = 'hf ' + cls;
      ['左', '中', '右'].forEach(function (k, i) {
        var e = document.createElement('div');
        e.className = (i === 1 ? 'c' : i === 2 ? 'r' : 'l');
        e.textContent = (o && o[k]) || '';
        d.appendChild(e);
      });
      return d;
    };
    /* ★帯の 高さは 出してから 測る★（決め打ちしない） */
    var 見本 = 帯(S.h || S.f, 'hfH');
    document.body.appendChild(見本);
    var 帯高 = 見本.getBoundingClientRect().height;
    document.body.removeChild(見本);

    var 使える = S.使える - (S.h ? 帯高 + 4 : 0) - (S.f ? 帯高 + 4 : 0);
    if (!(使える > 40)) return;                     /* ★入らないなら 何もしない★ */

    var rows = [].slice.call(tbl.rows);
    var 組 = [], 今 = [], 高 = 0;
    for (var i = 0; i < rows.length; i++) {
      var rh = rows[i].getBoundingClientRect().height;
      if (今.length && 高 + rh > 使える) { 組.push(今); 今 = []; 高 = 0; }
      今.push(rows[i]); 高 += rh;
    }
    if (今.length) 組.push(今);
    var N = 組.length;
    var 入れる = function (t, p) {
      return String(t || '').split(S.P).join(String(p)).split(S.N).join(String(N));
    };
    var 写す = function (o, p) {
      return o ? { 左: 入れる(o.左, p), 中: 入れる(o.中, p), 右: 入れる(o.右, p) } : null;
    };
    var 入れ物 = document.createElement('div');
    for (var g = 0; g < 組.length; g++) {
      var box = document.createElement('div');
      box.className = 'pgbox';
      if (S.h) box.appendChild(帯(写す(S.h, g + 1), 'hfH'));
      var t2 = document.createElement('table');
      var tb = document.createElement('tbody');
      for (var k = 0; k < 組[g].length; k++) tb.appendChild(組[g][k]);
      t2.appendChild(tb);
      box.appendChild(t2);
      if (S.f) box.appendChild(帯(写す(S.f, g + 1), 'hfF'));
      入れ物.appendChild(box);
    }
    tbl.parentNode.replaceChild(入れ物, tbl);
  }

  var api = { buildHtml: buildHtml, usedRange: usedRange, shown: shown, colName: colName };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.GridPrint = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
