/* grid-xlsx.js — グリッド(book.html)の持ち方 → xlsx-io.js が食う形 への変換と、書き出し前の警告。
 *
 *  なぜ分けてあるか:
 *    ・変換と警告は【純関数】にしておけば、ブラウザを立ち上げずにテストできる。
 *    ・HyperFormula には触らない。配列を返す式かどうかは呼び出し側(book.html)が
 *      isArrayFormula コールバックで教える＝このファイルはエンジンに依存しない。
 *
 *  ★書き出しで「落ちる物」は黙って落とさず、何が起きるかを具体的に伝える(セル番地つき)。
 *    実測(2026-08-01)で分かっている落ちる物は2つ:
 *      ① 太字・文字色・背景色・罫線 … 同梱SheetJS CE 0.20.3 が書けない(往復させると消える)
 *         ※表示形式・セルの結合・列幅は【書ける】ので、そこは落とさない。
 *      ② 配列を返す式のスピル      … Excelで開くと1セル分の値になる
 *         (metadata.xml は正しく書けているが、セルの cm 属性を注入してもスピルしなかった＝原因は別。
 *          台帳 file_roundtrip_known)
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.GridXlsx = factory();
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  function colLetter(c) { var s = ''; c += 1; while (c > 0) { s = String.fromCharCode(64 + ((c - 1) % 26 + 1)) + s; c = Math.floor((c - 1) / 26); } return s; }
  function addr(r, c) { return colLetter(c) + (r + 1); }

  /* ★数値かどうかの判定は book.html の toHFVal と【同じ規則】にする。ここがズレると、
   *  落としたファイルの =SUM(...) がグリッドの画面と違う数字を出す＝一番やってはいけない壊れ方になる。
   *    toHFVal: 文字列で trim()!=='' かつ !isNaN(v) なら parseFloat(v)
   *  この規則の帰結（どちらもグリッドと同じになるので正しい）:
   *    ・'0007' → 7。グリッドの画面は「0007」と出るが、エンジンは既に 7 として計算している。
   *      ここで文字列のまま書くと Excel の合計だけがグリッドと食い違う。実Excelも打ち込んだ 0007 は 7 にする。
   *    ・'1,234' → isNaN なので文字列のまま。グリッドのエンジンも文字列として扱う（台帳 R11）。
   * ★2026-08-05 変更: '2026-07-31' のような【日付】だけは、Excel と同じ「数（シリアル値）」にする。
   *   理由: グリッド側で日付が日付として扱われず、+30 が 2056 になる等の不具合があったため
   *   グリッドを Excel と同じ扱いに直した。この規則の目的は「画面と書き出しを一致させる」ことなので、
   *   書き出しも同じ数にするのが正しい。日付でない物（007-1234 等）は今までどおり文字のまま。
   */
  /* ★日付は Excel と同じ「1899-12-30 から数えて何日目か」の数にする。
     book.html の dateSerial / toHFVal と ★同じ規則★でないと、
     画面は日付なのに落としたファイルはただの文字、という食い違いが出る。
     守り: tests/grid-date.test.mjs（両方を実際に動かして突き合わせる）。 */
  function dateSerial(y, mo, d) {
    return Math.round((Date.UTC(y, mo - 1, d) - Date.UTC(1899, 11, 30)) / 86400000);
  }
  function parseDateStr(s) {
    if (typeof s !== 'string') return null;
    var m = s.match(/^(\d{4})([/-])(\d{1,2})\2(\d{1,2})$/);
    if (!m) return null;
    var y = parseInt(m[1], 10), mo = parseInt(m[3], 10), d = parseInt(m[4], 10);
    if (mo < 1 || mo > 12 || d < 1 || d > 31) return null;
    var dt = new Date(y, mo - 1, d);
    if (dt.getFullYear() !== y || dt.getMonth() !== mo - 1 || dt.getDate() !== d) return null;
    return { y: y, m: mo, d: d };
  }
  /* 打った形（/ か -）に合わせた日付の表示形式。日付でなければ null。 */
  function dateFmtFor(raw) {
    if (typeof raw !== 'string') return null;
    var pd = parseDateStr(raw);
    if (!pd) return null;
    return raw.indexOf('-') >= 0 ? 'yyyy-mm-dd' : 'yyyy/m/d';
  }
  function asValue(raw) {
    if (raw === null || raw === undefined) return '';
    var s = String(raw);
    if (s === '') return '';
    if (s === 'TRUE') return true;
    if (s === 'FALSE') return false;
    var pd = parseDateStr(s);
    if (pd) return dateSerial(pd.y, pd.m, pd.d);
    if (s.trim() !== '' && !isNaN(s)) return parseFloat(s);
    return s;
  }
  function typeOf(v) { return typeof v === 'number' ? 'n' : (typeof v === 'boolean' ? 'b' : 's'); }

  var STYLE_KEYS = ['bold', 'italic', 'underline', 'strike', 'color', 'bgColor', 'border'];
  function hasLostStyle(cell) {
    for (var i = 0; i < STYLE_KEYS.length; i++) {
      var v = cell[STYLE_KEYS[i]];
      if (v !== undefined && v !== null && v !== false && v !== '') return true;
    }
    return false;
  }

  /* グリッドの sheets[] → xlsx-io.writeBook が食う book
     グリッド: sheets[i] = { name, data: { 'r,c': {v,f,d,numFmt,merged,mergeEnd,…} }, colW: {ci:px} }
     ★セルの持ち方: f が '=' 始まりなら式。そうでなければ v が打ち込んだ生の値、d が表示。 */
  function gridToBook(sheets) {
    return {
      sheets: (sheets || []).map(function (sh) {
        var cells = {}, merges = [], cols = [];
        Object.keys(sh.data || {}).forEach(function (key) {
          var p = key.split(','), r = parseInt(p[0], 10), c = parseInt(p[1], 10);
          if (isNaN(r) || isNaN(c)) return;
          var cell = sh.data[key] || {};
          var isF = typeof cell.f === 'string' && cell.f.charAt(0) === '=';
          var out = {};
          if (isF) {
            out.f = cell.f;
            var cached = asValue(cell.d);                  // 計算済みの値(Excelは開いた瞬間これを見せ、再計算で上書きする)
            out.v = cached; out.t = typeOf(cached);
          } else {
            var val = asValue(cell.v !== undefined && cell.v !== '' ? cell.v : cell.d);
            if (val === '' ) return;                        // 空セルは書かない
            out.v = val; out.t = typeOf(val);
          }
          if (cell.numFmt) out.z = cell.numFmt;             // ★表示形式は書ける(実測)
          /* ★日付は「数」で書くので、表示形式が無いと Excel で 46265 という裸の数字に見える。
             人が書式を決めていない時だけ、打った形（2026/8/31 か 2026-08-31）に合わせて付ける。
             ★人が選んだ書式は上書きしない。★ */
          else if (!isF && dateFmtFor(cell.v !== undefined && cell.v !== '' ? cell.v : cell.d)) {
            out.z = dateFmtFor(cell.v !== undefined && cell.v !== '' ? cell.v : cell.d);
          }
          cells[addr(r, c)] = out;
          if (cell.merged && cell.mergeEnd) {
            merges.push({ s: { r: cell.merged.r, c: cell.merged.c }, e: { r: cell.mergeEnd.r, c: cell.mergeEnd.c } });
          }
        });
        Object.keys(sh.colW || {}).forEach(function (ci) {
          var i = parseInt(ci, 10);
          if (!isNaN(i)) cols[i] = { wpx: sh.colW[ci] };    // ★列幅も書ける(実測)
        });
        for (var i = 0; i < cols.length; i++) if (!cols[i]) cols[i] = {};
        return { name: sh.name || 'Sheet1', cells: cells, merges: merges, cols: cols };
      })
    };
  }

  /* 書き出す前に「何が落ちるか」を集める。opts.isArrayFormula(sheetIdx, r, c) -> bool */
  function exportWarnings(sheets, opts) {
    opts = opts || {};
    var isArr = opts.isArrayFormula || function () { return false; };
    var styleCells = [], arrayCells = [];
    (sheets || []).forEach(function (sh, si) {
      Object.keys(sh.data || {}).forEach(function (key) {
        var p = key.split(','), r = parseInt(p[0], 10), c = parseInt(p[1], 10);
        if (isNaN(r) || isNaN(c)) return;
        var cell = sh.data[key] || {};
        var where = (sh.name || 'Sheet1') + '!' + addr(r, c);
        if (hasLostStyle(cell)) styleCells.push(where);
        if (typeof cell.f === 'string' && cell.f.charAt(0) === '=' && isArr(si, r, c)) arrayCells.push(where);
      });
    });
    var out = [];
    if (styleCells.length) {
      out.push({
        kind: 'style-lost', cells: styleCells,
        msg: '太字・文字色・背景色・罫線はExcelファイルに入りません'
          + '（表示形式・セルの結合・列幅は入ります）: ' + listCells(styleCells)
      });
    }
    if (arrayCells.length) {
      out.push({
        kind: 'array-spill', cells: arrayCells,
        msg: '配列を返す式が' + arrayCells.length + '件あります。Excelでは1セル分の値になります: ' + listCells(arrayCells)
      });
    }
    return out;
  }

  // 番地は最大8件まで出し、それ以上は「ほか◯件」にする(客が自分で直せる形にしつつ、長すぎない)
  function listCells(list) {
    var head = list.slice(0, 8).join(', ');
    return list.length > 8 ? head + ' ほか' + (list.length - 8) + '件' : head;
  }

  return { gridToBook: gridToBook, exportWarnings: exportWarnings, addr: addr, asValue: asValue, listCells: listCells, dateFmtFor: dateFmtFor, dateSerial: dateSerial };
}));
