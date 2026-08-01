/* xlsx-io.js — 数式入り xlsx の「書き出し / 読み戻し」の唯一の口。
 *
 *  ★置き場所について
 *    今はわざと tests/xlsx-harness/ の下に置いている。lib/ に置くと
 *    scripts/stamp-build.mjs が lib/ 配下の全 .js を内容ハッシュに含めるため、
 *    どのHTMLからも読まれていないのに全HTMLの ?v= が動いてしまう。
 *    誰も使わないコードを配信物に混ぜないための措置。
 *    ★P2でグリッド(book.html)が実際にこれを呼ぶ時に lib/ へ移す。その時が初めての配信。
 *
 *  ★UIには繋いでいない。book.html の saveXlsx() は今もスタブのまま(今回は触らない)。
 *
 *  依存: SheetJS CE 0.20.3(リポジトリ同梱の lib/xlsx.full.min.js)
 *  対応: Node / ブラウザ 両方
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory(require('../../lib/xlsx.full.min.js'));
  } else {
    root.XlsxIO = factory(root.XLSX);
  }
}(typeof self !== 'undefined' ? self : this, function (XLSX) {
  'use strict';

  /* ══ 新しい関数の接頭辞 _xlfn. ═══════════════════════════════════════
   *  ★実測(Excel 365 16.0.20228 / 2026-07-31)で分かった一番大事な事:
   *    XLOOKUP や IFS のような新しい関数を、そのままの名前で xlsx に書くと
   *    **Excelはそのファイルを開けない**(1本混ざっただけでブックごと開けない)。
   *    xlsx の中では _xlfn. を付けた名前で保存する決まりのため。
   *    下の一覧は「_xlfn. を付けたら実Excelが開いて正しく計算した」ことを1関数ずつ確かめた結果。
   *    (SORT/FILTER は _xlfn._xlws. でも通るが、_xlfn. で通ることを実測したのでこちらに統一)
   *  ★LET / LAMBDA は関数名だけでなく引数名にも _xlpm. が要る(_xlfn.LET(_xlpm.x,2,_xlpm.x*3)=6 を実測)。
   *    引数名の付け替えはここではやらない=黙って壊れたファイルを作らないよう、書き出しを止める。
   */
  var XLFN = ['XLOOKUP', 'XMATCH', 'CONCAT', 'TEXTJOIN', 'TEXTBEFORE', 'TEXTAFTER', 'TEXTSPLIT',
    'VALUETOTEXT', 'ARRAYTOTEXT', 'IFS', 'IFNA', 'SWITCH', 'MAXIFS', 'MINIFS',
    'SORT', 'SORTBY', 'UNIQUE', 'FILTER', 'SEQUENCE', 'RANDARRAY',
    'TOCOL', 'TOROW', 'VSTACK', 'HSTACK', 'CHOOSECOLS', 'CHOOSEROWS', 'TAKE', 'DROP', 'EXPAND',
    'NUMBERVALUE', 'ENCODEURL', 'AGGREGATE', 'FORECAST.LINEAR', 'RANK.EQ', 'PERCENTILE.INC',
    'MODE.SNGL', 'BINOM.DIST'];
  var XLFN_SET = {};
  XLFN.forEach(function (n) { XLFN_SET[n] = 1; });
  var NEEDS_XLPM = { LET: 1, LAMBDA: 1 };

  /* ══ 別名(日本語UIの表示名 → ファイルに入る本名) ═════════════════════
   *  ★実測(Excel 365 16.0.20228 / 2026-08-01): 半角→全角の関数の本名は DBCS。
   *    JIS は日本語UIの表示名でしかなく、US-English構文/ファイルの中では通らない
   *    (=JIS(A1) を .Formula で入れると #NAME? になる)。
   *    Excel自身が「表示名=JIS / 保存名=DBCS」で持っているので、書き出す時に本名へ直す。
   *    エンジン側は convertFormula(exally-formula.js) が同じ変換をしている＝入口と出口の両方で本名に寄せる。
   */
  var ALIAS = { JIS: 'DBCS' };

  /* 式の中の関数名だけを見て接頭辞を付ける。
     ・文字列リテラル("...")の中は触らない
     ・すでに _xl… が付いている物は触らない
     ・後読み(?<=)は使わない(古いiOS Safariで正規表現ごと壊れるため) */
  function applyXlfn(formula) {
    var out = '', i = 0, n = formula.length;
    while (i < n) {
      var ch = formula[i];
      if (ch === '"') {                       // 文字列リテラルはそのまま通す
        out += ch; i++;
        while (i < n) { out += formula[i]; if (formula[i] === '"' && formula[i + 1] !== '"') { i++; break; } if (formula[i] === '"') { out += formula[i + 1]; i += 2; } else i++; }
        continue;
      }
      if (/[A-Za-z_]/.test(ch)) {
        var j = i;
        while (j < n && /[A-Za-z0-9_.]/.test(formula[j])) j++;
        var word = formula.slice(i, j);
        var k = j;
        while (k < n && formula[k] === ' ') k++;
        var isCall = formula[k] === '(';
        var upper = word.toUpperCase();
        if (isCall && NEEDS_XLPM[upper] && word.indexOf('_xl') !== 0) {
          throw new Error(upper + ' は引数名に _xlpm. が要るため、この書き出しでは未対応です(壊れたxlsxを作らないために止めました)');
        }
        // ★別名は先に本名へ寄せる(JIS → DBCS)。そのまま書くとExcelが #NAME? にする。
        if (isCall && word.indexOf('_xl') !== 0 && ALIAS[upper]) { out += ALIAS[upper]; i = j; continue; }
        if (isCall && word.indexOf('_xl') !== 0 && XLFN_SET[upper]) out += '_xlfn.' + upper;
        else out += word;
        i = j;
        continue;
      }
      out += ch; i++;
    }
    return out;
  }
  function stripXlfn(formula) {
    return String(formula).replace(/_xlfn\._xlws\./g, '').replace(/_xlfn\./g, '').replace(/_xlpm\./g, '');
  }

  function colName(c) {
    var s = '';
    c = c + 1;
    while (c > 0) { var m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); }
    return s;
  }

  /* book = { sheets: [ { name, cells: { 'A1': {f?:'=SUM(..)', v?:値, t?:'n'|'s'|'b'} } } ] }
   *  ★f は '=' 始まりで受ける(グリッドの持ち方に合わせる)。SheetJSは '=' 無しで持つので剥がす。
   *  ★v は「計算済みの値(キャッシュ値)」。Excelはこれを開いた瞬間に表示し、再計算で上書きする。
   */
  function writeBook(book) {
    var wb = XLSX.utils.book_new();
    (book.sheets || []).forEach(function (sh) {
      var ws = {};
      var maxR = 0, maxC = 0;
      Object.keys(sh.cells || {}).forEach(function (addr) {
        var spec = sh.cells[addr];
        var rc = XLSX.utils.decode_cell(addr);
        if (rc.r > maxR) maxR = rc.r;
        if (rc.c > maxC) maxC = rc.c;
        var cell = {};
        var v = spec.v;
        if (spec.t) cell.t = spec.t;
        else if (typeof v === 'number') cell.t = 'n';
        else if (typeof v === 'boolean') cell.t = 'b';
        else cell.t = 's';
        if (v !== undefined && v !== null) cell.v = v;
        if (spec.f) cell.f = applyXlfn(String(spec.f).replace(/^=/, ''));   // ★新関数に _xlfn. を付ける
        if (cell.v === undefined && !cell.f) return;
        // ★式だけでキャッシュ値が無いセルは v を作らない。
        //   空文字の v を持つ文字列セル + 式 の組み合わせは Excel が開けないファイルになる(実測)。
        //   計算はアプリ側にさせる、という意味でもこちらが正しい。
        if (cell.v === undefined) { cell.t = spec.t || 'n'; delete cell.v; }
        ws[addr] = cell;
      });
      ws['!ref'] = 'A1:' + colName(maxC) + (maxR + 1);
      XLSX.utils.book_append_sheet(wb, ws, sh.name || 'Sheet1');
    });
    return XLSX.write(wb, { bookType: 'xlsx', type: 'buffer', cellFormula: true });
  }

  function readBook(buf) {
    var wb = XLSX.read(buf, { type: 'buffer', cellFormula: true, cellNF: true });
    var sheets = wb.SheetNames.map(function (name) {
      var ws = wb.Sheets[name];
      var cells = {};
      Object.keys(ws).forEach(function (addr) {
        if (addr[0] === '!') return;
        var c = ws[addr];
        var o = { t: c.t };
        if (c.v !== undefined) o.v = c.v;
        if (c.f) o.f = '=' + stripXlfn(c.f);   // ★_xlfn. を外してグリッド側の持ち方('='付き)に戻す
        cells[addr] = o;
      });
      return { name: name, cells: cells };
    });
    return { sheets: sheets };
  }

  return {
    writeBook: writeBook, readBook: readBook, sheetjsVersion: XLSX.version,
    applyXlfn: applyXlfn, stripXlfn: stripXlfn, xlfnNames: XLFN
  };
}));
