/* excel-version.js — 「その式は、相手のExcelで動くか」を判定する。
 *
 *  ★止めない。警告するだけ。客の作業を止めない。
 *  ★どのセルかまで出す（客が自分で直せる形にする）。
 *  ★判定できない物は黙って通す（誤警告を出さない方に倒す）。
 *
 *  警告は2種類ある。★1つ目が本体★
 *   ① そもそも Excel に無い関数（23個）… 版に関係なく必ず #NAME? になる。★365の客でも壊れる★
 *      HyperFormula / Google Sheets 由来の名前。実Excel 16.0.20228 に COM で1つずつ聞いて機械判定した結果。
 *   ② その版にはまだ無い関数        … 選んだ版より新しい関数。公式の版マーカーから。
 *      出典: Microsoft「Excel functions (alphabetical)」の版マーカー
 *            https://support.microsoft.com/en-us/office/excel-functions-alphabetical-b3944572-255d-4efb-bb96-c6d90033e188
 *      「版マーカー 2013 = Excel 2013 以降で使える」という意味なので、選んだ版 < マーカー なら動かない。
 *
 *  ★版マーカーが無い関数は「どの版にもある」として扱う＝警告を出さない。
 *  ★Mac / Online は 365 と同じ扱い（Macは一部関数の提供が遅れることがあるが、そこは割り切り。台帳 version_scope）。
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.ExcelVersion = factory();
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  /* ① どの版のExcelにも無い関数（2026-08-01 実測・実Excel 16.0.20228 に COM で問い合わせ）
        ★版セレクタと無関係に常時警告する。これが本体。 */
  var NOT_IN_EXCEL = ['ARRAYFORMULA', 'ARRAY_CONSTRAIN', 'CHIDISTRT', 'CHIINVRT', 'COUNTUNIQUE',
    'COVARIANCEP', 'COVARIANCES', 'FDISTRT', 'FINVRT', 'INTERVAL', 'ISBINARY', 'LOGNORMINV',
    'MAXPOOL', 'MEDIANPOOL', 'POISSONDIST', 'SKEWP', 'STDEVS', 'TDIST2T', 'TDISTRT', 'TINV2T',
    'VARS', 'VERSION', 'WEIBULLDIST'];

  /* ② 公式の版マーカー（2016以降）。★ここは「Excelが持っている全部」を書いておく。
        うちが提供していない関数が混ざっていても害はなく、将来 HF が増やした時に自動で効く。 */
  var MIN_VER = {
    // 2016
    'FORECAST.ETS': 2016, 'FORECAST.ETS.CONFINT': 2016, 'FORECAST.ETS.SEASONALITY': 2016,
    'FORECAST.ETS.STAT': 2016, 'FORECAST.LINEAR': 2016,
    // 2019
    CONCAT: 2019, IFS: 2019, MAXIFS: 2019, MINIFS: 2019, SWITCH: 2019, TEXTJOIN: 2019,
    // 2021
    ARRAYTOTEXT: 2021, FILTER: 2021, LET: 2021, RANDARRAY: 2021, SEQUENCE: 2021, SORT: 2021,
    SORTBY: 2021, UNIQUE: 2021, VALUETOTEXT: 2021, XLOOKUP: 2021, XMATCH: 2021,
    // 2024
    BYCOL: 2024, BYROW: 2024, CHOOSECOLS: 2024, CHOOSEROWS: 2024, DROP: 2024, EXPAND: 2024,
    HSTACK: 2024, IMAGE: 2024, ISOMITTED: 2024, LAMBDA: 2024, MAKEARRAY: 2024, MAP: 2024,
    REDUCE: 2024, SCAN: 2024, TAKE: 2024, TEXTSPLIT: 2024, TEXTAFTER: 2024, TEXTBEFORE: 2024,
    TOCOL: 2024, TOROW: 2024, VSTACK: 2024, WRAPCOLS: 2024, WRAPROWS: 2024
  };

  // 版セレクタ(localStorage の exally_excel_version)の値 → 比較用の数
  var VER_NUM = {
    excel_2016: 2016, excel_2019: 2019, excel_2021: 2021, excel_2024: 2024,
    excel_365: 9999, excel_online: 9999, excel_mac: 9999,
    excel_none: 9999   // 「Excelなし」＝版の警告は出さない。①(Excelに無い関数)だけ出る
  };
  var VER_LBL = {
    excel_2016: 'Excel 2016', excel_2019: 'Excel 2019', excel_2021: 'Excel 2021',
    excel_2024: 'Excel 2024', excel_365: 'Excel 365', excel_online: 'Excel Online',
    excel_mac: 'Excel（Mac）', excel_none: 'Excelなし'
  };

  var NOT_SET = {};
  NOT_IN_EXCEL.forEach(function (n) { NOT_SET[n] = 1; });

  /* 式の中の関数名を拾う。★文字列リテラルの中は見ない（"XLOOKUP(" を関数と数えない）。
     ★後読み(?<=)は使わない＝旧iOS Safari で正規表現ごと壊れるため。
     xlsx-io.js の applyXlfn と同じ歩き方にしてある。 */
  function functionsIn(formula) {
    var out = [], i = 0, s = String(formula || ''), n = s.length;
    while (i < n) {
      var ch = s.charAt(i);
      if (ch === '"') {
        i++;
        while (i < n) { if (s.charAt(i) === '"' && s.charAt(i + 1) !== '"') { i++; break; } i += (s.charAt(i) === '"') ? 2 : 1; }
        continue;
      }
      if (/[A-Za-z_]/.test(ch)) {
        var j = i;
        while (j < n && /[A-Za-z0-9_.]/.test(s.charAt(j))) j++;
        var word = s.slice(i, j), k = j;
        while (k < n && s.charAt(k) === ' ') k++;
        if (s.charAt(k) === '(') out.push(word.toUpperCase());
        i = j;
        continue;
      }
      i++;
    }
    return out;
  }

  /* 1つの式を見て、警告を返す（空配列＝問題なし）。verKey は版セレクタの値。 */
  function checkFormula(formula, verKey) {
    var want = VER_NUM[verKey];
    if (want === undefined) want = 9999;              // 知らない値なら版の警告は出さない
    var seen = {}, out = [];
    functionsIn(formula).forEach(function (name) {
      if (seen[name]) return;
      seen[name] = 1;
      if (NOT_SET[name]) {
        out.push({
          name: name, kind: 'not-in-excel',
          msg: name + ' は Excel には無い関数です（この画面では計算できますが、Excelで開くと #NAME? になります）'
        });
        return;
      }
      var min = MIN_VER[name];
      if (min && want < min) {
        out.push({
          name: name, kind: 'too-new', min: min,
          msg: name + ' は ' + min + ' 以降の関数です（' + (VER_LBL[verKey] || verKey) + ' では #NAME? になります）'
        });
      }
    });
    return out;
  }

  function colLetter(c) { var s = ''; c += 1; while (c > 0) { s = String.fromCharCode(64 + ((c - 1) % 26 + 1)) + s; c = Math.floor((c - 1) / 26); } return s; }
  function addr(r, c) { return colLetter(c) + (r + 1); }

  /* ブック全体を見て、セル番地つきでまとめる。書き出し前の警告に使う。 */
  function checkBook(sheets, verKey) {
    var byMsg = {};
    (sheets || []).forEach(function (sh) {
      Object.keys(sh.data || {}).forEach(function (key) {
        var p = key.split(','), r = parseInt(p[0], 10), c = parseInt(p[1], 10);
        if (isNaN(r) || isNaN(c)) return;
        var cell = sh.data[key] || {};
        if (typeof cell.f !== 'string' || cell.f.charAt(0) !== '=') return;
        checkFormula(cell.f, verKey).forEach(function (w) {
          var k = w.kind + '|' + w.name;
          if (!byMsg[k]) byMsg[k] = { kind: w.kind, name: w.name, msg: w.msg, cells: [] };
          byMsg[k].cells.push((sh.name || 'Sheet1') + '!' + addr(r, c));
        });
      });
    });
    // ★Excelに無い関数(本体)を先に出す
    return Object.keys(byMsg).map(function (k) { return byMsg[k]; })
      .sort(function (a, b) { return (a.kind === 'not-in-excel' ? 0 : 1) - (b.kind === 'not-in-excel' ? 0 : 1); });
  }

  return {
    checkFormula: checkFormula, checkBook: checkBook, functionsIn: functionsIn, addr: addr,
    NOT_IN_EXCEL: NOT_IN_EXCEL, MIN_VER: MIN_VER, VER_NUM: VER_NUM, VER_LBL: VER_LBL
  };
}));
