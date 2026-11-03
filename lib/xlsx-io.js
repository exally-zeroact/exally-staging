/* xlsx-io.js — 数式入り xlsx の「書き出し / 読み戻し」の唯一の口。
 *
 *  ★2026-08-01 に tests/xlsx-harness/ から lib/ へ移した＝ここが初めての配信。
 *    グリッド(book.html)の「Excelに書き出す」がこれを呼ぶようになったため。
 *    ハーネス(tests/xlsx-harness/)も同じこの1本を使う＝書き出しの口は1つだけ。
 *    ★書き出しの正しさ(実Excelで開けるか)は tools/verify-workbook-excel.ps1 で確認している。
 *
 *  ここが持っている「実Excelに聞いて分かった事」:
 *    ・新しい関数は xlsx の中では _xlfn. を付けた名前で保存する(付け忘れるとExcelが式を壊す/開けない)
 *    ・日本語UIの表示名は本名に直す(JIS→DBCS / YEN→DOLLAR)。表示名のまま書くと #NAME? になる
 *    ・LET / LAMBDA は引数名にも _xlpm. が要るので、黙って壊れたファイルを作らず書き出しを止める
 *
 *  依存: SheetJS CE 0.20.3(リポジトリ同梱の lib/xlsx.full.min.js)
 *  対応: Node / ブラウザ 両方
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory(require('./xlsx.full.min.js'));
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
    /* ★★WRAPROWS / WRAPCOLS —— 上の 一族と 同じ なのに ★この 2個だけ 抜けて いました★★
         （2026-09-20 実Excel 16.0 build 20326 で 実測
          紙 `docs/measured/golden-wrap-excel-2026-09-20.tsv`）
       ★何が 起きて いたか★
         お客さんは 打てる（台にも 皮にも 在る）／画面で 計算できる／保存も 出来る
         ⇒ でも 書き出しは ★裸★（`WRAPROWS(A1:A6,3)`）
         ⇒★★相手の Excel で #NAME?（-2146826259）★★
         ⇒★溢れ先も 全部 空＝★表が 丸ごと 消えます★★
           （うちの 画面では 2行3列の 表／実Excel では 誤り 1マスだけ）
       ★対照★（同じ 紙で 同時に 測りました）
         `_xlfn.SEQUENCE(3)` ＝ 1/2/3 に 溢れる（★付くのが 正しい★）
         `TRANSPOSE(A1:A3)`  ＝ 1/2/3 に 溢れる（★裸で 正しい★）
         ⇒★「裸だから 割れた」では 無く ★名前ごとに 要る／要らない★★
       ★なぜ 09-16 の 測りで 漏れたか★
         分母の 紙 `docs/measured/xlfn-namae.txt`（396個）に
         ★`WRAPROWS` `WRAPCOLS` が 入って いません★（台は 最初から 持って いた）
         ⇒★分母から 45個 漏れて いました★（残りは まだ 測って いません） */
    'WRAPROWS', 'WRAPCOLS',
    /* ★★FORMULATEXT —— ★台が 持つ のに 名簿に 無かった 4個目★★
         （2026-09-20 実Excel 16.0 build 20326 に 打たせて 実測・59）
       実Excel が 書く 字 ＝ `_xlfn.FORMULATEXT(B1)`
       ★但し `cm` は 無し＝★溢れません★
         ⇒割れても ★その マスだけ #NAME?★（表は 消えません）
       ★★2026-09-21 直し ── ★この「溢れません」は 言い過ぎでした★★
         ＝59 で 打ったのは `=FORMULATEXT(B1)`＝★1マスを 渡した 形★だけ
         ⇒★範囲を 渡すと 溢れます★（★75 で 実測★）
             `=FORMULATEXT(A1:A3)` ⇒ ★3マス★ `#N/A` / `=1+1` / `#N/A`
             （A2 だけが 式＝★式で ない マスは #N/A★）
         ⇒★`cm` が 付くかは ★渡した 形★で 変わる＝1つの 形で 決めては いけない★
         ⇒★古い 字は 消して いません★（外れた 見込みは 捨てない）
         ⇒★WRAPROWS / WRAPCOLS / MODE.MULT より 軽い★
       ★見つけ方★ ＝ ★門が 自分で 赤に しました（2回目）★
         （紙が 55個 → 62個 に 増えた 瞬間）
       ★点入りか どうかは 見分けに なりません★
         `FORMULATEXT` `XLOOKUP` `XOR` `SHEET` は 点が 無い のに `_xlfn.` が 要る */
    /* ★★ISFORMULA —— ★`FORMULATEXT` と 同じ 2013年の 兄弟★★
         （2026-09-20 実Excel 16.0 build 20326 に 打たせて 実測・61）
       実Excel が 書く 字 ＝ `_xlfn.ISFORMULA(A1:A3)`  cm 在り  ref="H1337:H1339"
       ★`cm` が 在る＝溢れます★ ⇒★割れると 表が 丸ごと 消えます★
         （`FORMULATEXT` は ★1マスを 渡した 時は★ 溢れない
           ⇒★範囲を 渡すと 溢れます★＝2026-09-21 に 直しました）
       ★見つけ方★ ＝ ★台の 376個を 全部 実Excel に 聞いた★（総当たり）
         ★兄弟を 見なかった★ ⇒ `FORMULATEXT` を 見つけた 時に
         `ISFORMULA` も 見て いれば 1回で 済んで いました
         ⇒★★総当たりに して いなければ 見つかって いません★★ */
    'ISFORMULA',
    'FORMULATEXT',
    'NUMBERVALUE', 'ENCODEURL', 'AGGREGATE', 'FORECAST.LINEAR', 'RANK.EQ', 'RANK.AVG', 'PERCENTILE.INC',
    'MODE.SNGL', 'BINOM.DIST', 'PERMUTATIONA',
    /* ★★MODE.MULT —— ★`MODE.SNGL` は 在る のに こちらだけ 抜けて いました★★
         （2026-09-20 実Excel 16.0 build 20326 に 打たせて 実測
          紙 `docs/measured/golden-jitsu-excel-no-shirushi-2026-09-20.tsv` H43）
       実Excel が 書く 字 ＝ `_xlfn.MODE.MULT(A1:A6)`（ref="H43:H44"・cm 在り）
       ★溢れる 式★なので 裸だと #NAME? で ★表が 丸ごと 消えます★
       ★同じ 回に 聞いた 残り 20個は 全部 裸で 正しい★
         ASC AVERAGEIFS BAHTTEXT DBCS DOLLAR FIXED FORECAST INDIRECT KURT
         LEFTB LENB LOOKUP MDETERM MIDB OFFSET PERCENTRANK PERMUT PHONETIC
         RIGHTB TRIMMEAN
       ⇒★「古い 物は 要らない」が 成り立ち、★点入りの 新しい 形だけ 要る★★
       ⇒★これを 見つけたのは ★門その もの★です
         （紙が 増えた 瞬間 `xlfn-morenashi` が 自分で 赤に なりました） */
    'MODE.MULT',
    /* ★★2026-09-16 に ★実Excel に 全部 聞いて★ 足した 91個★★
       ★司さん（2026-09-16）「Exally と Excel に 引き渡しても ちゃんと 動くか
         ★実際に 動作確認しながら★ やれよ」
       ★測った 道★ `docs/measured/toru-xlfn-zenbu.ps1`（台が 知る 394個を 実Excel に 打たせた）
       ★分かった 事★ … 実Excel は ★104個★に `_xlfn.` を 要求するが
                      うちの 一覧は ★ 39個★しか 持って いなかった
                      ＝★★ 91個が 書き出すと #NAME? に なって いた★★
       ★紙が 全部 緑でも 見つかりません★（台の 中では 正しく 計算できる）
       ★残りの 27個は 参照が 要るので `(1)` では 打てず 未測定★ */
    'ACOT', 'ACOTH', 'ARABIC', 'BASE', 'BETA.DIST', 'BETA.INV', 'BINOM.DIST.RANGE', 'BINOM.INV', 'BITAND', 'BITLSHIFT', 'BITOR', 'BITRSHIFT', 'BITXOR', 'CEILING.MATH', 'CEILING.PRECISE', 'CHISQ.DIST', 'CHISQ.DIST.RT', 'CHISQ.INV', 'CHISQ.INV.RT', 'CHISQ.TEST', 'COMBINA', 'CONFIDENCE.NORM', 'CONFIDENCE.T', 'COT', 'COTH', 'COVARIANCE.P', 'COVARIANCE.S', 'CSC', 'CSCH', 'DAYS', 'DECIMAL', 'ERF.PRECISE', 'ERFC.PRECISE', 'EXPON.DIST', 'F.DIST', 'F.DIST.RT', 'F.INV', 'F.INV.RT', 'F.TEST', 'FLOOR.MATH', 'FLOOR.PRECISE', 'GAMMA', 'GAMMA.DIST', 'GAMMA.INV', 'GAMMALN.PRECISE', 'GAUSS', 'HYPGEOM.DIST', 'IMCOSH', 'IMCOT', 'IMCSC', 'IMCSCH', 'IMSEC', 'IMSECH', 'IMSINH', 'IMTAN', 'ISOMITTED', 'ISOWEEKNUM', 'LOGNORM.DIST', 'LOGNORM.INV', 'NEGBINOM.DIST', 'NORM.DIST', 'NORM.INV', 'NORM.S.DIST', 'NORM.S.INV', 'PDURATION', 'PERCENTILE.EXC', 'PHI', 'POISSON.DIST', 'QUARTILE.EXC', 'QUARTILE.INC', 'RRI', 'SEC', 'SECH', 'SHEET', 'SHEETS', 'SKEW.P', 'STDEV.P', 'STDEV.S', 'T.DIST', 'T.DIST.2T', 'T.DIST.RT', 'T.INV', 'T.INV.2T', 'T.TEST', 'UNICHAR', 'UNICODE', 'VAR.P', 'VAR.S', 'WEIBULL.DIST', 'XOR', 'Z.TEST'];
  //  ★PERMUTATIONA は 2026-08-02(第3波P3)に追加。実Excelで開いたら その式だけ #NAME? になった
  //    (同じP3の FORECAST/GESTEP/INTERCEPT/IRR/KURT/MDETERM/MODE/PERCENTRANK/PERMUT/TRIMMEAN は接頭辞不要だった)。
  //  ★RANK.AVG は 2026-08-01(第3波P2)に追加。RANK.EQ は入っていたのに RANK.AVG だけ抜けており、
  //    書き出したブックを実Excelで開くと その式だけ #NAME? になっていた
  //    (tools/verify-workbook-excel.ps1 が拾った＝画面のテストだけでは絶対に見つからない類)。
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
  var ALIAS = { JIS: 'DBCS', YEN: 'DOLLAR' };

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
  /* ★.xlsb は「裸の _xlws.」を 出してくる（2026-08-29 実測）★
   *   司さんの実物（.xlsb）を SheetJS で 読むと `_xlws.FILTER(...)` が そのまま 来る。
   *   ここは `_xlfn._xlws.` しか 外していなかったので すり抜け、★52本が #ERROR!★ に なっていた
   *   （実Excel の答えは 177,523 / 668,639 などの ★金額★）。
   *   ⇒ ★長い方から 順に 外す★（先に短い方を外すと `_xlfn.` だけ 消えて `_xlws.` が 残る）。
   *   ★同じ一覧を exally-formula.js の convertFormula も 持つ★＝あちらは エンジン直前の 最後の関所。
   *     tests/xlfn-strip.test.mjs が ★両方が 同じ物を 外すか★ を 突き合わせる。 */
  function stripXlfn(formula) {
    return String(formula)
      .replace(/_xlfn\._xlws\./g, '')
      .replace(/_xlfn\./g, '')
      .replace(/_xlws\./g, '')
      .replace(/_xludf\./g, '')
      .replace(/_xlpm\./g, '');
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
        /* ★★溢れ(スピル)を 運ぶ★★（2026-09-20）
             `F` ＝ 溢れの 範囲（'A1:A3'）⇒ SheetJS が `<f t="array" ref="A1:A3">` を 書く
             `D` ＝ 動く並びの 印   ⇒ SheetJS が `<c cm="1">` を 書く
             ★これが 無いと★ 頭が ★普通の 1マスの 式★に なり、隣に 値が 在る ので
             実Excel で 開いた 瞬間 ★#SPILL!★ です（2026-09-20 生の xml で 実測）。 */
        if (spec.F) cell.F = spec.F;
        if (spec.D) cell.D = true;
        if (cell.v === undefined && !cell.f) return;
        // ★式だけでキャッシュ値が無いセルは v を作らない。
        //   空文字の v を持つ文字列セル + 式 の組み合わせは Excel が開けないファイルになる(実測)。
        //   計算はアプリ側にさせる、という意味でもこちらが正しい。
        if (cell.v === undefined) { cell.t = spec.t || 'n'; delete cell.v; }
        // ★表示形式(z)。実Excelで開いて確かめたら、ここで運ばないと「G/標準」になって
        //   グリッドで「54,000」に見えていた物が「54000」で落ちる(2026-08-02 実測)。
        if (spec.z) cell.z = spec.z;
        ws[addr] = cell;
      });
      ws['!ref'] = 'A1:' + colName(maxC) + (maxR + 1);
      // ★セルの結合と列幅は SheetJS CE でも書ける(実測 2026-08-01)。太字/色/罫線は書けない。
      if (sh.merges && sh.merges.length) ws['!merges'] = sh.merges;
      if (sh.cols && sh.cols.length) ws['!cols'] = sh.cols;
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
        if (c.z) o.z = c.z;                    // 表示形式(読む側でも落とさない)
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
