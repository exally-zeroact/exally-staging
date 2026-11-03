/* shiki-tsunagi.js — ★自前で 書いた 関数を「自前の エンジン」に 繋ぐ 皮★（2026-09-14）
 *
 *  ★置き場★ … `lib/shiki-tsunagi.js`（土台 `lib/shiki-*.js` の 仲間）
 *    ・2026-09-14 に repo の 外（scratchpad）から ここへ 移しました。
 *    ・★外に 置いて いた 訳★＝前に 中途半端な 物を `git add -A` で 混ぜて 叱られた為。
 *      ⇒ 今回は ★61個が 実Excel と 突き合わせ済（yosoku 721/724・nokori 83/83 ほか）★
 *        に なってから 入れて います。
 *    ・★見張り★ … `node scripts/stamp-build.mjs` を 掛けてから commit する事
 *      （`book.html` から 読み込む 時は ★?v= が 要る★）
 *
 *  ★★これが「借り物を 外す」の 本丸です★★
 *    司さん 2026-09-13「借り物が あったら 商用など いろいろ 引っかかる。★自作で やれって 前から いよろが★」
 *    司さん 2026-09-14「★まず 借り物を なくすのが お前の 仕事や★」
 *
 *  ★★分かった 事（2026-09-14・経営者1 の 検算 → 私が 別の手で 裏を 取った）★★
 *    ・自前で 書いた 関数は ★65個★（`lib/formula-*-plug.js`）＝私は「7個」と 言って いた（10倍 少ない）
 *    ・★中身（計算）は もう 借り物から 独立して います★＝
 *      require を 差し替えて ★hyperformula を 読もうと したら 例外★ に した 状態で
 *      formula-extra / nokori / kane / yosoku / complex を 読み込み ⇒ ★83個の 部品が 動いた★
 *      ★借り物を 読もうと した 回数 = 0★
 *    ⇒★借り物に ぶら下がって いるのは 「皮」（-plug.js）だけ★
 *    ⇒★やる事は「65個を 書き直す」では なく「皮を 差し替える」★
 *
 *  ★★この 台の 仕事★★
 *    `lib/formula-*.js`（中身）を ★自前の エンジン（shiki-hyou / shiki-kansuu）★ から 呼べる形に する。
 *    ★中身は 1行も 触りません★＝[[feedback_mihon_no_michi_ga_futatsu_aru_toki_katahou_dake_naosu_na]]
 *
 *  ★★値の 形の 行き来★★
 *    自前の エンジン … { 型:'数'|'字'|'真偽'|'誤'|'空', 値 }  ／ 四角は { 溢れ, 行数, 列数, 並び }
 *    中身の 部品     … ★素の 値★（数は number・字は string・表は [[…],…]）
 *    ⇒ ここで ★ほどく／包む★ だけ します。
 *
 *  ★★見て いない 範囲（★書かない 皮は「全部 繋いだ」と 読まれる★）★★
 *    ・今 繋いで いるのは ★formula-extra の 13個だけ★（★残り 52個は まだ★）
 *    ・借り物側の `-plug.js` は ★そのまま★＝★対応表が 2本 在る 状態★です。
 *      ⇒★片方だけ 増えると 食い違う★ので ★見張りで 突き合わせます★
 *        （tests/shiki-tsunagi.test.mjs が ★両方に 同じ 関数名が 載っているか★ を 見る）
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(
    typeof require === 'function' ? require('./shiki-keisan.js') : root.ShikiKeisan,
    typeof require === 'function' ? require('./shiki-afure.js') : root.ShikiAfure,
    typeof require === 'function' ? require('./formula-extra.js') : root.FormulaExtra,
    typeof require === 'function' ? require('./bahttext.js') : root.Bahttext,
    typeof require === 'function' ? require('./formula-kane.js') : root.FormulaKane,
    typeof require === 'function' ? require('./formula-filterxml.js') : root.FormulaFilterxml,
    typeof require === 'function' ? require('./formula-cell.js') : root.FormulaCell,
    typeof require === 'function' ? require('./formula-yosoku.js') : root.FormulaYosoku,
    typeof require === 'function' ? require('./formula-nokori.js') : root.FormulaNokori,
    typeof require === 'function' ? require('./bessel.js') : root.Bessel);
  /* ★★窓の 名前は ★本物と 1字も 違わない★ 事★★（2026-09-18・★繋ぐ 前に 見つけた★）
       `lib/bahttext.js` .......... `root.BahtText`（★大文字 T★）
       `lib/formula-filterxml.js` . `root.FormulaFilterXml`（★大文字 X★）
     ★ここは 小文字で 書いて いました★＝★ブラウザでは `undefined`★
     ⇒★node の `require` では 通るので 今まで 出ませんでした★
     ⇒★両方 受けます★（★本物を 先に／古い 綴りも 残す★） */
  else root.ShikiTsunagi = factory(root.ShikiKeisan, root.ShikiAfure, root.FormulaExtra,
    root.BahtText || root.Bahttext, root.FormulaKane,
    root.FormulaFilterXml || root.FormulaFilterxml,
    root.FormulaCell, root.FormulaYosoku, root.FormulaNokori, root.Bessel);
})(typeof self !== 'undefined' ? self : this, function (計, 溢, 中, バーツ, 金, XML, セル, 予, 残, ベ) {
  'use strict';

  /* ══ ★自前の 値 → 素の 値★（中身の 部品に 渡す 形） ══ */
  function ほどく(v) {
    if (溢.溢れか(v)) return v.並び.map(function (段) { return 段.map(ほどく); });
    if (!v || !v.型) return v;
    if (v.型 === '空') return null;
    if (v.型 === '誤') return { 誤: v.値 };        /* ★誤りは 印を 付けて 渡す★ */
    return v.値;
  }

  /* ══ ★中身が 返す 誤りの 名前 → 実Excel の 書き方★ ══
     ★中身の 部品は `{誤り:'DIV_BY_ZERO'}` の 形で 返します★（2026-09-14 実測）。
     借り物の 皮も 同じ物を `ErrorType[v.誤り]` で 直して いました。
     ⇒★私は `{誤:…}` を 見て いて 拾えず、全部 #VALUE! に して いました★。 */
  var 誤りの名 = {
    DIV_BY_ZERO: '#DIV/0!', NA: '#N/A', VALUE: '#VALUE!', REF: '#REF!',
    NAME: '#NAME?', NUM: '#NUM!', NULL: '#NULL!', CYCLE: '#CYCLE!', ERROR: '#ERROR!',
    SPILL: '#SPILL!', CALC: '#CALC!'
  };
  function 誤りか(x) {
    return !!(x && typeof x === 'object' && !Array.isArray(x) && (x.誤り || x.誤));
  }
  function 誤りを直す(x) {
    var n = x.誤り || x.誤;
    return 計.誤(誤りの名[n] || (String(n).charAt(0) === '#' ? String(n) : '#VALUE!'));
  }

  /* ══ ★素の 値 → 自前の 値★（エンジンへ 返す 形） ══ */
  function 包む(x) {
    if (x === null || x === undefined) return 計.空;
    if (誤りか(x)) return 誤りを直す(x);        /* ★中身の 誤りを 先に 拾う★ */
    if (Array.isArray(x)) {
      var 並び = x.map(function (段) {
        return (Array.isArray(段) ? 段 : [段]).map(包む);
      });
      return 溢.溢れ(並び);
    }
    if (typeof x === 'number') return 計.数(x);
    if (typeof x === 'boolean') return 計.真偽(x);
    if (typeof x === 'string') {
      /* ★中身の 部品は 誤りを 字で 返す 事が 在る★（'#DIV/0!' など） */
      if (/^#[A-Z0-9/!?]+$/.test(x)) return 計.誤(x);
      return 計.字(x);
    }
    return 計.誤('#VALUE!');
  }

  /* ══ ★引数の 受け取り方★ ══
     ★四角と 直に 書いた 物を 分けない★＝中身の 部品は 表を 受ける 作りなので
     ★そのまま 素に して 渡す★（[[feedback_dasu_mae_ni_fumu_mon]] の 門2＝見ていない範囲を 書く）
     ⇒★ここで 型の 決まり（四角の 中の 字は 無視 等）は ★掛けて いません★★
       掛けるのは `shiki-kansuu.js` の 7個だけ。★混ぜない★ */
  function 素にする(引数たち) {
    return 引数たち.map(function (a) {
      /* ★四角は 形（行数×列数）を 保つ★
         ＝2026-09-14 に ここで ★必ず 縦一列に 潰して いました★。
         `ARRAYTOTEXT(A1:B3,1)` が 借り物 `{1,2;3,4;5,6}` に 対して
         ★`{1;2;3;4;5;6}` と 出た★のが それです。 */
      if (a.溢れ === true) return a.並び.map(function (段) { return 段.map(ほどく); });
      if (a.種 === '四角') {
        if (a.並び && a.行数 && a.列数) {
          /* ★2つの 形が 来ます★
             ①★段ごと（2次元）★ … 突き合わせの 道具や 溢れから 来る 物
             ②★平ら ＋ 行数×列数★ … ★土台（lib/shiki-hyou.js）の 四角★
                ＝土台の 並びは ★わざと 平ら★（`shiki-kansuu.js` の `ほどく` が
                  平らに 舐める 作りなので 2次元に すると 土台の 7個が 壊れる）
                ⇒★組み直すのは 皮の 仕事★。2026-09-14 に ここで
                  `=ARRAYTOTEXT(A1:B3,1)` と `=MINVERSE(A1:B2)` が ★#VALUE!★に なって いた。 */
          if (Array.isArray(a.並び[0])) {
            return a.並び.map(function (段) { return 段.map(ほどく); });
          }
          var 出 = [];
          for (var r = 0; r < a.行数; r++) {
            var 段 = [];
            for (var c = 0; c < a.列数; c++) 段.push(ほどく(a.並び[r * a.列数 + c]));
            出.push(段);
          }
          return 出;
        }
        return a.並び.map(function (v) { return [ほどく(v)]; });
      }
      return ほどく(a.値);
    });
  }

  /* ══ ★対応表★ ══
     ★関数名 → 中身の 部品 ＋ 引数の 渡し方★
     ★答えは 実Excel を 打って 確かめ済み★（`lib/formula-extra-plug.js` の 頭に 実測が 在る） */
  var 表 = {
    /* ★組（条件の 対）を 作って 渡す★＝AVERAGEIFS(平均する, 範囲1, 条件1, …)
       ★条件は 表に しない★（借り物の 皮も `evaluateAst` の 生の 値を 渡して いる） */
    'AVERAGEIFS': function (素) {
      var 組 = [];
      for (var i = 1; i + 1 < 素.length; i += 2) 組.push([表にする(素[i]), 素[i + 1]]);
      return 中.条件つき平均(表にする(素[0]), 組);
    },
    /* ★TAKE と DROP は 同じ 部品＝最後の「落とすか」で 分ける★ */
    'TAKE': function (素) { return 中.切り出す(表にする(素[0]), 数に(素[1]), 数に(素[2]), false); },
    'DROP': function (素) { return 中.切り出す(表にする(素[0]), 数に(素[1]), 数に(素[2]), true); },
    /* ★CHOOSECOLS と CHOOSEROWS も 同じ 部品＝「列か」で 分ける★
       ★2つ目 以降を 全部 数に ほどいて 並べる★ */
    'CHOOSECOLS': function (素) { return 中.選び出す(表にする(素[0]), 番号たち(素), true); },
    'CHOOSEROWS': function (素) { return 中.選び出す(表にする(素[0]), 番号たち(素), false); },
    /* ★TOCOL と TOROW も 同じ 部品＝「縦か」で 分ける★ */
    'TOCOL': function (素) { return 中.一本にする(表にする(素[0]), 数に(素[1]) || 0, !!素[2], true); },
    'TOROW': function (素) { return 中.一本にする(表にする(素[0]), 数に(素[1]) || 0, !!素[2], false); },
    /* ★WRAPROWS と WRAPCOLS も 同じ 部品＝「行でか」で 分ける★ */
    'WRAPROWS': function (素) { return 中.折り返す(表にする(素[0]), 数に(素[1]), 素[2], true); },
    'WRAPCOLS': function (素) { return 中.折り返す(表にする(素[0]), 数に(素[1]), 素[2], false); },
    'EXPAND': function (素) { return 中.広げる(表にする(素[0]), 数に(素[1]), 数に(素[2]), 素[3]); },
    'ARRAYTOTEXT': function (素) { return 中.表を字に(表にする(素[0]), 数に(素[1])); },
    'MODE.MULT': function (素) { return 中.最頻値たち(表にする(素[0])); },
    /* ★実Excel の 決まり（台帳 tests/fixtures/bahttext-golden.json の「変な物」）★
         空の字 → ★#VALUE!★ ／ 空マス → 0 ／ "123" → 数として 読む ／ TRUE=1・FALSE=0 */
    'BAHTTEXT': function (素) {
      if (!バーツ) return null;
      var x = 素[0];
      if (x === null || x === undefined) return バーツ.字にする(0);   /* ★空マスは 0★ */
      if (typeof x === 'boolean') return バーツ.字にする(x ? 1 : 0);
      if (typeof x === 'string') {
        if (x.trim() === '') return { 誤り: 'VALUE' };                 /* ★空の字は #VALUE!★ */
        var n = Number(x);
        if (!isFinite(n)) return { 誤り: 'VALUE' };                    /* ★数に ならない 字★ */
        return バーツ.字にする(n);
      }
      var m = 数に(x);
      if (m === null) return { 誤り: 'VALUE' };
      return バーツ.字にする(m);
    }
  };

  /* ══ ★お金の 22個★ ══
     ★呼び方は `lib/formula-kane-plug.js` の 112〜129行 を そのまま 写しました★（当て推量 0）。
     ★答えは `docs/measured/kansuu46/golden-kane-2026-09-07.tsv`★
       ＝実Excel 16.0 build 20326 に 打たせた ★700本（うち この22個ぶん 604本）★。
     ★お金の 計算は 一番 間違えては いけない★ので ★借り物と 同じ★では 足りず、
     ★実Excel の 実測と 直に★ 突き合わせます。 */
  var 日 = function (n) { return 金.数から日(n); };
  var 金の表 = [
    ['ACCRINT',   8, function (a) { return 金.経過利息(a[0], a[1], a[2], a[3], a[4], a[5], a[6], a[7] === null ? true : a[7]); }],
    ['ACCRINTM',  5, function (a) { return 金.満期一括の経過利息(a[0], a[1], a[2], a[3], a[4]); }],
    ['AMORDEGRC', 7, function (a) { return 金.仏定率(a[0], a[1], a[2], a[3], a[4], a[5], a[6]); }],
    ['AMORLINC',  7, function (a) { return 金.仏定額(a[0], a[1], a[2], a[3], a[4], a[5], a[6]); }],
    ['COUPDAYBS', 4, function (a) { return 金.前からの日数(日(a[0]), 日(a[1]), a[2], a[3] === null ? 0 : a[3]); }],
    ['COUPDAYS',  4, function (a) { return 金.期間の日数(日(a[0]), 日(a[1]), a[2], a[3] === null ? 0 : a[3]); }],
    ['COUPDAYSNC',4, function (a) { return 金.次までの日数(日(a[0]), 日(a[1]), a[2], a[3] === null ? 0 : a[3]); }],
    ['COUPNCD',   4, function (a) { var p = 金.次の利払日(日(a[0]), 日(a[1]), a[2]); return 金.日から数(p.y, p.m, p.d); }],
    ['COUPNUM',   4, function (a) { return 金.利払回数(日(a[0]), 日(a[1]), a[2]); }],
    ['COUPPCD',   4, function (a) { var p = 金.前の利払日(日(a[0]), 日(a[1]), a[2]); return 金.日から数(p.y, p.m, p.d); }],
    ['DISC',      5, function (a) { return 金.割引率(a[0], a[1], a[2], a[3], a[4]); }],
    ['DURATION',  6, function (a) { return 金.期間(a[0], a[1], a[2], a[3], a[4], a[5]); }],
    ['INTRATE',   5, function (a) { return 金.利率(a[0], a[1], a[2], a[3], a[4]); }],
    ['MDURATION', 6, function (a) { return 金.修正期間(a[0], a[1], a[2], a[3], a[4], a[5]); }],
    ['PRICE',     7, function (a) { return 金.価格(a[0], a[1], a[2], a[3], a[4], a[5], a[6]); }],
    ['PRICEDISC', 5, function (a) { return 金.割引債の価格(a[0], a[1], a[2], a[3], a[4]); }],
    ['PRICEMAT',  6, function (a) { return 金.満期一括の価格(a[0], a[1], a[2], a[3], a[4], a[5]); }],
    ['RECEIVED',  5, function (a) { return 金.受取額(a[0], a[1], a[2], a[3], a[4]); }],
    ['VDB',       7, function (a) { return 金.可変定率(a[0], a[1], a[2], a[3], a[4], a[5] === null ? 2 : a[5], a[6] === null ? false : a[6]); }],
    ['YIELD',     7, function (a) { return 金.利回り(a[0], a[1], a[2], a[3], a[4], a[5], a[6]); }],
    ['YIELDDISC', 5, function (a) { return 金.割引債の利回り(a[0], a[1], a[2], a[3], a[4]); }],
    ['YIELDMAT',  6, function (a) { return 金.満期一括の利回り(a[0], a[1], a[2], a[3], a[4], a[5]); }],
    /* ★★2026-09-19 ... ODDL の 2個★★（★台に 無い 27 → 25★）
         ＝`lib/formula-kane.js` の `足した名前()` が 正本／ここは 繋ぐ 所だけ
         ＝★basis を 省くと 0★（実Excel と 同じ） */
    ['ODDLPRICE', 8, function (a) { return 金.最終端数の価格(a[0], a[1], a[2], a[3], a[4], a[5], a[6], a[7] === null ? 0 : a[7]); }],
    ['ODDLYIELD', 8, function (a) { return 金.最終端数の利回り(a[0], a[1], a[2], a[3], a[4], a[5], a[6], a[7] === null ? 0 : a[7]); }],
    /* ★★2026-09-19 ... ODDF の 2個★★（★台に 無い 27 → 25★・★保留 0個★）
         ＝`lib/formula-kane.js` の `足した名前()` が 正本／ここは 繋ぐ 所だけ
         ＝★basis を 省くと 0★（実Excel と 同じ） */
    ['ODDFPRICE', 9, function (a) { return 金.初回端数の価格(a[0], a[1], a[2], a[3], a[4], a[5], a[6], a[7], a[8] === null ? 0 : a[8]); }],
    ['ODDFYIELD', 9, function (a) { return 金.初回端数の利回り(a[0], a[1], a[2], a[3], a[4], a[5], a[6], a[7], a[8] === null ? 0 : a[8]); }]
  ];
  /* ★お金は 数（と 真偽）だけを 受けます★＝省いた 引数は null */
  for (var gi = 0; gi < 金の表.length; gi++) {
    (function (組) {
      表[組[0]] = function (素) {
        var a = [];
        for (var i = 0; i < 組[1]; i++) {
          var x = 素[i];
          if (x === undefined || x === null) { a.push(null); continue; }
          if (typeof x === 'boolean') { a.push(x); continue; }
          var n = 数に(x);
          a.push(n);
        }
        try { return 組[2](a); } catch (e) { return { 誤り: 'NUM' }; }
      };
    }(金の表[gi]));
  }

  /* ══ ★FILTERXML（1個）★ ══
     ★呼び方は `lib/formula-filterxml-plug.js` の 実物を 写しました★（当て推量 0）：
       `F.取り出す(xml, 道筋, 道具)` ／ 道具 ＝ { DOMParser, XPathResult }
     ★XML を 読む 道具が 要ります★＝
       画面では 窓の 物／node では jsdom（借り物の 皮も 同じ形）。
     ★道具が 無ければ 計算しません★＝★半分 合う 答えを 出さない★
     ★答えは 実Excel の 実測★＝`docs/measured/kansuu46/golden-filterxml-2026-09-07.tsv`（37行） */
  var XML部品 = null;
  function XML道具を渡す(d) { XML部品 = d; }
  function XML道具() {
    if (XML部品 && XML部品.DOMParser && XML部品.XPathResult) return XML部品;
    if (typeof self !== 'undefined' && self.DOMParser && self.XPathResult) return self;
    return null;
  }
  表['FILTERXML'] = function (素) {
    var 道具 = XML道具();
    if (!道具) return { 誤り: 'VALUE' };      /* ★道具が 無ければ 計算しない★ */
    return XML.取り出す(素[0], 素[1], 道具);
  };

  /* ══ ★CELL（1個）★ ══
     ★呼び方は `lib/formula-cell-plug.js` の 実物を 写しました★（当て推量 0）：
       `F.セルの事(何を, { 行, 列, 別のシートか, シート名, ブック名, 中身, 空か,
                          表示形式, そろえ, 幅の点, 既定の点, 隠れ, ロックなし })`
     ★CELL は「マスの 事」を 聞く 関数★＝★表の 本体と 見た目が 要ります★。
     ⇒★マスの 事を 渡す 口★ を 作りました。★渡されなければ 計算しません★
       （半分 合う 答えを 出さない＝[[feedback_hanbun_au_keisan_wa_dasuna]]）
     ★答えは 実Excel の 実測★＝golden-cell2-2026-09-07.tsv（106行）
                              ＋ golden-zero-torinaoshi-2026-09-08.tsv（201行） */
  var マスの事を出す = null;
  function マスの事を渡す(f) { マスの事を出す = f; }
  表['CELL'] = function (素, 生, 場) {
    var 何を = 素[0];
    if (何を === null || 何を === undefined) return { 誤り: 'VALUE' };
    if (typeof 何を !== 'string') return { 誤り: 'VALUE' };
    /* ★2つ目は 無くてもよい★（実測 `=CELL("row")` は 今 いる マス） */
    var 事 = null;
    if (typeof マスの事を出す === 'function') {
      /* ★画面が 見た目を 渡して くれる 時★（book.html が 繋ぐ） */
      事 = マスの事を出す(素.length > 1 ? 素[1] : null);
    }
    if (!事) {
      /* ★見た目が 無くても 答えます★（2026-09-14／経営者1 の 指摘で 直した）
         ＝借り物の 皮（`lib/formula-cell-plug.js`）は ★既定の 見た目で 埋めて 答えて★ いました。
           自前の 皮だけ #VALUE! を 返して いた＝★同じ CELL が 2つの 振る舞い★。
         ★既定は「半分 合う 答え」では ありません★＝
           ★実Excel の まっさらな シートの 本当の 答え★（紙の 実測 … width 8／protect 1）。
         ★どの マスかは 土台が 渡します★（引数の `番地`）。無ければ 答えない。 */
      var m = 生 && 生.length > 1 ? 生[1] : null;
      var 番地 = m && m.番地 ? m.番地 : null;
      var 中 = m ? m.値 : null;
      if (m && m.種 === '四角') {
        /* ★四角を 渡されたら 左上★（実測 `=CELL("row",A1:C3)` → 1／`address` → $A$1） */
        中 = (m.並び && m.並び.length) ? (Array.isArray(m.並び[0]) ? m.並び[0][0] : m.並び[0]) : null;
      }
      if (!番地) {
        /* ★引数が 無い＝今 いる マス★（実測 `=CELL("row")` → 1）
           ★土台が 教えて くれた 時だけ★ 答える（★当て推量で 1 と 言わない★） */
        if (素.length <= 1 && 場 && 場.今のマス) { 番地 = 場.今のマス; 中 = null; }
      }
      if (!番地) return { 誤り: 'VALUE' };
      事 = セル.見た目を埋める({
        行: 番地.行, 列: 番地.列,
        中身: (中 && 中.型 === '空') ? null : (中 ? 中.値 : null),
        空か: !!(中 && 中.型 === '空') || 中 === null || 中 === undefined,
      });
      if (中 && 中.型 === '誤') 事.中身 = { 誤り: String(中.値) };
    }
    var 答 = セル.セルの事(何を, 事);
    /* ★中身を 聞かれて そのマスが 誤りなら その 誤りを そのまま 返す★（皮の 実物と 同じ） */
    if (String(何を).trim().toLowerCase() === 'contents' && 事.中身 && 事.中身.誤り) return 事.中身;
    return 答;
  };

  /* ══ ★yosoku（10個）★ ══
     ★呼び方は `lib/formula-yosoku-plug.js` から ★機械で 抜きました★（手で 写して いない）
     ★答えは 実Excel の 実測★＝答えの紙 だけで ★884本★
       （golden-convert 243／convert3 147／convert2 67／linest-hyou 119／yosoku-3kansuu 33 …）
     ★INFO と RANDARRAY は 場や でたらめに 依ります★＝紙の 実測で 確かめる */
  var 予の表 = [
    ['CONVERT',      function (素) { return 予.単位を変える(素[0], 素[1], 素[2]); }],  /* ★数に しない★＝型を 潰すと 真偽/字の #VALUE! が 消える（借り物の 皮も 生で 渡す） */
    ['ERF.PRECISE',  function (素) { return 予.誤差(数に(素[0])); }],
    ['ERFC.PRECISE', function (素) { return 予.誤差の残り(数に(素[0])); }],
    ['PERCENTOF',    function (素) { return 予.割合(表にする(素[0]), 表にする(素[1])); }],
    ['TREND',        function (素) { return 予.直線予測(表にする(素[0]), 素[1] == null ? null : 表にする(素[1]), 素[2] == null ? null : 表にする(素[2]), 素[3] == null ? true : !!素[3]); }],
    ['GROWTH',       function (素) { return 予.増える予測(表にする(素[0]), 素[1] == null ? null : 表にする(素[1]), 素[2] == null ? null : 表にする(素[2]), 素[3] == null ? true : !!素[3]); }],
    ['LINEST',       function (素) { return 予.直線の係数(表にする(素[0]), 素[1] == null ? null : 表にする(素[1]), 素[2] == null ? true : !!素[2], 素[3] == null ? false : !!素[3]); }],
    ['LOGEST',       function (素) { return 予.増えるの係数(表にする(素[0]), 素[1] == null ? null : 表にする(素[1]), 素[2] == null ? true : !!素[2], 素[3] == null ? false : !!素[3]); }],
    ['INFO',         function (素) { return 予.場の事(素[0]); }],
    ['RANDARRAY',    function (素) { return 予.でたらめの表(数に(素[0]), 数に(素[1]), 数に(素[2]), 数に(素[3]), 素[4] == null ? false : !!素[4]); }]
  ];
  for (var yi = 0; yi < 予の表.length; yi++) {
    (function (組) { 表[組[0]] = function (素) { try { return 組[1](素); } catch (e) { return { 誤り: 'VALUE' }; } }; }(予の表[yi]));
  }

  /* ══ ★nokori（14個）★ ══
     ★呼び方は `lib/formula-nokori-plug.js` を 1行ずつ 読んで 写しました★（当て推量 0）
     ★ERROR.TYPE だけ 別★＝★誤りを 受け取る 関数★なので
       `呼ぶ` の「渡された 誤りは そのまま 伝える」を ★通しては いけない★（下の 誤りを受ける 名簿） */
  function 字に(x) {
    if (x === null || x === undefined) return '';
    if (Array.isArray(x)) return 字に(x.length && Array.isArray(x[0]) ? x[0][0] : x[0]);
    return String(x);
  }
  /* ★誤りの 印（#DIV/0! 等）を エンジン式の 名（DIV_BY_ZERO 等）に 戻す★
     ＝借り物側の 皮は `v.type`（エンジン式）を 渡して いました。同じ物を 渡します。
     ★`#N/A` は 印の まま 渡すと `N/A` に なって 表に 無く、7 では なく #N/A に なります★ */
  var 印から名 = (function () {
    var o = {};
    for (var k in 誤りの名) if (Object.prototype.hasOwnProperty.call(誤りの名, k)) o[誤りの名[k]] = k;
    return o;
  }());
  /* ★★BESSEL 4つ★★（2026-09-16）
       ★覚えた 係数表は 1つも 使って いません★
         ＝級数（項の 比）・漸化式・ハンケルの 漸近形だけ
       ★★実Excel の 方が 間違って いる 所が 3件 在ります★★
         … `docs/measured/bessel-no-kotae.md`
         … 逃げ道の 門 `tests/jitsuexcel-ga-machigai.test.mjs`
       ★x ≦ 0 の K・Y は #NUM!★（`lib/bessel.js` が null を 返します） */
  var ベッセルの表 = [
    ['BESSELI', function (素) { return ベ.BESSELI(数に(素[0]), Math.trunc(数に(素[1]))); }],
    ['BESSELJ', function (素) { return ベ.BESSELJ(数に(素[0]), Math.trunc(数に(素[1]))); }],
    ['BESSELK', function (素) {
      var v = ベ.BESSELK(数に(素[0]), Math.trunc(数に(素[1])));
      return v === null ? { 誤り: 'NUM' } : v;
    }],
    ['BESSELY', function (素) {
      var v = ベ.BESSELY(数に(素[0]), Math.trunc(数に(素[1])));
      return v === null ? { 誤り: 'NUM' } : v;
    }]
  ];
  for (var bi = 0; bi < ベッセルの表.length; bi++) {
    (function (組) {
      表[組[0]] = function (素) { try { return 組[1](素); } catch (e) { return { 誤り: 'VALUE' }; } };
    }(ベッセルの表[bi]));
  }

  var 残の表 = [
    ['FINDB',           function (素) { return 残.探すB(字に(素[0]), 字に(素[1]), 素[2] == null ? 1 : 数に(素[2])); }],
    ['SEARCHB',         function (素) { return 残.探すB大小なし(字に(素[0]), 字に(素[1]), 素[2] == null ? 1 : 数に(素[2])); }],
    ['REPLACEB',        function (素) { return 残.入れ替えB(字に(素[0]), 数に(素[1]), 数に(素[2]), 字に(素[3])); }],
    ['TEXTSPLIT',       function (素) { var 縦 = 字に(素[2]); return 残.区切って分ける(字に(素[0]), 字に(素[1]), 縦 === '' ? null : 縦); }],
    ['REGEXTEST',       function (素) { return 残.正規で調べる(字に(素[0]), 字に(素[1]), 数に(素[2]) === 1); }],
    ['REGEXEXTRACT',    function (素) { return 残.正規で取り出す(字に(素[0]), 字に(素[1]), 数に(素[3]) === 1); }],
    ['REGEXREPLACE',    function (素) { return 残.正規で入れ替える(字に(素[0]), 字に(素[1]), 字に(素[2]), 数に(素[4]) === 1); }],
    ['SORTBY',          function (素) { return 残.別の列で並べる(表にする(素[0]), 表にする(素[1]), 素[2] == null ? 1 : 数に(素[2])); }],
    ['MUNIT',           function (素) { return 残.単位行列(数に(素[0])); }],
    ['MINVERSE',        function (素) { return 残.逆行列(表にする(素[0])); }],
    ['PERCENTRANK.INC', function (素) { return 残.順位の割合(表にする(素[0]), 数に(素[1]), 素[2] == null ? undefined : 数に(素[2]), true); }],
    ['PERCENTRANK.EXC', function (素) { return 残.順位の割合(表にする(素[0]), 数に(素[1]), 素[2] == null ? undefined : 数に(素[2]), false); }],
    ['PROB',            function (素) { return 残.確率(表にする(素[0]), 表にする(素[1]), 数に(素[2]), 素[3] == null ? undefined : 数に(素[3])); }],
    ['ERROR.TYPE',      function (素) {
      var x = 素[0];
      if (Array.isArray(x)) x = x.length && Array.isArray(x[0]) ? x[0][0] : x[0];
      if (!(x && typeof x === 'object' && x.誤)) return { 誤り: 'NA' };   /* ★誤りで ない＝#N/A★（借り物側と 同じ） */
      return 残.誤りを番号に(印から名[x.誤] || String(x.誤));
    }]
  ];
  for (var ni = 0; ni < 残の表.length; ni++) {
    (function (組) { 表[組[0]] = function (素) { try { return 組[1](素); } catch (e) { return { 誤り: 'VALUE' }; } }; }(残の表[ni]));
  }

  /* ★素の 値を 表（[[…],…]）に する★＝中身の 部品は 表を 受ける 作り */
  function 表にする(x) {
    if (x === null || x === undefined) return [[x]];
    if (Array.isArray(x)) return 中.表にする ? 中.表にする(x) : x;
    return [[x]];
  }
  /* ★数に する（省いた 引数は null）★ */
  function 数に(x) {
    if (x === null || x === undefined) return null;
    if (Array.isArray(x)) return 数に(x.length && Array.isArray(x[0]) ? x[0][0] : x[0]);
    var n = Number(x);
    return isFinite(n) ? n : null;
  }
  /* ★2つ目 以降を 全部 数に ほどいて 並べる★（CHOOSECOLS / CHOOSEROWS） */
  function 番号たち(素) {
    var 出 = [];
    for (var i = 1; i < 素.length; i++) {
      var t = 表にする(素[i]);
      for (var r = 0; r < t.length; r++) {
        for (var c = 0; c < t[r].length; c++) {
          var n = Number(t[r][c]);
          if (isFinite(n)) 出.push(n);
        }
      }
    }
    return 出;
  }

  /* ★誤りを 受け取る 関数★＝ここに 載せた 物だけ 誤りを そのまま 渡す */
  var 誤りを受ける = { 'ERROR.TYPE': true };

  /* ★知らない 名前は 呼ばない★＝★半分 合う 計算を 出さない★ */
  function 呼ぶ(名前, 引数たち, 場) {
    var f = 表[String(名前).toUpperCase()];
    if (!f) return null;                      /* ★null＝この 皮は 知らない★ */
    /* ★渡された 誤りは そのまま 伝える★
       ＝2026-09-14 実測。`=BAHTTEXT(1/0)` は 実Excel で ★#DIV/0!★。
       私は これを 握り潰して ★0バーツ★ と 答えて いました。 */
    /* ★誤りを「受け取る」関数は 素通しの 外★
       ＝`=ERROR.TYPE(1/0)` は 実Excel で ★2★。素通しすると ★#DIV/0!★ を 返して しまう。 */
    if (!誤りを受ける[String(名前).toUpperCase()]) {
      for (var i = 0; i < 引数たち.length; i++) {
        var a = 引数たち[i];
        if (a && a.種 === '直' && a.値 && a.値.型 === '誤') return a.値;
        if (a && a.種 === 'マス' && a.値 && a.値.型 === '誤') return a.値;
      }
    }
    try {
      /* ★生の 引数も 渡す★＝`種`（直/マス/四角）や ★番地★を 要る 関数が 在る（CELL）。
         ★今まで の 関数は 2つ目を 見ないので 何も 変わりません★ */
      return 包む(f(素にする(引数たち), 引数たち, 場 || {}));
    } catch (e) {
      /* ★★`process` は ブラウザに 在りません★★（2026-09-18・★繋ぐ 前に 見つけた★）
           ★node では 通ります★＝★今まで 1度も 出ませんでした★
           ⇒★★ここが 投げると ★皮が 答える 65個 が 全部 落ちます★★★
           ⇒★記憶「JS層は 2階建て」／「読み込む != 登録される」の 一族★ */
      if (typeof process !== 'undefined' && process.env && process.env.MISERU) {
        console.error('  ★落ちた★ ' + 名前 + ' ... ' + e.message);
      }
      return 計.誤('#VALUE!');
    }
  }

  /* ★繋いだ 関数の 名前（見張りが 数える）★ */
  function 名前たち() { return Object.keys(表); }

  return { 呼ぶ: 呼ぶ, 名前たち: 名前たち, 表: 表, ほどく: ほどく, 包む: 包む,
           XML道具を渡す: XML道具を渡す, マスの事を渡す: マスの事を渡す };
});
