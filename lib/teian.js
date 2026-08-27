/* teian.js — ★8 提案＝うちから「こう直せますよ」と言う★（純関数・0円）
 *
 *  ★正本（2026-08-22 司さん決定／指示役 08-27）★
 *    ★気づくのは機械。AIは説明するだけ★
 *      ＝提案を ★見つけるのは 0円★（AIを1回も呼ばない）。AIは 言葉にするだけ。
 *    ★競合は「聞かれたら答える」＝受け身。自分から言わない★ ←★ここが差★
 *
 *  ★実物で 数えた材料（2026-08-27・代行計算表2026.xlsb）★
 *    式 15,126本 ／ ★INDEX+MATCH 6,832本★ ／ IF 2,548本 ／ VLOOKUP 0本
 *    ★同じ MATCH を 1つの式の中で 3回 書いている★物が 多数
 *      例) =IF(INDEX(計算!V2:V32, MATCH(給料1!B4, 計算!A2:A32, 0))<=0,"",
 *              INDEX(計算!E2:E32, MATCH(給料1!B4,計算!A2:A32, 0))
 *             -INDEX(計算!S2:S32, MATCH(B4,計算!A2:A32,0)))
 *    ★IFの入れ子は 最大1段★（この本には 深い入れ子は 無かった＝★無い物は 出さない★）
 *
 *  ★決まり★
 *    ・★勝手に直さない★＝提案するだけ。直すのは 客が押してから（差分プレビューを通す）
 *    ・★相手のExcelの版で 使えない直し方は 出さない★（lib/excel-version.js に聞く）
 *    ・★お金の数字を 提案文に書かない★
 *    ・★客に見せる字に ★ を書かない★
 */
(function (root) {
  'use strict';

  var EV = (typeof require === 'function' && typeof module === 'object')
    ? require('./excel-version.js')
    : root.ExcelVersion;

  function 列の字(c) {
    var s = '';
    c = Number(c);
    while (c >= 0) { s = String.fromCharCode(65 + (c % 26)) + s; c = Math.floor(c / 26) - 1; }
    return s;
  }
  function セルの名(rc) {
    var p = String(rc).split(',');
    return 列の字(Number(p[1])) + (Number(p[0]) + 1);
  }

  /** その版で その関数が 使えるか（分からない時は ★使えない側に倒す★＝嘘の提案をしない） */
  function 使えるか(関数名, 版) {
    try {
      if (!EV || !EV.checkFormula) return false;
      /* ★知らない版は 勧めない★＝checkFormula は 知らない版を 365 とみなす作りなので、
         ここで ★版そのものを 確かめる★（分からない物に 勧めるのが いちばん危ない） */
      var 版一覧 = EV.VER_NUM || {};
      if (!Object.prototype.hasOwnProperty.call(版一覧, 版 || 'excel_365')) return false;
      /* ★checkFormula は「使えない物の一覧」を そのまま返す★（配列）。
         ★1つも返らなければ 使える★。★分からない時は 使えない側に倒す★＝嘘の提案をしない。 */
      var 悪い = EV.checkFormula('=' + 関数名 + '(1)', 版 || 'excel_365');
      if (!悪い || typeof 悪い.length !== 'number') return false;
      return 悪い.length === 0;
    } catch (e) { return false; }
  }

  /** 1つの式の中で 同じ引数のMATCHを 何回 書いているか */
  function 同じMATCHの数(f) {
    var 出 = {};
    var re = /MATCH\s*\(([^()]*(?:\([^()]*\)[^()]*)*)\)/gi;
    var m;
    while ((m = re.exec(f))) {
      var 中 = m[1].replace(/\s+/g, '');
      出[中] = (出[中] || 0) + 1;
    }
    var 一番 = 0;
    for (var k in 出) if (Object.prototype.hasOwnProperty.call(出, k) && 出[k] > 一番) 一番 = 出[k];
    return 一番;
  }

  /** IF の入れ子の深さ（★「IFが何回 出るか」ではなく 本当の深さ★） */
  function IFの深さ(f) {
    var s = String(f || '');
    var 深さ = 0, 最大 = 0;
    var 積み = [];
    for (var i = 0; i < s.length; i++) {
      if (s.charAt(i) === '(') {
        var 前 = s.slice(Math.max(0, i - 3), i).toUpperCase();
        var isIF = /(^|[^A-Z])IF$/.test(前);
        積み.push(isIF);
        if (isIF) { 深さ++; if (深さ > 最大) 最大 = 深さ; }
      } else if (s.charAt(i) === ')') {
        if (積み.pop()) 深さ--;
      }
    }
    return 最大;
  }

  /* ══ ★見つける（0円・AIを1回も呼ばない）★ ══════════════════ */
  var 決まり = {
    IFの深さ: 4,        /* これ以上 深い入れ子は 読めない */
    出す数: 5,          /* 一度に出す 提案の数（多すぎると 誰も読まない） */
    例の数: 3,          /* 1つの提案に付ける 場所の例 */
  };

  /**
   * @param {Array} sheets
   * @param {{版:string, 出す数:number}} opt  版 … 'excel_365' 等（相手のExcelの版）
   * @returns {{提案:Array, 見た式:number, かかった秒:number}}
   */
  function 見つける(sheets, opt) {
    opt = opt || {};
    var 版 = opt.版 || 'excel_365';
    var t0 = (typeof performance !== 'undefined' && performance.now) ? performance.now() : Date.now();
    var 数 = { indexMatch: 0, matchの重複: 0, vlookup: 0, ifの入れ子: 0 };
    var 例 = { indexMatch: [], matchの重複: [], vlookup: [], ifの入れ子: [] };
    var 見た式 = 0;
    for (var si = 0; si < (sheets || []).length; si++) {
      var sh = sheets[si], data = sh.data || {};
      for (var k in data) {
        if (!Object.prototype.hasOwnProperty.call(data, k)) continue;
        var f = data[k] && data[k].f;
        if (!f || String(f).charAt(0) !== '=') continue;
        見た式++;
        var 場所 = sh.name + '!' + セルの名(k);
        if (/\bVLOOKUP\s*\(/i.test(f)) {
          数.vlookup++;
          if (例.vlookup.length < 決まり.例の数) 例.vlookup.push({ 場所: 場所, 式: String(f) });
        } else if (/\bINDEX\s*\(/i.test(f) && /\bMATCH\s*\(/i.test(f)) {
          数.indexMatch++;
          if (例.indexMatch.length < 決まり.例の数) 例.indexMatch.push({ 場所: 場所, 式: String(f) });
        }
        if (同じMATCHの数(f) >= 2) {
          数.matchの重複++;
          if (例.matchの重複.length < 決まり.例の数) 例.matchの重複.push({ 場所: 場所, 式: String(f) });
        }
        if (IFの深さ(f) >= 決まり.IFの深さ) {
          数.ifの入れ子++;
          if (例.ifの入れ子.length < 決まり.例の数) 例.ifの入れ子.push({ 場所: 場所, 式: String(f) });
        }
      }
    }

    var 提案 = [];
    var XL = 使えるか('XLOOKUP', 版);
    var IFS = 使えるか('IFS', 版);
    if (数.vlookup) {
      提案.push(作る('vlookup_to_xlookup', 数.vlookup, 例.vlookup, XL, 版));
    }
    if (数.indexMatch) {
      提案.push(作る('indexmatch_to_xlookup', 数.indexMatch, 例.indexMatch, XL, 版));
    }
    if (数.matchの重複) {
      提案.push(作る('match_no_juufuku', 数.matchの重複, 例.matchの重複, XL, 版));
    }
    if (数.ifの入れ子) {
      提案.push(作る('if_nest', 数.ifの入れ子, 例.ifの入れ子, IFS, 版));
    }
    /* ★多い順★（効く所から出す）／★出しすぎない★ */
    提案.sort(function (a, b) { return b.何本 - a.何本; });
    提案 = 提案.slice(0, opt.出す数 || 決まり.出す数);
    var t1 = (typeof performance !== 'undefined' && performance.now) ? performance.now() : Date.now();
    return { 提案: 提案, 見た式: 見た式, かかった秒: Math.max(0.01, Math.round((t1 - t0) / 10) / 100) };
  }

  /* ★客に見せる言葉は ここ1か所で作る★（画面に散らさない・★に注意） */
  var 言い方 = {
    vlookup_to_xlookup: {
      題: 'VLOOKUP を もっと短く書けます',
      なに: '左から数えて◯列目、という書き方は 列を1つ足すだけで ずれます。',
      こう: 'XLOOKUP なら「どの列を見て、どの列を返すか」を そのまま書けます。',
    },
    indexmatch_to_xlookup: {
      題: 'INDEX と MATCH の組み合わせを 1つにできます',
      なに: '同じ場所を 2つの関数で 指しているので、直す時に 両方 直す必要があります。',
      こう: 'XLOOKUP なら 1つで書けます。式が短くなり、直す所も 1か所になります。',
    },
    match_no_juufuku: {
      題: '同じ探し物を 1つの式の中で 何度もしています',
      なに: '同じ MATCH を 2回以上 書いているので、その回数だけ 探し直しています。',
      こう: 'XLOOKUP に置き換えると 1回で済みます。式も短くなります。',
    },
    if_nest: {
      題: 'IF の入れ子が 深くなっています',
      なに: '入れ子が深いと、どの条件で どうなるかが 読めなくなります。',
      こう: 'IFS を使うと 条件と答えを 横に並べて書けます。',
    },
  };

  function 作る(種類, 何本, 例, 使える, 版) {
    var w = 言い方[種類];
    return {
      種類: 種類,
      題: w.題,
      何本: 何本,
      場所: 例.map(function (x) { return x.場所; }),
      例: 例[0] ? 例[0].式 : '',
      本文: w.なに + (何本 > 1 ? ('この形が ' + 何本 + 'か所 あります。') : ''),
      直し方: 使える ? w.こう : null,
      /* ★相手の版で 使えない直し方は 出さない★（嘘の提案をしない） */
      使えない理由: 使える ? null : ('お使いの Excel では この直し方が 使えません（' + 版 + '）。'),
    };
  }

  /** ★向こうから出す1行★（押していないのに 出る所の字） */
  function 知らせの字(結果) {
    if (!結果 || !結果.提案 || !結果.提案.length) return null;
    /* ★直し方が無い物は 向こうから 出さない★
       ＝お使いの版で 直せない事を こちらから言っても、客は 何も出来ない。
       （押して見た時には 出す＝★隠してはいない★） */
    var 出せる = 結果.提案.filter(function (t) { return !!t.直し方; });
    if (!出せる.length) return null;
    var 合計 = 0;
    for (var i = 0; i < 出せる.length; i++) 合計 += 出せる[i].何本;
    return {
      題: 'もっと短く書ける所が あります',
      本文: '式を ' + 結果.見た式 + '本 見て、' + 出せる.length + '通り・'
        + 合計 + 'か所 見つけました。AIは使っていません。',
      つぎ: '中身を出します。直すかどうかは 見てから 決めてください。',
    };
  }

  var api = {
    決まり: 決まり,
    見つける: 見つける,
    知らせの字: 知らせの字,
    同じMATCHの数: 同じMATCHの数,
    IFの深さ: IFの深さ,
    使えるか: 使えるか,
    列の字: 列の字,
    セルの名: セルの名,
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.Teian = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
