/* horu.js — ★7 掘る＝AIが「もっと見せて」と言えるようにする★（純関数・0円）
 *
 *  ★正本（2026-08-22 司さん決定）★
 *    ★AIが「もっと見せて」と言えるようにする … 範囲を指定→その場で返す。
 *      何度でも掘れる＝実質 全部見られる／掘っても値段はほぼ増えない★
 *
 *  ★指示役の決まり（2026-08-26）★
 *    ・★答えには「どのセルを見て言ったか」を必ず付ける★（見ていない物を 見たと言わせない）
 *    ・★お金の数字（合計・請求額・給与）はAIに出させない★
 *      ＝AIは {{シート!セル}} と書く。★値を入れるのは うちの計算★
 *    ・★1回にAIへ渡すのは 2万トークンまで★（2026-08-09 司さんの決定）
 *    ・★掘った回数は 数え場に残す★
 *
 *  ★暴走させない★
 *    ・1回に返すセルは ★上限まで★（超えたら 減らして「ここまで返した」と書く）
 *    ・掘れる回数は ★上限まで★（超えたら 掘らずに 今ある材料で答えさせる）
 *    ・★止めた時は 黙らない★（何回で止めたかを 客にもAIにも言う）
 */
(function (root) {
  'use strict';

  var 決まり = {
    掘れる回数: 5,        /* 1相談で 掘れる回数（超えたら 今ある材料で答えさせる） */
    セル数: 2000,    /* 1回に返すセルの数 */
    字数: 12000,     /* 1回に返す字の数（2万トークンの内側に収める） */
  };

  function 列の番号(字) {
    var s = String(字 || '').toUpperCase();
    if (!/^[A-Z]{1,3}$/.test(s)) return -1;
    var c = 0;
    for (var i = 0; i < s.length; i++) c = c * 26 + (s.charCodeAt(i) - 64);
    return c - 1;
  }
  function 列の字(c) {
    var s = '';
    c = Number(c);
    while (c >= 0) { s = String.fromCharCode(65 + (c % 26)) + s; c = Math.floor(c / 26) - 1; }
    return s;
  }

  /** 'A1:D50' / 'A1' / 'シート名!A1:D50' を { シート, r0,c0,r1,c1 } に */
  function 範囲を読む(字, 既定のシート) {
    var s = String(字 || '').trim();
    var シート = 既定のシート || '';
    var m0 = /^(?:'((?:[^']|'')+)'|([^!]+))!(.+)$/.exec(s);
    if (m0) {
      シート = (m0[1] !== undefined) ? m0[1].replace(/''/g, "'") : m0[2];
      s = m0[3];
    }
    var m = /^\$?([A-Z]{1,3})\$?(\d{1,7})(?::\$?([A-Z]{1,3})\$?(\d{1,7}))?$/i.exec(s.trim());
    if (!m) return null;
    var c0 = 列の番号(m[1]), r0 = Number(m[2]) - 1;
    var c1 = m[3] ? 列の番号(m[3]) : c0;
    var r1 = m[4] ? Number(m[4]) - 1 : r0;
    if (c0 < 0 || c1 < 0 || r0 < 0 || r1 < 0) return null;
    return {
      シート: シート,
      r0: Math.min(r0, r1), c0: Math.min(c0, c1),
      r1: Math.max(r0, r1), c1: Math.max(c0, c1),
    };
  }

  /**
   * ★AIの返事から「もっと見せて」を読む★
   * 返し方（地図で教えてある形）… {"もっと見せて":["計算!A1:D50", "売上表!A1:C10"]}
   * @returns {{掘る:Array<string>, 答え:string}}
   */
  function 頼みを読む(text) {
    var s = String(text == null ? '' : text);
    var 掘る = [];
    var m = s.match(/\{[\s\S]*?"もっと見せて"[\s\S]*?\}/);
    if (m) {
      try {
        var o = JSON.parse(m[0]);
        var a = (o && o['もっと見せて']) || [];
        for (var i = 0; i < a.length; i++) if (a[i]) 掘る.push(String(a[i]));
      } catch (e) { /* 読めない頼みは 無い物として扱う（答えとして出す） */ }
    }
    return { 掘る: 掘る, 答え: 掘る.length ? '' : s };
  }

  /**
   * ★頼まれた範囲を その場で返す★（0円・AIを呼ばない）
   * @returns {{字:string, 返したセル:number, 頼まれたセル:number, 減らした:boolean, なぜ:string}}
   */
  function 範囲を出す(sheets, 頼み, opt) {
    opt = opt || {};
    var 上限セル = opt.セル数 || 決まり.セル数;
    var 上限字 = opt.字数 || 決まり.字数;
    var 出 = [];
    var 返した = 0, 頼まれた = 0, 減らした = false, なぜ = '';
    for (var i = 0; i < 頼み.length; i++) {
      var 範囲 = 範囲を読む(頼み[i], opt.今のシート);
      if (!範囲) { 出.push('（' + 頼み[i] + ' … 場所の書き方が 読めませんでした）'); continue; }
      var sh = null;
      for (var k = 0; k < sheets.length; k++) if (sheets[k].name === 範囲.シート) { sh = sheets[k]; break; }
      if (!sh && !範囲.シート && sheets.length) sh = sheets[0];
      if (!sh) { 出.push('（' + 頼み[i] + ' … そのシートは 在りません）'); continue; }
      var 幅 = (範囲.c1 - 範囲.c0 + 1), 高 = (範囲.r1 - 範囲.r0 + 1);
      頼まれた += 幅 * 高;
      var r1 = 範囲.r1;
      if (幅 * 高 > 上限セル) {
        r1 = 範囲.r0 + Math.max(0, Math.floor(上限セル / Math.max(1, 幅)) - 1);
        減らした = true;
        なぜ = '大きすぎるので 行を減らしました';
      }
      出.push('## ' + sh.name + '!' + 列の字(範囲.c0) + (範囲.r0 + 1)
        + ':' + 列の字(範囲.c1) + (r1 + 1));
      for (var r = 範囲.r0; r <= r1; r++) {
        var 行 = [];
        for (var c = 範囲.c0; c <= 範囲.c1; c++) {
          var cell = (sh.data || {})[r + ',' + c];
          var v = '';
          if (cell) {
            /* ★式が在れば 式も見せる★（AIが「どこを見ているか」を辿れる） */
            if (cell.f && String(cell.f).charAt(0) === '=') v = String(cell.f);
            else if (cell.v !== undefined && cell.v !== null) v = String(cell.v);
          }
          行.push(v);
          返した++;
        }
        出.push(列の字(範囲.c0) + (r + 1) + '\t' + 行.join('\t'));
      }
      if (r1 < 範囲.r1) 出.push('（' + (範囲.r1 - r1) + '行は 出していません。要るなら もう一度 頼んでください）');
    }
    var 字 = 出.join('\n');
    if (字.length > 上限字) {
      字 = 字.slice(0, 上限字) + '\n（長いので ここまでにしました。続きは もう一度 頼んでください）';
      減らした = true;
      なぜ = なぜ || '長すぎるので 途中までにしました';
    }
    return { 字: 字, 返したセル: 返した, 頼まれたセル: 頼まれた, 減らした: 減らした, なぜ: なぜ };
  }

  /** ★掘れる回数を 使い切ったか★（使い切ったら 掘らずに 今ある材料で答えさせる） */
  function もう掘れないか(掘った回数, opt) {
    var 上限 = (opt && opt.掘れる回数) || 決まり.掘れる回数;
    return 掘った回数 >= 上限;
  }
  function もう掘れない時の言葉(掘った回数, opt) {
    var 上限 = (opt && opt.掘れる回数) || 決まり.掘れる回数;
    return 'これ以上は 掘れません（' + 上限 + '回まで）。'
      + '今 見た所（' + 掘った回数 + '回ぶん）だけで 答えてください。'
      + '足りない所は「ここが分かりません」と はっきり書いてください。';
  }

  /* ══ ★答えの決まり★ ══════════════════════════════════════
     ①★どのセルを見て言ったか★が 無い答えは 通さない
     ②★金額は AIに書かせない★＝{{シート!セル}} を うちの計算結果に差し替える */

  var 見た所の印 = '見た所:';

  /** ★答えに「見た所」が付いているか★（付いていなければ 付け直しを頼む） */
  function 見た所が在るか(答え) {
    return String(答え || '').indexOf(見た所の印) >= 0;
  }
  function 見た所を付けてと言う() {
    return '答えの最後に「' + 見た所の印 + ' シート名!セル, …」を 必ず書いてください。'
      + '見ていないセルは 書かないでください。';
  }

  /**
   * ★{{シート!セル}} を うちの計算結果に差し替える★
   *   ＝★お金の数字を AIに出させない★（合計・請求額・給与）
   * @param {Function} 値をとる (シート, r, c) => 値
   * @returns {{字:string, 差し込んだ:number, 見つからない:Array}}
   */
  function 値を差し込む(答え, sheets, 値をとる) {
    var s = String(答え == null ? '' : 答え);
    var 差し込んだ = 0;
    var 見つからない = [];
    var 出 = s.replace(/\{\{([^{}]+)\}\}/g, function (全, 中) {
      var 範囲 = 範囲を読む(中.trim(), '');
      if (!範囲) { 見つからない.push(中.trim()); return '（' + 中.trim() + ' … 場所が 読めません）'; }
      var sh = null;
      for (var k = 0; k < sheets.length; k++) if (sheets[k].name === 範囲.シート) { sh = sheets[k]; break; }
      if (!sh) { 見つからない.push(中.trim()); return '（' + 中.trim() + ' … そのシートは 在りません）'; }
      var v = 値をとる ? 値をとる(sh, 範囲.r0, 範囲.c0) : undefined;
      if (v === undefined || v === null || v === '') {
        見つからない.push(中.trim());
        return '（' + 中.trim() + ' … 空です）';
      }
      差し込んだ++;
      return String(v);
    });
    return { 字: 出, 差し込んだ: 差し込んだ, 見つからない: 見つからない };
  }

  var api = {
    決まり: 決まり,
    範囲を読む: 範囲を読む,
    頼みを読む: 頼みを読む,
    範囲を出す: 範囲を出す,
    もう掘れないか: もう掘れないか,
    もう掘れない時の言葉: もう掘れない時の言葉,
    見た所の印: 見た所の印,
    見た所が在るか: 見た所が在るか,
    見た所を付けてと言う: 見た所を付けてと言う,
    値を差し込む: 値を差し込む,
    列の字: 列の字,
    列の番号: 列の番号,
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.Horu = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
