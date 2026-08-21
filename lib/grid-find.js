/* grid-find.js — ★検索と置換（純関数）★
 *
 *  ★実Excel 16.0.20228 を COM で動かして測った真値（2026-08-21）★
 *
 *  ① ★置換は「式」を見る（答えではない）★
 *       A1=2 ／ B1==A1*2（答え4）に 「2」→「9」
 *       → A1=9 ／ ★B1 の式が =A1*9 になり 答えが 81★
 *  ② ★大文字と小文字は 既定で 区別しない★（abc/ABC/Abc すべて置き換わる）
 *       区別する設定にすると abc だけ
 *  ③ ★一部でも置き換える（既定）★（「あいうえお」→「Xうえお」）
 *       「セル全体が同じ時だけ」にすると 「あい」だけ
 *  ④ ★全角と半角は 別物★（ABC を探しても ＡＢＣ は置き換わらない）
 *  ⑤ ★* と ? はワイルドカード★（「あ*い」で あい・あうい・あ*い が全部あたる）
 *       ★~* と書くと 「*」そのもの★（あ*い だけ）／★? は1文字★（あうい だけ）
 *  ⑥ ★次を探すのは「選んだセルの次」から★（A1:B2 の A1・B1・A2 に x が在り、
 *       A1 を選んで探すと ★B1★ が出る）
 *  ⑦ ★★式の中の数の一部まで置き換わる★★
 *       A1=10 ／ A2==A1+100 に 「10」→「20」
 *       → A2 が ★=A1+200★ になり 答えが 110 → ★220★
 *       ＝ ★Excel 自身が こうなる★。うちも同じにするが、
 *         ★黙って値が変わる★ ので ★式を何本 書き換えたかを必ず知らせる★（画面側の仕事）。
 */
(function (root) {
  'use strict';

  /** 探す相手の字。★式のセルは 式の文字を見る★（実測①） */
  function textOf(cell) {
    if (!cell) return '';
    if (typeof cell.f === 'string' && cell.f.charAt(0) === '=') return cell.f;
    var v = cell.v;
    if (v === undefined || v === null) return '';
    return String(v);
  }

  /** Excel の書き方（* ? ~）を 正規表現にする */
  function toRegExp(what, opts) {
    opts = opts || {};
    var s = String(what === undefined || what === null ? '' : what);
    var out = '';
    for (var i = 0; i < s.length; i++) {
      var ch = s.charAt(i);
      if (ch === '~') {                       /* ★~ の次は そのままの字★（実測⑤） */
        var next = s.charAt(i + 1);
        if (next === '*' || next === '?' || next === '~') { out += next.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); i++; continue; }
        out += '~'; continue;
      }
      if (ch === '*') { out += '[\\s\\S]*'; continue; }   /* いくつでも */
      if (ch === '?') { out += '[\\s\\S]'; continue; }    /* 1文字 */
      out += ch.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    }
    /* ★セル全体が同じ時だけ★ なら 端を止める */
    if (opts.whole) out = '^' + out + '$';
    /* ★大文字と小文字は 既定で 区別しない★（実測②） */
    return new RegExp(out, opts.matchCase ? 'g' : 'gi');
  }

  /** そのセルが あたるか */
  function match(cell, what, opts) {
    if (what === undefined || what === null || what === '') return false;
    var t = textOf(cell);
    if (t === '') return false;
    return toRegExp(what, opts).test(t);
  }

  /**
   * そのセルの字を置き換えた結果を返す（変わらなければ null）
   * ★式のセルは 式の文字を置き換える★＝答えが変わる（実測①⑦）
   */
  function replacedText(cell, what, to, opts) {
    if (what === undefined || what === null || what === '') return null;
    var t = textOf(cell);
    if (t === '') return null;
    var re = toRegExp(what, opts);
    var next = t.replace(re, String(to === undefined || to === null ? '' : to));
    return next === t ? null : next;
  }

  /**
   * ★次を探す★ … 今いる所の「次」から 行の順に探す（実測⑥）
   * @param {function(number,number):Object} get
   * @param {number} rows,cols 見る広さ
   * @param {{r:number,c:number}} 今 いま選んでいる所
   * @returns {{r:number,c:number}|null}
   */
  function findNext(get, rows, cols, 今, what, opts) {
    if (what === undefined || what === null || what === '') return null;
    var 全部 = rows * cols;
    var start = (今.r * cols + 今.c + 1) % 全部;      /* ★次から★ */
    for (var n = 0; n < 全部; n++) {
      var i = (start + n) % 全部;
      var r = Math.floor(i / cols), c = i % cols;
      if (match(get(r, c), what, opts)) return { r: r, c: c };
    }
    return null;
  }

  /** その式が 式かどうか（何本 書き換えたかを数えるのに使う） */
  function isFormula(cell) {
    return !!(cell && typeof cell.f === 'string' && cell.f.charAt(0) === '=');
  }

  var api = { textOf: textOf, toRegExp: toRegExp, match: match, replacedText: replacedText, findNext: findNext, isFormula: isFormula };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.GridFind = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
