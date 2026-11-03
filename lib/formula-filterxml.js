/* formula-filterxml.js — ★FILTERXML（XML から 取り出す）★（2026-09-07）
 *
 *  ★★XPath は 自分で 作りません★★
 *    ブラウザは ★本物の XPath★を 持っています（`document.evaluate`）。
 *    ⇒★自分で 少しだけ 作ると ★通る 式と 通らない 式が 混ざる★★
 *      （2026-09-07 の 昼に 私が「出さない」と 決めた 理由が それでした）
 *    ⇒★★ブラウザの 物を そのまま 使えば ★全部 通ります★★★
 *    ⇒ 試験は jsdom（同じ `document.evaluate` を 持っている）で 押す
 *
 *  ★★答えは 実Excel に 打たせて 取った★★
 *    `docs/measured/kansuu46/golden-filterxml-2026-09-07.tsv`（37本）
 *    Excel 16.0 build 20326（UI 1041）
 *
 *  ★実測で 分かった 実Excel の くせ★
 *    ・見つかった 分を ★縦1列★で 返す（//b は 3行・//c は 2行）
 *    ・★属性も 取れる★（//c/@id → x）
 *    ・★`text()` は #VALUE!★（字の かたまりは 返さない）
 *    ・★1つも 見つからなければ #VALUE!★（空では ない）
 *    ・★数を 返す 式（count(//b)）も #VALUE!★（node で ない）
 *    ・壊れた XML／空の XML／空の XPath／おかしな XPath／数 … 全部 #VALUE!
 *
 *  見張り: tests/formula-filterxml.test.mjs
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.FormulaFilterXml = factory();
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var 誤り = function (種) { return { 誤り: 種 }; };

  /** ★XML と XPath から 取り出す★
   *  @param 部品 { DOMParser, XPathResult }（画面は window／試験は jsdom を 入れる）
   *  @returns {Array<Array<string>>|{誤り:string}} 縦1列の 表 */
  function 取り出す(xml, xpath, 部品) {
    /* ★字で ない 物は 断る★（実測 … FILTERXML(123,"//b") は #VALUE!） */
    if (typeof xml !== 'string' || xml === '') return 誤り('VALUE');
    if (typeof xpath !== 'string' || xpath.trim() === '') return 誤り('VALUE');
    var P = 部品 && 部品.DOMParser;
    var R = 部品 && 部品.XPathResult;
    if (!P || !R) return 誤り('VALUE');

    var doc;
    try { doc = new P().parseFromString(xml, 'text/xml'); } catch (e) { return 誤り('VALUE'); }
    if (!doc) return 誤り('VALUE');
    /* ★壊れた XML は `parsererror` が 混ざる★（ブラウザも jsdom も 同じ 出し方） */
    try {
      if (doc.getElementsByTagName('parsererror').length) return 誤り('VALUE');
    } catch (e) { return 誤り('VALUE'); }
    if (!doc.documentElement) return 誤り('VALUE');

    var 結;
    try {
      結 = doc.evaluate(xpath, doc, null, R.ORDERED_NODE_SNAPSHOT_TYPE, null);
    } catch (e) {
      /* ★おかしな XPath／node で ない 答え（count(...)）は ここで 投げる★ */
      return 誤り('VALUE');
    }
    if (!結 || typeof 結.snapshotLength !== 'number') return 誤り('VALUE');
    var n = 結.snapshotLength;
    if (!n) return 誤り('VALUE');           /* ★1つも 無ければ #VALUE!★（空では ない） */

    var 出 = [];
    for (var i = 0; i < n; i++) {
      var 節 = 結.snapshotItem(i);
      if (!節) return 誤り('VALUE');
      /* ★要素（1）と 属性（2）だけ★＝実測 … `text()`（3）は #VALUE! */
      if (節.nodeType !== 1 && 節.nodeType !== 2) return 誤り('VALUE');
      var v = (節.nodeType === 2) ? (節.value != null ? 節.value : 節.textContent) : 節.textContent;
      出.push([v == null ? '' : String(v)]);
    }
    return 出;
  }

  function 足した名前() { return ['FILTERXML']; }

  return { 取り出す: 取り出す, 足した名前: 足した名前 };
}));
