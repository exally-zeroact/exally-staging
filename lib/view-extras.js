/* view-extras.js — ★表示タブの「表示」組★（ルーラー／ナビゲーション／セルにフォーカス）2026-08-30
 *
 *  ★なぜ 今ごろ 作るか★
 *    2026-08-30 に 実Excel（16.0 build 20326）を ★測り直した★ら、
 *    ★「表示｜表示」という 組が まるごと 正本から 抜けていた★。
 *    抜けた 訳は ★自分の 絞り込み★（組の名前が タブ名と 同じ物を 捨てていた）。
 *    ⇒ 目盛線・見出し・数式バー・ルーラー・ナビゲーション・セルにフォーカス
 *      ＝★Excelの ごく 基本★が 7個 まるごと 無かった。
 *
 *  ★ここに 入れる物★（計算だけ・画面は 触らない＝試験しやすくする）
 *    ・ルーラーの 目盛り（cm）
 *    ・ナビゲーションの 一覧（シート／名前／テーブル）
 *    ・セルにフォーカスの 帯（どの 行・列を 塗るか）
 *
 *  ★実測（うちの 画面で 測って 決めた）★
 *    1cm = 96 / 2.54 = 37.7952… px（CSSの 決まり。96dpi）
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.ViewExtras = factory();
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var CM = 96 / 2.54;   /* 1cm の px（CSSの 決まり） */

  /**
   * ★ルーラーの 目盛り★
   * @param {number} 幅px   目盛りを 引く 長さ
   * @param {number} [倍率] 画面の 拡大率（1 = 100%）
   * @returns {{x:number, 大きい:boolean, 札:string}[]}
   */
  function 目盛り(幅px, 倍率) {
    var z = (typeof 倍率 === 'number' && 倍率 > 0) ? 倍率 : 1;
    var 出 = [];
    if (!(幅px > 0)) return 出;
    var 刻み = CM * z / 2;                     /* 0.5cm ごと */
    for (var i = 0; ; i++) {
      var x = i * 刻み;
      if (x > 幅px) break;
      var 大 = (i % 2 === 0);                  /* 1cm ちょうど */
      出.push({ x: x, 大きい: 大, 札: 大 ? String(i / 2) : '' });
    }
    return 出;
  }

  /**
   * ★ナビゲーションの 一覧★
   * @param {{sheets:{name:string}[], activeSheet:number,
   *          names?:{name:string, ref:string}[], tables?:{name:string, ref:string}[]}} 本
   */
  function 一覧(本) {
    var 出 = [];
    var s = (本 && 本.sheets) || [];
    var 章 = { 見出し: 'シート', 行: [] };
    for (var i = 0; i < s.length; i++) {
      章.行.push({ 名: s[i].name, 印: (i === 本.activeSheet) ? '今' : '', 種: 'シート', 先: i });
    }
    出.push(章);

    var n = (本 && 本.names) || [];
    var 章2 = { 見出し: '名前', 行: [] };
    for (var j = 0; j < n.length; j++) 章2.行.push({ 名: n[j].name, 印: n[j].ref || '', 種: '名前', 先: n[j].ref });
    出.push(章2);

    var t = (本 && 本.tables) || [];
    var 章3 = { 見出し: 'テーブル', 行: [] };
    for (var k = 0; k < t.length; k++) 章3.行.push({ 名: t[k].name, 印: t[k].ref || '', 種: 'テーブル', 先: t[k].ref });
    出.push(章3);
    return 出;
  }

  /** 一覧の 中身が いくつ 在るか（0なら「まだ 何も ありません」と 出す） */
  function 一覧の数(章たち) {
    var n = 0;
    for (var i = 0; i < (章たち || []).length; i++) n += 章たち[i].行.length;
    return n;
  }

  /**
   * ★セルにフォーカス★＝今の セルの 行と 列を 帯で 塗る
   * ★選んだ 範囲では なく「今の 1つの セル」を 見る★（Excelと 同じ）
   * @returns {{行:number, 列:number}|null}
   */
  function 帯(今) {
    if (!今 || typeof 今.r !== 'number' || typeof 今.c !== 'number') return null;
    if (今.r < 0 || 今.c < 0) return null;
    return { 行: 今.r, 列: 今.c };
  }

  return { 目盛り: 目盛り, 一覧: 一覧, 一覧の数: 一覧の数, 帯: 帯, CM: CM };
}));
