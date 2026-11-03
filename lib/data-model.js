/* data-model.js — ★データ モデル（表どうしを つなげる）★ 2026-08-30
 *
 *  ★真値（実Excel 16.0 で 実測）★ … tools/measure-xml.ps1 と 追加の 実測
 *    ・表（ListObject）は 名前を 付けられる／`ListColumns` で 列の 名が 取れる
 *    ・★表の 既定の 名前は 「テーブル3」★（日本語＋番号）
 *    ・`Workbook.Model.ModelTables.Count` … ★はじめ 0★
 *      （＝★表を 作っただけでは モデルに 入らない★）
 *    ・`Model.ModelRelationships.Count` … ★0★
 *    ・`Model.AddConnection` … ★COM から 呼べなかった★（引数の 形が 合わない）
 *      ⇒ ★中の 動きは 測れていない★ので ★うちで 決めて うちで 測る★
 *
 *  ★うちの 決め★
 *    ・つながり ＝ ★{ 左の表, 左の列, 右の表, 右の列 }★ の 1本
 *    ・★片方が「1つずつ」で ないと つなげない★（同じ 値が 2回 出る 側は 親に 出来ない）
 *      ＝実Excel も 同じ 決まり（「一意」でないと 親に 出来ない）
 *    ・つなげた あとは ★左の 表に 右の 列を 足して 見せる★（VLOOKUP の 束）
 *
 *  ★お金の 話★
 *    つなげる のは ★値を 持ってくる★だけ。★合計や 割り算は うちの 計算で する★
 *    （AIに 金額を 覚えさせない・`lib/chizu.js` の 決まりと 同じ）
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.DataModel = factory();
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  function 字(v) { return (v === null || v === undefined) ? '' : String(v); }

  /** その 列の 値が ★1つずつ（同じ 値が 2回 出ない）★か */
  function 一つずつか(表, 列) {
    var 見た = {};
    for (var r = 1; r < 表.行たち.length; r++) {          /* 0行目は 見出し */
      var v = 字(表.行たち[r][列]);
      if (v === '') continue;
      if (見た[v]) return false;
      見た[v] = true;
    }
    return true;
  }

  /** 列の 名前から 番号（無ければ -1） */
  function 列の番号(表, 名) {
    var 見出し = 表.行たち[0] || [];
    for (var i = 0; i < 見出し.length; i++) if (字(見出し[i]) === 字(名)) return i;
    return -1;
  }

  /** ★つなげられるか★を 見る（理由も 返す） */
  function つなげるか(左表, 左列名, 右表, 右列名) {
    var lc = 列の番号(左表, 左列名), rc = 列の番号(右表, 右列名);
    if (lc < 0) return { 出来る: false, 訳: '左の 表に「' + 左列名 + '」が ありません' };
    if (rc < 0) return { 出来る: false, 訳: '右の 表に「' + 右列名 + '」が ありません' };
    if (!一つずつか(右表, rc)) {
      return { 出来る: false,
        訳: '右の「' + 右列名 + '」は 同じ 値が 2回 以上 出ます'
          + '（つなぐ 相手は 1つずつで ないと いけません）' };
    }
    /* 合う 値が 1つも 無ければ 教える（黙って 空を 返さない） */
    var 右の値 = {};
    for (var r = 1; r < 右表.行たち.length; r++) 右の値[字(右表.行たち[r][rc])] = true;
    var 合う = 0;
    for (var l = 1; l < 左表.行たち.length; l++) if (右の値[字(左表.行たち[l][lc])]) 合う++;
    if (!合う) return { 出来る: false, 訳: '合う 値が 1つも ありません' };
    return { 出来る: true, 訳: '合う 行 ' + 合う + '個', 左列: lc, 右列: rc, 合う: 合う };
  }

  /** つながりを 足す（同じ 物は 2つ 作らない） */
  function つなぎを足す(たち, 左表名, 左列名, 右表名, 右列名) {
    for (var i = 0; i < たち.length; i++) {
      var v = たち[i];
      if (v.左表 === 左表名 && v.左列 === 左列名 && v.右表 === 右表名 && v.右列 === 右列名) return v;
    }
    var 新 = { 左表: 左表名, 左列: 左列名, 右表: 右表名, 右列: 右列名 };
    たち.push(新);
    return 新;
  }
  function つなぎを消す(たち, i) {
    if (i < 0 || i >= たち.length) return false;
    たち.splice(i, 1);
    return true;
  }

  /** ★つなげて 1つの 表に する★
   *  @returns [[見出し…],[値…],…]（左の 表 ＋ 右の 表の 列。★合わない 行も 残す★）
   */
  function つなげる(左表, 左列名, 右表, 右列名) {
    var 見 = つなげるか(左表, 左列名, 右表, 右列名);
    if (!見.出来る) return null;
    var lc = 見.左列, rc = 見.右列;
    var 右見出し = 右表.行たち[0] || [];
    var 右引き = {};
    for (var r = 1; r < 右表.行たち.length; r++) {
      右引き[字(右表.行たち[r][rc])] = 右表.行たち[r];
    }
    var 出 = [];
    var 頭 = (左表.行たち[0] || []).slice();
    for (var c = 0; c < 右見出し.length; c++) {
      if (c === rc) continue;                                  /* 同じ 列は 2回 出さない */
      頭.push(字(右表.名) + '.' + 字(右見出し[c]));
    }
    出.push(頭);
    for (var l = 1; l < 左表.行たち.length; l++) {
      var 行 = 左表.行たち[l].slice();
      var 相 = 右引き[字(左表.行たち[l][lc])];
      for (var c2 = 0; c2 < 右見出し.length; c2++) {
        if (c2 === rc) continue;
        行.push(相 ? 相[c2] : '');                             /* ★合わない 行も 残す★ */
      }
      出.push(行);
    }
    return 出;
  }

  return {
    一つずつか: 一つずつか, 列の番号: 列の番号, つなげるか: つなげるか,
    つなぎを足す: つなぎを足す, つなぎを消す: つなぎを消す, つなげる: つなげる,
  };
}));
