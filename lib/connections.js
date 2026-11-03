/* connections.js — ★クエリと接続（どこから 持ってきたか の 台帳）★ 2026-08-30
 *
 *  ★真値（実Excel 16.0 で 実測）★ … tools/measure-data3.ps1
 *    ・`Workbook.Connections.Count` … はじめ ★0★
 *    ・`Workbook.Queries.Count` … はじめ ★0★
 *    ・`RefreshAll()` … つなぎが 0でも ★通る（何も 起きない）★
 *    ・外の ブックへの つなぎ（`LinkSources(1)`）
 *        - はじめ ★何も 無い（null）★
 *        - `='C:\ない\[ないブック.xlsx]Sheet1'!A1` と 書くと ★1件★に なり
 *          中身は ★`C:\ない\ないブック.xlsx`★（ブックまで。シート名は 入らない）
 *        - その セルの 答えは ★`#REF!`★（★消えていない のに 読めない★）
 *
 *  ★違う所（わざと・出来ない事は 出来ないと 書く）★
 *    ・実Excel の つなぎは ★ファイルの 場所を 覚えて 後から 読み直す★。
 *      ブラウザは ★人が 選んだ ファイルしか 読めない★（場所を 覚えて 勝手に 開けない）。
 *      ⇒ うちの 台帳は ★どこから 来たか★を 覚え、更新は ★もう一度 選んでもらう★。
 *        これを ★画面に そう 書く★。
 *    ・「画像から」は ★字を 読み取る 仕組みが 要る★ので まだ。
 *    ・「株式・通貨・地理」は ★Microsoft の サービスに つながる★物なので うちには 無い。
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.Connections = factory();
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  /* つなぎの 種類（★人が 選び直せる か どうか★を 持つ） */
  var 種類たち = [
    { 種: 'csv',   名: 'CSV / テキスト', 選び直せる: true,  説明: 'パソコンの ファイルから' },
    { 種: 'range', 名: 'この ブックの 表', 選び直せる: false, 説明: '同じ ブックの 中の 範囲から' },
    { 種: 'web',   名: 'Web から',        選び直せる: true,  説明: 'ページの 表から' },
    { 種: 'excel',名: '外の ブック',      選び直せる: true,  説明: '別の Excel ファイルから' },
  ];

  /** 名前を 決める（同じ 名前を 2つ 作らない） */
  function 名前を決める(たち, ひな) {
    var 使った = {};
    for (var i = 0; i < (たち || []).length; i++) 使った[String(たち[i].名)] = true;
    if (!使った[ひな]) return ひな;
    var n = 2;
    while (使った[ひな + ' ' + n]) n++;
    return ひな + ' ' + n;
  }

  /** つなぎを 足す */
  function 足す(たち, 種, 名, 元, 中身) {
    var v = {
      種: 種, 名: 名前を決める(たち, String(名 || 'つなぎ')),
      元: String(元 || ''),                 /* どこから 来たか（ファイル名・URL・範囲） */
      行数: (中身 && 中身.length) ? 中身.length : 0,
      列数: (中身 && 中身[0]) ? 中身[0].length : 0,
      更新回数: 0,
      置き場: null,                          /* {シート, r, c}＝出した 所 */
    };
    たち.push(v);
    return v;
  }

  function 探す(たち, 名) {
    for (var i = 0; i < (たち || []).length; i++) if (たち[i].名 === 名) return たち[i];
    return null;
  }
  function 消す(たち, 名) {
    for (var i = 0; i < (たち || []).length; i++) {
      if (たち[i].名 === 名) { たち.splice(i, 1); return true; }
    }
    return false;
  }

  /** その 種類は 押すだけで 更新 出来るか（★出来ない物は そう 言う★） */
  function 更新できるか(種) {
    for (var i = 0; i < 種類たち.length; i++) {
      if (種類たち[i].種 === 種) return !種類たち[i].選び直せる || 種 === 'web';
    }
    return false;
  }

  /** 更新の 説明（画面に 出す 字） */
  function 更新の説明(種) {
    if (種 === 'range') return 'この ブックの 中なので そのまま 数え直せます';
    if (種 === 'web') return 'もう一度 読みに 行きます';
    return 'ブラウザは 前の ファイルを 勝手に 開けません＝もう一度 選んでください';
  }

  /** 外の ブックへの つなぎを 式から 見つける
   *  実測＝`='C:\\ない\\[ないブック.xlsx]Sheet1'!A1` → ★`C:\\ない\\ないブック.xlsx`★
   */
  function 外のブックを探す(式たち) {
    var 出 = [];
    var 見た = {};
    for (var i = 0; i < (式たち || []).length; i++) {
      var f = String(式たち[i] || '');
      var re = /'?([^'\[\]]*)\[([^\]]+)\]/g;
      var m;
      while ((m = re.exec(f)) !== null) {
        var 道 = (m[1] || '') + m[2];
        if (見た[道]) continue;
        見た[道] = true;
        出.push(道);
      }
    }
    return 出;
  }

  return {
    種類たち: 種類たち, 名前を決める: 名前を決める, 足す: 足す, 探す: 探す, 消す: 消す,
    更新できるか: 更新できるか, 更新の説明: 更新の説明, 外のブックを探す: 外のブックを探す,
  };
}));
