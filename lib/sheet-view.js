/* sheet-view.js — ★シート ビュー／ユーザー設定のビュー（見え方に 名前を 付けて 覚える）★ 2026-08-30
 *
 *  ★真値（実Excel 16.0 で 実測）★ … tools/measure-sheetview.ps1
 *
 *  ● ユーザー設定のビュー（Workbook.CustomViews）★ブック 全体★
 *      ・`Add('うちの見え方', True, True)` … 2つ目=印刷設定も／3つ目=隠した行も
 *      ・`PrintSettings=True` / `RowColSettings=True`（実測）
 *      ・★覚えて 戻せる★＝3・4行目を 隠し Zoom=75 で 足した後、
 *        隠すのを やめ Zoom=100 に しても `Show()` で
 *        ★3行目 隠れている=True・Zoom=75★ に ★戻った★
 *
 *  ● シート ビュー（Worksheet.NamedSheetViews）★シート 1枚ごと★
 *      ・`Count` は はじめ 0／`Add('わたしの見え方')` で 1 に なる
 *      ・`Add` は ★足した 物を 返す★（`.Name` が 読める）
 *      ・`Exit()` で 抜けられる／`Active` は 何も 入っていない 時 ★空★
 *
 *  ★違う所（わざと）★
 *    実Excelの シート ビューは ★共同編集（何人かで 同時に 開く）★の 為の 物で、
 *    「自分の 絞り込みが 他の人に 見えない」事が 値打ち。
 *    ★うちは 1人で 使う★ので その 部分は ★無い★。
 *    ⇒ うちの シート ビューは ★名前を 付けた 絞り込み・並べ替えの 覚え書き★＝
 *      切り替えても ★元の 見え方が 壊れない★ 所は 同じに した。
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.SheetView = factory();
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  /** 名前を 決める（同じ 名前を 2つ 作らない。実Excel も 同じ 名前は 断る） */
  function 名前を決める(もうある, ひな) {
    var 使った = {};
    for (var i = 0; i < (もうある || []).length; i++) 使った[String(もうある[i])] = true;
    if (!使った[ひな]) return ひな;
    var n = 2;
    while (使った[ひな + ' ' + n]) n++;
    return ひな + ' ' + n;
  }

  /** 今の 見え方を 写し取る（★元の 物を 持たない＝後で 変わっても 覚えた物は 変わらない★） */
  function 写し取る(いま) {
    var o = いま || {};
    return {
      絞り: 束を写す(o.絞り),
      手で隠した行: 束を写す(o.手で隠した行),
      手で隠した列: 束を写す(o.手で隠した列),
      並べ方: o.並べ方 ? { 列: o.並べ方.列, 逆: !!o.並べ方.逆 } : null,
      倍率: (o.倍率 === undefined || o.倍率 === null) ? 1 : Number(o.倍率),
      固定行: Number(o.固定行 || 0),
      固定列: Number(o.固定列 || 0),
    };
  }
  function 束を写す(x) {
    var 出 = {};
    for (var k in (x || {})) if (Object.prototype.hasOwnProperty.call(x, k) && x[k]) 出[k] = true;
    return 出;
  }

  /** 足す（実測＝Add は 足した 物を 返す） */
  function 足す(たち, 名, いま) {
    var 名たち = (たち || []).map(function (v) { return v.名; });
    var 決めた = 名前を決める(名たち, String(名 || 'ビュー'));
    var v = { 名: 決めた, 見え方: 写し取る(いま) };
    たち.push(v);
    return v;
  }

  /** 名前で 探す */
  function 探す(たち, 名) {
    for (var i = 0; i < (たち || []).length; i++) if (たち[i].名 === 名) return たち[i];
    return null;
  }

  /** 消す（消せたら true） */
  function 消す(たち, 名) {
    for (var i = 0; i < (たち || []).length; i++) {
      if (たち[i].名 === 名) { たち.splice(i, 1); return true; }
    }
    return false;
  }

  /** 覚えた 見え方を 取り出す（★写しを 返す＝取り出した 後に いじっても 元は 壊れない★） */
  function 取り出す(たち, 名) {
    var v = 探す(たち, 名);
    return v ? 写し取る(v.見え方) : null;
  }

  return {
    名前を決める: 名前を決める, 写し取る: 写し取る,
    足す: 足す, 探す: 探す, 消す: 消す, 取り出す: 取り出す,
  };
}));
