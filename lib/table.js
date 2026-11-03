/* table.js — ★テーブル（Ctrl+T）★ 2026-08-30
 *
 *  ★真値（実Excel 16.0 で 実測）★ … tools/measure-table.ps1 / tools/measure-table2.ps1
 *
 *    A1:C4（1行目が 見出し）で 作った ListObject:
 *      名前            = ★テーブル1★（2つ目は テーブル2）
 *      範囲            = A1:C4（そのまま）
 *      見出し行        = ★True★
 *      集計行          = ★False★
 *      自動フィルタ    = ★True★
 *      しま（行）      = ★True★  ／ しま（列）= False
 *      最初の列 / 最後の列 = False / False
 *      見た目の名      = TableStyleMedium2
 *      列の名          = 月 / 売上 / 原価
 *
 *    ★1行目が 見出しか どうかの 見立て（xlGuess で 実測）★
 *      ①1行目が 字   → 見出しとして 使う（範囲は そのまま・列の名は その字）
 *      ②1行目も 数   → ★見出し行を 1行 足す★（D1:E2 → ★D1:E3★）・列の名は ★列1・列2★
 *
 *    ★集計行を 出した時★（実測）
 *      ・行が 1つ 増える（A1:C4 → A1:C5）
 *      ・一番 左は ★'集計'（ただの字）★
 *      ・★一番 右の列だけ★ =SUBTOTAL(109,[列名])
 *      ・間の 列は ★空★
 *      （2列の テーブルでも 同じ＝'集計' と 右端の SUBTOTAL だけ）
 *
 *    ★SUBTOTAL の 109★（実測）
 *      2行目を 隠して … SUBTOTAL(109,H1:H3)=4 ／ SUBTOTAL(9,H1:H3)=6
 *      ＝109 は ★隠した行を 数えない★
 *
 *  ★うちと 違う所（黙らない）★
 *    実Excel は 集計の 式を ★[列名]★ で 書く（構造化参照）。
 *    うちの 計算は まだ [列名] を 読めないので ★A1の 形★で 書く。
 *    ＝出る数は 同じ。書き出した時の 式の 見た目が 違う。
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.Table = factory();
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  /* うちの 色（Excelの TableStyleMedium2 は 写さない） */
  var 色 = {
    見出し: '#3D9E72',
    見出しの字: '#ffffff',
    しま: '#F0FAF4',       /* 1つ おきの 薄い 帯 */
    枠: '#C8ECD8',
  };

  function 数か(v) {
    if (v === null || v === undefined || v === '') return false;
    if (typeof v === 'number') return isFinite(v);
    var s = String(v).replace(/[,¥￥\s]/g, '');
    if (s === '') return false;
    return isFinite(Number(s));
  }

  /** ★名前を 決める★＝テーブル1・テーブル2 …（実測どおり／空いている 一番小さい 番号） */
  function 名前を決める(もうある) {
    var 使った = {};
    for (var i = 0; i < (もうある || []).length; i++) 使った[String(もうある[i])] = true;
    for (var n = 1; n < 100000; n++) {
      var 名 = 'テーブル' + n;
      if (!使った[名]) return 名;
    }
    return 'テーブル';
  }

  /** ★1行目を 見出しに するか★（実測の 見立て）
   *   1行目に 字が 1つでも 在り、数が 1つも 無ければ 見出し。 */
  function 見出しか(一行目) {
    var 字あり = false;
    for (var i = 0; i < (一行目 || []).length; i++) {
      var v = 一行目[i];
      if (v === null || v === undefined || v === '') continue;
      if (数か(v)) return false;          /* ★数が 混ざっていたら 見出しでは ない★ */
      字あり = true;
    }
    return 字あり;
  }

  /** 列の 名前（見出しが 無い時は ★列1・列2…★＝実測②） */
  function 列名を作る(一行目, 幅, 見出しがある) {
    var 出 = [];
    for (var i = 0; i < 幅; i++) {
      if (見出しがある) {
        var v = (一行目 || [])[i];
        出.push((v === null || v === undefined || v === '') ? ('列' + (i + 1)) : String(v));
      } else {
        出.push('列' + (i + 1));
      }
    }
    /* ★同じ名前は 使えない★（Excelも 直す）＝後ろに 2,3… を 付ける */
    var 使った = {};
    for (var j = 0; j < 出.length; j++) {
      var 名 = 出[j], n = 2;
      while (使った[名]) { 名 = 出[j] + n; n++; }
      使った[名] = true;
      出[j] = 名;
    }
    return 出;
  }

  /** ★しまに するか★＝見出しの 次の行を 1本目として 1つ おき（実測 ShowTableStyleRowStripes=True）
   *  @param 何行目 … 中身の 何行目か（0はじまり） */
  function しまか(何行目) { return (何行目 % 2) === 1; }

  /** 集計の 式（★一番 右の列だけ★＝実測） */
  function 集計の式(範囲の字) { return '=SUBTOTAL(109,' + 範囲の字 + ')'; }

  /** 集計行に 何を 入れるか（実測＝左端は '集計'・右端だけ 式・間は 空） */
  function 集計行の中身(列の数, 右端の範囲) {
    var 出 = [];
    for (var i = 0; i < 列の数; i++) {
      if (i === 0) 出.push('集計');
      else if (i === 列の数 - 1) 出.push(集計の式(右端の範囲));
      else 出.push('');
    }
    /* ★1列だけの テーブル★＝左端 かつ 右端。実Excel は '集計' を 出す（左が 勝つ） */
    return 出;
  }

  /** その セルが テーブルの どこか（'見出し' / '中身' / '集計' / null） */
  function どこか(表, r, c) {
    if (!表) return null;
    if (r < 表.r1 || r > 表.r2 || c < 表.c1 || c > 表.c2) return null;
    if (表.見出し行 && r === 表.r1) return '見出し';
    if (表.集計行 && r === 表.r2) return '集計';
    return '中身';
  }

  /** 重なっている テーブルが 在るか（★2つの テーブルを 重ねない★） */
  function 重なる(表たち, r1, c1, r2, c2) {
    for (var i = 0; i < (表たち || []).length; i++) {
      var t = 表たち[i];
      if (r1 <= t.r2 && r2 >= t.r1 && c1 <= t.c2 && c2 >= t.c1) return t;
    }
    return null;
  }

  return {
    名前を決める: 名前を決める, 見出しか: 見出しか, 列名を作る: 列名を作る,
    しまか: しまか, 集計の式: 集計の式, 集計行の中身: 集計行の中身,
    どこか: どこか, 重なる: 重なる, 色: 色, _数か: 数か,
  };
}));
