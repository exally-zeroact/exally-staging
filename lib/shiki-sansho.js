/* shiki-sansho.js — ★参照（A1・$A$1・A1:B2・板!A1）を 読む★（2026-09-11）
 *
 *  ★★土台を 自分で 作る ③枚目の 前半★★
 *    土台は 5つ … ①字に 切る ②形に する ③★頼りの 地図★ ④順番と 計算 ⑤溢れ
 *    地図を 作るには ★まず 参照が 読めないと 始まりません★
 *
 *  ★★借り物の 中は 1文字も 読んで いません★★
 *    答えは ★実Excel に 聞いて 測りました★（2026-09-11）
 *      A1 → 行1 列1 ／ Z9 → 行9 列26 ／ AA1 → 列27 ／ XFD1 → 列16384
 *      A1048576 → 行1048576 ／ $A$1 $A1 A$1 … ★どれも A1 と 同じ 場所★
 *      A1:B2 → 行1..2 列1..2 ／ A:A → 列1・行 1048576本 ／ 1:1 → 行1・列 16384本
 *
 *  ★返す 物★
 *    {ok, 板, 上, 左, 下, 右, 列全体, 行全体}   … 行・列は ★0から 数える★（画面と 同じ）
 *    読めなければ {ok:false}
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.ShikiSansho = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var 行の数 = 1048576;   /* ★実Excel に 聞いた★ */
  var 列の数 = 16384;     /* ★実Excel に 聞いた（XFD = 16384）★ */

  /** ★A → 0／Z → 25／AA → 26★ */
  function 列を数に(字) {
    var n = 0;
    for (var i = 0; i < 字.length; i++) {
      var c = 字.charCodeAt(i);
      if (c >= 97) c -= 32;                       /* 小文字も 受ける */
      if (c < 65 || c > 90) return -1;
      n = n * 26 + (c - 64);
    }
    return n - 1;
  }

  /** ★0 → A／25 → Z／26 → AA★ */
  function 数を列に(n) {
    var s = '';
    n = n + 1;
    while (n > 0) { var t = (n - 1) % 26; s = String.fromCharCode(65 + t) + s; n = (n - t - 1) / 26; }
    return s;
  }

  /** ★1つの マス（A1 / $A$1）★ … 返す {行, 列, 行固定, 列固定} */
  function 一つ(字) {
    var i = 0, 列固定 = false, 行固定 = false;
    if (字.charAt(i) === '$') { 列固定 = true; i++; }
    var j = i;
    while (j < 字.length && /[A-Za-z]/.test(字.charAt(j))) j++;
    if (j === i) return null;
    var 列 = 列を数に(字.slice(i, j));
    if (列 < 0 || 列 >= 列の数) return null;
    if (字.charAt(j) === '$') { 行固定 = true; j++; }
    var k = j;
    while (k < 字.length && /[0-9]/.test(字.charAt(k))) k++;
    if (k === j || k !== 字.length) return null;
    var 行 = Number(字.slice(j, k)) - 1;
    if (!(行 >= 0) || 行 >= 行の数) return null;
    return { 行: 行, 列: 列, 行固定: 行固定, 列固定: 列固定 };
  }

  /** ★列だけ（A / $A）★ */
  function 列だけ(字) {
    var i = 0;
    if (字.charAt(i) === '$') i++;
    var j = i;
    while (j < 字.length && /[A-Za-z]/.test(字.charAt(j))) j++;
    if (j === i || j !== 字.length) return null;
    var 列 = 列を数に(字.slice(i, j));
    return (列 >= 0 && 列 < 列の数) ? 列 : null;
  }

  /** ★行だけ（1 / $1）★ */
  function 行だけ(字) {
    var i = 0;
    if (字.charAt(i) === '$') i++;
    if (i >= 字.length) return null;
    for (var j = i; j < 字.length; j++) if (!/[0-9]/.test(字.charAt(j))) return null;
    var 行 = Number(字.slice(i)) - 1;
    return (行 >= 0 && 行 < 行の数) ? 行 : null;
  }

  /**
   * ★参照を 読む★
   * @param {string} 字 … `A1` `$A$1` `A1:B2` `Sheet1!A1` `'あ い'!A1:B2` `A:A` `1:1`
   */
  function 読む(字) {
    var s = String(字 == null ? '' : 字);
    var 板 = null;

    /* ★板の 名前★＝`!` の 前（'…' で 囲む 事も 在る） */
    var 感 = -1;
    if (s.charAt(0) === "'") {
      var e = 1;
      while (e < s.length) {
        if (s.charAt(e) === "'") { if (s.charAt(e + 1) === "'") { e += 2; continue; } break; }
        e++;
      }
      if (s.charAt(e) !== "'" || s.charAt(e + 1) !== '!') return { ok: false };
      板 = s.slice(1, e).split("''").join("'");
      s = s.slice(e + 2);
    } else {
      感 = s.indexOf('!');
      if (感 >= 0) { 板 = s.slice(0, 感); s = s.slice(感 + 1); }
    }
    if (板 !== null && !板.length) return { ok: false };

    var 二 = s.split(':');
    if (二.length === 1) {
      var a = 一つ(二[0]);
      if (!a) return { ok: false };
      return { ok: true, 板: 板, 上: a.行, 左: a.列, 下: a.行, 右: a.列, 列全体: false, 行全体: false };
    }
    if (二.length !== 2) return { ok: false };

    /* ★A1:B2★ */
    var p = 一つ(二[0]), q = 一つ(二[1]);
    if (p && q) {
      return { ok: true, 板: 板,
        上: Math.min(p.行, q.行), 下: Math.max(p.行, q.行),
        左: Math.min(p.列, q.列), 右: Math.max(p.列, q.列),
        列全体: false, 行全体: false };
    }
    /* ★A:A（列 まるごと）★ */
    var c1 = 列だけ(二[0]), c2 = 列だけ(二[1]);
    if (c1 !== null && c2 !== null) {
      return { ok: true, 板: 板, 上: 0, 下: 行の数 - 1,
        左: Math.min(c1, c2), 右: Math.max(c1, c2), 列全体: true, 行全体: false };
    }
    /* ★1:1（行 まるごと）★ */
    var r1 = 行だけ(二[0]), r2 = 行だけ(二[1]);
    if (r1 !== null && r2 !== null) {
      return { ok: true, 板: 板, 左: 0, 右: 列の数 - 1,
        上: Math.min(r1, r2), 下: Math.max(r1, r2), 列全体: false, 行全体: true };
    }
    return { ok: false };
  }

  /** ★何マス 分か★（地図の 大きさを 見る為） */
  function マスの数(r) {
    if (!r || !r.ok) return 0;
    return (r.下 - r.上 + 1) * (r.右 - r.左 + 1);
  }

  return { 読む: 読む, マスの数: マスの数, 列を数に: 列を数に, 数を列に: 数を列に,
    行の数: 行の数, 列の数: 列の数 };
});
