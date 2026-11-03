/* timeline.js — ★タイムライン（挿入→フィルター→タイムライン）★ 2026-08-30
 *
 *  ★実Excel で 測れた 事★ … tools/measure-insert.ps1
 *    ・タイムラインは ★日付の 列を 持つ ピボットが 要る★
 *    ・★スライサーの 仲間★（SlicerCaches に 入る）
 *    ・★形・大きさは COM から 読めなかった★（スライサーと 同じ）
 *      ⇒ ★見た目は うちの 決め★。★絞る 働きは 同じ★に する。
 *
 *  ★うちの 決め★
 *    ・区切りは ★年・四半期・月・日★の 4つ（実Excel と 同じ 4つ）
 *    ・押した 所から 押した 所まで（★間を まとめて 選べる★）
 *    ・絞る 道は ★見出しの ▼ と 同じ 1本★（同じ物を 2つ 持たない）
 *
 *  ★日付の 読み方★
 *    `2026-08-30` `2026/8/30` `2026年8月30日` と Excel の 通し番号を 読む。
 *    ★読めない 物は 黙って 捨てない★＝「読めない」として 数える。
 *
 *  ★Excel の 通し番号（実測 2026-08-30）★
 *    1→1900-01-01／2→1900-01-02／59→1900-02-28／
 *    ★60→1900-02-29（実際には 無い 日）★／61→1900-03-01／
 *    45900→2025-08-31／46000→2025-12-09
 *    ＝★Excel は 1900年を うるう年だと 思っている★（昔からの 決まり）。
 *      うちは 59以下と 61以上を 別々に 数え、★60 は「読めない」と 言う★。
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.Timeline = factory();
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var 区切りたち = [
    { 種: '年', 名: '年' },
    { 種: '四半期', 名: '四半期' },
    { 種: '月', 名: '月' },
    { 種: '日', 名: '日' },
  ];

  /** その 書式は 日付の 書式か（y/m/d が 入っていて 通貨や % では ない） */
  function 日付の書式か(書式) {
    var f = String(書式 === null || 書式 === undefined ? '' : 書式);
    if (!f) return false;
    if (/[%¥$]/.test(f)) return false;
    return /y|m|d|e|g|年|月|日/.test(f);
  }

  /** 値を 日付に する（読めなければ null）
   *  @param 書式 その セルの 書式（★ただの 数は 書式が 日付の 時だけ 通し番号と 見る★）
   *
   *  ★08-30 実ブラウザで 見つけた★＝書式を 見ないと
   *    ★金額の 列（100・200…）を 日付の 列だと 思い込む★（100 → 1900-04-09）。
   *    ⇒ ★ただの 数は 書式が 日付の 時だけ★ 通し番号として 読む。
   */
  function 日付にする(v, 書式) {
    if (v === null || v === undefined || v === '') return null;
    if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
    var s = String(v).trim();
    /* ★Excel の 通し番号★（実測 2026-08-30）
     *   1 → 1900-01-01 ／ 2 → 1900-01-02 ／ 59 → 1900-02-28
     *   ★60 → 1900-02-29（実際には 無い 日）★
     *   61 → 1900-03-01 ／ 45900 → 2025-08-31 ／ 46000 → 2025-12-09
     *   ＝★1900年を うるう年だと 思っている★ぶん、60 から 先が 1日 ずれる。
     *   ⇒ 59以下は 1899-12-31 から／61以上は 1899-12-30 から 数える。
     *     ★60 は 実際に 無い 日★なので ★読めないと 言う★（黙って 別の 日に しない）。 */
    if (/^\d+(\.\d+)?$/.test(s)) {
      if (!日付の書式か(書式)) return null;      /* ★ただの 数は 日付に しない★ */
      var n = Math.floor(Number(s));
      if (n === 60) return null;                      /* ★1900-02-29 は 無い★ */
      if (n >= 1 && n <= 59) {
        return new Date(Date.UTC(1899, 11, 31) + n * 86400000);
      }
      if (n >= 61 && n < 2958466) {
        return new Date(Date.UTC(1899, 11, 30) + n * 86400000);
      }
      return null;
    }
    var m = /^(\d{4})[-/年](\d{1,2})[-/月](\d{1,2})日?$/.exec(s);
    if (m) {
      var d = new Date(Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3])));
      return isNaN(d.getTime()) ? null : d;
    }
    var m2 = /^(\d{4})[-/年](\d{1,2})月?$/.exec(s);
    if (m2) {
      var d2 = new Date(Date.UTC(Number(m2[1]), Number(m2[2]) - 1, 1));
      return isNaN(d2.getTime()) ? null : d2;
    }
    return null;
  }

  /** その 日付の 区切りの 名前（年/四半期/月/日） */
  function 区切りの名(日, 種) {
    var y = 日.getUTCFullYear(), m = 日.getUTCMonth() + 1, d = 日.getUTCDate();
    if (種 === '年') return String(y) + '年';
    if (種 === '四半期') return String(y) + '年 Q' + (Math.floor((m - 1) / 3) + 1);
    if (種 === '月') return String(y) + '年' + m + '月';
    return String(y) + '年' + m + '月' + d + '日';
  }

  /** 並べる為の 順（同じ 区切りなら 同じ 数） */
  function 区切りの順(日, 種) {
    var y = 日.getUTCFullYear(), m = 日.getUTCMonth() + 1, d = 日.getUTCDate();
    if (種 === '年') return y * 10000;
    if (種 === '四半期') return y * 10000 + (Math.floor((m - 1) / 3) + 1) * 100;
    if (種 === '月') return y * 10000 + m * 100;
    return y * 10000 + m * 100 + d;
  }

  /** ★目盛りを 作る★
   *  @param 値たち 日付の 列の 値（{値, 行} でも 値だけでも 良い）
   *  @returns {目盛り:[{名,順,行たち}], 読めない, 読めた}
   */
  function 目盛りを作る(値たち, 種) {
    var 束 = {};
    var 読めない = 0, 読めた = 0;
    for (var i = 0; i < (値たち || []).length; i++) {
      var x = 値たち[i];
      var v = (x && typeof x === 'object' && '値' in x) ? x.値 : x;
      var 行 = (x && typeof x === 'object' && '行' in x) ? x.行 : i;
      var 書 = (x && typeof x === 'object') ? x.書式 : undefined;
      var 日 = 日付にする(v, 書);
      if (!日) { 読めない++; continue; }
      読めた++;
      var 名 = 区切りの名(日, 種);
      if (!束[名]) 束[名] = { 名: 名, 順: 区切りの順(日, 種), 行たち: [] };
      束[名].行たち.push(行);
    }
    var 出 = Object.keys(束).map(function (k) { return 束[k]; });
    出.sort(function (a, b) { return a.順 - b.順; });
    return { 目盛り: 出, 読めない: 読めない, 読めた: 読めた };
  }

  /** ★選んだ 幅（始め〜終わり）に 入る 行★を 出す */
  function 選んだ行(目盛り, 始め, 終わり) {
    if (始め === null || 始め === undefined) return null;   /* null＝ぜんぶ 見せる */
    var a = Math.min(始め, 終わり === undefined || 終わり === null ? 始め : 終わり);
    var b = Math.max(始め, 終わり === undefined || 終わり === null ? 始め : 終わり);
    var 出 = {};
    for (var i = a; i <= b && i < 目盛り.length; i++) {
      if (i < 0) continue;
      for (var j = 0; j < 目盛り[i].行たち.length; j++) 出[目盛り[i].行たち[j]] = true;
    }
    return 出;
  }

  return {
    区切りたち: 区切りたち, 日付にする: 日付にする, 日付の書式か: 日付の書式か,
    区切りの名: 区切りの名, 区切りの順: 区切りの順,
    目盛りを作る: 目盛りを作る, 選んだ行: 選んだ行,
  };
}));
