/* shiki-afure.js — ★1つの 式が 何マスにも 広がる（溢れ）★（2026-09-13）
 *
 *  ★★土台を 自分で 作る ⑤枚目★★
 *    土台は ①字に切る ②形にする ③頼りの地図 ④順番と計算 ★⑤溢れ★ ＋本体。
 *    ★司さん 2026-09-13「借り物が あったら 商用など いろいろ 引っかかる。全部 やれや」★
 *
 *  ★★答えは 全部 実Excel に 打って 測りました★★
 *    （2026-09-13・★16通り★／道具 `docs/measured/toru-afure.ps1`
 *      紙 `docs/measured/golden-afure-2026-09-13.tsv`／Excel 16.0 build 20326）
 *
 *  ★★当て推量なら 外して いた 物★★
 *    溢れる先に `=""`（★見た目は 空★）が 在るだけで → ★#SPILL!★
 *      ＝★目で 見ても 何も 無いのに 溢れられない★
 *    `=A1:A3+B1:B2`（形が 違う 四角どうし）→ ★11・22・#N/A★
 *      ＝★全体が 誤りに ならず、足りない 所だけ #N/A★
 *    `=@A1:A3` → ★1★（★@ を 付けると 溢れない＝1マスに なる★）
 *    `=IF(A1:A3>1,1,0)` → ★0・1・1★（★関数の 引数に 四角を 渡しても 溢れる★）
 *    `=A1:A3&"x"` → ★1x・2x・3x★（つなぎ も ひとつずつ）
 *
 *  ★★この 台が 見て いない 範囲★★
 *    ・溢れた 先が また 溢れる（重なり合い）の 細かい 決まり
 *    ・SORT/UNIQUE など ★大きさが 中身で 決まる★ 関数（まだ 作って いない）
 *    ・板を またぐ 四角
 *    ⇒★出来て いない 物を 出来た 顔で 混ぜない★
 *
 *  ★値の 形★ `shiki-keisan.js` と 同じ { 型, 値 }。
 *  ★溢れた 物の 形★ { 溢れ: true, 行数, 列数, 並び: [[値,…],…] }
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(
    typeof require === 'function' ? require('./shiki-keisan.js') : root.ShikiKeisan);
  else root.ShikiAfure = factory(root.ShikiKeisan);
})(typeof self !== 'undefined' ? self : this, function (計) {
  'use strict';

  function 溢れ(並び) {
    return { 溢れ: true, 行数: 並び.length, 列数: 並び.length ? 並び[0].length : 0, 並び: 並び };
  }
  function 溢れか(v) { return !!(v && v.溢れ === true); }

  /* ★1マスに して 読む★（溢れない 所へ 渡す時） */
  function 先頭(v) {
    if (!溢れか(v)) return v;
    if (!v.行数 || !v.列数) return 計.誤('#VALUE!');
    return v.並び[0][0];
  }

  /* ══ ★形が 違う 四角どうしを 重ねる★ ══
     ★実測＝`=A1:A3+B1:B2` → 11・22・★#N/A★★
       ＝全体を 誤りに せず ★足りない 所だけ #N/A★
     ★大きい 方に 合わせる★（1x1 は どこまでも 伸びる） */
  function 重ねる(左, 右, 作る) {
    var l = 溢れか(左), r = 溢れか(右);
    if (!l && !r) return 作る(左, 右);
    var 行数 = Math.max(l ? 左.行数 : 1, r ? 右.行数 : 1);
    var 列数 = Math.max(l ? 左.列数 : 1, r ? 右.列数 : 1);
    var 並び = [];
    for (var i = 0; i < 行数; i++) {
      var 段 = [];
      for (var j = 0; j < 列数; j++) {
        var a = 取る(左, i, j, l), b = 取る(右, i, j, r);
        段.push((a === null || b === null) ? 計.誤('#N/A') : 作る(a, b));
      }
      並び.push(段);
    }
    return 溢れ(並び);
  }
  function 取る(v, i, j, 溢れてるか) {
    if (!溢れてるか) return v;                       /* ★1つの 値は どこまでも 伸びる★ */
    var 行 = (v.行数 === 1) ? 0 : i;
    var 列 = (v.列数 === 1) ? 0 : j;
    if (行 >= v.行数 || 列 >= v.列数) return null;   /* ★足りない＝#N/A★ */
    return v.並び[行][列];
  }

  /* ★1つの 値に 同じ事を する（前置きの − など）★ */
  function ひとつずつ(v, 作る) {
    if (!溢れか(v)) return 作る(v);
    var 並び = [];
    for (var i = 0; i < v.行数; i++) {
      var 段 = [];
      for (var j = 0; j < v.列数; j++) 段.push(作る(v.並び[i][j]));
      並び.push(段);
    }
    return 溢れ(並び);
  }

  /* ══ ★溢れられるか を 見る★ ══
     ★実測＝溢れる先に 物が 在れば #SPILL!★
       数でも 字でも ★`=""`（見た目は 空）でも 駄目★
     ⇒★「空マス」だけが 溢れてよい★（空の 字は 空マスでは ない） */
  function 溢れられるか(表, 元名, v, 番地を出す) {
    if (!溢れか(v)) return { よい: true };
    var 塞ぐ = [];
    for (var i = 0; i < v.行数; i++) {
      for (var j = 0; j < v.列数; j++) {
        if (i === 0 && j === 0) continue;          /* 自分の マスは 除く */
        var 名 = 番地を出す(i, j);
        var m = 表[名];
        if (!m) continue;                          /* 何も 無い＝溢れてよい */
        if (m.打った字 === '' || m.打った字 == null) continue;  /* ★本当に 空★ */
        塞ぐ.push(名);                             /* ★数・字・式・=""  全部 駄目★ */
      }
    }
    return 塞ぐ.length ? { よい: false, 塞ぐ: 塞ぐ } : { よい: true };
  }

  return {
    溢れ: 溢れ, 溢れか: 溢れか, 先頭: 先頭,
    重ねる: 重ねる, ひとつずつ: ひとつずつ,
    溢れられるか: 溢れられるか
  };
});
