/* menu-place.js — ★出したメニューを 画面の中に収める（純関数）★
 *
 *  ★なぜ在るか（2026-08-21 指示役の実測）★
 *    右クリックのメニューが ★高さ743px★／画面は ★619px★。
 *    前の書き方は「下がはみ出すなら y−高さ に置く」だけだったので
 *    ★上端が −470px＝画面の上へ470px はみ出し★、中で動かす事も出来なかった。
 *    ⇒ ★13項目が 客の手に届いていなかった★（並べ替え・絞り込み・固定・印刷・入力の決まり を含む）
 *    ★これは「DOMに在る≠読める」の4回目★。
 *
 *  ★決め事★
 *    ① 入り切らないなら ★中で動かせるようにする★（高さの上限＋スクロール）
 *    ② はみ出す時は ★出し直す★（下がだめなら上向き、それでもだめなら 端に貼る）
 *    ③ ★上端・左端が 画面の外に出る事は 絶対に無い★
 *
 *  ★測る所を 画面から切り離した理由★
 *    jsdom には 幅も高さも無いので 画面では測れない。
 *    ★計算だけを ここに出して、どの画面の高さでも 数え切れるようにする★。
 */
(function (root) {
  'use strict';

  /**
   * @param {Object} o
   *   x, y      … 押した所
   *   w, h      … メニューの 幅・高さ（中身そのままの大きさ）
   *   winW,winH … 画面の 幅・高さ
   *   余白      … 画面のふちから空ける分（既定8）
   * @returns {{left:number, top:number, maxHeight:number|null, 中で動かす:boolean, 見える高さ:number}}
   */
  function place(o) {
    var 余白 = (o.余白 === undefined || o.余白 === null) ? 8 : o.余白;
    var winW = o.winW, winH = o.winH;
    var 使える高さ = Math.max(0, winH - 余白 * 2);
    var 使える幅 = Math.max(0, winW - 余白 * 2);

    /* ① 入り切らないなら 中で動かせるようにする */
    var 中で動かす = o.h > 使える高さ;
    var maxHeight = 中で動かす ? 使える高さ : null;
    var 見える高さ = Math.min(o.h, 使える高さ);
    var 見える幅 = Math.min(o.w, 使える幅);

    /* ② 下に入るならそのまま。入らないなら 上向きに出し直す */
    var top = o.y;
    if (top + 見える高さ > winH - 余白) top = o.y - 見える高さ;
    /* ③ それでも入らないなら 端に貼る（★上端が画面の外に出ない★） */
    if (top < 余白) top = 余白;
    if (top + 見える高さ > winH - 余白) top = Math.max(余白, winH - 余白 - 見える高さ);

    var left = o.x;
    if (left + 見える幅 > winW - 余白) left = o.x - 見える幅;
    if (left < 余白) left = 余白;
    if (left + 見える幅 > winW - 余白) left = Math.max(余白, winW - 余白 - 見える幅);

    return { left: left, top: top, maxHeight: maxHeight, 中で動かす: 中で動かす, 見える高さ: 見える高さ };
  }

  /** ★全部の項目に 手が届くか★（届かないなら なぜ届かないかを返す）
      第2引数に 出した位置を渡すと それを見る（★判定その物を 壊せるようにするため★）。 */
  function 届くか(o, 出した位置) {
    var p = 出した位置 || place(o);
    if (p.top < 0) return { ok: false, why: '上端が画面の外（' + p.top + 'px）' };
    if (p.left < 0) return { ok: false, why: '左端が画面の外（' + p.left + 'px）' };
    if (p.top + p.見える高さ > o.winH) return { ok: false, why: '下端が画面の外' };
    if (p.left + Math.min(o.w, o.winW) > o.winW) return { ok: false, why: '右端が画面の外' };
    if (o.h > p.見える高さ && !p.中で動かす) return { ok: false, why: '入り切らないのに 中で動かせない' };
    return { ok: true, why: '' };
  }

  var api = { place: place, 届くか: 届くか };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.MenuPlace = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
