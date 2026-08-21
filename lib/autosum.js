/* autosum.js — ★Alt+= が当てる範囲を決める（純関数）★
 *
 *  ★実Excel 16.0 を 指示役が実際に押して測った真値（2026-08-21）★
 *    ⓐ A1〜A10 に数10個 → A11 で Alt+=  → =SUM(A1:A10)
 *    ⓑ C1〜C3 数・C4 空・C5 数 → C6 で   → =SUM(C5)      ★空きで止まる（C1〜C3は拾わない）★
 *    ⓒ F1〜F3 縦・C8〜E8 横 → F8 で      → =SUM(C8:E8)   ★横が勝つ★
 *
 *  ★この3通りから読み取れる決め方★
 *    ① ★すぐ上★ が数なら、上へ 数が続く間だけ たどる（空きで止まる）→ =SUM(上:下)
 *       ⓒで縦が採られなかったのは、F8 の ★すぐ上（F7）が空★ だから。
 *       F1〜F3 に数が在っても ★離れている物は拾わない★（ⓑの「空きで止まる」と同じ）
 *    ② 上が使えなければ ★すぐ左★ を同じように見る → =SUM(左:右)
 *    ③ どちらも無ければ ★分かりません★（何も入れない）
 *
 *  ★測れなかった事（勝手に決めない・報告に出す）★
 *    ・★ⓓ すぐ上 と すぐ左 の両方に数が在る時 どちらを採るか★
 *      ＝ 2026-08-21 指示役が ★3回 試して 3回とも測れなかった★
 *        （Excel が編集モードから戻らず 式が読めない）。★未測定のまま★。
 *      ⇒ ★決め打ちで黙って選ばない★。上の3通りと矛盾しない「上を先に見る」で出すが、
 *        ★どちらを取ったかを 画面に出して 直せるようにする★（ambiguous:true を返す）。
 *    ・Excel は式を入れた後 ★編集中のまま待つ★（Enterで確定）。
 *      うちは ★そのまま入れる★。ここは測った物ではない＝違いとして記録しておく。
 */
(function (root) {
  'use strict';

  /** その中身が「数」として足せるか。
   *  ★式のセルは 答え(d)を見る★（=SUM(...) の答えが数ならその列は数の並びとして扱う）。 */
  function isNum(cell) {
    if (!cell) return false;
    var v = (cell.d !== undefined && cell.d !== null && cell.d !== '') ? cell.d : cell.v;
    if (v === undefined || v === null || v === '') return false;
    if (typeof v === 'number') return isFinite(v);
    var s = String(v).trim();
    if (!s) return false;
    if (s.charAt(0) === '#') return false;                 // #REF! などは数ではない
    return !isNaN(Number(s.replace(/,/g, '')));
  }

  function colName(i) {
    var s = '', n = i;
    do { s = String.fromCharCode(65 + (n % 26)) + s; n = Math.floor(n / 26) - 1; } while (n >= 0);
    return s;
  }

  /**
   * @param {function(number,number):Object} get  (row,col) → セル（0起算）
   * @param {number} r  今いる行（0起算）
   * @param {number} c  今いる列（0起算）
   * @returns {{formula:string, dir:'up'|'left', from:string, to:string}|null}
   *          当てる所が見つからなければ null（★分かりません★）
   */
  function suggest(get, r, c) {
    /* ① 上：すぐ上が数なら、数が続く間だけ たどる（空きで止まる＝実測ⓑ） */
    /* ★ⓓ 上と左の両方が すぐ隣に在るか★（未測定なので 黙って選ばない・出して見せる） */
    var 上あり = (r > 0 && isNum(get(r - 1, c)));
    var 左あり = (c > 0 && isNum(get(r, c - 1)));
    var まよう = 上あり && 左あり;
    if (上あり) {
      var top = r - 1;
      while (top > 0 && isNum(get(top - 1, c))) top--;
      var a = colName(c) + (top + 1), b = colName(c) + r;
      return { formula: '=SUM(' + (a === b ? a : a + ':' + b) + ')', dir: 'up', from: a, to: b, ambiguous: まよう };
    }
    /* ② 左：同じ見方（実測ⓒ＝上が空なら左が採られる） */
    if (左あり) {
      var left = c - 1;
      while (left > 0 && isNum(get(r, left - 1))) left--;
      var a2 = colName(left) + (r + 1), b2 = colName(c - 1) + (r + 1);
      return { formula: '=SUM(' + (a2 === b2 ? a2 : a2 + ':' + b2) + ')', dir: 'left', from: a2, to: b2, ambiguous: まよう };
    }
    /* ③ ★分かりません★（勝手に広げない） */
    return null;
  }

  var api = { suggest: suggest, isNum: isNum, colName: colName };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.AutoSum = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
