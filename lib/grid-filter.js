/* grid-filter.js — ★絞り込み（フィルター）の判定（純関数）★
 *
 *  ★実Excel 16.0.20228 を COM で動かして測った真値（2026-08-21）★
 *    ① 「りんご」で絞ると ★行3と行5が Hidden=True★。★中身は消えない・隠れるだけ★
 *       ★見出し行（1行目）は隠れない★
 *    ② ★合計は 絞っても変わらない★
 *         絞る前   SUM=15 ／ SUBTOTAL(9)=15 ／ SUBTOTAL(109)=15
 *         絞った後 ★SUM=15（そのまま）★ ／ SUBTOTAL(9)=12 ／ SUBTOTAL(109)=12
 *       ＝ ★隠れた行も SUM は足す★（合計が黙って小さくならない）
 *    ③ ★フィルターで隠れた行と 手で隠した行は 別物★
 *         手で行を隠した時 SUBTOTAL(9)=15（変わらない）／SUBTOTAL(109)=14
 *       ＝ SUBTOTAL(9) は ★フィルターだけ★ 効く。だから2つを混ぜて持ってはいけない。
 *    ④ 解除すると ★全部の行が戻る★（1:False 〜 6:False）
 *
 *  ★うちの決まり★
 *    ・絞り込みで隠す行は sheet.filterHidden（★sheet.hiddenRows とは別★）に持つ。
 *      混ぜると「解除したら 人が手で隠した行まで出てくる」事になる。
 *    ・★合計は触らない★（隠すのは見た目だけ。計算からは外さない＝実測どおり）
 */
(function (root) {
  'use strict';

  /** セルの「見えている字」（比べるのに使う） */
  function text(cell) {
    if (!cell) return '';
    var v = (cell.d !== undefined && cell.d !== null && cell.d !== '') ? cell.d : cell.v;
    return (v === undefined || v === null) ? '' : String(v).trim();
  }

  /**
   * 「選んだセルの値」で絞った時に 隠す行を返す
   * @param {function(number,number):Object} get (row,col) → セル
   * @param {{r1:number,c1:number,r2:number,c2:number}} rng 表の範囲
   * @param {boolean} 見出しあり 先頭行を見出しとして扱うか
   * @param {number} keyCol 見る列
   * @param {string} 値 この値と同じ行だけ残す
   * @returns {{hide:number[], keep:number[], value:string}}
   */
  function byValue(get, rng, 見出しあり, keyCol, 値) {
    var 先頭 = rng.r1 + (見出しあり ? 1 : 0);
    var hide = [], keep = [];
    var 見たい = String(値 === undefined || 値 === null ? '' : 値).trim();
    for (var r = 先頭; r <= rng.r2; r++) {
      if (text(get(r, keyCol)) === 見たい) keep.push(r); else hide.push(r);
    }
    /* ★見出し行は隠さない★（実測: 1行目は False のまま） */
    return { hide: hide, keep: keep, value: 見たい };
  }

  /** その列に入っている値の種類（絞る候補。同じ物は1つにまとめる） */
  function values(get, rng, 見出しあり, keyCol) {
    var 先頭 = rng.r1 + (見出しあり ? 1 : 0);
    var 見た = {}, 並び = [];
    for (var r = 先頭; r <= rng.r2; r++) {
      var t = text(get(r, keyCol));
      if (見た[t]) continue;
      見た[t] = true; 並び.push(t);
    }
    return 並び;
  }

  var api = { byValue: byValue, values: values, text: text };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.GridFilter = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
