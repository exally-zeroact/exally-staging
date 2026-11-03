/* sparkline.js — ★スパークライン（セルの中の 小さな グラフ）★ 2026-08-30
 *
 *  ★真値（実Excel 16.0 で 実測）★ … tools/measure-sparkline.ps1 / measure-spark2.ps1
 *    種類＝折れ線1・縦棒2・勝敗3
 *
 *      種類    線の太さ  高い点  低い点  最初  最後  印   ★マイナス★  縦軸の型  横軸  空の扱い
 *      折れ線   0.75     False   False  False False False   False        2       False    1
 *      縦棒     0.75     False   False  False False False   False        2       False    1
 *      勝敗     0.75     False   False  False False False  ★True★       2       False    1
 *
 *    ⇒ ①★飾りは 何も 出さない★（高い点・低い点・最初・最後・印 は 全部 False）
 *      ②★勝敗だけ マイナスを 別に 出す★（Negative.Visible=True）
 *      ③縦軸の型＝★2★（Excelの XlSparkScale では 2＝スパークラインごと
 *         ＝画面の「自動（各スパークライン）」と 同じ）
 *      ④横軸は ★出さない★（Axes.Horizontal.Axis.Visible=False）
 *      ⑤空の扱い＝★1★（XlDisplayBlanksAs の 1＝描かない＝★隙間を あける★）
 *
 *  ★見せ方★＝Excelの 絵も 配色も 1つも 写さない。色は うちの緑。
 *  ★出来ない事は 黙らない★＝数が 1つも 無ければ 描かずに false を 返す。
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.Sparkline = factory();
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var 色 = {
    ふつう: '#3D9E72',
    マイナス: '#C0392B',   /* ★勝敗の マイナスだけ 別（実測 Negative.Visible=True）★ */
  };
  var 線の太さ = 0.75;      /* 実測 */

  function 数にする(v) {
    if (v === null || v === undefined || v === '') return null;
    if (typeof v === 'number') return isFinite(v) ? v : null;
    var s = String(v).replace(/[,¥￥\s]/g, '');
    if (s === '') return null;
    var n = Number(s);
    return isFinite(n) ? n : null;
  }

  /** 並んだ 値を 数に する（空は null＝★隙間★＝実測⑤） */
  function 値にする(並び) {
    var 出 = [];
    for (var i = 0; i < (並び || []).length; i++) 出.push(数にする(並び[i]));
    return 出;
  }

  /**
   * セルの 中に 描く
   * @param 種類 … 'line' | 'column' | 'winloss'
   * @returns true（描いた）／false（★描く物が 無い★）
   */
  function 描く(ctx, 値, 種類, x, y, w, h) {
    値 = 値 || [];
    var ある = 0;
    for (var i = 0; i < 値.length; i++) if (値[i] !== null) ある++;
    if (!ある || w < 6 || h < 4) return false;

    var 余 = 2;
    var X = x + 余, Y = y + 余, W = w - 余 * 2, H = h - 余 * 2;
    if (W <= 0 || H <= 0) return false;

    var 個数 = 値.length;
    var 一枠 = W / Math.max(1, 個数);

    if (種類 === 'winloss') {
      /* ★勝敗★＝プラスは 上・マイナスは 下・0は 描かない（高さは 同じ） */
      var 真ん中 = Y + H / 2;
      var 棒幅2 = Math.max(1, 一枠 * 0.7);
      for (var j = 0; j < 個数; j++) {
        var v = 値[j];
        if (v === null || v === 0) continue;          /* ★0は 勝ちでも 負けでもない★ */
        ctx.fillStyle = v > 0 ? 色.ふつう : 色.マイナス;
        var 高 = H / 2 - 1;
        ctx.fillRect(X + 一枠 * j + (一枠 - 棒幅2) / 2, v > 0 ? (真ん中 - 高) : 真ん中, 棒幅2, 高);
      }
      return true;
    }

    /* ★縦軸は スパークラインごと（実測③）★＝この1本の 中だけで 最小/最大を 取る */
    var 最小 = null, 最大 = null;
    for (var k = 0; k < 個数; k++) {
      var v2 = 値[k];
      if (v2 === null) continue;
      if (最小 === null || v2 < 最小) 最小 = v2;
      if (最大 === null || v2 > 最大) 最大 = v2;
    }
    if (最小 === 最大) { 最小 = 最小 - 1; 最大 = 最大 + 1; }
    var py = function (v3) { return Y + H - (v3 - 最小) / (最大 - 最小) * H; };

    if (種類 === 'column') {
      var 棒幅 = Math.max(1, 一枠 * 0.7);
      /* 0 が 範囲の 中に 在れば そこを 底に する（Excel も マイナスは 下に 伸びる） */
      var 底 = (最小 <= 0 && 最大 >= 0) ? py(0) : py(最小);
      for (var m = 0; m < 個数; m++) {
        var v4 = 値[m];
        if (v4 === null) continue;                    /* ★空は 描かない（実測⑤）★ */
        ctx.fillStyle = v4 < 0 ? 色.マイナス : 色.ふつう;
        var y1 = py(v4);
        ctx.fillRect(X + 一枠 * m + (一枠 - 棒幅) / 2, Math.min(底, y1), 棒幅, Math.max(1, Math.abs(y1 - 底)));
      }
      return true;
    }

    /* 折れ線＝★空の所で 線を 切る（実測⑤ 隙間）★ */
    ctx.strokeStyle = 色.ふつう;
    ctx.lineWidth = 線の太さ;
    ctx.beginPath();
    var つながり = false;
    for (var n = 0; n < 個数; n++) {
      var v5 = 値[n];
      var cx = X + 一枠 * n + 一枠 / 2;
      if (v5 === null) { つながり = false; continue; }
      if (!つながり) { ctx.moveTo(cx, py(v5)); つながり = true; }
      else ctx.lineTo(cx, py(v5));
    }
    ctx.stroke();
    return true;
  }

  return { 描く: 描く, 値にする: 値にする, 色: 色, 線の太さ: 線の太さ, _数にする: 数にする };
}));
