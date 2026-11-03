/* ink.js — ★手書き（描画タブ）★ 2026-08-30
 *
 *  ★真値★＝実Excelの リボンを 機械で 取った 正本（docs/excel-ribbon-flat.tsv）に
 *  ★ペンの 色と 太さが そのまま 書いてある★:
 *      ペン: 黒、0.35 mm ／ ペン: 赤、0.35 mm ／ 鉛筆書き: 灰色、0.5 mm
 *      蛍光ペン: 黄, 6mm、テキストに位置を合わせる: オフ
 *      万年筆: 濃い青、1 mm ／ ブラシ ペン: 緑、1 mm
 *  ⇒ ★色も 太さも その とおりに する★（ここは 見た目では なく ★道具の 仕様★なので 合わせる）
 *
 *  ★mm → 画面の 点★ … 1inch = 25.4mm ＝ 96点（実測: Excelの 1inch = 72pt／画面は 96dpi）
 *      0.35mm ≒ 1.32点 ／ 0.5mm ≒ 1.89点 ／ 1mm ≒ 3.78点 ／ 6mm ≒ 22.68点
 *
 *  ★前の 決めを 直した（2026-08-30）★
 *      前は「うちは 表計算に 絞る」ので 描画タブを 出さない と していた。
 *      ★司さん「Excel全機能全能力が Exallyに 入って」「サボるな」★＝
 *      ブラウザの canvas で ★本当に 書ける★ので 作った。
 *      （出来ない物＝インクを 図形/数式に 変える＝形の 読み取りが 要る。まだ 出していない）
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.Ink = factory();
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var 一mmの点 = 96 / 25.4;                  /* ★1mm ＝ 3.779527… 点★ */
  function mmを点に(mm) { return Math.round(mm * 一mmの点 * 100) / 100; }

  /* ★実Excelの 項目名に 書いてある とおり★ */
  var ペンたち = [
    { 名: 'ペン（黒）',   色: '#000000', mm: 0.35, 透け: 1,   Excel: 'ペン: 黒、0.35 mm' },
    { 名: 'ペン（赤）',   色: '#FF0000', mm: 0.35, 透け: 1,   Excel: 'ペン: 赤、0.35 mm' },
    { 名: '鉛筆',         色: '#808080', mm: 0.5,  透け: 1,   Excel: '鉛筆書き: 灰色、0.5 mm' },
    { 名: '蛍光ペン',     色: '#FFFF00', mm: 6,    透け: 0.4, Excel: '蛍光ペン: 黄, 6mm' },
    { 名: '万年筆',       色: '#00008B', mm: 1,    透け: 1,   Excel: '万年筆: 濃い青、1 mm' },
    { 名: 'ブラシ ペン',  色: '#008000', mm: 1,    透け: 1,   Excel: 'ブラシ ペン: 緑、1 mm' },
  ];

  /** 1本の 線（ストローク）を 作る */
  function 線を始める(ペン, x, y) {
    return { 色: ペン.色, 太さ: mmを点に(ペン.mm), 透け: ペン.透け, 点: [{ x: x, y: y }], 名: ペン.名 };
  }
  function 線を伸ばす(線, x, y) {
    if (!線) return null;
    var 前 = 線.点[線.点.length - 1];
    /* ★近すぎる 点は 入れない★（重くなるだけ） */
    if (前 && Math.abs(前.x - x) < 1 && Math.abs(前.y - y) < 1) return 線;
    線.点.push({ x: x, y: y });
    return 線;
  }

  /** 描く（★shift は 画面の ずれ＝スクロールに 合わせる★） */
  function 描く(ctx, 線, ずれX, ずれY) {
    if (!線 || !線.点 || 線.点.length < 1) return false;
    ずれX = ずれX || 0; ずれY = ずれY || 0;
    ctx.save();
    ctx.globalAlpha = (線.透け === undefined) ? 1 : 線.透け;
    ctx.strokeStyle = 線.色;
    ctx.lineWidth = 線.太さ;
    ctx.lineCap = 'round';
    ctx.lineJoin = 'round';
    ctx.beginPath();
    for (var i = 0; i < 線.点.length; i++) {
      var p = 線.点[i];
      if (i === 0) ctx.moveTo(p.x + ずれX, p.y + ずれY);
      else ctx.lineTo(p.x + ずれX, p.y + ずれY);
    }
    if (線.点.length === 1) {
      /* ★1点だけの 時も 見えるように 打つ★ */
      ctx.lineTo(線.点[0].x + ずれX + 0.1, 線.点[0].y + ずれY + 0.1);
    }
    ctx.stroke();
    ctx.restore();
    return true;
  }

  /** その 点に 近い 線（消しゴム・選ぶ で 使う） */
  function 近い線(線たち, x, y, どれくらい) {
    var 幅 = どれくらい || 6;
    for (var i = (線たち || []).length - 1; i >= 0; i--) {   /* ★後から 引いた 線が 上★ */
      var s = 線たち[i];
      for (var j = 0; j < s.点.length; j++) {
        var p = s.点[j];
        var あそび = Math.max(幅, (s.太さ || 1) / 2 + 2);
        if (Math.abs(p.x - x) <= あそび && Math.abs(p.y - y) <= あそび) return s;
      }
    }
    return null;
  }

  /** 四角の 中に 入っている 線（なげなわ選択） */
  function 中に在る線(線たち, x1, y1, x2, y2) {
    var 左 = Math.min(x1, x2), 右 = Math.max(x1, x2);
    var 上 = Math.min(y1, y2), 下 = Math.max(y1, y2);
    var 出 = [];
    for (var i = 0; i < (線たち || []).length; i++) {
      var s = 線たち[i], 全部 = true;
      for (var j = 0; j < s.点.length; j++) {
        var p = s.点[j];
        if (p.x < 左 || p.x > 右 || p.y < 上 || p.y > 下) { 全部 = false; break; }
      }
      if (全部 && s.点.length) 出.push(s);
    }
    return 出;
  }

  return {
    ペンたち: ペンたち, mmを点に: mmを点に, 一mmの点: 一mmの点,
    線を始める: 線を始める, 線を伸ばす: 線を伸ばす, 描く: 描く,
    近い線: 近い線, 中に在る線: 中に在る線,
  };
}));
