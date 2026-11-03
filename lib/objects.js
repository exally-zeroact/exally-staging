/* objects.js — ★シートの 上に 浮かぶ 物（画像・図形・テキスト）★ 2026-08-30
 *
 *  ★真値（実Excel 16.0 で 実測）★ … tools/measure-shapes.ps1 / tools/measure-colors.ps1
 *    ・図形（四角）… 名前 'Rectangle 1'／塗り ★#156082★／線 ★#042433★／線の太さ ★1.5★
 *    ・テキスト ボックス … 名前 'TextBox 2'／塗り ★#FFFFFF★（見える）／線も 見える
 *    ・重なりの 順（ZOrderPosition）… ★1 から★／前面へ で 2・背面へ で 1
 *    ・図形の 種類の 番号 … 四角=1・角丸=5・円=9・三角=7・右矢印=33・星5=92
 *
 *  ★色は 写さない★（司さん「訴えられんような 見せ方で 同じように」）
 *    ＝形と 働きは 同じに するが、色は ★うちの 緑★を 使う。実測値は 上に 書き残す。
 *
 *  ★重なりは 1 から★（実Excelと 同じ）＝一番 後ろが 1。
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.SheetObjects = factory();
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  /* うちの 色（Excelの 色は 使わない。実測値は 見出しに 残した） */
  var 色 = {
    塗り: '#D9F0E4',
    線: '#3D9E72',
    字: '#1A2B22',
    えらんだ: '#1E88E5',
  };
  var 線の太さ = 1.5;        /* ★実測と 同じ★ */

  var 形たち = [
    { 名: '四角', Excel: 1 },
    { 名: '角丸四角', Excel: 5 },
    { 名: '丸', Excel: 9 },
    { 名: '三角', Excel: 7 },
    { 名: '右矢印', Excel: 33 },
    { 名: '星', Excel: 92 },
    { 名: '線', Excel: null },
  ];

  /** 名前を 決める（実Excel は 'Rectangle 1' のように 種類＋番号） */
  function 名前を決める(もうある, 種類) {
    var n = 1;
    var 使った = {};
    for (var i = 0; i < (もうある || []).length; i++) 使った[String(もうある[i])] = true;
    while (使った[種類 + ' ' + n]) n++;
    return 種類 + ' ' + n;
  }

  /** 重なりの 順を そろえる（★1 から 続き番号★＝実Excelと 同じ） */
  function 順をそろえる(物たち) {
    var 並び = (物たち || []).slice().sort(function (a, b) { return (a.z || 0) - (b.z || 0); });
    for (var i = 0; i < 並び.length; i++) 並び[i].z = i + 1;
    return 並び.length;
  }
  /** 前面へ（一番 上に する） */
  function 前面へ(物たち, 物) {
    if (!物) return 0;
    物.z = 1e9;
    順をそろえる(物たち);
    return 物.z;
  }
  /** 背面へ（一番 下に する） */
  function 背面へ(物たち, 物) {
    if (!物) return 0;
    物.z = -1;
    順をそろえる(物たち);
    return 物.z;
  }

  /** その 点に 在る 物（★上に 在る 物が 勝つ★） */
  function どれを押したか(物たち, x, y) {
    var 見つけ = null;
    for (var i = 0; i < (物たち || []).length; i++) {
      var o = 物たち[i];
      if (x < o.x || x > o.x + o.w || y < o.y || y > o.y + o.h) continue;
      if (!見つけ || (o.z || 0) > (見つけ.z || 0)) 見つけ = o;
    }
    return 見つけ;
  }

  /** 描く（canvas の ctx へ。★セルの 上に 出す★） */
  function 描く(ctx, o, えらんでいるか) {
    if (!o || o.w <= 0 || o.h <= 0) return false;
    ctx.save();
    /* ★回転★（実測＝90 を 足していく）＝真ん中を 軸に 回す */
    if (o.回転) {
      ctx.translate(o.x + o.w / 2, o.y + o.h / 2);
      ctx.rotate(o.回転 * Math.PI / 180);
      ctx.translate(-(o.x + o.w / 2), -(o.y + o.h / 2));
    }
    ctx.lineWidth = 線の太さ;
    ctx.strokeStyle = o.線 || 色.線;
    ctx.fillStyle = o.塗り || 色.塗り;
    if (o.種類 === '画像' && o.画像) {
      try { ctx.drawImage(o.画像, o.x, o.y, o.w, o.h); } catch (e) { /* まだ 読めていない */ }
    } else if (o.種類 === 'テキスト') {
      ctx.fillStyle = '#FFFFFF';
      ctx.fillRect(o.x, o.y, o.w, o.h);
      ctx.strokeRect(o.x, o.y, o.w, o.h);
    } else if (o.種類 === '丸') {
      ctx.beginPath();
      ctx.ellipse(o.x + o.w / 2, o.y + o.h / 2, o.w / 2, o.h / 2, 0, 0, Math.PI * 2);
      ctx.fill(); ctx.stroke();
    } else if (o.種類 === '三角') {
      ctx.beginPath();
      ctx.moveTo(o.x + o.w / 2, o.y);
      ctx.lineTo(o.x + o.w, o.y + o.h);
      ctx.lineTo(o.x, o.y + o.h);
      ctx.closePath(); ctx.fill(); ctx.stroke();
    } else if (o.種類 === '右矢印') {
      var 首 = o.h * 0.3, 頭 = o.w * 0.35;
      ctx.beginPath();
      ctx.moveTo(o.x, o.y + o.h / 2 - 首 / 2);
      ctx.lineTo(o.x + o.w - 頭, o.y + o.h / 2 - 首 / 2);
      ctx.lineTo(o.x + o.w - 頭, o.y);
      ctx.lineTo(o.x + o.w, o.y + o.h / 2);
      ctx.lineTo(o.x + o.w - 頭, o.y + o.h);
      ctx.lineTo(o.x + o.w - 頭, o.y + o.h / 2 + 首 / 2);
      ctx.lineTo(o.x, o.y + o.h / 2 + 首 / 2);
      ctx.closePath(); ctx.fill(); ctx.stroke();
    } else if (o.種類 === '星') {
      var cx = o.x + o.w / 2, cy = o.y + o.h / 2;
      var R = Math.min(o.w, o.h) / 2, r = R * 0.4;
      ctx.beginPath();
      for (var k = 0; k < 10; k++) {
        var 角 = -Math.PI / 2 + k * Math.PI / 5;
        var 長 = (k % 2 === 0) ? R : r;
        var px = cx + Math.cos(角) * 長, py = cy + Math.sin(角) * 長;
        if (k === 0) ctx.moveTo(px, py); else ctx.lineTo(px, py);
      }
      ctx.closePath(); ctx.fill(); ctx.stroke();
    } else if (o.種類 === 'グループ') {
      /* ★まとめた 物★＝中の 物を そのまま 描き、まわりに 薄い 枠 */
      for (var g = 0; g < (o.中身 || []).length; g++) {
        var 子 = o.中身[g];
        描く(ctx, { 種類: 子.物.種類, x: o.x + 子.ずれX, y: o.y + 子.ずれY,
          w: 子.物.w, h: 子.物.h, 字: 子.物.字, 塗り: 子.物.塗り, 線: 子.物.線,
          画像: 子.物.画像, 回転: 子.物.回転 }, false);
      }
      ctx.strokeStyle = 色.線;
      ctx.setLineDash([2, 3]);
      ctx.strokeRect(o.x, o.y, o.w, o.h);
      ctx.setLineDash([]);
    } else if (o.種類 === '線') {
      ctx.beginPath();
      ctx.moveTo(o.x, o.y);
      ctx.lineTo(o.x + o.w, o.y + o.h);
      ctx.stroke();
    } else if (o.種類 === '角丸四角') {
      var r2 = Math.min(12, o.w / 4, o.h / 4);
      ctx.beginPath();
      ctx.moveTo(o.x + r2, o.y);
      ctx.lineTo(o.x + o.w - r2, o.y);
      ctx.quadraticCurveTo(o.x + o.w, o.y, o.x + o.w, o.y + r2);
      ctx.lineTo(o.x + o.w, o.y + o.h - r2);
      ctx.quadraticCurveTo(o.x + o.w, o.y + o.h, o.x + o.w - r2, o.y + o.h);
      ctx.lineTo(o.x + r2, o.y + o.h);
      ctx.quadraticCurveTo(o.x, o.y + o.h, o.x, o.y + o.h - r2);
      ctx.lineTo(o.x, o.y + r2);
      ctx.quadraticCurveTo(o.x, o.y, o.x + r2, o.y);
      ctx.closePath(); ctx.fill(); ctx.stroke();
    } else {
      ctx.fillRect(o.x, o.y, o.w, o.h);
      ctx.strokeRect(o.x, o.y, o.w, o.h);
    }
    /* 中の 字（実Excel も 図形に 字が 入る＝実測 TextFrame2 在り） */
    if (o.字) {
      ctx.fillStyle = 色.字;
      ctx.font = '13px "Noto Sans JP",sans-serif';
      ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
      var 行 = String(o.字).split('\n');
      var 上 = o.y + o.h / 2 - (行.length - 1) * 8;
      for (var i2 = 0; i2 < 行.length; i2++) ctx.fillText(行[i2], o.x + o.w / 2, 上 + i2 * 16);
    }
    /* 選んでいる 印（つまみ） */
    if (えらんでいるか) {
      ctx.strokeStyle = 色.えらんだ;
      ctx.lineWidth = 1;
      ctx.setLineDash([4, 3]);
      ctx.strokeRect(o.x - 2, o.y - 2, o.w + 4, o.h + 4);
      ctx.setLineDash([]);
      ctx.fillStyle = 色.えらんだ;
      ctx.fillRect(o.x + o.w - 4, o.y + o.h - 4, 8, 8);
    }
    ctx.restore();
    return true;
  }

  return {
    色: 色, 線の太さ: 線の太さ, 形たち: 形たち,
    名前を決める: 名前を決める, 順をそろえる: 順をそろえる,
    前面へ: 前面へ, 背面へ: 背面へ, どれを押したか: どれを押したか, 描く: 描く,
  };
}));
