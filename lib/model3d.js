/* model3d.js — ★3D モデル（挿入→図→3D モデル）★ 2026-08-30
 *
 *  ★実Excel で 測れた 事★
 *    3D モデルは ★形（Shapes）の 1つ★として 置かれる。
 *    ★中の 三角形の 数や 見せ方は COM から 読めなかった★
 *      ⇒ ★真似る 相手の 数字が 無い★ので ★うちで 決めて うちで 測る★。
 *
 *  ★うちの 決め★
 *    ・読む 形 … ★.obj（字だけの ファイル）★
 *        `v x y z`（点）／`f a b c`（面。4つ 以上でも 三角に 切る）
 *        ★.glb / .fbx は 中身が 読めない★（決まりが 大きい）ので ★読めないと 言う★
 *    ・見せ方 … ★線だけ（ワイヤーフレーム）★＝面の 影は 付けない
 *    ・向き … 横（Y軸）と 縦（X軸）で 回す
 *    ・大きさ … ★対角線★を 箱に 合わせる
 *        ＝★どんな 大きさの モデルでも・どんな 向きに 回しても はみ出さない★
 *        （一番 長い 辺で 合わせると 回した 時に はみ出す＝08-30 実測）
 *
 *  ★出来ない事は 出来ないと 書く★
 *    ・色・材質・影は ★付けない★（.obj の 色は 別の ファイル）
 *    ・アニメは ★無い★
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.Model3D = factory();
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  /** ★.obj を 読む★（点と 面だけ。読めなければ null） */
  function objを読む(字) {
    var s = String(字 === null || 字 === undefined ? '' : 字);
    if (!s.trim()) return null;
    var 点 = [], 面 = [];
    var 行たち = s.split(/\r?\n/);
    for (var i = 0; i < 行たち.length; i++) {
      var t = 行たち[i].trim();
      if (!t || t.charAt(0) === '#') continue;
      var 部 = t.split(/\s+/);
      if (部[0] === 'v') {
        var x = Number(部[1]), y = Number(部[2]), z = Number(部[3]);
        if (!isFinite(x) || !isFinite(y) || !isFinite(z)) continue;
        点.push({ x: x, y: y, z: z });
      } else if (部[0] === 'f') {
        var 番 = [];
        for (var j = 1; j < 部.length; j++) {
          /* `1` `1/2` `1//3` `1/2/3` の どれでも 先頭だけ 使う */
          var n = parseInt(String(部[j]).split('/')[0], 10);
          if (!isFinite(n)) continue;
          if (n < 0) n = 点.length + n + 1;              /* 負の 番号は 後ろから */
          番.push(n - 1);                                 /* obj は 1 から */
        }
        /* 4つ 以上は ★三角に 切る★ */
        for (var k = 2; k < 番.length; k++) 面.push([番[0], 番[k - 1], 番[k]]);
      }
    }
    if (!点.length) return null;
    return { 点: 点, 面: 面 };
  }

  /** 全部の 点が 入る 箱 */
  function 箱(形) {
    var b = { x1: Infinity, y1: Infinity, z1: Infinity, x2: -Infinity, y2: -Infinity, z2: -Infinity };
    for (var i = 0; i < 形.点.length; i++) {
      var p = 形.点[i];
      if (p.x < b.x1) b.x1 = p.x; if (p.x > b.x2) b.x2 = p.x;
      if (p.y < b.y1) b.y1 = p.y; if (p.y > b.y2) b.y2 = p.y;
      if (p.z < b.z1) b.z1 = p.z; if (p.z > b.z2) b.z2 = p.z;
    }
    return b;
  }

  /** ★画面の 点に 直す★（回して・大きさを 合わせて・平らに する） */
  function 画面の点(形, 横, 縦, W, H) {
    var b = 箱(形);
    var cx = (b.x1 + b.x2) / 2, cy = (b.y1 + b.y2) / 2, cz = (b.z1 + b.z2) / 2;
    /* ★どんな 向きに 回しても 箱に 入る★ように ★対角線★で 大きさを 決める
       （08-30 実測＝一番 長い 辺で 合わせると 回した 時に はみ出した） */
    var dx = b.x2 - b.x1, dy = b.y2 - b.y1, dz = b.z2 - b.z1;
    var 幅 = Math.sqrt(dx * dx + dy * dy + dz * dz) || 1;
    var 倍 = (Math.min(W, H) * 0.9) / 幅;
    var ch = Math.cos(横), sh = Math.sin(横);
    var cv = Math.cos(縦), sv = Math.sin(縦);
    var 出 = [];
    for (var i = 0; i < 形.点.length; i++) {
      var p = 形.点[i];
      var x = p.x - cx, y = p.y - cy, z = p.z - cz;
      /* 横（Y軸）に 回す */
      var x1 = x * ch + z * sh;
      var z1 = -x * sh + z * ch;
      /* 縦（X軸）に 回す */
      var y1 = y * cv - z1 * sv;
      var z2 = y * sv + z1 * cv;
      出.push({ x: W / 2 + x1 * 倍, y: H / 2 - y1 * 倍, z: z2 });
    }
    return 出;
  }

  /** 線（辺）を 出す（★同じ 辺を 2回 出さない★） */
  function 辺たち(形) {
    var 見た = {};
    var 出 = [];
    for (var i = 0; i < 形.面.length; i++) {
      var f = 形.面[i];
      for (var j = 0; j < f.length; j++) {
        var a = f[j], b = f[(j + 1) % f.length];
        if (a === undefined || b === undefined) continue;
        var k = (a < b) ? (a + '-' + b) : (b + '-' + a);
        if (見た[k]) continue;
        見た[k] = true;
        出.push([a, b]);
      }
    }
    return 出;
  }

  /** ★見本の 立方体★（ファイルが 無くても 試せる） */
  function 立方体() {
    return objを読む([
      'v -1 -1 -1', 'v 1 -1 -1', 'v 1 1 -1', 'v -1 1 -1',
      'v -1 -1 1', 'v 1 -1 1', 'v 1 1 1', 'v -1 1 1',
      'f 1 2 3 4', 'f 5 6 7 8', 'f 1 2 6 5', 'f 2 3 7 6',
      'f 3 4 8 7', 'f 4 1 5 8',
    ].join('\n'));
  }

  return {
    objを読む: objを読む, 箱: 箱, 画面の点: 画面の点, 辺たち: 辺たち, 立方体: 立方体,
  };
}));
