/* ink-shape.js — ★手書きを 図形に する（描画→変換→インクを図形に変換）★ 2026-08-30
 *
 *  ★実Excel で 測れた 事★ … tools/measure-review2.ps1 ほか
 *    ・インクは ★形（Shapes）の 1つとして 数える★
 *    ・★変換の 中の 動きは COM から 見えない★（画面の ボタンだけ）
 *      ⇒ ★真似る 相手の 数字が 無い★ので、うちは ★自分で 決めて 自分で 測る★。
 *
 *  ★うちの 決め方（全部 数で 決める・当てずっぽうを しない）★
 *    ① 線を ★角で 切る★（向きが 大きく 変わる 所）
 *    ② ★閉じているか★（始めと 終わりが 近いか）を 見る
 *    ③ ★角の 数★と ★丸さ★で 形を 決める
 *         丸さ ＝ 4π×面積 ÷ まわりの長さ²（★真円で 1.0★）
 *    ④ 決めた 形の ★囲み★を そのまま 図形の 大きさに する
 *
 *  ★境目（実際に 測って 決めた）★
 *    ・閉じている … 始めと 終わりの 間が ★囲みの 対角線の 20%以下★
 *    ・角 … 向きの 変わりが ★60度 以上★
 *    ・丸 … 閉じていて 丸さ ★0.80 以上★
 *    ・三角 … 閉じていて 角 3個
 *    ・四角 … 閉じていて 角 4個（または 丸さ 0.60〜0.80 で 角 4個 前後）
 *    ・矢印 … 開いていて ★終わりの 近くに 戻る 角が 2つ★
 *    ・線 … それ以外で 開いている
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.InkShape = factory();
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var 閉じる割合 = 0.20;      /* 始めと 終わりが 対角線の 20%以下 なら 閉じている */
  var 角の度 = 60;            /* 向きが 60度 以上 変わったら 角 */
  var 丸の丸さ = 0.80;        /* 丸さ 0.80 以上 なら 丸 */

  function 囲み(点たち) {
    var 左 = null, 右 = null, 上 = null, 下 = null;
    for (var i = 0; i < 点たち.length; i++) {
      var p = 点たち[i];
      if (左 === null || p.x < 左) 左 = p.x;
      if (右 === null || p.x > 右) 右 = p.x;
      if (上 === null || p.y < 上) 上 = p.y;
      if (下 === null || p.y > 下) 下 = p.y;
    }
    return { 左: 左 || 0, 右: 右 || 0, 上: 上 || 0, 下: 下 || 0,
             w: (右 - 左) || 0, h: (下 - 上) || 0 };
  }

  /** まわりの 長さ */
  function 長さ(点たち) {
    var 出 = 0;
    for (var i = 1; i < 点たち.length; i++) {
      var dx = 点たち[i].x - 点たち[i - 1].x, dy = 点たち[i].y - 点たち[i - 1].y;
      出 += Math.sqrt(dx * dx + dy * dy);
    }
    return 出;
  }

  /** 面積（★靴ひもの 式★＝閉じた 形の 面積） */
  function 面積(点たち) {
    var 出 = 0;
    for (var i = 0; i < 点たち.length; i++) {
      var a = 点たち[i], b = 点たち[(i + 1) % 点たち.length];
      出 += a.x * b.y - b.x * a.y;
    }
    return Math.abs(出) / 2;
  }

  /** 点を まびく（近すぎる 点を 捨てる＝手の ふるえを 消す） */
  function まびく(点たち, 間) {
    if (!点たち || !点たち.length) return [];
    間 = 間 || 4;
    var 出 = [点たち[0]];
    for (var i = 1; i < 点たち.length; i++) {
      var 前 = 出[出.length - 1];
      var dx = 点たち[i].x - 前.x, dy = 点たち[i].y - 前.y;
      if (Math.sqrt(dx * dx + dy * dy) >= 間) 出.push(点たち[i]);
    }
    if (出.length < 2 && 点たち.length > 1) 出.push(点たち[点たち.length - 1]);
    return 出;
  }

  /** 角を 数える（向きが 大きく 変わる 所）
   *  ★閉じた 線は 端でも 回り込んで 見る★
   *  （08-30 実測＝回り込まないと 四角が 角3個・三角が 角2個に なり、形を 間違えた） */
  function 角たち(点たち, 閉じ) {
    var 出 = [];
    if (点たち.length < 3) return 出;
    var 幅 = Math.max(1, Math.round(点たち.length / 24));
    var n = 点たち.length;
    var 端 = 閉じ ? 0 : 幅;
    var 取 = function (k) {
      if (!閉じ) return 点たち[k];
      return 点たち[((k % n) + n) % n];
    };
    for (var i = 端; i < (閉じ ? n : n - 幅); i++) {
      var a = 取(i - 幅), b = 取(i), c = 取(i + 幅);
      var a1 = Math.atan2(b.y - a.y, b.x - a.x);
      var a2 = Math.atan2(c.y - b.y, c.x - b.x);
      var d = Math.abs(a2 - a1) * 180 / Math.PI;
      while (d > 180) d = 360 - d;
      if (d >= 角の度) {
        /* 近くの 角は 1つに まとめる */
        var 前 = 出[出.length - 1];
        if (前 && (i - 前.i) < 幅 * 2) {
          if (d > 前.度) { 出[出.length - 1] = { i: i, x: b.x, y: b.y, 度: d }; }
        } else {
          出.push({ i: i, x: b.x, y: b.y, 度: d });
        }
      }
    }
    /* ★閉じた 線は 端と 端も 隣どうし★＝1つの 角を 2回 数えない
       （08-30 実測＝これが 無いと 四角が 5個・三角が 4個に なった） */
    if (閉じ && 出.length >= 2) {
      var 頭 = 出[0], 尻 = 出[出.length - 1];
      if ((n - 尻.i) + 頭.i < 幅 * 2) {
        if (尻.度 > 頭.度) 出[0] = 尻;
        出.pop();
      }
    }
    return 出;
  }

  /** 閉じているか */
  function 閉じているか(点たち) {
    if (点たち.length < 3) return false;
    var w = 囲み(点たち);
    var 対角 = Math.sqrt(w.w * w.w + w.h * w.h) || 1;
    var a = 点たち[0], b = 点たち[点たち.length - 1];
    var dx = b.x - a.x, dy = b.y - a.y;
    return Math.sqrt(dx * dx + dy * dy) <= 対角 * 閉じる割合;
  }

  /** 丸さ（★真円で 1.0★） */
  function 丸さ(点たち) {
    var L = 長さ(点たち);
    if (L <= 0) return 0;
    return (4 * Math.PI * 面積(点たち)) / (L * L);
  }

  /** ★形を 決める★
   *  @returns {形, 角, 丸さ, 閉じ, 囲み}  形＝'丸'|'三角'|'四角'|'右矢印'|'線'
   */
  function 形を決める(点たち) {
    var 点 = まびく(点たち, 4);
    var w = 囲み(点);
    var 閉 = 閉じているか(点);
    var 角 = 角たち(点, 閉);
    var 丸 = 閉 ? 丸さ(点) : 0;
    var 形 = '線';
    if (点.length < 3) 形 = '線';
    else if (閉) {
      if (丸 >= 丸の丸さ) 形 = '丸';
      else if (角.length <= 3) 形 = '三角';
      else 形 = '四角';   /* 角 4個 以上＝四角（実測＝きれいな 四角は 4個・三角は 3個） */
    } else {
      /* 開いていて ★終わりの 近くに 戻る 角が 2つ★＝矢印 */
      var 終 = 点[点.length - 1];
      var 対角 = Math.sqrt(w.w * w.w + w.h * w.h) || 1;
      var 近い角 = 角.filter(function (v) {
        var dx = v.x - 終.x, dy = v.y - 終.y;
        return Math.sqrt(dx * dx + dy * dy) <= 対角 * 0.35;
      });
      形 = (近い角.length >= 2) ? '右矢印' : '線';
    }
    return { 形: 形, 角: 角.length, 丸さ: 丸, 閉じ: 閉, 囲み: w };
  }

  return {
    閉じる割合: 閉じる割合, 角の度: 角の度, 丸の丸さ: 丸の丸さ,
    囲み: 囲み, 長さ: 長さ, 面積: 面積, まびく: まびく,
    角たち: 角たち, 閉じているか: 閉じているか, 丸さ: 丸さ, 形を決める: 形を決める,
  };
}));
