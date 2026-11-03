/* diagram.js — ★SmartArt（図解）・アイコン・数式・チェックボックス★ 2026-08-30
 *
 *  ★真値（実Excel 16.0 で 実測）★ … tools/measure-insert.ps1
 *
 *  ● SmartArt（挿入→図→SmartArt）
 *      ・型（レイアウト）の 数 … ★159個★
 *      ・1〜8番目の 名前 … カード型リスト／左右交替積み上げ六角形／画像リスト／
 *        線区切りリスト／縦方向箇条書きリスト／縦方向リスト／横方向箇条書きリスト／
 *        箇条書き記号アクセント リスト
 *      ・置いた 時の 形の 名前 … `Diagram 1`
 *      ・★中の 節（Nodes）の 数は はじめ 5★／`Add()` で 6 に なる
 *
 *  ● チェック ボックス（挿入→コントロール）
 *      ・新しい 形は ★セルの 中の TRUE / FALSE★（型 Boolean）
 *      ・画面に 出る 字 … ★`TRUE` / `FALSE`（大文字）★
 *      ・`=COUNTIF(A1:A2,TRUE)` … 1 ／ ★`=SUM(A1:A2)` … 0★ ／ ★`=A1+A2` … 1★
 *      ・古い（フォーム）チェックボックス … 名前 `Check Box 1`／
 *        大きさ ★100.125 × 20.25★／はじめの 値 ★-4146（どちらでもない）★／字 `チェック 1`
 *
 *  ● アイコン（挿入→図→アイコン）
 *      ・★COM から 一覧を 出せない★（絵は Microsoft の クラウドから 来る）
 *      ⇒ ★うちは うちで 描く★（線だけの 単純な 形。写さない）
 *
 *  ● 数式（挿入→記号と特殊文字→数式）
 *      ・テキスト ボックスは 足せる。★中の 数式（Math）は COM から 読めなかった（0）★
 *      ⇒ ★書き方（分数・累乗・平方根・添え字）を 決めて うちで 描く★
 *
 *  ★色は 写さない★＝形と 働きは 同じに するが 色は うちの 緑。
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.Diagram = factory();
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  /* ★実Excel の 1〜8番目の 名前（実測）★＝書き残すだけ。うちは 下の 8種を 作る */
  var 実Excelの型8 = ['カード型リスト', '左右交替積み上げ六角形', '画像リスト', '線区切りリスト',
    '縦方向箇条書きリスト', '縦方向リスト', '横方向箇条書きリスト', '箇条書き記号アクセント リスト'];
  var 実Excelの型の数 = 159;
  var 実Excelの節の数 = 5;          /* ★置いた 時 5個★ */

  /* うちの 図解（8種）＝形の 並べ方だけ 決める。中身の 字は 人が 入れる */
  var 図解たち = [
    { 名: '縦のリスト',   種: 'list-v' },
    { 名: '横のリスト',   種: 'list-h' },
    { 名: '手順（→）',   種: 'process' },
    { 名: 'ぐるぐる（輪）', 種: 'cycle' },
    { 名: '組織図（木）', 種: 'tree' },
    { 名: 'ピラミッド',   種: 'pyramid' },
    { 名: 'つり合い（天秤）', 種: 'balance' },
    { 名: '入れ子（丸）', 種: 'venn' },
  ];

  /** 新しい 図解を 作る（★節は 5個から★＝実Excel と 同じ） */
  function 図解を作る(種, 字たち) {
    var 節 = (字たち && 字たち.length) ? 字たち.slice()
      : ['項目1', '項目2', '項目3', '項目4', '項目5'];   /* ★5個★ */
    return { 種類: '図解', 図: 種 || 'list-v', 節: 節, x: 0, y: 0, w: 320, h: 220 };
  }

  /** 節の 置き場所を 数える（描く 側は これを 使うだけ） */
  function 節の場所(図, w, h, 数) {
    var 出 = [];
    var i;
    if (数 <= 0) return 出;
    if (図 === 'list-v' || 図 === 'process' && false) { /* 縦に 積む */
      var hh = h / 数;
      for (i = 0; i < 数; i++) 出.push({ x: 0, y: i * hh, w: w, h: hh * 0.86, 矢: false });
      return 出;
    }
    if (図 === 'list-h') {
      var ww = w / 数;
      for (i = 0; i < 数; i++) 出.push({ x: i * ww, y: h * 0.3, w: ww * 0.9, h: h * 0.4, 矢: false });
      return 出;
    }
    if (図 === 'process') {              /* 横に 並べて あいだに 矢印 */
      var w2 = w / 数;
      for (i = 0; i < 数; i++) 出.push({ x: i * w2, y: h * 0.3, w: w2 * 0.78, h: h * 0.4, 矢: i < 数 - 1 });
      return 出;
    }
    if (図 === 'cycle') {                /* 輪に 並べる */
      var cx = w / 2, cy = h / 2;
      var R = Math.min(w, h) * 0.34;
      var 大 = Math.min(w, h) * 0.24;
      for (i = 0; i < 数; i++) {
        var 角 = -Math.PI / 2 + i * 2 * Math.PI / 数;
        出.push({ x: cx + Math.cos(角) * R - 大 / 2, y: cy + Math.sin(角) * R - 大 / 2,
                  w: 大, h: 大, 丸: true, 矢: true });
      }
      return 出;
    }
    if (図 === 'tree') {                 /* 1つ目が 親・残りが 子 */
      var 子 = 数 - 1;
      出.push({ x: w * 0.3, y: 0, w: w * 0.4, h: h * 0.26, 矢: false });
      if (子 > 0) {
        var w3 = w / 子;
        for (i = 0; i < 子; i++)
          出.push({ x: i * w3, y: h * 0.6, w: w3 * 0.86, h: h * 0.26, 線: 0 });
      }
      return 出;
    }
    if (図 === 'pyramid') {              /* 下ほど 広い */
      var hh2 = h / 数;
      for (i = 0; i < 数; i++) {
        var 幅 = w * (i + 1) / 数;
        出.push({ x: (w - 幅) / 2, y: i * hh2, w: 幅, h: hh2 * 0.86, 台形: true });
      }
      return 出;
    }
    if (図 === 'balance') {              /* 左右に 分ける */
      var 半 = Math.ceil(数 / 2);
      for (i = 0; i < 数; i++) {
        var 左か = i < 半;
        var j = 左か ? i : i - 半;
        var 個 = 左か ? 半 : (数 - 半);
        出.push({ x: 左か ? 0 : w * 0.55, y: (h / Math.max(1, 個)) * j,
                  w: w * 0.45, h: (h / Math.max(1, 個)) * 0.8 });
      }
      return 出;
    }
    /* venn＝重なる 丸 */
    var R2 = Math.min(w, h) * 0.28;
    for (i = 0; i < 数; i++) {
      var 角2 = -Math.PI / 2 + i * 2 * Math.PI / 数;
      var r2 = 数 === 1 ? 0 : R2 * 0.62;
      出.push({ x: w / 2 + Math.cos(角2) * r2 - R2, y: h / 2 + Math.sin(角2) * r2 - R2,
                w: R2 * 2, h: R2 * 2, 丸: true, 透け: 0.55 });
    }
    return 出;
  }

  /* ── アイコン（★線だけで うちが 描く★・写さない）── */
  var アイコンたち = [
    { 名: '人',     種: 'person' },
    { 名: '家',     種: 'home' },
    { 名: '星',     種: 'star' },
    { 名: 'チェック', 種: 'check' },
    { 名: '時計',   種: 'clock' },
    { 名: '封筒',   種: 'mail' },
    { 名: 'グラフ', 種: 'chart' },
    { 名: '歯車',   種: 'gear' },
    { 名: '矢印',   種: 'arrow' },
    { 名: '丸',     種: 'circle' },
    { 名: '四角',   種: 'square' },
    { 名: '三角',   種: 'triangle' },
  ];

  /* ── 数式（うちの 書き方）── */
  /** `a/b` `a^b` `sqrt(a)` `a_b` を 組み立てて 並べる物に する */
  function 数式を組む(字) {
    var s = String(字 || '');
    var 出 = [];
    var i = 0;
    while (i < s.length) {
      var m = /^sqrt\(([^)]*)\)/.exec(s.slice(i));
      if (m) { 出.push({ 型: '根', 中: m[1] }); i += m[0].length; continue; }
      m = /^([A-Za-z0-9.]+)\/([A-Za-z0-9.]+)/.exec(s.slice(i));
      if (m) { 出.push({ 型: '分数', 上: m[1], 下: m[2] }); i += m[0].length; continue; }
      m = /^([A-Za-z0-9.]+)\^([A-Za-z0-9.]+)/.exec(s.slice(i));
      if (m) { 出.push({ 型: '累乗', 元: m[1], 肩: m[2] }); i += m[0].length; continue; }
      m = /^([A-Za-z0-9.]+)_([A-Za-z0-9.]+)/.exec(s.slice(i));
      if (m) { 出.push({ 型: '添え字', 元: m[1], 下: m[2] }); i += m[0].length; continue; }
      /* ふつうの 字（次の 記号まで） */
      m = /^[^]/.exec(s.slice(i));
      var 前 = 出[出.length - 1];
      if (前 && 前.型 === '字') 前.字 += m[0];
      else 出.push({ 型: '字', 字: m[0] });
      i += 1;
    }
    return 出;
  }

  return {
    実Excelの型8: 実Excelの型8, 実Excelの型の数: 実Excelの型の数, 実Excelの節の数: 実Excelの節の数,
    図解たち: 図解たち, 図解を作る: 図解を作る, 節の場所: 節の場所,
    アイコンたち: アイコンたち, 数式を組む: 数式を組む,
  };
}));
