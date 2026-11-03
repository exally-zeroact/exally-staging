/* theme.js — ★テーマ（配色・フォント・効果）と 色の 濃淡★ 2026-08-30
 *
 *  ★真値（実Excel 16.0 で 実測）★ … tools/measure-theme.ps1
 *
 *  ● 既定（Office）の 12色 … 役割の 名前と 並びは ★実Excel と 同じ★
 *      dk1 #000000 ／ lt1 #FFFFFF ／ dk2 #0E2841 ／ lt2 #E8E8E8
 *      accent1 #156082 ／ accent2 #E97132 ／ accent3 #196B24
 *      accent4 #0F9ED5 ／ accent5 #A02B93 ／ accent6 #4EA72E
 *      hlink #467886 ／ folHlink #96607D
 *  ● フォント … 見出し `Aptos Display`／本文 `Aptos Narrow`
 *      ただし ★セルに 実際に 入るのは 游ゴシック 11★（Application.StandardFont）
 *  ● 塗りの 色を テーマで 決めると そのまま その 色に なる（accent1 → #156082）
 *
 *  ★色の 濃淡（TintAndShade）は ★数まで 合わせた★★
 *      昔からの Windows の HLS（H・L・S を ★0〜240 の 整数★で 持つ）で
 *        t>0 … L' = ★floor(L × (1−t) + 240 × t)★
 *        t<0 … L' = ★floor(L × (1+t))★
 *      ⇒ 実測 15通り（accent1 を 12段・accent2 を 3段）と ★15/15 ぴったり★ 合った。
 *        （四捨五入だと 7/15・最大 2 ずれる。★切り捨てが 正しい★）
 *      例）#156082 … +0.4 → #44B3E1 ／ −0.25 → #104861 ／ +0.8 → #C0E6F5
 *
 *  ★★どちらの 本の 話かで 分かれます★★（2026-09-21 経営者1 と 一致）
 *    ・★うちが 新しく 作る 本★ ... ★うちの 配色★（既定の 字は `#1A2B22`）
 *        ＝下の 「色は 写さない」は ★この 話★ です。
 *    ・★お客さんが 持ち込んだ 本★ ... ★その 本の `xl/theme/theme1.xml` を そのまま 使う★
 *        ＝これは 「よその 配色を 写す」では なく ★お客さんの 書類を そのまま 出す★ 事。
 *        ＝読む 所 `lib/xlsx-kazari.js` の `テーマを読む()`（2026-09-21）
 *    ⇒★★持ち込んだ 本の 色を うちの 配色に 置き換えては いけません★★
 *
 *  ★色は 写さない★（司さん「訴えられんような 見せ方で 同じように」）
 *      ＝★仕組み（12の 役割・濃淡の 数え方）は 同じ★に するが、
 *        ★配色そのものは うちの 物★を 使う。実測値は 上に 書き残す。
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.Theme = factory();
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var HLSMAX = 240, RGBMAX = 255;

  /* ★役割の 名前と 並び★＝実Excel と 同じ 12個 */
  var 役割 = ['dk1', 'lt1', 'dk2', 'lt2',
    'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6',
    'hlink', 'folHlink'];
  var 役割の日本語 = {
    dk1: '濃い1', lt1: '薄い1', dk2: '濃い2', lt2: '薄い2',
    accent1: '差し色1', accent2: '差し色2', accent3: '差し色3',
    accent4: '差し色4', accent5: '差し色5', accent6: '差し色6',
    hlink: 'リンク', folHlink: '見たリンク',
  };

  /* ★実Excel の 既定（Office）★＝★使わない。書き残すだけ★ */
  var 実Excelの既定 = ['#000000', '#FFFFFF', '#0E2841', '#E8E8E8',
    '#156082', '#E97132', '#196B24', '#0F9ED5', '#A02B93', '#4EA72E',
    '#467886', '#96607D'];

  /* ★うちの 配色★（形は 同じ 12個・色は うちの 物） */
  var 配色たち = [
    { 名: 'みどり（既定）', 色: ['#1A2B22', '#FFFFFF', '#2E7D54', '#E8F5EE',
        '#3D9E72', '#E0A33E', '#4F86C6', '#C46B6B', '#8A6BC4', '#6BAF8A',
        '#2E7D54', '#7A5C8A'] },
    { 名: 'あお',           色: ['#16232E', '#FFFFFF', '#22506E', '#E8F1F7',
        '#3D82A8', '#E08B3E', '#6BA3C6', '#C4736B', '#7A6BC4', '#5FA8B5',
        '#22506E', '#6B5C8A'] },
    { 名: 'すみ（白黒）',   色: ['#1A1A1A', '#FFFFFF', '#3A3A3A', '#EFEFEF',
        '#5A5A5A', '#8A8A8A', '#7A7A7A', '#9A9A9A', '#6A6A6A', '#AAAAAA',
        '#3A3A3A', '#6A6A6A'] },
    { 名: 'あたたかい',     色: ['#2B211A', '#FFFFFF', '#7D4E2E', '#F7EFE8',
        '#B5763D', '#3D9E72', '#C6A24F', '#A85F5F', '#8A6B4F', '#C68A5F',
        '#7D4E2E', '#8A5C5C'] },
  ];

  /* フォントの 組（見出し・本文）＝★仕組みは 同じ・字は うちの 物★ */
  var フォント組たち = [
    { 名: 'ゴシック（既定）', 見出し: '"Noto Sans JP", sans-serif', 本文: '"Noto Sans JP", sans-serif' },
    { 名: '明朝',             見出し: '"Noto Serif JP", serif',     本文: '"Noto Serif JP", serif' },
    { 名: '見出しは明朝',     見出し: '"Noto Serif JP", serif',     本文: '"Noto Sans JP", sans-serif' },
  ];

  /* 効果（実Excel は 図形の 影・光り方。うちは 図形の 線と 影） */
  var 効果たち = [
    { 名: 'すっきり', 線の太さ: 1.5, 影: 0 },
    { 名: 'かげ付き', 線の太さ: 1.5, 影: 4 },
    { 名: 'ふとい線', 線の太さ: 3,   影: 0 },
  ];

  /* ─────────── 色の 濃淡（実測と 15/15 合う） ─────────── */

  function _十六を数に(hex) {
    var s = String(hex).replace('#', '');
    return [parseInt(s.slice(0, 2), 16), parseInt(s.slice(2, 4), 16), parseInt(s.slice(4, 6), 16)];
  }
  function _数を十六に(a) {
    return '#' + a.map(function (v) {
      var n = Math.max(0, Math.min(255, v));
      return (n < 16 ? '0' : '') + n.toString(16).toUpperCase();
    }).join('');
  }

  /** RGB → HLS（0〜240 の 整数。昔からの Windows の やり方） */
  function RGBをHLSへ(R, G, B) {
    var mx = Math.max(R, G, B), mn = Math.min(R, G, B);
    var L = Math.floor((((mx + mn) * HLSMAX) + RGBMAX) / (2 * RGBMAX));
    var H, S;
    if (mx === mn) { S = 0; H = 160; }
    else {
      if (L <= HLSMAX / 2) S = Math.floor((((mx - mn) * HLSMAX) + ((mx + mn) / 2)) / (mx + mn));
      else S = Math.floor((((mx - mn) * HLSMAX) + ((2 * RGBMAX - mx - mn) / 2)) / (2 * RGBMAX - mx - mn));
      var Rd = Math.floor((((mx - R) * (HLSMAX / 6)) + ((mx - mn) / 2)) / (mx - mn));
      var Gd = Math.floor((((mx - G) * (HLSMAX / 6)) + ((mx - mn) / 2)) / (mx - mn));
      var Bd = Math.floor((((mx - B) * (HLSMAX / 6)) + ((mx - mn) / 2)) / (mx - mn));
      if (R === mx) H = Bd - Gd;
      else if (G === mx) H = (HLSMAX / 3) + Rd - Bd;
      else H = ((2 * HLSMAX) / 3) + Gd - Rd;
      if (H < 0) H += HLSMAX;
      if (H > HLSMAX) H -= HLSMAX;
    }
    return [H, L, S];
  }
  function _色合いから(n1, n2, hue) {
    if (hue < 0) hue += HLSMAX;
    if (hue > HLSMAX) hue -= HLSMAX;
    if (hue < HLSMAX / 6) return n1 + Math.floor((((n2 - n1) * hue) + (HLSMAX / 12)) / (HLSMAX / 6));
    if (hue < HLSMAX / 2) return n2;
    if (hue < (HLSMAX * 2) / 3)
      return n1 + Math.floor((((n2 - n1) * (((HLSMAX * 2) / 3) - hue)) + (HLSMAX / 12)) / (HLSMAX / 6));
    return n1;
  }
  /** HLS（0〜240）→ RGB */
  function HLSをRGBへ(H, L, S) {
    if (S === 0) {
      var v = Math.floor((L * RGBMAX) / HLSMAX);
      return [v, v, v];
    }
    var M2 = (L <= HLSMAX / 2)
      ? Math.floor((L * (HLSMAX + S) + HLSMAX / 2) / HLSMAX)
      : L + S - Math.floor((L * S + HLSMAX / 2) / HLSMAX);
    var M1 = 2 * L - M2;
    return [
      Math.floor((_色合いから(M1, M2, H + HLSMAX / 3) * RGBMAX + HLSMAX / 2) / HLSMAX),
      Math.floor((_色合いから(M1, M2, H) * RGBMAX + HLSMAX / 2) / HLSMAX),
      Math.floor((_色合いから(M1, M2, H - HLSMAX / 3) * RGBMAX + HLSMAX / 2) / HLSMAX),
    ];
  }

  /** ★濃淡（TintAndShade）★ … −1〜+1。＋で 薄く・−で 濃く。
   *  ★切り捨て★（実測 15/15。四捨五入だと 7/15 で ずれる） */
  function 濃淡(色, t) {
    /* ★0 は その色の まま★＝RGB→HLS→RGB は ★1 ずれる事が 在る★
       （#156082 が #156182 に なった）。実Excel は 0 なら そのままの 色。 */
    if (!t) return _数を十六に(_十六を数に(色));
    var a = _十六を数に(色);
    var hls = RGBをHLSへ(a[0], a[1], a[2]);
    var H = hls[0], L = hls[1], S = hls[2];
    t = Number(t) || 0;
    if (t > 1) t = 1;
    if (t < -1) t = -1;
    L = (t > 0) ? Math.floor(L * (1 - t) + HLSMAX * t) : Math.floor(L * (1 + t));
    if (L < 0) L = 0;
    if (L > HLSMAX) L = HLSMAX;
    return _数を十六に(HLSをRGBへ(H, L, S));
  }

  /** 配色を 名前で 引く（無ければ 1つ目） */
  function 配色を引く(名) {
    for (var i = 0; i < 配色たち.length; i++) if (配色たち[i].名 === 名) return 配色たち[i];
    return 配色たち[0];
  }

  /** 役割の 名前から 色を 出す（accent1 など） */
  function 役割の色(配色, 役, t) {
    var i = 役割.indexOf(役);
    if (i < 0) return null;
    var c = (配色 && 配色.色) ? 配色.色[i] : 配色たち[0].色[i];
    return (t === undefined || t === 0) ? c : 濃淡(c, t);
  }

  /** その 配色の「よく使う 10段」（実Excel の 色の 一覧と 同じ 並べ方） */
  function 見本の段(配色, 役) {
    var 段 = [0.8, 0.6, 0.4, 0, -0.25, -0.5];
    var 出 = [];
    for (var i = 0; i < 段.length; i++) 出.push(役割の色(配色, 役, 段[i]));
    return 出;
  }

  return {
    役割: 役割, 役割の日本語: 役割の日本語, 実Excelの既定: 実Excelの既定,
    配色たち: 配色たち, フォント組たち: フォント組たち, 効果たち: 効果たち,
    濃淡: 濃淡, RGBをHLSへ: RGBをHLSへ, HLSをRGBへ: HLSをRGBへ,
    配色を引く: 配色を引く, 役割の色: 役割の色, 見本の段: 見本の段,
  };
}));
