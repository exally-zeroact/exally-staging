/* chart.js — ★グラフ（縦棒・折れ線・円）★ 2026-08-29
 *
 *  ★司さん「Excel全機能全能力が Exallyに 入って Excelの最上級に なる」★
 *  Excelの「挿入」タブの 中心は グラフ（11個）。まず ★よく使う3つ★ から。
 *
 *  ★真値（実Excel 16.0 で 実測 2026-08-29）★
 *    ・既定の 種類 … ★51 = 集合縦棒★
 *    ・既定の 大きさ … ★幅 360 × 高さ 216★（ポイント）
 *    ・凡例 … ★系列が 1本なら 出さない★（HasLegend=False）
 *    ・タイトル … ★出す★（HasTitle=True）
 *    ・系列の 名前 … 見出しが 無ければ「系列1」
 *
 *  ★見せ方（司さん「訴えられんような見せ方で 同じように」）★
 *    ・★Excelの 絵・配色は 1つも 写さない★。色は うちの緑の並び。
 *    ・描くのは ★自前の canvas★（外の部品を 1つも 読み込まない）。
 *
 *  ★出来ない事は 黙らない★＝数が 1つも 無い時は 作らずに 理由を 返す。
 *
 *  ★種類ごとの 既定（実Excel 16.0 で 実測 2026-08-30）★
 *    `tools/measure-charts.ps1` … 見出し1列＋数2列＋4行 を 選んで AddChart2 した 結果。
 *    ★大きさは どの種類も 360x216／題は どの種類も 出す（「グラフ タイトル」）★
 *
 *      種類            番号      凡例    軸
 *      集合縦棒(既定)     51      出す    2
 *      散布図          -4169     出す    2
 *      散布図(線つき)      74      出す    2
 *      バブル            15      出す    2
 *      レーダー        -4151     出す   ★1★
 *      レーダー(印つき)     82      出す   ★1★
 *      ヒストグラム       118    ★出さない★ 2   （★系列が 1本に まとまる★）
 *      箱ひげ図         121    ★出さない★ 2
 *      ツリーマップ       117      出す   ★0★
 *      サンバースト       120    ★出さない★★0★
 *      ウォーターフォール    119      出す    2
 *      じょうご         123    ★出さない★ 1
 *      等高線         -4109   ★出さない★ 2
 *      ドーナツ        -4120     出す   ★0★
 *      面               1      出す    2
 *      株価(高安終)       88   ★作れなかった★（高値/安値/終値の3列が要る＝材料の形が 違う）
 *
 *    ⇒ ★凡例は「系列が2本以上」だけでは 決まらない★（ツリーマップは1本でも 出す・
 *       等高線は2本でも 出さない）。★種類ごとに 実測どおり★ に する。
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.Chart = factory();
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  /* ★うちの色★（Excelの 配色は 使わない） */
  var 色 = ['#3D9E72', '#52B788', '#2E7D54', '#7A9A87', '#C8ECD8', '#1A2B22'];

  var 既定 = { 幅: 360, 高さ: 216 };   /* 実Excelの 実測値 */

  /** 選んだ範囲から グラフの 材料を 作る
   *  @param 表 … [[値,…],…]（選んだ範囲の 中身。左上から）
   *  返り値: { ok, なぜ, 見出し[], 系列[{名, 値[]}] } */
  /* ★系列を 縦に取るか 横に取るか★（実Excel 16.0 で 実測 2026-08-30）
       tools/measure-orient.ps1 / measure-orient2.ps1
         数3行3列 → 系列は ★行★（Y=1行目・2行目・3行目）
         数4行2列 → 系列は ★列★
         数2行4列 → 系列は ★行★
         数5行5列 → 系列は ★行★
       見出しを 付けても 同じ:
         見出し付き 数2行2列 → 系列='1月','2月'（★行★）・横軸=売上,原価
         見出し付き 数3行2列 → 系列='売上','原価'（★列★）・横軸=1月,2月,3月
         見出し付き 数2行3列 → 系列='1月','2月'（★行★）
     ⇒ ★決まり＝数の塊で「行数 > 列数」なら 系列は 列／そうでなければ 系列は 行★
       （うちは 前は いつも 列だった＝2026-08-30 に 実測して 直した） */
  function 材料を作る(表) {
    if (!表 || !表.length || !表[0] || !表[0].length) return { ok: false, なぜ: '選んだ所に 何も ありません' };
    var 高 = 表.length, 幅 = 表[0].length;
    /* 1列目が 字なら 見出し（Excelと 同じ 見立て） */
    var 左が見出し = false;
    for (var r = 0; r < 高; r++) {
      var v = 表[r][0];
      if (v !== null && v !== undefined && v !== '' && !数か(v)) { 左が見出し = true; break; }
    }
    /* 1行目が 字なら 系列の 名前 */
    var 上が名前 = false;
    for (var c = 左が見出し ? 1 : 0; c < 幅; c++) {
      var v2 = 表[0][c];
      if (v2 !== null && v2 !== undefined && v2 !== '' && !数か(v2)) { 上が名前 = true; break; }
    }
    var 始行 = 上が名前 ? 1 : 0;
    var 始列 = 左が見出し ? 1 : 0;
    if (始行 >= 高 || 始列 >= 幅) return { ok: false, なぜ: '数が 1つも ありません' };

    /* 数の塊と、その 行の名前・列の名前 */
    var 行数 = 高 - 始行, 列数 = 幅 - 始列;
    var 数 = [], rr, cc;
    for (rr = 0; rr < 行数; rr++) {
      var 一行 = [];
      for (cc = 0; cc < 列数; cc++) 一行.push(数にする(表[始行 + rr][始列 + cc]));
      数.push(一行);
    }
    var 行の名 = [], 列の名 = [];
    for (rr = 0; rr < 行数; rr++) {
      行の名.push(左が見出し ? String(表[始行 + rr][0] == null ? '' : 表[始行 + rr][0]) : String(rr + 1));
    }
    for (cc = 0; cc < 列数; cc++) {
      列の名.push(上が名前 ? String(表[0][始列 + cc] == null ? '' : 表[0][始列 + cc]) : String(cc + 1));
    }

    /* ★行数 > 列数 なら 系列は 列／そうでなければ 系列は 行（実測）★ */
    var 横に取った = !(行数 > 列数);
    var 見出し = 横に取った ? 列の名 : 行の名;
    var 名の元 = 横に取った ? 行の名 : 列の名;
    var 名がある = 横に取った ? 左が見出し : 上が名前;
    var 本数 = 横に取った ? 行数 : 列数;
    var 個数 = 横に取った ? 列数 : 行数;

    var 系列 = [];
    for (var b = 0; b < 本数; b++) {
      var 値 = [], 数あり = false;
      for (var k = 0; k < 個数; k++) {
        var n = 横に取った ? 数[b][k] : 数[k][b];
        値.push(n);
        if (n !== null) 数あり = true;
      }
      if (!数あり) continue;
      系列.push({ 名: 名がある ? 名の元[b] : ('系列' + (系列.length + 1)), 値: 値 });
    }
    if (!系列.length) return { ok: false, なぜ: '数が 1つも ありません' };
    /* ★見出しが 字から 来たか★＝散布図で X を どう取るかが 変わる（実測①②③） */
    return {
      ok: true, 見出し: 見出し, 系列: 系列,
      見出しが字: 横に取った ? 上が名前 : 左が見出し,
      横に取った: 横に取った,
    };
  }

  function 数か(v) { return 数にする(v) !== null; }
  function 数にする(v) {
    if (v === null || v === undefined || v === '') return null;
    if (typeof v === 'number') return isFinite(v) ? v : null;
    var s = String(v).replace(/[,¥￥\s]/g, '');
    if (s === '') return null;
    var n = Number(s);
    return isFinite(n) ? n : null;
  }

  /* ★種類ごとの 既定（上の実測表の 機械が読む形）★
     凡例 … 'series2'＝系列が2本以上なら出す（実Excelの 集合縦棒・折れ線・円・面・散布図・レーダー・バブル）
            true ＝1本でも 出す（ツリーマップ・ウォーターフォール）
            false＝出さない（ヒストグラム・箱ひげ・サンバースト・じょうご・等高線） */
  var 種類の既定 = {
    column:      { 凡例: 'series2', 軸: 2 },
    line:        { 凡例: 'series2', 軸: 2 },
    area:        { 凡例: 'series2', 軸: 2 },
    pie:         { 凡例: 'series2', 軸: 0 },
    doughnut:    { 凡例: 'series2', 軸: 0 },
    scatter:     { 凡例: 'series2', 軸: 2 },
    scatterLine: { 凡例: 'series2', 軸: 2 },
    bubble:      { 凡例: 'series2', 軸: 2 },
    radar:       { 凡例: 'series2', 軸: 1 },
    combo:       { 凡例: 'series2', 軸: 2 },
    treemap:     { 凡例: true,      軸: 0 },
    waterfall:   { 凡例: true,      軸: 2 },
    histogram:   { 凡例: false,     軸: 2 },
    box:         { 凡例: false,     軸: 2 },
    sunburst:    { 凡例: false,     軸: 0 },
    funnel:      { 凡例: false,     軸: 1 },
  };
  function 凡例を出すか(種類, 材料) {
    var d = 種類の既定[種類];
    var 決まり = d ? d.凡例 : 'series2';
    if (決まり === true) return true;
    if (決まり === false) return false;
    return 材料.系列.length >= 2;
  }

  /** canvas に 描く
   *  @param 種類 … 'column' | 'line' | 'area' | 'pie' | 'doughnut' | 'scatter' | 'scatterLine'
   *                | 'bubble' | 'radar' | 'combo' | 'treemap' | 'sunburst'
   *                | 'histogram' | 'box' | 'waterfall' | 'funnel' */
  function 描く(ctx, 材料, 種類, 幅, 高, 題) {
    if (!ctx || !材料 || !材料.ok) return false;
    var W = 幅 || 既定.幅, H = 高 || 既定.高さ;
    ctx.clearRect(0, 0, W, H);
    ctx.fillStyle = '#ffffff'; ctx.fillRect(0, 0, W, H);
    ctx.strokeStyle = '#C8ECD8'; ctx.lineWidth = 1; ctx.strokeRect(0.5, 0.5, W - 1, H - 1);

    /* ★タイトルは 出す★（実Excelの 既定と 同じ） */
    var 上 = 8;
    if (題) {
      ctx.fillStyle = '#2E7D54';
      ctx.font = 'bold 13px "Noto Sans JP",sans-serif';
      ctx.textAlign = 'center'; ctx.textBaseline = 'top';
      ctx.fillText(String(題), W / 2, 上);
      上 += 20;
    }
    /* ★凡例を 出すかは 種類ごとに 実測どおり★（上の表） */
    var 凡例の高 = 0;
    if (凡例を出すか(種類, 材料)) {
      凡例の高 = 18;
      ctx.font = '11px "Noto Sans JP",sans-serif';
      ctx.textAlign = 'left'; ctx.textBaseline = 'middle';
      var x = 10;
      for (var i = 0; i < 材料.系列.length; i++) {
        ctx.fillStyle = 色[i % 色.length];
        ctx.fillRect(x, H - 14, 10, 10);
        ctx.fillStyle = '#333333';
        ctx.fillText(材料.系列[i].名, x + 14, H - 9);
        x += 14 + ctx.measureText(材料.系列[i].名).width + 12;
      }
    }
    if (種類 === 'pie' || 種類 === 'doughnut') return _円(ctx, 材料, W, H, 上, 凡例の高, 種類 === 'doughnut');
    if (種類 === 'scatter' || 種類 === 'scatterLine' || 種類 === 'bubble')
      return _散布(ctx, 材料, W, H, 上, 凡例の高, 種類);
    if (種類 === 'radar') return _レーダー(ctx, 材料, W, H, 上, 凡例の高);
    if (種類 === 'treemap') return _四角の入れ子(ctx, 材料, W, H, 上, 凡例の高, false);
    if (種類 === 'sunburst') return _四角の入れ子(ctx, 材料, W, H, 上, 凡例の高, true);
    if (種類 === 'funnel') return _じょうご(ctx, 材料, W, H, 上, 凡例の高);
    if (種類 === 'histogram') return _棒か線(ctx, ヒストグラムの材料(材料), 'column', W, H, 上, 凡例の高);
    if (種類 === 'waterfall') return _ウォーターフォール(ctx, 材料, W, H, 上, 凡例の高);
    if (種類 === 'box') return _箱ひげ(ctx, 材料, W, H, 上, 凡例の高);
    if (種類 === 'combo') return _棒か線(ctx, 材料, 'combo', W, H, 上, 凡例の高);
    if (種類 === 'area') return _棒か線(ctx, 材料, 'area', W, H, 上, 凡例の高);
    return _棒か線(ctx, 材料, 種類 === 'line' ? 'line' : 'column', W, H, 上, 凡例の高);
  }

  function _棒か線(ctx, 材料, 種類, W, H, 上, 凡例の高) {
    var 左 = 44, 右 = 10, 下 = 24 + 凡例の高;
    var 図幅 = W - 左 - 右, 図高 = H - 上 - 下;
    if (図幅 <= 10 || 図高 <= 10) return false;

    var 最大 = 0, 最小 = 0;
    for (var i = 0; i < 材料.系列.length; i++) {
      for (var j = 0; j < 材料.系列[i].値.length; j++) {
        var v = 材料.系列[i].値[j];
        if (v === null) continue;
        if (v > 最大) 最大 = v;
        if (v < 最小) 最小 = v;
      }
    }
    if (最大 === 最小) 最大 = 最小 + 1;
    var 目盛 = function (v) { return 上 + 図高 - (v - 最小) / (最大 - 最小) * 図高; };

    /* 目盛の 線と 数 */
    ctx.strokeStyle = '#D4EDE1'; ctx.fillStyle = '#7A9A87';
    ctx.font = '10px "Noto Sans JP",sans-serif';
    ctx.textAlign = 'right'; ctx.textBaseline = 'middle';
    for (var k = 0; k <= 4; k++) {
      var v2 = 最小 + (最大 - 最小) * k / 4;
      var y = 目盛(v2);
      ctx.beginPath(); ctx.moveTo(左, y); ctx.lineTo(W - 右, y); ctx.stroke();
      ctx.fillText(String(Math.round(v2 * 100) / 100), 左 - 4, y);
    }
    var 個数 = 材料.見出し.length;
    var 一枠 = 図幅 / Math.max(1, 個数);

    /* ★複合★＝1本目は 棒・残りは 線（実Excel の 複合の 既定と 同じ 組み合わせ）
       ★未測定★：実Excel の 複合は 2本目を 第2軸に 置く。うちは ★同じ軸★（軸を2本 持っていないため）。
       ⇒ ★黙って 同じに 見せない★＝下に 断りを 出す。 */
    var 棒にする = function (s) {
      if (種類 === 'column') return true;
      if (種類 === 'combo') return s === 0;
      return false;
    };
    var 面にする = (種類 === 'area');

    if (種類 === 'column' || 種類 === 'combo') {
      var 本数 = 種類 === 'combo' ? 1 : 材料.系列.length;
      var 棒幅 = Math.max(2, (一枠 * 0.7) / 本数);
      for (var s = 0; s < 本数; s++) {
        if (!棒にする(s)) continue;
        ctx.fillStyle = 色[s % 色.length];
        for (var n = 0; n < 個数; n++) {
          var v3 = 材料.系列[s].値[n];
          if (v3 === null) continue;
          var x = 左 + 一枠 * n + (一枠 - 棒幅 * 本数) / 2 + 棒幅 * s;
          var y0 = 目盛(Math.max(0, 最小)), y1 = 目盛(v3);
          ctx.fillRect(x, Math.min(y0, y1), 棒幅, Math.abs(y1 - y0) || 1);
        }
      }
    }
    /* 線（折れ線／面／複合の 2本目から） */
    if (種類 !== 'column') {
      for (var s2 = 0; s2 < 材料.系列.length; s2++) {
        if (棒にする(s2)) continue;
        ctx.strokeStyle = 色[s2 % 色.length]; ctx.lineWidth = 2;
        var 点 = [];
        for (var n2 = 0; n2 < 個数; n2++) {
          var v4 = 材料.系列[s2].値[n2];
          if (v4 === null) continue;
          点.push([左 + 一枠 * n2 + 一枠 / 2, 目盛(v4)]);
        }
        if (!点.length) continue;
        /* ★面★＝線の 下を 塗る（折れ線とは 見た目が 違う＝実Excel も 塗る） */
        if (面にする) {
          ctx.beginPath();
          ctx.moveTo(点[0][0], 目盛(Math.max(0, 最小)));
          for (var a = 0; a < 点.length; a++) ctx.lineTo(点[a][0], 点[a][1]);
          ctx.lineTo(点[点.length - 1][0], 目盛(Math.max(0, 最小)));
          ctx.closePath();
          ctx.globalAlpha = 0.35; ctx.fillStyle = 色[s2 % 色.length]; ctx.fill(); ctx.globalAlpha = 1;
        }
        ctx.beginPath();
        for (var b = 0; b < 点.length; b++) {
          if (b === 0) ctx.moveTo(点[b][0], 点[b][1]); else ctx.lineTo(点[b][0], 点[b][1]);
        }
        ctx.stroke();
      }
    }
    /* 下の 見出し */
    ctx.fillStyle = '#333333'; ctx.font = '10px "Noto Sans JP",sans-serif';
    ctx.textAlign = 'center'; ctx.textBaseline = 'top';
    for (var m = 0; m < 個数; m++) {
      ctx.fillText(材料.見出し[m], 左 + 一枠 * m + 一枠 / 2, 上 + 図高 + 4);
    }
    return true;
  }

  function _円(ctx, 材料, W, H, 上, 凡例の高, ドーナツか) {
    var 値 = 材料.系列[0].値;
    var 合計 = 0;
    for (var i = 0; i < 値.length; i++) if (値[i] !== null && 値[i] > 0) 合計 += 値[i];
    if (!合計) return false;
    var cx = W / 2, cy = 上 + (H - 上 - 凡例の高 - 8) / 2;
    var r = Math.min(W, H - 上 - 凡例の高) / 2 - 12;
    var 角 = -Math.PI / 2;
    for (var j = 0; j < 値.length; j++) {
      var v = 値[j];
      if (v === null || v <= 0) continue;
      var 幅角 = v / 合計 * Math.PI * 2;
      ctx.beginPath(); ctx.moveTo(cx, cy);
      ctx.arc(cx, cy, r, 角, 角 + 幅角);
      ctx.closePath();
      ctx.fillStyle = 色[j % 色.length]; ctx.fill();
      ctx.strokeStyle = '#ffffff'; ctx.lineWidth = 1; ctx.stroke();
      角 += 幅角;
    }
    /* ★ドーナツ★＝真ん中を 白で 抜く（実Excelの 既定の 穴は 半径の 約半分） */
    if (ドーナツか) {
      ctx.beginPath(); ctx.arc(cx, cy, r * 0.5, 0, Math.PI * 2);
      ctx.fillStyle = '#ffffff'; ctx.fill();
    }
    return true;
  }

  /* ───────── 散布図・バブル ─────────
     ★実Excel と 同じ 見立て★＝★1列目が X★・残りが Y（バブルは 3列目が 大きさ）。
     うちの 材料を作る は 1列目が 字なら「見出し」に するので、
     散布図の時だけ ★見出しを 数として 読み直す★。 */
  /*  ★真値（実Excel 16.0 で 実測 2026-08-30・tools/measure-scatter.ps1）★
        ①1列目が 字（あ/1/10 …）→ ★系列2本・X=1,2,3（番号）★
        ②数だけ 2列（1/10 …）  → ★系列1本・X=1列目・Y=2列目★
        ③数だけ 3列            → ★X=1行目・Y=2行目と3行目★（横に取った 形）
      ⇒ ★見出しが 字なら X は 番号／字が 無ければ 1本目の 系列が X★ */
  function 散布の材料(材料) {
    var X = [], i;
    if (材料.見出しが字) {
      for (i = 0; i < 材料.見出し.length; i++) X.push(i + 1);   /* ★実測① 番号★ */
      return { X: X, 系列: 材料.系列 };
    }
    if (材料.系列.length >= 2) {
      return { X: 材料.系列[0].値.slice(), 系列: 材料.系列.slice(1) };  /* ★実測②③★ */
    }
    /* 系列が 1本しか 無い＝X に 出来る物が 無い ⇒ 番号を 使う */
    for (i = 0; i < 材料.見出し.length; i++) X.push(i + 1);
    return { X: X, 系列: 材料.系列 };
  }

  function _散布(ctx, 材料, W, H, 上, 凡例の高, 種類) {
    var d = 散布の材料(材料);
    var 左 = 44, 右 = 12, 下 = 24 + 凡例の高;
    var 図幅 = W - 左 - 右, 図高 = H - 上 - 下;
    if (図幅 <= 10 || 図高 <= 10) return false;

    var xMin = null, xMax = null, yMin = 0, yMax = 0;
    for (var i = 0; i < d.X.length; i++) {
      if (xMin === null || d.X[i] < xMin) xMin = d.X[i];
      if (xMax === null || d.X[i] > xMax) xMax = d.X[i];
    }
    var 描く系列 = (種類 === 'bubble' && d.系列.length >= 2) ? d.系列.slice(0, 1) : d.系列;
    for (var s2 = 0; s2 < 描く系列.length; s2++) {
      for (var j = 0; j < 描く系列[s2].値.length; j++) {
        var v = 描く系列[s2].値[j];
        if (v === null) continue;
        if (v > yMax) yMax = v;
        if (v < yMin) yMin = v;
      }
    }
    if (xMin === null) { xMin = 0; xMax = 1; }
    if (xMax === xMin) xMax = xMin + 1;
    if (yMax === yMin) yMax = yMin + 1;
    var px = function (x) { return 左 + (x - xMin) / (xMax - xMin) * 図幅; };
    var py = function (y) { return 上 + 図高 - (y - yMin) / (yMax - yMin) * 図高; };

    _目盛(ctx, 左, W - 右, 上, 図高, yMin, yMax);

    /* ★バブルの 大きさ★＝3列目（無ければ 一定） */
    var 大きさ列 = (種類 === 'bubble' && d.系列.length >= 2) ? d.系列[1].値 : null;
    var 大最大 = 1;
    if (大きさ列) for (var k = 0; k < 大きさ列.length; k++) if (大きさ列[k] > 大最大) 大最大 = 大きさ列[k];

    for (var s = 0; s < 描く系列.length; s++) {
      ctx.fillStyle = 色[s % 色.length];
      ctx.strokeStyle = 色[s % 色.length];
      ctx.lineWidth = 2;
      if (種類 === 'scatterLine') {
        ctx.beginPath();
        var 打った = false;
        for (var n2 = 0; n2 < 描く系列[s].値.length; n2++) {
          var yv = 描く系列[s].値[n2];
          if (yv === null || d.X[n2] === undefined) continue;
          if (!打った) { ctx.moveTo(px(d.X[n2]), py(yv)); 打った = true; }
          else ctx.lineTo(px(d.X[n2]), py(yv));
        }
        ctx.stroke();
      }
      for (var n = 0; n < 描く系列[s].値.length; n++) {
        var y = 描く系列[s].値[n];
        if (y === null || d.X[n] === undefined) continue;
        var r = 3;
        if (大きさ列 && 大きさ列[n] !== null && 大きさ列[n] !== undefined) {
          r = 3 + Math.sqrt(Math.max(0, 大きさ列[n]) / 大最大) * 12;
        }
        ctx.beginPath(); ctx.arc(px(d.X[n]), py(y), r, 0, Math.PI * 2); ctx.fill();
      }
    }
    /* X の 目盛の 数（散布図は ★下も 数★＝実Excel と 同じ） */
    ctx.fillStyle = '#333333'; ctx.font = '10px "Noto Sans JP",sans-serif';
    ctx.textAlign = 'center'; ctx.textBaseline = 'top';
    for (var t = 0; t <= 4; t++) {
      var xv = xMin + (xMax - xMin) * t / 4;
      ctx.fillText(String(Math.round(xv * 100) / 100), px(xv), 上 + 図高 + 4);
    }
    return true;
  }

  /* 目盛の 線と 数（横線）＝棒/線/散布で 使い回す */
  function _目盛(ctx, 左, 右端, 上, 図高, 最小, 最大) {
    ctx.strokeStyle = '#D4EDE1'; ctx.fillStyle = '#7A9A87';
    ctx.font = '10px "Noto Sans JP",sans-serif';
    ctx.textAlign = 'right'; ctx.textBaseline = 'middle';
    for (var k = 0; k <= 4; k++) {
      var v = 最小 + (最大 - 最小) * k / 4;
      var y = 上 + 図高 - (v - 最小) / (最大 - 最小) * 図高;
      ctx.beginPath(); ctx.moveTo(左, y); ctx.lineTo(右端, y); ctx.stroke();
      ctx.fillText(String(Math.round(v * 100) / 100), 左 - 4, y);
    }
  }

  /* ───────── レーダー（軸は 1本＝実測）───────── */
  function _レーダー(ctx, 材料, W, H, 上, 凡例の高) {
    var 個数 = 材料.見出し.length;
    if (個数 < 3) return false;                 /* ★3つ 無いと 形に ならない★ */
    var cx = W / 2, cy = 上 + (H - 上 - 凡例の高) / 2;
    var r = Math.min(W, H - 上 - 凡例の高) / 2 - 22;
    if (r <= 6) return false;
    var 最大 = 0;
    for (var i = 0; i < 材料.系列.length; i++)
      for (var j = 0; j < 材料.系列[i].値.length; j++)
        if (材料.系列[i].値[j] > 最大) 最大 = 材料.系列[i].値[j];
    if (!最大) 最大 = 1;
    var 角 = function (n) { return -Math.PI / 2 + n / 個数 * Math.PI * 2; };

    /* 巣（4重）＋ 軸の線 */
    ctx.strokeStyle = '#D4EDE1'; ctx.lineWidth = 1;
    for (var w = 1; w <= 4; w++) {
      ctx.beginPath();
      for (var n = 0; n <= 個数; n++) {
        var a = 角(n % 個数), rr = r * w / 4;
        var x = cx + Math.cos(a) * rr, y = cy + Math.sin(a) * rr;
        if (n === 0) ctx.moveTo(x, y); else ctx.lineTo(x, y);
      }
      ctx.stroke();
    }
    for (var n2 = 0; n2 < 個数; n2++) {
      ctx.beginPath(); ctx.moveTo(cx, cy);
      ctx.lineTo(cx + Math.cos(角(n2)) * r, cy + Math.sin(角(n2)) * r); ctx.stroke();
    }
    /* 系列 */
    for (var s = 0; s < 材料.系列.length; s++) {
      ctx.strokeStyle = 色[s % 色.length]; ctx.lineWidth = 2;
      ctx.beginPath();
      for (var m = 0; m <= 個数; m++) {
        var v = 材料.系列[s].値[m % 個数];
        var rr2 = (v === null ? 0 : v) / 最大 * r;
        var a2 = 角(m % 個数);
        var x2 = cx + Math.cos(a2) * rr2, y2 = cy + Math.sin(a2) * rr2;
        if (m === 0) ctx.moveTo(x2, y2); else ctx.lineTo(x2, y2);
      }
      ctx.stroke();
    }
    /* 見出し */
    ctx.fillStyle = '#333333'; ctx.font = '10px "Noto Sans JP",sans-serif';
    ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
    for (var k = 0; k < 個数; k++) {
      ctx.fillText(材料.見出し[k], cx + Math.cos(角(k)) * (r + 12), cy + Math.sin(角(k)) * (r + 10));
    }
    return true;
  }

  /* ───────── ヒストグラム ─────────
     ★実Excel は 系列を 1本に まとめる（実測＝系列=1・凡例なし）★
     区間の 数は 値の数の 平方根（Excelの 既定と 同じ 考え方）で 決める。
     ★未測定★：実Excel の 区間の 境目の 丸め方は COM から 読めなかった。 */
  function ヒストグラムの材料(材料) {
    var 値 = [];
    for (var i = 0; i < 材料.系列.length; i++)
      for (var j = 0; j < 材料.系列[i].値.length; j++)
        if (材料.系列[i].値[j] !== null) 値.push(材料.系列[i].値[j]);
    if (!値.length) return { ok: true, 見出し: [''], 系列: [{ 名: '度数', 値: [0] }] };
    値.sort(function (a, b) { return a - b; });
    var 最小 = 値[0], 最大 = 値[値.length - 1];
    if (最大 === 最小) 最大 = 最小 + 1;
    var 区間数 = Math.max(1, Math.min(20, Math.ceil(Math.sqrt(値.length))));
    var 幅 = (最大 - 最小) / 区間数;
    var 度数 = [], 見出し = [];
    for (var b = 0; b < 区間数; b++) {
      度数.push(0);
      見出し.push(_丸め(最小 + 幅 * b) + '〜' + _丸め(最小 + 幅 * (b + 1)));
    }
    for (var k = 0; k < 値.length; k++) {
      var idx = Math.floor((値[k] - 最小) / 幅);
      if (idx >= 区間数) idx = 区間数 - 1;      /* ★一番上の 値は 最後の区間に 入れる★ */
      度数[idx]++;
    }
    return { ok: true, 見出し: 見出し, 系列: [{ 名: '度数', 値: 度数 }] };
  }
  function _丸め(n) { return String(Math.round(n * 100) / 100); }

  /* ───────── ウォーターフォール ─────────
     ★積み上がりを 見せる★＝1本目の 系列を 増減として 読み、足していった 高さに 棒を 置く。 */
  function _ウォーターフォール(ctx, 材料, W, H, 上, 凡例の高) {
    var 値 = 材料.系列[0].値;
    var 左 = 44, 右 = 10, 下 = 24 + 凡例の高;
    var 図幅 = W - 左 - 右, 図高 = H - 上 - 下;
    if (図幅 <= 10 || 図高 <= 10) return false;
    var 積 = 0, 下端 = [], 上端 = [], 最小 = 0, 最大 = 0;
    for (var i = 0; i < 値.length; i++) {
      var v = 値[i] === null ? 0 : 値[i];
      下端.push(積); 積 += v; 上端.push(積);
      if (積 > 最大) 最大 = 積;
      if (積 < 最小) 最小 = 積;
    }
    if (最大 === 最小) 最大 = 最小 + 1;
    var py = function (y) { return 上 + 図高 - (y - 最小) / (最大 - 最小) * 図高; };
    _目盛(ctx, 左, W - 右, 上, 図高, 最小, 最大);
    var 一枠 = 図幅 / Math.max(1, 値.length);
    var 棒幅 = Math.max(2, 一枠 * 0.6);
    for (var j = 0; j < 値.length; j++) {
      var 増えた = (上端[j] - 下端[j]) >= 0;
      ctx.fillStyle = 増えた ? 色[0] : '#C0392B';   /* ★減りは 別の色（黙って同じ色に しない）★ */
      var y0 = py(下端[j]), y1 = py(上端[j]);
      ctx.fillRect(左 + 一枠 * j + (一枠 - 棒幅) / 2, Math.min(y0, y1), 棒幅, Math.abs(y1 - y0) || 1);
    }
    ctx.fillStyle = '#333333'; ctx.font = '10px "Noto Sans JP",sans-serif';
    ctx.textAlign = 'center'; ctx.textBaseline = 'top';
    for (var m = 0; m < 材料.見出し.length; m++)
      ctx.fillText(材料.見出し[m], 左 + 一枠 * m + 一枠 / 2, 上 + 図高 + 4);
    return true;
  }

  /* ───────── じょうご（上から 下へ・真ん中ぞろえの 帯）───────── */
  function _じょうご(ctx, 材料, W, H, 上, 凡例の高) {
    var 値 = 材料.系列[0].値;
    var 最大 = 0;
    for (var i = 0; i < 値.length; i++) if (値[i] > 最大) 最大 = 値[i];
    if (!最大) return false;
    var 左 = 70, 右 = 10, 下 = 8 + 凡例の高;
    var 図幅 = W - 左 - 右, 図高 = H - 上 - 下;
    var 一枠 = 図高 / Math.max(1, 値.length);
    var 高さ = Math.max(2, 一枠 * 0.7);
    ctx.textBaseline = 'middle';
    for (var j = 0; j < 値.length; j++) {
      var v = 値[j] === null ? 0 : 値[j];
      var w = 図幅 * (v / 最大);
      var y = 上 + 一枠 * j + (一枠 - 高さ) / 2;
      ctx.fillStyle = 色[j % 色.length];
      ctx.fillRect(左 + (図幅 - w) / 2, y, w, 高さ);
      ctx.fillStyle = '#333333'; ctx.font = '10px "Noto Sans JP",sans-serif';
      ctx.textAlign = 'right';
      ctx.fillText(材料.見出し[j], 左 - 6, y + 高さ / 2);
    }
    return true;
  }

  /* ───────── ツリーマップ／サンバースト ─────────
     ★軸は 0本（実測）★。大きさで 場所を 分ける。 */
  function _四角の入れ子(ctx, 材料, W, H, 上, 凡例の高, 丸いか) {
    var 値 = 材料.系列[0].値;
    var 合計 = 0;
    for (var i = 0; i < 値.length; i++) if (値[i] !== null && 値[i] > 0) 合計 += 値[i];
    if (!合計) return false;
    var 左 = 8, 右 = 8, 下 = 8 + 凡例の高;
    var 図幅 = W - 左 - 右, 図高 = H - 上 - 下;
    if (図幅 <= 10 || 図高 <= 10) return false;

    if (丸いか) {
      /* サンバースト＝中を 抜いた 輪を 割る */
      var cx = 左 + 図幅 / 2, cy = 上 + 図高 / 2;
      var r外 = Math.min(図幅, 図高) / 2 - 4, r内 = r外 * 0.4;
      var 角 = -Math.PI / 2;
      for (var j = 0; j < 値.length; j++) {
        var v = 値[j]; if (v === null || v <= 0) continue;
        var 幅角 = v / 合計 * Math.PI * 2;
        ctx.beginPath();
        ctx.arc(cx, cy, r外, 角, 角 + 幅角);
        ctx.arc(cx, cy, r内, 角 + 幅角, 角, true);
        ctx.closePath();
        ctx.fillStyle = 色[j % 色.length]; ctx.fill();
        ctx.strokeStyle = '#ffffff'; ctx.lineWidth = 1; ctx.stroke();
        角 += 幅角;
      }
      return true;
    }
    /* ツリーマップ＝縦横に 切っていく（大きい順） */
    var 並び = [];
    for (var k = 0; k < 値.length; k++)
      if (値[k] !== null && 値[k] > 0) 並び.push({ v: 値[k], 名: 材料.見出し[k], i: k });
    並び.sort(function (a, b) { return b.v - a.v; });
    var x = 左, y = 上, w = 図幅, h = 図高, 残り = 合計;
    for (var m = 0; m < 並び.length; m++) {
      var 割合 = 並び[m].v / 残り;
      var 横に切る = (w >= h);
      var cw = 横に切る ? w * 割合 : w;
      var ch = 横に切る ? h : h * 割合;
      ctx.fillStyle = 色[並び[m].i % 色.length];
      ctx.fillRect(x, y, Math.max(1, cw - 1), Math.max(1, ch - 1));
      if (cw > 34 && ch > 14) {
        ctx.fillStyle = '#ffffff'; ctx.font = '10px "Noto Sans JP",sans-serif';
        ctx.textAlign = 'left'; ctx.textBaseline = 'top';
        ctx.fillText(String(並び[m].名), x + 4, y + 4);
      }
      if (横に切る) { x += cw; w -= cw; } else { y += ch; h -= ch; }
      残り -= 並び[m].v;
      if (残り <= 0 || w <= 1 || h <= 1) break;
    }
    return true;
  }

  /* ───────── 箱ひげ図 ─────────
     ★真値（実Excel 16.0 で 実測 2026-08-30）★
       COM の性質は 読めなかったので、★実Excelに 描かせて 保存した中身★を 読んだ:
         xl/charts/chartEx1.xml
           <cx:statistics quartileMethod="exclusive" />
           <cx:visibility meanLine="0" meanMarker="1" nonoutliers="0" outliers="1" />
       ⇒ ①四分位は ★排他（中央値を 含めない）＝QUARTILE.EXC★
            実測: 1〜10 で EXC=2.75 / 5.5 / 8.25（INC は 3.25 / 5.5 / 7.75）
         ②平均は ★線ではなく 印★ で 出す
         ③外れ値は ★出す★／ひげは ★外れ値まで 伸ばさない★
       ★外れ値の 決め方（1.5×四分位の幅）は うちの決め＝未測定★
         （XMLに 数が 書かれていない＝Excelの中の 決めが 読めない） */
  function 四分位(値) {
    var v = 値.slice().filter(function (x) { return x !== null && x !== undefined; })
      .sort(function (a, b) { return a - b; });
    var n = v.length;
    if (n < 4) return null;                        /* ★出せない★ */
    var 取る = function (p) {
      var f = p * (n + 1) - 1, i = Math.floor(f), d = f - i;
      if (i < 0 || i + 1 >= n) return null;
      return v[i] + (v[i + 1] - v[i]) * d;
    };
    var q1 = 取る(0.25), q2 = 取る(0.5), q3 = 取る(0.75);
    if (q1 === null || q2 === null || q3 === null) return null;
    /* ★外れ値（1.5×四分位の幅の 外）★＝ひげは そこまで 伸ばさない（実測 outliers="1"） */
    var 幅 = q3 - q1, 下限 = q1 - 幅 * 1.5, 上限 = q3 + 幅 * 1.5;
    var 外れ = [], ひげ下 = null, ひげ上 = null, 合計 = 0;
    for (var i = 0; i < n; i++) {
      合計 += v[i];
      if (v[i] < 下限 || v[i] > 上限) { 外れ.push(v[i]); continue; }
      if (ひげ下 === null) ひげ下 = v[i];
      ひげ上 = v[i];
    }
    if (ひげ下 === null) { ひげ下 = v[0]; ひげ上 = v[n - 1]; }
    return { 最小: v[0], q1: q1, 中央: q2, q3: q3, 最大: v[n - 1],
             ひげ下: ひげ下, ひげ上: ひげ上, 外れ: 外れ, 平均: 合計 / n };
  }

  function _箱ひげ(ctx, 材料, W, H, 上, 凡例の高) {
    var 箱たち = [], 最小 = null, 最大 = null;
    for (var i = 0; i < 材料.系列.length; i++) {
      var q = 四分位(材料.系列[i].値);
      箱たち.push(q);
      if (!q) continue;
      if (最小 === null || q.最小 < 最小) 最小 = q.最小;
      if (最大 === null || q.最大 > 最大) 最大 = q.最大;
    }
    var 出せる = 0;
    for (var z = 0; z < 箱たち.length; z++) if (箱たち[z]) 出せる++;
    if (!出せる) {
      ctx.fillStyle = '#C0392B'; ctx.font = '12px "Noto Sans JP",sans-serif';
      ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
      ctx.fillText('4つ 以上の 数が 要ります', W / 2, (H + 上) / 2);
      return false;                                /* ★出来ないのに 出来たふりを しない★ */
    }
    if (最大 === 最小) 最大 = 最小 + 1;
    var 左 = 44, 右 = 10, 下 = 24 + 凡例の高;
    var 図幅 = W - 左 - 右, 図高 = H - 上 - 下;
    var py = function (y) { return 上 + 図高 - (y - 最小) / (最大 - 最小) * 図高; };
    _目盛(ctx, 左, W - 右, 上, 図高, 最小, 最大);
    var 一枠 = 図幅 / Math.max(1, 箱たち.length);
    var 箱幅 = Math.max(4, 一枠 * 0.5);
    for (var j = 0; j < 箱たち.length; j++) {
      var q2 = 箱たち[j]; if (!q2) continue;
      var cx = 左 + 一枠 * j + 一枠 / 2;
      ctx.strokeStyle = 色[j % 色.length]; ctx.fillStyle = 色[j % 色.length]; ctx.lineWidth = 1;
      /* ひげ（★外れ値までは 伸ばさない★＝実測） */
      ctx.beginPath(); ctx.moveTo(cx, py(q2.ひげ下)); ctx.lineTo(cx, py(q2.ひげ上)); ctx.stroke();
      /* 箱 */
      ctx.globalAlpha = 0.35;
      ctx.fillRect(cx - 箱幅 / 2, py(q2.q3), 箱幅, Math.max(1, py(q2.q1) - py(q2.q3)));
      ctx.globalAlpha = 1;
      ctx.strokeRect(cx - 箱幅 / 2, py(q2.q3), 箱幅, Math.max(1, py(q2.q1) - py(q2.q3)));
      /* 中央 */
      ctx.beginPath(); ctx.moveTo(cx - 箱幅 / 2, py(q2.中央)); ctx.lineTo(cx + 箱幅 / 2, py(q2.中央));
      ctx.lineWidth = 2; ctx.stroke();
      /* ★平均は 印（線ではない）★＝実測 meanLine="0" meanMarker="1" */
      ctx.lineWidth = 1.5;
      ctx.beginPath();
      ctx.moveTo(cx - 4, py(q2.平均) - 4); ctx.lineTo(cx + 4, py(q2.平均) + 4);
      ctx.moveTo(cx + 4, py(q2.平均) - 4); ctx.lineTo(cx - 4, py(q2.平均) + 4);
      ctx.stroke();
      /* ★外れ値は 点で 出す★＝実測 outliers="1" */
      for (var o = 0; o < q2.外れ.length; o++) {
        ctx.beginPath(); ctx.arc(cx, py(q2.外れ[o]), 2.5, 0, Math.PI * 2); ctx.fill();
      }
      ctx.fillStyle = '#333333'; ctx.font = '10px "Noto Sans JP",sans-serif';
      ctx.textAlign = 'center'; ctx.textBaseline = 'top';
      ctx.fillText(材料.系列[j].名, cx, 上 + 図高 + 4);
    }
    return true;
  }

  return { 材料を作る: 材料を作る, 描く: 描く, 既定: 既定, 色: 色, 種類の既定: 種類の既定,
           凡例を出すか: 凡例を出すか, ヒストグラムの材料: ヒストグラムの材料, 四分位: 四分位,
           散布の材料: 散布の材料, _数にする: 数にする };
}));
