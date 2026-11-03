/* formula-yosoku.js — ★予測・統計・単位★（2026-09-07）
 *
 *  ★★どこから 来たか★★
 *    `docs/measured/exally-missing-2026-09-07.txt`（実Excel に 在って まだ 動かない 43個）
 *    ⇒ そこから ★外へ 出ない・決めが 要らない★物を 測って 作った。
 *
 *  ★★答えは 全部 実Excel に 打たせて 取った★★
 *    docs/measured/kansuu46/golden-yosoku-2026-09-07.tsv
 *    docs/measured/kansuu46/golden-convert-2026-09-07.tsv（単位の 係数 243本）
 *    docs/measured/kansuu46/golden-convert3-2026-09-07.tsv（接頭辞が 付く 単位 147本）
 *    Excel 16.0 build 20326（UI 1041）
 *
 *  ★★CONVERT は「無かった」のでは なく「間違っていた」★★
 *    本番の JS層は ★答えを 4桁で 丸めていました★
 *      =CONVERT(1,"lbm","kg") … ★0.4536★（実Excel は ★0.45359237★）
 *      =CONVERT(1,"kibyte","byte") … ★#N/A★（実Excel は ★1024★）
 *    ⇒★「動かない」より 見つけにくい＝★それらしい 数が 出る★★
 *    ⇒ 単位の 表は ★手で 書かない★（lib/tanni-hyou.js を 実測から 機械で 作る）
 *
 *  ★ここは 純粋な 計算だけ★（エンジンに 繋ぐのは lib/formula-yosoku-plug.js）
 *
 *  見張り: tests/formula-yosoku.test.mjs
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory(require('./tanni-hyou.js'));
  } else {
    root.FormulaYosoku = factory(root.TanniHyou);
  }
}(typeof self !== 'undefined' ? self : this, function (H) {
  'use strict';

  var 誤り = function (種) { return { 誤り: 種 }; };

  function 表にする(v) {
    if (!Array.isArray(v)) return [[v]];
    if (!v.length) return [[]];
    if (!Array.isArray(v[0])) return v.map(function (x) { return [x]; });
    return v.map(function (r) { return r.slice(); });
  }
  function 平らに(v) {
    var 出 = [], A = 表にする(v);
    for (var r = 0; r < A.length; r++) for (var c = 0; c < A[r].length; c++) 出.push(A[r][c]);
    return 出;
  }
  function 数だけ(v) {
    return 平らに(v).map(Number).filter(function (x) { return isFinite(x); });
  }

  /* ════════ 単位を 変える（CONVERT） ════════ */

  /** ★単位の 名前を 読む★＝接頭辞が 付いているかを 見る
   *  ★大文字小文字は 区別する★（実測 … "KM" も "Km" も #N/A） */
  function 単位を読む(名) {
    if (H.単位[名]) return { 倍: H.単位[名].倍, 種: H.単位[名].種 };
    var i, 素;
    /* 2進の 接頭辞（★情報の 単位にしか 付かない★＝実測 "Yim" は #N/A） */
    var 二 = Object.keys(H.二進接頭).sort(function (a, b) { return b.length - a.length; });
    for (i = 0; i < 二.length; i++) {
      if (名.indexOf(二[i]) !== 0) continue;
      素 = 名.slice(二[i].length);
      if (H.単位[素] && H.単位[素].種 === 'bit') {
        return { 倍: H.二進接頭[二[i]] * H.単位[素].倍, 種: H.単位[素].種 };
      }
    }
    /* 10進の 接頭辞（★付く 単位は 実測で 決まっている★） */
    var 十 = Object.keys(H.接頭).sort(function (a, b) { return b.length - a.length; });
    for (i = 0; i < 十.length; i++) {
      if (名.indexOf(十[i]) !== 0) continue;
      素 = 名.slice(十[i].length);
      var u = H.単位[素];
      if (u && u.頭) {
        /* ★面積は 2乗・体積は 3乗★（実測 … km2 は 1,000,000 倍） */
        return { 倍: Math.pow(H.接頭[十[i]], u.頭) * u.倍, 種: u.種 };
      }
    }
    return null;
  }
  function 温度を読む(名) {
    if (H.温度[名]) return { 傾き: H.温度[名].傾き, ずれ: H.温度[名].ずれ };
    var 十 = Object.keys(H.接頭).sort(function (a, b) { return b.length - a.length; });
    for (var i = 0; i < 十.length; i++) {
      if (名.indexOf(十[i]) !== 0) continue;
      var 素 = 名.slice(十[i].length);
      var t = H.温度[素];
      /* ★1 kK ＝ 1000 K★＝目盛りが 粗く なる ⇒ 傾きも ずれも ★割る★
         （掛けると 1kK が 0.001K に なって 逆さまに なる＝実測で 踏んだ） */
      if (t && t.頭) return { 傾き: t.傾き / H.接頭[十[i]], ずれ: t.ずれ / H.接頭[十[i]] };
    }
    return null;
  }
  /** ★単位を 変える★
   *  実測 … 数で ない 物は #VALUE!／空は #N/A／種類が 違えば #N/A */
  function 単位を変える(数, 元, 先) {
    if (typeof 数 === 'boolean' || (typeof 数 === 'string' && 数 !== '' && !isFinite(Number(数)))) {
      return 誤り('VALUE');
    }
    if (数 === null || 数 === undefined || 数 === '') return 誤り('NA');
    var n = Number(数);
    if (!isFinite(n)) return 誤り('VALUE');
    元 = String(元); 先 = String(先);
    var a = 温度を読む(元), b = 温度を読む(先);
    if (a || b) {
      if (!a || !b) return 誤り('NA');
      var C = (n - a.ずれ) / a.傾き;          /* まず セ氏に する */
      return C * b.傾き + b.ずれ;
    }
    var x = 単位を読む(元), y = 単位を読む(先);
    if (!x || !y || x.種 !== y.種) return 誤り('NA');
    return n * x.倍 / y.倍;
  }

  /* ════════ 最小二乗（TREND / GROWTH / LOGEST の 土台） ════════ */

  /** ★連立方程式を 解く★（掃き出し法） */
  /* ★主が 立たない 列は「0 に して 落とす」★（実Excel と 同じ）
     ＝2026-09-14 実測。点が 1つ だけの `=LINEST(2)` は 実Excel ★0★／
       うちは ★#VALUE!★ を 返して いた（紙 golden-346:1758 ほか ★22本★）。
       傾きが 決まらない 時 実Excel は ★その 係数を 0★ と し、残りを 出す。
     ★落ちた★ … 落とした 列の 番号を 入れて 返す（階数＝列数 − 落ちた数）*/
  function 解く(A, b, 落ちた) {
    var n = b.length, i, j, k;
    var M = A.map(function (r, ri) { return r.concat([b[ri]]); });
    /* ★大きさの 目安★＝この 中の 一番 大きい 値。これに 対して 小さい 主は「立たない」
       ★なぜ 割合（相対）で 見るか★＝元の 数が 大きくても 小さくても 同じ 判断に する為。
         前は ★絶対値 1e-300★ で 見て いた＝材料の 桁が 変われば 判断も 変わる。
         repo の 前例 … `exally-formula.js` の `_xlRound` が `Math.abs(v) * 1e-12`（相対）／
                        この ファイルの 369行も `Math.abs(sum) * 1e-17`（相対）
       ★★1e-12 の 根拠＝振って 測った★★（2026-09-14／答えの 紙 746本を 13通りで 通した）
         境 = 0 / 1e-300 / 1e-20 / 1e-18 / 1e-15 / 1e-13 / 1e-12 / 1e-11 /
              1e-9 / 1e-7 / 1e-6 / 1e-3 … ★どれも 同じ 714一致★
         境 = ★1e-2★ … ★711一致（LINEST が 3本 崩れた）★
         ⇒★安全なのは 0〜1e-3。壊れ始めるのは 1e-2★
         ⇒ 1e-12 は ★壊れ目より 9桁 下★＝真ん中より 十分 下側に 取って 在る
       ★正直に★＝この 746本は 0〜1e-3 を ★区別できません★。
         つまり「746本 緑」は 1e-12 でなければ ならない 証拠では ない。
         ★1e-12 を 選んだ 訳は「repo が 既に 使って いる 相対の 値と 揃える」★ */
    var 目安 = 0;
    for (i = 0; i < n; i++) for (j = 0; j < n; j++) if (Math.abs(M[i][j]) > 目安) 目安 = Math.abs(M[i][j]);
    var 境 = (目安 > 0 ? 目安 : 1) * 1e-12;
    var 段 = 0;                          /* 次に 使う 段 */
    var 主の段 = new Array(n).fill(-1);  /* 列 i の 主は どの 段か */
    for (i = 0; i < n; i++) {
      var p = -1, 最大 = 0;
      for (j = 段; j < n; j++) if (Math.abs(M[j][i]) > 最大) { 最大 = Math.abs(M[j][i]); p = j; }
      if (p < 0 || 最大 <= 境) { if (落ちた) 落ちた.push(i); continue; }   /* ★立たない＝落とす★ */
      var t = M[段]; M[段] = M[p]; M[p] = t;
      for (j = i + 1; j <= n; j++) M[段][j] /= M[段][i];
      M[段][i] = 1;
      for (k = 0; k < n; k++) {
        if (k === 段 || M[k][i] === 0) continue;
        var f = M[k][i];
        for (j = i; j <= n; j++) M[k][j] -= f * M[段][j];
      }
      主の段[i] = 段;
      段++;
    }
    var 出 = new Array(n).fill(0);       /* ★落とした 列は 0★ */
    for (i = 0; i < n; i++) if (主の段[i] >= 0) 出[i] = M[主の段[i]][n];
    return 出;
  }

  /** ★回帰★ y = b0 + b1·x1 + … （切片を 使うかは const で 決める）
   *  @returns {{係数:number[], 統計:object}|null} 係数[0]=切片 */
  function 回帰(ys, xs, 切片あり) {
    var n = ys.length, 次 = xs[0].length, i, j, k;
    var 列 = 次 + (切片あり ? 1 : 0);
    if (n < 1 || 列 < 1) return null;
    var X = [];
    for (i = 0; i < n; i++) {
      var r = 切片あり ? [1] : [];
      for (j = 0; j < 次; j++) r.push(xs[i][j]);
      X.push(r);
    }
    var A = [], b = [];
    for (i = 0; i < 列; i++) {
      A.push(new Array(列).fill(0));
      b.push(0);
      for (k = 0; k < n; k++) b[i] += X[k][i] * ys[k];
    }
    for (i = 0; i < 列; i++) for (j = 0; j < 列; j++) {
      var s = 0;
      for (k = 0; k < n; k++) s += X[k][i] * X[k][j];
      A[i][j] = s;
    }
    var 落ちた = [];
    var 解 = 解く(A, b, 落ちた);
    if (!解) return null;
    var 階数 = 列 - 落ちた.length;        /* ★立った 列の 数★ */
    var 係数 = 切片あり ? 解 : [0].concat(解);

    /* ★統計★（LINEST・LOGEST の 3〜5行目） */
    var 予測 = [], 残差二乗 = 0, i2;
    for (i2 = 0; i2 < n; i2++) {
      var v = 切片あり ? 解[0] : 0;
      for (j = 0; j < 次; j++) v += 解[切片あり ? j + 1 : j] * xs[i2][j];
      予測.push(v);
      残差二乗 += (ys[i2] - v) * (ys[i2] - v);
    }
    var 平均 = ys.reduce(function (a2, b2) { return a2 + b2; }, 0) / n;
    var 全二乗 = 0;
    for (i2 = 0; i2 < n; i2++) {
      var d = 切片あり ? (ys[i2] - 平均) : ys[i2];
      全二乗 += d * d;
    }
    var 回帰二乗 = 全二乗 - 残差二乗;
    var 自由度 = n - 階数;               /* ★落とした 列は 使って いない＝階数で 引く★ */
    var 分散 = 自由度 > 0 ? 残差二乗 / 自由度 : 0;
    /* 係数の 標準誤差＝(X'X)^-1 の 対角 × 分散 */
    var 逆 = [];
    for (i = 0; i < 列; i++) {
      var e = new Array(列).fill(0); e[i] = 1;
      var c = 解く(A.map(function (r2) { return r2.slice(); }), e);
      逆.push(c ? c[i] : 0);
    }
    var 標準誤差 = 逆.map(function (v2) { return Math.sqrt(Math.max(0, v2 * 分散)); });
    return {
      係数: 係数,
      切片あり: 切片あり,
      統計: {
        決定係数: 全二乗 > 0 ? 回帰二乗 / 全二乗 : 0,
        標準誤差y: Math.sqrt(分散),
        F: (分散 > 0 && 次 > 0) ? (回帰二乗 / 次) / 分散 : 0,
        自由度: 自由度,
        回帰二乗和: 回帰二乗,
        残差二乗和: 残差二乗,
        係数の標準誤差: 標準誤差
      }
    };
  }

  function xを整える(ys, xs) {
    var n = ys.length, i;
    if (!xs || !xs.length) {
      var 出 = [];
      for (i = 0; i < n; i++) 出.push([i + 1]);
      return 出;
    }
    var A = 表にする(xs);
    /* 縦1列 か 横1行 か 表 */
    if (A.length === n && A[0].length >= 1) return A.map(function (r) { return r.map(Number); });
    if (A.length === 1 && A[0].length === n) return A[0].map(function (v) { return [Number(v)]; });
    var 平 = 平らに(A).map(Number);
    if (平.length === n) return 平.map(function (v) { return [v]; });
    return null;
  }

  /** TREND … 直線で 予測する（縦1列で 返す） */
  function 直線予測(yの表, xの表, 新xの表, 切片あり) {
    var ys = 数だけ(yの表);
    if (!ys.length) return 誤り('VALUE');
    var xs = xを整える(ys, xの表);
    if (!xs) return 誤り('REF');
    var r = 回帰(ys, xs, 切片あり === undefined || 切片あり === null ? true : !!切片あり);
    if (!r) return 誤り('VALUE');
    var 新 = (新xの表 === undefined || 新xの表 === null || (Array.isArray(新xの表) && !平らに(新xの表).length))
      ? xs : xを整える(new Array(平らに(新xの表).length / xs[0].length).fill(0), 新xの表);
    if (!新) return 誤り('REF');
    return 新.map(function (row) {
      var v = r.係数[0];
      for (var j = 0; j < row.length; j++) v += r.係数[j + 1] * row[j];
      return [v];
    });
  }
  /** GROWTH … 増える 形で 予測する */
  function 増える予測(yの表, xの表, 新xの表, 切片あり) {
    var ys = 数だけ(yの表);
    if (!ys.length) return 誤り('VALUE');
    for (var i = 0; i < ys.length; i++) if (!(ys[i] > 0)) return 誤り('NUM');
    var 対数 = ys.map(function (v) { return Math.log(v); });
    var xs = xを整える(対数, xの表);
    if (!xs) return 誤り('REF');
    var r = 回帰(対数, xs, 切片あり === undefined || 切片あり === null ? true : !!切片あり);
    if (!r) return 誤り('VALUE');
    var 新 = (新xの表 === undefined || 新xの表 === null || (Array.isArray(新xの表) && !平らに(新xの表).length))
      ? xs : xを整える(new Array(平らに(新xの表).length / xs[0].length).fill(0), 新xの表);
    if (!新) return 誤り('REF');
    return 新.map(function (row) {
      var v = r.係数[0];
      for (var j = 0; j < row.length; j++) v += r.係数[j + 1] * row[j];
      return [Math.exp(v)];
    });
  }
  /** ★LINEST … 直線の 係数（Excel と 同じ 並び＝逆順・最後が 切片）★（2026-09-09）
   *
   *  ★なぜ ここに 足したか★
   *    今まで LINEST は `exally-formula.js` の JS層が ★裸の =LINEST(範囲,範囲) だけ★を 拾い、
   *    ★傾き 1つ★を 返して いた。だから
   *      ・`=INDEX(LINEST(…),1,1)` は ★#NAME?★（エンジンに 積んで いない）
   *      ・★x が 2本以上（重回帰）で 静かに 違う 答え★
   *        =LINEST(G1:G6,H1:I6) … 実Excel 0.7708333333333329 ／ うち ★7.390243902439025★
   *    ⇒ 下の `増えるの係数`（LOGEST）は ★同じ 土台★で 重回帰も 統計の 行も 出せて いる
   *      （LOGEST/TREND/GROWTH は 実Excel と ★33本 全部 一致★）
   *    ⇒★★LINEST は それの log と exp を 外しただけ＝断らずに 直せる★★
   *
   *  ★実Excel の 実測★（`docs/measured/golden-linest-hyou-2026-09-09.tsv`・119本）
   *    ・★必ず 表を 返す★（1つの 数では ない）
   *    ・★列 ＝ x の 本数 ＋ 1★（x1本→2列 ／ x2本→3列）
   *    ・★行 ＝ 既定 1行 ／ 統計を TRUE に すると 5行★
   *    ・5行 … 1行 係数と 切片 ／ 2行 それぞれの 標準誤差 ／
   *            3行 決定係数 と y の 標準誤差 ／ 4行 F と 自由度 ／ 5行 回帰と 残差の 平方和
   *    ・★x2本の 3行3列以降は #N/A★（★余りは 埋め物★・空でも 0 でも ない）
   */
  function 直線の係数(yの表, xの表, 切片あり, 統計も) {
    var ys = 数だけ(yの表);
    if (!ys.length) return 誤り('VALUE');
    var xs = xを整える(ys, xの表);
    if (!xs) return 誤り('REF');
    var 切 = 切片あり === undefined || 切片あり === null ? true : !!切片あり;
    var r = 回帰(ys, xs, 切);
    if (!r) return 誤り('VALUE');
    var 次 = xs[0].length;
    /* ★実Excel の 並びは mn…m1, b★（★後ろの x から 前へ★） */
    var 行1 = [];
    for (var j = 次; j >= 1; j--) 行1.push(r.係数[j]);
    行1.push(r.係数[0]);
    if (!統計も) return [行1];
    var se = r.統計.係数の標準誤差;
    var 行2 = [];
    for (var k = 次; k >= 1; k--) 行2.push(se[切 ? k : k - 1]);
    行2.push(切 ? se[0] : 0);
    /* ★余りは #N/A★（実Excel の 実測どおり） */
    var 埋 = function (a) { while (a.length < 次 + 1) a.push(誤り('NA')); return a; };
    /* ★★自由度が 0 なら F は #NUM!★★（実Excel の 実測）
       =INDEX(LINEST(K1:K2,L1:L2,TRUE,TRUE),4,1) … 点が 2つだけ＝自由度 0
         実Excel ★#NUM!★ ／ 直す 前の うち ★0★
       ⇒★割れない 物を 0 と 言わない★ */
    var F = r.統計.自由度 > 0 ? r.統計.F : 誤り('NUM');
    return [
      行1, 行2,
      埋([r.統計.決定係数, r.統計.標準誤差y]),
      埋([F, r.統計.自由度]),
      埋([r.統計.回帰二乗和, r.統計.残差二乗和])
    ];
  }

  /** LOGEST … 増える 形の 係数（Excel と 同じ 並び＝逆順・最後が 切片） */
  function 増えるの係数(yの表, xの表, 切片あり, 統計も) {
    var ys = 数だけ(yの表);
    if (!ys.length) return 誤り('VALUE');
    for (var i = 0; i < ys.length; i++) if (!(ys[i] > 0)) return 誤り('NUM');
    var 対数 = ys.map(function (v) { return Math.log(v); });
    var xs = xを整える(対数, xの表);
    if (!xs) return 誤り('REF');
    var 切 = 切片あり === undefined || 切片あり === null ? true : !!切片あり;
    var r = 回帰(対数, xs, 切);
    if (!r) return 誤り('VALUE');
    var 次 = xs[0].length;
    var 行1 = [];
    for (var j = 次; j >= 1; j--) 行1.push(Math.exp(r.係数[j]));
    行1.push(Math.exp(r.係数[0]));
    if (!統計も) return [行1];
    var se = r.統計.係数の標準誤差;
    var 行2 = [];
    for (var k = 次; k >= 1; k--) 行2.push(Math.abs(行1[次 - k]) * se[切 ? k : k - 1]);
    行2.push(切 ? Math.abs(行1[次]) * se[0] : 0);
    var 埋 = function (a) { while (a.length < 次 + 1) a.push(''); return a; };
    return [
      行1, 行2,
      埋([r.統計.決定係数, r.統計.標準誤差y]),
      埋([r.統計.F, r.統計.自由度]),
      埋([r.統計.回帰二乗和, r.統計.残差二乗和])
    ];
  }

  /* ════════ 誤差関数 ════════
     ★不完全ガンマ関数から 出す★（級数と 連分数を 使い分ける）
     ⇒ どちらも ★倍精度いっぱい★まで 詰められる形 */
  function 対数ガンマ(x) {
    var c = [76.18009172947146, -86.50532032941677, 24.01409824083091,
      -1.231739572450155, 0.1208650973866179e-2, -0.5395239384953e-5];
    var y = x, t = x + 5.5;
    t -= (x + 0.5) * Math.log(t);
    var s = 1.000000000190015;
    for (var j = 0; j < 6; j++) s += c[j] / ++y;
    return -t + Math.log(2.5066282746310005 * s / x);
  }
  function ガンマ級数(a, x) {
    var ap = a, sum = 1 / a, del = sum;
    for (var n = 1; n <= 1000; n++) {
      ap++; del *= x / ap; sum += del;
      if (Math.abs(del) < Math.abs(sum) * 1e-17) break;
    }
    return sum * Math.exp(-x + a * Math.log(x) - 対数ガンマ(a));
  }
  function ガンマ連分数(a, x) {
    var 極小 = 1e-300;
    var b = x + 1 - a, c = 1 / 極小, d = 1 / b, h = d;
    for (var i = 1; i <= 1000; i++) {
      var an = -i * (i - a);
      b += 2; d = an * d + b; if (Math.abs(d) < 極小) d = 極小;
      c = b + an / c; if (Math.abs(c) < 極小) c = 極小;
      d = 1 / d;
      var del = d * c; h *= del;
      if (Math.abs(del - 1) < 1e-17) break;
    }
    return Math.exp(-x + a * Math.log(x) - 対数ガンマ(a)) * h;
  }
  function 誤差(x) {
    x = Number(x);
    if (!isFinite(x)) return 誤り('VALUE');
    if (x === 0) return 0;
    var a = Math.abs(x);
    var v = (a * a < 1.5) ? ガンマ級数(0.5, a * a) : 1 - ガンマ連分数(0.5, a * a);
    return x < 0 ? -v : v;
  }
  function 誤差の残り(x) {
    x = Number(x);
    if (!isFinite(x)) return 誤り('VALUE');
    if (x === 0) return 1;
    var a = Math.abs(x);
    var q = (a * a < 1.5) ? 1 - ガンマ級数(0.5, a * a) : ガンマ連分数(0.5, a * a);
    return x < 0 ? 2 - q : q;
  }

  /* ════════ その他 ════════ */

  /** PERCENTOF … 一部が 全体の 何割か */
  function 割合(一部, 全体) {
    var a = 数だけ(一部), b = 数だけ(全体);
    if (!b.length) return 誤り('DIV_BY_ZERO');
    var 和 = function (v) { return v.reduce(function (x, y) { return x + y; }, 0); };
    var 下 = 和(b);
    if (下 === 0) return 誤り('DIV_BY_ZERO');
    return 和(a) / 下;
  }

  /** RANDARRAY … でたらめの 表（★数は 毎回 変わる／形と 範囲を 決める★） */
  function でたらめの表(行, 列, 下, 上, 整数か) {
    行 = (行 === null || 行 === undefined || 行 === '') ? 1 : Math.floor(Number(行));
    列 = (列 === null || 列 === undefined || 列 === '') ? 1 : Math.floor(Number(列));
    if (!(行 >= 1) || !(列 >= 1)) return 誤り('VALUE');
    var 使う = !(下 === null || 下 === undefined || 下 === '') || !(上 === null || 上 === undefined || 上 === '');
    var lo = 使う ? Number(下 || 0) : 0;
    var hi = 使う ? Number(上 || 0) : 1;
    if (使う && lo > hi) return 誤り('VALUE');
    var 出 = [];
    for (var r = 0; r < 行; r++) {
      var 一行 = [];
      for (var c = 0; c < 列; c++) {
        var v;
        if (!使う) v = Math.random();
        else if (整数か) v = Math.floor(Math.random() * (hi - lo + 1)) + lo;
        else v = lo + Math.random() * (hi - lo);
        一行.push(v);
      }
      出.push(一行);
    }
    return 出;
  }

  /* ════════ INFO（今の 場の 事を 返す） ════════
     ★★何を 返すかは 司さんの 決め★★（2026-09-07「全部やって」で 決まった）
       実Excel（16.0 build 20326・日本語）に 打たせた 答え
         numfile   … 1            ⇒ Exally は ★ブックの シートの 数★
         recalc    … 自動          ⇒ Exally は 常に 自動＝★自動★
         release   … 16.0         ⇒ Exally は ★自分の 版★（Excel の 版を 名乗ると 嘘に なる）
         system    … pcdos        ⇒ Windows なら pcdos／Mac なら mac（実Excel と 同じ 言葉）
         osversion … Windows (64-bit) NT 10.00 ⇒ 分かる 範囲で 同じ 形
       ★返せない 物★
         directory … ★パソコンの フォルダ★＝ブラウザの 中には 無い ⇒ #N/A
         memavail / memused / totmem … 実Excel も 今は #N/A
       ⇒★★無い 物を それらしく 作らない★★（#N/A で 返す）
     ★場の 事は 呼ぶ側が 渡す★＝ここは 純粋なまま（試験で そのまま 押せる） */
  function 場の事(何を, 場) {
    var k = String(何を 
      === null || 何を === undefined ? '' : 何を).trim().toLowerCase();
    var e = 場 || {};
    if (k === 'numfile') return Number(e.シート数 || 0) || 1;
    if (k === 'recalc') return '自動';
    if (k === 'release') return String(e.版 || '');
    if (k === 'system') return e.台 === 'mac' ? 'mac' : 'pcdos';
    if (k === 'osversion') return String(e.OS || '');
    if (k === 'origin') return '$A:' + String(e.左上 || '$A$1');
    /* ★無い 物は 作らない★ */
    if (k === 'directory' || k === 'memavail' || k === 'memused' || k === 'totmem') return 誤り('NA');
    return 誤り('VALUE');
  }

  /* ★数える（見張り用）★手で 並べない＝ここが 正本★ */
  function 足した名前() {
    /* ★LINEST は 2026-09-09 に 足した★
       （それまでは exally-formula.js の JS層が ★傾き 1つ★だけ 返して いた） */
    return ['CONVERT', 'TREND', 'GROWTH', 'LOGEST', 'LINEST',
      'ERF.PRECISE', 'ERFC.PRECISE', 'PERCENTOF', 'RANDARRAY', 'INFO'];
  }
  function 保留の名前() {
    return ['TRIMRANGE',
      'FORECAST.ETS', 'FORECAST.ETS.CONFINT', 'FORECAST.ETS.SEASONALITY', 'FORECAST.ETS.STAT'];
  }

  return {
    単位を読む: 単位を読む, 温度を読む: 温度を読む, 単位を変える: 単位を変える,
    回帰: 回帰, 直線予測: 直線予測, 増える予測: 増える予測, 増えるの係数: 増えるの係数, 直線の係数: 直線の係数,
    誤差: 誤差, 誤差の残り: 誤差の残り, 割合: 割合, でたらめの表: でたらめの表, 場の事: 場の事,
    表にする: 表にする, 数だけ: 数だけ,
    足した名前: 足した名前, 保留の名前: 保留の名前
  };
}));
