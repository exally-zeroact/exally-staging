/* formula-extra.js — ★実Excel に 在って 計算エンジンに 無い 関数★ 2026-08-31
 *
 *  ★司さんの 方針（2026-08-30）★
 *    「Excel を 細胞分解レベルまで 網羅して 把握した上で 持ち込み パクる」
 *
 *  ★実測（2026-08-31・実ブラウザで 数えた）★
 *    計算エンジン（HyperFormula）に 登録されている 関数 … ★457個★
 *    ★事務で 使う 物を 98個 見て、62個が 無かった★
 *
 *  ★ここで 足す 物★＝★うちだけで 作れて、答えが 決まっている 物★
 *    ・AVERAGEIFS      … 条件つきの 平均（毎日 使う）
 *    ・TAKE / DROP     … 表の 頭・尻を 取る／落とす
 *    ・CHOOSECOLS / CHOOSEROWS … 列・行を 選ぶ
 *    ・TOCOL / TOROW   … 表を 1列／1行に する
 *    ・WRAPROWS / WRAPCOLS … 1列を 表に 折り返す
 *    ・EXPAND          … 表を 広げる
 *    ・ARRAYTOTEXT     … 表を 字に する
 *    ・MODE.MULT       … 最頻値（複数）
 *
 *  ★足さない 物と その 理由★（★「これから」とも 書かない★）
 *    ・LAMBDA / LET / MAP / REDUCE / BYROW / BYCOL / SCAN / MAKEARRAY
 *        … ★式を 値として 渡す 仕組み★が エンジンに 無い（作るなら エンジン側）
 *    ・STOCKHISTORY / FIELDVALUE / WEBSERVICE / RTD / IMAGE
 *        … ★外の データ配信が 要る★
 *    ・PHONETIC … ★ふりがなは セルに 付いた データ★＝エンジンでは 見えない
 *        （うちは 画面側に「ふりがなを入れる」が 在る）
 *    ・BAHTTEXT … ★2026-09-04 に 足した★（lib/bahttext.js）
 *        ★前は「タイ語専用」と 書いて 足していなかった★＝★理由に なっていなかった★
 *        （司さん「Excelが 出来て うちが 出来んって事は 絶対に 無い」）
 *        ★答えは 実Excel に 116通り 打たせて 合わせた★
 *    ・YEN … ★日本語版の 画面の 名前★（2026-09-05 実Excel 16.0 build 20326・UI 1041 で 実測）
 *          `=YEN(1234)`（英語名）＝#NAME? ／ ★FormulaLocal `=YEN(1234)` ＝ ¥1,234★
 *          ⇒★中身は DOLLAR と 同じ★。★お客さんの 画面には YEN と 出る★
 *          ⇒★08-31 に「存在しない」と 書いたのは ★英語名しか 見ていなかった★から
 *        円は `DOLLAR` が 出す（実測＝DOLLAR(-1234.5,1) → ¥-1,234.5）。
 *        ★★棚は「足さない」ではなく ★別名で動く★（2026-09-05 夜に 移した）★★
 *          ＝★お客さんが YEN( と 打てば そのまま 動く★（実測 =YEN(1234.5) → ¥1,235）
 *            `exally-formula.js` の `convertFormula` が 本名 DOLLAR に 直している
 *          ⇒★「足さない」の 棚に 置いていた 間、AIへ 渡る 紙では
 *            ★「Exally内では まだ 動かない」の 列に 並んでいた★＝★AIが 客に 嘘を 言う★所だった
 *          ⇒★★「自分では 作っていない」と「動かない」は 別★★
 *          見張り … tests/kansuu-tana.test.mjs（★棚を engine に 押して 確かめる★）
 *
 *  ★形（spill）を 保つ★
 *    2026-08-29 に ★FILTER が 縦1列しか 返せず 月別合計が 全部 0★に なった。
 *    ⇒ ★2列以上でも 形を 保つ★事を 見張りで 確かめる。
 *
 *  見張り: tests/formula-extra.test.mjs
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.FormulaExtra = factory();
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  /* ══ ここは ★純粋な 計算★ だけ ══════════════════════
     エンジンに 繋ぐ 所（登録）は 下の つなぐ() で やる。
     ★分けてある理由★＝node で そのまま 試験できる。 */

  function 数か(v) { return typeof v === 'number' && isFinite(v); }

  /** 表（2次元の 配列）に そろえる */
  function 表にする(v) {
    if (!Array.isArray(v)) return [[v]];
    if (!v.length) return [[]];
    if (!Array.isArray(v[0])) return [v.slice()];
    return v.map(function (r) { return r.slice(); });
  }

  /** ★条件つきの 平均★（AVERAGEIFS）
   *  @param 平均する  数の 表
   *  @param 組        [条件の表, 条件] を 並べた 物
   *  ★1つも 当たらなければ #DIV/0!★（実Excel と 同じ）
   */
  function 条件つき平均(平均する, 組) {
    var A = 表にする(平均する);
    var 足 = 0, 数 = 0;
    for (var r = 0; r < A.length; r++) {
      for (var c = 0; c < A[r].length; c++) {
        var 当たる = true;
        for (var k = 0; k < 組.length; k++) {
          var B = 表にする(組[k][0]);
          var v = (B[r] === undefined) ? undefined : B[r][c];
          if (!条件に合うか(v, 組[k][1])) { 当たる = false; break; }
        }
        if (!当たる) continue;
        var x = A[r][c];
        if (!数か(x)) continue;
        足 += x; 数++;
      }
    }
    if (!数) return { 誤り: 'DIV_BY_ZERO' };
    return 足 / 数;
  }

  /** 条件の 見方（実Excel と 同じ＝">=10" や "*円" が 使える） */
  function 条件に合うか(v, 条件) {
    if (条件 === null || 条件 === undefined) return v === null || v === undefined || v === '';
    if (typeof 条件 === 'number') return Number(v) === 条件;
    var s = String(条件);
    var m = s.match(/^(<=|>=|<>|<|>|=)(.*)$/);
    if (m) {
      var 印 = m[1], 右 = m[2];
      var a = 数か(v) ? v : (String(v) === '' ? null : (isNaN(Number(v)) ? String(v) : Number(v)));
      var b = (右 === '') ? null : (isNaN(Number(右)) ? 右 : Number(右));
      if (印 === '=') return 等しいか(a, b);
      if (印 === '<>') return !等しいか(a, b);
      if (a === null || b === null) return false;
      if (typeof a !== typeof b) return false;
      if (印 === '<') return a < b;
      if (印 === '>') return a > b;
      if (印 === '<=') return a <= b;
      if (印 === '>=') return a >= b;
      return false;
    }
    /* ★* と ? が 使える★（実Excel と 同じ） */
    if (/[*?]/.test(s)) {
      var 型 = '^' + s.replace(/[.+^${}()|[\]\\]/g, '\\$&')
        .split('*').join('[\\s\\S]*').split('?').join('[\\s\\S]') + '$';
      return new RegExp(型, 'i').test(String(v === null || v === undefined ? '' : v));
    }
    return 等しいか(v, isNaN(Number(s)) ? s : Number(s));
  }
  function 等しいか(a, b) {
    if (typeof a === 'string' && typeof b === 'string') return a.toLowerCase() === b.toLowerCase();
    return a === b;
  }

  /** ★頭から N個★（TAKE）／★落とす★（DROP）
   *  ★マイナスなら 尻から★（実Excel と 同じ） */
  function 切り出す(表, 行数, 列数, 落とすか) {
    var A = 表にする(表);
    var 行 = A.length, 列 = A[0] ? A[0].length : 0;
    function 幅(n, 全) {
      if (n === null || n === undefined) return [0, 全];
      n = Math.trunc(n);
      if (落とすか) {
        if (n >= 0) return [Math.min(n, 全), 全];
        return [0, Math.max(0, 全 + n)];
      }
      if (n >= 0) return [0, Math.min(n, 全)];
      return [Math.max(0, 全 + n), 全];
    }
    var r = 幅(行数, 行), c = 幅(列数, 列);
    var 出 = [];
    for (var i = r[0]; i < r[1]; i++) {
      var 行の物 = [];
      for (var j = c[0]; j < c[1]; j++) 行の物.push(A[i][j]);
      出.push(行の物);
    }
    /* ★★中身が 無くなったら ★#CALC!★★★（2026-09-15・実Excel の 実測）
         `=TAKE(A1:A5,0)` `=TAKE(A1:A5,0.5)` `=TAKE(A1:A5,-0.5)`
         `=DROP(A1:A5,5)` `=DROP(A1:A5,6)` `=DROP(A1:A5,45294)`
       ⇒★6本とも 実Excel は `#CALC!`★／★前は `#VALUE!` を 返して いた★
       ★紙★ `docs/measured/kansuu46/golden-346-2026-09-08.tsv`
       ★見張り★ `tests/shiki-kansuu-kami.test.mjs`
       ★`#CALC!` は「空の 表に なった」時の 誤り★＝★`#VALUE!`（渡し方が 悪い）とは 別★ */
    if (!出.length || !出[0].length) return { 誤り: 'CALC' };
    return 出;
  }

  /** ★列を 選ぶ★（CHOOSECOLS）／★行を 選ぶ★（CHOOSEROWS）
   *  ★1から 数える／マイナスは 尻から★（実Excel と 同じ） */
  function 選び出す(表, 番号たち, 列か) {
    var A = 表にする(表);
    var 全 = 列か ? (A[0] ? A[0].length : 0) : A.length;
    var 番 = [];
    for (var i = 0; i < 番号たち.length; i++) {
      var n = Math.trunc(番号たち[i]);
      if (n === 0) return { 誤り: 'VALUE' };
      var k = n > 0 ? n - 1 : 全 + n;
      if (k < 0 || k >= 全) return { 誤り: 'VALUE' };
      番.push(k);
    }
    if (!番.length) return { 誤り: 'VALUE' };
    if (列か) {
      return A.map(function (r) { return 番.map(function (k) { return r[k]; }); });
    }
    return 番.map(function (k) { return A[k].slice(); });
  }

  /** ★1列／1行に する★（TOCOL / TOROW）
   *  @param 無視 0＝全部 / 1＝空を とばす / 2＝誤りを とばす / 3＝両方
   *  @param 列で読むか true＝列の順に 読む（実Excel の 3つ目の 引数） */
  function 一本にする(表, 無視, 列で読むか, 縦か) {
    var A = 表にする(表);
    var 行 = A.length, 列 = A[0] ? A[0].length : 0;
    /* ★★2つ目は『大きさ』では なく『無視する 印』＝0〜3 だけ★★（2026-09-08）
       ⇒ 私は ここを ★大きさの 引数だと 読み違えて いました★
       ★実Excel に 15通り 打った★（紙 golden-hikisuu-mihari-2026-09-08.tsv）
         n=-2 -1 … ★#VALUE!★ ／ n=-0.5 0 0.5 1 1.5 2 2.5 3 3.5 … 通る ／ n=4 5 6 45294 … ★#VALUE!★
       ⇒★切り捨てて 0〜3 の 外なら #VALUE!★（−0.5 は 切り捨てて 0＝通る）
       ⇒ 直す前は ★何を 渡しても 受けて いた★＝`=TOCOL(A1:B5,DATE(2024,1,3))` が 答えを 返して いた */
    var 印 = (無視 === null || 無視 === undefined || 無視 === '') ? 0 : Number(無視);
    if (!isFinite(印)) return { 誤り: 'VALUE' };
    印 = Math.trunc(印);
    if (印 < 0 || 印 > 3) return { 誤り: 'VALUE' };
    var 出 = [];
    function 入れる(v) {
      var 空 = (v === null || v === undefined || v === '');
      var 誤 = !!(v && v.誤り);
      if ((印 === 1 || 印 === 3) && 空) return;
      if ((印 === 2 || 印 === 3) && 誤) return;
      出.push(v);
    }
    if (列で読むか) {
      for (var c = 0; c < 列; c++) for (var r = 0; r < 行; r++) 入れる(A[r][c]);
    } else {
      for (var r2 = 0; r2 < 行; r2++) for (var c2 = 0; c2 < 列; c2++) 入れる(A[r2][c2]);
    }
    if (!出.length) return { 誤り: 'VALUE' };
    /* ★形を 保つ★＝縦は 1列の 表・横は 1行の 表（2026-08-29 の FILTER の 失敗を 繰り返さない） */
    return 縦か ? 出.map(function (v) { return [v]; }) : [出];
  }

  /** ★1本を 表に 折り返す★（WRAPROWS / WRAPCOLS） */
  function 折り返す(表, 幅, 埋め, 行でか) {
    var A = 表にする(表);
    var 平 = [];
    for (var r = 0; r < A.length; r++) for (var c = 0; c < A[r].length; c++) 平.push(A[r][c]);
    /* ★実Excel は ここで ★#NUM!★ と 言う（#VALUE! では ない）★（2026-09-08 実測）
         =WRAPROWS(A1:A5,n) … n=-2 -1 -0.5 0 0.5 … ★#NUM!★ ／ n=1 から 通る
         =WRAPCOLS も 同じ
       ⇒ うちは #VALUE! を 返して いた＝★答えは 断れて いたが 字が 違った★ */
    var w = Math.trunc(Number(幅));
    if (!(w >= 1)) return { 誤り: 'NUM' };
    var 本数 = Math.ceil(平.length / w);
    var 出 = [];
    for (var i = 0; i < 本数; i++) {
      var 一本 = [];
      for (var j = 0; j < w; j++) {
        var k = i * w + j;
        一本.push(k < 平.length ? 平[k] : (埋め === undefined ? { 誤り: 'NA' } : 埋め));
      }
      出.push(一本);
    }
    if (行でか) return 出;
    /* 列で 折り返す＝行と列を 入れ替える */
    var 転 = [];
    for (var y = 0; y < w; y++) {
      var 行の物 = [];
      for (var x = 0; x < 出.length; x++) 行の物.push(出[x][y]);
      転.push(行の物);
    }
    return 転;
  }

  /** ★表を 広げる★（EXPAND） */
  function 広げる(表, 行数, 列数, 埋め) {
    var A = 表にする(表);
    var 行 = A.length, 列 = A[0] ? A[0].length : 0;
    var R = (行数 === null || 行数 === undefined) ? 行 : Math.trunc(行数);
    var C = (列数 === null || 列数 === undefined) ? 列 : Math.trunc(列数);
    if (R < 行 || C < 列) return { 誤り: 'VALUE' };
    var 埋 = (埋め === undefined) ? { 誤り: 'NA' } : 埋め;
    var 出 = [];
    for (var r = 0; r < R; r++) {
      var 行の物 = [];
      for (var c = 0; c < C; c++) {
        行の物.push((r < 行 && c < 列) ? A[r][c] : 埋);
      }
      出.push(行の物);
    }
    return 出;
  }

  /** ★表を 字に する★（ARRAYTOTEXT）
   *  @param 形 0＝簡単（カンマ区切り）／1＝きっちり（{}付き・実Excel と 同じ） */
  function 表を字に(表, 形) {
    var A = 表にする(表);
    function 一つ(v) {
      if (v === null || v === undefined) return '';
      if (typeof v === 'string') return 形 === 1 ? '"' + v + '"' : v;
      if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
      return String(v);
    }
    if (形 !== 1) {
      var 平 = [];
      for (var r = 0; r < A.length; r++) for (var c = 0; c < A[r].length; c++) 平.push(一つ(A[r][c]));
      return 平.join(', ');
    }
    return '{' + A.map(function (row) {
      return row.map(一つ).join(',');
    }).join(';') + '}';
  }

  /** ★最頻値（複数）★（MODE.MULT）＝★縦1列で 返す★（実Excel と 同じ） */
  function 最頻値たち(表) {
    var A = 表にする(表);
    var 数え = [], 順 = [];
    for (var r = 0; r < A.length; r++) {
      for (var c = 0; c < A[r].length; c++) {
        var v = A[r][c];
        if (!数か(v)) continue;
        var i = 順.indexOf(v);
        if (i < 0) { 順.push(v); 数え.push(1); } else 数え[i]++;
      }
    }
    var 最大 = 0;
    for (var k = 0; k < 数え.length; k++) if (数え[k] > 最大) 最大 = 数え[k];
    if (最大 < 2) return { 誤り: 'NA' };          /* ★1回ずつなら #N/A★（実Excel と 同じ） */
    var 出 = [];
    for (var m = 0; m < 順.length; m++) if (数え[m] === 最大) 出.push([順[m]]);
    return 出;
  }

  /** 数える（見張り用） */
  function 数える() {
    return {
      足す: ['AVERAGEIFS', 'TAKE', 'DROP', 'CHOOSECOLS', 'CHOOSEROWS',
        'TOCOL', 'TOROW', 'WRAPROWS', 'WRAPCOLS', 'EXPAND',
        'ARRAYTOTEXT', 'MODE.MULT'],
      /* ★ここ（このファイル）では 計算していないが、繋いである 物★
         ＝★別の lib に 中身が 在る 物★。★繋ぐ 数と 合わせる 為に 数える★
         （書かないと「繋ぐ 数が 1つ 多い」と 見えて、見張りが 嘘に なる） */
      別のlibで足す: { 'BAHTTEXT': 'lib/bahttext.js（実Excel に 116通り 打たせて 合わせた）' },
      /* ★★別名で 動く 物★★（2026-09-05 に この 棚を 足した）
         ＝★ここでは 作っていないが、お客さんが 打てば ちゃんと 答える★
         ★なぜ 棚を 分けたか★
           YEN を 「足さない」に 置いていた ⇒ AIへ 渡る 紙で
           ★「Exally内では まだ 動かない」の 列に 並んでいた★
           ⇒★実際は ¥1,235 が 出るのに、AIが お客さんに「動きません」と 言う★
           ⇒★「足していない」と「動かない」は 別★＝棚を 分ける */
      /* ★★「別名で動く」＝★お客さんが 打てば そのまま 動く 物★★★
         ・別の 名前に 直して 動かす 物（YEN→DOLLAR・JIS→DBCS …）
         ・★別の 道（JS層）が 答える 物★（LET・MAP・REDUCE・SCAN・MAKEARRAY）
         ⇒ どちらも ★AIは 断っては いけない★＝紙では 同じ 列に 出す */
      /* ★★3つ目の 棚＝「形に よって 出る」★★（2026-09-08 に 足した）
         ★なぜ 棚を 分けたか（指示役の 決定）★
           ・「動く」に 入れると ★数が 嘘に なる★
             ⇒ 数は 1本 良く なるが ★お客さんの 式は 黙って 止まる★
           ・「動かない」に 置くと ★測って 分かった 事を 1つも 渡さない★
           ⇒★★『出る 形』と『まだの 形』を ★1行の 中で 名指しする★★
         ★AI は この 棚を「動く」とも「動かない」とも 言っては いけない★
           ⇒★★『この 形なら 出ます／この 形は まだです』と 言う★★ */
      形で動く: {
        'AREAS': '★範囲が 幾つに 分かれているかを 数える★。'
          + '★出る 形★＝ふつうの 範囲（B2:D4）／1マス／列ぜんぶ（A:A）／行ぜんぶ（1:1）／'
          + 'とびとび（(B2:D4,E5,F6:I9) → 3）／重なり（B2:D4 C3:E5 → 1）／'
          + 'シート付き（別の シートを またぐ とびとびは #VALUE!）。'
          + '★まだの 形★＝★かっこが 入っている 物 全部★'
          + '（=AREAS(INDEX(...)) の ような 関数入り／シート名に かっこが 在る 物）と、'
          + '★重ならない 重なり（=AREAS(B2:D4 A1)）★ ⇒ #NAME? に なる。'
          + '★訳★ 関数が 指せたかは 字だけでは 決まらず'
          + '（INDEX(...,1,1) は 1／INDEX(...,99,1) は #REF!）、'
          + 'どちらも 計算すると 誤りに なるので 見分けられない。'
          + '重ならない 重なりは 実Excel が #NULL! を 返すが うちの エンジンに その 誤りが 無い。'
          + '間違った 数を 出すより「まだ」と 言う（2026-09-08 実測 89本）',
        'ISOMITTED': '★LAMBDA の 引数が 省かれたかを 見る★。'
          + '★出る 形★＝引数を 書いた 物 全部 → ★FALSE★'
          + '（=ISOMITTED(A1)／=ISOMITTED(2)／=ISOMITTED("あ")／'
          + '★空の マスも FALSE★＝=ISOMITTED(C1)／=LET(x,1,ISOMITTED(x))／'
          + 'LAMBDA に 渡した 物 =LAMBDA(x,y,ISOMITTED(y))(1,2)）。'
          + '★まだの 形★＝★カンマで 省いた 物★ ⇒ 誤りに なる'
          + '（=LAMBDA(x,y,ISOMITTED(y))(1,)／=LAMBDA(x,y,ISOMITTED(x))(,2)／'
          + '=LAMBDA(x,y,z,ISOMITTED(y))(1,,3) … 実Excel は ★TRUE★）。'
          + '★訳★ うちの LAMBDA は 字を 置き換えて 展開する 作りで、'
          + '省かれた 引数が ISOMITTED(()) に なり 見分けられない。'
          + '間違った TRUE/FALSE を 出すより「まだ」と 言う。'
          + '★『空の マス』と『省かれた 引数』は 別物★'
          + '（2026-09-08 実Excel 実測 … docs/measured/kansuu46/toru-isomitted.ps1）',
      },
      別名で動く: {
        /* ★2026-09-06 実測＝押したら 動いた★（前は「動かない」棚に 居た） */
        'LET': '★打てば そのまま 動く★＝=LET(x,2,x*3) → 6（実測）。'
          + '名前を 付けて 使い回す 書き方が そのまま 通る',
        'MAP': '★打てば そのまま 動く★＝=MAP(A1:A3,LAMBDA(v,v*2)) が 答える（実測）',
        'REDUCE': '★打てば そのまま 動く★＝=REDUCE(0,A1:A3,LAMBDA(a,b,a+b)) → 9（実測）',
        'SCAN': '★打てば そのまま 動く★＝途中の 計算を 並べて 返す（実測）',
        'MAKEARRAY': '★打てば そのまま 動く★＝=MAKEARRAY(2,2,LAMBDA(r,c,r*c)) が 答える（実測）',
        /* ★★2026-09-19 実測＝「足さない」棚から 移した 4個★★
             ★出どころ★ 見張りを ★本番と 同じ 3段★（JS層 → 台 → 借り物）に 揃えたら 出た
             ★押した 字（材料 A1=1 A2=2 B1=10 B2=20）★ ... ★8 / 8 とも 答えた★ */
        'BYROW': '★打てば そのまま 動く★＝=BYROW(A1:B2,LAMBDA(r,SUM(r))) → 11（実測）。'
          + '=SUM(BYROW(...)) → 33。★行ごとに まとめて 縦に 並べて 返す★',
        'BYCOL': '★打てば そのまま 動く★＝=BYCOL(A1:B2,LAMBDA(c,SUM(c))) → 3（実測）。'
          + '=SUM(BYCOL(...)) → 33。★列ごとに まとめて 横に 並べて 返す★',
        'LAMBDA': '★その場で 呼ぶ 書き方は 動く★＝=LAMBDA(x,x+1)(2) → 3（実測）。'
          + '★裸の =LAMBDA(x,x+1) は #CALC!★＝★実Excel と 同じ 字★（2026-09-19 実測）。'
          + 'セルに 名前として しまう 使い方は まだ 出来ない',
        'PHONETIC': '★打てば そのまま 動く★＝=PHONETIC(A1) も =PHONETIC(A1:A2) も ★空の 字★'
          + '（実Excel と 同じ・2026-09-19 実測）。'
          + '★但し ふりがなの 入った ブックを 読んだ 時は まだ 違う★'
          + '＝xlsx の <rPh> を 読んで いない（実Excel は ヤマダ を 返す）',
        /* ★★手で 挙げていません★★＝tests/betsumei-zenbu.test.mjs が
           `convertFormula` から 機械で 拾い、1つずつ engine に 押して 出した 物
           （2026-09-06 指示役「残りを 手で 挙げないで ください／出してから 数える」）
           ★ここに 足し忘れると その日に 赤に なる★ */
        'JIS': '★打てば そのまま 動く★＝Exally が ★DBCS★（本名）に 直して 答える。'
          + '=JIS("あ") → あ。JIS は 日本語UIの 表示名で、ファイル／英語の 構文では DBCS'
          + '（YEN と 同じ 家／直している 場所＝exally-formula.js の convertFormula）',
        'NORMSDIST': '★打てば そのまま 動く★＝古い 名前。Exally が '
          + 'NORMSDIST(x,TRUE()) に 直して 答える（=NORMSDIST(1) → 0.8413…）。'
          + '新しい 名前は NORM.S.DIST（直している 場所＝convertFormula）',
        'BETADIST': '★打てば そのまま 動く★＝古い 名前。Exally が '
          + 'BETA.DIST(…,TRUE()) に 直して 答える（=BETADIST(0.5,1,2) → 0.75）'
          + '（直している 場所＝convertFormula）',
        'HYPGEOMDIST': '★打てば そのまま 動く★＝古い 名前。Exally が 4つ目の 引数に '
          + 'FALSE() を 足して 答える（=HYPGEOMDIST(1,2,3,4) → 0.5）'
          + '（直している 場所＝convertFormula）',
        'NEGBINOMDIST': '★打てば そのまま 動く★＝古い 名前。Exally が 引数を 足して '
          + '答える（=NEGBINOMDIST(1,2,0.5) → 0.25）（直している 場所＝convertFormula）',
        'ISREF': '★マスを 指しているか★を 見る。'
          + '★そのまま 動く 形★＝A1／A1:B2／A:A／1:1／Sheet1!A1（→TRUE）と、'
          + '数／文字／TRUE／FALSE（→FALSE）'
          + '（2026-09-08 に 実Excel の 16本と 突き合わせて ★16本 全部 一致★）。'
          + '★関数が 入っている 形は まだ★＝#NAME? に なる'
          + '（=ISREF(INDEX(...))／★シート名に かっこが 在る 物（\'売上(旧)\'!A1）も★）。'
          + '★指せたか どうかが 字だけでは 決まらない為（INDIRECT("A1") は TRUE／INDIRECT("zzz") は FALSE）。'
          + '間違った TRUE/FALSE を 出すより「まだ」と 言う（2026-09-07 実測 25本）',
        'YEN': '★打てば そのまま 動く★＝Exally が ★DOLLAR★（本名）に 直して 答える。'
          + '=YEN(1234.5) → ¥1,235。'
          + '★YEN は 日本語UIの 表示名★で、ファイル／英語の 構文では DOLLAR'
          + '（2026-09-05 実Excel 実測＝.FormulaLocal では ¥1,234 が 出る／'
          + '.Formula では #NAME?／直している 場所＝exally-formula.js の convertFormula）',
      },
      足さない: {
        /* ★★2026-09-06 実測＝ここに 書いてあった 5個は 動いていた★★
           LET／MAP／REDUCE／SCAN／MAKEARRAY ⇒ 下の「別名で動く」へ 移した
           ★なぜ 見えなかったか★＝私が ★引数の 形を 手で 決め打っていた★
             LAMBDA を 取る 関数は ★普通の 引数では 一度も 当たらない★
           ⇒ 見張り tests/kansuu-tana.test.mjs が engine に 押して 見る
           ★残る 3つは 押しても #NAME?★（実測） */
        /* ★★2026-09-19 ... LAMBDA / BYROW / BYCOL / PHONETIC を ★下の 棚へ 移しました★★★
             ★なぜ 今まで 気づかなかったか★
               見張り（kansuu-tana / kansuu-kabaa）が ★2段しか 押して いませんでした★
               ＝①JS層 と ③借り物 だけ／★②台（lib/shiki-hyou.js）を 見て いなかった★
               ⇒★本番では 動いて いるのに 「動かない」と 書いて いました★
               ⇒★2026-09-19 に 見張りを 3段に 揃えた その日に 4個 出ました★
             ★PHONETIC は 同じ日に 直しました★（範囲で #VALUE! だった） */
        'FIELDVALUE': '株価や 地図の カードから 項目を 引く 物。そのカードの 仕組みが 無い',
        'RTD': '向こうから 押してくる 相手に 繋ぐ 物。★その 相手（配信の 元）が うちには 無い★＝繋ぐ 先が 無い',
        /* ★★2026-09-07＝★式は 書いたのに 出さないと 決めた 4個★★★
           ★ここに 書く 訳（指示役）★
             書かないと お客さんは ★#NAME? を 食らうだけ★で
             ★「無い」のか「まだ」なのか 分からない★
             ⇒★「まだ」と「訳」と「いつ 出すか」の 3つを 書く★
           ★決めた 訳★＝実Excel の 答え 700本と 突き合わせたら
             ★合う 形と 合わない 形が 同じ 関数の 中で 分かれた★
             ⇒★合う 形だけ 出す事が 出来ない（引数で 変わる）★
             ⇒★半分 合う 計算は 出さない★（債券の 値段＝★お金が 動く★）
           ★調べた 記録★ … docs/measured/kansuu46/README.md */
        /* ★★ODDFPRICE / ODDFYIELD は 2026-09-19 に 棚から 降ろしました★★
             ＝経営者1 の ★414本で 外れ 0本★（向こうの 分母・向こうの 道具）
             ＝★わざと 壊したら 53 / 414★＝★緑が 本物★
             ＝私の 紙でも 48/48 と 24/24 ／ awaseru 700/700 ／ 総なめ 354本 赤0
             ⇒★「日数の 数え方 2・3 や 準期間 3つ以上で ずれる」は ★直りました★★
             ⇒★★これで この 棚は 空です＝保留 0個★★ */
        /* ★★ODDLPRICE / ODDLYIELD は 2026-09-19 に 棚から 降ろしました★★
             ＝経営者1 の ★275本で 外れ 0本★（向こうの 分母・向こうの 道具）
             ＝★わざと 壊したら 258本 赤★＝★緑が 本物★
             ＝私の 紙でも 12/12 と 12/12
             ⇒★「3期以上で ずれる」は ★直りました★★
               ＝準の並びを 最終利払日から 1つずつ に した
               ＝NL を その 期の 日数（basis ごと）に した
               ＝DSC を 期の 残り ＋ 後ろの 期 ＋ 端 に した
             ⇒★棚に 残すと ★お客さんに 嘘を 言います★★（`kansuu-tana` が 捕まえた） */
        /* ★★2026-09-07（その2）＝★お客さんに「まだ」と 言える ように 書く★★★
           ★どれも 実Excel に 打たせて 答えを 取った上で 出さないと 決めた★
             ⇒ docs/measured/kansuu46/golden-yosoku-2026-09-07.tsv */
        'TRIMRANGE': '範囲の 端の 空きを 落とす 物。'
          + ' 範囲の 大きさを その場で 変える 作りが エンジンに まだ 無い（2026-09-07 実測）',
        'FORECAST.ETS': '季節を 見つけて 先を 読む 物。'
          + ' 実Excel と 同じ 答えに するには 季節の 見つけ方まで 揃える 必要が 有り、'
          + ' まだ 揃っていない。近い 答えを 出すと ★予測が 少し ずれる★ ので 出さない（2026-09-07 実測）',
        'FORECAST.ETS.CONFINT': '先を 読んだ 時の 幅。季節の 見つけ方まで 揃わないと ずれる（2026-09-07 実測）',
        'FORECAST.ETS.SEASONALITY': '季節の 長さを 見つける 物。見つけ方が 実Excel と 揃っていない（2026-09-07 実測）',
        'FORECAST.ETS.STAT': '先を 読む 時の 中の 数値。季節の 見つけ方まで 揃わないと ずれる（2026-09-07 実測）',
        /* ★★2026-09-07（その3）＝★動かない 物を 1つ残らず ここに 載せる★★★
           ★訳（指示役）★
             ここに 載っていない 関数は ★AI にとって「知らない 関数」★です。
             ⇒★知らないので 止めようが ない＝AI が 勧めて しまう★
             ⇒ お客さんが 打つ ⇒ `#NAME?` ⇒★「出来ると 言ったのに 出来ない」★
             ⇒★★動くのに「動かない」と 言うのと ★同じ 家の 裏返し★★★
           ★見張り★ … tests/ugokanai-osu.test.mjs が
             ★「動かない」と 数えた 物★と ★この 棚★が ★同じ 集まりか★を 毎回 数える */
        'CALL': 'パソコンの 中の 外部の 部品を 呼ぶ 物。ブラウザの 中には その 仕組みが 無い',
        'REGISTER.ID': 'パソコンの 中の 外部の 部品の 番号を 取る 物。ブラウザには その 仕組みが 無い',
        'CUBEVALUE': '分析用の 倉庫（キューブ）に 繋いで 値を 取る 物。繋ぐ 相手が 無い',
        'CUBEMEMBER': '分析用の 倉庫（キューブ）の 項目を 指す 物。繋ぐ 相手が 無い',
        'CUBEMEMBERPROPERTY': '分析用の 倉庫の 項目の 中身を 取る 物。繋ぐ 相手が 無い',
        'CUBERANKEDMEMBER': '分析用の 倉庫から 順位で 項目を 選ぶ 物。繋ぐ 相手が 無い',
        'CUBESET': '分析用の 倉庫の 項目の まとまりを 作る 物。繋ぐ 相手が 無い',
        'CUBESETCOUNT': '分析用の 倉庫の まとまりの 数を 数える 物。繋ぐ 相手が 無い',
        'CUBEKPIMEMBER': '分析用の 倉庫の 目標の 値を 取る 物。繋ぐ 相手が 無い',
        'GETPIVOTDATA': 'ピボットテーブルから 値を 引く 物。ピボットテーブル 自体が まだ 無い',
        'PIVOTBY': 'ピボットの 表を 式で 作る 物。ピボットテーブル 自体が まだ 無い',
        'GROUPBY': 'まとめて 集計する 表を 式で 作る 物。ピボットテーブル 自体が まだ 無い',
        'IMAGE': 'ウェブの 住所の 絵を セルに 出す 物。★絵を セルの 上に 置く 仕組み★が まだ 無い'
      }
    };
  }

  return {
    条件つき平均: 条件つき平均, 条件に合うか: 条件に合うか,
    切り出す: 切り出す, 選び出す: 選び出す, 一本にする: 一本にする,
    折り返す: 折り返す, 広げる: 広げる, 表を字に: 表を字に,
    最頻値たち: 最頻値たち,
    表にする: 表にする, 数える: 数える
  };
}));
