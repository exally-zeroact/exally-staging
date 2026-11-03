/* shoshiki.js — ★書式の 字を 読んで 数を 字に する 台★（2026-09-15）
 *
 *  ★★なぜ 1つに するか★★（指示役1 の 決め 2026-09-15）
 *    ★TEXT() も 画面も 同じ 書式を 解きます★。2か所で 解くと ★道が 2本★＝
 *    今日 だけで 3回 踏んだ 型です（見本／empty-ref-zero／書式）。
 *    ⇒★台は ここ 1つ★。★TEXT が 呼ぶ★／★画面（`js/book-open.js`）も ここを 呼ぶ★。
 *
 *  ★★作る 分は 測って 決めました★★（★全部 作ってから 測らない★）
 *    `docs/measured/golden-jitsubutsu-shoshiki-2026-09-15.tsv`（repo に 入れない 紙）
 *      ㋐ TEXT 740回 … ★aaa（731）／m/d（9）の 2種類だけ★
 *      ㋑ 画面 67,542マス … ★18種★（上位 6種で 95.6%／記号だけで 98.1%）
 *    ⇒★覚える 物は 6つ★
 *      ①`G/標準` ②`#,##0` と `0`（小数 0〜3桁）③`_x` の 詰め物
 *      ④`;` の 区切りと `[色]` ⑤`(` `)` を そのまま ⑥日付（m d yyyy aaa "字" ;@）
 *
 *  ★★答えは 全部 実Excel に 打って 測りました★★
 *    `docs/measured/kansuu46/golden-shoshiki-dai-2026-09-15.tsv`（192行）
 *    ★192行 中 162行は マスの 見た目と TEXT() が 同じ 字★
 *    ★違う 30行は 全部「その 値を その 書式で 出せない」時★
 *      ⇒★この 台は「出せない」を 返すだけ★
 *        ・TEXT()  … ★#VALUE!★ に する
 *        ・画面    … ★#### で 埋める★
 *      ⇒★どちらに するかは 呼ぶ側★＝★だから 台は 1つで 足ります★
 *
 *  ★★測って 決めた 細かい 所★★（当て推量なら 外して いた）
 *    ・`_x` … ★空きを 1つ★（`#,##0_ ` の 答えは ★後ろに 空白★）
 *    ・丸め … 0.5→1 ／ -0.5→-1 ／ 1.5→2 ／ ★-0.004→0（-0 に しない）★
 *    ・`;` の 2つ目（負）は ★-0.004 でも 使う★ … `#,##0_);[赤](#,##0)` → ★(0)★
 *    ・`#,##0` は 3桁区切り ／ ★`0.00` は 区切らない★
 *    ・`m/d` … 0 → ★1/0★（1900年の 0日）
 *
 *  ★★まだ 入れて いない 物★★（★出来て いない 物を 出来た 顔で 混ぜない★）
 *    ・`%` `E+` 分数 `[h]` `ggge`（和暦）… ★実物が 使って いません★
 *      ⇒★要る と 分かってから 足します★（実Excel に 打って から）
 *    ・色は ★名前を 返すだけ★（塗るのは 呼ぶ側）
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.Shoshiki = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  /* ★この国の 色の 名前★（`[赤]`）と 世界共通（`[Red]`） */
  var 色の名 = {
    '赤': 'Red', '青': 'Blue', '緑': 'Green', '黄': 'Yellow',
    '紫': 'Magenta', '水': 'Cyan', '黒': 'Black', '白': 'White',
    'RED': 'Red', 'BLUE': 'Blue', 'GREEN': 'Green', 'YELLOW': 'Yellow',
    'MAGENTA': 'Magenta', 'CYAN': 'Cyan', 'BLACK': 'Black', 'WHITE': 'White',
  };
  var 曜日 = ['日', '月', '火', '水', '木', '金', '土'];
  /* ★ddd・mmm は 日本語 UI でも 英語★（`exally-formula.js` の 頭の 実測・16.0.20228） */
  var 英曜日 = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  var 英月 = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];

  /* ══ ★書式を `;` で 区切る★ ══
     ★引用符と 角括弧の 中の `;` は 区切りでは ありません★ */
  function 区切る(書式) {
    var 出 = [], 今 = '', 中 = 0, i = 0;
    while (i < 書式.length) {
      var c = 書式.charAt(i);
      if (c === '"') {                          /* 引用符の 中は そのまま */
        var j = 書式.indexOf('"', i + 1);
        if (j < 0) j = 書式.length;
        今 += 書式.slice(i, j + 1); i = j + 1; continue;
      }
      if (c === '\\' || c === '_' || c === '*') { 今 += 書式.slice(i, i + 2); i += 2; continue; }
      if (c === '[') 中++;
      if (c === ']') 中--;
      if (c === ';' && 中 <= 0) { 出.push(今); 今 = ''; i++; continue; }
      今 += c; i++;
    }
    出.push(今);
    return 出;
  }

  /* ══ ★1つの 区画から 色と 中身を 分ける★ ══ */
  function 区画を読む(s) {
    var 色 = null;
    var t = String(s).replace(/\[([^\]]*)\]/g, function (全, 中) {
      var k = 色の名[中] || 色の名[String(中).toUpperCase()];
      if (k) { 色 = k; return ''; }
      /* ★★国の 番号は 捨てます★★（2026-09-15 実測で 見つけた）
         ＝`m/d\([$-411]aaa\)` の `[$-411]` は ★日本語で 出せ★の 印で、
           ★出る 字には なりません★。
         ★司さんの 実物で 721マス★＝★台と SheetJS の 違いの 721/730★が これでした。 */
      if (/^\$-/.test(中) || /^\$/.test(中)) return '';
      return 全;                                /* ★色でも 国でも ない 角括弧は 残す★（[h] 等） */
    });
    return { 色: 色, 字: t };
  }

  /* ══ ★★書式を 掛ける 前に 15桁に 丸める★★ ══（2026-09-15 に 台へ 入れた）
     ★実Excel は 15桁に 丸めてから 出します★（2026-09-11 に 実物で 見つけた 決まり）
       34779.49999999999 … 実Excel ★34,780★ ／ 丸めないと ★34,779★＝★お金が 1円 ずれる★
     ★前は 呼ぶ側が やって いました★（`book.html:3440` の `_十五桁`）
       ⇒★`js/book-open.js:378` には 入って いません★＝★同じ 事を する 道が 2本／片方だけ 直って いた★
     ⇒★台の 中に 入れました★＝★呼ぶ側が 忘れられません★ */
  function 十五桁(n) {
    if (typeof n !== 'number' || !isFinite(n) || n === 0) return n;
    var v = Number(n.toPrecision(15));
    return isFinite(v) ? v : n;
  }

  /* ══ ★数を 15桁に 丸めて 字に する★（`G/標準`）══
     ★実測★ 0.520833333333333 → ★0.520833333★（11桁で 切れる）
       ＝実Excel の `G/標準` は ★11桁ぶんしか 見せません★ */
  function 標準の字(n) {
    if (!isFinite(n)) return String(n);
    if (n === 0) return '0';
    var a = Math.abs(n);
    if (a >= 1e11 || (a < 1e-4 && a > 0)) {
      /* ★指数は 実物が 使って いません★＝★素の まま 出します★（棚） */
      return String(n);
    }
    /* ★★11字まで★★（実測 0.520833333333333 → ★0.520833333★＝★ちょうど 11字★）
       ＝★桁数では なく 字数★（小数点も 1字に 数える）
         1234.5 → 1234.500000 → 後ろの 0 を 落として ★1234.5★ */
    var 整数桁 = (a < 1) ? 1 : (Math.floor(Math.log10(a)) + 1);
    var 小数桁 = Math.max(0, 11 - 整数桁 - 1);
    var s = n.toFixed(Math.min(小数桁, 20));
    if (s.indexOf('.') >= 0) s = s.replace(/0+$/, '').replace(/\.$/, '');
    return s;
  }

  /* ══ ★数の 区画を 当てる★ ══ */
  function 数を当てる(n, 字) {
    /* ★小数の 桁数★＝小数点の 後ろの 0 と # の 数 */
    var 点 = 字.indexOf('.');
    var 小数桁 = 0;
    if (点 >= 0) {
      var 後 = 字.slice(点 + 1);
      for (var i = 0; i < 後.length; i++) {
        var c = 後.charAt(i);
        if (c === '0' || c === '#' || c === '?') 小数桁++;
        else if (c === '"') { var j = 後.indexOf('"', i + 1); i = (j < 0 ? 後.length : j); }
        else if (c === '\\' || c === '_' || c === '*') i++;
      }
    }
    var 区切りか = /#,#|0,0/.test(字.slice(0, 点 < 0 ? 字.length : 点));
    /* ★丸めは 0 から 遠い 方へ★（実測 0.5→1 ／ -0.5→-1 ／ 1.5→2） */
    var m = Math.pow(10, 小数桁);
    var r = Math.round(Math.abs(n) * m + 1e-9) / m;
    var s = r.toFixed(小数桁);
    var 整数 = s, 小数 = '';
    var p = s.indexOf('.');
    if (p >= 0) { 整数 = s.slice(0, p); 小数 = s.slice(p + 1); }
    if (区切りか) 整数 = 整数.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    /* ★整数の 所に 0 が 何個 要るか★（`0000000` の ような 埋め） */
    var 整数型 = 字.slice(0, 点 < 0 ? 字.length : 点);
    var 要る0 = (整数型.match(/0/g) || []).length;
    while (整数.replace(/,/g, '').length < 要る0) 整数 = '0' + 整数;
    /* ★整数が 0 で `#` しか 無い 時は 何も 出さない★（実測 `=TEXT(0,"#")` → 空） */
    if (Number(s) === 0 && 要る0 === 0 && /[#?]/.test(整数型)) 整数 = '';
    return { 整数: 整数, 小数: 小数, 小数桁: 小数桁 };
  }

  /* ══ ★通し番号を 日に する★（`lib/formula-soto.js` と 同じ 形）══ */
  function 数から日(n) {
    var d = Math.floor(n);
    var t = n - d;
    /* ★★1900年の 起点は 2つ 在ります★★（2026-09-15 実測で 直した）
       ＝実Excel には ★1900-02-29 という 無い 日★が 在ります（通し 60）。
         通し 1〜59  … ★1899-12-31 から★  （1 → 1900-01-01）
         通し 61〜   … ★1899-12-30 から★  （61 → 1900-03-01）
       ★前は 全部 1899-12-30 から★に して ★1月が 12月に なって いました★
         （実測 `m/d` で 通し 0 → 紙 ★1/0★ ／ うち ★12/0★）
       ★通し 0 は「1900年 1月 0日」★＝実Excel が そう 見せます（実測 `m/d` → 1/0） */
    var 基 = (d < 60) ? Date.UTC(1899, 11, 31) : Date.UTC(1899, 11, 30);
    var v = new Date(基 + d * 86400000);
    /* ★★通し 60 は ★1900-02-29★＝本当の 暦には 無い 日★★（2026-09-16 実測）
         `=TEXT(60,"yyyy-mm-dd")` → 実Excel ★1900-02-29★ ／ うちは ★1900-02-28★
         `=DAY(60)`               → 実Excel ★29★         ／ うちは ★28★
       ＝★起点を 2つに 分けても 60 だけは どちらでも 出せません★
         （59 の 次の 日で、61 の 前の 日＝★2つの 起点の 境目そのもの★）
       ⇒★ここだけ 名指しで 返します★（曜日は 通し番号の 割り算なので そのまま 合う） */
    var 二月二十九 = (d === 60);
    var 秒 = Math.round(t * 86400);
    /* ★★曜日は 通し番号の 割り算★★（実測 … 通し 0→土 ／ 1→日 ／ 45292→月）
       ＝★本当の 暦の 曜日では ありません★（無い 日の 分 ずれて います）
       ⇒★実Excel と 同じに する には 通し番号で 数える★ */
    var 曜 = ((d % 7) + 6) % 7;
    return {
      y: (d === 0) ? 1900 : (二月二十九 ? 1900 : v.getUTCFullYear()),
      m: (d === 0) ? 1 : (二月二十九 ? 2 : v.getUTCMonth() + 1),
      d: (d === 0) ? 0 : (二月二十九 ? 29 : v.getUTCDate()),
      w: 曜,
      h: Math.floor(秒 / 3600), mi: Math.floor(秒 % 3600 / 60), s: 秒 % 60,
      通し: d,
    };
  }
  function 埋2(x) { return (x < 10 ? '0' : '') + x; }

  /* ══ ★日付の 区画を 当てる★ ══ */
  /* ══ ★★元号★★（2026-10-02）══
       ★記憶では 書いて いません★＝実Excel 16.0.20430 の 真値（境目の 前日と 当日を 全部 打った）
         1912/7/29 M45 → 7/30 T1 ／ 1926/12/24 T15 → 12/25 S1
         1989/1/7  S64 → 1/8  H1 ／ 2019/4/30  H31 → 5/1   R1 ／ 1900/1/1 M33
       ★元年は 「元」 では なく 1★（TEXT も マスの 書式も）
       `始` ＝その 元号の 最初の 日（年・月・日）／`引く` ＝西暦から 引く 数 */
  var 元号 = [
    { 始: [2019, 5, 1], 引く: 2018, g: 'R', gg: '令', ggg: '令和' },
    { 始: [1989, 1, 8], 引く: 1988, g: 'H', gg: '平', ggg: '平成' },
    { 始: [1926, 12, 25], 引く: 1925, g: 'S', gg: '昭', ggg: '昭和' },
    { 始: [1912, 7, 30], 引く: 1911, g: 'T', gg: '大', ggg: '大正' },
    { 始: [0, 1, 1], 引く: 1867, g: 'M', gg: '明', ggg: '明治' },
  ];
  function 元号を引く(日) {
    var 鍵 = 日.y * 10000 + 日.m * 100 + 日.d;
    for (var i = 0; i < 元号.length; i++) {
      var s = 元号[i].始;
      if (鍵 >= s[0] * 10000 + s[1] * 100 + s[2]) return 元号[i];
    }
    return 元号[元号.length - 1];
  }

  /* ══ ★★経過時間（[h] [m] [s]）★★（2026-10-02・実Excel の 真値 56本）══
       ★丸めは 秒で 1回だけ★、大きい 単位へは ★切り捨て★で 割る
         0.999999 日 ＝ 86,399.91秒 ⇒ `[h]:mm` ★24:00★ ／ `[h]` ★24★（秒で 丸めて 86,400）
         1/24/60*59.5/60 ＝ 59.4999...秒（浮動小数）⇒ `[h]:mm:ss` ★0:00:59★ ／ `[h]:mm` ★0:00★
       ★`ss.00` が 在れば 1/100秒で 丸める★ ⇒ 0.999999 は ★23:59:59.91★（繰り上がらない）
       ★負は `[s]` `[h]` の 1つだけの 書式なら 符号付き（-129600 ／ -36）★
         それ以外は ★出せない★（TEXT は #VALUE!／マスは ####・実測） */
  /** ★★和暦か 経過時間の 書式か★★（2026-10-04・★判じは ここ 1つ★＝TEXT と 画面の 両方が 呼ぶ）
   *  ★当たらない（今の 道に 残る）★ ... 引用符の 中・逃がした 字（\g）・_x・*x の 字／
   *    `G/標準` `General` の G・e ／ 指数の E（`0.00E+00` `0e-0`）／
   *    角括弧でも h/m/s で ない 物（`[Red]` `[$-411]` `[>=100]` `[DBNum1]`）／角括弧の 無い `h:mm`
   *  ★当たる★ ... 外に 出た `g` `e`（和暦）／ `[h]` `[mm]` `[ss]` 等（経過） */
  function 和暦か経過か(書式) {
    var s = String(書式 == null ? '' : 書式)
      .replace(/"[^"]*"/g, '')                 /* 引用符の 中 */
      .replace(/[\\_*]./g, '')                 /* 逃がし・詰め物・繰り返しの 字 */
      .replace(/G\/標準|General/gi, '');        /* 標準の 名前 */
    if (/\[(h+|m+|s+)\]/i.test(s)) return true;
    s = s.replace(/\[[^\]]*\]/g, '')            /* 色・国・条件 等の 角括弧 */
      .replace(/[eE][+-]/g, '');                /* 指数 */
    return /[gGeE]/.test(s);
  }

  function 経過の書式か(字) {
    return /\[(h+|m+|s+)\]/i.test(String(字).replace(/"[^"]*"/g, ''));
  }

  function 日を当てる(n, 字) {
    var 経過 = 経過の書式か(字);
    var 負 = false;
    if (n < 0) {
      /* ★負は 日に できません★（実測 → #### ／ #VALUE!）
         ★但し `[s]` `[h]` `[m]` だけの 書式は 符号付きで 出ます★（実測 -1.5 → -129600 ／ -36） */
      var 裸の字 = String(字).replace(/"[^"]*"/g, '');
      if (!(経過 && /^\s*\[(h+|m+|s+)\]\s*$/i.test(裸の字))) return null;
      負 = true; n = -n;
    }
    /* ══ ★★日付の 上限★★（2026-10-04・実Excel 16.0.20430・経営者 取った）══
         `=1E+21` ... `[h]:mm:ss` `[h]:mm:ss.00` `ge.m.d` は ★####★／★`[s]` は 字★（86400000000000000000000000）
         ⇒★単位 1つだけの 経過（`[s]` 等）は 上限を 持たない★（負の 時と 同じ 形）
         ⇒★単位 1つで ない 日付・時刻の 書式は 全部 9999/12/31（通し 2958465）までで 切る★
         ★境目も 測った★（10-04 夕・経営者）＝2958465.99999 までは 字（h:mm 23:59 ／ [h]:mm:ss 71003183:59:59・繰り上がらない）
           2958466 から ####：和暦・[h]:mm:ss・[m]:ss・yyyy/m/d・h:mm・m/d・d・yyyy・mmm・mmmm・aaa・aaaa・ddd・mm:ss・yyyy/m/d h:mm
           ★単位 1つの [h]・[s] は 上限 なし★（2958466 の [h] → 71003184）。★[m] だけは 未測定★
         ★前は 和暦・経過の 時だけ 切って いた★＝真値で yyyy/m/d 等も 同じ 上限と 分かった ので 広げた */
    var 単位一つ = 経過 && /^\s*\[(h+|m+|s+)\]\s*$/i.test(String(字).replace(/"[^"]*"/g, ''));
    if (n >= 2958466 && !単位一つ) return null;
    var 日 = 数から日(n);
    var 経 = null;
    if (経過) {
      var 小桁m = /s+\.(0+)/i.exec(字);
      var 小桁 = 小桁m ? 小桁m[1].length : 0;
      var 刻 = Math.pow(10, 小桁);
      var 全 = Math.round(n * 86400 * 刻);
      var 秒全 = Math.floor(全 / 刻);
      経 = {
        時全: Math.floor(秒全 / 3600), 分全: Math.floor(秒全 / 60), 秒全: 秒全,
        分: Math.floor(秒全 % 3600 / 60), 秒: 秒全 % 60,
        小: 全 % 刻, 小桁: 小桁,
        時か: /\[h+\]/i.test(字),
      };
    }
    /* ★大きい 数は 15桁＋0 で 出す★（実Excel `=1E+21` の `[s]` → 86400000000000000000000000）
       ＝String(8.64e25) は 指数の 形 '8.64e+25' に なる ⇒ 15桁を 取り 残りを 0 で 埋める */
    function 全桁(x) {
      if (Math.abs(x) < 1e15) return String(x);
      var m = /^(\d)(?:\.(\d+))?e\+(\d+)$/.exec(Math.abs(x).toPrecision(15));
      if (!m) return String(x);
      var 桁 = m[1] + (m[2] || ''), 指 = Number(m[3]);
      while (桁.length < 指 + 1) 桁 += '0';
      return 桁.slice(0, 指 + 1);
    }
    function 埋(x, 長) { var s = 全桁(x); while (s.length < 長) s = '0' + s; return s; }
    var 午前午後 = /AM\/PM|A\/P/i.test(String(字).replace(/"[^"]*"/g, ''));
    /* ★通し 0 は 1900年の 0日★（実測 `m/d` → 1/0） */
    var 出 = 負 ? '-' : '', i = 0;
    while (i < 字.length) {
      var c = 字.charAt(i);
      if (c === '"') { var j = 字.indexOf('"', i + 1); if (j < 0) j = 字.length; 出 += 字.slice(i + 1, j); i = j + 1; continue; }
      if (c === '\\') { 出 += 字.charAt(i + 1) || ''; i += 2; continue; }
      if (c === '_') { 出 += ' '; i += 2; continue; }
      if (c === '@') { i++; continue; }
      /* ★経過の 単位★（`[h]` `[hh]` `[m]` `[mm]` `[s]` `[ss]`） */
      var 括 = 経 ? /^\[(h+|m+|s+)\]/i.exec(字.slice(i)) : null;
      if (括) {
        var u = 括[1].charAt(0).toLowerCase(), 長u = 括[1].length;
        出 += 埋(u === 'h' ? 経.時全 : (u === 'm' ? 経.分全 : 経.秒全), 長u);
        i += 括[0].length; continue;
      }
      /* ★午前・午後★ `AM/PM` `A/P`（10-04 夕・実Excel：2958465 → 12:00 AM ／ .99999 → 11:59 PM） */
      if (/^AM\/PM/i.test(字.slice(i))) { 出 += (日.h < 12 ? 'AM' : 'PM'); i += 5; continue; }
      if (/^A\/P/i.test(字.slice(i))) { 出 += (日.h < 12 ? 'A' : 'P'); i += 3; continue; }
      /* ★和暦★ `g` `gg` `ggg` ／ `e` `ee` */
      var 和 = /^(g+|e+)/i.exec(字.slice(i));
      if (和) {
        var 号 = 元号を引く(日), 長w = 和[0].length;
        if (和[0].charAt(0).toLowerCase() === 'g') 出 += (長w >= 3 ? 号.ggg : (長w === 2 ? 号.gg : 号.g));
        else 出 += 埋(日.y - 号.引く, 長w >= 2 ? 2 : 1);
        i += 長w; continue;
      }
      var 続 = /^(y+|m+|d+|h+|s+|a+)/i.exec(字.slice(i));
      if (続) {
        var w = 続[0], k = w.charAt(0).toLowerCase(), 長 = w.length;
        /* ★経過の 書式の 中の `mm` `ss` は 「その 単位の 中の 余り」★ */
        if (経 && k === 'm' && 経.時か) { 出 += (長 >= 2 ? 埋2(経.分) : String(経.分)); i += 長; continue; }
        if (経 && k === 's') {
          出 += (長 >= 2 ? 埋2(経.秒) : String(経.秒));
          i += 長;
          if (経.小桁 && 字.charAt(i) === '.') { 出 += '.' + 埋(経.小, 経.小桁); i += 1 + 経.小桁; }
          continue;
        }
        if (k === 'y') 出 += (長 >= 3 ? String(日.y) : String(日.y % 100));
        /* ★ddd・dddd は 英語の 曜日★（日本語 UI でも 英語＝`exally-formula.js` の 頭の 実測 16.0.20228 ／
           10-04 夕 経営者の 真値 2958465 の ddd → Fri） */
        else if (k === 'd' && 長 >= 3) 出 += (長 === 3 ? 英曜日[日.w].slice(0, 3) : 英曜日[日.w]);
        else if (k === 'd') 出 += (長 >= 2 ? 埋2(日.d) : String(日.d));
        /* ★mmm 以上は いつも 月の 名★（英語＝同じ 実測／真値 mmm → Dec ／ mmmm → December ／ mmmmm は 頭の 1字） */
        else if (k === 'm' && 長 >= 3) 出 += (長 === 3 ? 英月[日.m - 1].slice(0, 3) : (長 === 4 ? 英月[日.m - 1] : 英月[日.m - 1].charAt(0)));
        else if (k === 'm') {
          /* ★★`m` は 2つの 意味を 持ちます★★（2026-09-15 実測で 直した）
             ★時の 隣なら「分」／それ以外は「月」★
               `=TEXT(0.520833…,"h:mm")` … 実Excel ★12:30★ ／ 直す前の うち ★12:01★
               （★01 は 1月★＝★月を 出して いました★）
             ★見分け方★ … ★前に `h` が 在る★ か ★後ろに `s` が 在る★ */
          var 前は時 = /h+[^ymdhsa"]*$/i.test(字.slice(0, i));
          var 後は秒 = /^[^ymdhsa"]*s+/i.test(字.slice(i + 長));
          if (前は時 || 後は秒) 出 += (長 >= 2 ? 埋2(日.mi) : String(日.mi));
          else 出 += (長 >= 2 ? 埋2(日.m) : String(日.m));
        }
        else if (k === 'h') {
          /* ★AM/PM が 在れば 12時間★（0時 → 12） */
          var 時 = 午前午後 ? (日.h % 12 === 0 ? 12 : 日.h % 12) : 日.h;
          出 += (長 >= 2 ? 埋2(時) : String(時));
        }
        else if (k === 's') 出 += (長 >= 2 ? 埋2(日.s) : String(日.s));
        else if (k === 'a') 出 += (長 >= 4 ? 曜日[日.w] + '曜日' : 曜日[日.w]);
        i += 長; continue;
      }
      出 += c; i++;
    }
    return 出;
  }

  /** ★★日付・時刻の 書式か★★（2026-10-04 夕・★判じは ここ 1つ★＝画面の fmtForDisplay が 呼ぶ）
   *  数の 区画（`@` で ない 区画）の どれかが 日付の 字なら 真。
   *  ★偽★ ... General・G/標準・指数の E（0.00E+00）・引用符の 中・逃がし・[Red]・[$-411]・
   *         [>=100]・[DBNum1]（★h/m/s で ない 角括弧は 先に 外す★）・#,##0"円"・0%・@ */
  function 日付の書式か(書式) {
    var f = String(書式 == null ? '' : 書式);
    if (f === '' || /^(General|G\/標準|標準)$/i.test(f)) return false;
    if (/[eE][+-]/.test(f.replace(/"[^"]*"/g, '').replace(/\\./g, ''))) return false;   /* 指数 */
    var 区 = 区切る(f).filter(function (x) { return !/@/.test(x.replace(/"[^"]*"/g, '')); });
    for (var i = 0; i < 区.length; i++) {
      var r = 区画を読む(区[i]);
      var s = r.字.replace(/\[(?!(?:h+|m+|s+)\])[^\]]*\]/gi, '');
      if (日付の字か(s)) return true;
    }
    return false;
  }

  /* ★日付の 書式か★＝`y m d h s a` が 引用符の 外に 在る */
  function 日付の字か(字) {
    var i = 0;
    字 = String(字).replace(/G\/標準/g, '');   /* ★`G/標準` の G は 元号では ない★ */
    while (i < 字.length) {
      var c = 字.charAt(i);
      if (c === '"') { var j = 字.indexOf('"', i + 1); i = (j < 0 ? 字.length : j + 1); continue; }
      if (c === '\\' || c === '_' || c === '*') { i += 2; continue; }
      if (/[yYdDhHsSaAgGeE]/.test(c)) return true;   /* ★g・e＝和暦（2026-10-02）★ */
      if (c === 'm' || c === 'M') return true;
      i++;
    }
    return false;
  }

  /* ══ ★数の 区画の 字を 組む★ ══ */
  function 数の字を組む(n, 字) {
    /* ★`%` は 100倍★（実測 `=TEXT(0.5,"0%")` → ★50%★）
       ★`%` の 数だけ 掛けます★（`0%%` は 10000倍＝実Excel の 決まり／★未測定★） */
    var 百 = (字.replace(/"[^"]*"/g, '').match(/%/g) || []).length;
    if (百) n = n * Math.pow(100, 百);
    var 部 = 数を当てる(n, 字);
    var 出 = '', i = 0, 数を出した = false;
    while (i < 字.length) {
      var c = 字.charAt(i);
      if (c === '"') { var j = 字.indexOf('"', i + 1); if (j < 0) j = 字.length; 出 += 字.slice(i + 1, j); i = j + 1; continue; }
      if (c === '\\') { 出 += 字.charAt(i + 1) || ''; i += 2; continue; }
      /* ★`_x` は 空きを 1つ★（実測 `#,##0_ ` の 答えは 後ろに 空白） */
      if (c === '_') { 出 += ' '; i += 2; continue; }
      if (c === '*') { i += 2; continue; }        /* ★埋め文字は 画面の 話★＝1つも 出さない */
      /* ★`@` は「素の 字」★（実測 `=TEXT(1,"@")` → ★1★） */
      if (c === '@') { 出 += 標準の字(n); i++; continue; }
      if (c === '0' || c === '#' || c === '?' || c === ',' || c === '.') {
        if (!数を出した) {
          出 += 部.整数 + (部.小数桁 ? '.' + 部.小数 : '');
          数を出した = true;
        }
        /* ★数の 形の 字は 読み飛ばす★（もう 出した） */
        while (i < 字.length && /[0#?,.]/.test(字.charAt(i))) i++;
        continue;
      }
      出 += c; i++;
    }
    return 出;
  }

  /**
   * ★書式を 当てる★
   * @param {number|string|boolean|null} 値
   * @param {string} 書式  … ★この国の 字★（`G/標準` `[赤]` …）
   * @returns {{出せる:boolean, 字:string, 色:string|null}}
   *   ★出せる false★＝★その 値を その 書式で 出せません★
   *     ⇒ TEXT() は ★#VALUE!★／画面は ★####★（★決めるのは 呼ぶ側★）
   */
  function 当てる(値, 書式) {
    var f = String(書式 == null ? '' : 書式);
    /* ★字・真偽は 素通り★（実測 `=TEXT("あ","0")` → あ ／ `=TEXT(TRUE,"0")` → TRUE） */
    if (typeof 値 === 'string') {
      var 区 = 区切る(f);
      var 字区 = 区.length >= 4 ? 区[3] : null;   /* 4つ目＝字の 区画 */
      if (字区 != null && /@/.test(字区)) {
        var y = 区画を読む(字区);
        return { 出せる: true, 字: y.字.replace(/@/g, 値).replace(/"/g, ''), 色: y.色 };
      }
      return { 出せる: true, 字: String(値), 色: null };
    }
    if (typeof 値 === 'boolean') return { 出せる: true, 字: 値 ? 'TRUE' : 'FALSE', 色: null };
    if (値 == null || 値 === '') 値 = 0;          /* ★空マスは 0★（実測 `=TEXT(A9,"0")` → 0） */
    var n = 十五桁(Number(値));
    if (!isFinite(n)) return { 出せる: false, 字: '', 色: null };

    if (f === '' ) return { 出せる: true, 字: '', 色: null };   /* 実測 `=TEXT(1234.5,"")` → 空 */
    if (f === 'G/標準' || f === 'General' || f === '標準') {
      return { 出せる: true, 字: 標準の字(n), 色: null };
    }
    /* ★`_` の 後ろに 字が 無い 書式は 通りません★（実測 `#,##0.0_` → #VALUE! ／ 画面は 素の 数） */
    if (/_$/.test(f)) return { 出せる: false, 字: '', 色: null };

    /* ★★まだ 作って いない 形は「出せない」と 言います★★（2026-09-15）
       ★半分 合う 字を 出す 方が 悪い★＝★#VALUE! の 方が まし★
         `ggge` `ge`（和暦）／`[h]` `[m]` `[s]`（24時間を 超える）／`?/?`（分数）／`E+`（指数）
       ★司さんの 実物は 1つも 使って いません★（TEXT 740回＝aaa と m/d だけ／画面 18種にも 無い）
       ⇒★要ると 分かったら 実Excel に 打って から 足します★（棚）
       ★実測の 答えは 紙に 在ります★
         `=TEXT(45292,"ggge年m月d日")` → 令和6年1月1日 ／ `=TEXT(1.5,"[h]:mm")` → 36:00
         `=TEXT(1.5,"# ?/?")` → 1 1/2 */
    /* ★★和暦（g・e）と 経過時間（[h] [m] [s]）は 2026-10-02 に 足しました★★
         ＝実Excel 16.0.20430 の 真値 155本（`tests/xlsx-harness/cases/61-text-wareki-keika.json`）
         ★分数・指数は まだ 棚★ */
    var 裸 = f.replace(/"[^"]*"/g, '');
    if (/\?\s*\/\s*\?|\d\s*\/\s*\?/.test(裸)) return { 出せる: false, 字: '', 色: null };
    if (/[eE][+-]/.test(裸)) return { 出せる: false, 字: '', 色: null };

    /* ★★`@` の 区画は「字の 区画」★★＝★数には 使いません★（2026-09-15 実測で 直した）
       `m/d;@` は ★2つに 見えて 数の 区画は 1つ★です。
       ★前は 負の 数に `@` を 当てて `@` と 出して いました★
       （実測 … 負の 数 × 日付の 書式 → ★マスは ####／TEXT は #VALUE!★＝★出せない★） */
    var 区2 = 区切る(f).filter(function (x) { return !/@/.test(x.replace(/"[^"]*"/g, '')); });
    if (!区2.length) {
      /* ★書式が `@` だけ＝「素の 字に して」★（実測 `=TEXT(1,"@")` → ★1★）
         ＝★数の 区画が 1つも 無い★ので 標準の 字を そのまま 返します */
      return { 出せる: true, 字: 標準の字(n), 色: null };
    }
    /* ★区画の 選び方★（実測 `-0.004` は ★負の 側★＝`(0)`） */
    var 選, 負を直すか = false;
    if (区2.length === 1) { 選 = 区2[0]; 負を直すか = false; }
    else if (n < 0 && 区2.length >= 2) { 選 = 区2[1]; 負を直すか = true; }
    else if (n === 0 && 区2.length >= 3) { 選 = 区2[2]; }
    else { 選 = 区2[0]; }
    var r = 区画を読む(選);
    if (r.字 === '') return { 出せる: true, 字: '', 色: r.色 };

    if (日付の字か(r.字)) {
      var d = 日を当てる(n, r.字);
      if (d === null) return { 出せる: false, 字: '', 色: r.色 };
      return { 出せる: true, 字: d, 色: r.色 };
    }
    /* ★負の 区画が 別に 在るなら 符号は 付けない★（実測 `(1,235)` に `-` は 付かない） */
    var v = 負を直すか ? Math.abs(n) : n;
    var s = 数の字を組む(Math.abs(v), r.字);
    /* ★★丸めて 0 に なったら 符号を 付けない★★（実測 2026-09-15）
       `=TEXT(-0.004,"#,##0_ ")` … 実Excel ★"0 "★ ／ ★"-0 " では ない★
       ⇒★`-0` を 客に 見せない★（★これは 当て推量では 出せませんでした★） */
    var 零か = !/[1-9]/.test(s);
    if (!負を直すか && v < 0 && s !== '' && !零か) s = '-' + s;
    return { 出せる: true, 字: s, 色: r.色 };
  }

  /* ★TEXT() が 呼ぶ★＝出せない なら #VALUE!
     ★★TEXT は「世界共通の 字」を 受けません★★（実測 2026-09-15）
       `=TEXT(1234.5,"General")` → ★#VALUE!★ ／ `=TEXT(1234.5,"G/標準")` → 通る
     ⇒★画面は 受ける（NumberFormat が 世界共通で 来る）／TEXT は 受けない★
       ＝★同じ 台でも 入口で 分かれます★（★台を 2つに する 訳には なりません★） */
  function 字にする(値, 書式) {
    if (String(書式) === 'General') return null;
    var r = 当てる(値, 書式);
    return r.出せる ? r.字 : null;                /* ★null＝呼ぶ側が #VALUE! に する★ */
  }

  /* ★画面が 呼ぶ★＝出せない なら 幅ぶんの `#` */
  function 画面の字(値, 書式, 幅) {
    var r = 当てる(値, 書式);
    if (r.出せる) return { 字: r.字, 色: r.色 };
    var n = Math.max(1, Math.floor(幅 || 10));
    return { 字: new Array(n + 1).join('#'), 色: r.色 };
  }

  return {
    当てる: 当てる, 字にする: 字にする, 画面の字: 画面の字, 和暦か経過か: 和暦か経過か, 日付の書式か: 日付の書式か,
    区切る: 区切る, 標準の字: 標準の字, 数から日: 数から日,
  };
});
