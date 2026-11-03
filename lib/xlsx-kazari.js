/* xlsx-kazari.js — ★マスの 飾り（太字・字の色・塗り・罫線）を ★自前で★ 読む★（2026-09-21）
 *
 *  ★★なぜ 自前か★★
 *    ★借り物（SheetJS 0.20.3）は 飾りを ほとんど 捨てます★（2026-09-21 実測）
 *      `cellStyles: true` を 付けても
 *        ・`c.s` ＝ ★塗りだけ★（`{patternType:'solid', fgColor:{rgb:'FFFF00'}}`）
 *        ・`wb.Styles.Borders` ＝ ★10個 とも `{}`★（★罫線が 空★）
 *        ・マスが どの `cellXf` を 指すかの ★番号を 捨てます★
 *      ⇒★太字・字の色・罫線は 1つも 来ません★
 *    ⇒生の `xl/styles.xml` には ★ちゃんと 在ります★（同じ日 実測）
 *        `<font><b/><sz val="11"/><color rgb="FFFF0000"/>＝`
 *        `<border><left style="medium">＝<bottom style="thick">＝`
 *    ⇒★借り物を 増やさず 外す 向きに 進みます★
 *
 *  ★★何が 出来て いなかったか★★（`golden-kazari-gamen-made-2026-09-21.tsv`）
 *    実Excel が 作った 飾り付きの ファイルを お客さんの 道で 開いて 数えたら
 *    ★届いた 7 / 13★ で、★太字・字の色・罫線2つ・塗り・図形★ が 画面に 出て いませんでした。
 *
 *  ★★この 本が する 事★★
 *    `読む(シートのxml, stylesのxml)` ⇒ `{ 'A1': { bold, color, bgColor, border }, ＝ }`
 *    ・`bold`    ＝ true/false
 *    ・`color`   ＝ `'#FF0000'`（字の 色）
 *    ・`bgColor` ＝ `'#FFFF00'`（塗り）
 *    ・`border`  ＝ `{ top:1|2, bottom:1|2, left:1|2, right:1|2 }`
 *                  ★1＝細い ／ 2＝太い★（`book.html:drawBorder` が そう 描きます）
 *    ★台（`sheets[].data`）が 前から 持って いる 形に 合わせて います★
 *      ＝★新しい 持ち方を 作りません★（画面の 描き手を 1行も 変えずに 出ます）
 *
 *  ★★正規表現を 使いません★★
 *    `indexOf` だけで 切ります。
 *    ＝2026-09-21 に ★逆斜線の 逃がしが 途中で 落ちる★のを 8回 踏みました。
 *    ＝★この 本は 貝殻を 通して 直す 事が 有り得る★ので 最初から 避けます。
 *
 *  ★★測って いない 事★★
 *    ・`<color theme="＝"/>` ＝ ★テーマの 色表を 読んで いません★
 *        ⇒`rgb` が 一緒に 書いて あれば そちらを 使い、無ければ ★付けません★
 *        ⇒★当て推量で 黒を 入れません★（間違った 色より 色なしの 方が まし）
 *    ・`<color indexed="＝"/>` ＝ ★64 は 「自動」★＝★付けません★
 *        ⇒それ以外の 番号は ★未測定★＝付けません
 *    ・斜めの 罫線（`<diagonal>`）＝ ★台に 持ち方が 有りません★＝読みません
 *    ・`applyFont` / `applyFill` / `applyBorder` ＝ ★見て いません★
 *        ⇒`cellXfs` の `fontId`/`fillId`/`borderId` を そのまま 使って います
 *        ⇒★実Excel が これらを どう 使い分けるかは 測って いません★
 *
 *  ★★名前を `読む` に しない★★（2026-09-21）
 *    `tests/unused-param.test.mjs` は ★関数の 名前で 呼び出しを 数えます★。
 *    `読む` は この repo に ★たくさん 在る 名前★ なので、
 *    3つ目の 引数を 足した 日に ★「誰も 渡して いない」と 誤報★ されました。
 *    ⇒`飾りを読む` に しました。★`lib/hairanai.js` の 頭にも 同じ 断りが 在ります★
 *    ⇒経営者1 も 同じ 日に `数える` で 踏んで います（★2人で 2回★）。
 *
 *  見張り: tests/xlsx-kazari.test.mjs
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.XlsxKazari = factory();
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  /** ★`<名 ＝>＝</名>` と `<名 ＝ />` を 順に 取り出す★（正規表現を 使わない）
   *    返す 物 ＝ `[{ 属: '<名 ＝', 中: '＝' }, ＝]`
   *    ★`<名` の 次が 字なら 別の 札★（`<border` が `<borders` に 当たらない為） */
  function 札たち(xml, 名) {
    var 出 = [], i = 0, 開 = '<' + 名;
    var s = String(xml || '');
    while (true) {
      var a = s.indexOf(開, i);
      if (a < 0) break;
      var 次 = s.charAt(a + 開.length);
      if (次 !== ' ' && 次 !== '>' && 次 !== '/' && 次 !== '\t' && 次 !== '\n' && 次 !== '\r') {
        i = a + 開.length; continue;
      }
      var g = s.indexOf('>', a);
      if (g < 0) break;
      if (s.charAt(g - 1) === '/') { 出.push({ 属: s.slice(a, g), 中: '' }); i = g + 1; continue; }
      var 閉 = '</' + 名 + '>';
      var e = s.indexOf(閉, g);
      if (e < 0) { 出.push({ 属: s.slice(a, g), 中: '' }); i = g + 1; continue; }
      出.push({ 属: s.slice(a, g), 中: s.slice(g + 1, e) });
      i = e + 閉.length;
    }
    return 出;
  }

  /** ★`名="＝"` の 中身★（無ければ null） */
  function 属(字, 名) {
    var s = String(字 || '');
    var k = ' ' + 名 + '="';
    var a = s.indexOf(k);
    if (a < 0) return null;
    var b = s.indexOf('"', a + k.length);
    if (b < 0) return null;
    return s.slice(a + k.length, b);
  }

  /** ★中に `<名` が 在るか★（`<b/>` の ような 印） */
  function 印が在る(中, 名) {
    var s = String(中 || '');
    var 開 = '<' + 名;
    var a = s.indexOf(開);
    if (a < 0) return false;
    var 次 = s.charAt(a + 開.length);
    return 次 === ' ' || 次 === '>' || 次 === '/';
  }

  /** ★`FFFF0000`（ARGB）⇒ `'#FF0000'`★
   *    ★8桁なら 頭の 2桁は 透け具合★＝色では ありません。
   *    ★6桁なら そのまま★。★それ以外は null★（★当て推量で 色を 作らない★） */
  function 色にする(rgb) {
    var s = String(rgb || '').trim();
    if (s.length === 8) s = s.slice(2);
    if (s.length !== 6) return null;
    for (var i = 0; i < 6; i++) {
      var c = s.charAt(i).toUpperCase();
      if (!((c >= '0' && c <= '9') || (c >= 'A' && c <= 'F'))) return null;
    }
    return '#' + s.toUpperCase();
  }

  /* ══ ★★テーマの 色★★ ══（2026-09-21）
       ★★実Excel が 書く 形★★（実測）
         `.xlsx` ... `<color theme="4"/>` ＝ ★番号だけ★
         `.xlsb` ... `07 04 00 00 15 60 82 ff` ＝ ★番号 ＋ 実際の R G B★
         ⇒`.xlsx` は `xl/theme/theme1.xml` を ★引かないと 色が 出ません★
       ★★番号 ⇒ 紙の 何番目か★★（経営者1 が 実Excel に 聞いて 逆に 引いた 物）
         紙 `docs/measured/xlsb-styles-no-ji.md`（⑩）
         | 番号 | 色 | 名 | ★紙の 何番目★ |
         |  0 | #ffffff | lt1 | ★1番目★ |
         |  1 | #000000 | dk1 | ★0番目★ |
         |  2 | #e8e8e8 | lt2 | ★3番目★ |
         |  3 | #0e2841 | dk2 | ★2番目★ |
         | 4-9 | accent1-6 | そのまま |
         ⇒★★0と1、2と3が 入れ替わって います★★
         ⇒★紙を 上から 数えて n 番目を 取ると 白黒が 逆に なります★
       ★★dk1 / lt1 は `srgbClr` では ありません★★
         `<a:sysClr val="windowText" lastClr="000000"/>` ⇒ ★`lastClr` を 見ます★
       ★★測って いない 事★★
         ・★★濃さ（tint）★★ ... `<color theme="4" tint="-0.5"/>` の 形は ★未測定★
           ⇒★元の 色だけ 出します★＝★濃さは 効きません★（色味は 合う／明るさが ずれる）
           ⇒★消すより 出す★に しました（`lib/objects.js` の 罫線と 同じ 決め）
           ⇒★★これは 「合って いる」では ありません★★＝測って もらう 所です
         ・テーマを ★別の テーマに 変えた 時★ ／ `hlink`（10）`folHlink`（11）
       ★読めなければ 何も 返しません★＝★当て推量で 黒を 入れない★ */
  var テーマの並び = ['lt1', 'dk1', 'lt2', 'dk2',
    'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];

  /* ══ ★★濃さ（tint）★★ ══（2026-09-21 経営者1 が 実Excel に 聞きました）
       紙 `docs/measured/golden-jitsu-excel-iro-no-koide-2026-09-21.tsv`
       ★★明るさ（HSL の L）を 動かします★★（色あい H と 彩り S は そのまま）
         t > 0 ... L2 = L * (1 - t) + t      （白へ 寄る）
         t < 0 ... L2 = L * (1 + t)          （黒へ 寄る）
       ★★★ぴったり 合う 式は 見つかって いません★★★
         実測 9個 中 ★8個は 差 1以内★ ／ ★t=0.75 だけ 差 2★（255段階の うち）
         明るさを 240刻み・255刻みに すると ★かえって 悪く なりました★（6/9・5/9）
         ⇒★★「合って います」とは 書きません★★＝★差 2までは 出ます★
         ⇒目には 見えない 差ですが ★嘘に しない★ 為に 書いて おきます
       ★★判子の `shade` とは 別の 計算です★★
         COM は 判子の 線の `TintAndShade` を -0.85 と 返しますが、
         ★この 式では 出ません★。★同じ 名前でも 別の 物★です。
         （判子の 側は `lib/xlsx-zukei.js` に 書いて あります）
       ★★`tint` は 半端な 実数で 書かれます★★
         `<fgColor theme="4" tint="-0.499984740745262"/>`
         ⇒★字で 比べないで ください★（-0.5 とは 書かれません）
       ★★測って いない 事★★ ... ★accent1 以外の 色で 同じ 式に なるか★（1色でしか 見て いません） */
  /** ★★色に 濃さを 掛ける★★＝★`lib/theme.js` の `濃淡()` を 呼ぶだけ★
   *    ★★自分で 書き直しません★★（2026-09-21 ここで 1回 踏みました）
   *      私は HSL（0〜1 の 実数）で 書き直し、★9個 中 2個が 差 2★ でした。
   *      経営者1 も 同じ 日に 同じ 書き直しを して 同じ 所で 外しました。
   *      ⇒`lib/theme.js` は 2026-08-30 に ★実測 15/15 で 合わせて あります★
   *      ⇒呼んだら ★9個 とも 差 0★（2026-09-21 実測）
   *    ★★何が 違ったか★★
   *      私 ... HSL（0〜1 の 実数）で 四捨五入
   *      正 ... ★Windows の HLS（0〜240 の 整数）で 切り捨て★
   *      ＝`lib/theme.js` の 頭に ★式まで 書いて あります★
   *    ⇒記憶「★作る前に 探せ★／★在るのに 呼んで いないを 探せ★」
   *    ★台が 無い 時は 色を そのまま 返します★（★当て推量で 作らない★） */
  function 濃さを掛ける(色, t) {
    if (!色 || !isFinite(t) || t === 0) return 色;
    var T = (typeof module === 'object' && module.exports)
      ? require('./theme.js')
      : (typeof self !== 'undefined' ? self.Theme : null);
    if (!T || typeof T.濃淡 !== 'function') return 色;
    var 出 = T.濃淡(色, t);
    return 出 ? String(出).toUpperCase() : 色;
  }
  /** ★`xl/theme/theme1.xml` ⇒ `{ 番号: '#RRGGBB' }`★（★番号は 上の 表の 番号★） */
  function テーマを読む(themeXml) {
    var 出 = {};
    var 入 = 切り出す(String(themeXml || ''), 'a:clrScheme');
    if (!入) return 出;
    for (var i = 0; i < テーマの並び.length; i++) {
      var t = 札たち(入, 'a:' + テーマの並び[i]);
      if (!t.length) continue;
      var c = 札たち(t[0].中, 'a:srgbClr');
      var v = c.length ? 属(c[0].属, 'val') : null;
      if (!v) {
        var sy = 札たち(t[0].中, 'a:sysClr');
        if (sy.length) v = 属(sy[0].属, 'lastClr');
      }
      var 色 = 色にする(v);
      if (色) 出[i] = 色;
    }
    return 出;
  }

  /** ★`<color ＝/>` から 色を 取る★（★取れなければ null＝付けません★） */
  function 色を取る(中, テーマ, 札の名) {
    var c = 札たち(中, 札の名 || 'color');
    if (!c.length) return null;
    var 直 = 色にする(属(c[0].属, 'rgb'));
    if (直) return 直;                              /* ★字で 書いて あれば そちら★ */
    if (!テーマ) return null;
    var n = parseInt(属(c[0].属, 'theme'), 10);
    if (!isFinite(n)) return null;
    var 元 = テーマ[n];
    if (!元) return null;                           /* ★無ければ 付けません★ */
    /* ★濃さ（tint）が 在れば 掛けます★＝上の 注（★差 2までは 出ます★） */
    var t = parseFloat(属(c[0].属, 'tint'));
    return isFinite(t) ? 濃さを掛ける(元, t) : 元;
  }

  /* ★★罫線の 太さ★★（`book.html:drawBorder` は 1＝細い ／ 2＝太い の 2段だけ）
       ★実Excel の 名前は 14通り 在ります★。★太い 側を 名指しで 並べます★
       ＝★名簿に 無い 物は 細い★に します（★見えなく するより 細く 出す★）
       ★測って いない 事★ ＝ ★この 振り分けが 実Excel の 見た目と 合うかは 未測定★ */
  var 太い側 = {
    medium: 1, thick: 1, double: 1,
    mediumDashed: 1, mediumDashDot: 1, mediumDashDotDot: 1,
  };
  function 太さ(style) {
    var s = String(style || '');
    if (!s || s === 'none') return 0;
    return 太い側[s] ? 2 : 1;
  }

  /** ★`xl/styles.xml` を 読んで 表に する★ */
  function 型を読む(stylesXml, テーマ) {
    var s = String(stylesXml || '');

    var 字体 = 札たち(切り出す(s, 'fonts'), 'font').map(function (f) {
      var o = {};
      if (印が在る(f.中, 'b')) o.bold = true;
      if (印が在る(f.中, 'i')) o.italic = true;
      if (印が在る(f.中, 'u')) o.underline = true;
      var c = 色を取る(f.中, テーマ);
      if (c) o.color = c;
      return o;
    });

    var 塗り = 札たち(切り出す(s, 'fills'), 'fill').map(function (f) {
      var p = 札たち(f.中, 'patternFill');
      if (!p.length) return {};
      if (属(p[0].属, 'patternType') !== 'solid') return {};
      var fg = 札たち(p[0].中, 'fgColor');
      if (!fg.length) return {};
      var c = 色を取る(p[0].中, テーマ, 'fgColor');
      return c ? { bgColor: c } : {};
    });

    var 罫 = 札たち(切り出す(s, 'borders'), 'border').map(function (b) {
      var o = {};
      ['top', 'bottom', 'left', 'right'].forEach(function (向) {
        var t = 札たち(b.中, 向);
        if (!t.length) return;
        var w = 太さ(属(t[0].属, 'style'));
        if (w) o[向] = w;
      });
      return o;
    });

    /* ★`cellXfs` です★（`cellStyleXfs` では ありません＝名前付きの 型の 方） */
    var 組 = 札たち(切り出す(s, 'cellXfs'), 'xf').map(function (x) {
      return {
        字体: parseInt(属(x.属, 'fontId'), 10) || 0,
        塗り: parseInt(属(x.属, 'fillId'), 10) || 0,
        罫: parseInt(属(x.属, 'borderId'), 10) || 0,
      };
    });

    return { 字体: 字体, 塗り: 塗り, 罫: 罫, 組: 組 };
  }

  /** ★`<名 ＝>` から `</名>` までを 中身ごと 取る★（無ければ 空） */
  function 切り出す(xml, 名) {
    var t = 札たち(xml, 名);
    return t.length ? t[0].中 : '';
  }

  /** ★シートの xml から ★マス ⇒ cellXfs の 番号★ を 作る★
   *    `<c r="A1" s="7">` の `r` と `s` だけ 見ます（★中身は 読みません★） */
  function マスの型番(sheetXml) {
    var 出 = {};
    var s = String(sheetXml || '');
    var i = 0;
    while (true) {
      var a = s.indexOf('<c ', i);
      if (a < 0) break;
      var g = s.indexOf('>', a);
      if (g < 0) break;
      var 札 = s.slice(a, g);
      var r = 属(札, 'r');
      var n = 属(札, 's');
      if (r && n !== null) {
        var v = parseInt(n, 10);
        if (isFinite(v)) 出[r] = v;
      }
      i = g + 1;
    }
    return 出;
  }

  /** ★★口★★ ＝ シートの xml と styles の xml から 飾りを 作る
   *    返す 物 ＝ `{ 'A1': { bold, color, bgColor, border }, ＝ }`
   *    ★飾りが 1つも 無い マスは 入れません★（★空の 物を 増やさない★） */
  function 飾りを読む(sheetXml, stylesXml, themeXml) {
    var 型 = 型を読む(stylesXml, テーマを読む(themeXml));
    var 番 = マスの型番(sheetXml);
    var 出 = {};
    Object.keys(番).forEach(function (ま) {
      var x = 型.組[番[ま]];
      if (!x) return;
      var o = {};
      /* ★★既定の 字体（0番）の 色は 付けません★★（2026-09-21）
           ★★何が 起きたか★★
             テーマの 色を 読む ように したら、
             `<font><color theme="1"/>` ＝ ★既定の 字の 色（黒）★ まで 拾い、
             ★`s=` の 在る マスだけ 黒 ／ 無い マスは Exally の 既定★ に なりました。
             ⇒★★絵で 見たら 黒と 灰が 混ざって いました★★（実Excel は 全部 黒）
             ⇒★数だけ 見て いたら 気づきません★（飾りの 在る マスが 9 ⇒ 12 に 増えただけ）
           ★★直し★★ ＝ 0番の 字体は ★Exally が 既に 描いて いる 物★なので 付けません。
           ★★まだ 合って いない 事★★
             Exally の 既定の 字の 色は `#1A2B22`／実Excel は `#000000`。
             ⇒★揃って は いますが 実Excel と 同じ 色では ありません★
             ⇒★画面の 既定の 色を 変えるかは 別の 1押し★（★全部の 本に 効きます★） */
      var f = (x.字体 === 0) ? null : 型.字体[x.字体];
      if (f) {
        if (f.bold) o.bold = true;
        if (f.italic) o.italic = true;
        if (f.underline) o.underline = true;
        if (f.color) o.color = f.color;
      }
      var g = 型.塗り[x.塗り];
      if (g && g.bgColor) o.bgColor = g.bgColor;
      var b = 型.罫[x.罫];
      if (b && Object.keys(b).length) o.border = b;
      if (Object.keys(o).length) 出[ま] = o;
    });
    return 出;
  }

  return { 飾りを読む: 飾りを読む, 型を読む: 型を読む, マスの型番: マスの型番,
    色にする: 色にする, 太さ: 太さ, テーマを読む: テーマを読む, テーマの並び: テーマの並び,
    濃さを掛ける: 濃さを掛ける };
}));
