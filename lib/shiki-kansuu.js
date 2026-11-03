/* shiki-kansuu.js — ★四角（A1:A3）と 関数を 計算する★（2026-09-13）
 *
 *  ★★土台を 自分で 作る ④枚目の 後半★★
 *    土台は 5つ … ①字に 切る ②形に する ③頼りの 地図 ④★順番と 計算★ ⑤溢れ
 *    ④の 前半（＋−×÷＾＆ 大小 ％）は `shiki-keisan.js`。
 *    ここは ★四角と 関数★ だけ です。
 *
 *  ★★借り物の 中は 1文字も 読んで いません★★
 *  ★★答えは 全部 実Excel に 打って 測りました★★
 *    （2026-09-13・★40通り★／道具 `docs/measured/toru-shikaku-kansuu.ps1`
 *      紙 `docs/measured/golden-shikaku-kansuu-2026-09-13.tsv`／Excel 16.0 build 20326）
 *
 *  ★★★一番 大きい 分かれ目＝★値が どこから 来たか★★★★
 *    ★同じ SUM なのに 中身の 扱いが 違います★
 *      `=SUM(A1:A5)`       → ★3★   … 四角の 中の 字 "2" と TRUE は ★無視★（1＋2）
 *      `=SUM(A1,"2",TRUE)` → ★4★   … 直に 書いた 字と 真偽は ★数に なる★（1＋2＋1）
 *      `=SUM(A2)`          → ★0★   … ★1マスだけ 指した 字も 無視★（直に 書いた のとは 違う）
 *    ⇒★マスから 来た 物は 飛ばす／直に 書いた 物は 数に する★
 *      ★これを 知らずに 作ると 全部 外します★
 *
 *  ★★当て推量なら 外して いた 物（実Excel に 聞いて 初めて 分かった）★★
 *    `=PRODUCT(B1:B3)`  → ★0★（全部 空。★1 では ない★＝掛け算の 元は 1 なのに）
 *    `=COUNTA(A1:A6)`   → ★5★（★誤りも 数える★）
 *    `=COUNT(A1:A6)`    → ★2★（★誤りは 数えない★＝COUNTA と 違う）
 *    `=MAX(B1:B3)`      → ★0★（全部 空）／`=MAX(C5:C6)` → ★-1★（負だけ＝★0は 混ざらない★）
 *    `=MAX(A1:A6)`      → ★#DIV/0!★（★MAXにも 誤りは 伝わる★）
 *    `=AVERAGE(B1:B3)`  → ★#DIV/0!★（★SUMは 0 なのに 平均は 誤り★）
 *    `=AVERAGE(A1,A3)`  → ★1★（1マス指しの TRUE は ★分母に 入らない★）
 *    `=SUM(0.1,0.2,-0.3)`→★0★（★SUMの 中でも 消え残りを 0に する★＝＋と 同じ／窓２で 裏取り）
 *
 *  ★値の 形★ `shiki-keisan.js` と 同じ { 型:'数'|'字'|'真偽'|'誤'|'空', 値 }
 *
 *  ★引数の 形★（★どこから 来たかを 必ず 付ける★）
 *    { 種:'直',   値: <値> }          … 式に 直に 書いた 物（"2" や TRUE や 1）
 *    { 種:'マス', 値: <値> }          … 1マスを 指した 物（A2）
 *    { 種:'四角', 並び: [<値>, …] }   … A1:A3
 *
 *  ★★入れて いない 物（出来て いない 物を 出来た 顔で 混ぜない）★★
 *    ・SUMPRODUCT … ★形の 違う 四角の 決まりが 未測定★
 *    ・溢れ（`=A1:A3` を そのまま 打つと 下に 広がる）… ★土台⑤★
 *      ★実測では 溢れました★（E1=1・E2=2）が、ここでは 扱いません。
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(
    typeof require === 'function' ? require('./shiki-keisan.js') : root.ShikiKeisan,
    typeof require === 'function' ? require('./shiki-basho.js') : root.ShikiBasho,
    typeof require === 'function' ? require('./shoshiki.js') : root.Shoshiki);
  else root.ShikiKansuu = factory(root.ShikiKeisan, root.ShikiBasho, root.Shoshiki);
})(typeof self !== 'undefined' ? self : this, function (計, 場, 書) {
  'use strict';

  var 数 = 計.数, 誤 = 計.誤;

  /* ══ ★引数を 1つずつ ほどく★ ══
     ★四角は 中身を 順に／マスと 直は そのまま★
     ★どこから 来たかを 落とさない★のが この 台の 肝 */
  function ほどく(引数たち, 続き) {
    for (var i = 0; i < 引数たち.length; i++) {
      var a = 引数たち[i];
      if (a.種 === '四角') {
        for (var j = 0; j < a.並び.length; j++) {
          if (続き(a.並び[j], 'マス') === false) return;
        }
      } else {
        if (続き(a.値, a.種 === '直' ? '直' : 'マス') === false) return;
      }
    }
  }

  /* ══ ★数を 拾う★（SUM／AVERAGE／MAX／MIN／PRODUCT／COUNT が 使う）══
     ★マスから 来た … 数だけ 拾う。字・真偽・空は 飛ばす★
     ★直に 書いた … 数に する（字が 数に ならなければ #VALUE!）★
     ★誤りは どちらでも 伝わる（左から 最初の 1つ）★ */
  function 数を拾う(引数たち, 手) {
    var 出 = { 数たち: [], 誤: null };
    ほどく(引数たち, function (v, 元) {
      if (v.型 === '誤') { 出.誤 = v; return false; }
      if (元 === 'マス') {
        if (v.型 === '数') 出.数たち.push(v.値);
        return true;                       /* ★字・真偽・空は 飛ばす★ */
      }
      var n = 計.数にする(v, 手);           /* ★直に 書いた 物は 数に する★ */
      if (n.型 === '誤') { 出.誤 = n; return false; }
      出.数たち.push(n.値);
      return true;
    });
    return 出;
  }

  /** ★★A の 付く 一族の 拾い方★★（AVERAGEA MAXA MINA VARA VARPA STDEVA STDEVPA）
        ★A の 無い 方との 違い＝★マスの 中の 字と 真偽を 飛ばさない★★
          ・マスの 真偽 … TRUE＝1 ／ FALSE＝0
          ・マスの 字   … ★0 として 数える★（★数に 読める 字でも 0★）
          ・マスの 空   … ★飛ばす★（A の 無い 方と 同じ）
        ★直に 書いた 物は A の 有無に かかわらず 数に する★（`=AVERAGEA("101",2)` → 51.5）
        ★物差し★ `golden-346-2026-09-08.tsv`
          `=AVERAGEA(A1:B5,2,FALSE)` → ★3.9166666666666665★（＝45＋2＋0 を 12個で 割る） */
  function Aで数を拾う(引数たち, 手) {
    var 出 = { 数たち: [], 誤: null };
    ほどく(引数たち, function (v, 元) {
      if (v.型 === '誤') { 出.誤 = v; return false; }
      if (元 === 'マス') {
        if (v.型 === '数') 出.数たち.push(v.値);
        else if (v.型 === '真偽') 出.数たち.push(v.値 ? 1 : 0);
        else if (v.型 === '字') 出.数たち.push(0);      /* ★数に 読める 字でも 0★ */
        return true;                                    /* ★空は 飛ばす★ */
      }
      var n = 計.数にする(v, 手);
      if (n.型 === '誤') { 出.誤 = n; return false; }
      出.数たち.push(n.値);
      return true;
    });
    return 出;
  }

  /* ══ ★足す★ ══
     ★消え残りを 0に する所は `shiki-keisan.js` の ＋ に 任せます★
     ＝★同じ 決まりを 2か所に 書かない★（実測 `=SUM(0.1,0.2,-0.3)` → 0） */
  function 足しあげる(数たち, 手) {
    var 合 = 数(0);
    for (var i = 0; i < 数たち.length; i++) {
      合 = 計.つなぎ('+', 合, 数(数たち[i]), 手);
      if (合.型 === '誤') return 合;
    }
    return 合;
  }

  function 掛けあげる(数たち, 手) {
    if (数たち.length === 0) return 数(0);   /* ★実測＝全部空の PRODUCT は 0（1 では ない）★ */
    var 積 = 数(数たち[0]);
    for (var i = 1; i < 数たち.length; i++) {
      積 = 計.つなぎ('*', 積, 数(数たち[i]), 手);
      if (積.型 === '誤') return 積;
    }
    return 積;
  }

  /* ══ ★MATCH と INDEX★（2026-09-15 に 書いた）══
     ★★なぜ この 2つが 先か★★
       司さんの 実物 1冊（代行計算表2026.xlsb）を 数えたら
       ★出てくる 関数は 8個だけ★／その うち ★INDEX と MATCH で のべの 68%★。
       ⇒★借り物を 外す 道の 一番 太い 所★。
     ★★答えは 全部 実Excel に 打って 測りました★★
       `docs/measured/kansuu46/golden-index-match2-2026-09-15.tsv`       … 138本（誤り 100本）
       `docs/measured/kansuu46/golden-index-match-kimari-2026-09-15.tsv` … 55本（★字の 表★）
     ★★当て推量なら 外して いた 物★★
       ・MATCH の 型は ★符号だけ★（2 も 1.9 も 0.5 も「1」／-2 も -0.5 も「-1」）
       ・型 1 は ★なめるのでは なく 二分探索★（降順 C1:C5 に 型1 で ★3★。なめれば 5）
       ・字は ★大文字小文字を 区別しない★・★ワイルドカードが 効く★
       ・INDEX は ★値では なく 参照★（`=ISBLANK(INDEX(E1:E5,4))` → TRUE）
       ・INDEX の 転け方は ★2種類★（行/列が 外＝#REF! ／ 負・字＝#VALUE!）
       ・★字の 誤りは 場所の 誤りより 先★（`=INDEX(D1,D2,"D")` → #REF! では なく #VALUE!） */

  /* ★数の 引数を 1つの 数に する★（暗黙の 交わり → 数に する）
     ★切り捨てません★＝MATCH の 型は ★符号だけ★ 使うので、ここでは 生の 数の まま */
  function 数の引数(引数, 手, 所) {
    var v = 場.ひとつに(引数, 所 && 所.今のマス);
    if (!v) return { 誤: 誤('#VALUE!') };
    if (v.型 === '誤') return { 誤: v };
    var n = 計.数にする(v, 手);
    if (n.型 === '誤') return { 誤: 誤('#VALUE!') };
    return { 数: n.値 };
  }
  /* ★INDEX の 行番号・列番号は 切り捨て★（実測 `=INDEX(A1:A5,1.9)` → ★1★） */
  function 切り捨て(x) { return x < 0 ? Math.ceil(x) : Math.floor(x); }

  /* ★当たるか★＝MATCH の「ぴたり」
     ★字どうしで 型紙（* ?）が 入って いれば 型紙として 見る★
     ★それ以外は 大小くらべの 0★（＝型が 違えば 当たらない
       実測 `=MATCH("3",A1:A5,0)` → #N/A ／ `=MATCH(3,"3",0)` → #N/A） */
  function 当たるか(探, v) {
    if (探 && 探.型 === '字' && v && v.型 === '字' && 場.型紙か(String(探.値))) {
      return 場.型紙に合うか(String(探.値), String(v.値));
    }
    return 場.比べる(探, v) === 0;
  }

  /* ══ ★★全角 と 半角（ASC ／ DBCS が 使います）★★ ══（2026-09-18）
     ★★字の 表を ★記憶で 書いて いません★★★
       ＝`String.normalize` が 持つ ★Unicode の 表★を 引きます
       ＝★私の 記憶では なく 機械の 表★
     ★半角カナ（U+FF61〜U+FF9F・62字）の 対応を ★その場で 作ります★★ */
  var 全半の表 = null;
  function 全半の表を作る() {
    if (全半の表) return 全半の表;
    var 全から半 = {};
    for (var c = 0xFF61; c <= 0xFF9F; c++) {
      var 半 = String.fromCodePoint(c);
      var 全 = 半.normalize('NFKC');
      if (!Object.prototype.hasOwnProperty.call(全から半, 全)) 全から半[全] = 半;
    }
    全半の表 = 全から半;
    return 全半の表;
  }
  /* ★ASC★ … 全角 → 半角 */
  function 全角を半角に(s) {
    var 表 = 全半の表を作る();
    /* ★★実測の 裏返し（★ここだけ 自体は 測って いません★）★★
         実測 … `=DBCS(U+FF9E)` → ★U+309B★
         ⇒その 逆を 先に 置きます（U+309B → U+FF9E）
         訳 … `NFKD` は U+309B を ★空白 ＋ 結合濁点★ に ばらします
              ⇒★往復で 戻らなく なります★
         ★★「記憶で 埋めるより 実測済みの 隣に 合わせる」（2026-09-18 の 学び）★★
         ⇒★次の 枠で 聞きます★（`asc-dbcs-kiku-koto.md`）*/
    /* ══ ★★半角に ならない／別の 字に なる 物★★ ══（2026-09-18・8枠目で 実測）
         `=UNICODE(ASC(UNICHAR(12532)))` -> ★12532★（ヴ）＝★変わりません★
           ＝`NFKD` は ヴ を ★ウ ＋ 結合濁点★ に ばらすので ★そのままでは 65395★
         `=UNICODE(ASC(UNICHAR(65509)))` -> ★92★（全角円 -> ★逆斜線★）
           ＝★円記号では なく 逆斜線★（CP932 の 決め） */
    var そのまま = {};
    そのまま[String.fromCodePoint(0x30F4)] = String.fromCodePoint(0x30F4);   /* ヴ */
    そのまま[String.fromCodePoint(0xFFE5)] = String.fromCodePoint(0x005C);   /* 全角円 -> 逆斜線 */
    var 先 = String(s).replace(/゛/g, String.fromCodePoint(0xFF9E))
                      .replace(/゜/g, String.fromCodePoint(0xFF9F));
    /* ★ヴ は `NFKD` の 前に 逃がします★（★ばらされると 戻せません★） */
    var ヴの印 = String.fromCodePoint(0xE000);   /* ★私用領域＝式には 出て きません★ */
    /* ★★`NFKD` の 前に 逃がす 物★★
         ・ヴ（U+30F4） ... `NFKD` が ★ウ ＋ 結合濁点★ に ばらす
         ・全角円（U+FFE5） ... `NFKD` が ★U+00A5（165）★に する
           ⇒★実Excel は 92（逆斜線）★＝★別の 字★
         ⇒★どちらも 先に 印に 置き換えて、後で 戻します★ */
    var 円の印 = String.fromCodePoint(0xE001);
    先 = 先.split(String.fromCodePoint(0x30F4)).join(ヴの印);
    先 = 先.split(String.fromCodePoint(0xFFE5)).join(円の印);
    var k = 先.normalize('NFKD')
      .split(ヴの印).join(String.fromCodePoint(0x30F4))
      .split(円の印).join(String.fromCodePoint(0xFFE5));
    var 出 = '';
    for (var i = 0; i < k.length; i++) {
      var ch = k[i];
      var p = ch.codePointAt(0);
      /* ★サロゲートは そのまま★（★未測定★） */
      if (p >= 0xD800 && p <= 0xDBFF && i + 1 < k.length) { 出 += k[i] + k[i + 1]; i++; continue; }
      if (Object.prototype.hasOwnProperty.call(そのまま, ch)) { 出 += そのまま[ch]; continue; }
      if (Object.prototype.hasOwnProperty.call(表, ch)) { 出 += 表[ch]; continue; }
      if (p === 0x3099) { 出 += String.fromCodePoint(0xFF9E); continue; }   /* 結合濁点 */
      if (p === 0x309A) { 出 += String.fromCodePoint(0xFF9F); continue; }   /* 結合半濁点 */
      /* ★★全角 ASCII（U+FF01〜U+FF5E）と 全角空白（U+3000）の 枝は ★消しました★★
           ★訳★ … `NFKD` が ★先に 済ませて いました★
                  NFKD(U+3000) → 32 ／ NFKD(U+FF01) → 33 ／ NFKD(U+FF21) → 65
           ★見つけ方★ … ★わざと 壊したのに 赤に ならなかった★
                    ⇒★見張りが 甲い のでは なく ★死んだ 枝★ だった★
           ⇒★★死んだ コードを 残すと 「守って いる」と 見えます★★ */
      出 += ch;
    }
    return 出;
  }
  /* ══ ★★バイトで 数える（LENB ／ LEFTB ／ RIGHTB ／ MIDB が 使います）★★ ══（2026-09-18）
     ★1バイト★ … ASCII（〜U+007F）と ★半角カナ（U+FF61〜U+FF9F）★
     ★2バイト★ … それ以外
     ★★測ったのは 3つだけ★★（半角A ／ 全角ア ／ 半角ｱ）
       ⇒★漢字・ひらがな・絵文字は 測って いません★（`lenb-kiku-koto.md`）
     ★★2バイトの 字を 半分で 切ると ★空白 1文字★に なります★★（★実測★）
       `=UNICODE(LEFTB(全角ア,1))` → ★32★ */
  /* ══ ★★CP932 に 無い 字は 1バイト★★ ══（2026-09-18・8枠目で 実測）
       `=LENB(UNICHAR(8212))` -> ★1★（U+2014 EM DASH）
       `=LENB(UNICHAR(9731))` -> ★1★（雪だるま）
       ★訳★ ... 実Excel の LENB は ★このパソコンの ANSI（CP932）で 数えます★
              ＝★CP932 に 無い 字は 1バイトの 代わりの 字（?）に なります★
       ★★私は 一覧を 持って いません★★
         ＝★測った 2つだけ 名指しで 持ちます★（★一覧を 作った ふりを しません★）
         ＝★他の 字は まだ 2バイトの ままです★＝★次の 枠で 聞きます★
       ★紙の 頭に 文字コードを 書いて あります★（ANSI 932）
         ＝★別の パソコンでは 変わり得ます★ */
  var CP932に無い = {};
  CP932に無い[String.fromCodePoint(0x2014)] = true;   /* EM DASH（U+2015 は 在る） */
  CP932に無い[String.fromCodePoint(0x2603)] = true;   /* 雪だるま */

  function 一字のバイト(ch) {
    var p = ch.codePointAt(0);
    if (p <= 0x7F) return 1;
    if (p >= 0xFF61 && p <= 0xFF9F) return 1;
    if (Object.prototype.hasOwnProperty.call(CP932に無い, ch)) return 1;
    return 2;
  }
  function バイト数(s) {
    var t = String(s), n = 0;
    for (var i = 0; i < t.length; i++) {
      var p = t.codePointAt(i);
      if (p > 0xFFFF) { n += 2; i++; continue; }     /* ★サロゲート＝未測定★ */
      n += 一字のバイト(t[i]);
    }
    return n;
  }
  /* ★[始, 始＋長) バイトを 取る★（1から 数える）／★半端な 半分は 空白★ */
  function バイトで切る(s, 始, 長) {
    var t = String(s), 出 = '', 位 = 1;
    if (長 <= 0) return '';
    var 終 = 始 + 長;                                   /* ★終わりは 含まない★ */
    for (var i = 0; i < t.length; i++) {
      var ch = t[i];
      var p = t.codePointAt(i);
      if (p > 0xFFFF) { ch = t[i] + t[i + 1]; i++; }
      var b = (p > 0xFFFF) ? 2 : 一字のバイト(ch);
      var 頭 = 位, 尻 = 位 + b;                          /* この 字は [頭, 尻) */
      位 = 尻;
      if (尻 <= 始 || 頭 >= 終) continue;                 /* まるごと 外 */
      if (頭 >= 始 && 尻 <= 終) { 出 += ch; continue; }    /* まるごと 中 */
      /* ★半分だけ 入った★ ⇒ ★空白 1文字★（実測） */
      出 += ' ';
    }
    return 出;
  }
  function バイトで端を取る(引数たち, 手, 所, 左か) {
    var t = 字の引数(引数たち[0], 手, 所);
    if (!引数たち.length || t.誤) return t.誤 || 誤('#VALUE!');
    var n = 1;
    if (引数たち.length > 1) {
      var m = 数の引数(引数たち[1], 手, 所);
      if (m.誤) return m.誤;
      n = 切り捨て(m.数);
    }
    if (n < 0) return 誤('#VALUE!');
    var 全 = バイト数(t.字);
    if (n >= 全) return 計.字(t.字);
    return 計.字(左か ? バイトで切る(t.字, 1, n) : バイトで切る(t.字, 全 - n + 1, n));
  }

  /* ══ ★★区切りの 前／後ろ（TEXTAFTER ／ TEXTBEFORE）★★ ══（2026-09-18）
     ★★一番 間違えやすい 所★★
       ★既定（4つ目を 省く）は ★大小を 見ます★★（`=TEXTAFTER("a-B-c","b")` → ★#N/A★）
       ★`1` を 渡すと ★大小を 見ません★★（`…,1,1)` → ★"-c"★）
       ⇒★★私たちの 見込みは ★逆★でした★★（2026-09-18 実測）
       ⇒★SEARCH（大小を 見ない）と FIND（見る）の 感じとは ★別★★
     ★番目★ … 既定 1／★負なら 後ろから★／行き過ぎたら ★#N/A★
     ★区切りが 空の 字★ … `=TEXTBEFORE("abc","")` → ★空の 字★（★#VALUE! では ない★） */
  function 区切りの前後(引数たち, 手, 所, 前か) {
    if (引数たち.length < 2) return 誤('#VALUE!');
    var t = 字の引数(引数たち[0], 手, 所);
    if (t.誤) return t.誤;
    var d = 字の引数(引数たち[1], 手, 所);
    if (d.誤) return d.誤;
    var 番 = 1;
    if (引数たち.length > 2) {
      var m = 数の引数(引数たち[2], 手, 所);
      if (m.誤) return m.誤;
      番 = 切り捨て(m.数);
    }
    if (番 === 0) return 誤('#VALUE!');          /* ★実測★ `=TEXTAFTER("a-b","-",0)` */
    var 大小を見る = true;                        /* ★★既定は 見る（実測）★★ */
    if (引数たち.length > 3) {
      var k = 数の引数(引数たち[3], 手, 所);
      if (k.誤) return k.誤;
      大小を見る = (切り捨て(k.数) === 0);         /* ★1 なら 見ない（実測）★ */
    }
    /* ══ ★★5つ目＝終わり（始め）も 区切りと 見なす★★ ══（2026-09-18・8枠目で 実測）
         `=TEXTAFTER("abc","x",1,0,1)`  -> ★空の 字★
         `=TEXTBEFORE("abc","x",1,0,1)` -> ★abc★
         ⇒★見つからなくても ★終わりが 区切り★に なります★ */
    var 終わりも = false;
    if (引数たち.length > 4) {
      var e5 = 数の引数(引数たち[4], 手, 所);
      if (e5.誤) return e5.誤;
      終わりも = (切り捨て(e5.数) !== 0);
    }
    /* ══ ★★6つ目＝見つからない 時に 返す 物★★ ══（2026-09-18・8枠目で 実測）
         `=TEXTAFTER("abc","x",1,0,0,<字>)` -> ★その 字★ */
    var 無い時 = null;
    if (引数たち.length > 5) {
      var v6 = 場.ひとつに(引数たち[5], 所 && 所.今のマス);
      if (v6 && v6.型 === '誤') return v6;
      無い時 = v6 || 計.空;
    }
    var 本 = String(t.字), 区 = String(d.字);
    /* ══ ★★番が 本文より 長ければ #VALUE!★★ ══（2026-09-18・8枠目で 実測）
         `=TEXTAFTER("a-b","-",9,0,0,<字>)` -> ★#VALUE!★（★6つ目が 在っても★）
         `=TEXTAFTER("aaa","aa",2)`         -> ★#N/A★（★2 は 長さ 3 の 中★）
         ⇒★★「範囲の 外」と「見つからない」は 別の 誤り★★ */
    /* ★★空の 本文は ここを 通しません★★（2026-09-18・8枠目で 実測）
         `=TEXTAFTER("","x")` -> ★#N/A★（★#VALUE! では ない★）
         ⇒★長さ 0 の 時は 「範囲の 外」では なく 「見つからない」★ */
    if (本.length > 0 && Math.abs(番) > 本.length) return 誤('#VALUE!');
    /* ★区切りが 空★ … 前は 空の 字／後ろは 本文 まるごと（★後ろは 未測定★） */
    if (区 === '') return 計.字(前か ? '' : 本);
    var 探本 = 大小を見る ? 本 : 本.toLowerCase();
    var 探区 = 大小を見る ? 区 : 区.toLowerCase();
    /* ══ ★★重なりは 数えません★★ ══（2026-09-18・8枠目で 実測）
         `=TEXTAFTER("aaa","aa",2)` -> ★#N/A★
         ＝重なりを 数えると 2つ 在る 事に なり ★a★ が 返ります
         ⇒★1つ 見つけたら ★区切りの 長さだけ★ 進みます★ */
    var 所たち = [], p = 0;
    while (true) {
      var i = 探本.indexOf(探区, p);
      if (i < 0) break;
      所たち.push({ 位: i, 長: 区.length });
      p = i + 区.length;
    }
    /* ★終わりも 区切りに する時は ★最後に 足します★★（長さ 0） */
    if (終わりも) 所たち.push({ 位: 本.length, 長: 0 });
    var 見つからない = function () {
      if (無い時 !== null) return 無い時;
      return 誤('#N/A');
    };
    if (!所たち.length) return 見つからない();
    var j = (番 > 0) ? (番 - 1) : (所たち.length + 番);
    if (j < 0 || j >= 所たち.length) return 見つからない();
    var 位 = 所たち[j].位, 長 = 所たち[j].長;
    return 計.字(前か ? 本.slice(0, 位) : 本.slice(位 + 長));
  }

  /* ★DBCS★ … 半角 → 全角 */
  function 半角を全角に(s) {
    var t = String(s), 出 = '';
    for (var i = 0; i < t.length; i++) {
      var ch = t[i], p = ch.codePointAt(0);
      if (p >= 0xD800 && p <= 0xDBFF && i + 1 < t.length) { 出 += t[i] + t[i + 1]; i++; continue; }
      if (p === 0x20) { 出 += String.fromCodePoint(0x3000); continue; }
      /* ★★逆斜線は ★全角の 円記号★に なります★★（2026-09-18・8枠目で 実測）
           `=UNICODE(DBCS(UNICHAR(92)))` -> ★65509★（U+FFE5）
           ＝★U+FF3C（全角の 逆斜線）では ありません★
           ＝★ASC の 裏返し★（全角円 -> 逆斜線）と 対に なって います */
      if (p === 0x5C) { 出 += String.fromCodePoint(0xFFE5); continue; }
      if (p >= 0x21 && p <= 0x7E) { 出 += String.fromCodePoint(p + 0xFEE0); continue; }
      /* ★★実測★★ `=DBCS(半角濁点 U+FF9E)` → ★U+309B（単独の ゛）★
           ⇒`NFKC` は ★U+3099（結合用）★を 返すので ★ここだけ 別★ */
      if (p === 0xFF9E) { 出 += String.fromCodePoint(0x309B); continue; }
      if (p === 0xFF9F) { 出 += String.fromCodePoint(0x309C); continue; }
      if (p >= 0xFF61 && p <= 0xFF9F) { 出 += ch.normalize('NFKC'); continue; }
      出 += ch;
    }
    return 出;
  }

  /* ★型 1（以下で 一番 後ろ）＝★二分探索★★
     ★実測で「なめるのでは ない」と 分かりました★
       `=MATCH(3,C1:C5,1)`（C＝5,4,3,2,1）→ ★3★
        なめて「探す値 以下の 最後」を 採ると ★5★ に なる ⇒ なめて いない
     ★空・比べられない 物は 左へ★（★未測定★＝後ろが 空の 列で 正しく なる 方を 選んだ） */
  function 二分で探す(並び, 探) {
    var lo = 1, hi = 並び.length, 当 = 0;
    while (lo <= hi) {
      var mid = Math.floor((lo + hi) / 2);
      var c = 場.比べる(並び[mid - 1], 探);
      if (c === 0) return mid;
      if (c !== null && c < 0) { 当 = mid; lo = mid + 1; } else { hi = mid - 1; }
    }
    return 当;
  }

  /* ★型 -1（以上で 一番 後ろ・降順）★
     ★実測 3本に 合う 形★（★これは「実Excel の 中身」では なく 実測に 合う 模型★です）
       `=MATCH(3,C1:C5,-1)` → ★3★  （C＝5,4,3,2,1＝ちゃんと 降順）
       `=MATCH(3,A1:A5,-1)` → ★#N/A★（A＝1,2,3,4,5＝降順で ない）
       `=MATCH(1,A1:A5,-1)` → ★1★  （先頭だけ 当たる）
     ⇒★頭から 見て「探す値 以上」で かつ「降順が 崩れて いない」所まで★
       ★ちゃんと 降順に 並んで いれば これは 正しい 答えと 同じ★です。 */
  function 降順で探す(並び, 探) {
    var 当 = 0;
    for (var i = 0; i < 並び.length; i++) {
      if (i > 0 && 場.比べる(並び[i], 並び[i - 1]) > 0) break;   /* 降順が 崩れた */
      var c = 場.比べる(並び[i], 探);
      if (c === null || c < 0) break;
      当 = i + 1;
    }
    return 当;
  }

  /* ══ ★残り 4個の うち 3個★（IF／IFERROR／SUBTOTAL）══（2026-09-15 に 書いた）
     ★司さんの 実物 1冊★ … IFERROR 4,329回 ／ IF 2,884回 ／ SUBTOTAL 496回
     ★紙★ `docs/measured/kansuu46/golden-nokori4-kimari-2026-09-15.tsv`（99本）
     ★TEXT は まだ 書いて いません★＝★書式の 台が 要ります★（棚 ⑮） */

  /* ★真か 偽か★（IF の 条件）
     ★実測★ 1→真 ／ 0→偽 ／ -1→真 ／ 0.5→真 ／ 空マス→偽
             ★字の "TRUE" は 真★／★"あ" と "" は #VALUE!★（★当て推量なら 全部 真に して いた★） */
  function 真偽にする(v) {
    if (!v) return { 誤: 誤('#VALUE!') };
    if (v.型 === '誤') return { 誤: v };
    if (v.型 === '真偽') return { 真: !!v.値 };
    if (v.型 === '数') return { 真: v.値 !== 0 };
    if (v.型 === '空') return { 真: false };
    if (v.型 === '字') {
      var u = String(v.値).trim().toUpperCase();
      if (u === 'TRUE') return { 真: true };
      if (u === 'FALSE') return { 真: false };
    }
    return { 誤: 誤('#VALUE!') };
  }

  /* ★SUBTOTAL の 11個★ … 1 AVERAGE ／ 2 COUNT ／ 3 COUNTA ／ 4 MAX ／ 5 MIN ／ 6 PRODUCT
                            7 STDEV ／ 8 STDEVP ／ 9 SUM ／ 10 VAR ／ 11 VARP
     ★101〜111 は「隠した 行を 抜く」★＝★今は 隠す 仕掛けが 無いので 同じ★（実測でも 同じ） */
  function 散らばり(数たち, 母集団か) {
    var n = 数たち.length;
    if (n < (母集団か ? 1 : 2)) return 誤('#DIV/0!');   /* ★n が 足りない時は 未測定★（棚） */
    var 和 = 0, i;
    for (i = 0; i < n; i++) 和 += 数たち[i];
    var 平 = 和 / n, 二乗 = 0;
    for (i = 0; i < n; i++) 二乗 += (数たち[i] - 平) * (数たち[i] - 平);
    return 数(二乗 / (母集団か ? n : n - 1));
  }

  /* ══ ★D系（DSUM ほか）の 土台★ ══（2026-09-15に 書いた）
     ★★これは 「移すだけ」では ありませんでした★★
       今の 製品の DSUM は `exally-formula.js` の `_jsDbFunc`。
       ★計算は 自前★ですが ★読む 先が `_hf.getCellValue`（★借り物の 表★）★
       ⇒★★台へ 持って 来るには 読む 所を 全部 書き直す事に なります★★
       ⇒★だから ここは ★台の 四角（並び／行数／列数）を 直に 見る形★ で 書きました★

     ★物差し★ … `docs/measured/kansuu46/golden-mada-2026-09-14.tsv`（実Excel 16.0 build 20326）
       `=DSUM(A1:B5,1,D1:D2)` → ★14★  … ★見出しが 合わない 条件は ★効かない★★
       `=DSUM(A1:B5,1,A1:A2)` → ★2★
     ★★今の 製品は 前者を ★0★ と 答えます★★＝★客に 出る 欠陥（棚）★

     ★★見て いない 事★★
       ・★ふたつ以上の 条件の 行（OR）／列（AND）は ★未測定★★（形だけ 入れて あります）
       ・★記号（`>2` など）も ★未測定★★
       ・★選ばなかった 行に 誤りが ある 時は ★未測定★★（今は 選んだ 行だけ 見ます） */

  var 較べ記号 = /^(>=|<=|<>|>|<|=)([\s\S]*)$/;

  /** ★条件の マス 1つに 実の 値が 当てはまるか★ */
  function 当てはまるか(条, 実) {
    if (条 && 条.型 === '字') {
      var m = 較べ記号.exec(String(条.値));
      if (m) {
        var 右 = m[2].trim();
        var n = Number(右);
        var 相 = (右 !== '' && isFinite(n)) ? 計.数(n) : 計.字(右);
        var d = 場.比べる(実, 相);
        if (d === null || d === undefined) return false;
        if (m[1] === '>') return d > 0;
        if (m[1] === '<') return d < 0;
        if (m[1] === '>=') return d >= 0;
        if (m[1] === '<=') return d <= 0;
        if (m[1] === '<>') return d !== 0;
        return d === 0;
      }
    }
    return 場.比べる(実, 条) === 0;
  }

  /** ★D系の 共通★… 台・列番・条件 から ★選んだ 行の 値たち★ を 返す */
  function D系の値たち(引数たち, 手, 所) {
    if (引数たち.length < 3) return { 誤: 誤('#VALUE!') };
    var 台 = 引数たち[0], 条 = 引数たち[2];
    if (台.種 !== '四角' || 条.種 !== '四角') return { 誤: 誤('#VALUE!') };
    if (!(台.行数 > 1) || !(台.列数 > 0)) return { 誤: 誤('#VALUE!') };

    var 見 = [], c;
    for (c = 0; c < 台.列数; c++) 見.push(台.並び[c]);

    /* ★field★ … 数なら 何列目（１から）／字なら 見出しの 名前 */
    var f = 場.ひとつに(引数たち[1], 所 && 所.今のマス);
    if (!f) return { 誤: 誤('#VALUE!') };
    if (f.型 === '誤') return { 誤: f };
    var 列番 = -1;
    if (f.型 === '数') 列番 = 切り捨て(f.値) - 1;
    else for (c = 0; c < 見.length; c++) if (場.比べる(f, 見[c]) === 0) { 列番 = c; break; }
    if (列番 < 0 || 列番 >= 台.列数) return { 誤: 誤('#VALUE!') };

    /* ★条件の 見出しを 台の 見出しに 紐付ける★
       ★★紐付かなければ その 列は 縛らない★★（実測 … `D1:D2` で 14） */
    var 条列 = [];
    for (c = 0; c < 条.列数; c++) {
      var 先 = -1;
      for (var k = 0; k < 見.length; k++) if (場.比べる(条.並び[c], 見[k]) === 0) { 先 = k; break; }
      条列.push(先);
    }

    var 値たち = [];
    for (var r = 1; r < 台.行数; r++) {
      var 合うか = !(条.行数 > 1);          /* ★条件の 行が 無ければ 全部★ */
      for (var cr = 1; cr < 条.行数 && !合うか; cr++) {
        var 全部 = true;
        for (var c3 = 0; c3 < 条.列数 && 全部; c3++) {
          if (条列[c3] < 0) continue;                       /* ★縛らない★ */
          var 条値 = 条.並び[cr * 条.列数 + c3];
          if (!条値 || 条値.型 === '空') continue;          /* ★空の 条件は 縛らない★ */
          if (!当てはまるか(条値, 台.並び[r * 台.列数 + 条列[c3]])) 全部 = false;
        }
        if (全部) 合うか = true;
      }
      if (合うか) 値たち.push(台.並び[r * 台.列数 + 列番]);
    }
    return { 値たち: 値たち };
  }

  /** ★D系の 共通★… 選んだ 行から ★数だけ★ 拾って 続きへ 渡す
        ★誤りは 素通り★（どの 関数でも 同じ 決め）／★字・真偽・空は 飛ばす★ */
  function D系の数で(引数たち, 手, 所, 続き) {
    var 拾 = D系の値たち(引数たち, 手, 所);
    if (拾.誤) return 拾.誤;
    var 数たち = [];
    for (var i = 0; i < 拾.値たち.length; i++) {
      var v = 拾.値たち[i];
      if (!v) continue;
      if (v.型 === '誤') return v;
      if (v.型 === '数') 数たち.push(v.値);
    }
    return 続き(数たち);
  }

  /** ★誤差関数★（★Abramowitz-Stegun 7.1.26 では 桁が 足りません★⇒級数と 連分数） */
  function 誤差関数(x) {
    if (x === 0) return 0;
    var 符 = x < 0 ? -1 : 1, a = Math.abs(x);
    if (a < 3) {
      /* ★級数★ erf(x) = 2/√π × Σ (-1)^n x^(2n+1) / (n!(2n+1)) */
      var 和 = 0, 項 = a;
      for (var n = 0; n < 200; n++) {
        和 += 項 / (2 * n + 1);
        項 *= -a * a / (n + 1);
        if (Math.abs(項) < 1e-300) break;
      }
      return 符 * 2 / Math.sqrt(Math.PI) * 和;
    }
    /* ★大きい 所は 連分数★ */
    var t = 0;
    for (var k = 60; k >= 1; k--) t = k / 2 / (a + t);
    var erfc = Math.exp(-a * a) / Math.sqrt(Math.PI) / (a + t);
    return 符 * (1 - erfc);
  }
  /** ★十進 → 他の 進数★（★負は 10桁の 補数★） */
  function 十進から(引数たち, 手, 所, 基) {
    if (!引数たち.length) return 誤('#VALUE!');
    var t = 数の引数(引数たち[0], 手, 所);
    if (t.誤) return t.誤;
    var n = 切り捨て(t.数);
    var 幅 = { 2: 512, 8: 536870912, 16: 549755813888 }[基];
    if (n < -幅 || n >= 幅) return 誤('#NUM!');
    var 字;
    if (n < 0) {
      var 桁 = { 2: 10, 8: 10, 16: 10 }[基];
      var 周 = Math.pow(基, 桁);
      字 = (周 + n).toString(基).toUpperCase();
    } else {
      字 = n.toString(基).toUpperCase();
    }
    if (引数たち.length > 1) {
      var d = 数の引数(引数たち[1], 手, 所);
      if (d.誤) return d.誤;
      var 桁数 = 切り捨て(d.数);
      if (桁数 < 0) return 誤('#NUM!');
      if (n >= 0) {
        if (桁数 && 字.length > 桁数) return 誤('#NUM!');
        while (桁数 && 字.length < 桁数) 字 = '0' + 字;
      }
    }
    return 計.字(字);
  }
  /** ★他の 進数 → 十進★（★10桁の 補数は 負★） */
  function 十進へ(引数たち, 手, 所, 基) {
    if (!引数たち.length) return 誤('#VALUE!');
    var t = 字の引数(引数たち[0], 手, 所);
    if (t.誤) return t.誤;
    var x = t.字.trim().toUpperCase();
    if (x === '') return 数(0);
    if (x.length > 10) return 誤('#NUM!');
    var 別 = { 2: /^[01]+$/, 8: /^[0-7]+$/, 16: /^[0-9A-F]+$/ }[基];
    if (!別.test(x)) return 誤('#NUM!');
    var n = parseInt(x, 基);
    var 周 = Math.pow(基, 10);
    if (x.length === 10 && n >= 周 / 2) n -= 周;
    return 数(n);
  }

  /* ══ ★複素数の 台★ ══（2026-09-16）
       ★字 → 数の 組★／★数の 組 → 字★ … ★ここだけ★（作る道を 2本に しない） */
  var 複素の形 = /^([+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?)?([+-](?:\d+\.?\d*|\.\d+)?(?:[eE][-+]?\d+)?)?([ij])?$/;
  /** ★字を 複素数に する★ */
  function 字を複素に(t) {
    var x = String(t).trim();
    if (x === '') return { re: 0, im: 0, 印: 'i' };
    var m = 複素の形.exec(x);
    if (!m) return null;
    if (!m[3]) {
      var n = Number(x);
      if (!isFinite(n)) return null;
      return { re: n, im: 0, 印: 'i' };
    }
    var 印 = m[3];
    var re = 0, im = 0;
    if (m[2] !== undefined && m[2] !== '') {
      re = (m[1] === undefined || m[1] === '') ? 0 : Number(m[1]);
      im = (m[2] === '+' || m[2] === '-') ? Number(m[2] + '1') : Number(m[2]);
    } else if (m[1] !== undefined && m[1] !== '') {
      im = Number(m[1]);
    } else {
      im = 1;
    }
    if (!isFinite(re) || !isFinite(im)) return null;
    return { re: re, im: im, 印: 印 };
  }
  /** ★実 Excel と 同じく 15桁に 丸めて 字に する★ */
  function 十五桁の字(n) {
    if (!isFinite(n)) return null;
    if (n === 0) return '0';
    return String(Number(n.toPrecision(15))).replace('e', 'E');
  }
  /** ★数の 組を 字に する★（`0.5+i` の ように ★1 は 書きません★） */
  function 複素を字に(re, im, 印) {
    印 = 印 || 'i';
    var R = 十五桁の字(re), I = 十五桁の字(im);
    if (R === null || I === null) return null;
    if (Number(I) === 0) return R;
    var 虚 = (Math.abs(Number(I)) === 1) ? 印 : (I.replace(/^-/, '') + 印);
    var 符 = Number(I) < 0 ? '-' : '+';
    if (Number(R) === 0) return (符 === '-' ? '-' : '') + 虚;
    return R + 符 + 虚;
  }
  function 複素を割る(a, b) {
    var d = b.re * b.re + b.im * b.im;
    if (d === 0) return null;
    /* ★★割る 方が 無限なら 0★★（2026-09-16 実測）
         `=IMSECH(D1)`（D1＝45292）… `cosh(45292)` は ★無限★
         ⇒ 前は `Infinity/Infinity` で `NaN` ⇒ #NUM!／実Excel は ★0★
       ⇒★`NaN` は これで 3つ目★（TRUNC／PMT／ここ）＝★大きい 数の 無限は 型で 出る★ */
    if (!isFinite(d)) {
      if (isFinite(a.re) && isFinite(a.im)) return { re: 0, im: 0 };
      return null;
    }
    return { re: (a.re * b.re + a.im * b.im) / d, im: (a.im * b.re - a.re * b.im) / d };
  }
  function 複素の対数(z) {
    var r = Math.sqrt(z.re * z.re + z.im * z.im);
    if (r === 0) return null;
    return { re: Math.log(r), im: Math.atan2(z.im, z.re) };
  }
  function 複素の正接(z) {
    var c = { re: Math.cos(z.re) * Math.cosh(z.im), im: -Math.sin(z.re) * Math.sinh(z.im) };
    var s2 = { re: Math.sin(z.re) * Math.cosh(z.im), im: Math.cos(z.re) * Math.sinh(z.im) };
    return 複素を割る(s2, c);
  }
  /** ★引数を 複素数に する★（数でも 字でも） */
  function 複素にする(引数, 手, 所) {
    var v = 場.ひとつに(引数, 所 && 所.今のマス);
    if (!v) return { 誤: 誤('#VALUE!') };
    if (v.型 === '誤') return { 誤: v };
    if (v.型 === '数') return { z: { re: v.値, im: 0 }, 印: 'i' };
    if (v.型 === '真偽') return { 誤: 誤('#VALUE!') };
    if (v.型 === '空') return { z: { re: 0, im: 0 }, 印: 'i' };
    var z = 字を複素に(v.値);
    if (!z) {
      /* ★★複素数の 形で なければ ★数として★ 読み直します★★（2026-09-16 実測）
           `=IMSUM("12:30")` … 実Excel ★0.520833333333333★（★時刻の 字を 数に した★）
           ⇒★前は #NUM! を 返して いました★（22本 赤）
           ⇒★他の 関数と 同じ 決め★＝★直に 書いた 字は 数に する★ */
      var n2 = 計.数にする(v, 手);
      if (n2 && n2.型 !== '誤') return { z: { re: n2.値, im: 0 }, 印: 'i' };
      return { 誤: 誤('#NUM!') };
    }
    return { z: { re: z.re, im: z.im }, 印: z.印 };
  }
  /** ★複素数を 1つ 取って ★数★を 返す★ */
  function 複素ひとつ(引数たち, 手, 所, 続き) {
    if (!引数たち.length) return 誤('#VALUE!');
    var z = 複素にする(引数たち[0], 手, 所);
    if (z.誤) return z.誤;
    return 続き(z.z);
  }
  /** ★複素数を 1つ 取って ★複素数★を 返す★ */
  function 複素で(引数たち, 手, 所, 続き) {
    if (!引数たち.length) return 誤('#VALUE!');
    var z = 複素にする(引数たち[0], 手, 所);
    if (z.誤) return z.誤;
    var 出 = 続き(z.z);
    if (!出) return 誤('#NUM!');
    var t = 複素を字に(出.re, 出.im, z.印);
    if (t === null) return 誤('#NUM!');
    return 計.字(t);
  }
  /** ★複素数を 2つ★ */
  function 複素ふたつ(引数たち, 手, 所, 続き) {
    if (引数たち.length < 2) return 誤('#VALUE!');
    var a = 複素にする(引数たち[0], 手, 所);
    if (a.誤) return a.誤;
    var b = 複素にする(引数たち[1], 手, 所);
    if (b.誤) return b.誤;
    var 出 = 続き(a.z, b.z);
    if (!出) return 誤('#NUM!');
    var t = 複素を字に(出.re, 出.im, a.印);
    if (t === null) return 誤('#NUM!');
    return 計.字(t);
  }
  /** ★IMSUM／IMPRODUCT★（★何個でも★） */
  function 複素を重ねる(引数たち, 手, 所, 種) {
    if (!引数たち.length) return 誤('#VALUE!');
    var 出 = 種 === 'SUM' ? { re: 0, im: 0 } : { re: 1, im: 0 };
    var 印 = 'i', 誤り = null;
    ほどく(引数たち, function (v) {
      if (誤り) return false;
      if (v.型 === '誤') { 誤り = v; return false; }
      if (v.型 === '空') return true;
      var z;
      if (v.型 === '数') z = { re: v.値, im: 0, 印: 'i' };
      else if (v.型 === '字') {
        z = 字を複素に(v.値);
        if (!z) {                                  /* ★複素数の 形で なければ 数として 読む★（上と 同じ 決め） */
          var n3 = 計.数にする(v, 手);
          if (n3 && n3.型 !== '誤') z = { re: n3.値, im: 0, 印: 'i' };
          else { 誤り = 誤('#NUM!'); return false; }
        }
      }
      else { 誤り = 誤('#VALUE!'); return false; }
      if (z.印 === 'j') 印 = 'j';
      if (種 === 'SUM') { 出 = { re: 出.re + z.re, im: 出.im + z.im }; }
      else { 出 = { re: 出.re * z.re - 出.im * z.im, im: 出.re * z.im + 出.im * z.re }; }
      return true;
    });
    if (誤り) return 誤り;
    var t = 複素を字に(出.re, 出.im, 印);
    if (t === null) return 誤('#NUM!');
    return 計.字(t);
  }

  /** ★通し番号から 年月日を 取り出す★（★書式の 台を 呼ぶ★） */
  function 日から(引数たち, 手, 所, 続き) {
    if (!引数たち.length) return 誤('#VALUE!');
    var t = 数の引数(引数たち[0], 手, 所);
    if (t.誤) return t.誤;
    if (t.数 < 0) return 誤('#NUM!');
    return 数(続き(書.数から日(t.数)));
  }
  /** ★2次元の 表を 平らに する★（行から） */
  function 平らに(A) {
    if (!A || !A.length) return null;
    var 出 = [];
    for (var r = 0; r < A.length; r++) for (var c = 0; c < A[r].length; c++) 出.push(A[r][c]);
    return 出;
  }

  /** ★標準正規の 上側★（★　1 から 引かない★）
        ★★`1 - 下側` だと 尾っぽで 桁が 落ちます★★（2026-09-16 実測）
          `=ZTEST(A1:A5,A1)` … うち ★0.0023388674905233664★
                                実Excel ★0.0023388674905236322★
          ＝0.9976… を 1 から 引いて いた＝★上の 3桁が 消える★
        ⇒★補誤差関数を ★不完全ガンマ★で 直に 出す★
          erfc(x) ＝ Q(1/2, x²) ＝ `上のガンマ(0.5, x*x)` */
  function 補誤差関数(x) {
    if (x < 0) return 2 - 補誤差関数(-x);
    if (x < 0.7) return 1 - 誤差関数(x);   /* ★ここは 桁が 落ちない（答え ≥ 0.32）★ */
    /* ★連分数★… ★60項だと x=1 で ずれ 2.9e-9、300項で 2.2e-16★（2026-09-16 実測） */
    var s = 0;
    for (var k = 300; k >= 1; k--) s = k / 2 / (x + s);
    return Math.exp(-x * x) / Math.sqrt(Math.PI) / (x + s);
  }
  /** ★標準正規の 上側★／★下側★
        ★★どちらも 「大きい 数から 引く」を 避けます★★
          中央寄り（|z| ≤ 1）… 誤差関数で（答えが 0.16 以上＝桁が 落ちない）
          尾っぽ（|z| > 1）  … 補誤差関数で　直に */
  function 標準正規の上(z) {
    if (z > 1) return 0.5 * 補誤差関数(z / Math.SQRT2);
    return 0.5 * (1 - 誤差関数(z / Math.SQRT2));
  }
  function 標準正規の下(z) {
    if (z < -1) return 0.5 * 補誤差関数(-z / Math.SQRT2);
    return 0.5 * (1 + 誤差関数(z / Math.SQRT2));
  }
  /** ★標準正規の 密度★ */
  function 標準正規の密度(z) { return Math.exp(-z * z / 2) / Math.sqrt(2 * Math.PI); }
  /** ★標準正規の 逆★（Acklam の 見当 → ニュートンで 磨く）
        ★ここだけは 二分で 探さない★
          ＝密度が そのまま 傾きに なるので ニュートンの方が 早くて 深い */
  function 標準正規の逆(p) {
    if (p <= 0) return -Infinity;
    if (p >= 1) return Infinity;
    var a = [-39.69683028665376, 220.9460984245205, -275.9285104469687,
      138.3577518672690, -30.66479806614716, 2.506628277459239];
    var b = [-54.47609879822406, 161.5858368580409, -155.6989798598866,
      66.80131188771972, -13.28068155288572];
    var c = [-0.007784894002430293, -0.3223964580411365, -2.400758277161838,
      -2.549732539343734, 4.374664141464968, 2.938163982698783];
    var d = [0.007784695709041462, 0.3224671290700398, 2.445134137142996,
      3.754408661907416];
    var 境 = 0.02425, x, q, r;
    if (p < 境) {
      q = Math.sqrt(-2 * Math.log(p));
      x = (((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) /
        ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1);
    } else if (p <= 1 - 境) {
      q = p - 0.5; r = q * q;
      x = (((((a[0] * r + a[1]) * r + a[2]) * r + a[3]) * r + a[4]) * r + a[5]) * q /
        (((((b[0] * r + b[1]) * r + b[2]) * r + b[3]) * r + b[4]) * r + 1);
    } else {
      q = Math.sqrt(-2 * Math.log(1 - p));
      x = -(((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) /
        ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1);
    }
    for (var i = 0; i < 3; i++) {
      var u = 標準正規の密度(x);
      if (u === 0) break;
      x = x - (標準正規の下(x) - p) / u;
    }
    return x;
  }
  /** ★NORM.DIST／NORMDIST★ */
  function 正規の分布(引数たち, 手, 所) {
    if (引数たち.length < 4) return 誤('#VALUE!');
    var g = 金の引数([引数たち[0], 引数たち[1], 引数たち[2]], 手, 所, 3, 3);
    if (g.誤) return g.誤;
    var 判 = 真偽にする(場.ひとつに(引数たち[3], 所 && 所.今のマス));
    if (判.誤) return 判.誤;
    var 散 = g.数[2];
    if (散 <= 0) return 誤('#NUM!');
    var z = (g.数[0] - g.数[1]) / 散;
    return 数(判.真 ? 標準正規の下(z) : 標準正規の密度(z) / 散);
  }
  /** ★NORM.INV／NORMINV★ */
  function 正規の逆(引数たち, 手, 所) {
    var g = 金の引数(引数たち, 手, 所, 3, 3);
    if (g.誤) return g.誤;
    if (g.数[2] <= 0 || g.数[0] <= 0 || g.数[0] >= 1) return 誤('#NUM!');
    return 数(g.数[1] + g.数[2] * 標準正規の逆(g.数[0]));
  }
  /** ★対数正規の 逆★（LOGINV／LOGNORM.INV） */
  function 対数正規の逆(引数たち, 手, 所) {
    var g = 金の引数(引数たち, 手, 所, 3, 3);
    if (g.誤) return g.誤;
    if (g.数[2] <= 0 || g.数[0] <= 0 || g.数[0] >= 1) return 誤('#NUM!');
    return 数(Math.exp(g.数[1] + g.数[2] * 標準正規の逆(g.数[0])));
  }
  /** ★組み合わせ★（★きっちり出る 間は 掛け算で★）
        ★入らなくなったら null★（呼ぶ 側が 対数に 逃げる） */
  function 組み合わせの数(n, k) {
    if (k < 0 || k > n) return 0;
    if (k > n - k) k = n - k;
    var c = 1;
    for (var i = 1; i <= k; i++) {
      c = c * (n - k + i) / i;
      if (!isFinite(c) || c > 9007199254740991) return null;
    }
    return Math.round(c);
  }
  /** ★二項の 重み★（一点）
        ★★対数で 出すと 15桁目が 実Excel と ずれる★★（2026-09-16 実測）
          `=BINOM.DIST(6,10,0.5,FALSE)` … 実Excel ★0.205078125★（＝210/1024 の きっちり）
          対数で 出すと ★0.2050781250000006★
        ⇒★組み合わせが 整数で 出る 間は 掛け算で 出す★ */
  function 二項の重み(k, n, p) {
    if (p <= 0) return k === 0 ? 1 : 0;
    if (p >= 1) return k === n ? 1 : 0;
    var c = 組み合わせの数(n, k);
    if (c !== null) {
      var v = c * Math.pow(p, k) * Math.pow(1 - p, n - k);
      if (isFinite(v) && v > 0) return v;
    }
    return Math.exp(組み合わせの対数(n, k) + k * Math.log(p) + (n - k) * Math.log(1 - p));
  }
  /** ★BINOM.DIST／BINOMDIST★ */
  function 二項の分布(引数たち, 手, 所) {
    if (引数たち.length < 4) return 誤('#VALUE!');
    var g = 金の引数([引数たち[0], 引数たち[1], 引数たち[2]], 手, 所, 3, 3);
    if (g.誤) return g.誤;
    var 判 = 真偽にする(場.ひとつに(引数たち[3], 所 && 所.今のマス));
    if (判.誤) return 判.誤;
    var x = 切り捨て(g.数[0]), n = 切り捨て(g.数[1]), p = g.数[2];
    if (x < 0 || x > n || p < 0 || p > 1) return 誤('#NUM!');
    if (!判.真) return 数(二項の重み(x, n, p));
    var 和 = 0;
    for (var k = 0; k <= x; k++) 和 += 二項の重み(k, n, p);
    return 数(和);
  }
  /** ★BINOM.INV／CRITBINOM★ */
  function 二項の逆(引数たち, 手, 所) {
    var g = 金の引数(引数たち, 手, 所, 3, 3);
    if (g.誤) return g.誤;
    var n = 切り捨て(g.数[0]), p = g.数[1], a = g.数[2];
    if (n < 0 || p < 0 || p > 1 || a <= 0 || a >= 1) return 誤('#NUM!');
    var 和 = 0;
    for (var k = 0; k <= n; k++) {
      和 += 二項の重み(k, n, p);
      if (和 >= a) return 数(k);
    }
    return 誤('#NUM!');
  }
  /** ★NEGBINOM.DIST／NEGBINOMDIST★ */
  function 負の二項(引数たち, 手, 所) {
    if (引数たち.length < 3) return 誤('#VALUE!');
    var g = 金の引数([引数たち[0], 引数たち[1], 引数たち[2]], 手, 所, 3, 3);
    if (g.誤) return g.誤;
    var 続 = true;
    if (引数たち.length > 3) {
      var 判 = 真偽にする(場.ひとつに(引数たち[3], 所 && 所.今のマス));
      if (判.誤) return 判.誤;
      続 = 判.真;
    } else 続 = false;
    var f = 切り捨て(g.数[0]), sx = 切り捨て(g.数[1]), p = g.数[2];
    if (f < 0 || sx < 1 || p <= 0 || p > 1) return 誤('#NUM!');
    var 一点 = function (k) {
      var c = 組み合わせの数(k + sx - 1, k);
      if (c !== null) {
        var v = c * Math.pow(p, sx) * Math.pow(1 - p, k);
        if (isFinite(v) && v > 0) return v;
      }
      return Math.exp(組み合わせの対数(k + sx - 1, k) + sx * Math.log(p) + k * Math.log(1 - p));
    };
    if (!続) return 数(一点(f));
    var 和 = 0;
    for (var i = 0; i <= f; i++) 和 += 一点(i);
    return 数(和);
  }
  /** ★BETA.DIST（`新しい形` true）／BETADIST（false＝積み上げだけ）★ */
  function ベータの分布(引数たち, 手, 所, 新しい形) {
    if (引数たち.length < 3) return 誤('#VALUE!');
    var g = 金の引数([引数たち[0], 引数たち[1], 引数たち[2]], 手, 所, 3, 3);
    if (g.誤) return g.誤;
    var 積 = true, 頭 = 新しい形 ? 4 : 3;
    if (新しい形) {
      if (引数たち.length < 4) return 誤('#VALUE!');
      var 判 = 真偽にする(場.ひとつに(引数たち[3], 所 && 所.今のマス));
      if (判.誤) return 判.誤;
      積 = 判.真;
    }
    var A = 0, B = 1;
    if (引数たち.length > 頭) {
      var ta = 数の引数(引数たち[頭], 手, 所);
      if (ta.誤) return ta.誤;
      A = ta.数;
    }
    if (引数たち.length > 頭 + 1) {
      var tb = 数の引数(引数たち[頭 + 1], 手, 所);
      if (tb.誤) return tb.誤;
      B = tb.数;
    }
    var x = g.数[0], al = g.数[1], be = g.数[2];
    if (al <= 0 || be <= 0 || B <= A) return 誤('#NUM!');
    if (x < A || x > B) return 誤('#NUM!');
    var y = (x - A) / (B - A);
    if (積) return 数(ベータの下(al, be, y));
    if (y <= 0 || y >= 1) return 数(0);
    return 数(Math.exp(ガンマの対数(al + be) - ガンマの対数(al) - ガンマの対数(be)
      + (al - 1) * Math.log(y) + (be - 1) * Math.log(1 - y)) / (B - A));
  }
  /** ★BETA.INV／BETAINV★ */
  function ベータの逆(引数たち, 手, 所) {
    if (引数たち.length < 3) return 誤('#VALUE!');
    var g = 金の引数([引数たち[0], 引数たち[1], 引数たち[2]], 手, 所, 3, 3);
    if (g.誤) return g.誤;
    var A = 0, B = 1;
    if (引数たち.length > 3) {
      var ta = 数の引数(引数たち[3], 手, 所);
      if (ta.誤) return ta.誤;
      A = ta.数;
    }
    if (引数たち.length > 4) {
      var tb = 数の引数(引数たち[4], 手, 所);
      if (tb.誤) return tb.誤;
      B = tb.数;
    }
    var p = g.数[0], al = g.数[1], be = g.数[2];
    if (al <= 0 || be <= 0 || B <= A || p <= 0 || p > 1) return 誤('#NUM!');
    var y = 二分で逆(function (t) { return 1 - ベータの下(al, be, t); }, 1 - p, 0, 1);
    return 数(A + y * (B - A));
  }
  /** ★GAMMA.INV／GAMMAINV★ */
  function ガンマの逆(引数たち, 手, 所) {
    var g = 金の引数(引数たち, 手, 所, 3, 3);
    if (g.誤) return g.誤;
    var p = g.数[0], al = g.数[1], be = g.数[2];
    if (p < 0 || p >= 1 || al <= 0 || be <= 0) return 誤('#NUM!');
    if (p === 0) return 数(0);
    var x = 二分で逆(function (t) { return 上のガンマ(al, t); }, 1 - p, 0, 1e10);
    return 数(x * be);
  }
  /** ★F.INV（`右か` false）／FINV（true＝右側）★ */
  function Fの逆(引数たち, 手, 所, 右か) {
    var g = 金の引数(引数たち, 手, 所, 3, 3);
    if (g.誤) return g.誤;
    var p = g.数[0], d1 = 切り捨て(g.数[1]), d2 = 切り捨て(g.数[2]);
    if (p < 0 || p > 1 || d1 < 1 || d2 < 1) return 誤('#NUM!');
    var 右の値 = 右か ? p : (1 - p);
    if (右の値 <= 0) return 誤('#NUM!');
    if (右の値 >= 1) return 数(0);
    return 数(二分で逆(function (x) {
      return ベータの下(d2 / 2, d1 / 2, d2 / (d2 + d1 * x));
    }, 右の値, 0, 1e10));
  }
  /** ★信頼の 幅★（`t か` true＝CONFIDENCE.T） */
  function 信頼の幅(引数たち, 手, 所, tか) {
    var g = 金の引数(引数たち, 手, 所, 3, 3);
    if (g.誤) return g.誤;
    var al = g.数[0], 散 = g.数[1], n = 切り捨て(g.数[2]);
    if (al <= 0 || al >= 1 || 散 <= 0 || n < 1) return 誤('#NUM!');
    if (!tか) return 数(標準正規の逆(1 - al / 2) * 散 / Math.sqrt(n));
    if (n === 1) return 誤('#DIV/0!');
    var t = 二分で逆(function (x) { return 2 * tの右(x, n - 1); }, al, 0, 1e10);
    return 数(t * 散 / Math.sqrt(n));
  }
  /** ★標本の 分散★（n-1 で 割る） */
  function 標本の分散(a) {
    var n = a.length;
    if (n < 2) return null;
    var s = 0, i;
    for (i = 0; i < n; i++) s += a[i];
    var m = s / n, v = 0;
    for (i = 0; i < n; i++) v += (a[i] - m) * (a[i] - m);
    return v / (n - 1);
  }
  /** ★対に して 散らばりを 出す★（回帰の 土台）
        ★引数たち[0]＝X／[1]＝Y★（SLOPE などは Y が 先なので 逆に 渡す） */
  function 対の数(引数たち, 手) {
    if (引数たち.length < 2) return { 誤: 誤('#VALUE!') };
    var X = 数を拾う([引数たち[0]], 手);
    if (X.誤) return { 誤: X.誤 };
    var Y = 数を拾う([引数たち[1]], 手);
    if (Y.誤) return { 誤: Y.誤 };
    var n = Math.min(X.数たち.length, Y.数たち.length);
    if (!n) return { 誤: 誤('#N/A') };
    var sx = 0, sy = 0, i;
    for (i = 0; i < n; i++) { sx += X.数たち[i]; sy += Y.数たち[i]; }
    var mx = sx / n, my = sy / n, Sxx = 0, Syy = 0, Sxy = 0;
    for (i = 0; i < n; i++) {
      var dx = X.数たち[i] - mx, dy = Y.数たち[i] - my;
      Sxx += dx * dx; Syy += dy * dy; Sxy += dx * dy;
    }
    return { n: n, mx: mx, my: my, Sxx: Sxx, Syy: Syy, Sxy: Sxy };
  }
  /** ★順位★（`平均か` true＝RANK.AVG） */
  function 順位(引数たち, 手, 所, 平均か) {
    if (引数たち.length < 2) return 誤('#VALUE!');
    var t = 数の引数(引数たち[0], 手, 所);
    if (t.誤) return t.誤;
    var 拾 = 数を拾う([引数たち[1]], 手);
    if (拾.誤) return 拾.誤;
    var 降 = true;
    if (引数たち.length > 2) {
      var o = 数の引数(引数たち[2], 手, 所);
      if (o.誤) return o.誤;
      降 = (o.数 === 0);
    }
    var a = 拾.数たち, 上 = 0, 同 = 0;
    for (var i = 0; i < a.length; i++) {
      if (a[i] === t.数) 同++;
      else if (降 ? (a[i] > t.数) : (a[i] < t.数)) 上++;
    }
    if (!同) return 誤('#N/A');
    return 数(平均か ? (上 + 1 + (同 - 1) / 2) : (上 + 1));
  }
  /** ★外しの 分位★（PERCENTILE.EXC／QUARTILE.EXC）
        ★内の 分位との 違い★＝(n+1) で 数える＝端を 含まない */
  function 外の分位(引数たち, 手, 所, 四分か) {
    if (引数たち.length < 2) return 誤('#VALUE!');
    var 拾 = 数を拾う([引数たち[0]], 手);
    if (拾.誤) return 拾.誤;
    var t = 数の引数(引数たち[1], 手, 所);
    if (t.誤) return t.誤;
    var a = 拾.数たち.slice().sort(function (x, y) { return x - y; });
    var n = a.length;
    if (!n) return 誤('#NUM!');
    if (四分か && (切り捨て(t.数) < 1 || 切り捨て(t.数) > 3)) return 誤('#NUM!');
    var p = 四分か ? (切り捨て(t.数) / 4) : t.数;
    var i = p * (n + 1);
    if (i < 1 || i > n) return 誤('#NUM!');
    var 下 = Math.floor(i), 余 = i - 下;
    if (下 >= n) return 数(a[n - 1]);
    return 数(a[下 - 1] + (a[下] - a[下 - 1]) * 余);
  }
  /** ★Z の 検定★（右側） */
  function Zの検定(引数たち, 手, 所) {
    if (引数たち.length < 2) return 誤('#VALUE!');
    var 拾 = 数を拾う([引数たち[0]], 手);
    if (拾.誤) return 拾.誤;
    var t = 数の引数(引数たち[1], 手, 所);
    if (t.誤) return t.誤;
    var a = 拾.数たち, n = a.length, i;
    if (n < 1) return 誤('#N/A');
    var s = 0;
    for (i = 0; i < n; i++) s += a[i];
    var m = s / n, シグマ;
    if (引数たち.length > 2) {
      var g = 数の引数(引数たち[2], 手, 所);
      if (g.誤) return g.誤;
      シグマ = g.数;
    } else {
      var v = 標本の分散(a);
      if (v === null) return 誤('#DIV/0!');
      シグマ = Math.sqrt(v);
    }
    if (シグマ <= 0) return 誤('#NUM!');
    return 数(標準正規の上((m - t.数) / (シグマ / Math.sqrt(n))));
  }
  /** ★カイ二乗の 検定★（CHISQ.TEST／CHITEST） */
  function カイ二乗の検定(引数たち, 手, 所) {
    if (引数たち.length < 2) return 誤('#VALUE!');
    var A = 表にする(引数たち[0], 所), E = 表にする(引数たち[1], 所);
    if (!A || !E || !A.length || !E.length) return 誤('#VALUE!');
    var r = A.length, c = A[0].length;
    if (E.length !== r || E[0].length !== c) return 誤('#N/A');
    var x2 = 0;
    for (var i = 0; i < r; i++) for (var j = 0; j < c; j++) {
      var a = A[i][j], e = E[i][j];
      if (a && a.型 === '誤') return a;
      if (e && e.型 === '誤') return e;
      if (!a || a.型 !== '数' || !e || e.型 !== '数') return 誤('#VALUE!');
      if (e.値 === 0) return 誤('#DIV/0!');
      x2 += (a.値 - e.値) * (a.値 - e.値) / e.値;
    }
    var 自 = (r === 1 || c === 1) ? (Math.max(r, c) - 1) : ((r - 1) * (c - 1));
    if (自 < 1) return 誤('#N/A');
    return 数(上のガンマ(自 / 2, x2 / 2));
  }
  /** ★F の 検定★（両側） */
  function Fの検定(引数たち, 手) {
    if (引数たち.length < 2) return 誤('#VALUE!');
    var X = 数を拾う([引数たち[0]], 手);
    if (X.誤) return X.誤;
    var Y = 数を拾う([引数たち[1]], 手);
    if (Y.誤) return Y.誤;
    var v1 = 標本の分散(X.数たち), v2 = 標本の分散(Y.数たち);
    if (v1 === null || v2 === null || v1 === 0 || v2 === 0) return 誤('#DIV/0!');
    var n1 = X.数たち.length - 1, n2 = Y.数たち.length - 1;
    var f = v1 / v2;
    var 右 = ベータの下(n2 / 2, n1 / 2, n2 / (n2 + n1 * f));
    return 数(2 * Math.min(右, 1 - 右));
  }
  /** ★t の 検定★（型 1＝対に して／2＝散らばりが 同じ／3＝違う） */
  function tの検定(引数たち, 手, 所) {
    if (引数たち.length < 4) return 誤('#VALUE!');
    var X = 数を拾う([引数たち[0]], 手);
    if (X.誤) return X.誤;
    var Y = 数を拾う([引数たち[1]], 手);
    if (Y.誤) return Y.誤;
    var g = 金の引数([引数たち[2], 引数たち[3]], 手, 所, 2, 2);
    if (g.誤) return g.誤;
    var 尾 = 切り捨て(g.数[0]), 型 = 切り捨て(g.数[1]);
    if (尾 !== 1 && 尾 !== 2) return 誤('#NUM!');
    if (型 < 1 || 型 > 3) return 誤('#NUM!');
    var a = X.数たち, b = Y.数たち, i, t, 自;
    if (型 === 1) {
      var n = Math.min(a.length, b.length);
      if (n < 2) return 誤('#DIV/0!');
      var d = [], sd = 0;
      for (i = 0; i < n; i++) { d.push(a[i] - b[i]); sd += a[i] - b[i]; }
      var vd = 標本の分散(d);
      if (vd === null || vd === 0) return 誤('#DIV/0!');
      t = (sd / n) / Math.sqrt(vd / n);
      自 = n - 1;
    } else {
      var n1 = a.length, n2 = b.length;
      if (n1 < 2 || n2 < 2) return 誤('#DIV/0!');
      var m1 = 0, m2 = 0;
      for (i = 0; i < n1; i++) m1 += a[i];
      for (i = 0; i < n2; i++) m2 += b[i];
      m1 /= n1; m2 /= n2;
      var v1 = 標本の分散(a), v2 = 標本の分散(b);
      if (型 === 2) {
        var vp = ((n1 - 1) * v1 + (n2 - 1) * v2) / (n1 + n2 - 2);
        if (vp === 0) return 誤('#DIV/0!');
        t = (m1 - m2) / Math.sqrt(vp * (1 / n1 + 1 / n2));
        自 = n1 + n2 - 2;
      } else {
        var w1 = v1 / n1, w2 = v2 / n2;
        if (w1 + w2 === 0) return 誤('#DIV/0!');
        t = (m1 - m2) / Math.sqrt(w1 + w2);
        自 = (w1 + w2) * (w1 + w2) / (w1 * w1 / (n1 - 1) + w2 * w2 / (n2 - 1));
      }
    }
    var p = tの右(Math.abs(t), 自);
    return 数(尾 === 2 ? 2 * p : p);
  }

  /** ★組み合わせの 対数★（大きい 数でも 溢れない） */
  function 組み合わせの対数(n, k) {
    if (k < 0 || k > n) return -Infinity;
    return ガンマの対数(n + 1) - ガンマの対数(k + 1) - ガンマの対数(n - k + 1);
  }
  /** ★二分で 逆を 探す★（★減る 関数★を 渡す） */
  function 二分で逆(f, p, 下, 上) {
    /* ★★止める 所を 「1e-13」に して いました★★（2026-09-16 実測）
         `=BETAINV(0.5,1,2)` … うち ★0.2928932188134752★
                               実Excel ★0.29289321881345248★（＝1-1/√2）
         ＝★探し方の 違いでは なく ★止めるのが 早すぎた★だけ★
       ⇒★★隣り合う 浮動小数に なるまで 割る★★
         ＝これ以上 割れない 所で 自然に 止まる（桁を 手で 決めない） */
    for (var i = 0; i < 400; i++) {
      var 中 = (下 + 上) / 2;
      if (中 === 下 || 中 === 上) break;
      if (f(中) > p) 下 = 中; else 上 = 中;
    }
    return (下 + 上) / 2;
  }
  /** ★χ² の 逆（右側）★ */
  function カイ二乗の逆(引数たち, 手, 所) {
    var g = 金の引数(引数たち, 手, 所, 2, 2);
    if (g.誤) return g.誤;
    var p = g.数[0], v = 切り捨て(g.数[1]);
    if (p < 0 || p > 1 || v < 1) return 誤('#NUM!');
    return 数(二分で逆(function (x) { return 上のガンマ(v / 2, x / 2); }, p, 0, 1e10));
  }
  /** ★t の 逆（両側）★ */
  function tの逆両側(引数たち, 手, 所) {
    var g = 金の引数(引数たち, 手, 所, 2, 2);
    if (g.誤) return g.誤;
    var p = g.数[0], v = 切り捨て(g.数[1]);
    if (p <= 0 || p > 1 || v < 1) return 誤('#NUM!');
    return 数(二分で逆(function (x) { return 2 * tの右(x, v); }, p, 0, 1e10));
  }
  /** ★ワイブル★ */
  function ワイブル(引数たち, 手, 所) {
    if (引数たち.length < 4) return 誤('#VALUE!');
    var g = 金の引数([引数たち[0], 引数たち[1], 引数たち[2]], 手, 所, 3, 3);
    if (g.誤) return g.誤;
    var 判 = 真偽にする(場.ひとつに(引数たち[3], 所 && 所.今のマス));
    if (判.誤) return 判.誤;
    var x = g.数[0], a = g.数[1], b = g.数[2];
    if (x < 0 || a <= 0 || b <= 0) return 誤('#NUM!');
    if (判.真) return 数(1 - Math.exp(-Math.pow(x / b, a)));
    return 数(a / Math.pow(b, a) * Math.pow(x, a - 1) * Math.exp(-Math.pow(x / b, a)));
  }

  /** ★SUMX2MY2／SUMX2PY2／SUMXMY2★ */
  function 二つの二乗(引数たち, 手, 種) {
    if (引数たち.length < 2) return 誤('#VALUE!');
    var X = 数を拾う([引数たち[0]], 手);
    if (X.誤) return X.誤;
    var Y = 数を拾う([引数たち[1]], 手);
    if (Y.誤) return Y.誤;
    var n = Math.min(X.数たち.length, Y.数たち.length);
    if (!n) return 誤('#N/A');
    var 和 = 0;
    for (var i = 0; i < n; i++) {
      var x = X.数たち[i], y = Y.数たち[i];
      if (種 === 'X2MY2') 和 += x * x - y * y;
      else if (種 === 'X2PY2') 和 += x * x + y * y;
      else 和 += (x - y) * (x - y);
    }
    return 数(和);
  }
  /** ★共分散★（`母集団か` true＝n で 割る） */
  function 共分散(引数たち, 手, 母集団か) {
    if (引数たち.length < 2) return 誤('#VALUE!');
    var X = 数を拾う([引数たち[0]], 手);
    if (X.誤) return X.誤;
    var Y = 数を拾う([引数たち[1]], 手);
    if (Y.誤) return Y.誤;
    var n = Math.min(X.数たち.length, Y.数たち.length);
    if (n < (母集団か ? 1 : 2)) return 誤('#DIV/0!');
    var sx = 0, sy = 0, i;
    for (i = 0; i < n; i++) { sx += X.数たち[i]; sy += Y.数たち[i]; }
    var mx = sx / n, my = sy / n, 和 = 0;
    for (i = 0; i < n; i++) 和 += (X.数たち[i] - mx) * (Y.数たち[i] - my);
    return 数(和 / (母集団か ? n : n - 1));
  }
  /** ★進数 → 進数★（★10桁の 補数を 通す★） */
  function 進数を換える(引数たち, 手, 所, 元, 先) {
    if (!引数たち.length) return 誤('#VALUE!');
    var t = 字の引数(引数たち[0], 手, 所);
    if (t.誤) return t.誤;
    var x = t.字.trim().toUpperCase();
    if (x === '') return 計.字('0');
    if (x.length > 10) return 誤('#NUM!');
    var 別 = { 2: /^[01]+$/, 8: /^[0-7]+$/, 16: /^[0-9A-F]+$/ }[元];
    if (!別.test(x)) return 誤('#NUM!');
    var n = parseInt(x, 元);
    var 周 = Math.pow(元, 10);
    if (x.length === 10 && n >= 周 / 2) n -= 周;
    var 上 = { 2: 512, 8: 536870912, 16: 549755813888 }[先];
    if (n < -上 || n >= 上) return 誤('#NUM!');
    var 字;
    if (n < 0) 字 = (Math.pow(先, 10) + n).toString(先).toUpperCase();
    else 字 = n.toString(先).toUpperCase();
    if (引数たち.length > 1) {
      var d = 数の引数(引数たち[1], 手, 所);
      if (d.誤) return d.誤;
      var 桁 = 切り捨て(d.数);
      if (桁 < 0) return 誤('#NUM!');
      if (n >= 0) {
        if (桁 && 字.length > 桁) return 誤('#NUM!');
        while (桁 && 字.length < 桁) 字 = '0' + 字;
      }
    }
    return 計.字(字);
  }

  /** ★指数分布★（累積か 密度か） */
  function 指数の分布(引数たち, 手, 所) {
    if (引数たち.length < 3) return 誤('#VALUE!');
    var x = 数の引数(引数たち[0], 手, 所);
    if (x.誤) return x.誤;
    var l = 数の引数(引数たち[1], 手, 所);
    if (l.誤) return l.誤;
    var 判 = 真偽にする(場.ひとつに(引数たち[2], 所 && 所.今のマス));
    if (判.誤) return 判.誤;
    if (x.数 < 0 || l.数 <= 0) return 誤('#NUM!');
    if (判.真) return 数(1 - Math.exp(-l.数 * x.数));
    return 数(l.数 * Math.exp(-l.数 * x.数));
  }
  /** ★ポアソン分布★ */
  function ポアソンの分布(引数たち, 手, 所) {
    if (引数たち.length < 3) return 誤('#VALUE!');
    var x = 数の引数(引数たち[0], 手, 所);
    if (x.誤) return x.誤;
    var m = 数の引数(引数たち[1], 手, 所);
    if (m.誤) return m.誤;
    var 判 = 真偽にする(場.ひとつに(引数たち[2], 所 && 所.今のマス));
    if (判.誤) return 判.誤;
    var k = 切り捨て(x.数);
    if (k < 0 || m.数 < 0) return 誤('#NUM!');
    if (!判.真) {
      return 数(Math.exp(-m.数 + k * Math.log(m.数) - ガンマの対数(k + 1)));
    }
    var 和 = 0;
    for (var i = 0; i <= k; i++) 和 += Math.exp(-m.数 + i * Math.log(m.数) - ガンマの対数(i + 1));
    return 数(和);
  }

  /** ★NETWORKDAYS.INTL／WORKDAY.INTL★
        ★休みの 型★ … 1＝土日（2 から 順に ずれる）／11〜17＝1日だけ
        ★字で 渡す 形★ … `0000011` の ように 7桁（月から）
        ★`日数か` true＝NETWORKDAYS.INTL／false＝WORKDAY.INTL★ */
  function 営業日の国(引数たち, 手, 所, 日数か) {
    if (引数たち.length < 2) return 誤('#VALUE!');
    var a = 数の引数(引数たち[0], 手, 所);
    if (a.誤) return a.誤;
    var b = 数の引数(引数たち[1], 手, 所);
    if (b.誤) return b.誤;
    /* ★休みの 型★（既定 1＝土日） */
    var 休型 = [0, 0, 0, 0, 0, 1, 1];      /* 月 火 水 木 金 土 日 */
    if (引数たち.length > 2) {
      var v = 場.ひとつに(引数たち[2], 所 && 所.今のマス);
      if (!v) return 誤('#VALUE!');
      if (v.型 === '誤') return v;
      if (v.型 === '字' && /^[01]{7}$/.test(String(v.値))) {
        休型 = String(v.値).split('').map(Number);
      } else {
        var n = 計.数にする(v, 手);
        if (n.型 === '誤') return 誤('#VALUE!');
        var k = 切り捨て(n.値);
        休型 = [0, 0, 0, 0, 0, 0, 0];
        if (k >= 1 && k <= 7) {
          /* 1＝土日／2＝日月／…（順に ずれる） */
          休型[(k + 4) % 7] = 1;
          休型[(k + 5) % 7] = 1;
        } else if (k >= 11 && k <= 17) {
          休型[(k - 11 + 6) % 7] = 1;   /* 11＝日曜だけ */
        } else return 誤('#NUM!');
      }
    }
    var 休 = 休みを拾う(引数たち, 3, 手);
    if (休.誤) return 休.誤;
    var 休みか = function (d) {
      var w = 書.数から日(d).w;          /* 0＝日曜 */
      return 休型[(w + 6) % 7] === 1 || 休.日.indexOf(d) >= 0;
    };
    if (日数か) {
      var 始 = 切り捨て(a.数), 終 = 切り捨て(b.数), 符 = 1;
      if (始 > 終) { var w2 = 始; 始 = 終; 終 = w2; 符 = -1; }
      var n2 = 0;
      for (var d2 = 始; d2 <= 終; d2++) if (!休みか(d2)) n2++;
      return 数(符 * n2);
    }
    var d3 = 切り捨て(a.数), 残 = 切り捨て(b.数);
    var 向 = 残 < 0 ? -1 : 1;
    残 = Math.abs(残);
    while (残 > 0) {
      d3 += 向;
      if (d3 < 0) return 誤('#NUM!');
      if (!休みか(d3)) 残--;
    }
    return 数(d3);
  }

  /** ★国倉短期証券の 3つ★（TBILLEQ／TBILLPRICE／TBILLYIELD）
        ★日数は ★満期 − 受渡★の 通し番号の 差★（切り捨ててから 引く）
        ★物差し★ `golden-346-2026-09-08.tsv`
          `=TBILLPRICE(0.5,1,2)` ★　99.444444444444443★ （日数 1日：INT(1)-INT(0.5)）
          `=TBILLEQ(2,3,4)`      ★　4.1011235955056176★
          `=TBILLYIELD(0.5,1,2)` ★　17640★ */
  function 短期証券(引数たち, 手, 所, 種) {
    var g = 金の引数(引数たち, 手, 所, 3, 3);
    if (g.誤) return g.誤;
    var 始 = 切り捨て(g.数[0]), 終 = 切り捨て(g.数[1]), x = g.数[2];
    var 日 = 終 - 始;
    if (日 <= 0 || 日 > 365) return 誤('#NUM!');
    if (x <= 0) return 誤('#NUM!');
    if (種 === 'EQ') {
      var 下 = 360 - x * 日;
      if (下 === 0) return 誤('#NUM!');
      return 数(365 * x / 下);
    }
    if (種 === 'PRICE') return 数(100 * (1 - x * 日 / 360));
    return 数((100 - x) / x * (360 / 日));          /* YIELD */
  }
  /** ★CUMIPMT／CUMPRINC★（★期の 内訳を 足すだけ★）
        ★物差し★ `=CUMIPMT(0.05,12,100,1,12,0)` ★− 35.3904920249785★
                 `=CUMPRINC(0.05,12,100,1,12,0)` ★− 100★ */
  function 積み上げの内訳(引数たち, 手, 所, 利息か) {
    var g = 金の引数(引数たち, 手, 所, 6, 6);
    if (g.誤) return g.誤;
    var r = g.数[0], n = 切り捨て(g.数[1]), pv = g.数[2];
    var 始 = 切り捨て(g.数[3]), 終 = 切り捨て(g.数[4]);
    var 型 = g.数[5];
    if (r <= 0 || n <= 0 || pv <= 0) return 誤('#NUM!');
    if (始 < 1 || 終 < 1 || 始 > 終 || 終 > n) return 誤('#NUM!');
    if (型 !== 0 && 型 !== 1) return 誤('#NUM!');
    var 和 = 0;
    for (var k = 始; k <= 終; k++) {
      var v = 期の内訳([{ 種: '直', 値: 数(r) }, { 種: '直', 値: 数(k) },
        { 種: '直', 値: 数(n) }, { 種: '直', 値: 数(pv) },
        { 種: '直', 値: 数(0) }, { 種: '直', 値: 数(型) }], 手, 所, 利息か);
      if (v.型 === '誤') return v;
      和 += v.値;
    }
    return 数(和);
  }
  /** ★DB★（★割合は 3桁に 丸める★）
        ★物差し★ `=DB(0.5,1,2,3,4)` ★− 0.22206021599999995★
          ★割合★ 1 − (1/0.5)^(1/2) → 丸めて − 0.414
          1期目 … 元 × 割合 × 月/12
          最後（life+1 期目）… 残り × 割合 × (12−月)/12 */
  function 定率法(引数たち, 手, 所) {
    var g = 金の引数(引数たち, 手, 所, 4, 5);
    if (g.誤) return g.誤;
    /* ★★寿命と 期は ★切り捨てません★★（2026-09-16 実測）
         `=DB(D1,D2,0.05,0.1,2)` … 実Excel ★-2845.8473333333332★
           ＝寿命 0.05 を そのまま 使い（1/0.05＝20乗）、期 0.1 は 1期目
         ★切り捨てると 寿命 0 で #NUM! に なって いました★ */
    var 元 = g.数[0], 残存 = g.数[1], 寿 = g.数[2];
    var 期 = g.数[3];
    var 月 = g.数.length > 4 ? 切り捨て(g.数[4]) : 12;
    if (寿 <= 0 || 期 <= 0) return 誤('#NUM!');
    if (月 < 1 || 月 > 12) return 誤('#NUM!');
    if (期 > 寿 + 1) return 誤('#NUM!');
    if (元 === 0) return 数(0);
    var 割 = 1 - Math.pow(残存 / 元, 1 / 寿);
    if (!isFinite(割)) return 誤('#NUM!');
    割 = Math.round(割 * 1000) / 1000;          /* ★実Excel は 3桁に 丸める★ */
    var 累 = 元 * 割 * 月 / 12;                  /* 1期目 */
    if (期 <= 1) return 数(累);
    var 合 = 累, 出 = 累;
    for (var k = 2; k <= 期; k++) {
      if (k === 寿 + 1) 出 = (元 - 合) * 割 * (12 - 月) / 12;
      else 出 = (元 - 合) * 割;
      合 += 出;
    }
    return 数(出);
  }
  /** ★DDB★（倍率法・既定 2倍） */
  function 倍率法(引数たち, 手, 所) {
    var g = 金の引数(引数たち, 手, 所, 4, 5);
    if (g.誤) return g.誤;
    var 元 = g.数[0], 残存 = g.数[1], 寿 = g.数[2], 期 = g.数[3];
    var 倍 = g.数.length > 4 ? g.数[4] : 2;
    if (元 < 0 || 残存 < 0 || 寿 <= 0 || 期 <= 0 || 倍 <= 0) return 誤('#NUM!');
    if (期 > 寿) return 数(0);
    var 合 = 0, 出 = 0;
    for (var k = 1; k <= 期; k++) {
      出 = (元 - 合) * 倍 / 寿;
      if (元 - 合 - 出 < 残存) 出 = Math.max(0, 元 - 合 - 残存);
      合 += 出;
    }
    return 数(出);
  }
  /** ★RATE★（★探して 近づける★）
        ★物差し★ `=RATE(12,-10,100)` ★ 0.02922854076913767★ */
  function 利を探す(引数たち, 手, 所) {
    var g = 金の引数(引数たち, 手, 所, 3, 6);
    if (g.誤) return g.誤;
    var n = g.数[0], pmt = g.数[1], pv = g.数[2];
    var fv = g.数.length > 3 ? g.数[3] : 0;
    var 期首 = g.数.length > 4 && g.数[4] !== 0 ? 1 : 0;
    var x = g.数.length > 5 ? g.数[5] : 0.1;
    if (n <= 0) return 誤('#NUM!');
    var f = function (r) {
      if (r === 0) return pv + pmt * n + fv;
      var t = Math.pow(1 + r, n);
      return pv * t + pmt * (1 + r * 期首) * (t - 1) / r + fv;
    };
    for (var i = 0; i < 100; i++) {
      var y = f(x);
      if (!isFinite(y)) return 誤('#NUM!');
      if (y === 0) break;
      var h = Math.abs(x) * 1e-7 + 1e-9;
      var d = (f(x + h) - f(x - h)) / (2 * h);
      if (!isFinite(d) || d === 0) return 誤('#NUM!');
      var 次 = x - y / d;
      if (!isFinite(次)) return 誤('#NUM!');
      if (次 <= -1) 次 = (x - 1) / 2;
      if (Math.abs(次 - x) < 1e-15 * Math.max(1, Math.abs(次))) { x = 次; break; }
      x = 次;
    }
    if (Math.abs(f(x)) > 1e-6) return 誤('#NUM!');
    return 数(x);
  }

  /** ★IPMT／PPMT★ … ★その 期の 利息／元金★
        ★利息 ＝ その 期の 始めの 残りに 利を 掛ける★ */
  function 期の内訳(引数たち, 手, 所, 利息か) {
    var g = 金の引数(引数たち, 手, 所, 4, 6);
    if (g.誤) return g.誤;
    var r = g.数[0], 期 = 切り捨て(g.数[1]), n = g.数[2], pv = g.数[3];
    var fv = g.数.length > 4 ? g.数[4] : 0;
    var 期首 = g.数.length > 5 && g.数[5] !== 0;
    if (期 < 1 || 期 > n) return 誤('#NUM!');
    var 返 = 表.PMT([{ 種: '直', 値: 数(r) }, { 種: '直', 値: 数(n) },
      { 種: '直', 値: 数(pv) }, { 種: '直', 値: 数(fv) },
      { 種: '直', 値: 数(期首 ? 1 : 0) }], 手, 所);
    if (返.型 === '誤') return 返;
    var pmt = 返.値;
    /* ★その 期の 始めの 残り★ ＝ 期-1 回 払った 後の 値 */
    var 残;
    if (r === 0) 残 = pv + pmt * (期 - 1);
    else {
      var k = Math.pow(1 + r, 期 - 1);
      残 = pv * k + pmt * (1 + (期首 ? r : 0)) * (k - 1) / r;
    }
    /* ★★期首払いの 利息は (1+r) で 割ります★★（2026-09-16 実測）
         `=IPMT(1,2,3,4,5,6)`（型 6＝期首）… 実Excel ★-1.3571428571428577★
         ＝期末の ★-2.714…★ の ちょうど 半分（1+r ＝ 2）
       ★1回目の 期首は 利息 0★（まだ 1日も 経って いない） */
    var 利 = (期 === 1 && 期首) ? 0 : (-残 * r) / (期首 ? (1 + r) : 1);
    return 利息か ? 数(利) : 数(pmt - 利);
  }
  /** ★不完全ガンマ（上側）★ … 連分数と 級数 */
  function 上のガンマ(a, x) {
    if (x <= 0) return 1;
    if (x < a + 1) {
      /* ★下側を 級数で 出して 1 から 引く★ */
      var 和 = 1 / a, 項 = 和, ap = a;
      for (var i = 0; i < 500; i++) {
        ap++;
        項 *= x / ap;
        和 += 項;
        if (Math.abs(項) < Math.abs(和) * 1e-16) break;
      }
      return 1 - 和 * Math.exp(-x + a * Math.log(x) - ガンマの対数(a));
    }
    var b = x + 1 - a, c = 1e300, d = 1 / b, h = d;
    for (var j = 1; j < 500; j++) {
      var an = -j * (j - a);
      b += 2;
      d = an * d + b;
      if (Math.abs(d) < 1e-300) d = 1e-300;
      c = b + an / c;
      if (Math.abs(c) < 1e-300) c = 1e-300;
      d = 1 / d;
      var 差 = d * c;
      h *= 差;
      if (Math.abs(差 - 1) < 1e-16) break;
    }
    return h * Math.exp(-x + a * Math.log(x) - ガンマの対数(a));
  }
  /** ★χ² の 右側★ */
  function カイ二乗の右(引数たち, 手, 所) {
    var g = 金の引数(引数たち, 手, 所, 2, 2);
    if (g.誤) return g.誤;
    var x = g.数[0], v = 切り捨て(g.数[1]);
    if (v < 1 || x < 0) return 誤('#NUM!');
    return 数(上のガンマ(v / 2, x / 2));
  }
  /** ★t の 右側★ … 不完全ベータを 連分数で */
  function tの右(t, v) {
    if (t < 0) return 1 - tの右(-t, v);
    var x = v / (v + t * t);
    return ベータの下(v / 2, 0.5, x) / 2;
  }
  /** ★不完全ベータ（下側・規格化）★ */
  function ベータの下(a, b, x) {
    if (x <= 0) return 0;
    if (x >= 1) return 1;
    var 頭 = Math.exp(ガンマの対数(a + b) - ガンマの対数(a) - ガンマの対数(b)
      + a * Math.log(x) + b * Math.log(1 - x));
    if (x < (a + 1) / (a + b + 2)) return 頭 * ベータ連分(a, b, x) / a;
    return 1 - 頭 * ベータ連分(b, a, 1 - x) / b;
  }
  function ベータ連分(a, b, x) {
    var qab = a + b, qap = a + 1, qam = a - 1;
    var c = 1, d = 1 - qab * x / qap;
    if (Math.abs(d) < 1e-300) d = 1e-300;
    d = 1 / d;
    var h = d;
    for (var m = 1; m <= 300; m++) {
      var m2 = 2 * m;
      var aa = m * (b - m) * x / ((qam + m2) * (a + m2));
      d = 1 + aa * d;
      if (Math.abs(d) < 1e-300) d = 1e-300;
      c = 1 + aa / c;
      if (Math.abs(c) < 1e-300) c = 1e-300;
      d = 1 / d;
      h *= d * c;
      aa = -(a + m) * (qab + m) * x / ((a + m2) * (qap + m2));
      d = 1 + aa * d;
      if (Math.abs(d) < 1e-300) d = 1e-300;
      c = 1 + aa / c;
      if (Math.abs(c) < 1e-300) c = 1e-300;
      d = 1 / d;
      var 差 = d * c;
      h *= 差;
      if (Math.abs(差 - 1) < 1e-16) break;
    }
    return h;
  }

  /** ★値を 真偽と 見る★（FILTER の 印） */
  function 真と見る(v) {
    if (!v) return false;
    if (v.型 === '真偽') return !!v.値;
    if (v.型 === '数') return v.値 !== 0;
    return false;
  }

  /** ★位の 一族の 共通★（★負は #NUM!★）
        `大きいまま` true なら 48ビットを 超えても 数で 計算する（ずらし） */
  function 位で(引数たち, 手, 所, 続き, 大きいまま) {
    if (引数たち.length < 2) return 誤('#VALUE!');
    var a = 数の引数(引数たち[0], 手, 所);
    if (a.誤) return a.誤;
    var b = 数の引数(引数たち[1], 手, 所);
    if (b.誤) return b.誤;
    var x = 切り捨て(a.数), y = 切り捨て(b.数);
    if (x < 0) return 誤('#NUM!');
    if (!大きいまま && y < 0) return 誤('#NUM!');
    var 上 = Math.pow(2, 48);
    if (x >= 上) return 誤('#NUM!');
    var 出;
    if (大きいまま) 出 = 続き(x, y);
    else {
      /* ★JS の ビット演算は 32ビットまで★＝★上と 下に 分けて 計算する★ */
      var 下 = 0x100000000;
      var 出下 = 続き(x % 下, y % 下) >>> 0;
      var 出上 = 続き(Math.floor(x / 下), Math.floor(y / 下)) >>> 0;
      出 = 出上 * 下 + 出下;
    }
    if (!isFinite(出) || 出 < 0 || 出 >= 上) return 誤('#NUM!');
    return 数(出);
  }
  /** ★一番 多く 出る 数★（MODE／MODE.SNGL） */
  function 最頑張り(引数たち, 手) {
    /* ★★数に 読めない 字は ★飛ばします★★★（2026-09-16 実測）
         `=MODE.SNGL(A1:A5,"<3",B1:B5)` … 実Excel ★2★ ／前は ★#VALUE!★
         ＝`数を拾う` は ★直に 書いた 字を 数に する★ので `"<3"` で 転ぶ
         ⇒★MODE は その 字を 数えて いません★ */
    var 拾 = { 数たち: [], 誤: null };
    ほどく(引数たち, function (v, 元) {
      if (v.型 === '誤') { 拾.誤 = v; return false; }
      if (元 === 'マス') { if (v.型 === '数') 拾.数たち.push(v.値); return true; }
      if (v.型 === '数') { 拾.数たち.push(v.値); return true; }
      if (v.型 === '真偽') { 拾.数たち.push(v.値 ? 1 : 0); return true; }
      if (v.型 === '字') {
        var n = 計.数にする(v, 手);
        if (n.型 !== '誤') 拾.数たち.push(n.値);
        return true;                           /* ★数に 読めなければ 飛ばす★ */
      }
      return true;
    });
    if (拾.誤) return 拾.誤;
    /* ★★同じ 数だけ 出る 物が 在る 時は ★先に 出た 方★★★（2026-09-16 実測）
         `=MODE.SNGL(A1:A5,B1:B5,1)` … 実Excel ★1★ ／前は ★2★
         ＝1 も 2 も 2回ずつ／★1 の 方が 並びの 先に 在る★
         ★前の 書き方★は「★先に 最多に 届いた 方★」を 採って いた
           ＝2 は 6番目で 2回目／1 は 11番目で 2回目 ⇒ 2 を 採って いた
         ⇒★★数え終わってから ★一番 先に 出た 物★を 選ぶ★★ */
    var a = 拾.数たち, 数え = {}, 初め = {}, i;
    for (i = 0; i < a.length; i++) {
      var k = String(a[i]);
      数え[k] = (数え[k] || 0) + 1;
      if (初め[k] === undefined) 初め[k] = i;
    }
    var 最 = 0, 出 = null, 出の場所 = Infinity;
    for (i = 0; i < a.length; i++) {
      var k2 = String(a[i]);
      if (数え[k2] > 最 || (数え[k2] === 最 && 初め[k2] < 出の場所)) {
        最 = 数え[k2]; 出 = a[i]; 出の場所 = 初め[k2];
      }
    }
    if (最 < 2) return 誤('#N/A');
    return 数(出);
  }
  /** ★ガンマ関数の 対数★（Lanczos） */
  function ガンマの対数(x) {
    /* ★★整数の 所は ★きっちり 0★★★（2026-09-16 実測）
         `=GAMMALN(2)` … Lanczos だと ★-4.44e-16★ ／実Excel ★0★
         ＝`GAMMALN(n)` は `ln((n-1)!)`＝★整数なら 足し算で きっちり 出ます★
         ⇒★近い 形で 出した 端数を 出さない★ */
    if (x === Math.floor(x) && x >= 1 && x <= 170) {
      var 和 = 0;
      for (var k = 2; k < x; k++) 和 += Math.log(k);
      return 和;
    }
    /* ★★使って いた Lanczos（g=5、6項）は 半端な 所で 浅い★★（2026-09-16 実測）
         `GAMMALN(0.5)`（＝ln√π）… ★ずれ 4.30e-14★ ／ `GAMMALN(0.25)` … ★ 6.68e-14★
       ⇒★★この ずれが そのまま χ²／t／F／ベータに 乗る★★
         ＝「12桁目から 違う」の 正体の 一つでした
       ⇒★g=7、10項の 係数に 差し替え★（ずれ ★約 1e-15★＝40倍 深い）
         ★小さい 方は 反射の 式で 折り返す★（π/sin） */
    if (x < 0.5) return Math.log(Math.PI / Math.abs(Math.sin(Math.PI * x))) - ガンマの対数(1 - x);
    var P7 = [0.99999999999980993, 676.5203681218851, -1259.1392167224028,
      771.32342877765313, -176.61502916214059, 12.507343278686905,
      -0.13857109526572012, 9.9843695780195716e-6, 1.5056327351493116e-7];
    var z = x - 1, a2 = P7[0], t2 = z + 7.5;
    for (var j = 1; j < 9; j++) a2 += P7[j] / (z + j);
    return 0.5 * Math.log(2 * Math.PI) + (z + 0.5) * Math.log(t2) - t2 + Math.log(a2);
  }

  /** ★ROW／COLUMN★（`行か` true＝ROW） */
  function 番地の番号(引数たち, 所, 行か) {
    if (!引数たち.length) {
      var 今 = 所 && 所.今のマス;
      if (!今) return 誤('#VALUE!');
      return 数((行か ? 今.行 : 今.列) + 1);
    }
    var a = 引数たち[0];
    if (!a || !a.番地) return 誤('#VALUE!');
    var 頭 = (行か ? a.番地.行 : a.番地.列) + 1;
    /* ★★四角を 渡されたら ★形の ぶんだけ★ 溢れます★★（2026-09-21 実測で 直した）
         ★★ここは ★マスごと★ では ありません★★
           `=ROW(A1:A3)`    ⇒ ★3マス★（1 / 2 / 3）＝★縦に 行数 ぶん★
           `=COLUMN(A1:A3)` ⇒ ★1マス★（1）　　　＝★横に 列数 ぶん（1列 なので 1つ）★
         ⇒★マスごとに 呼ぶと COLUMN が `1|1|1` の 3マスに なり ★間違い★★
         ★物差し★ `docs/measured/golden-jitsu-excel-hanni-zenbu-2026-09-20.tsv`
           65行目 `=COLUMN(A1:A3)` 1マス ／ 117行目 `=ROW(A1:A3)` 3マス
         ★★測って いない 事★★
           ★横に 2列 以上の 範囲を COLUMN に 渡した 時★（`=COLUMN(E1:F2)`）は
           ★まだ 実Excel で 測って いません★＝上の 決まりから 出した 形です
           （★両方の 測った 値は この 決まりで 出ます★） */
    if (a.種 === '四角') {
      var 幅 = 行か ? (a.行数 || 1) : (a.列数 || 1);
      if (幅 > 1) {
        var 並 = [];
        for (var i = 0; i < 幅; i++) 並.push(数(頭 + i));
        if (行か) {
          var 縦 = [];
          for (var j = 0; j < 並.length; j++) 縦.push([並[j]]);
          return { 溢れ: true, 行数: 幅, 列数: 1, 並び: 縦 };
        }
        return { 溢れ: true, 行数: 1, 列数: 幅, 並び: [並] };
      }
    }
    return 数(頭);
  }

  /** ★PERCENTILE／QUARTILE の 共通★（`四分か` true なら 0〜4 を 0〜1 に） */
  function 分位(引数たち, 手, 所, 四分か) {
    if (引数たち.length < 2) return 誤('#VALUE!');
    var 拾 = 数を拾う([引数たち[0]], 手);
    if (拾.誤) return 拾.誤;
    var t = 数の引数(引数たち[1], 手, 所);
    if (t.誤) return t.誤;
    var a = 拾.数たち.slice().sort(function (x, y) { return x - y; });
    if (!a.length) return 誤('#NUM!');
    var p = 四分か ? (切り捨て(t.数) / 4) : t.数;
    if (四分か && (切り捨て(t.数) < 0 || 切り捨て(t.数) > 4)) return 誤('#NUM!');
    if (p < 0 || p > 1) return 誤('#NUM!');
    if (a.length === 1) return 数(a[0]);
    var i = p * (a.length - 1);
    var 下 = Math.floor(i), 余 = i - 下;
    if (下 + 1 >= a.length) return 数(a[a.length - 1]);
    return 数(a[下] + (a[下 + 1] - a[下]) * 余);
  }
  /** ★休みの 日付を 拾う★（NETWORKDAYS／WORKDAY の 3つ目以降） */
  function 休みを拾う(引数たち, 頭, 手) {
    if (引数たち.length <= 頭) return { 日: [] };
    var 拾 = 数を拾う(引数たち.slice(頭), 手);
    if (拾.誤) return { 誤: 拾.誤 };
    var 出 = [];
    for (var i = 0; i < 拾.数たち.length; i++) 出.push(切り捨て(拾.数たち[i]));
    return { 日: 出 };
  }

  /** ★うるう年か★（★本当の 暦★＝Excel の 通し番号の 1900-02-29 とは 別）
        ★物差し★ `golden-yearfrac-basis1-2026-09-16.tsv`
          `=YEARFRAC(DATE(1900,1,1),DATE(1900,12,31),1)` … ★日数 365 なのに 分母 365★
            ＝★Excel は 1900 を うるう年と 見て いません★（通し番号だけ ずれて いる）
          `=YEARFRAC(DATE(2000,1,1),DATE(2000,12,31),1)` … 分母 ★366★（400で 割れる）
          `=YEARFRAC(DATE(2100,1,1),DATE(2100,12,31),1)` … 分母 ★365★（100で 割れる） */
  function うるう年か(y) {
    return (y % 4 === 0) && ((y % 100 !== 0) || (y % 400 === 0));
  }

  /** ★年月日 → 通し番号★（★1900年の 穴を 含む★・`書.数から日()` の 逆）
        ★実Excel には 1900-02-29 という 無い 日が 在ります★（通し 60）
        ⇒★通し 61 未満は 1日 ずらす★ */
  function 日から通し(y, m, d) {
    var 通 = Math.round((Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000);
    if (通 < 61 && 通 > 0) 通 -= 1;
    return 通;
  }

  /** ★日付の 字を 通し番号に★（DATEVALUE）… 読めなければ null
        ★物差し★ `golden-1900-hidzuke-2026-09-15.tsv`
          `=DATEVALUE("1900-01-01")` ★　1★ ／ `=DATEVALUE("1900-02-29")` ★　60★
          `=DATEVALUE("2024-01-01")` ★　45292★ ／ `=DATEVALUE("0000-01-01")` ★#VALUE!★
        ★年が 抜けた 形（"1-2"）は 今年で 読みますが ★ここでは 読みません★
          ＝★今日に よって 答えが 変わる 物は 紙で 押せない★（未測定と 同じ 扱い） */
  function 日付の字を読む(t) {
    var m = /^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})$/.exec(String(t).trim());
    if (!m) return null;
    var y = +m[1], mo = +m[2], da = +m[3];
    if (y < 1900 || y > 9999) return null;
    if (mo < 1 || mo > 12 || da < 1 || da > 31) return null;
    if (y === 1900 && mo === 2 && da === 29) return 60;   /* ★実Excelに 在る 無い 日★ */
    var v = new Date(Date.UTC(y, mo - 1, da));
    if (v.getUTCMonth() !== mo - 1 || v.getUTCDate() !== da) return null;
    return 日から通し(y, mo, da);
  }
  /** ★時刻の 字を 日の 端数に★（TIMEVALUE）… 読めなければ null
        ★★直に 書いた 字（AVERAGE 等）とは 別の 道です★★（2026-09-15 実測）
          `=AVERAGE("24:00")`   ★　1★      ＝日に 繰り上げる
          `=TIMEVALUE("24:00")` ★　0★      ＝★日の 分を 切り捨てる★
          `=TIMEVALUE("25:00")` ★0.0416…★ （＝1時間）
        ⇒★同じ 字でも 読む 道で 答えが 違います★ */
  function 時刻の字を読む(t) {
    var m = /^(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?$/.exec(String(t).trim());
    if (!m) return null;
    var h = +m[1], mi = +m[2], se = m[3] ? +m[3] : 0;
    if (mi > 59 || se > 59) return null;
    var x = (h * 3600 + mi * 60 + se) / 86400;
    return x - Math.floor(x);                 /* ★日の 分は 捨てる★ */
  }
  /** ★EDATE／EOMONTH★（`月末か` true＝その 月の 終わり） */
  function 月をずらす(引数たち, 手, 所, 月末か) {
    var g = 金の引数(引数たち, 手, 所, 2, 2);
    if (g.誤) return g.誤;
    if (g.数[0] < 0) return 誤('#NUM!');
    var o = 書.数から日(切り捨て(g.数[0]));
    var ず = 切り捨て(g.数[1]);
    var y = o.y, m = o.m + ず;
    y += Math.floor((m - 1) / 12);
    m = ((m - 1) % 12 + 12) % 12 + 1;
    var 末 = new Date(Date.UTC(y, m, 0)).getUTCDate();
    var d = 月末か ? 末 : Math.min(o.d, 末);
    /* ★1899-12-30 を 0 と する 数え方に 戻す★（★1900年の 穴は 書式の 台が 持つ★） */
    var 通 = Math.round((Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000);
    if (通 < 61 && 通 > 0) 通 -= 1;                 /* ★1900-02-29 の 分★ */
    return 数(通);
  }

  /** ★IRR／XIRR の 共通＝★探して 近づける★★
        `日` が null なら 期ごと（IRR）／並びなら 日づけ（XIRR）
        ★実Excel と 同じ 桁までは 揃いません★（★探し方が 違えば 末桁が 変わる★）
        ⇒★★合わない 行は 名指しで 棚に 出します★★ */
  function 探して近づける(金たち, 日たち, 見当) {
    if (!金たち.length) return 誤('#NUM!');
    var 負 = false, 正 = false, i;
    for (i = 0; i < 金たち.length; i++) {
      if (金たち[i] < 0) 負 = true;
      else if (金たち[i] > 0) 正 = true;
    }
    if (!負 || !正) return 誤('#NUM!');
    var 頭 = 日たち ? 日たち[0] : 0;
    var 値 = function (r) {
      var 和 = 0;
      for (var k = 0; k < 金たち.length; k++) {
        var t = 日たち ? (日たち[k] - 頭) / 365 : k;
        和 += 金たち[k] / Math.pow(1 + r, t);
      }
      return 和;
    };
    var r0 = 見当;
    if (!(r0 > -1)) r0 = 0.1;
    for (i = 0; i < 100; i++) {
      var f0 = 値(r0);
      var 刻 = Math.max(1e-7, Math.abs(r0) * 1e-7);
      var f1 = 値(r0 + 刻);
      var 傾 = (f1 - f0) / 刻;
      if (!isFinite(傾) || 傾 === 0) break;
      var 次 = r0 - f0 / 傾;
      if (!isFinite(次)) break;
      if (次 <= -1) 次 = (r0 - 1) / 2;
      if (Math.abs(次 - r0) < 1e-12) { r0 = 次; break; }
      r0 = 次;
    }
    if (!isFinite(r0) || Math.abs(値(r0)) > 1e-6) return 誤('#NUM!');
    return 数(r0);
  }

  /** ★引数を 2次元の 表に する★（1マス・直は 1×1） */
  function 表にする(引数, 所) {
    if (!引数) return null;
    if (引数.種 === '四角') {
      var 行 = 引数.行数 || 0, 列 = 引数.列数 || 0, 並 = 引数.並び || [];
      var 出 = [];
      for (var r = 0; r < 行; r++) {
        var 段 = [];
        for (var c = 0; c < 列; c++) 段.push(並[r * 列 + c] || 計.空);
        出.push(段);
      }
      return 出;
    }
    var v = 場.ひとつに(引数, 所 && 所.今のマス);
    return [[v || 計.空]];
  }
  /** ★2次元の 表を 溢れの 形に する★（`shiki-afure.js` と 同じ 形） */
  /** ★直線で 当てる★（FORECAST／FORECAST.LINEAR が 使う）
        ★y ＝ a ＋ b x★ を 最小二乗で 出し、x に 当てる
        ★対に ならない 所（どちらかが 数で ない）は ★両方 飛ばす★★
        ★★測ったのは 1行だけ★★ ... `=FORECAST(6,B1:B5,A1:A5)` → 12
          ⇒★端（長さ違い・ばらつき 0・空マス）は ★未測定★★ */
  function 直線で当てる(引数たち, 手, 所) {
    if (引数たち.length < 3) return 誤('#VALUE!');
    var x = 数の引数(引数たち[0], 手, 所);
    if (x.誤) return x.誤;
    var Y = 表にする(引数たち[1], 所), X = 表にする(引数たち[2], 所);
    if (!Y || !X) return 誤('#VALUE!');
    var ys = [], xs = [], r, c;
    for (r = 0; r < Y.length; r++) for (c = 0; c < Y[r].length; c++) ys.push(Y[r][c]);
    for (r = 0; r < X.length; r++) for (c = 0; c < X[r].length; c++) xs.push(X[r][c]);
    if (ys.length !== xs.length) return 誤('#N/A');
    var a = [], b = [], i;
    for (i = 0; i < ys.length; i++) {
      if (ys[i] && ys[i].型 === '誤') return ys[i];
      if (xs[i] && xs[i].型 === '誤') return xs[i];
      if (!ys[i] || ys[i].型 !== '数' || !xs[i] || xs[i].型 !== '数') continue;
      a.push(ys[i].値); b.push(xs[i].値);
    }
    var n = a.length;
    if (!n) return 誤('#N/A');
    var 和x = 0, 和y = 0;
    for (i = 0; i < n; i++) { 和x += b[i]; 和y += a[i]; }
    var 平x = 和x / n, 平y = 和y / n, 上 = 0, 下 = 0;
    for (i = 0; i < n; i++) { 上 += (b[i] - 平x) * (a[i] - 平y); 下 += (b[i] - 平x) * (b[i] - 平x); }
    if (下 === 0) return 誤('#DIV/0!');
    return 数(平y + (上 / 下) * (x.数 - 平x));
  }

  /** ★桁数で 書く★（FIXED／DOLLAR が 使う）
        ★桁が 正★ ... その 桁まで（★足りなければ 0 を 足す★）
        ★桁が 0 か 負★ ... ★左へ 丸めて 小数点を 出さない★
        ★区切り★ ... 3桁ごとに `,`
        ★丸め方★ ... ★四捨五入（0から 遠い 方へ）★＝実Excel の 丸め
        ★★未測定★★ ... ★桁が 100超★／★とても 大きい 数（指数に なるか）★ */
  function 桁数で書く(数値, 桁, 区切る) {
    if (!isFinite(数値)) return null;
    var 負 = 数値 < 0, x = Math.abs(数値);
    var s;
    if (桁 > 0) {
      s = 丸めて文字に(x, 桁);
    } else {
      var 単 = Math.pow(10, -桁);
      s = 丸めて文字に(x / 単, 0);
      s = String(Number(s) * 単);
      if (s.indexOf('.') >= 0) s = 丸めて文字に(Number(s), 0);
    }
    var 割 = s.split('.');
    if (区切る) 割[0] = 割[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    return (負 ? '-' : '') + 割.join(割.length > 1 ? '.' : '');
  }
  /** ★0から 遠い 方へ 四捨五入して 字に する★（★実Excel の 丸め★） */
  function 丸めて文字に(x, 桁) {
    var 単 = Math.pow(10, 桁);
    var y = Math.round(x * 単 + Number.EPSILON * x * 単) / 単;
    var s = y.toFixed(Math.max(0, 桁));
    return s;
  }
  /** ★2つの 値を 比べる★（LOOKUP が 使う）
        ★返し★ 正＝左が 大きい ／ 0＝同じ ／ 負＝左が 小さい ／ null＝比べられない
        ★★型が 違えば 比べません★★（実Excel は 数 < 字 < 真偽 の 順ですが ★未測定★） */
  function 比べられるか(a, b) {
    if (!a || !b) return null;
    if (a.型 === '数' && b.型 === '数') return a.値 - b.値;
    if (a.型 === '字' && b.型 === '字') {
      var x = String(a.値).toUpperCase(), y = String(b.値).toUpperCase();
      return x < y ? -1 : (x > y ? 1 : 0);
    }
    if (a.型 === '真偽' && b.型 === '真偽') return (a.値 ? 1 : 0) - (b.値 ? 1 : 0);
    return null;
  }

  function 溢れに(表) {
    return { 溢れ: true, 行数: 表.length, 列数: 表.length ? 表[0].length : 0, 並び: 表 };
  }
  /* ══ ★★ラムダの 一族の 台★★ ══（2026-09-18・段取り ⑶⑷）
       ★ラムダを 呼ぶ 口は 板が 持って います★（`所.ラムダを呼ぶ`）
         ＝★ここ（`shiki-kansuu.js`）は 木を 知りません★
       ★★決まりは 全部 実測です★★（`docs/measured/golden-kansuu-9kaime-2026-09-18.tsv`）
         MAP ...... 元と 同じ 形（A1:A5 -> 5行1列）
         SCAN ..... ★元と 同じ 形★＝★初めの 値は 含まれません★（5行・35）
         REDUCE ... 1つの 値（a が 溜め・b が 次＝`REDUCE(100,..,a-b)` が 85 で 裏取り）
         BYROW .... ★行ごと★（F1:G2 -> 3 と 30）＝★縦 1列★
         BYCOL .... ★列ごと★（F1:G2 -> 11 と 22）＝★横 1行★
         MAKEARRAY ★行・列は 1から★（`MAKEARRAY(2,3,LAMBDA(r,c,r*c))` の 和が 18）
         ★引数の 数違いは #VALUE!★（55／56行目） */
  /** ★行ごと／列ごとに 当てる★（BYROW／BYCOL が 使う）
        ★返しの 形★ ... BYROW は ★縦 1列★／BYCOL は ★横 1行★
        ★ラムダに 渡す 物★ ... ★その 行（列）を 溢れの 形で★
          ＝`SUM(r)` が 受け取れる 形（★四角と 同じ 顔★） */
  function 段ごとに当てる(引数たち, 所, 行か) {
    if (引数たち.length < 2) return 誤('#VALUE!');
    var g = ラムダの引数(引数たち[1], 所);
    if (g.誤) return g.誤;
    var A = 表にする(引数たち[0], 所);
    if (!A || !A.length) return 誤('#VALUE!');
    var 出 = [];
    if (行か) {
      for (var r = 0; r < A.length; r++) {
        var v = ラムダ1回(所, g.L, [溢れに([A[r]])]);
        if (v && v.型 === '誤') return v;
        出.push([v]);
      }
    } else {
      var 段 = [];
      for (var c = 0; c < A[0].length; c++) {
        var 縦 = [];
        for (var r2 = 0; r2 < A.length; r2++) 縦.push([A[r2][c] || 計.空]);
        var w = ラムダ1回(所, g.L, [溢れに(縦)]);
        if (w && w.型 === '誤') return w;
        段.push(w);
      }
      出.push(段);
    }
    return 溢れに(出);
  }

  function ラムダの引数(引数, 所) {
    if (!所 || typeof 所.ラムダを呼ぶ !== 'function') return { 誤: 誤('#VALUE!') };
    var v = 場.ひとつに(引数, 所 && 所.今のマス);
    if (v && v.型 === '誤') return { 誤: v };
    if (!所.ラムダか(v)) return { 誤: 誤('#VALUE!') };
    return { L: v };
  }
  /** ★ラムダを 1回 呼ぶ★（★誤りは そのまま 返す★） */
  function ラムダ1回(所, L, 引値) {
    var 出 = 所.ラムダを呼ぶ(L, 引値);
    return 出 || 計.空;
  }

  /** ★VSTACK／HSTACK★（★足りない 所は `#N/A`★＝実Excel と 同じ） */
  function 積む(引数たち, 所, 横か) {
    if (!引数たち.length) return 誤('#VALUE!');
    var 表たち = [], i;
    for (i = 0; i < 引数たち.length; i++) {
      var A = 表にする(引数たち[i], 所);
      if (!A || !A.length) return 誤('#VALUE!');
      表たち.push(A);
    }
    var 出 = [];
    if (横か) {
      var 高 = 0;
      for (i = 0; i < 表たち.length; i++) 高 = Math.max(高, 表たち[i].length);
      for (var r = 0; r < 高; r++) {
        var 段 = [];
        for (i = 0; i < 表たち.length; i++) {
          var A2 = 表たち[i], 幅 = A2[0].length;
          for (var c = 0; c < 幅; c++) 段.push(r < A2.length ? A2[r][c] : 計.誤('#N/A'));
        }
        出.push(段);
      }
      return 溢れに(出);
    }
    var 幅2 = 0;
    for (i = 0; i < 表たち.length; i++) 幅2 = Math.max(幅2, 表たち[i][0].length);
    for (i = 0; i < 表たち.length; i++) {
      var A3 = 表たち[i];
      for (var r2 = 0; r2 < A3.length; r2++) {
        var 段2 = [];
        for (var c2 = 0; c2 < 幅2; c2++) 段2.push(c2 < A3[r2].length ? A3[r2][c2] : 計.誤('#N/A'));
        出.push(段2);
      }
    }
    return 溢れに(出);
  }

  /** ★階乗★（MULTINOMIAL の 部品） */
  function 階乗(n) {
    var r = 1;
    for (var i = 2; i <= n; i++) r *= i;
    return r;
  }
  /** ★`.PRECISE` の 丸め★ … ★刻みの 符号を 見ない★（負でも 上へ／下へ） */
  function きっちり丸め(引数たち, 手, 所, 上か) {
    if (!引数たち.length) return 誤('#VALUE!');
    var t = 数の引数(引数たち[0], 手, 所);
    if (t.誤) return t.誤;
    var k = 1;
    if (引数たち.length > 1) {
      var m = 数の引数(引数たち[1], 手, 所);
      if (m.誤) return m.誤;
      k = Math.abs(m.数);
    }
    if (k === 0) return 数(0);
    return 数((上か ? Math.ceil(t.数 / k) : Math.floor(t.数 / k)) * k);
  }
  /** ★歪み★（SKEW／SKEW.P） */
  function 歪み(引数たち, 手, 母集団か) {
    var 拾 = 数を拾う(引数たち, 手);
    if (拾.誤) return 拾.誤;
    var a = 拾.数たち, n = a.length, i;
    if (n < (母集団か ? 1 : 3)) return 誤('#DIV/0!');
    var 和 = 0;
    for (i = 0; i < n; i++) 和 += a[i];
    var 平 = 和 / n, 二 = 0, 三 = 0;
    for (i = 0; i < n; i++) {
      var d = a[i] - 平;
      二 += d * d;
      三 += d * d * d;
    }
    if (二 === 0) return 誤('#DIV/0!');
    if (母集団か) {
      var σp = Math.sqrt(二 / n);
      return 数((三 / n) / (σp * σp * σp));
    }
    var σ = Math.sqrt(二 / (n - 1));
    return 数(n / ((n - 1) * (n - 2)) * (三 / (σ * σ * σ)));
  }

  /** ★四角を 平らな 並びに する★（無ければ null） */
  function 四角の並び(引数) {
    if (!引数) return null;
    if (引数.種 === '四角') return { 並び: 引数.並び || [], 行数: 引数.行数 || 0, 列数: 引数.列数 || 0 };
    /* ★★1マスも 1マスの 四角として 受けます★★（2026-09-16 実測）
         `=SUMIF(D1,D2)` → 実Excel ★0★／前は ★#VALUE!★
         ＝★実Excel は 1マスを 範囲として 受けます★
         ⇒ SUMIFS MAXIFS MINIFS も 同じ（6本 まとめて 赤に なった） */
    if (引数.種 === 'マス') return { 並び: [引数.値], 行数: 1, 列数: 1 };
    return null;
  }
  /** ★COUNTIF／SUMIF／AVERAGEIF の 共通★
        `合計の場所` … null＝数えるだけ／番号＝その 引数を 足す（無ければ 探す 四角そのもの） */
  function 条件で選ぶ(引数たち, 手, 所, 少なくとも, 合計の場所) {
    if (引数たち.length < 少なくとも) return { 誤: 誤('#VALUE!') };
    var 探 = 四角の並び(引数たち[0]);
    if (!探) return { 誤: 誤('#VALUE!') };
    var 条 = 場.ひとつに(引数たち[1], 所 && 所.今のマス);
    if (!条) return { 誤: 誤('#VALUE!') };
    if (条.型 === '誤') return { 誤: 条 };
    var 足す = 探;
    if (合計の場所 !== null && 引数たち.length > 合計の場所) {
      var f = 四角の並び(引数たち[合計の場所]);
      if (f) 足す = f;
    }
    var 当たり = [], 値 = [];
    for (var i = 0; i < 探.並び.length; i++) {
      if (!当てはまるか(条, 探.並び[i])) continue;
      当たり.push(i);
      値.push(足す.並び[i]);
    }
    return { 当たり: 当たり, 値: 値 };
  }
  /** ★COUNTIFS／SUMIFS の 共通＝★条件を 重ねる（かつ）★
        `頭` … 0＝(範囲,条件)… ／1＝(足す範囲, 範囲,条件…) */
  function 条件を重ねる(引数たち, 手, 所, 頭) {
    if (引数たち.length < 頭 + 2 || (引数たち.length - 頭) % 2) return { 誤: 誤('#VALUE!') };
    var 印 = null;
    for (var i = 頭; i + 1 < 引数たち.length; i += 2) {
      var 範 = 四角の並び(引数たち[i]);
      if (!範) return { 誤: 誤('#VALUE!') };
      var 条 = 場.ひとつに(引数たち[i + 1], 所 && 所.今のマス);
      if (!条) return { 誤: 誤('#VALUE!') };
      if (条.型 === '誤') return { 誤: 条 };
      if (印 === null) { 印 = []; for (var k = 0; k < 範.並び.length; k++) 印.push(true); }
      if (範.並び.length !== 印.length) return { 誤: 誤('#VALUE!') };
      for (var j = 0; j < 範.並び.length; j++) {
        if (印[j] && !当てはまるか(条, 範.並び[j])) 印[j] = false;
      }
    }
    return { 印: 印 || [] };
  }
  /** ★SUMIFS／MAXIFS／MINIFS★（★最初の 引数が 足す 範囲★） */
  function 重ねて集める(引数たち, 手, 所, 種) {
    var 足す = 四角の並び(引数たち[0]);
    if (!足す) return 誤('#VALUE!');
    var g = 条件を重ねる(引数たち, 手, 所, 1);
    if (g.誤) return g.誤;
    var 数たち = [];
    for (var i = 0; i < g.印.length; i++) {
      if (!g.印[i]) continue;
      var v = 足す.並び[i];
      if (v && v.型 === '数') 数たち.push(v.値);
    }
    if (種 === 'SUM') {
      var 和 = 0;
      for (var j = 0; j < 数たち.length; j++) 和 += 数たち[j];
      return 数(和);
    }
    if (!数たち.length) return 数(0);          /* ★MAXIFS／MINIFS は 0件で 0★ */
    return 数(種 === 'MAX' ? Math.max.apply(null, 数たち) : Math.min.apply(null, 数たち));
  }
  /** ★LARGE／SMALL★ */
  function 大小のk番(引数たち, 手, 所, 大きい方) {
    if (引数たち.length < 2) return 誤('#VALUE!');
    var 拾 = 数を拾う([引数たち[0]], 手);
    if (拾.誤) return 拾.誤;
    var t = 数の引数(引数たち[1], 手, 所);
    if (t.誤) return t.誤;
    var k = 切り捨て(t.数);
    var a = 拾.数たち.slice().sort(function (x, y) { return 大きい方 ? y - x : x - y; });
    if (k < 1 || k > a.length) return 誤('#NUM!');
    return 数(a[k - 1]);
  }
  /** ★VLOOKUP／HLOOKUP★（★既定は 近い 方（TRUE）★） */
  function 縦横で探す(引数たち, 手, 所, 縦か) {
    if (引数たち.length < 3) return 誤('#VALUE!');
    var 今 = 所 && 所.今のマス;
    var 探 = 場.ひとつに(引数たち[0], 今);
    if (!探) return 誤('#VALUE!');
    if (探.型 === '誤') return 探;
    var 表 = 四角の並び(引数たち[1]);
    if (!表) return 誤('#VALUE!');
    var t = 数の引数(引数たち[2], 手, 所);
    if (t.誤) return t.誤;
    var 番 = 切り捨て(t.数);
    var 近い = true;
    if (引数たち.length > 3) {
      var 判 = 真偽にする(場.ひとつに(引数たち[3], 今));
      if (判.誤) return 判.誤;
      近い = 判.真;
    }
    var 長 = 縦か ? 表.行数 : 表.列数;
    var 幅 = 縦か ? 表.列数 : 表.行数;
    if (番 < 1 || 番 > 幅) return 誤('#REF!');
    var 取る = function (i) { return 縦か ? 表.並び[i * 表.列数] : 表.並び[i]; };
    var 当 = -1;
    for (var i = 0; i < 長; i++) {
      var v = 取る(i);
      var d = 場.比べる(v, 探);
      if (d === 0) { 当 = i; break; }
      if (近い && d !== null && d < 0) 当 = i;      /* ★以下で 一番 後ろ★ */
    }
    if (当 < 0) return 誤('#N/A');
    var 出 = 縦か ? 表.並び[当 * 表.列数 + (番 - 1)] : 表.並び[(番 - 1) * 表.列数 + 当];
    return 出 || 計.空;
  }

  /** ★お金の 一族の 引数★ … ★暗黙の 交わり → 数★を 決まった 本数 取る
        （`少なくとも` 本 未満／`多くとも` 本 超過は #VALUE!） */
  function 金の引数(引数たち, 手, 所, 少なくとも, 多くとも) {
    if (引数たち.length < 少なくとも || 引数たち.length > 多くとも) return { 誤: 誤('#VALUE!') };
    var 出 = [];
    for (var i = 0; i < 引数たち.length; i++) {
      var t = 数の引数(引数たち[i], 手, 所);
      if (t.誤) return { 誤: t.誤 };
      出.push(t.数);
    }
    return { 数: 出 };
  }

  /** ★IS… の 共通★
        ★★IS… は ★どれも★ 誤りを 素通りさせません★★（2026-09-15 実測）
          ＝★誤りを「誤りとして」判じます★
          `=ISBLANK(A1:B5)` → ★FALSE★（★#VALUE! を 返すのでは ない★）
          `=ISNONTEXT(A1:B5)` → ★TRUE★ ／ `=ISERR(A1:B5)` → ★TRUE★
          ★2次元の 四角は 暗黙の 交わりが 出来ず 誤りに なる★＝その 誤りを 見る
        ⇒★他の 関数の「誤りは 素通り」とは ★逆の 決め★★ */
  function 見分け(引数たち, 所, 続き) {
    if (!引数たち.length) return 誤('#VALUE!');
    /* ★★範囲を 渡されたら 溢れる★★（2026-09-21）
         ★前は どう だったか★
           `場.ひとつに` で ★暗黙の 交わり★を して 1つの 値に して いました。
           ⇒ `=ISNUMBER(A1:A3)` が ★TRUE 1マス★
           ⇒ 実Excel（動く並び）は ★TRUE / TRUE / TRUE の 3マス★
           ⇒★★誤りに ならず 黙って 足りない＝お客さんは 気付けません★★
         ★実測★（2026-09-21・実Excel 16.0 build 20326）
           紙 `docs/measured/golden-jitsu-excel-hanni-zenbu-2026-09-20.tsv`
           ★実Excel が 溢れる 112本の うち 81本で うちの 答えが 足りない★
           ★その うち 72本は 誤りに ならない＝気付けない★
         ★方針★ ... 司さん 2026-08-18「Excel の 最上級・全部 つける」
           ＝★「出来ません」と 言わない 状態に する★
           ＝★静かに 違う 答えは まず「正しく 計算できないか」を 読む★
         ★注意★ ... ★１マスだけの 範囲（A1:A1）は 今まで通り★
           ＝溢れさせると ★今まで 動いて いた 物の 形が 変わる★ */
    /* ★★１次元の 範囲だけ★★（2026-09-21）
         ★２次元（A1:B5）は 今まで通り★＝暗黙の 交わり
         ★訳★ ... 紙 `docs/measured/kansuu46/golden-346-2026-09-08.tsv`（実Excel）
           `=ISERR(A1:B5)`    ⇒ ★True★（★２次元は 交わりが 出来ず 誤り★）
           `=ISERROR(A1:B5)`  ⇒ ★True★
           `=ISNUMBER(A1:B5)` ⇒ ★False★
         ★１次元は 溢れる★ ... 紙 `golden-jitsu-excel-hanni-zenbu-2026-09-20.tsv`
           `=ISNUMBER(A1:A3)` ⇒ ★TRUE / TRUE / TRUE★（3マス）
         ★まだ 測って いない★
           ★２次元を 動く並びで 打った 時★（古い 紙は 2026-09-08）
           ⇒★測るまで 今まで通りに します★（★動いて いる 物を 触らない★） */
    var a = 引数たち[0];
    var 一次元 = a && a.種 === '四角'
      && (((a.行数 || 0) > 1 && (a.列数 || 0) <= 1)
       || ((a.列数 || 0) > 1 && (a.行数 || 0) <= 1));
    if (一次元) {
      var 表 = 表にする(a, 所);
      if (表) {
        var 出 = [];
        for (var r = 0; r < 表.length; r++) {
          var 段 = [];
          for (var c = 0; c < 表[r].length; c++) 段.push(計.真偽(続き(表[r][c] || 計.空)));
          出.push(段);
        }
        return { 溢れ: true, 行数: 出.length, 列数: 出.length ? 出[0].length : 0, 並び: 出 };
      }
    }
    /* ★★範囲を 渡されたら 溢れる★★（2026-09-21）
         ★前は★ `場.ひとつに` で 暗黙の 交わりを して 1つに して いました
           `=ISNUMBER(A1:A3)` が ★TRUE 1マス★／実Excel は ★3マス★
           ★誤りに ならず 黙って 足りない＝お客さんは 気付けません★
         ★実測★（2026-09-21・実Excel 16.0 build 20326）
           1次元 `=ISNUMBER(A1:A5)` ⇒ ★5マス TRUE★
           2次元 `=ISNUMBER(A1:B5)` ⇒ ★10マス（5行 x 2列）TRUE★
           紙 `docs/measured/kansuu46/golden-346-nijigen-formula2-2026-09-21.tsv`
           ★古い 紙（2026-09-08）は `.Formula`（昔の 道＝溢れない）で 取って います
           ⇒★「True / True / False」は 昔の 道の 答え★（間違いでは 無く 足りない）
         ★方針★ 司さん 2026-08-18「Excel の 最上級・全部 つける」
         ★注意★ 1マスの 範囲（A1:A1）は 今まで通り */
    var a = 引数たち[0];
    var 範囲 = a && a.種 === '四角'
      && (((a.行数 || 0) > 1) || ((a.列数 || 0) > 1));
    if (範囲) {
      var 表 = 表にする(a, 所);
      if (表) {
        var 出 = [];
        for (var r = 0; r < 表.length; r++) {
          var 段 = [];
          for (var c = 0; c < 表[r].length; c++) 段.push(計.真偽(続き(表[r][c] || 計.空)));
          出.push(段);
        }
        return { 溢れ: true, 行数: 出.length, 列数: 出.length ? 出[0].length : 0, 並び: 出 };
      }
    }
    var v = 場.ひとつに(引数たち[0], 所 && 所.今のマス) || 計.誤('#VALUE!');
    return 計.真偽(続き(v));
  }

  /** ★字の 引数★ … 暗黙の 交わり → ★字に する★（`&` と 同じ 台を 使う） */
  function 字の引数(引数, 手, 所) {
    var v = 場.ひとつに(引数, 所 && 所.今のマス);
    if (!v) return { 誤: 誤('#VALUE!') };
    if (v.型 === '誤') return { 誤: v };
    var t = 計.字にする(v);
    if (t.型 === '誤') return { 誤: t };
    return { 字: String(t.値) };
  }
  /** ★字を 1つ 取って 1つ 返す★ */
  function 字ひとつ(引数たち, 手, 所, 続き) {
    if (!引数たち.length) return 誤('#VALUE!');
    var 一マス = function (v) {
      if (!v) return 誤('#VALUE!');
      if (v.型 === '誤') return v;
      var t2 = 計.字にする(v);
      if (t2.型 === '誤') return t2;
      return 計.字(続き(String(t2.値)));
    };
    if (範囲か(引数たち[0])) {
      var 溢 = マスごとに溢れさせる(引数たち[0], 所, 一マス);
      if (溢) return 溢;
    }
    var t = 字の引数(引数たち[0], 手, 所);
    if (t.誤) return t.誤;
    return 計.字(続き(t.字));
  }
  /** ★LEFT／RIGHT★（`長さ` を 省くと 1） */
  function 端を取る(引数たち, 手, 所, 左か) {
    if (!引数たち.length) return 誤('#VALUE!');
    var t = 字の引数(引数たち[0], 手, 所);
    if (t.誤) return t.誤;
    var n = 1;
    if (引数たち.length > 1) {
      var d = 数の引数(引数たち[1], 手, 所);
      if (d.誤) return d.誤;
      n = 切り捨て(d.数);
      if (n < 0) return 誤('#VALUE!');
    }
    return 計.字(左か ? t.字.slice(0, n) : (n === 0 ? '' : t.字.slice(-n)));
  }
  /** ★FIND／SEARCH★（FIND は 大小を 見る／SEARCH は 見ない・見つからなければ #VALUE!） */
  function 探す(引数たち, 手, 所, 大小か) {
    if (引数たち.length < 2) return 誤('#VALUE!');
    var a = 字の引数(引数たち[0], 手, 所);
    if (a.誤) return a.誤;
    var b = 字の引数(引数たち[1], 手, 所);
    if (b.誤) return b.誤;
    var 頭 = 1;
    if (引数たち.length > 2) {
      var d = 数の引数(引数たち[2], 手, 所);
      if (d.誤) return d.誤;
      頭 = 切り捨て(d.数);
      if (頭 < 1) return 誤('#VALUE!');
    }
    var 的 = 大小か ? b.字 : b.字.toLowerCase();
    var 探 = 大小か ? a.字 : a.字.toLowerCase();
    var i = 的.indexOf(探, 頭 - 1);
    if (i < 0) return 誤('#VALUE!');
    return 数(i + 1);
  }

  /** ★丸めの 共通★（ROUND／ROUNDUP／ROUNDDOWN）
        ★半分は 0から 遠い 方へ★＝JS の `Math.round` は 負で 違う（-2.5 → -2）ので 自分で 書く */
  function 丸める(引数たち, 手, 所, 向) {
    if (!引数たち.length) return 誤('#VALUE!');
    var t = 数の引数(引数たち[0], 手, 所);
    if (t.誤) return t.誤;
    var 桁 = 0;
    if (引数たち.length > 1) {
      var d = 数の引数(引数たち[1], 手, 所);
      if (d.誤) return d.誤;
      桁 = 切り捨て(d.数);
    }
    var p = Math.pow(10, 桁);
    if (!isFinite(p)) return 数(桁 > 0 ? t.数 : 0);
    var x = t.数 * p;
    if (!isFinite(x)) return 数(t.数);
    var 符 = x < 0 ? -1 : 1, a = Math.abs(x), r;
    if (向 === 'U') r = Math.ceil(a);
    else if (向 === 'D') r = Math.floor(a);
    else r = Math.floor(a + 0.5);          /* ★半分は 外へ★ */
    return 数(符 * r / p);
  }

  /** ★CEILING.MATH／FLOOR.MATH★
        `上か` true＝CEILING／false＝FLOOR
        ★3つ目（向き）は ★負の 数の 時だけ★ 効く★（0 以外＝0から 遠い 方へ） */
  function 数学丸め(引数たち, 手, 所, 上か) {
    if (!引数たち.length) return 誤('#VALUE!');
    var t = 数の引数(引数たち[0], 手, 所);
    if (t.誤) return t.誤;
    var k = 1;
    if (引数たち.length > 1) {
      var m = 数の引数(引数たち[1], 手, 所);
      if (m.誤) return m.誤;
      k = m.数;
    }
    var 向 = 0;
    if (引数たち.length > 2) {
      var d = 数の引数(引数たち[2], 手, 所);
      if (d.誤) return d.誤;
      向 = d.数;
    }
    if (k === 0) return 数(0);
    var x = t.数;
    if (x >= 0) return 数((上か ? Math.ceil(x / Math.abs(k)) : Math.floor(x / Math.abs(k))) * Math.abs(k));
    var 外へ = 向 !== 0;
    var 上げる = 上か ? !外へ : 外へ;      /* 負の 時 … 向き 0＝0に 近い 方／0以外＝遠い 方 */
    var q = x / Math.abs(k);
    return 数((上げる ? Math.ceil(q) : Math.floor(q)) * Math.abs(k));
  }

  /** ★GCD／LCM の 共通★（★切り捨ててから★・負は #NUM!） */
  function 整数の組(引数たち, 手, 続き, 元) {
    var 拾 = 数を拾う(引数たち, 手);
    if (拾.誤) return 拾.誤;
    var 出 = 元;
    for (var i = 0; i < 拾.数たち.length; i++) {
      var n = 切り捨て(拾.数たち[i]);
      if (n < 0) return 誤('#NUM!');
      出 = 続き(出, n);
      if (!isFinite(出)) return 誤('#NUM!');
    }
    return 数(出);
  }

  /** ★数を 1つ 取って 1つ 返す★（`続き` が null を 返したら #NUM!） */
  /* ══ ★★範囲を 渡されたら 1マスずつ 当てて 溢れさせる★★ ══（2026-09-21）
       ★なぜ 要るか★
         実Excel（動く並び）は `=ABS(A1:A3)` を ★3マス★に 溢れさせます。
         うちは 暗黙の 交わりで ★1マス★でした。
         ⇒★誤りに ならず 黙って 足りない＝お客さんは 気付けません★
       ★誤りも そのまま 並べます★（実Excel が そう して います）
         `=ACOTH(A1:A3)` ⇒ ★#NUM! / 0.549306 / 0.549306★（1マス目だけ 誤り）
         紙 `docs/measured/golden-jitsu-excel-hanni-zenbu-2026-09-20.tsv`
       ★注意★ 1マスの 範囲（A1:A1）は 今まで通り */
  function 範囲か(引数) {
    return !!(引数 && 引数.種 === '四角'
      && (((引数.行数 || 0) > 1) || ((引数.列数 || 0) > 1)));
  }
  function マスごとに溢れさせる(引数, 所, 一マス) {
    var 表 = 表にする(引数, 所);
    if (!表) return null;
    var 出 = [];
    for (var r = 0; r < 表.length; r++) {
      var 段 = [];
      for (var c = 0; c < 表[r].length; c++) 段.push(一マス(表[r][c] || 計.空));
      出.push(段);
    }
    return { 溢れ: true, 行数: 出.length, 列数: 出.length ? 出[0].length : 0, 並び: 出 };
  }

  /* ══ ★★中身では なく ★場所★を 1つずつ 渡して 溢れさせる★★ ══（2026-09-21）
       ★なぜ 上の `マスごとに溢れさせる` では 駄目か★
         上の 道具は ★値だけ★を `{ 種:'直', 値: 中身 }` で 渡します。
         ⇒`ISFORMULA` は ★値を 1度も 見て いません★＝見るのは
           ★そのマスに 打って ある 字★（`打った字`）です。
         ⇒値だけ 渡すと ★どのマスか 分からず `#VALUE!`★ に なります。
         ★2026-09-21 に 実際 そう なりました★
           `=ISFORMULA(A1:A3)` 実Excel ★FALSE / FALSE / FALSE★
           うち（値だけ 渡した 時）★#VALUE! / #VALUE! / #VALUE!★
           ⇒★数は 3マスで 合うのに 中身が 全部 誤り★＝★数だけ 見る 門では 捕まりません★
       ★この 道具が 渡す 物★ ... `{ 種:'マス', 値, 番地, 打った字 }`
         ＝★1マスの 参照と 同じ 顔★（`lib/shiki-hyou.js` が 作る 物と 同じ）
       ★並びは 行が 先★（`打った字(i)` の `i` も 行が 先＝`lib/shiki-hyou.js` と 同じ） */
  function マスの顔(値, 角, r, c, 列数, 打) {
    var i = r * 列数 + c;
    return {
      種: 'マス', 値: 値,
      番地: { 行: 角.行 + r, 列: 角.列 + c, 板: 角.板 },
      打った字: 打 ? function () { return 打(i); } : null,
    };
  }
  function 場所ごとに溢れさせる(引数, 所, 一マス) {
    var 角 = 引数 && 引数.番地;
    if (!角) return null;            /* ★どこか 分からない 物は 溢れさせません★ */
    var 表 = 表にする(引数, 所);
    if (!表 || !表.length) return null;
    var 列数 = 引数.列数 || 表[0].length;
    var 打 = (typeof 引数.打った字 === 'function') ? 引数.打った字 : null;
    var 出 = [];
    for (var r = 0; r < 表.length; r++) {
      var 段 = [];
      for (var c = 0; c < 表[r].length; c++) {
        段.push(一マス(マスの顔(表[r][c] || 計.空, 角, r, c, 列数, 打)));
      }
      出.push(段);
    }
    return { 溢れ: true, 行数: 出.length, 列数: 出.length ? 出[0].length : 0, 並び: 出 };
  }

  function 数ひとつ(引数たち, 手, 所, 続き) {
    if (!引数たち.length) return 誤('#VALUE!');
    /* ★一マスぶんの 計算を 切り出します★（範囲の 時に 1マスずつ 当てる 為） */
    var 一マス = function (v) {
      if (!v) return 誤('#VALUE!');
      if (v.型 === '誤') return v;
      var n = 計.数にする(v, 手);
      if (n.型 === '誤') return 誤('#VALUE!');
      var y = 続き(n.値);
      if (y === null || !isFinite(y)) return 誤('#NUM!');
      return 数(y);
    };
    if (範囲か(引数たち[0])) {
      var 溢 = マスごとに溢れさせる(引数たち[0], 所, 一マス);
      if (溢) return 溢;
    }
    var t = 数の引数(引数たち[0], 手, 所);
    if (t.誤) return t.誤;
    var 出 = 続き(t.数);
    if (出 === null || !isFinite(出)) return 誤('#NUM!');
    return 数(出);
  }
  /** ★数を 2つ 取って 1つ 返す★ */
  function 数ふたつ(引数たち, 手, 所, 続き) {
    if (引数たち.length < 2) return 誤('#VALUE!');
    var a = 数の引数(引数たち[0], 手, 所);
    if (a.誤) return a.誤;
    var b = 数の引数(引数たち[1], 手, 所);
    if (b.誤) return b.誤;
    var 出 = 続き(a.数, b.数);
    if (出 === null || !isFinite(出)) return 誤('#NUM!');
    return 数(出);
  }

  /** ★OR／XOR／AND の 共通★
        ★拾い方★ 数は ≠0／真偽は そのまま／★直の 字は TRUE FALSE と 数に 読める 物だけ★
                  ★マスの 字・真偽・空は 飛ばす★（四角の 決めと 同じ）
        ★1つも 拾えなければ #VALUE!★ */
  function 真偽の一族(引数たち, 手, 種) {
    var 並び = [], 誤り = null;
    ほどく(引数たち, function (v, 元) {
      if (v.型 === '誤') { 誤り = v; return false; }
      if (元 === 'マス') {
        if (v.型 === '数') 並び.push(v.値 !== 0);
        else if (v.型 === '真偽') 並び.push(!!v.値);
        return true;                                   /* ★字・空は 飛ばす★ */
      }
      if (v.型 === '真偽') { 並び.push(!!v.値); return true; }
      if (v.型 === '数') { 並び.push(v.値 !== 0); return true; }
      if (v.型 === '字') {
        var u = String(v.値).trim().toUpperCase();
        if (u === 'TRUE') { 並び.push(true); return true; }
        if (u === 'FALSE') { 並び.push(false); return true; }
        /* ★★数に 読める 字も 飛ばします★★（2026-09-15・実Excel の 実測）
             `=XOR("101",2)` … ★実Excel TRUE★
             ⇒ "101" を 数に すると 真が 2つ ⇒ XOR は FALSE に なって しまう
             ⇒★実Excel は "101" を ★数えて いません★★
             ★`=OR("101",2)` は どちらでも TRUE★＝★OR では 見分けが 付かない★
             ⇒★★XOR が 1本で 正体を 出しました★★ */
        return true;                                   /* ★字は 全部 飛ばす★ */
      }
      return true;                                     /* 空は 飛ばす */
    });
    if (誤り) return 誤り;
    if (!並び.length) return 誤('#VALUE!');
    var i;
    if (種 === 'AND') {
      for (i = 0; i < 並び.length; i++) if (!並び[i]) return 計.真偽(false);
      return 計.真偽(true);
    }
    if (種 === 'OR') {
      for (i = 0; i < 並び.length; i++) if (並び[i]) return 計.真偽(true);
      return 計.真偽(false);
    }
    var 数え = 0;
    for (i = 0; i < 並び.length; i++) if (並び[i]) 数え++;
    return 計.真偽(数え % 2 === 1);
  }

  /** ★A の 付く 散らばり★（VARA VARPA STDEVA STDEVPA） */
  function A散らばりで(引数たち, 手, 母集団か, 根か) {
    var 拾 = Aで数を拾う(引数たち, 手);
    if (拾.誤) return 拾.誤;
    var 分 = 散らばり(拾.数たち, 母集団か);
    if (分.型 === '誤') return 分;
    return 根か ? 数(Math.sqrt(分.値)) : 分;
  }

  /** ★散らばりの 一族の 共通★
        `母集団か` … true＝n で 割る（`.P`）／false＝n−1（`.S`）
        `根か`     … true＝平方根を 取る（STDEV系）／false＝そのまま（VAR系） */
  function 散らばりで(引数たち, 手, 母集団か, 根か) {
    var 拾 = 数を拾う(引数たち, 手);
    if (拾.誤) return 拾.誤;
    var 分 = 散らばり(拾.数たち, 母集団か);
    if (分.型 === '誤') return 分;
    return 根か ? 数(Math.sqrt(分.値)) : 分;
  }

  /* ══ ★関数たち★ ══ */
  var 表 = {
    /* ══ ★★よく 使う 9個★★ ══（2026-09-16） */
    CONCATENATE: function (引数たち, 手, 所) { return 表.CONCAT(引数たち, 手, 所); },
    /* ★TRUE()／FALSE()★ … ★括弧付きは 関数★（裸の TRUE は 名前で 処理済み） */
    TRUE: function () { return 計.真偽(true); },
    FALSE: function () { return 計.真偽(false); },
    /* ★★SUMPRODUCT★★
         ★四角の 中の 字・真偽は ★ 0 と して 掛ける★★（★飛ばしません★）
         ★物差し★ `golden-shikaku-kansuu-2026-09-13.tsv`
           材料 A1=1 ・ A2="2"（字）・ A3=TRUE ・ A4=空 ・ A5=2
           `=SUMPRODUCT(A1:A5,A1:A5)` → ★ 5★（＝1×1 ＋ 0 ＋ 0 ＋ 0 ＋ 2×2）
           ★飛ばして いたら 1+4＝5 に は なるが、形が 合わなく なります★
         ★形が 違う 四角を 渡されたら #VALUE!★ */
    SUMPRODUCT: function (引数たち, 手, 所) {
      if (!引数たち.length) return 誤('#VALUE!');
      var 表たち = [], i, j;
      for (i = 0; i < 引数たち.length; i++) {
        var t = 表にする(引数たち[i], 所);
        if (!t || !t.length) return 誤('#VALUE!');
        表たち.push(t);
      }
      var 行 = 表たち[0].length, 列 = 表たち[0][0].length;
      for (i = 1; i < 表たち.length; i++) {
        if (表たち[i].length !== 行 || 表たち[i][0].length !== 列) return 誤('#VALUE!');
      }
      var 和 = 0;
      for (var r = 0; r < 行; r++) {
        for (var c = 0; c < 列; c++) {
          var 積 = 1;
          for (i = 0; i < 表たち.length; i++) {
            var v = 表たち[i][r][c];
            if (v && v.型 === '誤') return v;
            /* ★字・真偽・空は 0★（上の 覚書き） */
            積 *= (v && v.型 === '数') ? v.値 : 0;
          }
          和 += 積;
        }
      }
      return 数(和);
    },
    /* ★FORMULATEXT★ … ★その マスに 打って ある 字★を 返す
         ★引数は 前から `打った字` を 持って います★（SUBTOTAL の 為に 入れて あった）
         ★式で なければ #N/A★ */
    FORMULATEXT: function (引数たち) {
      if (!引数たち.length) return 誤('#VALUE!');
      var a = 引数たち[0];
      if (!a || typeof a.打った字 !== 'function') return 誤('#N/A');
      var t = String(a.打った字(0) || '');
      if (t.charAt(0) !== '=') return 誤('#N/A');
      return 計.字(t);
    },
    /* ★★走らせる たびに 変わる 物★★
         ★紙では 押せません★（紙は 取った 時の 値を 焼いて いる）
         ⇒★名指しで まだに 出して あります★
         ★でも 無いと お客さんが 困る★＝★書かない 理由に は なりません★ */
    RAND: function () { return 数(Math.random()); },
    RANDBETWEEN: function (引数たち, 手, 所) {
      var g = 金の引数(引数たち, 手, 所, 2, 2);
      if (g.誤) return g.誤;
      var 下 = Math.ceil(g.数[0]), 上 = Math.floor(g.数[1]);
      if (下 > 上) return 誤('#NUM!');
      return 数(下 + Math.floor(Math.random() * (上 - 下 + 1)));
    },
    TODAY: function () {
      var d = new Date();
      return 数(日から通し(d.getFullYear(), d.getMonth() + 1, d.getDate()));
    },
    NOW: function () {
      var d = new Date();
      var 通 = 日から通し(d.getFullYear(), d.getMonth() + 1, d.getDate());
      var 秒 = d.getHours() * 3600 + d.getMinutes() * 60 + d.getSeconds();
      return 数(通 + 秒 / 86400);
    },

    /* ══ ★★お金の 残り 9個★★ ══（2026-09-16）
       ★物差し★ `golden-346-2026-09-08.tsv` ほか */
    TBILLEQ: function (引数たち, 手, 所) { return 短期証券(引数たち, 手, 所, 'EQ'); },
    TBILLPRICE: function (引数たち, 手, 所) { return 短期証券(引数たち, 手, 所, 'PRICE'); },
    TBILLYIELD: function (引数たち, 手, 所) { return 短期証券(引数たち, 手, 所, 'YIELD'); },
    CUMIPMT: function (引数たち, 手, 所) { return 積み上げの内訳(引数たち, 手, 所, true); },
    CUMPRINC: function (引数たち, 手, 所) { return 積み上げの内訳(引数たち, 手, 所, false); },
    DB: function (引数たち, 手, 所) { return 定率法(引数たち, 手, 所); },
    DDB: function (引数たち, 手, 所) { return 倍率法(引数たち, 手, 所); },
    RATE: function (引数たち, 手, 所) { return 利を探す(引数たち, 手, 所); },
    /* ★SERIESSUM★ … Σ aᵢ × x^(n + i×m)
         ★物差し★ `=SERIESSUM(0.5,1,2,3)` ★ 1.5★ ／ `=SERIESSUM(2,3,4,5)` ★ 40★ */
    SERIESSUM: function (引数たち, 手, 所) {
      if (引数たち.length < 4) return 誤('#VALUE!');
      var g = 金の引数([引数たち[0], 引数たち[1], 引数たち[2]], 手, 所, 3, 3);
      if (g.誤) return g.誤;
      var 拾 = 数を拾う([引数たち[3]], 手);
      if (拾.誤) return 拾.誤;
      var x = g.数[0], n = g.数[1], m = g.数[2], 和 = 0;
      for (var i = 0; i < 拾.数たち.length; i++) {
        和 += 拾.数たち[i] * Math.pow(x, n + i * m);
      }
      return 数(和);
    },

    /* ══ ★★日付の 字を 数に する 三つ★★ ══（2026-09-16）
       ★物差し★ `golden-1900-hidzuke-2026-09-15.tsv` */
    DATE: function (引数たち, 手, 所) {
      var g = 金の引数(引数たち, 手, 所, 3, 3);
      if (g.誤) return g.誤;
      var y = 切り捨て(g.数[0]), mo = 切り捨て(g.数[1]), da = 切り捨て(g.数[2]);
      if (y < 0 || y > 9999) return 誤('#NUM!');
      if (y < 1900) y += 1900;                /* ★実Excel の 決め（0〜1899 は 1900を 足す）★ */
      /* ★月・日は 繰り上がる★（`=DATE(2024,13,1)` は 2025年1月1日） */
      var y2 = y + Math.floor((mo - 1) / 12);
      var m2 = ((mo - 1) % 12 + 12) % 12 + 1;
      var 通 = 日から通し(y2, m2, 1) + (da - 1);
      if (!isFinite(通) || 通 < 0) return 誤('#NUM!');
      return 数(通);
    },
    DATEVALUE: function (引数たち, 手, 所) {
      var t = 字の引数(引数たち[0], 手, 所);
      if (t.誤) return t.誤;
      var n = 日付の字を読む(t.字);
      /* ★★時刻だけの 字は ★0★★（2026-09-16 実測）
           `=DATEVALUE("12:30")` … 実Excel ★0★（#VALUE! では 無い）
           ★逆は 違う★ `=TIMEVALUE("2024-01-15")` … ★#VALUE!★ */
      if (n === null && 時刻の字を読む(t.字) !== null) return 数(0);
      return n === null ? 誤('#VALUE!') : 数(n);
    },
    TIMEVALUE: function (引数たち, 手, 所) {
      var t = 字の引数(引数たち[0], 手, 所);
      if (t.誤) return t.誤;
      var n = 時刻の字を読む(t.字);
      /* ★★日付だけの 字は ★0★★（2026-09-16 実測）
           `=TIMEVALUE("2024-01-15")` … 実Excel ★0★（#VALUE! では 無い）
           ★DATEVALUE の 時刻と 向かい合い★（どちらも 0） */
      if (n === null && 日付の字を読む(t.字) !== null) return 数(0);
      return n === null ? 誤('#VALUE!') : 数(n);
    },

    /* ══ ★★統計の 一族（残り）★★ ══（2026-09-16）
       ★土台は 既に 在る ものだけ★
         誤差関数／ガンマの対数／上のガンマ／ベータの下／tの右／二分で逆
         ＝★作る 道を 何本も 作らない★
       ★物差し★ `golden-346-2026-09-08.tsv` ほか */
    'NORM.DIST': function (引数たち, 手, 所) { return 正規の分布(引数たち, 手, 所); },
    NORMDIST: function (引数たち, 手, 所) { return 正規の分布(引数たち, 手, 所); },
    'NORM.INV': function (引数たち, 手, 所) { return 正規の逆(引数たち, 手, 所); },
    NORMINV: function (引数たち, 手, 所) { return 正規の逆(引数たち, 手, 所); },
    'NORM.S.INV': function (引数たち, 手, 所) {
      var g = 金の引数(引数たち, 手, 所, 1, 1);
      if (g.誤) return g.誤;
      if (g.数[0] <= 0 || g.数[0] >= 1) return 誤('#NUM!');
      return 数(標準正規の逆(g.数[0]));
    },
    NORMSINV: function (引数たち, 手, 所) {
      var g = 金の引数(引数たち, 手, 所, 1, 1);
      if (g.誤) return g.誤;
      if (g.数[0] <= 0 || g.数[0] >= 1) return 誤('#NUM!');
      return 数(標準正規の逆(g.数[0]));
    },
    LOGINV: function (引数たち, 手, 所) { return 対数正規の逆(引数たち, 手, 所); },
    'LOGNORM.INV': function (引数たち, 手, 所) { return 対数正規の逆(引数たち, 手, 所); },
    /* ★二項の 一族★ */
    'BINOM.DIST': function (引数たち, 手, 所) { return 二項の分布(引数たち, 手, 所); },
    BINOMDIST: function (引数たち, 手, 所) { return 二項の分布(引数たち, 手, 所); },
    'BINOM.DIST.RANGE': function (引数たち, 手, 所) {
      var g = 金の引数(引数たち, 手, 所, 3, 4);
      if (g.誤) return g.誤;
      var n = 切り捨て(g.数[0]), p = g.数[1];
      var s1 = 切り捨て(g.数[2]);
      var s2 = g.数.length > 3 ? 切り捨て(g.数[3]) : s1;
      if (n < 0 || p < 0 || p > 1) return 誤('#NUM!');
      if (s1 < 0 || s1 > n || s2 < s1 || s2 > n) return 誤('#NUM!');
      var 和 = 0;
      for (var k = s1; k <= s2; k++) 和 += 二項の重み(k, n, p);
      return 数(和);
    },
    'BINOM.INV': function (引数たち, 手, 所) { return 二項の逆(引数たち, 手, 所); },
    CRITBINOM: function (引数たち, 手, 所) { return 二項の逆(引数たち, 手, 所); },
    'NEGBINOM.DIST': function (引数たち, 手, 所) { return 負の二項(引数たち, 手, 所); },
    NEGBINOMDIST: function (引数たち, 手, 所) { return 負の二項(引数たち, 手, 所); },
    HYPGEOMDIST: function (引数たち, 手, 所) {
      var g = 金の引数(引数たち, 手, 所, 4, 4);
      if (g.誤) return g.誤;
      var x = 切り捨て(g.数[0]), n = 切り捨て(g.数[1]);
      var M = 切り捨て(g.数[2]), N = 切り捨て(g.数[3]);
      if (x < 0 || n < 0 || M < 0 || N < 0) return 誤('#NUM!');
      if (n > N || M > N || x > n || x > M) return 誤('#NUM!');
      if (n - x > N - M) return 誤('#NUM!');
      var c1 = 組み合わせの数(M, x), c2 = 組み合わせの数(N - M, n - x);
      var c3 = 組み合わせの数(N, n);
      if (c1 !== null && c2 !== null && c3 !== null && c3 !== 0) return 数(c1 * c2 / c3);
      return 数(Math.exp(組み合わせの対数(M, x) + 組み合わせの対数(N - M, n - x)
        - 組み合わせの対数(N, n)));
    },
    /* ★ベータと ガンマ★ */
    'BETA.DIST': function (引数たち, 手, 所) { return ベータの分布(引数たち, 手, 所, true); },
    BETADIST: function (引数たち, 手, 所) { return ベータの分布(引数たち, 手, 所, false); },
    'BETA.INV': function (引数たち, 手, 所) { return ベータの逆(引数たち, 手, 所); },
    BETAINV: function (引数たち, 手, 所) { return ベータの逆(引数たち, 手, 所); },
    GAMMADIST: function (引数たち, 手, 所) { return 表['GAMMA.DIST'](引数たち, 手, 所); },
    'GAMMA.INV': function (引数たち, 手, 所) { return ガンマの逆(引数たち, 手, 所); },
    GAMMAINV: function (引数たち, 手, 所) { return ガンマの逆(引数たち, 手, 所); },
    /* ★左側の 逆★ */
    'CHISQ.INV': function (引数たち, 手, 所) {
      var g = 金の引数(引数たち, 手, 所, 2, 2);
      if (g.誤) return g.誤;
      var p = g.数[0], v = 切り捨て(g.数[1]);
      if (p < 0 || p >= 1 || v < 1) return 誤('#NUM!');
      return 数(二分で逆(function (x) { return 上のガンマ(v / 2, x / 2); }, 1 - p, 0, 1e10));
    },
    'F.INV': function (引数たち, 手, 所) { return Fの逆(引数たち, 手, 所, false); },
    FINV: function (引数たち, 手, 所) { return Fの逆(引数たち, 手, 所, true); },
    'T.INV': function (引数たち, 手, 所) {
      var g = 金の引数(引数たち, 手, 所, 2, 2);
      if (g.誤) return g.誤;
      var p = g.数[0], v = 切り捨て(g.数[1]);
      if (p <= 0 || p >= 1 || v < 1) return 誤('#NUM!');
      /* ★★真ん中を 二分で 探すな★★（2026-09-16 実測）
           `=T.INV(0.5,1)` … うち ★-1.05e-8★ ／実Excel ★0★
           ＝t の 右側は 0 の 周りで 平ら＝★数の 揺れで 場所が 決まる★
         ⇒★左右は 対称★＝★半分だけ 探して 折り返す★ */
      if (p === 0.5) return 数(0);
      var 右か = p > 0.5;
      var q = 右か ? (1 - p) : p;                 /* ★片側の 確からしさ★ */
      var x = 二分で逆(function (w) { return tの右(w, v); }, q, 0, 1e10);
      return 数(右か ? x : -x);
    },
    /* ★信頼の 幅★ */
    CONFIDENCE: function (引数たち, 手, 所) { return 信頼の幅(引数たち, 手, 所, false); },
    'CONFIDENCE.NORM': function (引数たち, 手, 所) { return 信頼の幅(引数たち, 手, 所, false); },
    'CONFIDENCE.T': function (引数たち, 手, 所) { return 信頼の幅(引数たち, 手, 所, true); },
    /* ★回帰と 相関★（★SLOPE などは Y が 先★＝実Excel と 同じ 順） */
    CORREL: function (引数たち, 手) {
      var d = 対の数(引数たち, 手);
      if (d.誤) return d.誤;
      if (d.Sxx === 0 || d.Syy === 0) return 誤('#DIV/0!');
      return 数(d.Sxy / Math.sqrt(d.Sxx * d.Syy));
    },
    PEARSON: function (引数たち, 手) {
      var d = 対の数(引数たち, 手);
      if (d.誤) return d.誤;
      if (d.Sxx === 0 || d.Syy === 0) return 誤('#DIV/0!');
      return 数(d.Sxy / Math.sqrt(d.Sxx * d.Syy));
    },
    RSQ: function (引数たち, 手) {
      var d = 対の数(引数たち, 手);
      if (d.誤) return d.誤;
      if (d.Sxx === 0 || d.Syy === 0) return 誤('#DIV/0!');
      var r = d.Sxy / Math.sqrt(d.Sxx * d.Syy);
      return 数(r * r);
    },
    SLOPE: function (引数たち, 手) {
      if (引数たち.length < 2) return 誤('#VALUE!');
      var d = 対の数([引数たち[1], 引数たち[0]], 手);
      if (d.誤) return d.誤;
      if (d.Sxx === 0) return 誤('#DIV/0!');
      return 数(d.Sxy / d.Sxx);
    },
    INTERCEPT: function (引数たち, 手) {
      if (引数たち.length < 2) return 誤('#VALUE!');
      var d = 対の数([引数たち[1], 引数たち[0]], 手);
      if (d.誤) return d.誤;
      if (d.Sxx === 0) return 誤('#DIV/0!');
      return 数(d.my - (d.Sxy / d.Sxx) * d.mx);
    },
    STEYX: function (引数たち, 手) {
      if (引数たち.length < 2) return 誤('#VALUE!');
      var d = 対の数([引数たち[1], 引数たち[0]], 手);
      if (d.誤) return d.誤;
      if (d.n < 3 || d.Sxx === 0) return 誤('#DIV/0!');
      var 残 = d.Syy - d.Sxy * d.Sxy / d.Sxx;
      if (残 < 0) 残 = 0;
      return 数(Math.sqrt(残 / (d.n - 2)));
    },
    /* ★検定★ */
    'Z.TEST': function (引数たち, 手, 所) { return Zの検定(引数たち, 手, 所); },
    ZTEST: function (引数たち, 手, 所) { return Zの検定(引数たち, 手, 所); },
    'CHISQ.TEST': function (引数たち, 手, 所) { return カイ二乗の検定(引数たち, 手, 所); },
    CHITEST: function (引数たち, 手, 所) { return カイ二乗の検定(引数たち, 手, 所); },
    'F.TEST': function (引数たち, 手) { return Fの検定(引数たち, 手); },
    FTEST: function (引数たち, 手) { return Fの検定(引数たち, 手); },
    'T.TEST': function (引数たち, 手, 所) { return tの検定(引数たち, 手, 所); },
    TTEST: function (引数たち, 手, 所) { return tの検定(引数たち, 手, 所); },
    /* ★順位と 分位と 度数★ */
    RANK: function (引数たち, 手, 所) { return 順位(引数たち, 手, 所, false); },
    'RANK.EQ': function (引数たち, 手, 所) { return 順位(引数たち, 手, 所, false); },
    'RANK.AVG': function (引数たち, 手, 所) { return 順位(引数たち, 手, 所, true); },
    'PERCENTILE.EXC': function (引数たち, 手, 所) { return 外の分位(引数たち, 手, 所, false); },
    'QUARTILE.EXC': function (引数たち, 手, 所) { return 外の分位(引数たち, 手, 所, true); },
    FREQUENCY: function (引数たち, 手) {
      if (引数たち.length < 2) return 誤('#VALUE!');
      var D = 数を拾う([引数たち[0]], 手);
      if (D.誤) return D.誤;
      var B = 数を拾う([引数たち[1]], 手);
      if (B.誤) return B.誤;
      var 仕 = B.数たち, 数え = [], i;
      for (i = 0; i <= 仕.length; i++) 数え.push(0);
      for (var j = 0; j < D.数たち.length; j++) {
        var v = D.数たち[j], k = 仕.length;
        for (var m = 0; m < 仕.length; m++) if (v <= 仕[m]) { k = m; break; }
        数え[k]++;
      }
      var 出 = [];
      for (i = 0; i < 数え.length; i++) 出.push([数(数え[i])]);
      return 溢れに(出);
    },

    /* ★#N/A を 作るだけ の 関数★（2026-09-15 に 入れた）
       ★実測★ `=NA()+1` → ★#N/A★／`=NA()+1/0` → ★#N/A★／`=NA()=NA()` → ★#N/A★
               （golden-zero-torinaoshi-2026-09-08.tsv ／ .Value2 は -2146826246）
               `=ERROR.TYPE(NA())` → ★7★（golden-2026-09-06.tsv）
       ★入れた 訳★＝2026-09-15 の 突き合わせで
         ★借り物だけが 正しく 答える 10通り★の 1本が `=ERROR.TYPE(NA())` でした
         （紙 7 ／ うち ★5★＝`NA()` を 知らず #NAME? に なって いた）
       ★引数は 取りません★（実Excel も `NA()` は 引数なし） */
    NA: function () { return 誤('#N/A'); },

    /* ★TEXT★ … ★書式の 台（`lib/shoshiki.js`）を 呼ぶだけ★
       ★★台は 1つ★★（指示役1 の 決め 2026-09-15）＝★画面（`js/book-open.js`）も 同じ 台★
       ★司さんの 実物では 740回★／★使う 書式は 2種類だけ★（aaa 731 ／ m/d 9）
       ★実測★（`docs/measured/kansuu46/golden-nokori4-kimari-2026-09-15.tsv` ほか）
         `=TEXT("あ","0")`        → ★あ★     （字は 素通り）
         `=TEXT(TRUE,"0")`        → ★TRUE★   （真偽も 素通り）
         `=TEXT(A9,"0")`          → ★0★      （空マスは 0）
         `=TEXT(1/0,"0")`         → ★#DIV/0!★（誤りは 通り抜ける）
         `=TEXT(1234.5,"")`       → ★空★
         `=TEXT(1234.5,"General")`→ ★#VALUE!★（★この国の Excel では 通らない★）
         `=TEXT(-0.004,"#,##0_ ")`→ ★"0 "★   （★-0 に しない★） */
    TEXT: function (引数たち, 手, 所) {
      var 今 = 所 && 所.今のマス;
      if (引数たち.length < 2) return 誤('#VALUE!');
      var v = 場.ひとつに(引数たち[0], 今);
      if (!v) return 誤('#VALUE!');
      if (v.型 === '誤') return v;
      var f = 場.ひとつに(引数たち[1], 今);
      if (!f) return 誤('#VALUE!');
      if (f.型 === '誤') return f;
      var 素 = (v.型 === '空') ? 0 : v.値;
      if (v.型 === '真偽') 素 = !!v.値;
      var 出 = 書.字にする(素, (f.型 === '空') ? '' : String(f.値));
      if (出 === null) return 誤('#VALUE!');
      return 計.字(出);
    },


    /* ★IF★
       ★実測で 分かった 事★（当て推量なら 4つ 外して いた）
         `=IF("TRUE",1,2)` → ★1★   （★字の TRUE は 真に なる★）
         `=IF("あ",1,2)`   → ★#VALUE!★ ／ `=IF("",1,2)` → ★#VALUE!★
         `=IF(FALSE,1)`    → ★FALSE★（★引数 2つの 時の 偽は FALSE★）
         `=IF(TRUE,,2)`    → ★0★   （★枝を 省くと 0★）
         `=IF(FALSE,1/0,2)`→ ★2★   （★選ばない 枝の 誤りは 見ない★）
       ★選ばない 枝も 計算して います★＝★答えは 同じ★（返さないので）。
         違うのは ★速さ★だけ。★遅さが 出たら そこで 直します★（今は 測って いません）。 */
    IF: function (引数たち, 手, 所) {
      var 今 = 所 && 所.今のマス;
      if (引数たち.length < 2) return 誤('#VALUE!');
      var 判 = 真偽にする(場.ひとつに(引数たち[0], 今));
      if (判.誤) return 判.誤;
      if (!判.真 && 引数たち.length < 3) return 計.真偽(false);   /* 実測 `=IF(FALSE,1)` → FALSE */
      var 枝 = 判.真 ? 引数たち[1] : 引数たち[2];
      var v = 場.ひとつに(枝, 今);
      return v || 計.空;
    },

    /* ★IFERROR★
       ★実測★ `=IFERROR("","x")` → ★空の 字★（★空の 字は 誤りでは ない★）
               `=IFERROR(A9,"x")` → ★0★（★空マスは 誤りでは ない★）
               `=IFERROR(1/0,)`   → ★0★（★2つ目を 省くと 0★）
               `=IFERROR(1)`      → ★実Excel が 式として 受け付けない★ */
    IFERROR: function (引数たち, 手, 所) {
      var 今 = 所 && 所.今のマス;
      if (引数たち.length < 2) return 誤('#VALUE!');   /* 実Excel は 式として 受け付けない */
      var v = 場.ひとつに(引数たち[0], 今);
      if (!v || v.型 !== '誤') return v || 計.空;
      var w = 場.ひとつに(引数たち[1], 今);
      return w || 計.空;
    },

    /* ★SUBTOTAL★ … ★本体は「入れ子の SUBTOTAL を 数えない」事★
       ★実測★ A6 に `=SUBTOTAL(9,A1:A5)` を 置いて
                 `=SUBTOTAL(9,A1:A6)` → ★15★ ／ `=SUM(A1:A6)` → ★30★
                 `=SUBTOTAL(2,A1:A6)` → ★5★（個数でも 数えない）
       ★実測（呼び方）★ `"9"` の 字でも 通る ／ `9.5` は 切り捨て
                         `0` と `12` は ★#VALUE!★ ／ 範囲で なく 数は ★式として 受け付けない★ */
    /* ══ ★★AGGREGATE★★ ══（2026-09-18）
       ★前は お客さんの 画面で #NAME? でした★（★台に 無い 50個の うちの 1個★）

       ★★紙（実Excel の 実測・12本）★★
         `docs/measured/golden-kansuu-7kaime-2026-09-18.tsv`
         ★材料★ A1:A5 ＝ 1,2,3,4,5 ／ ★B1:B5 ＝ 1,2,(#DIV/0!),4,5★
         ★10 / 12 が 見込み どおり★（★外れた 2本は 下★）
           =AGGREGATE(9,0,B1:B5)        → ★#DIV/0!★（選択 0 ＝ 何も 無視しない）
           =AGGREGATE(9,2,B1:B5)        → ★12★（誤りを 無視）
           =AGGREGATE(9,6,B1:B5)        → ★12★（誤りを 無視）
           =AGGREGATE(9,4,B1:B5)        → ★#DIV/0!★（何も 無視しない）
           =AGGREGATE(1,6,B1:B5)        → ★3★（12 ÷ 4）
           =AGGREGATE(14,6,B1:B5,2)     → ★4★（2番目に 大きい）
           =AGGREGATE(15,6,B1:B5,2)     → ★2★（2番目に 小さい）
           =AGGREGATE(9,6,B1:B5,2)      → ★#VALUE!★（★余分な 引数★）
           =AGGREGATE(14,6,B1:B5)       → ★#VALUE!★（★k が 要るのに 無い★）
           =AGGREGATE(9,0,A1:A5,B1:B5)  → ★#DIV/0!★（範囲 2つ）

       ★★★外れた 2本（★消しません★）★★★
         ①`=AGGREGATE(9,7,B1:B5)` → ★★12★★
            ★私の 見込み★ … #DIV/0!（「選択 7 ＝ ★隠した 行だけ★ 無視」と 読んだ）
            ⇒★★実は 誤りも 無視して います★★
            ⇒★「選択 4 と 7 は 同じに なる はず」という 見込みも 一緒に 外れました★
               （4 は #DIV/0! ／ 7 は 12）
         ②`=AGGREGATE(19,6,B1:B5,1)` → ★★1.25★★
            ★私の 見込み★ … 1.75（★機能 19 ＝ QUARTILE.★INC★ と 読んだ★）
            ⇒★★機能 19 は QUARTILE.★EXC★ です★★
            ★確かめ★ {1,2,4,5} で EXC(1) ＝ (4+1)×0.25 ＝ 1.25番目 ⇒ ★1.25★
                     INC(1) なら 1.75 に なります

       ★★選択（2つ目）の 決め★★（★実測から★）
         ★0・4★ … ★何も 無視しない★（★誤りが 在れば その 誤り★）
         ★2・6・7★ … ★誤りを 無視する★
         ★1・3・5★ … ★測って いません★（★2・6・7 と 同じに して います★）
         ⇒★訳★ … Excel の 覚書きでは 1〜3 が「隠した 行」、4〜7 が その 組み合わせ
                  ★但し 原文を 見て いません＝記憶です★
                  ⇒★★7 が 誤りを 無視する のは 測りました★★
                  ⇒★1・3・5 は `docs/measured/kansuu46/aggregate-kiku-koto.md` に 名指し★

       ★★19個の 機能は ★台に 全部 在ります★★（数え直しました）
         1 AVERAGE ／ 2 COUNT ／ 3 COUNTA ／ 4 MAX ／ 5 MIN ／ 6 PRODUCT ／
         7 STDEV.S ／ 8 STDEV.P ／ 9 SUM ／ 10 VAR.S ／ 11 VAR.P ／ 12 MEDIAN ／
         13 MODE.SNGL ／ 14 LARGE ／ 15 SMALL ／ 16 PERCENTILE.INC ／
         17 QUARTILE.INC ／ 18 PERCENTILE.EXC ／ ★19 QUARTILE.EXC★
         ⇒★呼び分けるだけ★（★中の 計算を 書き直して いません★）

       ★★k が 要る 機能★★ … ★14・15・16・17・18・19★
         ★要るのに 無い★ ⇒ #VALUE!（実測）／★要らないのに 在る★ ⇒ #VALUE!（実測） */
    AGGREGATE: function (引数たち, 手, 所) {
      if (引数たち.length < 3) return 誤('#VALUE!');
      var f = 数の引数(引数たち[0], 手, 所);
      if (f.誤) return f.誤;
      var o = 数の引数(引数たち[1], 手, 所);
      if (o.誤) return o.誤;
      var 機能 = 切り捨て(f.数), 選択 = 切り捨て(o.数);
      if (機能 < 1 || 機能 > 19) return 誤('#VALUE!');
      if (選択 < 0 || 選択 > 7) return 誤('#VALUE!');

      var kが要る = (機能 >= 14 && 機能 <= 19);
      var 範囲 = 引数たち.slice(2);
      var k = null;
      if (kが要る) {
        /* ★★ここに 門を 置きません★★（★置いて から 消しました★）
             ★わざと 外したら 赤に ならなかった★ ⇒ 調べたら
             ★LARGE 自身が 引数 0個でも 1個でも ★#VALUE!★ を 返します★
             ⇒★★見張りが 甲い のでは なく ★本当に 壊れて いない★★
             ⇒★死んだ コードを 残すと「守って いる」と 見えます★
             ★紙の `=AGGREGATE(14,6,B1:B5)` → #VALUE! は そのまま 合って います★ */
        k = 範囲.pop();
      } else if (範囲.length > 1) {
        /* ★余分な 引数は #VALUE!★（★実測 `=AGGREGATE(9,6,B1:B5,2)`★）
           ★但し ★範囲 2つ★は 通ります★（★実測 `=AGGREGATE(9,0,A1:A5,B1:B5)` → #DIV/0!★）
           ⇒★見分け★ … ★四角で ない 物が 混じって いたら 余分な 引数★ */
        for (var q = 1; q < 範囲.length; q++) {
          if (!範囲[q] || 範囲[q].種 !== '四角') return 誤('#VALUE!');
        }
      }

      /* ══ ★★誤りを 飛ばすのは 2・3・6・7 だけ★★ ══（2026-09-18・8枠目で 実測）
           `=AGGREGATE(9,1,B1:B5)` -> ★#DIV/0!★（★1 は 飛ばさない★）
           `=AGGREGATE(9,3,B1:B5)` -> ★12★    （★3 は 飛ばす★）
           `=AGGREGATE(9,5,B1:B5)` -> ★#DIV/0!★（★5 は 飛ばさない★）
           `=AGGREGATE(9,6,A1:A5,B1:B5,A1:A5)` -> ★42★（15+12+15＝6 は 飛ばす）
           ★前★ 「0 と 4 以外は 飛ばす」と 書いて いました ⇒ ★1 と 5 を 間違えて いた★
           ★訳★ ... ★選択は「何を 無視するか」の 組み合わせ★
                   ＝★誤りを 無視するのは 2・3・6・7 だけ★ */
      var 誤りを飛ばす = (選択 === 2 || 選択 === 3 || 選択 === 6 || 選択 === 7);
      if (!誤りを飛ばす) {
        for (var i = 0; i < 範囲.length; i++) {
          var a = 範囲[i];
          if (a && a.種 === '直' && a.値 && a.値.型 === '誤') return a.値;
          if (a && a.種 === '四角') {
            for (var j = 0; j < (a.並び || []).length; j++) {
              if (a.並び[j] && a.並び[j].型 === '誤') return a.並び[j];
            }
          } else if (a && a.値 && a.値.型 === '誤') { return a.値; }
        }
      }
      /* ★飛ばす 時は 誤りを 抜いた 四角に 作り直す★ */
      var 中 = [];
      for (var m = 0; m < 範囲.length; m++) {
        var b = 範囲[m];
        if (!誤りを飛ばす || !b || b.種 !== '四角') { 中.push(b); continue; }
        var 残 = [];
        for (var n = 0; n < (b.並び || []).length; n++) {
          if (b.並び[n] && b.並び[n].型 === '誤') continue;
          残.push(b.並び[n]);
        }
        中.push({ 種: '四角', 並び: 残, 行数: 残.length, 列数: 1 });
      }

      var 名 = ['AVERAGE', 'COUNT', 'COUNTA', 'MAX', 'MIN', 'PRODUCT',
                'STDEV.S', 'STDEV.P', 'SUM', 'VAR.S', 'VAR.P', 'MEDIAN',
                'MODE.SNGL', 'LARGE', 'SMALL', 'PERCENTILE.INC',
                'QUARTILE.INC', 'PERCENTILE.EXC', 'QUARTILE.EXC'][機能 - 1];
      var やる = 表[名];
      if (!やる) return 誤('#VALUE!');                 /* ★台に 無い＝起きない はず★ */
      return やる(kが要る ? 中.concat([k]) : 中, 手, 所);
    },

    SUBTOTAL: function (引数たち, 手, 所) {
      if (引数たち.length < 2) return 誤('#VALUE!');
      var t = 数の引数(引数たち[0], 手, 所);
      if (t.誤) return t.誤;
      var 番 = 切り捨て(t.数);
      if (番 >= 101 && 番 <= 111) 番 -= 100;   /* ★隠した 行を 抜く 形★＝今は 同じ（実測でも 同じ） */
      if (番 < 1 || 番 > 11) return 誤('#VALUE!');

      /* ★入れ子の SUBTOTAL を 外す★（★値だけでは 分からない＝打った 字を 見る★） */
      var 中 = [];
      for (var i = 1; i < 引数たち.length; i++) {
        var a = 引数たち[i];
        /* ★誤りは そのまま 返す★（2026-09-15＝実物で 1本 見つけた）
           ＝ は ★#REF!★（★#VALUE! では ない★）
           ★誤りは どの 関数でも 素通りする★＝SUBTOTAL の 決まりでは なく 誤りの 決まり */
        if (a.種 === '直' && a.値 && a.値.型 === '誤') return a.値;
        if (a.種 === '直') return 誤('#VALUE!');       /* 実Excel は 式として 受け付けない */
        if (a.種 !== '四角' || typeof a.打った字 !== 'function') { 中.push(a); continue; }
        var 残 = [];
        for (var j = 0; j < a.並び.length; j++) {
          if (/^=\s*SUBTOTAL\s*\(/i.test(a.打った字(j))) continue;
          残.push(a.並び[j]);
        }
        中.push({ 種: '四角', 並び: 残, 行数: 残.length, 列数: 1 });
      }

      if (番 === 1) return 表.AVERAGE(中, 手);
      if (番 === 2) return 表.COUNT(中, 手);
      if (番 === 3) return 表.COUNTA(中, 手);
      if (番 === 4) return 表.MAX(中, 手);
      if (番 === 5) return 表.MIN(中, 手);
      if (番 === 6) return 表.PRODUCT(中, 手);
      if (番 === 9) return 表.SUM(中, 手);
      var 拾 = 数を拾う(中, 手);
      if (拾.誤) return 拾.誤;
      var 分 = 散らばり(拾.数たち, 番 === 8 || 番 === 11);
      if (分.型 === '誤') return 分;
      return (番 === 10 || 番 === 11) ? 分 : 数(Math.sqrt(分.値));
    },


    /* ★★DSUM★★ … ★417個の うち ★最初の 1個★★（2026-09-15）
       ★指示役1「まず 1個 移して 手間を 出せ」★
       ★実測★ `=DSUM(A1:B5,1,D1:D2)` → ★14★ ／ `=DSUM(A1:B5,1,A1:A2)` → ★2★
       ★誤りは 素通り★（★未測定★＝台の 決め。SUM／MAX／AVERAGE は 実測で そう） */
    DSUM: function (引数たち, 手, 所) {
      var 拾 = D系の値たち(引数たち, 手, 所);
      if (拾.誤) return 拾.誤;
      var 和 = 0;
      for (var i = 0; i < 拾.値たち.length; i++) {
        var v = 拾.値たち[i];
        if (!v) continue;
        if (v.型 === '誤') return v;
        if (v.型 === '数') 和 += v.値;
      }
      return 数(和);
    },

    /* ★★D系の 残り 11個★★（2026-09-15）
       ★土台（`D系の値たち`）は DSUM の 時に 書きました★＝ここは ★拾い方だけ★
       ★物差し★ … 実Excel（`tests/shiki-kansuu-dkei.test.mjs`＝紙 21行＋17行）
       ★実Excel に 聞いて 初めて 分かった 物（★当て推量なら 外して いた★）★
         `=DPRODUCT(…)` 当たる 行 0件 → ★0★（★掛け算の 元は 1 なのに★）
         `=DGET(…)`     0件 → ★#VALUE!★／2件以上 → ★#NUM!★（★別の 誤り★）
         `=DCOUNT(…)`   0件 → ★0★／`=DAVERAGE(…)` 0件 → ★#DIV/0!★
       ★COUNTA は 誤りも 数える★（`COUNTA` と 同じ 決め） */

    /** ★選んだ 行の 数だけ 拾う★（★字と 真偽と 空は 飛ばす★＝四角の 決めと 同じ） */
    DAVERAGE: function (引数たち, 手, 所) { return D系の数で(引数たち, 手, 所, function (数たち) {
      if (!数たち.length) return 誤('#DIV/0!');
      var 和 = 0; for (var i = 0; i < 数たち.length; i++) 和 += 数たち[i];
      return 数(和 / 数たち.length);
    }); },

    DCOUNT: function (引数たち, 手, 所) { return D系の数で(引数たち, 手, 所, function (数たち) {
      return 数(数たち.length);
    }); },

    /* ★DCOUNTA は 誤りも 数える★＝`数で` では なく ★値そのもの★を 見る */
    DCOUNTA: function (引数たち, 手, 所) {
      var 拾 = D系の値たち(引数たち, 手, 所);
      if (拾.誤) return 拾.誤;
      var n = 0;
      for (var i = 0; i < 拾.値たち.length; i++) {
        var v = 拾.値たち[i];
        if (v && v.型 !== '空') n++;
      }
      return 数(n);
    },

    /* ★DGET は 0件と 2件以上で ★誤りが 違う★★（実Excel の 実測） */
    DGET: function (引数たち, 手, 所) {
      var 拾 = D系の値たち(引数たち, 手, 所);
      if (拾.誤) return 拾.誤;
      var 生 = [];
      for (var i = 0; i < 拾.値たち.length; i++) {
        var v = 拾.値たち[i];
        if (v && v.型 !== '空') 生.push(v);
      }
      if (生.length === 0) return 誤('#VALUE!');
      if (生.length > 1) return 誤('#NUM!');
      if (生[0].型 === '誤') return 生[0];
      return 生[0];
    },

    DMAX: function (引数たち, 手, 所) { return D系の数で(引数たち, 手, 所, function (数たち) {
      if (!数たち.length) return 数(0);          /* ★MAX と 同じ 決め★（実測 `=MAX(B1:B3)` → 0） */
      return 数(Math.max.apply(null, 数たち));
    }); },

    DMIN: function (引数たち, 手, 所) { return D系の数で(引数たち, 手, 所, function (数たち) {
      if (!数たち.length) return 数(0);
      return 数(Math.min.apply(null, 数たち));
    }); },

    /* ★0件は ★0★★（実Excel の 実測・★1 では ない★） */
    DPRODUCT: function (引数たち, 手, 所) { return D系の数で(引数たち, 手, 所, function (数たち) {
      if (!数たち.length) return 数(0);
      var 積 = 1; for (var i = 0; i < 数たち.length; i++) 積 *= 数たち[i];
      return 数(積);
    }); },

    DSTDEV: function (引数たち, 手, 所) { return D系の数で(引数たち, 手, 所, function (数たち) {
      var 分 = 散らばり(数たち, false);
      return (分.型 === '誤') ? 分 : 数(Math.sqrt(分.値));
    }); },
    DSTDEVP: function (引数たち, 手, 所) { return D系の数で(引数たち, 手, 所, function (数たち) {
      var 分 = 散らばり(数たち, true);
      return (分.型 === '誤') ? 分 : 数(Math.sqrt(分.値));
    }); },
    DVAR: function (引数たち, 手, 所) { return D系の数で(引数たち, 手, 所, function (数たち) {
      return 散らばり(数たち, false);
    }); },
    DVARP: function (引数たち, 手, 所) { return D系の数で(引数たち, 手, 所, function (数たち) {
      return 散らばり(数たち, true);
    }); },

    /* ══ ★★散らばりの 一族★★ ══（2026-09-15）
       ★物差し★ … `docs/measured/kansuu46/golden-346-2026-09-08.tsv`（実Excel 16.0 build 20326）
       ★見張り★ … `tests/shiki-kansuu-kami.test.mjs`（★紙 60枚 全部で 押す★）
       ★土台は もう 在ります★ … `数を拾う`（★四角の 中の 字と 真偽は 飛ばす★）／`散らばり`
       ★★`.S` と `.P` の 違い★★ … `.S`（標本）は n−1／`.P`（母集団）は n
         ⇒`散らばり(数たち, 母集団か)` の 第2引数だけ 違います
       ★別名★ VAR＝VAR.S ／ VARP＝VAR.P ／ STDEV＝STDEV.S ／ STDEVP＝STDEV.P
         （★1つの 中身を 2つの 名前で 指す＝二重実装に しない★） */
    VAR: function (引数たち, 手, 所) { return 散らばりで(引数たち, 手, false, false); },
    'VAR.S': function (引数たち, 手, 所) { return 散らばりで(引数たち, 手, false, false); },
    VARP: function (引数たち, 手, 所) { return 散らばりで(引数たち, 手, true, false); },
    'VAR.P': function (引数たち, 手, 所) { return 散らばりで(引数たち, 手, true, false); },
    STDEV: function (引数たち, 手, 所) { return 散らばりで(引数たち, 手, false, true); },
    'STDEV.S': function (引数たち, 手, 所) { return 散らばりで(引数たち, 手, false, true); },
    STDEVP: function (引数たち, 手, 所) { return 散らばりで(引数たち, 手, true, true); },
    'STDEV.P': function (引数たち, 手, 所) { return 散らばりで(引数たち, 手, true, true); },

    /* ★MEDIAN★ … 真ん中（偶数なら 2つの 平均） */
    MEDIAN: function (引数たち, 手) {
      var 拾 = 数を拾う(引数たち, 手);
      if (拾.誤) return 拾.誤;
      var a = 拾.数たち.slice().sort(function (x, y) { return x - y; });
      if (!a.length) return 誤('#NUM!');
      var m = a.length >> 1;
      return 数(a.length % 2 ? a[m] : (a[m - 1] + a[m]) / 2);
    },

    /* ★SUMSQ★ … 二乗の 和 */
    SUMSQ: function (引数たち, 手) {
      var 拾 = 数を拾う(引数たち, 手);
      if (拾.誤) return 拾.誤;
      var 和 = 0;
      for (var i = 0; i < 拾.数たち.length; i++) 和 += 拾.数たち[i] * 拾.数たち[i];
      return 数(和);
    },

    /* ★DEVSQ★ … 平均からの ずれの 二乗の 和 */
    DEVSQ: function (引数たち, 手) {
      var 拾 = 数を拾う(引数たち, 手);
      if (拾.誤) return 拾.誤;
      var a = 拾.数たち;
      if (!a.length) return 誤('#NUM!');
      var 和 = 0, i;
      for (i = 0; i < a.length; i++) 和 += a[i];
      var 平 = 和 / a.length, 二 = 0;
      for (i = 0; i < a.length; i++) 二 += (a[i] - 平) * (a[i] - 平);
      return 数(二);
    },

    /* ★AVEDEV★ … 平均からの ずれ（絶対値）の 平均 */
    AVEDEV: function (引数たち, 手) {
      var 拾 = 数を拾う(引数たち, 手);
      if (拾.誤) return 拾.誤;
      var a = 拾.数たち;
      if (!a.length) return 誤('#NUM!');
      var 和 = 0, i;
      for (i = 0; i < a.length; i++) 和 += a[i];
      var 平 = 和 / a.length, 差 = 0;
      for (i = 0; i < a.length; i++) 差 += Math.abs(a[i] - 平);
      return 数(差 / a.length);
    },

    /* ★GEOMEAN★ … 掛けて n乗根（★0や 負が 混ざると #NUM!★） */
    GEOMEAN: function (引数たち, 手) {
      var 拾 = 数を拾う(引数たち, 手);
      if (拾.誤) return 拾.誤;
      var a = 拾.数たち;
      if (!a.length) return 誤('#NUM!');
      /* ★★形を 2つ 作って 測りました★★（2026-09-15・紙 37行）
           ㋐★掛けて n乗根★     … ★36/37★  ←★こちらを 採った★
           ㋑`EXP(LN の 平均)`  … 35/37
           ★実Excel 自身が 1つの 形では ありません★
             `=GEOMEAN(0.5,1)` … 実Excel `0.70710678118654757`＝★√0.5 そのもの★（㋐が 合う）
             `=GEOMEAN(2,TRUE)` … 実Excel `1.4142135623730949`＝★√2 より 1つ 下★（㋑が 合う）
           ⇒★★両方 合う 形は 作れません★★⇒★合う 本数の 多い 方を 採り、残り 1本は 名指しで 棚★
           ⇒`tests/shiki-kansuu-kami.test.mjs` の `まだ` に `=GEOMEAN(2,TRUE)` */
      var 積 = 1;
      for (var i = 0; i < a.length; i++) {
        if (a[i] <= 0) return 誤('#NUM!');
        積 *= a[i];
      }
      return 数(Math.pow(積, 1 / a.length));
    },

    /* ★HARMEAN★ … 逆数の 平均の 逆数（★0や 負が 混ざると #NUM!★） */
    HARMEAN: function (引数たち, 手) {
      var 拾 = 数を拾う(引数たち, 手);
      if (拾.誤) return 拾.誤;
      var a = 拾.数たち;
      if (!a.length) return 誤('#NUM!');
      /* ★`n ÷ Σ(1/x)` では なく ★`1 ÷ (Σ(1/x) ÷ n)`★★（2026-09-15 実測）
           `=HARMEAN(1,2,3,4,5,6)` … 前 `2.4489795918367347` ／実Excel `2.4489795918367352`
           ⇒★後の 形で ★同じ double★★（GEOMEAN と 同じ 型の 話） */
      var 逆 = 0;
      for (var i = 0; i < a.length; i++) {
        if (a[i] <= 0) return 誤('#NUM!');
        逆 += 1 / a[i];
      }
      return 数(1 / (逆 / a.length));
    },

    /* ══ ★★分布の 残り★★ ══（2026-09-16）
       ★物差し★ `golden-346-2026-09-08.tsv`
       ★左側／右側／逆引きを ★不完全ガンマ・ベータ★ 1つの 台で 出します★ */
    'F.DIST.RT': function (引数たち, 手, 所) {
      var g = 金の引数(引数たち, 手, 所, 3, 3);
      if (g.誤) return g.誤;
      var x = g.数[0], d1 = 切り捨て(g.数[1]), d2 = 切り捨て(g.数[2]);
      if (x < 0 || d1 < 1 || d2 < 1) return 誤('#NUM!');
      return 数(1 - ベータの下(d1 / 2, d2 / 2, d1 * x / (d1 * x + d2)));
    },
    FDIST: function (引数たち, 手, 所) { return 表['F.DIST.RT'](引数たち, 手, 所); },
    'F.DIST': function (引数たち, 手, 所) {
      if (引数たち.length < 4) return 誤('#VALUE!');
      var g = 金の引数([引数たち[0], 引数たち[1], 引数たち[2]], 手, 所, 3, 3);
      if (g.誤) return g.誤;
      var 判 = 真偽にする(場.ひとつに(引数たち[3], 所 && 所.今のマス));
      if (判.誤) return 判.誤;
      var x = g.数[0], d1 = 切り捨て(g.数[1]), d2 = 切り捨て(g.数[2]);
      if (x < 0 || d1 < 1 || d2 < 1) return 誤('#NUM!');
      if (判.真) return 数(ベータの下(d1 / 2, d2 / 2, d1 * x / (d1 * x + d2)));
      var 頭 = Math.exp(ガンマの対数((d1 + d2) / 2) - ガンマの対数(d1 / 2) - ガンマの対数(d2 / 2));
      return 数(頭 * Math.pow(d1 / d2, d1 / 2) * Math.pow(x, d1 / 2 - 1)
        * Math.pow(1 + d1 * x / d2, -(d1 + d2) / 2));
    },
    'CHISQ.INV.RT': function (引数たち, 手, 所) { return カイ二乗の逆(引数たち, 手, 所); },
    CHIINV: function (引数たち, 手, 所) { return カイ二乗の逆(引数たち, 手, 所); },
    'T.INV.2T': function (引数たち, 手, 所) { return tの逆両側(引数たち, 手, 所); },
    TINV: function (引数たち, 手, 所) { return tの逆両側(引数たち, 手, 所); },
    'F.INV.RT': function (引数たち, 手, 所) {
      var g = 金の引数(引数たち, 手, 所, 3, 3);
      if (g.誤) return g.誤;
      var p = g.数[0], d1 = 切り捨て(g.数[1]), d2 = 切り捨て(g.数[2]);
      if (p < 0 || p > 1 || d1 < 1 || d2 < 1) return 誤('#NUM!');
      return 数(二分で逆(function (x) {
        return 1 - ベータの下(d1 / 2, d2 / 2, d1 * x / (d1 * x + d2));
      }, p, 0, 1e12));
    },
    TDIST: function (引数たち, 手, 所) {
      var g = 金の引数(引数たち, 手, 所, 3, 3);
      if (g.誤) return g.誤;
      var t = g.数[0], v = 切り捨て(g.数[1]), 尾 = 切り捨て(g.数[2]);
      if (t < 0 || v < 1 || (尾 !== 1 && 尾 !== 2)) return 誤('#NUM!');
      return 数(尾 * tの右(t, v));
    },
    LOGNORMDIST: function (引数たち, 手, 所) {
      var g = 金の引数(引数たち, 手, 所, 3, 3);
      if (g.誤) return g.誤;
      var x = g.数[0], m = g.数[1], sd = g.数[2];
      if (x <= 0 || sd <= 0) return 誤('#NUM!');
      return 数((1 + 誤差関数((Math.log(x) - m) / (sd * Math.SQRT2))) / 2);
    },
    'LOGNORM.DIST': function (引数たち, 手, 所) {
      if (引数たち.length < 4) return 誤('#VALUE!');
      var g = 金の引数([引数たち[0], 引数たち[1], 引数たち[2]], 手, 所, 3, 3);
      if (g.誤) return g.誤;
      var 判 = 真偽にする(場.ひとつに(引数たち[3], 所 && 所.今のマス));
      if (判.誤) return 判.誤;
      var x = g.数[0], m = g.数[1], sd = g.数[2];
      if (x <= 0 || sd <= 0) return 誤('#NUM!');
      if (判.真) return 数((1 + 誤差関数((Math.log(x) - m) / (sd * Math.SQRT2))) / 2);
      var z = (Math.log(x) - m) / sd;
      return 数(Math.exp(-z * z / 2) / (x * sd * Math.sqrt(2 * Math.PI)));
    },
    'WEIBULL.DIST': function (引数たち, 手, 所) { return ワイブル(引数たち, 手, 所); },
    WEIBULL: function (引数たち, 手, 所) { return ワイブル(引数たち, 手, 所); },
    'GAMMA.DIST': function (引数たち, 手, 所) {
      if (引数たち.length < 4) return 誤('#VALUE!');
      var g = 金の引数([引数たち[0], 引数たち[1], 引数たち[2]], 手, 所, 3, 3);
      if (g.誤) return g.誤;
      var 判 = 真偽にする(場.ひとつに(引数たち[3], 所 && 所.今のマス));
      if (判.誤) return 判.誤;
      var x = g.数[0], a2 = g.数[1], b2 = g.数[2];
      if (x < 0 || a2 <= 0 || b2 <= 0) return 誤('#NUM!');
      if (判.真) return 数(1 - 上のガンマ(a2, x / b2));
      return 数(Math.exp((a2 - 1) * Math.log(x) - x / b2 - a2 * Math.log(b2) - ガンマの対数(a2)));
    },
    'HYPGEOM.DIST': function (引数たち, 手, 所) {
      if (引数たち.length < 5) return 誤('#VALUE!');
      var g = 金の引数([引数たち[0], 引数たち[1], 引数たち[2], 引数たち[3]], 手, 所, 4, 4);
      if (g.誤) return g.誤;
      var 判 = 真偽にする(場.ひとつに(引数たち[4], 所 && 所.今のマス));
      if (判.誤) return 判.誤;
      var k = 切り捨て(g.数[0]), n = 切り捨て(g.数[1]);
      var K = 切り捨て(g.数[2]), N = 切り捨て(g.数[3]);
      if (k < 0 || n < 1 || K < 0 || N < 1 || n > N || K > N || k > n || k > K) return 誤('#NUM!');
      var 密 = function (i) {
        return Math.exp(組み合わせの対数(K, i) + 組み合わせの対数(N - K, n - i)
          - 組み合わせの対数(N, n));
      };
      if (!判.真) return 数(密(k));
      var 和 = 0;
      for (var i = 0; i <= k; i++) if (n - i <= N - K) 和 += 密(i);
      return 数(和);
    },

    /* ══ ★★二乗和・共分散・進数の 続き★★ ══（2026-09-16）
       ★物差し★ `golden-346-2026-09-08.tsv` */
    SUMX2MY2: function (引数たち, 手) { return 二つの二乗(引数たち, 手, 'X2MY2'); },
    SUMX2PY2: function (引数たち, 手) { return 二つの二乗(引数たち, 手, 'X2PY2'); },
    SUMXMY2: function (引数たち, 手) { return 二つの二乗(引数たち, 手, 'XMY2'); },
    COVAR: function (引数たち, 手) { return 共分散(引数たち, 手, true); },
    'COVARIANCE.P': function (引数たち, 手) { return 共分散(引数たち, 手, true); },
    'COVARIANCE.S': function (引数たち, 手) { return 共分散(引数たち, 手, false); },
    STANDARDIZE: function (引数たち, 手, 所) {
      var g = 金の引数(引数たち, 手, 所, 3, 3);
      if (g.誤) return g.誤;
      if (g.数[2] <= 0) return 誤('#NUM!');
      return 数((g.数[0] - g.数[1]) / g.数[2]);
    },
    PDURATION: function (引数たち, 手, 所) {
      var g = 金の引数(引数たち, 手, 所, 3, 3);
      if (g.誤) return g.誤;
      var r = g.数[0], pv = g.数[1], fv = g.数[2];
      if (r <= 0 || pv <= 0 || fv <= 0) return 誤('#NUM!');
      return 数((Math.log(fv) - Math.log(pv)) / Math.log(1 + r));
    },
    ISPMT: function (引数たち, 手, 所) {
      var g = 金の引数(引数たち, 手, 所, 4, 4);
      if (g.誤) return g.誤;
      var r = g.数[0], 期 = g.数[1], n = g.数[2], pv = g.数[3];
      if (n === 0) return 誤('#DIV/0!');
      return 数(pv * r * (期 / n - 1));
    },
    FACTDOUBLE: function (引数たち, 手, 所) {
      return 数ひとつ(引数たち, 手, 所, function (x) {
        var n = 切り捨て(x);
        if (n < -1) return null;
        if (n <= 0) return 1;
        var r = 1;
        for (var i = n; i > 1; i -= 2) r *= i;
        return r;
      });
    },
    ISFORMULA: function (引数たち) {
      /* ★この 台は ★指した 先に 式が 在るか★ を まだ 持ちません★
           ⇒★引数が 指した 場所か どうかだけ 見て `FALSE`★（★半分 合う 答えを 出さない★）
           ★紙の 6行は どれも 式で ない マスを 指して います★ */
      if (!引数たち.length) return 誤('#VALUE!');
      var a = 引数たち[0];
      if (a.種 !== 'マス' && a.種 !== '四角') return 誤('#VALUE!');
      /* ★★台は マスに ★打った 字★を 持って います★★（2026-09-16）
           ⇒★`=` で 始まれば 式★＝★本当に 見られます★
           ★最初 私は「台は 持って いない」と 決めつけて ★いつも FALSE★を 返しました★
           ⇒`=ISFORMULA(D1)`（D1＝`=DATE(2024,1,1)`）… 実Excel ★TRUE★ で 赤に なった
           ⇒★★「持って いない」と 決める 前に 口を 探す★★ */
      if (typeof a.打った字 !== 'function') return 計.真偽(false);
      var 字 = String(a.打った字(0) || '');
      return 計.真偽(字.charAt(0) === '=');
    },
    /* ★進数の 組み合わせ★ */
    OCT2HEX: function (引数たち, 手, 所) { return 進数を換える(引数たち, 手, 所, 8, 16); },
    OCT2BIN: function (引数たち, 手, 所) { return 進数を換える(引数たち, 手, 所, 8, 2); },
    HEX2OCT: function (引数たち, 手, 所) { return 進数を換える(引数たち, 手, 所, 16, 8); },
    HEX2BIN: function (引数たち, 手, 所) { return 進数を換える(引数たち, 手, 所, 16, 2); },
    BIN2HEX: function (引数たち, 手, 所) { return 進数を換える(引数たち, 手, 所, 2, 16); },
    BIN2OCT: function (引数たち, 手, 所) { return 進数を換える(引数たち, 手, 所, 2, 8); },

    /* ══ ★★行列と 分布と お金の 続き★★ ══（2026-09-16）
       ★物差し★ `golden-346-2026-09-08.tsv` */
    MMULT: function (引数たち, 手, 所) {
      if (引数たち.length < 2) return 誤('#VALUE!');
      var A = 表にする(引数たち[0], 所), B = 表にする(引数たち[1], 所);
      if (!A || !B || !A.length || !B.length) return 誤('#VALUE!');
      if (A[0].length !== B.length) return 誤('#VALUE!');
      var 出 = [];
      for (var r = 0; r < A.length; r++) {
        var 段 = [];
        for (var c = 0; c < B[0].length; c++) {
          var 和 = 0;
          for (var k = 0; k < B.length; k++) {
            var x = A[r][k], y = B[k][c];
            if (!x || x.型 !== '数' || !y || y.型 !== '数') return 誤('#VALUE!');
            和 += x.値 * y.値;
          }
          段.push(数(和));
        }
        出.push(段);
      }
      return 溢れに(出);
    },
    /* ══ ★★台に 無かった 49個の 続き（1束目）★★ ══（2026-09-18）
       ★どこから 取ったか★ ... `docs/measured/golden-346-2026-09-08.tsv`（★46個の 紙★）
       ★★但し 1つの 関数に つき 1行しか 在りません★★
       ⇒★★合わせられるのは その 1行だけ★★
       ⇒★端（誤り・負・空・長さ違い）は ★未測定★★＝★各関数の 下に 名指しで 書きます★ */

    /** ★MDETERM★ ... 行列式
        ★紙★ `=MDETERM(A1:B2)` → 0（A1:B2 ＝ 1,2／2,4）
        ★書き方★ ★部分ピボットの ガウス消去★
          ＝★余因子展開は n! で 増えます★（8×8 で 4万回）
        ★★未測定★★ ... 正方で ない／字が 混ざる／空マス（#VALUE! の はず） */
    MDETERM: function (引数たち, 手, 所) {
      if (!引数たち.length) return 誤('#VALUE!');
      var A = 表にする(引数たち[0], 所);
      if (!A || !A.length || !A[0].length) return 誤('#VALUE!');
      var n = A.length, r, c;
      /* ══ ★★誤りは 正方の 検めより 先★★ ══（2026-09-18・9枠目 1行目で 実測）
           ★実Excel★ `=MDETERM(A1:B5)`（5行2列・★B3 に `=1/0`★）-> ★#DIV/0!★
           ★私の 見込み★ #VALUE!（正方で ないから）⇒★外れました★
           ★訳★ ... ★正方で ない 事より 先に ★中の 誤りが 通り抜けます★★
           ★★私は 見込みを 書く 時 B3 が 四角の 中に 在る 事を 忘れて いました★★
             ＝★材料を 見ずに 形だけで 見込みを 書いた★
           ★まだ 聞いて いない★ ... ★誤りの 無い 非正方★（#VALUE! の はず） */
      for (r = 0; r < A.length; r++) {
        for (c = 0; c < A[r].length; c++) {
          if (A[r][c] && A[r][c].型 === '誤') return A[r][c];
        }
      }
      if (A[0].length !== n) return 誤('#VALUE!');
      var m = [];
      for (r = 0; r < n; r++) {
        m.push([]);
        for (c = 0; c < n; c++) {
          var x = A[r][c];
          if (!x) return 誤('#VALUE!');
          if (x.型 !== '数') return 誤('#VALUE!');
          m[r].push(x.値);
        }
      }
      var 出 = 1;
      for (c = 0; c < n; c++) {
        var 一番 = c;
        for (r = c + 1; r < n; r++) if (Math.abs(m[r][c]) > Math.abs(m[一番][c])) 一番 = r;
        if (m[一番][c] === 0) return 数(0);
        if (一番 !== c) { var 仮 = m[c]; m[c] = m[一番]; m[一番] = 仮; 出 = -出; }
        出 *= m[c][c];
        for (r = c + 1; r < n; r++) {
          var 倍 = m[r][c] / m[c][c];
          for (var k = c; k < n; k++) m[r][k] -= 倍 * m[c][k];
        }
      }
      if (!isFinite(出)) return 誤('#NUM!');
      return 数(出);
    },

    /** ★KURT★ ... 尖り（★標本★）
        ★紙★ `=KURT(A1:A5)` → -1.1999999999999984（A1:A5 ＝ 1,2,3,4,5）
        ★式★ n(n+1)/((n-1)(n-2)(n-3)) × Σ((x-平)/s)^4 − 3(n-1)²/((n-2)(n-3))
          （s ＝ ★標本の 標準偏差（n-1 で 割る）★）
        ★★紙の 字が -1.2 では なく -1.1999999999999984★★
          ＝★丸めずに 出る 消え残り★＝★この 順で 計算した 時に 出ます★
        ★★未測定★★ ... n が 4未満（#DIV/0! の はず）／s が 0 */
    KURT: function (引数たち, 手) {
      var 拾 = 数を拾う(引数たち, 手);
      if (拾.誤) return 拾.誤;
      var a = 拾.数たち, n = a.length, i;
      if (n < 4) return 誤('#DIV/0!');
      var 和 = 0;
      for (i = 0; i < n; i++) 和 += a[i];
      var 平 = 和 / n, 二 = 0;
      for (i = 0; i < n; i++) 二 += (a[i] - 平) * (a[i] - 平);
      var s = Math.sqrt(二 / (n - 1));
      if (s === 0) return 誤('#DIV/0!');
      var 四 = 0;
      for (i = 0; i < n; i++) { var z = (a[i] - 平) / s; 四 += z * z * z * z; }
      var 出 = n * (n + 1) / ((n - 1) * (n - 2) * (n - 3)) * 四
        - 3 * (n - 1) * (n - 1) / ((n - 2) * (n - 3));
      if (!isFinite(出)) return 誤('#NUM!');
      return 数(出);
    },

    /** ★TRIMMEAN★ ... 端を 落とした 平均
        ★紙★ `=TRIMMEAN(A1:A5,0.2)` → 3
        ★落とす 数★ ＝ ★floor(n × 割合 ÷ 2)★ を ★上下 それぞれ★
          ＝5 × 0.2 ÷ 2 ＝ 0.5 ⇒ 切り捨て 0 ⇒★1つも 落ちない★⇒平均 3 ✔
        ★★未測定★★ ... 割合が 0未満／1以上（#NUM! の はず）／全部 落ちる 時 */
    TRIMMEAN: function (引数たち, 手, 所) {
      if (引数たち.length < 2) return 誤('#VALUE!');
      var w = 数の引数(引数たち[引数たち.length - 1], 手, 所);
      if (w.誤) return w.誤;
      if (w.数 < 0 || w.数 >= 1) return 誤('#NUM!');
      var 拾 = 数を拾う(引数たち.slice(0, 引数たち.length - 1), 手);
      if (拾.誤) return 拾.誤;
      var a = 拾.数たち.slice().sort(function (x, y) { return x - y; });
      var n = a.length;
      if (!n) return 誤('#NUM!');
      var 落 = Math.floor(n * w.数 / 2);
      var 残 = a.slice(落, n - 落);
      if (!残.length) return 誤('#NUM!');
      var 和 = 0;
      for (var i = 0; i < 残.length; i++) 和 += 残[i];
      return 数(和 / 残.length);
    },

    /** ★FORECAST／FORECAST.LINEAR★ ... 直線で 当てる
        ★紙★ `=FORECAST(6,B1:B5,A1:A5)` → 12
          （B ＝ 2,4,6,8,10 ／ A ＝ 1,2,3,4,5 ⇒ y ＝ 2x ⇒ x=6 で 12）
        ★★引数の 順に 気を つける★★ ... (x, ★y の 並び★, ★x の 並び★)＝★y が 先★ */
    FORECAST: function (引数たち, 手, 所) { return 直線で当てる(引数たち, 手, 所); },
    'FORECAST.LINEAR': function (引数たち, 手, 所) { return 直線で当てる(引数たち, 手, 所); },

    /** ★NUMBERVALUE★ ... 字を 数に する（★区切りを 口で 決める★）
        ★紙★ `=NUMBERVALUE("1.5")` → 1.5
        ★引数★ (字, 小数の 印, 桁の 印)／★既定★ 小数 "." ／ 桁 ","
        ★決まり★ ・桁の 印は 全部 落とす ・終わりの ％は 100で 割る（重ねられる）
                  ・空の 字は 0 ・印が 2字 以上なら 1字目だけ
        ★★未測定★★ ... 小数の 印が 2回 出る／印が 同じ 字／土地で 既定が 変わるか */
    NUMBERVALUE: function (引数たち, 手, 所) {
      if (!引数たち.length) return 誤('#VALUE!');
      var t = 字の引数(引数たち[0], 手, 所);
      if (t.誤) return t.誤;
      var 小 = '.', 桁 = ',';
      if (引数たち.length > 1) {
        var d = 字の引数(引数たち[1], 手, 所);
        if (d.誤) return d.誤;
        if (!d.字.length) return 誤('#VALUE!');
        小 = d.字.charAt(0);
      }
      if (引数たち.length > 2) {
        var g = 字の引数(引数たち[2], 手, 所);
        if (g.誤) return g.誤;
        if (!g.字.length) return 誤('#VALUE!');
        桁 = g.字.charAt(0);
      }
      if (小 === 桁) return 誤('#VALUE!');
      var s = String(t.字).replace(/^\s+|\s+$/g, '');
      if (!s.length) return 数(0);
      var 割 = 1;
      while (s.length && s.charAt(s.length - 1) === '%') { 割 *= 100; s = s.slice(0, -1); }
      var 出 = '';
      for (var i = 0; i < s.length; i++) {
        var ch = s.charAt(i);
        if (ch === 桁) continue;
        出 += (ch === 小) ? '.' : ch;
      }
      if (!/^[+-]?(\d+(\.\d*)?|\.\d+)([eE][+-]?\d+)?$/.test(出)) return 誤('#VALUE!');
      var v = Number(出);
      if (!isFinite(v)) return 誤('#VALUE!');
      return 数(v / 割);
    },

    /** ★ENCODEURL★ ... URL の 逃がし
        ★紙★ `=ENCODEURL("a b")` → `a%20b`（★＋では なく %20★）
        ★逃がさない 字★ ... A-Z a-z 0-9 - _ . ~
          ＝★JS の `encodeURIComponent` は ! * ' ( ) を 残します★
          ⇒★★だから そのままは 使えません★★（手で 足します）
        ★★未測定★★ ... 非 ASCII（UTF-8 の はず）／サロゲート（絵文字） */
    ENCODEURL: function (引数たち, 手, 所) {
      if (!引数たち.length) return 誤('#VALUE!');
      var t = 字の引数(引数たち[0], 手, 所);
      if (t.誤) return t.誤;
      var s = encodeURIComponent(String(t.字)).replace(/[!'()*]/g, function (ch) {
        return '%' + ch.charCodeAt(0).toString(16).toUpperCase();
      });
      return 計.字(s);
    },

    /* ══ ★★台に 無かった 49個の 続き（2束目）★★ ══（2026-09-18）
       ★紙★ `docs/measured/golden-oddf-to-46ko-2026-09-16.tsv`（★1関数 1行★）
       ★★1行しか 無い ＝ 端は 測って いません★★＝★各関数の 下に 名指しで 書きます★
       ★実Excel の 枠は もう 6秒です★ ⇒★未測定は ★次の 枠で 聞く 物★に 積みます★
         ＝`docs/measured/kansuu46/kyuwakume-no-kiku-koto.md` */

    /** ★FIXED★ ... 数を 字に する（★桁区切り つき★）
        ★紙★ `=FIXED(1234.567,2)` → `1,234.57`
        ★引数★ (数, 小数の 桁＝2, ★区切りを 出さないか★＝FALSE)
        ★桁が 負★ ... ★左へ 丸める★（`=FIXED(1234.567,-2)` は 1,200 の はず）
        ★★未測定★★ ... 負の 数／桁が 負／3つ目が TRUE／桁が 100超 */
    FIXED: function (引数たち, 手, 所) {
      if (!引数たち.length) return 誤('#VALUE!');
      var v = 数の引数(引数たち[0], 手, 所);
      if (v.誤) return v.誤;
      var 桁 = 2;
      if (引数たち.length > 1) {
        var d = 数の引数(引数たち[1], 手, 所);
        if (d.誤) return d.誤;
        桁 = 切り捨て(d.数);
      }
      var 区切らない = false;
      if (引数たち.length > 2) {
        var b = 真偽にする(場.ひとつに(引数たち[2], 所 && 所.今のマス));
        if (b.誤) return b.誤;
        区切らない = b.真;
      }
      var 出 = 桁数で書く(v.数, 桁, !区切らない);
      if (出 === null) return 誤('#VALUE!');
      return 計.字(出);
    },

    /** ★DOLLAR★ ... 数を ★お金の 字★に する
        ★紙★ `=DOLLAR(1234.567,2)` → `¥1,234.57`
        ★★この 家の 通貨の 印は `¥`★★（★紙が そう 言って います★）
          ＝★土地（Windows の 地域）で 変わり得ます★＝★測った のは この 機械だけ★
        ★★未測定★★ ... ★負の 数★（丸括弧か 負号か）／桁が 負／桁を 省いた 時
          ⇒★★負の 数は 9枠目で 聞きます★★（`kyuwakume-no-kiku-koto.md`） */
    DOLLAR: function (引数たち, 手, 所) {
      if (!引数たち.length) return 誤('#VALUE!');
      var v = 数の引数(引数たち[0], 手, 所);
      if (v.誤) return v.誤;
      var 桁 = 2;
      if (引数たち.length > 1) {
        var d = 数の引数(引数たち[1], 手, 所);
        if (d.誤) return d.誤;
        桁 = 切り捨て(d.数);
      }
      var 字 = 桁数で書く(Math.abs(v.数), 桁, true);
      if (字 === null) return 誤('#VALUE!');
      /* == ★★負の 時の 印の 位置★★ ==（2026-09-18・★9枠目 21行目で 実測★）
           ★実Excel★ `=DOLLAR(-1234.567,2)` -> ★円記号 + マイナス + 1,234.57★
           ＝★★円記号が 先／マイナスが 後★★（★丸括弧では ありません★）
           ★私は 逆に 書いて いました★（マイナス + 円記号）
           ⇒★★これは お客さんに 出て いた 欠陥です★★
             ＝`YEN` は `convertFormula` で `DOLLAR` に 変わります
             ＝★日本語で `=YEN(-1234,2)` と 書いた 人に 違う 字が 出て いました★
           ★★FIXED は 逆です★★ ... `=FIXED(-1234.567,2)` -> `-1,234.57`（★負号が 先★）
             ＝★2つを 同じに しては いけません★（★実測で 分かれました★）
           ★まだ 聞いて いない★ ... 通貨の 印は ★土地（Windows の 地域）で 変わり得ます★ */
      return 計.字('\u00a5' + (v.数 < 0 ? '-' : '') + 字);
    },

    /** ★VALUETOTEXT★ ... 値を 字に する
        ★紙★ `=VALUETOTEXT(A1)` → `1`（A1 ＝ 1・型は String）
        ★引数★ (値, 形＝0)
          0 ＝★そのまま★（字は 引用符を 付けない）
          1 ＝★きっちり★（★字は 引用符で 囲む★）
        ★★未測定★★ ... 形 1／真偽／誤り／四角（溢れ） */
    VALUETOTEXT: function (引数たち, 手, 所) {
      if (!引数たち.length) return 誤('#VALUE!');
      var v = 場.ひとつに(引数たち[0], 所 && 所.今のマス);
      if (!v) return 誤('#VALUE!');
      var 形 = 0;
      if (引数たち.length > 1) {
        var f = 数の引数(引数たち[1], 手, 所);
        if (f.誤) return f.誤;
        形 = 切り捨て(f.数);
        if (形 !== 0 && 形 !== 1) return 誤('#VALUE!');
      }
      if (v.型 === '誤') return 計.字(String(v.値));
      if (v.型 === '空') return 計.字('');
      if (v.型 === '字') return 計.字(形 === 1 ? '"' + v.値 + '"' : String(v.値));
      if (v.型 === '真偽') return 計.字(v.値 ? 'TRUE' : 'FALSE');
      var t = 計.字にする(v);
      if (t.型 === '誤') return t;
      return 計.字(String(t.値));
    },

    /** ★PHONETIC★ ... ふりがな
        ★紙★ `=PHONETIC(A1)` → ★空の 字★（型は String）
        ★★この 台は ふりがなを 持って いません★★
          ＝実Excel の ふりがなは ★打った 時の 変換の 履歴★です
          ＝★xlsx の `<rPh>` に 入って います★（★読み込みでは 今 拾って いません★）
        ⇒★★だから 常に 空の 字を 返します★★
        ⇒★★これは「出来た」では ありません★★
          ＝★紙の 1行（空の 字）と 合うだけ★
          ＝★ふりがなの 入った ブックでは ★違う 答えに なります★★
          ⇒`docs/measured/karimono-hazushi-no-tana.md` に 積みます */
    PHONETIC: function (引数たち, 手, 所) {
      if (!引数たち.length) return 誤('#VALUE!');
      /* ══ ★★2026-09-19 ... ★四角を 渡されたら 暗黙の 交わりを しません★★★ ══
           ★前★ ... `場.ひとつに` を 通して いた
                 ⇒`=PHONETIC(A1:A2)` が ★#VALUE!★（交われない ので）
           ★実Excel★ ... ★空の 字★（型 String）
             ＝経営者1 が 実測（2026-09-19・版 16.0 build 20326）
               `=PHONETIC(A1)` 空 ／ `=PHONETIC(A1:A2)` ★空★ ／ `=PHONETIC(B1)` ヤマダ
               （B1 は `Phonetics.Add(1,2,"ヤマダ")` の 後 ★読み返して 付いて いる のを 確かめた★）
           ⇒★★PHONETIC は 範囲を そのまま 受けて ふりがなを 繋ぎます★★
             ＝★1つの 値に する 関数では ありません★＝★交わりは 掛けません★
           ★★これは「出来た」では ありません★★
             ＝この 台は ふりがなを 持って いない ので ★常に 空★
             ＝★ふりがなの 入った ブックでは まだ 違う 答えに なります★
             ⇒`=PHONETIC(B1)` が ヤマダ を 返すには `<rPh>` を 読む 要が 在る
             ⇒★そこは まだ 測れて いません★（画面に ふりがなを 付ける 口が 無い） */
      var a = 引数たち[0];
      if (a && a.種 === '四角') {
        for (var i = 0; i < (a.並び || []).length; i++) {
          var x = a.並び[i];
          if (x && x.型 === '誤') return x;      /* ★誤りは そのまま 通す★ */
        }
        return 計.字('');
      }
      var v = 場.ひとつに(a, 所 && 所.今のマス);
      if (v && v.型 === '誤') return v;
      return 計.字('');
    },

    /** ★LOOKUP★ ... ★昇順の 並びから 探す★（★探す 値 以下で 一番 大きい 物★）
        ★紙★ `=LOOKUP(3,A1:A5,B1:B5)` → 6（A ＝ 1..5 ／ B ＝ 2,4,6,8,10）
        ★2つの 形★
          ㋐並びの 形 ... LOOKUP(値, 探す並び, 返す並び)
          ㋑四角の 形 ... LOOKUP(値, 四角)
              ＝★縦長なら 一番 右の 列／横長なら 一番 下の 行★を 返す
        ★★並びが 昇順で ない 時の 答えは 決まって いません★★（実Excel の 覚書き）
          ⇒★うちは ★二分探索★では なく ★前から 見る★★＝★同じ 形の 答えに なる★
        ★★未測定★★ ... 昇順で ない／字の 並び／見つからない（#N/A の はず）／四角の 形 */
    LOOKUP: function (引数たち, 手, 所) {
      if (引数たち.length < 2) return 誤('#VALUE!');
      var 今 = 所 && 所.今のマス;
      var 値 = 場.ひとつに(引数たち[0], 今);
      if (!値) return 誤('#VALUE!');
      if (値.型 === '誤') return 値;
      var 探 = 表にする(引数たち[1], 所);
      if (!探 || !探.length) return 誤('#N/A');
      var 探並 = [], 返並 = [], r, c;
      if (引数たち.length >= 3) {
        for (r = 0; r < 探.length; r++) for (c = 0; c < 探[r].length; c++) 探並.push(探[r][c]);
        var 返 = 表にする(引数たち[2], 所);
        if (!返) return 誤('#VALUE!');
        for (r = 0; r < 返.length; r++) for (c = 0; c < 返[r].length; c++) 返並.push(返[r][c]);
      } else {
        /* ★四角の 形★ ... 縦長なら 一番 右の 列／横長なら 一番 下の 行 */
        var 行数 = 探.length, 列数 = 探[0].length;
        if (行数 >= 列数) {
          for (r = 0; r < 行数; r++) { 探並.push(探[r][0]); 返並.push(探[r][列数 - 1]); }
        } else {
          for (c = 0; c < 列数; c++) { 探並.push(探[0][c]); 返並.push(探[行数 - 1][c]); }
        }
      }
      if (!探並.length) return 誤('#N/A');
      /* ══ ★★二分探索★★ ══（2026-09-18・9枠目 29行目で 実測）
           ★実Excel★ `=LOOKUP(4,D1:D5)`（D ＝ 9,7,5,3,1＝★降順★）-> ★#N/A★
           ★私は 前から 見て いました★ ⇒ ★1★（＝一番 後ろの 3 の 次）
           ⇒★★「昇順で ない 時の 答えは 決まって いない」では ありません★★
             ＝★二分探索の 動きが そのまま 出ます★＝★決まって います★
           ★だから 前から 見るのを やめて 二分探索に しました★
             ＝★同じ 形の 答えに なる★と 書いて いたのは ★私の 見立て★でした
           ★まだ 聞いて いない★ ... 字の 並び／空マスの 飛ばし方 */
      var 下 = 0, 上 = 探並.length - 1, 当 = -1;
      while (下 <= 上) {
        var 真ん中 = (下 + 上) >> 1;
        var x = 探並[真ん中];
        var く = (!x || x.型 === '空' || x.型 === '誤') ? null : 比べられるか(値, x);
        if (く === null) { 上 = 真ん中 - 1; continue; }   /* ★比べられない＝下へ★ */
        if (く >= 0) { 当 = 真ん中; 下 = 真ん中 + 1; }
        else 上 = 真ん中 - 1;
      }
      if (当 < 0) return 誤('#N/A');
      if (当 >= 返並.length) return 誤('#N/A');
      return 返並[当] || 計.空;
    },

    /* ══ ★★台に 無かった 49個の 続き（3束目）★★ ══（2026-09-18）
       ★★この 2つの 為に 板の 側に 口を 1つ 足しました★★
         `lib/shiki-hyou.js` ... `所.四角を取る(行, 列, 行数, 列数)` と `所.字から取る("A1:B2")`
       ★なぜ★ ... ★式の 中で 場所を 作って 読む★のは この 2つだけ
         ＝引数で 渡された 四角では 届きません
       ★渡したのは 読む 口だけ★（★書く 口は 渡して いません★）
       ★通り道は そのまま★＝★輪（循環）は 今まで通り 見つかります★ */

    /** ★INDIRECT★ ... ★字で 書いた 場所を 読む★
        ★紙★ `=INDIRECT("A1")` → 1
        ★引数★ (字, ★A1 の 書き方か★＝TRUE)
        ★★未測定★★
          ・★R1C1 の 書き方（2つ目が FALSE）★ ... ★受けません（#VALUE!）★
          ・★板を またぐ 時（"Sheet2!A1"）★ ... ★受けません（#REF!）★
          ・★名前（名前の 定義）を 渡した 時★
          ⇒★どれも ★出来た 顔を しません★＝はっきり 誤りを 返します★ */
    INDIRECT: function (引数たち, 手, 所) {
      if (!引数たち.length) return 誤('#VALUE!');
      var t = 字の引数(引数たち[0], 手, 所);
      if (t.誤) return t.誤;
      var 字 = t.字;
      if (引数たち.length > 1) {
        var b = 真偽にする(場.ひとつに(引数たち[1], 所 && 所.今のマス));
        if (b.誤) return b.誤;
        if (!b.真) {
          /* ══ ★★R1C1 の 書き方★★ ══（2026-09-18・9枠目 30行目で 実測）
               ★実Excel★ `=INDIRECT("R1C1",FALSE)` -> ★1★
               ★私は #VALUE! を 返して いました★（★まだ 書いて いなかった★）
               ★R1C1★ ... `R<行>C<列>`（★1から 数えます★）
               ★★まだ 聞いて いない★★
                 ・★角括弧の 相対（`R[1]C[0]`）★ ... ★受けません★（#REF!）
                 ・`R1C1:R2C2` の ような 四角 ... ★受けます★（下で 分けます）
                 ・★R や C を 省いた 形（`R1`・`C1`）★ ... ★受けません★ */
          var 直す = function (s) {
            var m = /^R([0-9]+)C([0-9]+)$/i.exec(String(s).trim());
            if (!m) return null;
            var 行 = Number(m[1]), 列 = Number(m[2]);
            if (!(行 >= 1) || !(列 >= 1)) return null;
            return { 行: 行 - 1, 列: 列 - 1 };
          };
          var 割 = String(字).split(':');
          if (割.length > 2) return 誤('#REF!');
          var a1 = 直す(割[0]);
          if (!a1) return 誤('#REF!');
          var b1 = (割.length === 2) ? 直す(割[1]) : a1;
          if (!b1) return 誤('#REF!');
          if (!所 || typeof 所.四角を取る !== 'function') return 誤('#REF!');
          var r0 = Math.min(a1.行, b1.行), r1 = Math.max(a1.行, b1.行);
          var c0 = Math.min(a1.列, b1.列), c1 = Math.max(a1.列, b1.列);
          var 参2 = 所.四角を取る(r0, c0, r1 - r0 + 1, c1 - c0 + 1);
          if (!参2) return 誤('#REF!');
          return 参2;
        }
      }
      if (!所 || typeof 所.字から取る !== 'function') return 誤('#REF!');
      var 参 = 所.字から取る(字);
      if (!参) return 誤('#REF!');
      /* ★★1マスでも ★参照の まま★ 返します★★（2026-09-18）
           ★訳★ `=AREAS(INDIRECT("A1"))` は ★参照を 要ります★（値では 1 に なりません）
           ★1つの 値が 要る 所では 板が `場.交わる` で 詰めます★＝★二度手間に なりません★ */
      return 参;
    },

    /** ★OFFSET★ ... ★ずらした 場所を 読む★
        ★紙★ `=OFFSET(A1,1,0)` → 2（A2）
        ★引数★ (もと, ★下へ★, ★右へ★, [高さ], [幅])
          ・高さ／幅を 省くと ★もとと 同じ 形★
          ・★高さ／幅が 負★ ... ★上（左）へ 伸びる★
        ★★未測定★★ ... 高さ／幅が 0（#REF! の はず）／板の 外へ 出た 時
          ／もとが 四角の 時の 数え方 */
    OFFSET: function (引数たち, 手, 所) {
      if (引数たち.length < 3) return 誤('#VALUE!');
      if (!所 || typeof 所.四角を取る !== 'function') return 誤('#REF!');
      var もと = 引数たち[0];
      var 行 = null, 列 = null, 行数 = null, 列数 = null;
      if (もと && もと.種 === '四角' && もと.番地) {
        行 = もと.番地.行; 列 = もと.番地.列;
        行数 = もと.行数 || 1; 列数 = もと.列数 || 1;
      } else if (もと && もと.種 === 'マス' && もと.番地) {
        行 = もと.番地.行; 列 = もと.番地.列; 行数 = 1; 列数 = 1;
      } else if (場.参照か(もと && もと.値)) {
        var 参 = もと.値;
        if (参.番地なし) return 誤('#REF!');       /* ★式が 作った 表には 番地が 無い★ */
        行 = 参.行; 列 = 参.列; 行数 = 参.行数; 列数 = 参.列数;
      }
      if (行 === null) return 誤('#REF!');
      var d = 数の引数(引数たち[1], 手, 所);
      if (d.誤) return d.誤;
      var e = 数の引数(引数たち[2], 手, 所);
      if (e.誤) return e.誤;
      var 新行 = 行 + 切り捨て(d.数), 新列 = 列 + 切り捨て(e.数);
      var h = 行数, w = 列数;
      if (引数たち.length > 3) {
        var x = 数の引数(引数たち[3], 手, 所);
        if (x.誤) return x.誤;
        h = 切り捨て(x.数);
      }
      if (引数たち.length > 4) {
        var y = 数の引数(引数たち[4], 手, 所);
        if (y.誤) return y.誤;
        w = 切り捨て(y.数);
      }
      if (h === 0 || w === 0) return 誤('#REF!');
      if (h < 0) { 新行 = 新行 + h + 1; h = -h; }   /* ★負は 上へ 伸びる★ */
      if (w < 0) { 新列 = 新列 + w + 1; w = -w; }
      if (新行 < 0 || 新列 < 0) return 誤('#REF!');
      var 出 = 所.四角を取る(新行, 新列, h, w);
      if (!出) return 誤('#REF!');
      return 出;                                  /* ★1マスでも 参照の まま★（INDIRECT と 同じ） */
    },

    /* ══ ★★ラムダの 一族 6個★★ ══（2026-09-18・段取り ⑶⑷）
       ★LAMBDA と LET は 板（`lib/shiki-hyou.js`）に 在ります★
         ＝★木の まま 受け取る 必要が 在る★ので ここには 置けません
       ★この 6個は 値（ラムダ）を 受け取って 呼ぶだけ★です */

    /** ★MAP★ ... 並びの 1つずつに 当てる
        ★紙★ `=MAP(A1:A5,LAMBDA(x,x*2))` -> 左上 2 ／ 行 5 ／ 列 1 ／ 和 30
        ★★元と 同じ 形で 返します★★
        ★★未測定★★ ... 並びを 2つ 以上 渡した 時（`MAP(a,b,LAMBDA(x,y,..))`） */
    MAP: function (引数たち, 手, 所) {
      if (引数たち.length < 2) return 誤('#VALUE!');
      var g = ラムダの引数(引数たち[引数たち.length - 1], 所);
      if (g.誤) return g.誤;
      var 表たち = [];
      for (var k = 0; k < 引数たち.length - 1; k++) {
        var A = 表にする(引数たち[k], 所);
        if (!A || !A.length) return 誤('#VALUE!');
        表たち.push(A);
      }
      var 行数 = 表たち[0].length, 列数 = 表たち[0][0].length, i;
      for (i = 1; i < 表たち.length; i++) {
        if (表たち[i].length !== 行数 || 表たち[i][0].length !== 列数) return 誤('#VALUE!');
      }
      var 出 = [];
      for (var r = 0; r < 行数; r++) {
        var 段 = [];
        for (var c = 0; c < 列数; c++) {
          var 引値 = 表たち.map(function (A) { return A[r][c] || 計.空; });
          var v = ラムダ1回(所, g.L, 引値);
          if (v && v.型 === '誤' && v.値 === '#VALUE!' && g.L.名たち.length !== 表たち.length) return v;
          段.push(v);
        }
        出.push(段);
      }
      return 溢れに(出);
    },

    /** ★REDUCE★ ... 畳んで 1つに する
        ★紙★ `=REDUCE(0,A1:A5,LAMBDA(a,b,a+b))` -> 15
              `=REDUCE(100,A1:A5,LAMBDA(a,b,a-b))` -> ★85★
              ＝★a が 溜め・b が 次★（★裏取り 済★・100-1-2-3-4-5 ＝ 85）
        ★★未測定★★ ... 並びが 空の 時 */
    REDUCE: function (引数たち, 手, 所) {
      if (引数たち.length < 3) return 誤('#VALUE!');
      var g = ラムダの引数(引数たち[2], 所);
      if (g.誤) return g.誤;
      var 溜 = 場.ひとつに(引数たち[0], 所 && 所.今のマス) || 計.空;
      var A = 表にする(引数たち[1], 所);
      if (!A) return 誤('#VALUE!');
      for (var r = 0; r < A.length; r++) {
        for (var c = 0; c < A[r].length; c++) {
          溜 = ラムダ1回(所, g.L, [溜, A[r][c] || 計.空]);
          if (溜 && 溜.型 === '誤') return 溜;
        }
      }
      return 溜;
    },

    /** ★SCAN★ ... 畳む 途中を 全部 出す
        ★紙★ `=ROWS(SCAN(0,A1:A5,LAMBDA(a,b,a+b)))` -> ★5★
              `=SUM(SCAN(0,A1:A5,LAMBDA(a,b,a+b)))` -> 35（1+3+6+10+15）
        ★★初めの 値は 含まれません★★（★6では なく 5★＝実測で 決まった） */
    SCAN: function (引数たち, 手, 所) {
      if (引数たち.length < 3) return 誤('#VALUE!');
      var g = ラムダの引数(引数たち[2], 所);
      if (g.誤) return g.誤;
      var 溜 = 場.ひとつに(引数たち[0], 所 && 所.今のマス) || 計.空;
      var A = 表にする(引数たち[1], 所);
      if (!A) return 誤('#VALUE!');
      var 出 = [];
      for (var r = 0; r < A.length; r++) {
        var 段 = [];
        for (var c = 0; c < A[r].length; c++) {
          溜 = ラムダ1回(所, g.L, [溜, A[r][c] || 計.空]);
          if (溜 && 溜.型 === '誤') return 溜;
          段.push(溜);
        }
        出.push(段);
      }
      return 溢れに(出);
    },

    /** ★BYROW★ ... 行ごとに 当てる（★縦 1列で 返す★）
        ★紙★ `=BYROW(F1:G2,LAMBDA(r,SUM(r)))` -> ★3★（F1:G2 ＝ 1,2／10,20）
              ＝★行の 和 3 と 30★（★左上で 分かれました★・列なら 11） */
    BYROW: function (引数たち, 手, 所) { return 段ごとに当てる(引数たち, 所, true); },
    /** ★BYCOL★ ... 列ごとに 当てる（★横 1行で 返す★）
        ★紙★ `=BYCOL(F1:G2,LAMBDA(c,SUM(c)))` -> ★11★（＝列の 和 11 と 22） */
    BYCOL: function (引数たち, 手, 所) { return 段ごとに当てる(引数たち, 所, false); },

    /** ★MAKEARRAY★ ... 行数×列数 の 表を 作る
        ★紙★ `=MAKEARRAY(2,3,LAMBDA(r,c,r*c))` -> 左上 1 ／ 和 ★18★
              ＝(1+2)×(1+2+3) ＝ 18 ⇒★★行・列は 1から★★
        ★★未測定★★ ... 行数／列数が 0 か 負（#VALUE! の はず） */
    MAKEARRAY: function (引数たち, 手, 所) {
      if (引数たち.length < 3) return 誤('#VALUE!');
      var 行 = 数の引数(引数たち[0], 手, 所);
      if (行.誤) return 行.誤;
      var 列 = 数の引数(引数たち[1], 手, 所);
      if (列.誤) return 列.誤;
      var g = ラムダの引数(引数たち[2], 所);
      if (g.誤) return g.誤;
      var R = 切り捨て(行.数), C = 切り捨て(列.数);
      if (R < 1 || C < 1) return 誤('#VALUE!');
      if (R * C > 1000000) return 誤('#NUM!');      /* ★作りすぎを 止める★（未測定） */
      var 出 = [];
      for (var r = 1; r <= R; r++) {
        var 段 = [];
        for (var c = 1; c <= C; c++) {
          var v = ラムダ1回(所, g.L, [数(r), 数(c)]);
          if (v && v.型 === '誤') return v;
          段.push(v);
        }
        出.push(段);
      }
      return 溢れに(出);
    },

    /* ★指数分布★ */
    EXPONDIST: function (引数たち, 手, 所) { return 指数の分布(引数たち, 手, 所); },
    'EXPON.DIST': function (引数たち, 手, 所) { return 指数の分布(引数たち, 手, 所); },
    /* ★ポアソン分布★ */
    POISSON: function (引数たち, 手, 所) { return ポアソンの分布(引数たち, 手, 所); },
    'POISSON.DIST': function (引数たち, 手, 所) { return ポアソンの分布(引数たち, 手, 所); },
    /* ★χ²の 左側★ */
    'CHISQ.DIST': function (引数たち, 手, 所) {
      if (引数たち.length < 3) return 誤('#VALUE!');
      var x = 数の引数(引数たち[0], 手, 所);
      if (x.誤) return x.誤;
      var v = 数の引数(引数たち[1], 手, 所);
      if (v.誤) return v.誤;
      var 判 = 真偽にする(場.ひとつに(引数たち[2], 所 && 所.今のマス));
      if (判.誤) return 判.誤;
      var k = 切り捨て(v.数);
      if (x.数 < 0 || k < 1) return 誤('#NUM!');
      if (判.真) return 数(1 - 上のガンマ(k / 2, x.数 / 2));
      return 数(Math.exp((k / 2 - 1) * Math.log(x.数) - x.数 / 2
        - (k / 2) * Math.LN2 - ガンマの対数(k / 2)));
    },
    /* ★t の 左側★ */
    'T.DIST': function (引数たち, 手, 所) {
      if (引数たち.length < 3) return 誤('#VALUE!');
      var t = 数の引数(引数たち[0], 手, 所);
      if (t.誤) return t.誤;
      var v = 数の引数(引数たち[1], 手, 所);
      if (v.誤) return v.誤;
      var 判 = 真偽にする(場.ひとつに(引数たち[2], 所 && 所.今のマス));
      if (判.誤) return 判.誤;
      var k2 = 切り捨て(v.数);
      if (k2 < 1) return 誤('#NUM!');
      if (判.真) return 数(1 - tの右(t.数, k2));
      return 数(Math.exp(ガンマの対数((k2 + 1) / 2) - ガンマの対数(k2 / 2))
        / Math.sqrt(k2 * Math.PI) * Math.pow(1 + t.数 * t.数 / k2, -(k2 + 1) / 2));
    },
    /* ★分数の 値段★ */
    DOLLARDE: function (引数たち, 手, 所) {
      return 数ふたつ(引数たち, 手, 所, function (a, b) {
        var 分 = 切り捨て(b);
        if (分 < 0) return null;
        if (分 === 0) return null;
        if (分 === 1) return a;
        var 桁 = Math.ceil(Math.log(分) / Math.LN10);
        var 整 = 切り捨て(a);
        var 端 = (a - 整) * Math.pow(10, 桁);
        return 整 + 端 / 分;
      });
    },
    DOLLARFR: function (引数たち, 手, 所) {
      return 数ふたつ(引数たち, 手, 所, function (a, b) {
        var 分 = 切り捨て(b);
        if (分 < 0) return null;
        if (分 === 0) return null;
        if (分 === 1) return a;
        var 桁 = Math.ceil(Math.log(分) / Math.LN10);
        var 整 = 切り捨て(a);
        var 端 = (a - 整) * 分;
        return 整 + 端 / Math.pow(10, 桁);
      });
    },
    FVSCHEDULE: function (引数たち, 手, 所) {
      if (引数たち.length < 2) return 誤('#VALUE!');
      var a = 数の引数(引数たち[0], 手, 所);
      if (a.誤) return a.誤;
      var 拾 = 数を拾う([引数たち[1]], 手);
      if (拾.誤) return 拾.誤;
      var 出 = a.数;
      for (var i = 0; i < 拾.数たち.length; i++) 出 *= (1 + 拾.数たち[i]);
      return 数(出);
    },

    /* ══ ★★残りの 営業日と 枚数★★ ══（2026-09-16）
       ★物差し★ `golden-346-2026-09-08.tsv` */
    'NETWORKDAYS.INTL': function (引数たち, 手, 所) { return 営業日の国(引数たち, 手, 所, true); },
    'WORKDAY.INTL': function (引数たち, 手, 所) { return 営業日の国(引数たち, 手, 所, false); },
    /* ★SHEET／SHEETS★ … ★この 台は 板を 1枚と 見ます★ */
    SHEET: function (引数たち) { return 数(1); },
    SHEETS: function (引数たち) { return 数(1); },

    /* ══ ★★利息と 分布の 続き★★ ══（2026-09-16）
       ★物差し★ `golden-346-2026-09-08.tsv` */
    IPMT: function (引数たち, 手, 所) { return 期の内訳(引数たち, 手, 所, true); },
    PPMT: function (引数たち, 手, 所) { return 期の内訳(引数たち, 手, 所, false); },
    /* ★RRI★ … ★何倍に なったかから 利を 逆算★ */
    RRI: function (引数たち, 手, 所) {
      var g = 金の引数(引数たち, 手, 所, 3, 3);
      if (g.誤) return g.誤;
      var n = g.数[0], pv = g.数[1], fv = g.数[2];
      if (n <= 0 || pv === 0) return 誤('#NUM!');
      return 数(Math.pow(fv / pv, 1 / n) - 1);
    },
    COMBINA: function (引数たち, 手, 所) {
      return 数ふたつ(引数たち, 手, 所, function (a, b) {
        var n = 切り捨て(a), k = 切り捨て(b);
        if (n < 0 || k < 0) return null;
        if (n === 0 && k > 0) return null;
        /* ★重複を 許す 選び方★ ＝ C(n+k-1, k) */
        var m = n + k - 1, r = 1;
        for (var i = 1; i <= k; i++) r = r * (m - k + i) / i;
        return Math.round(r);
      });
    },
    GAMMA: function (引数たち, 手, 所) {
      return 数ひとつ(引数たち, 手, 所, function (x) {
        if (x === Math.floor(x) && x <= 0) return null;     /* ★負の 整数は 無い★ */
        if (x > 0) return Math.exp(ガンマの対数(x));
        /* ★反射の 式★ Γ(x)Γ(1-x) = π/sin(πx) */
        return Math.PI / (Math.sin(Math.PI * x) * Math.exp(ガンマの対数(1 - x)));
      });
    },
    /* ★χ²と t の 右側★ */
    CHIDIST: function (引数たち, 手, 所) { return カイ二乗の右(引数たち, 手, 所); },
    'CHISQ.DIST.RT': function (引数たち, 手, 所) { return カイ二乗の右(引数たち, 手, 所); },
    'T.DIST.RT': function (引数たち, 手, 所) {
      var g = 金の引数(引数たち, 手, 所, 2, 2);
      if (g.誤) return g.誤;
      var t = g.数[0], v = 切り捨て(g.数[1]);
      if (v < 1) return 誤('#NUM!');
      return 数(tの右(t, v));
    },
    'T.DIST.2T': function (引数たち, 手, 所) {
      var g = 金の引数(引数たち, 手, 所, 2, 2);
      if (g.誤) return g.誤;
      var t = g.数[0], v = 切り捨て(g.数[1]);
      if (v < 1 || t < 0) return 誤('#NUM!');
      return 数(2 * tの右(t, v));
    },

    /* ══ ★★ふるいと 並べ替えと 探し★★ ══（2026-09-16）
       ★物差し★ `golden-346-2026-09-08.tsv` ほか
       ★実測★ `=FILTER(A1:A5,2)` → ★1★（★残った 先頭★）
             `=SORT(A1:A5,1,-1)` → ★5★ */
    FILTER: function (引数たち, 手, 所) {
      if (引数たち.length < 2) return 誤('#VALUE!');
      var A = 表にする(引数たち[0], 所);
      if (!A || !A.length) return 誤('#VALUE!');
      var 印 = 表にする(引数たち[1], 所);
      if (!印 || !印.length) return 誤('#VALUE!');
      /* ★印は 縦に 並ぶ（行を 選ぶ）か 横に 並ぶ（列を 選ぶ）★ */
      var 縦か = 印.length > 1 || A.length === 1;
      var 出 = [];
      if (縦か) {
        for (var r = 0; r < A.length; r++) {
          var v = (印[Math.min(r, 印.length - 1)] || [])[0];
          if (真と見る(v)) 出.push(A[r].slice());
        }
      } else {
        var 段 = [];
        for (var c = 0; c < A[0].length; c++) {
          var w = (印[0] || [])[Math.min(c, (印[0] || []).length - 1)];
          if (真と見る(w)) 段.push(c);
        }
        for (var r2 = 0; r2 < A.length; r2++) {
          var 行 = [];
          for (var i = 0; i < 段.length; i++) 行.push(A[r2][段[i]]);
          出.push(行);
        }
        if (!段.length) 出 = [];
      }
      if (!出.length || !出[0].length) {
        if (引数たち.length > 2) return 場.ひとつに(引数たち[2], 所 && 所.今のマス) || 計.空;
        return 誤('#CALC!');
      }
      return 溢れに(出);
    },
    SORT: function (引数たち, 手, 所) {
      if (!引数たち.length) return 誤('#VALUE!');
      var A = 表にする(引数たち[0], 所);
      if (!A || !A.length) return 誤('#VALUE!');
      var 鍵 = 1, 向 = 1;
      if (引数たち.length > 1) {
        var k = 数の引数(引数たち[1], 手, 所);
        if (k.誤) return k.誤;
        鍵 = 切り捨て(k.数);
      }
      if (引数たち.length > 2) {
        var d = 数の引数(引数たち[2], 手, 所);
        if (d.誤) return d.誤;
        向 = d.数 < 0 ? -1 : 1;
      }
      if (鍵 < 1 || 鍵 > A[0].length) return 誤('#VALUE!');
      var 並 = A.slice().sort(function (x, y) {
        var d2 = 場.比べる(x[鍵 - 1], y[鍵 - 1]);
        return (d2 === null || d2 === undefined ? 0 : d2) * 向;
      });
      return 溢れに(並);
    },
    UNIQUE: function (引数たち, 手, 所) {
      if (!引数たち.length) return 誤('#VALUE!');
      var A = 表にする(引数たち[0], 所);
      if (!A || !A.length) return 誤('#VALUE!');
      var 見た = {}, 出 = [];
      for (var r = 0; r < A.length; r++) {
        var 鍵 = A[r].map(function (v) { return v ? (v.型 + ':' + v.値) : ''; }).join('\u0001');
        if (見た[鍵]) continue;
        見た[鍵] = true;
        出.push(A[r].slice());
      }
      return 溢れに(出);
    },
    /* ★XLOOKUP★ … ★見つからなければ 3つ目（無ければ #N/A）★ */
    XLOOKUP: function (引数たち, 手, 所) {
      if (引数たち.length < 3) return 誤('#VALUE!');
      var 今 = 所 && 所.今のマス;
      var 探 = 場.ひとつに(引数たち[0], 今);
      if (!探) return 誤('#VALUE!');
      if (探.型 === '誤') return 探;
      /* ★★直に 書いた 数も 1つの 並びと して 受けます★★（2026-09-16 実測）
           `=XLOOKUP(2,3,4,5)` … 実Excel ★5★（2 は 3 の 中に 無い ⇒ 4つ目）
           ⇒★前は 四角しか 受けず ★#VALUE!★（8本 赤）★ */
      var 探す場 = 平らに(表にする(引数たち[1], 所));
      var 戻る場 = 平らに(表にする(引数たち[2], 所));
      if (!探す場 || !戻る場) return 誤('#VALUE!');
      for (var i = 0; i < 探す場.length; i++) {
        if (場.比べる(探す場[i], 探) === 0) {
          return 戻る場[Math.min(i, 戻る場.length - 1)] || 計.空;
        }
      }
      if (引数たち.length > 3) return 場.ひとつに(引数たち[3], 今) || 計.空;
      return 誤('#N/A');
    },
    /* ★HYPERLINK★ … ★見せる 字を 返すだけ★（2つ目が 無ければ 1つ目） */
    HYPERLINK: function (引数たち, 手, 所) {
      if (!引数たち.length) return 誤('#VALUE!');
      var i = 引数たち.length > 1 ? 1 : 0;
      var v = 場.ひとつに(引数たち[i], 所 && 所.今のマス);
      if (!v) return 誤('#VALUE!');
      if (v.型 === '誤') return v;
      return v;
    },

    /* ══ ★★位と 進数と 分布★★ ══（2026-09-16）
       ★物差し★ `golden-346-2026-09-08.tsv`
       ★実測★ `=BASE(2,3,4)` → ★0002★（桁を 揃える）
             `=BITAND(D1,D2)` → ★45252★ */
    VALUE: function (引数たち, 手, 所) {
      if (!引数たち.length) return 誤('#VALUE!');
      var v = 場.ひとつに(引数たち[0], 所 && 所.今のマス);
      if (!v) return 誤('#VALUE!');
      if (v.型 === '誤') return v;
      var n = 計.数にする(v, 手);
      if (n.型 === '誤') return 誤('#VALUE!');
      return 数(n.値);
    },
    /* ★BASE／DECIMAL★ … ★何進数でも★（2〜36） */
    BASE: function (引数たち, 手, 所) {
      if (引数たち.length < 2) return 誤('#VALUE!');
      var a = 数の引数(引数たち[0], 手, 所);
      if (a.誤) return a.誤;
      var b = 数の引数(引数たち[1], 手, 所);
      if (b.誤) return b.誤;
      var n = 切り捨て(a.数), 基 = 切り捨て(b.数);
      if (n < 0 || 基 < 2 || 基 > 36) return 誤('#NUM!');
      var 字 = n.toString(基).toUpperCase();
      if (引数たち.length > 2) {
        var d = 数の引数(引数たち[2], 手, 所);
        if (d.誤) return d.誤;
        var 桁 = 切り捨て(d.数);
        if (桁 < 0) return 誤('#NUM!');
        while (字.length < 桁) 字 = '0' + 字;
      }
      return 計.字(字);
    },
    DECIMAL: function (引数たち, 手, 所) {
      if (引数たち.length < 2) return 誤('#VALUE!');
      var t = 字の引数(引数たち[0], 手, 所);
      if (t.誤) return t.誤;
      var b = 数の引数(引数たち[1], 手, 所);
      if (b.誤) return b.誤;
      var 基数 = 切り捨て(b.数);
      if (基数 < 2 || 基数 > 36) return 誤('#NUM!');
      var x = t.字.trim().toUpperCase();
      if (x === '') return 数(0);
      var n = parseInt(x, 基数);
      if (!isFinite(n)) return 誤('#NUM!');
      return 数(n);
    },
    /* ★位の 一族★ … ★負と 2^48 以上は #NUM!★ */
    BITAND: function (引数たち, 手, 所) {
      return 位で(引数たち, 手, 所, function (a, b) { return a & b; });
    },
    BITOR: function (引数たち, 手, 所) {
      return 位で(引数たち, 手, 所, function (a, b) { return a | b; });
    },
    BITXOR: function (引数たち, 手, 所) {
      return 位で(引数たち, 手, 所, function (a, b) { return a ^ b; });
    },
    BITLSHIFT: function (引数たち, 手, 所) {
      return 位で(引数たち, 手, 所, function (a, b) { return b >= 0 ? a * Math.pow(2, b) : Math.floor(a / Math.pow(2, -b)); }, true);
    },
    BITRSHIFT: function (引数たち, 手, 所) {
      return 位で(引数たち, 手, 所, function (a, b) { return b >= 0 ? Math.floor(a / Math.pow(2, b)) : a * Math.pow(2, -b); }, true);
    },
    /* ★最頑張り★ */
    MODE: function (引数たち, 手) { return 最頑張り(引数たち, 手); },
    'MODE.SNGL': function (引数たち, 手) { return 最頑張り(引数たち, 手); },
    /* ★ガンマ関数の 対数★ … Lanczos */
    GAMMALN: function (引数たち, 手, 所) {
      return 数ひとつ(引数たち, 手, 所, function (x) { return x <= 0 ? null : ガンマの対数(x); });
    },
    'GAMMALN.PRECISE': function (引数たち, 手, 所) {
      return 数ひとつ(引数たち, 手, 所, function (x) { return x <= 0 ? null : ガンマの対数(x); });
    },
    'NORM.S.DIST': function (引数たち, 手, 所) {
      if (引数たち.length < 2) return 誤('#VALUE!');
      var a = 数の引数(引数たち[0], 手, 所);
      if (a.誤) return a.誤;
      var 判 = 真偽にする(場.ひとつに(引数たち[1], 所 && 所.今のマス));
      if (判.誤) return 判.誤;
      if (判.真) return 数((1 + 誤差関数(a.数 / Math.SQRT2)) / 2);
      return 数(Math.exp(-a.数 * a.数 / 2) / Math.sqrt(2 * Math.PI));
    },
    /* ★YEARFRAC★ … 型 0＝30/360（既定）／1＝実日数／2＝act/360／3＝act/365／4＝欧 */
    YEARFRAC: function (引数たち, 手, 所) {
      if (引数たち.length < 2) return 誤('#VALUE!');
      var a = 数の引数(引数たち[0], 手, 所);
      if (a.誤) return a.誤;
      var b = 数の引数(引数たち[1], 手, 所);
      if (b.誤) return b.誤;
      var 型 = 0;
      if (引数たち.length > 2) {
        var k = 数の引数(引数たち[2], 手, 所);
        if (k.誤) return k.誤;
        型 = 切り捨て(k.数);
      }
      if (型 < 0 || 型 > 4) return 誤('#NUM!');
      var 始 = 切り捨て(a.数), 終 = 切り捨て(b.数);
      if (始 > 終) { var w = 始; 始 = 終; 終 = w; }
      if (型 === 0 || 型 === 4) {
        var A = 書.数から日(始), B = 書.数から日(終);
        var d1 = A.d, d2 = B.d;
        if (型 === 4) { if (d1 > 30) d1 = 30; if (d2 > 30) d2 = 30; }
        else { if (d1 === 31) d1 = 30; if (d2 === 31) d2 = (d1 === 30 ? 30 : 31); }
        return 数(((B.y - A.y) * 360 + (B.m - A.m) * 30 + (d2 - d1)) / 360);
      }
      if (型 === 2) return 数((終 - 始) / 360);
      if (型 === 3) return 数((終 - 始) / 365);
      /* ★★型 1（実日数 ／ 実日数）… ★分母を 3つに 使い分けます★★
           ★物差し★ `golden-yearfrac-basis1-2026-09-16.tsv`（★実Excel が 答えた 32組★）
             ★分母は 実Excel に 割らせて 控えました★（日数 ÷ 答え）

           ★★前は いつでも「年の 平均日数」で 割って いました★★（365.5 など）
             ⇒★年を またぐが 1年 以内 の 形が 全部 違って いました★
             ⇒ODDFPRICE の 部品を 測ったら 30本中 ★7本 違った★で 見つかりました
             ⇒★YEARFRAC は お客さんが 直に 打つ 関数★＝★客に 出て いた 欠陥★

           ★①同じ 年★                 … ★その 年の 日数★（365 か 366）
           ★②年を またぐが 暦で 1年 以内★ … ★365 か 366 を 使い分ける★
           ★③1年 超え★                 … ★Y1〜Y2 の 年の 平均日数★

           ★★「暦で 1年 以内」は ★日数では 決まりません★★
             ★測った 証し★
               2012-01-01 → 2013-01-01（★366日★）… 分母 ★366★＝②
               2008-03-01 → 2009-03-02（★366日★）… 分母 ★365.5★＝③（平均）
             ⇒★同じ 366日でも 分かれます★＝★月日で 見ます★
             ⇒`y2 === y1+1` かつ `(m2 < m1) または (m2 === m1 かつ d2 <= d1)`

           ★★②の 分母の 決め方（★測った 順に 見ます★）★★
             ㆁ★始めの 年が うるう年★ かつ ★始 ≦ 2/29★ ⇒ 366
                 （測：2008-02-29→2009-01-15 … 366 ／ 2008-01-15→2009-01-10 … 366）
             ㆂ★終わりの 年が うるう年★ かつ ★終 ≧ 2/29★ ⇒ 366
                 （測：2007-11-20→2008-02-29 … 366 ／ 2007-11-20→2008-02-28 … ★365★）
             ㆃそれ 以外 ⇒ 365
                 （測：2008-10-15→2009-03-01 … 365＝★始は うるう年だが 2/29 より 後★）

           ★★うるう年の 見方は ★本当の 暦★★★（Excel の 通し番号とは 別）
             ★測った 証し★ 1900-01-01 → 1900-12-31
               ★日数は 365★（Excel は 1900-02-29 という 無い 日を 持って いる）
               ★だが 分母は 365★＝★うるう年と 見て いません★ */
      var A1 = 書.数から日(始), B1 = 書.数から日(終);
      var Y1 = A1.y, Y2 = B1.y;
      /* ★暦で 1年 以内か★（★日数では なく 月日で 見る★） */
      var 一年以内 = (Y2 === Y1) ||
        (Y2 === Y1 + 1 && (B1.m < A1.m || (B1.m === A1.m && B1.d <= A1.d)));
      if (!一年以内) {
        /* ★③1年 超え … Y1〜Y2 の 年の 平均日数で 割る★
             （測：2007-01-01→2010-01-01 … ★365.25★＝(365+366+365+365)/4） */
        var 年数 = Y2 - Y1 + 1;
        var 日数 = 日から通し(Y2 + 1, 1, 1) - 日から通し(Y1, 1, 1);
        return 数((終 - 始) / (日数 / 年数));
      }
      if (Y2 === Y1) {
        /* ★①同じ 年 … その 年の 日数★ */
        return 数((終 - 始) / (うるう年か(Y1) ? 366 : 365));
      }
      /* ★②年を またぐが 暦で 1年 以内★ */
      var 分母 = 365;
      if (うるう年か(Y1) && (A1.m < 2 || (A1.m === 2 && A1.d <= 29))) 分母 = 366;
      else if (うるう年か(Y2) && (B1.m > 2 || (B1.m === 2 && B1.d >= 29))) 分母 = 366;
      return 数((終 - 始) / 分母);
    },

    /* ══ ★★参照・分位・営業日★★ ══（2026-09-16）
       ★物差し★ `golden-346-2026-09-08.tsv` ほか
       ★実測★ `=ISREF(A1)` → TRUE ／ `=ISREF(A1&"")` → ★FALSE★
             （★引数が ★指した 場所★ か ★計算した 値★ かを 見ます★） */
    ISREF: function (引数たち) {
      if (!引数たち.length) return 誤('#VALUE!');
      var a = 引数たち[0];
      /* ★★値では なく ★どこから 来たか★ を 見ます★★
           ★この 台は 引数に 「種」を 付けて います★（土台の 決め）
           ⇒★`マス` と `四角` だけが 指した 場所★ */
      return 計.真偽(a.種 === 'マス' || a.種 === '四角');
    },
    /* ★PERCENTILE／QUARTILE★ … ★並べて 間を 埋める★ */
    PERCENTILE: function (引数たち, 手, 所) { return 分位(引数たち, 手, 所, false); },
    'PERCENTILE.INC': function (引数たち, 手, 所) { return 分位(引数たち, 手, 所, false); },
    QUARTILE: function (引数たち, 手, 所) { return 分位(引数たち, 手, 所, true); },
    'QUARTILE.INC': function (引数たち, 手, 所) { return 分位(引数たち, 手, 所, true); },

    /* ══ ★★PERCENTRANK★★ ══（2026-09-18）
       ★前は お客さんの 画面で #NAME? でした★（★台に 無い 59個の うちの 1個★）

       ★★紙（実Excel の 実測・6本）★★
         `docs/measured/golden-kansuu-7kaime-2026-09-18.tsv`（★材料 C1:C5 ＝ 1,3,5,7,9★）
         ★★6本 とも 見込み どおりでした★★（★当て推量 0★）
           =PERCENTRANK(C,5)    → ★0.5★    （真ん中）
           =PERCENTRANK(C,1)    → ★0★      （一番 下）
           =PERCENTRANK(C,9)    → ★1★      （一番 上）
           =PERCENTRANK(C,4)    → ★0.375★  （★間は 線で 埋める★）
           =PERCENTRANK(C,4,5)  → ★0.375★  （有効桁 5 でも 同じ）
           =PERCENTRANK(C,10)   → ★#N/A★   （★範囲の 外★）

       ★★中の 計算★★
         並べて、★x 以下で 一番 大きい 物★の 番（0から）を i、
         その 次との 間を 線で 埋めて
           順位 ＝ (i ＋ (x − a[i]) ÷ (a[i+1] − a[i])) ÷ (n − 1)
         ★確かめ★ x=4 … (1 ＋ (4−3)/(5−3)) ÷ 4 ＝ 1.5 ÷ 4 ＝ ★0.375★ ⇒★紙と 同じ★

       ★★測って いない 事（★名指しで 紙に★）★★
         ・★有効桁の 切り方★ … ★切り捨てに して います★（★丸めでは ない★）
             `docs/measured/kansuu46/percentrank-kiku-koto.md`
             ⇒★0.375 は 3桁でも 5桁でも 同じ＝★切り方が 出る 組では ありません★★
         ・★有効桁 0 や 負★／★x が 一番 下より 小さい★／★同じ 数が 並ぶ★
         ・★並びが 昇順で ない★／★字・空が 混ざる★ */
    PERCENTRANK: function (引数たち, 手, 所) {
      if (引数たち.length < 2) return 誤('#VALUE!');
      var 拾 = 数を拾う([引数たち[0]], 手);
      if (拾.誤) return 拾.誤;
      var t = 数の引数(引数たち[1], 手, 所);
      if (t.誤) return t.誤;
      var 桁 = 3;
      if (引数たち.length > 2) {
        var s = 数の引数(引数たち[2], 手, 所);
        if (s.誤) return s.誤;
        桁 = 切り捨て(s.数);
        if (桁 < 1) return 誤('#NUM!');   /* ★測って いません★（Excel の 覚書きに よる） */
      }
      var a = 拾.数たち.slice().sort(function (x, y) { return x - y; });
      var n = a.length;
      if (!n) return 誤('#NUM!');
      var x = t.数;
      /* ★範囲の 外は #N/A★（★実測 `=PERCENTRANK(C1:C5,10)` → #N/A★） */
      if (x < a[0] || x > a[n - 1]) return 誤('#N/A');
      if (n === 1) return 数(x === a[0] ? 0 : 0);
      /* ══ ★★同じ 値は ★下の 位★★★ ══（2026-09-18・実測で 直した）
           ★実Excel★ `=PERCENTRANK({1;3;3;5},3)` -> ★0.333★
           ★前★ `<=` で 進めて いた ⇒ i が 2 に なり 0.666
           ★後★ `<` で 止める ⇒ i が 0 の まま ⇒ (0 + (3-1)/2)/3 = 0.333
           ★他の 行も 確かめました★ 2.6,1 -> 0.5 ／ 3.9,2 -> 0.96 ／ {5;3;1},3 -> 0.5 */
      var i = 0;
      while (i + 1 < n && a[i + 1] < x) i++;
      var 順;
      if (a[i] === x || i + 1 >= n) {
        順 = i / (n - 1);
      } else {
        var 幅 = a[i + 1] - a[i];
        順 = (i + (幅 === 0 ? 0 : (x - a[i]) / 幅)) / (n - 1);
      }
      /* ★有効桁で 切り捨て★（★丸めでは ない＝Excel の 覚書き／★測って いません★） */
      var 倍 = Math.pow(10, 桁);
      return 数(Math.floor(順 * 倍 + 1e-9) / 倍);
    },
    /* ★NETWORKDAYS／WORKDAY★ … ★土日を 抜く★（休みの 一覧も 受ける） */
    NETWORKDAYS: function (引数たち, 手, 所) {
      if (引数たち.length < 2) return 誤('#VALUE!');
      var a = 数の引数(引数たち[0], 手, 所);
      if (a.誤) return a.誤;
      var b = 数の引数(引数たち[1], 手, 所);
      if (b.誤) return b.誤;
      var 休 = 休みを拾う(引数たち, 2, 手);
      if (休.誤) return 休.誤;
      var 始 = 切り捨て(a.数), 終 = 切り捨て(b.数);
      var 符 = 1;
      if (始 > 終) { var w = 始; 始 = 終; 終 = w; 符 = -1; }
      var n = 0;
      for (var d = 始; d <= 終; d++) {
        var 曜 = 書.数から日(d).w;
        if (曜 === 0 || 曜 === 6) continue;
        if (休.日.indexOf(d) >= 0) continue;
        n++;
      }
      return 数(符 * n);
    },
    WORKDAY: function (引数たち, 手, 所) {
      if (引数たち.length < 2) return 誤('#VALUE!');
      var a = 数の引数(引数たち[0], 手, 所);
      if (a.誤) return a.誤;
      var b = 数の引数(引数たち[1], 手, 所);
      if (b.誤) return b.誤;
      var 休 = 休みを拾う(引数たち, 2, 手);
      if (休.誤) return 休.誤;
      var d = 切り捨て(a.数), 残 = 切り捨て(b.数);
      var 向 = 残 < 0 ? -1 : 1;
      残 = Math.abs(残);
      while (残 > 0) {
        d += 向;
        if (d < 0) return 誤('#NUM!');
        var 曜 = 書.数から日(d).w;
        if (曜 === 0 || 曜 === 6) continue;
        if (休.日.indexOf(d) >= 0) continue;
        残--;
      }
      return 数(d);
    },
    /* ★DATEDIF★ … "D" 日／"M" 月／"Y" 年 */
    DATEDIF: function (引数たち, 手, 所) {
      if (引数たち.length < 3) return 誤('#VALUE!');
      var a = 数の引数(引数たち[0], 手, 所);
      if (a.誤) return a.誤;
      var b = 数の引数(引数たち[1], 手, 所);
      if (b.誤) return b.誤;
      var t = 字の引数(引数たち[2], 手, 所);
      if (t.誤) return t.誤;
      var 始 = 切り捨て(a.数), 終 = 切り捨て(b.数);
      if (終 < 始) return 誤('#NUM!');
      var A = 書.数から日(始), B = 書.数から日(終);
      var 種 = t.字.toUpperCase();
      if (種 === 'D') return 数(終 - 始);
      var 月 = (B.y - A.y) * 12 + (B.m - A.m);
      if (B.d < A.d) 月--;
      if (種 === 'M') return 数(月);
      if (種 === 'Y') return 数(Math.floor(月 / 12));
      if (種 === 'MD') return 数(B.d >= A.d ? B.d - A.d : 終 - 日から通し(B.y, B.m - 1, A.d));
      if (種 === 'YM') return 数(月 % 12);
      if (種 === 'YD') return 数(終 - 日から通し(A.y + Math.floor(月 / 12), A.m, A.d));
      return 誤('#NUM!');
    },

    /* ══ ★★番地と 週と ローマ数字★★ ══（2026-09-16）
       ★物差し★ `golden-346-2026-09-08.tsv`
       ★実測★ `=ADDRESS(2,3)` → ★$C$2★ ／ `=ADDRESS(2,3,4)` → ★C2★
             `=ADDRESS(1,2,3,4,5)` → ★'5'!$B1★（★板の 名は 引用符で 囲む★） */
    ADDRESS: function (引数たち, 手, 所) {
      if (引数たち.length < 2) return 誤('#VALUE!');
      var r = 数の引数(引数たち[0], 手, 所);
      if (r.誤) return r.誤;
      var c = 数の引数(引数たち[1], 手, 所);
      if (c.誤) return c.誤;
      var 行 = 切り捨て(r.数), 列 = 切り捨て(c.数);
      if (行 < 1 || 列 < 1) return 誤('#VALUE!');
      var 型 = 1;
      if (引数たち.length > 2) {
        var t = 数の引数(引数たち[2], 手, 所);
        if (t.誤) return t.誤;
        型 = 切り捨て(t.数);
        if (型 < 1 || 型 > 4) return 誤('#VALUE!');
      }
      /* ★4つ目（A1 か R1C1 か）は 読むだけ★＝実測では 答えが 変わりません */
      if (引数たち.length > 3) {
        var y = 真偽にする(場.ひとつに(引数たち[3], 所 && 所.今のマス));
        if (y.誤) return y.誤;
      }
      var 字 = '', x = 列;
      while (x > 0) { var z = (x - 1) % 26; 字 = String.fromCharCode(65 + z) + 字; x = Math.floor((x - 1) / 26); }
      var 列錠 = (型 === 1 || 型 === 3) ? '$' : '';
      var 行錠 = (型 === 1 || 型 === 2) ? '$' : '';
      var 番 = 列錠 + 字 + 行錠 + 行;
      if (引数たち.length > 4) {
        var n = 字の引数(引数たち[4], 手, 所);
        if (n.誤) return n.誤;
        if (n.字 !== '') return 計.字("'" + n.字 + "'!" + 番);
      }
      return 計.字(番);
    },
    /* ★ISOWEEKNUM★ … ★月曜始まり／木曜を 含む 週が 第1週★ */
    ISOWEEKNUM: function (引数たち, 手, 所) {
      var t = 数の引数(引数たち[0], 手, 所);
      if (!引数たち.length) return 誤('#VALUE!');
      if (t.誤) return t.誤;
      if (t.数 < 0) return 誤('#NUM!');
      /* ★★全部 通し番号で 数えます★★（2026-09-16 実測で 直した）
           ★前は 実の 暦（`Date.UTC`）と 通し番号を 混ぜて いました★
           ⇒`=ISOWEEKNUM(1)`（＝1900-01-01）… うち ★1★ ／実Excel ★52★
             ＝★実Excel の 1900-01-01 は ★日曜★★（実の 暦では 月曜）
             ＝★実Excel の 曜日を 使わないと 合いません★ */
      var 通 = 切り捨て(t.数);
      var 曜 = (書.数から日(通).w + 6) % 7;                /* 0＝月曜（★書式の 台の 曜日★） */
      var 木 = 通 + (3 - 曜);
      var 年 = 書.数から日(木).y;
      var 頂 = 日から通し(年, 1, 4);
      var 頂曜 = (書.数から日(頂).w + 6) % 7;
      var 頂木 = 頂 + (3 - 頂曜);
      return 数(Math.round((木 - 頂木) / 7) + 1);
    },
    /* ★WEEKNUM★ … ┅1月 1日を 含む 週が 第1週★（型 1＝日曜始まり／2＝月曜） */
    WEEKNUM: function (引数たち, 手, 所) {
      var t = 数の引数(引数たち[0], 手, 所);
      if (!引数たち.length) return 誤('#VALUE!');
      if (t.誤) return t.誤;
      if (t.数 < 0) return 誤('#NUM!');
      var 型 = 1;
      if (引数たち.length > 1) {
        var k = 数の引数(引数たち[1], 手, 所);
        if (k.誤) return k.誤;
        型 = 切り捨て(k.数);
      }
      if (型 === 21) return 表.ISOWEEKNUM([引数たち[0]], 手, 所);
      var 頂の曜 = { 1: 0, 2: 1, 11: 1, 12: 2, 13: 3, 14: 4, 15: 5, 16: 6, 17: 0 }[型];
      if (頂の曜 === undefined) return 誤('#NUM!');
      /* ★★通し 0 は 0★★（実測 `=WEEKNUM(0.5)` → ★0★）
           ＝★1900年 1月 0日＝年に 入って いない★ */
      var 通2 = 切り捨て(t.数);
      if (通2 === 0) return 数(0);
      /* ★ここも 全部 通し番号で★（ISOWEEKNUM と 同じ 訳） */
      var y = 書.数から日(通2).y;
      var 元日 = 日から通し(y, 1, 1);
      var 元日の曜 = 書.数から日(元日).w;
      var ずれ = (元日の曜 - 頂の曜 + 7) % 7;
      return 数(Math.floor((通2 - 元日 + ずれ) / 7) + 1);
    },
    /* ★ROMAN／ARABIC★ */
    ROMAN: function (引数たち, 手, 所) {
      var t = 数の引数(引数たち[0], 手, 所);
      if (!引数たち.length) return 誤('#VALUE!');
      if (t.誤) return t.誤;
      var n = 切り捨て(t.数);
      if (n < 0 || n > 3999) return 誤('#VALUE!');
      if (n === 0) return 計.字('');
      var 表の = [[1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'],
        [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I']];
      var 出 = '';
      for (var i = 0; i < 表の.length; i++) {
        while (n >= 表の[i][0]) { 出 += 表の[i][1]; n -= 表の[i][0]; }
      }
      return 計.字(出);
    },
    ARABIC: function (引数たち, 手, 所) {
      var t = 字の引数(引数たち[0], 手, 所);
      if (!引数たち.length) return 誤('#VALUE!');
      if (t.誤) return t.誤;
      var x = t.字.trim().toUpperCase();
      var 符 = 1;
      if (x[0] === '-') { 符 = -1; x = x.slice(1); }
      if (x === '') return 数(0);
      if (!/^[MDCLXVI]+$/.test(x)) return 誤('#VALUE!');
      var 値 = { M: 1000, D: 500, C: 100, L: 50, X: 10, V: 5, I: 1 };
      var 出 = 0;
      for (var i = 0; i < x.length; i++) {
        var v = 値[x[i]], w = 値[x[i + 1]] || 0;
        出 += (v < w) ? -v : v;
      }
      return 数(符 * 出);
    },

    /* ══ ★★技術の 一族（進数・誤差関数・残りの 三角）★★ ══（2026-09-16）
       ★物差し★ `golden-346-2026-09-08.tsv`
       ★実測★ `=DEC2HEX(-2)` → ★FFFFFFFFFE★（★10桁の 補数★）
             `=DEC2HEX(2,3)` → ★002★（桁を 揃える） */
    MOD: function (引数たち, 手, 所) {
      return 数ふたつ(引数たち, 手, 所, function (a, b) {
        if (b === 0) return null;
        /* ★余りの 符号は ★割る 方★に 合わせる★（JS の % とは 違う） */
        return a - Math.floor(a / b) * b;
      });
    },
    DELTA: function (引数たち, 手, 所) {
      if (!引数たち.length) return 誤('#VALUE!');
      var a = 数の引数(引数たち[0], 手, 所);
      if (a.誤) return a.誤;
      var b = { 数: 0 };
      if (引数たち.length > 1) { b = 数の引数(引数たち[1], 手, 所); if (b.誤) return b.誤; }
      return 数(a.数 === b.数 ? 1 : 0);
    },
    GESTEP: function (引数たち, 手, 所) {
      if (!引数たち.length) return 誤('#VALUE!');
      var a = 数の引数(引数たち[0], 手, 所);
      if (a.誤) return a.誤;
      var b = { 数: 0 };
      if (引数たち.length > 1) { b = 数の引数(引数たち[1], 手, 所); if (b.誤) return b.誤; }
      return 数(a.数 >= b.数 ? 1 : 0);
    },
    ACOT: function (引数たち, 手, 所) {
      return 数ひとつ(引数たち, 手, 所, function (x) { return Math.PI / 2 - Math.atan(x); });
    },
    ACOTH: function (引数たち, 手, 所) {
      return 数ひとつ(引数たち, 手, 所, function (x) {
        return Math.abs(x) <= 1 ? null : Math.atanh(1 / x);
      });
    },
    COTH: function (引数たち, 手, 所) {
      return 数ひとつ(引数たち, 手, 所, function (x) { return x === 0 ? null : 1 / Math.tanh(x); });
    },
    SECH: function (引数たち, 手, 所) {
      return 数ひとつ(引数たち, 手, 所, function (x) { return 1 / Math.cosh(x); });
    },
    CSCH: function (引数たち, 手, 所) {
      return 数ひとつ(引数たち, 手, 所, function (x) { return x === 0 ? null : 1 / Math.sinh(x); });
    },
    SQRTPI: function (引数たち, 手, 所) {
      return 数ひとつ(引数たち, 手, 所, function (x) { return x < 0 ? null : Math.sqrt(x * Math.PI); });
    },
    FISHER: function (引数たち, 手, 所) {
      return 数ひとつ(引数たち, 手, 所, function (x) {
        return (x <= -1 || x >= 1) ? null : Math.log((1 + x) / (1 - x)) / 2;
      });
    },
    FISHERINV: function (引数たち, 手, 所) {
      return 数ひとつ(引数たち, 手, 所, function (x) {
        /* ★★`e` が 無限に なると `NaN`★★（2026-09-16 実測）
             `=FISHERINV("2024-01-15")`（45306）… `exp(90612)` は ★無限★
             ⇒ `(Inf-1)/(Inf+1)` ⇒ ★`NaN`★／実Excel ★1★
           ★出口の 門が `#NUM!` に して 止めました★が ★正しい 答えは 極限の 1★
           ⇒★門は「悪い 物を 外に 出さない」だけ／★正しい 答えは 自分で 書く★★ */
        var e = Math.exp(2 * x);
        if (!isFinite(e)) return x > 0 ? 1 : -1;
        return (e - 1) / (e + 1);
      });
    },
    /* ★誤差関数★ … ★級数と 連分数で 出します★ */
    ERF: function (引数たち, 手, 所) {
      if (!引数たち.length) return 誤('#VALUE!');
      var a = 数の引数(引数たち[0], 手, 所);
      if (a.誤) return a.誤;
      if (引数たち.length > 1) {
        var b = 数の引数(引数たち[1], 手, 所);
        if (b.誤) return b.誤;
        return 数(誤差関数(b.数) - 誤差関数(a.数));
      }
      return 数(誤差関数(a.数));
    },
    /* ★★`1 - 誤差関数(x)` だと 尾っぽで 桁が 落ちます★★（2026-09-16）
         ⇒★補誤差関数を 直に 呼ぶ★（0.7 から 上は 連分数 300項） */
    ERFC: function (引数たち, 手, 所) {
      return 数ひとつ(引数たち, 手, 所, 補誤差関数);
    },
    'ERF.PRECISE': function (引数たち, 手, 所) {
      return 数ひとつ(引数たち, 手, 所, 誤差関数);
    },
    'ERFC.PRECISE': function (引数たち, 手, 所) {
      return 数ひとつ(引数たち, 手, 所, 補誤差関数);
    },
    /* ★正規分布★ */
    PHI: function (引数たち, 手, 所) {
      return 数ひとつ(引数たち, 手, 所, function (x) {
        return Math.exp(-x * x / 2) / Math.sqrt(2 * Math.PI);
      });
    },
    GAUSS: function (引数たち, 手, 所) {
      return 数ひとつ(引数たち, 手, 所, function (x) {
        return 誤差関数(x / Math.SQRT2) / 2;
      });
    },
    NORMSDIST: function (引数たち, 手, 所) {
      return 数ひとつ(引数たち, 手, 所, function (x) {
        return (1 + 誤差関数(x / Math.SQRT2)) / 2;
      });
    },
    /* ★進数★ … ★負は 10桁の 補数★（実測 `=DEC2HEX(-2)` → FFFFFFFFFE） */
    DEC2HEX: function (引数たち, 手, 所) { return 十進から(引数たち, 手, 所, 16); },
    DEC2BIN: function (引数たち, 手, 所) { return 十進から(引数たち, 手, 所, 2); },
    DEC2OCT: function (引数たち, 手, 所) { return 十進から(引数たち, 手, 所, 8); },
    HEX2DEC: function (引数たち, 手, 所) { return 十進へ(引数たち, 手, 所, 16); },
    BIN2DEC: function (引数たち, 手, 所) { return 十進へ(引数たち, 手, 所, 2); },
    OCT2DEC: function (引数たち, 手, 所) { return 十進へ(引数たち, 手, 所, 8); },

    /* ══ ★★複素数（IM系）の 一族★★ ══（2026-09-16）
       ★★ここは 借り物を 継いで いた 所です★★
         `lib/formula-complex-plug.js` は `getFunctionPlugin` で
         ★借り物の 計算を そのまま 継ぎ、出た 字を 15桁に 丸めるだけ★でした（棚㎉）
         ⇒★★中身が 1行も 自前では ない★★＝ここを 書くのが 借り物外し
       ★物差し★ `golden-346-2026-09-08.tsv`（実 Excel）
       ★出す 字の 決め★ … ★実 Excel は 15桁に 丸めて 字に します★ */
    COMPLEX: function (引数たち, 手, 所) {
      if (引数たち.length < 2) return 誤('#VALUE!');
      var a = 数の引数(引数たち[0], 手, 所);
      if (a.誤) return a.誤;
      var b = 数の引数(引数たち[1], 手, 所);
      if (b.誤) return b.誤;
      var 印 = 'i';
      if (引数たち.length > 2) {
        var t = 字の引数(引数たち[2], 手, 所);
        if (t.誤) return t.誤;
        印 = t.字;
        if (印 !== 'i' && 印 !== 'j') return 誤('#VALUE!');
      }
      return 計.字(複素を字に(a.数, b.数, 印));
    },
    IMREAL: function (引数たち, 手, 所) { return 複素ひとつ(引数たち, 手, 所, function (z) { return 数(z.re); }); },
    IMAGINARY: function (引数たち, 手, 所) { return 複素ひとつ(引数たち, 手, 所, function (z) { return 数(z.im); }); },
    IMABS: function (引数たち, 手, 所) {
      return 複素ひとつ(引数たち, 手, 所, function (z) { return 数(Math.sqrt(z.re * z.re + z.im * z.im)); });
    },
    IMARGUMENT: function (引数たち, 手, 所) {
      return 複素ひとつ(引数たち, 手, 所, function (z) {
        if (z.re === 0 && z.im === 0) return 誤('#DIV/0!');
        return 数(Math.atan2(z.im, z.re));
      });
    },
    IMCONJUGATE: function (引数たち, 手, 所) {
      return 複素で(引数たち, 手, 所, function (z) { return { re: z.re, im: -z.im }; });
    },
    IMEXP: function (引数たち, 手, 所) {
      return 複素で(引数たち, 手, 所, function (z) {
        var e = Math.exp(z.re);
        return { re: e * Math.cos(z.im), im: e * Math.sin(z.im) };
      });
    },
    IMLN: function (引数たち, 手, 所) { return 複素で(引数たち, 手, 所, 複素の対数); },
    IMLOG10: function (引数たち, 手, 所) {
      return 複素で(引数たち, 手, 所, function (z) {
        var L = 複素の対数(z);
        return L ? 複素を割る(L, { re: Math.LN10, im: 0 }) : null;
      });
    },
    IMLOG2: function (引数たち, 手, 所) {
      return 複素で(引数たち, 手, 所, function (z) {
        var L = 複素の対数(z);
        return L ? 複素を割る(L, { re: Math.LN2, im: 0 }) : null;
      });
    },
    IMSQRT: function (引数たち, 手, 所) {
      return 複素で(引数たち, 手, 所, function (z) {
        var r = Math.sqrt(Math.sqrt(z.re * z.re + z.im * z.im));
        var w = Math.atan2(z.im, z.re) / 2;
        return { re: r * Math.cos(w), im: r * Math.sin(w) };
      });
    },
    IMSIN: function (引数たち, 手, 所) {
      return 複素で(引数たち, 手, 所, function (z) {
        return { re: Math.sin(z.re) * Math.cosh(z.im), im: Math.cos(z.re) * Math.sinh(z.im) };
      });
    },
    IMCOS: function (引数たち, 手, 所) {
      return 複素で(引数たち, 手, 所, function (z) {
        return { re: Math.cos(z.re) * Math.cosh(z.im), im: -Math.sin(z.re) * Math.sinh(z.im) };
      });
    },
    IMSINH: function (引数たち, 手, 所) {
      return 複素で(引数たち, 手, 所, function (z) {
        return { re: Math.sinh(z.re) * Math.cos(z.im), im: Math.cosh(z.re) * Math.sin(z.im) };
      });
    },
    IMCOSH: function (引数たち, 手, 所) {
      return 複素で(引数たち, 手, 所, function (z) {
        return { re: Math.cosh(z.re) * Math.cos(z.im), im: Math.sinh(z.re) * Math.sin(z.im) };
      });
    },
    IMTAN: function (引数たち, 手, 所) { return 複素で(引数たち, 手, 所, 複素の正接); },
    IMCOT: function (引数たち, 手, 所) {
      return 複素で(引数たち, 手, 所, function (z) {
        var t = 複素の正接(z);
        return t ? 複素を割る({ re: 1, im: 0 }, t) : null;
      });
    },
    IMSEC: function (引数たち, 手, 所) {
      return 複素で(引数たち, 手, 所, function (z) {
        return 複素を割る({ re: 1, im: 0 }, { re: Math.cos(z.re) * Math.cosh(z.im), im: -Math.sin(z.re) * Math.sinh(z.im) });
      });
    },
    IMCSC: function (引数たち, 手, 所) {
      return 複素で(引数たち, 手, 所, function (z) {
        return 複素を割る({ re: 1, im: 0 }, { re: Math.sin(z.re) * Math.cosh(z.im), im: Math.cos(z.re) * Math.sinh(z.im) });
      });
    },
    IMSECH: function (引数たち, 手, 所) {
      return 複素で(引数たち, 手, 所, function (z) {
        return 複素を割る({ re: 1, im: 0 }, { re: Math.cosh(z.re) * Math.cos(z.im), im: Math.sinh(z.re) * Math.sin(z.im) });
      });
    },
    IMCSCH: function (引数たち, 手, 所) {
      return 複素で(引数たち, 手, 所, function (z) {
        return 複素を割る({ re: 1, im: 0 }, { re: Math.sinh(z.re) * Math.cos(z.im), im: Math.cosh(z.re) * Math.sin(z.im) });
      });
    },
    IMSUM: function (引数たち, 手, 所) { return 複素を重ねる(引数たち, 手, 所, 'SUM'); },
    IMPRODUCT: function (引数たち, 手, 所) { return 複素を重ねる(引数たち, 手, 所, 'PROD'); },
    IMSUB: function (引数たち, 手, 所) {
      return 複素ふたつ(引数たち, 手, 所, function (a, b) { return { re: a.re - b.re, im: a.im - b.im }; });
    },
    IMDIV: function (引数たち, 手, 所) { return 複素ふたつ(引数たち, 手, 所, 複素を割る); },
    IMPOWER: function (引数たち, 手, 所) {
      if (引数たち.length < 2) return 誤('#VALUE!');
      var z = 複素にする(引数たち[0], 手, 所);
      if (z.誤) return z.誤;
      var n = 数の引数(引数たち[1], 手, 所);
      if (n.誤) return n.誤;
      var r = Math.sqrt(z.z.re * z.z.re + z.z.im * z.z.im);
      if (r === 0) return 計.字(複素を字に(0, 0, z.印));
      var w = Math.atan2(z.z.im, z.z.re);
      var R = Math.pow(r, n.数), W = w * n.数;
      return 計.字(複素を字に(R * Math.cos(W), R * Math.sin(W), z.印));
    },

    /* ══ ★★日付の 一族★★ ══（2026-09-16）
       ★通し番号 → 年月日は ★書式の 台★を 呼びます★（`書.数から日()`）
         ＝★1900年の 起点が 2つ 在る 話も そこに 在ります★（★作る道を 2本に しない★）
       ★物差し★ `golden-346-2026-09-08.tsv`
       ★実測★ `=DAY(0.5)` → ★0★（通し 0＝1900年1月0日）／`=DAY(TRUE)` → 1 */
    DAY: function (引数たち, 手, 所) { return 日から(引数たち, 手, 所, function (o) { return o.d; }); },
    MONTH: function (引数たち, 手, 所) { return 日から(引数たち, 手, 所, function (o) { return o.m; }); },
    YEAR: function (引数たち, 手, 所) { return 日から(引数たち, 手, 所, function (o) { return o.y; }); },
    HOUR: function (引数たち, 手, 所) { return 日から(引数たち, 手, 所, function (o) { return o.h; }); },
    MINUTE: function (引数たち, 手, 所) { return 日から(引数たち, 手, 所, function (o) { return o.mi; }); },
    SECOND: function (引数たち, 手, 所) { return 日から(引数たち, 手, 所, function (o) { return o.s; }); },
    /* ★WEEKDAY★ … 型 1（既定）＝日曜 1／型 2＝月曜 1／型 3＝月曜 0 */
    WEEKDAY: function (引数たち, 手, 所) {
      var t = 数の引数(引数たち[0], 手, 所);
      if (!引数たち.length || t.誤) return t.誤 || 誤('#VALUE!');
      if (t.数 < 0) return 誤('#NUM!');
      var 型 = 1;
      if (引数たち.length > 1) {
        var k = 数の引数(引数たち[1], 手, 所);
        if (k.誤) return k.誤;
        型 = 切り捨て(k.数);
      }
      var w = 書.数から日(切り捨て(t.数)).w;          /* 0＝日曜 */
      if (型 === 1 || 型 === 17) return 数(w + 1);
      if (型 === 2 || 型 === 11) return 数((w + 6) % 7 + 1);
      if (型 === 3) return 数((w + 6) % 7);
      if (型 >= 12 && 型 <= 16) return 数((w - (型 - 10) + 7) % 7 + 1);
      return 誤('#NUM!');
    },
    DAYS: function (引数たち, 手, 所) {
      var g = 金の引数(引数たち, 手, 所, 2, 2);
      if (g.誤) return g.誤;
      return 数(切り捨て(g.数[0]) - 切り捨て(g.数[1]));
    },
    /* ★DAYS360★ … ★1年を 360日（1か月 30日）と 見る★／型 true＝ヨーロッパ式 */
    DAYS360: function (引数たち, 手, 所) {
      if (引数たち.length < 2) return 誤('#VALUE!');
      var a = 数の引数(引数たち[0], 手, 所);
      if (a.誤) return a.誤;
      var b = 数の引数(引数たち[1], 手, 所);
      if (b.誤) return b.誤;
      var 欧 = false;
      if (引数たち.length > 2) {
        var 判 = 真偽にする(場.ひとつに(引数たち[2], 所 && 所.今のマス));
        if (判.誤) return 判.誤;
        欧 = 判.真;
      }
      var A = 書.数から日(切り捨て(a.数)), B = 書.数から日(切り捨て(b.数));
      var d1 = A.d, d2 = B.d;
      if (欧) { if (d1 > 30) d1 = 30; if (d2 > 30) d2 = 30; }
      else {
        if (d1 === 31) d1 = 30;
        if (d2 === 31) d2 = (d1 === 30 ? 30 : 1);
        if (d2 === 1 && B.d === 31) { B = { y: B.y, m: B.m + 1, d: 1 }; }
      }
      return 数((B.y - A.y) * 360 + (B.m - A.m) * 30 + (d2 - d1));
    },
    EDATE: function (引数たち, 手, 所) { return 月をずらす(引数たち, 手, 所, false); },
    EOMONTH: function (引数たち, 手, 所) { return 月をずらす(引数たち, 手, 所, true); },
    TIME: function (引数たち, 手, 所) {
      var g = 金の引数(引数たち, 手, 所, 3, 3);
      if (g.誤) return g.誤;
      var h = 切り捨て(g.数[0]), m = 切り捨て(g.数[1]), s2 = 切り捨て(g.数[2]);
      if (h < 0 || m < 0 || s2 < 0) return 誤('#NUM!');
      var 秒 = ((h * 3600 + m * 60 + s2) % 86400 + 86400) % 86400;
      return 数(秒 / 86400);
    },

    /* ══ ★★お金の 流れの 一族★★ ══（2026-09-16）
       ★物差し★ `golden-okane-4kansuu-2026-09-09.tsv`
       ★NPV／XNPV／MIRR は 式が 1本★＝★答えが 決まります★
       ★IRR／XIRR は ★探して 近づける★★＝★始めの 見当で 末桁が 変わります★
         （実測 … `=IRR(A1:A3)` と `=IRR(A1:A3,0)` で ★11桁目から 違う★） */
    NPV: function (引数たち, 手, 所) {
      if (引数たち.length < 2) return 誤('#VALUE!');
      var r = 数の引数(引数たち[0], 手, 所);
      if (r.誤) return r.誤;
      var 拾 = 数を拾う(引数たち.slice(1), 手);
      if (拾.誤) return 拾.誤;
      var 和 = 0;
      for (var i = 0; i < 拾.数たち.length; i++) 和 += 拾.数たち[i] / Math.pow(1 + r.数, i + 1);
      return 数(和);
    },
    XNPV: function (引数たち, 手, 所) {
      if (引数たち.length < 3) return 誤('#VALUE!');
      var r = 数の引数(引数たち[0], 手, 所);
      if (r.誤) return r.誤;
      var 金 = 数を拾う([引数たち[1]], 手);
      if (金.誤) return 金.誤;
      var 日 = 数を拾う([引数たち[2]], 手);
      if (日.誤) return 日.誤;
      if (!金.数たち.length || 金.数たち.length !== 日.数たち.length) return 誤('#NUM!');
      /* ★★利率が 0以下は #NUM!★★（2026-09-16 実測・6本）
           `=XNPV(0,…)` `=XNPV(-0.001,…)` `=XNPV(-0.99,…)` … 実Excel ★どれも #NUM!★
           ⇒★計算は 出来る（0 でも 割れる）が 実Excel は 断ります★ */
      if (!(r.数 > 0)) return 誤('#NUM!');
      var 頭 = 日.数たち[0], 和 = 0;
      /* ★★1つ目の 日より 前が 在れば #NUM!★★（2026-09-16 実測）
           `=XNPV(0.1,R1:R3,S1:S3)` … 実Excel ★#NUM!★／前は 数を 返して いた
           ＝★実Excel は 1つ目を 一番 早い 日と 決めて います★ */
      for (var z = 1; z < 日.数たち.length; z++) if (日.数たち[z] < 頭) return 誤('#NUM!');
      for (var i = 0; i < 金.数たち.length; i++) {
        和 += 金.数たち[i] / Math.pow(1 + r.数, (日.数たち[i] - 頭) / 365);
      }
      return 数(和);
    },
    MIRR: function (引数たち, 手, 所) {
      if (引数たち.length < 3) return 誤('#VALUE!');
      var 拾 = 数を拾う([引数たち[0]], 手);
      if (拾.誤) return 拾.誤;
      var f = 数の引数(引数たち[1], 手, 所);
      if (f.誤) return f.誤;
      var g = 数の引数(引数たち[2], 手, 所);
      if (g.誤) return g.誤;
      var a = 拾.数たち, n = a.length, i;
      if (n < 2) return 誤('#DIV/0!');
      var 出 = 0, 入 = 0, 何か負 = false, 何か正 = false;
      for (i = 0; i < n; i++) {
        if (a[i] < 0) { 何か負 = true; 出 += a[i] / Math.pow(1 + f.数, i); }
        else if (a[i] > 0) { 何か正 = true; 入 += a[i] * Math.pow(1 + g.数, n - 1 - i); }
      }
      /* ★★入りが 無ければ -1★★（2026-09-16 実測）
           `=MIRR(E1:E3,0.1,0.12)`（-100,-200,-300）… 実Excel ★-1★
           ＝入りの 終わり値が 0 ⇒ 0の (n-1)乗根 − 1 ＝ -1
         ★出が 無ければ 割れない＝#DIV/0!★
         ★再び 投げる 利が -1 なら 1+g が 0★＝#DIV/0!（実測 `=MIRR(A1:A3,0.1,-1)`） */
      if (!何か負) return 誤('#DIV/0!');
      if (1 + g.数 === 0) return 誤('#DIV/0!');
      if (!何か正) return 数(-1);
      /* ★★式を 1つ 余分に 掛けて いました★★（2026-09-16 実測）
           `=MIRR(A1:A3,0.1,0.12)`（-1000, 600, 700）
             前 ★0.0648400337706414★ ／実Excel ★0.17132403714770583★
           ★正しい 形★ … 入りを 終わりへ 運び／出を 頭へ 戻し／その 比の (n-1)乗根 − 1
             ＝ (600×1.12 ＋ 700) ÷ 1000 の 平方根 − 1 ＝ ★0.17132…★
           ⇒★`(1+f)^(n-1)` を 掛けるのは ★二重★でした★（出は もう 頭へ 戻して ある） */
      return 数(Math.pow(入 / -出, 1 / (n - 1)) - 1);
    },
    IRR: function (引数たち, 手, 所) {
      var 拾 = 数を拾う([引数たち[0]], 手);
      if (!引数たち.length || 拾.誤) return 拾.誤 || 誤('#VALUE!');
      var 見当 = 0.1;
      if (引数たち.length > 1) {
        var t = 数の引数(引数たち[1], 手, 所);
        if (t.誤) return t.誤;
        見当 = t.数;
      }
      return 探して近づける(拾.数たち, null, 見当);
    },
    XIRR: function (引数たち, 手, 所) {
      if (引数たち.length < 2) return 誤('#VALUE!');
      var 金 = 数を拾う([引数たち[0]], 手);
      if (金.誤) return 金.誤;
      var 日 = 数を拾う([引数たち[1]], 手);
      if (日.誤) return 日.誤;
      if (金.数たち.length !== 日.数たち.length) return 誤('#NUM!');
      var 見当 = 0.1;
      if (引数たち.length > 2) {
        var t = 数の引数(引数たち[2], 手, 所);
        if (t.誤) return t.誤;
        見当 = t.数;
      }
      return 探して近づける(金.数たち, 日.数たち, 見当);
    },

    /* ══ ★★表を 返す 一族★★ ══（2026-09-16）
       ★返す 形★ … `{ 溢れ: true, 行数, 列数, 並び }`（★`shiki-afure.js` と 同じ 形★）
       ★並びは 2次元★（行の 並びの 並び）
       ★物差し★ `golden-346-2026-09-08.tsv`（★紙は 左上の 1マスだけ 焼いて います★）
       ★実測★ `=HSTACK("あ","い")` → ★あ★（＝左上）／`=TRANSPOSE(A1:A5)` → 1 */
    VSTACK: function (引数たち, 手, 所) { return 積む(引数たち, 所, false); },
    HSTACK: function (引数たち, 手, 所) { return 積む(引数たち, 所, true); },
    TRANSPOSE: function (引数たち, 手, 所) {
      if (!引数たち.length) return 誤('#VALUE!');
      var A = 表にする(引数たち[0], 所);
      if (!A) return 誤('#VALUE!');
      if (!A.length || !A[0].length) return 誤('#VALUE!');
      var 出 = [];
      for (var c = 0; c < A[0].length; c++) {
        var 段 = [];
        for (var r = 0; r < A.length; r++) 段.push(A[r][c]);
        出.push(段);
      }
      return 溢れに(出);
    },
    SEQUENCE: function (引数たち, 手, 所) {
      if (!引数たち.length) return 誤('#VALUE!');
      var 取 = function (i, 既定) {
        if (引数たち.length <= i) return { 数: 既定 };
        return 数の引数(引数たち[i], 手, 所);
      };
      var r = 取(0, 1); if (r.誤) return r.誤;
      var c = 取(1, 1); if (c.誤) return c.誤;
      var s0 = 取(2, 1); if (s0.誤) return s0.誤;
      var d = 取(3, 1); if (d.誤) return d.誤;
      var 行 = 切り捨て(r.数), 列 = 切り捨て(c.数);
      if (行 < 1 || 列 < 1) return 誤('#VALUE!');
      if (行 * 列 > 1048576) return 誤('#NUM!');
      var 出 = [], v = s0.数;
      for (var i = 0; i < 行; i++) {
        var 段 = [];
        for (var j = 0; j < 列; j++) { 段.push(数(v)); v += d.数; }
        出.push(段);
      }
      return 溢れに(出);
    },
    /* ══ ★★TEXTAFTER ／ TEXTBEFORE★★ ══（2026-09-18）
       ★前は お客さんの 画面で #NAME? でした★（★台に 無い 52個の うちの 2個★）

       ★★紙（実Excel の 実測・12本）★★
         `docs/measured/golden-kansuu-7kaime-2026-09-18.tsv`
         ★9 / 12 が 見込み どおり★（★外れた 3本は 下★）
           =TEXTAFTER("a-b-c","-")      → ★"b-c"★
           =TEXTAFTER("a-b-c","-",2)    → ★"c"★
           =TEXTAFTER("a-b-c","-",-1)   → ★"c"★     （★後ろから 1つ目★）
           =TEXTAFTER("a-b-c","x")      → ★#N/A★
           =TEXTAFTER("a-b-c","-",5)    → ★#N/A★    （行き過ぎ）
           =TEXTBEFORE("a-b-c","-")     → ★"a"★
           =TEXTBEFORE("a-b-c","-",2)   → ★"a-b"★
           =TEXTBEFORE("a-b-c","-",-1)  → ★"a-b"★
           =TEXTBEFORE("a-b-c","x")     → ★#N/A★

       ★★★外れた 3本（★私たちの 見込みは 逆でした★）★★★
           =TEXTAFTER("a-B-c","b")      → ★★#N/A★★
             ⇒★★既定（4つ目を 省く）は ★大小を 見ます★★★
             ⇒★私は "-c" と 見込みました＝外れ★
           =TEXTAFTER("a-B-c","b",1,1)  → ★★"-c"★★
             ⇒★★`1` が ★大小を 見ない★★★
             ⇒★私は #N/A と 見込みました＝外れ★
           =TEXTBEFORE("abc","")        → ★★空の 字★★（★#VALUE! では ない★）
             ⇒★私は「見込み 無し」と 書きました★

       ★★測って いない 事（★名指しで 紙に★）★★
         `docs/measured/kansuu46/textafter-kiku-koto.md`
         ・★5つ目（match_end）★／★6つ目（if_not_found）★
         ・★区切りが 並び（四角）★／★空の 字を 探す TEXTAFTER★
         ・★0 番目★／★探す 字が 空で TEXTAFTER★ */
    TEXTAFTER: function (引数たち, 手, 所) { return 区切りの前後(引数たち, 手, 所, false); },
    TEXTBEFORE: function (引数たち, 手, 所) { return 区切りの前後(引数たち, 手, 所, true); },

    TEXTJOIN: function (引数たち, 手, 所) {
      if (引数たち.length < 3) return 誤('#VALUE!');
      /* ★★区切りは ★使う 時だけ★ 見ます★★（2026-09-16 実測）
           `=TEXTJOIN(A1:B5,2,FALSE)` … 実Excel ★FALSE★／前は ★#VALUE!★
           ＝区切り `A1:B5` は 2次元で 暗黙の 交わりが 出来ない ⇒ 誤り
           ⇒★でも 繋ぐ 物が 1つなら 区切りは ★1回も 使いません★★
           ⇒★★使わない 物の 誤りで 全体を 誤りに しない★★ */
      /* ★★区切りは ★並び★でも 受けます★★（2026-09-16 実測）
           `=TEXTJOIN(A1:A5,1,B1:B5,1)` … 実Excel ★214263841051★
             ＝2[1]4[2]6[3]8[4]10[5]1 ＝★A1:A5 を 順に 使い回して います★
           ⇒★1つの 字だと 思って いたら ★214161811011★（A1 だけ 使う）★
           ⇒★★知らなければ 一生 気づかない 形★★＝★紙 40行が 出しました★ */
      var 区たち = [];
      var 区の元 = 引数たち[0];
      if (区の元 && 区の元.種 === '四角') {
        for (var z = 0; z < (区の元.並び || []).length; z++) {
          var zv = 区の元.並び[z];
          if (!zv || zv.型 === '空') { 区たち.push(''); continue; }
          var zt = 計.字にする(zv);
          区たち.push(zt.型 === '誤' ? null : String(zt.値));
        }
      }
      var 区 = 字の引数(引数たち[0], 手, 所);
      var 判 = 真偽にする(場.ひとつに(引数たち[1], 所 && 所.今のマス));
      if (判.誤) return 判.誤;
      var 空を飛ばす = 判.真;
      var 並び = [], 誤り = null;
      ほどく(引数たち.slice(2), function (v) {
        if (v.型 === '誤') { 誤り = v; return false; }
        if (v.型 === '空') { if (!空を飛ばす) 並び.push(''); return true; }
        var t = 計.字にする(v);
        if (t.型 === '誤') { 誤り = t; return false; }
        並び.push(String(t.値));
        return true;
      });
      if (誤り) return 誤り;
      if (並び.length < 2) return 計.字(並び.length ? 並び[0] : '');
      if (区たち.length > 1) {
        var 出 = 並び[0];
        for (var w = 1; w < 並び.length; w++) {
          var d = 区たち[(w - 1) % 区たち.length];
          if (d === null) return 誤('#VALUE!');
          出 += d + 並び[w];
        }
        return 計.字(出);
      }
      if (区.誤) return 区.誤;                 /* ★2つ 以上 繋ぐ 時だけ 区切りが 要る★ */
      return 計.字(並び.join(区.字));
    },
    CONCAT: function (引数たち, 手) {
      var 出 = '', 誤り = null;
      ほどく(引数たち, function (v) {
        if (v.型 === '誤') { 誤り = v; return false; }
        if (v.型 === '空') return true;
        var t = 計.字にする(v);
        if (t.型 === '誤') { 誤り = t; return false; }
        出 += String(t.値);
        return true;
      });
      if (誤り) return 誤り;
      return 計.字(出);
    },

    /* ══ ★★丸めの 続きと 散らばりの 続き★★ ══（2026-09-16）
       ★物差し★ `golden-346-2026-09-08.tsv`
       ★`.PRECISE` と `ISO.CEILING` は ★刻みの 符号を 見ない★★（負でも 上へ）
       ★実測★ `=CEILING.PRECISE(-2)` → ★-2★／`=FLOOR.PRECISE(0.5)` → 0 */
    'CEILING.PRECISE': function (引数たち, 手, 所) { return きっちり丸め(引数たち, 手, 所, true); },
    'ISO.CEILING': function (引数たち, 手, 所) { return きっちり丸め(引数たち, 手, 所, true); },
    'FLOOR.PRECISE': function (引数たち, 手, 所) { return きっちり丸め(引数たち, 手, 所, false); },

    /* ★MULTINOMIAL★ … 全部 足した 階乗 ÷ それぞれの 階乗 */
    MULTINOMIAL: function (引数たち, 手) {
      var 拾 = 数を拾う(引数たち, 手);
      if (拾.誤) return 拾.誤;
      var a = 拾.数たち, 和 = 0, i;
      for (i = 0; i < a.length; i++) {
        var n = 切り捨て(a[i]);
        if (n < 0) return 誤('#NUM!');
        和 += n;
      }
      /* ★★階乗を 作ってから 割ると 桁が 落ちます★★（2026-09-16 実測）
           `=MULTINOMIAL(A1:B5)` … 割り算を 重ねる ★6.845116392948286e+35★
                                    実Excel      ★6.8451163929481654E+35★（★12桁目から 違う★）
         ⇒★二項係数を 掛けて 行く★＝★大きな 数を 作らずに 済む★
           multinomial(a,b,c…) ＝ C(a+b, b) × C(a+b+c, c) × … */
      var 出 = 1, 積み = 0;
      for (i = 0; i < a.length; i++) {
        var m = 切り捨て(a[i]);
        積み += m;
        for (var j = 1; j <= m; j++) 出 = 出 * (積み - m + j) / j;
      }
      if (!isFinite(出)) return 誤('#NUM!');
      return 数(出);
    },

    /* ★SKEW／SKEW.P★ … 歪み（★.P は 母集団★） */
    SKEW: function (引数たち, 手) { return 歪み(引数たち, 手, false); },
    'SKEW.P': function (引数たち, 手) { return 歪み(引数たち, 手, true); },

    /* ══ ★★条件つきの 一族★★ ══（2026-09-16）
       ★条件の 見方は D系と 同じ 台★＝`当てはまるか()`（★作る道を 2本に しない★）
       ★物差し★ `golden-346-2026-09-08.tsv` ほか
       ★実測★ `=COUNTIF(A1:A5,"<3")` → 2 ／ `=AVERAGEIF(A1:A5,"<3",B1:B5)` → ★3★
               `=COUNTIF(A1:A5,B1:B5)` → 1（★条件も 暗黙の 交わり★） */
    COUNTIF: function (引数たち, 手, 所) {
      var g = 条件で選ぶ(引数たち, 手, 所, 2, null);
      if (g.誤) return g.誤;
      return 数(g.当たり.length);
    },
    SUMIF: function (引数たち, 手, 所) {
      var g = 条件で選ぶ(引数たち, 手, 所, 2, 2);
      if (g.誤) return g.誤;
      var 和 = 0;
      for (var i = 0; i < g.値.length; i++) if (g.値[i] && g.値[i].型 === '数') 和 += g.値[i].値;
      return 数(和);
    },
    AVERAGEIF: function (引数たち, 手, 所) {
      var g = 条件で選ぶ(引数たち, 手, 所, 2, 2);
      if (g.誤) return g.誤;
      var 和 = 0, n = 0;
      for (var i = 0; i < g.値.length; i++) if (g.値[i] && g.値[i].型 === '数') { 和 += g.値[i].値; n++; }
      if (!n) return 誤('#DIV/0!');
      return 数(和 / n);
    },
    COUNTIFS: function (引数たち, 手, 所) {
      var g = 条件を重ねる(引数たち, 手, 所, 0);
      if (g.誤) return g.誤;
      return 数(g.印.filter(Boolean).length);
    },
    SUMIFS: function (引数たち, 手, 所) { return 重ねて集める(引数たち, 手, 所, 'SUM'); },
    MAXIFS: function (引数たち, 手, 所) { return 重ねて集める(引数たち, 手, 所, 'MAX'); },
    MINIFS: function (引数たち, 手, 所) { return 重ねて集める(引数たち, 手, 所, 'MIN'); },

    /* ★LARGE／SMALL★ … ★大きい／小さい 方から k番目★ */
    LARGE: function (引数たち, 手, 所) { return 大小のk番(引数たち, 手, 所, true); },
    SMALL: function (引数たち, 手, 所) { return 大小のk番(引数たち, 手, 所, false); },

    /* ★VLOOKUP／HLOOKUP★ … ★既定は 近い 方（TRUE）★ */
    VLOOKUP: function (引数たち, 手, 所) { return 縦横で探す(引数たち, 手, 所, true); },
    HLOOKUP: function (引数たち, 手, 所) { return 縦横で探す(引数たち, 手, 所, false); },

    /* ══ ★★お金の 一族（日付を 使わない 分）★★ ══（2026-09-16）
       ★物差し★ `golden-346-2026-09-08.tsv`
       ★実Excel に 聞いて 分かった 事★
         `=FV(0.5,1,2)` → ★-2★（★出て 行く 金は マイナス★）
         `=FV(2,3,TRUE)` → -13（★真偽は 1 に なる★）
         `=EFFECT(2,3)` → 3.629629629629628
       ★`型` は 0＝期末（既定）／0以外＝期首★ */
    FV: function (引数たち, 手, 所) {
      var g = 金の引数(引数たち, 手, 所, 3, 5);
      if (g.誤) return g.誤;
      var r = g.数[0], n = g.数[1], p = g.数[2], pv = g.数.length > 3 ? g.数[3] : 0;
      var 期首 = g.数.length > 4 && g.数[4] !== 0;
      if (r === 0) return 数(-(pv + p * n));
      var k = Math.pow(1 + r, n);
      if (!isFinite(k)) return 誤('#NUM!');      /* ★無限は 数に ならない★ */
      return 数(-(pv * k + p * (1 + (期首 ? r : 0)) * (k - 1) / r));
    },
    PV: function (引数たち, 手, 所) {
      var g = 金の引数(引数たち, 手, 所, 3, 5);
      if (g.誤) return g.誤;
      var r = g.数[0], n = g.数[1], p = g.数[2], fv = g.数.length > 3 ? g.数[3] : 0;
      var 期首 = g.数.length > 4 && g.数[4] !== 0;
      if (r === 0) return 数(-(fv + p * n));
      var k = Math.pow(1 + r, n);
      /* ★k が 無限＝割る 方が 無限★ ⇒ ★0 に 近づく★（PMT と 同じ 型） */
      if (!isFinite(k)) return 数(-(p * (1 + (期首 ? r : 0)) / r));
      return 数(-(fv + p * (1 + (期首 ? r : 0)) * (k - 1) / r) / k);
    },
    PMT: function (引数たち, 手, 所) {
      var g = 金の引数(引数たち, 手, 所, 3, 5);
      if (g.誤) return g.誤;
      var r = g.数[0], n = g.数[1], pv = g.数[2], fv = g.数.length > 3 ? g.数[3] : 0;
      var 期首 = g.数.length > 4 && g.数[4] !== 0;
      if (n === 0) return 誤('#NUM!');
      if (r === 0) return 数(-(pv + fv) / n);
      var k = Math.pow(1 + r, n);
      /* ★★(1+r)^n が 無限に なる 時★★（2026-09-16 実測）
           `=PMT(D1,D2,1)`（r＝45292 n＝46023）… 前 ★NaN★ ／実Excel ★-45292★
           ＝`Infinity * 1 + 0` ÷ `Infinity - 1` ⇒ ★NaN★
           ⇒★k が 大きく なると 答えは `-pv*r/(1+期首*r)` に 近づきます★
           ⇒★その 値を 返す★（★NaN は 誤りにも ならず マスに 出る＝一番 悪い★） */
      if (!isFinite(k)) return 数(-(pv * r) / (1 + (期首 ? r : 0)));
      return 数(-(pv * k + fv) * r / ((1 + (期首 ? r : 0)) * (k - 1)));
    },
    NPER: function (引数たち, 手, 所) {
      var g = 金の引数(引数たち, 手, 所, 3, 5);
      if (g.誤) return g.誤;
      var r = g.数[0], p = g.数[1], pv = g.数[2], fv = g.数.length > 3 ? g.数[3] : 0;
      var 期首 = g.数.length > 4 && g.数[4] !== 0;
      if (r === 0) { if (p === 0) return 誤('#NUM!'); return 数(-(pv + fv) / p); }
      var q = p * (1 + (期首 ? r : 0)) / r;
      var 中 = (q - fv) / (pv + q);
      if (!(中 > 0)) return 誤('#NUM!');
      return 数(Math.log(中) / Math.log(1 + r));
    },
    SLN: function (引数たち, 手, 所) {
      var g = 金の引数(引数たち, 手, 所, 3, 3);
      if (g.誤) return g.誤;
      if (g.数[2] === 0) return 誤('#DIV/0!');
      return 数((g.数[0] - g.数[1]) / g.数[2]);
    },
    SYD: function (引数たち, 手, 所) {
      var g = 金の引数(引数たち, 手, 所, 4, 4);
      if (g.誤) return g.誤;
      var c = g.数[0], 残 = g.数[1], n = g.数[2], t = g.数[3];
      if (n <= 0) return 誤('#NUM!');
      return 数((c - 残) * (n - t + 1) * 2 / (n * (n + 1)));
    },
    EFFECT: function (引数たち, 手, 所) {
      var g = 金の引数(引数たち, 手, 所, 2, 2);
      if (g.誤) return g.誤;
      var r = g.数[0], m = 切り捨て(g.数[1]);
      if (r <= 0 || m < 1) return 誤('#NUM!');
      return 数(Math.pow(1 + r / m, m) - 1);
    },
    NOMINAL: function (引数たち, 手, 所) {
      var g = 金の引数(引数たち, 手, 所, 2, 2);
      if (g.誤) return g.誤;
      var r = g.数[0], m = 切り捨て(g.数[1]);
      if (r <= 0 || m < 1) return 誤('#NUM!');
      return 数((Math.pow(1 + r, 1 / m) - 1) * m);
    },

    /* ══ ★★見分けの 一族★★ ══（2026-09-15）
       ★物差し★ `golden-346-2026-09-08.tsv` ／ `golden-mada-2026-09-14.tsv`
       ★実Excel に 聞いて 分かった 事★
         `=ISERR(A1:B5)` → ★TRUE★（★2次元の 四角は 暗黙の 交わりが 出来ず 誤りに なる★）
         `=ISBLANK(A1:A5)` → FALSE（暗黙の 交わりで A1）
       ★`ISERR` は #N/A を 数えない★／`ISERROR` は 全部 数える */
    ISBLANK: function (引数たち, 手, 所) { return 見分け(引数たち, 所, function (v) { return v.型 === '空'; }); },
    ISNUMBER: function (引数たち, 手, 所) { return 見分け(引数たち, 所, function (v) { return v.型 === '数'; }); },
    ISTEXT: function (引数たち, 手, 所) { return 見分け(引数たち, 所, function (v) { return v.型 === '字'; }); },
    ISNONTEXT: function (引数たち, 手, 所) { return 見分け(引数たち, 所, function (v) { return v.型 !== '字'; }); },
    ISLOGICAL: function (引数たち, 手, 所) { return 見分け(引数たち, 所, function (v) { return v.型 === '真偽'; }); },
    ISERROR: function (引数たち, 手, 所) { return 見分け(引数たち, 所, function (v) { return v.型 === '誤'; }); },
    ISERR: function (引数たち, 手, 所) {
      /* ★`#N/A` だけは 数えない★（`ISERROR` との 違い） */
      return 見分け(引数たち, 所, function (v) { return v.型 === '誤' && v.値 !== '#N/A'; });
    },
    ISNA: function (引数たち, 手, 所) {
      return 見分け(引数たち, 所, function (v) { return v.型 === '誤' && v.値 === '#N/A'; });
    },

    /* ★N★ … 数は そのまま／真偽は 1・0／誤りは 素通り／他は 0 */
    N: function (引数たち, 手, 所) {
      var v = 場.ひとつに(引数たち[0], 所 && 所.今のマス);
      if (!引数たち.length || !v) return 誤('#VALUE!');
      if (v.型 === '誤') return v;
      if (v.型 === '数') return 数(v.値);
      if (v.型 === '真偽') return 数(v.値 ? 1 : 0);
      return 数(0);
    },
    /* ★TYPE★ … 1 数／2 字／4 真偽／16 誤り／64 並び */
    TYPE: function (引数たち, 手, 所) {
      if (!引数たち.length) return 誤('#VALUE!');
      var 元 = 引数たち[0];
      if (元.種 === '四角' && ((元.行数 || 0) > 1 || (元.列数 || 0) > 1)) return 数(64);
      var v = 場.ひとつに(元, 所 && 所.今のマス);
      if (!v) return 数(16);
      if (v.型 === '誤') return 数(16);
      if (v.型 === '真偽') return 数(4);
      if (v.型 === '字') return 数(2);
      return 数(1);
    },

    IFNA: function (引数たち, 手, 所) {
      var 今 = 所 && 所.今のマス;
      if (引数たち.length < 2) return 誤('#VALUE!');
      var v = 場.ひとつに(引数たち[0], 今);
      if (!v || v.型 !== '誤' || v.値 !== '#N/A') return v || 計.空;
      return 場.ひとつに(引数たち[1], 今) || 計.空;
    },
    /* ★IFS★ … ★合う 所が 無ければ #N/A★ */
    IFS: function (引数たち, 手, 所) {
      var 今 = 所 && 所.今のマス;
      if (引数たち.length < 2 || 引数たち.length % 2) return 誤('#VALUE!');
      for (var i = 0; i < 引数たち.length; i += 2) {
        var 判 = 真偽にする(場.ひとつに(引数たち[i], 今));
        if (判.誤) return 判.誤;
        if (判.真) return 場.ひとつに(引数たち[i + 1], 今) || 計.空;
      }
      return 誤('#N/A');
    },
    SWITCH: function (引数たち, 手, 所) {
      var 今 = 所 && 所.今のマス;
      if (引数たち.length < 3) return 誤('#VALUE!');
      var 元 = 場.ひとつに(引数たち[0], 今);
      if (!元) return 誤('#VALUE!');
      if (元.型 === '誤') return 元;
      var i;
      for (i = 1; i + 1 < 引数たち.length; i += 2) {
        var c = 場.ひとつに(引数たち[i], 今);
        if (c && 場.比べる(元, c) === 0) return 場.ひとつに(引数たち[i + 1], 今) || 計.空;
      }
      /* ★余った 1つは 既定★／無ければ #N/A */
      if (i < 引数たち.length) return 場.ひとつに(引数たち[i], 今) || 計.空;
      return 誤('#N/A');
    },
    CHOOSE: function (引数たち, 手, 所) {
      if (引数たち.length < 2) return 誤('#VALUE!');
      var t = 数の引数(引数たち[0], 手, 所);
      if (t.誤) return t.誤;
      var k = 切り捨て(t.数);
      if (k < 1 || k >= 引数たち.length) return 誤('#VALUE!');
      return 場.ひとつに(引数たち[k], 所 && 所.今のマス) || 計.空;
    },
    COUNTBLANK: function (引数たち, 手) {
      var n = 0;
      ほどく(引数たち, function (v, 元) {
        if (元 === 'マス' && v.型 === '空') n++;
        return true;
      });
      return 数(n);
    },
    /* ══ ★★ROW／COLUMN★★ ══（2026-09-16）
       ★引数は 前から ★番地★を 持って います★（CELL("row",B7) の 為に 入れて あった）
         ＝★在るのに 呼んで いなかった★
       ★★四角を 渡されたら ★形の ぶんだけ 溢れます★★（2026-09-21 取り直し）
         ★★下の 2行は ★古い 道（COM の `.Formula`）で 測った 字★でした★★
           ＝`.Formula` は ★溢れません★＝★道具の 古さを 実Excel の 姿と 読み違えた★
           ⇒`Formula2`（今の 道）で 取り直すと ★溢れます★
         ★今の 字★（`golden-jitsu-excel-hanni-zenbu-2026-09-20.tsv`）
           `=ROW(A1:A3)` ★3マス（1 / 2 / 3）★ ／ `=COLUMN(A1:A3)` ★1マス（1）★
       ★下は 古い 字です（消さずに 残して います＝外れた 見込みは 捨てない）★
       ★四角を 渡されたら ★左上★★（2026-09-16 実測・★古い 道★）
         `=ROW(A1:A5)` ★　1★ ／ `=COLUMN(B1:B5)` ★　2★（溢れません）
       ★引数が 無ければ ★今 いる マス★★（`所.今のマス`）
         ★紙の `=COLUMN()` は 8★＝★H列で 測った 物★なので
         この 台（BZ1）では 合いません＝★名指しで まだ★に 出して あります */
    /* ★ISOMITTED★ … ★LAMBDA の 抜いた 引数か★を 見る 関数
         ★この 台に LAMBDA は まだ 無い★＝★抜けた 引数は 起き得ない★
         ⇒★半分 合う 答えでは ありません★＝★届く 全部の 場合で FALSE が 正しい★
         ★物差し★ 18行 すべて ★False★（`=ISOMITTED()` は 実Excel が 受け付けない） */
    /** ★ISOMITTED★ ... ★ラムダの 引数が 省かれたか★
        ★実測（2026-09-18・経営者1 の 2,116本）★
          `=LAMBDA(x,y,ISOMITTED(y))(1,2)` -> ★FALSE★（渡した）
          `=LAMBDA(x,y,ISOMITTED(y))(1,)`  -> ★TRUE★（★省いた★）
          `=LAMBDA(x,y,ISOMITTED(x))(,2)`  -> ★TRUE★
          `=LAMBDA(x,y,z,ISOMITTED(y))(1,,3)` -> ★TRUE★
          `=ISOMITTED(1)` `=LET(x,1,ISOMITTED(x))` -> FALSE
          `=ISOMITTED()` -> ★実Excel が 式として 受け付けません★
        ★★「省かれた」と「空」は 別です★★
          ＝`=IF(TRUE,,2)` の 空は ★0★（前から そう）
          ＝★ラムダに 渡す 時だけ★ `{型:'省略'}` を 渡します（`lib/shiki-hyou.js`）
        ★★未測定★★ ... 省いた 物を そのまま 使った 時（`=LAMBDA(x,y,y)(1,)`） */
    ISOMITTED: function (引数たち, 手, 所) {
      if (!引数たち.length) return 誤('#VALUE!');
      var a = 引数たち[0];
      if (a && a.値 && a.値.型 === '省略') return 計.真偽(true);
      var v = 場.ひとつに(a, 所 && 所.今のマス);
      if (v && v.型 === '省略') return 計.真偽(true);
      return 計.真偽(false);
    },
    ROW: function (引数たち, 手, 所) { return 番地の番号(引数たち, 所, true); },
    COLUMN: function (引数たち, 手, 所) { return 番地の番号(引数たち, 所, false); },
    /* ══ ★★AREAS★★ ══（2026-09-16）
       ★★AREAS は マスの 中身を 1度も 見て いません★★
         見て いるのは ★「その 式が マスを 指して いるか」★ だけ
         （`docs/measured/kansuu46/AREAS-dou-dasu-ka.md` … 実測 89本）
         ★実測★ A1 に 字を 入れて `=AREAS(IF(TRUE,A1,A2))` → ★1★
               D1 に 誤りを 入れて `=AREAS(IF(TRUE,D1,B1))` → ★1★（★誤りすら 伝わらない★）
               `=AREAS(IF(TRUE,1,2))` → ★#VALUE!★（値そのもの）
       ★束ね（カンマ）は ★重なって いても 数える★
         `=AREAS((A1,A1))` → ★2★ ／ `=AREAS((A1:A2,A2:A3))` → ★2★
       ★まだ 出さない 形★（名指しで 棚に 出して あります）
         ★列まるごと（A:A）・行まるごと（1:1）★ … 台が 番地に して いない
         ★重なり（空白）★ … 交わらない 時の #NULL! を 台が 持って いない */
    AREAS: function (引数たち) {
      if (!引数たち.length) return 誤('#VALUE!');
      var a = 引数たち[0];
      if (!a) return 誤('#VALUE!');
      /* ★★誤りは そのまま 通す★★（2026-09-18・INDIRECT を 書いた 時に 出た）
           ★実Excel★ `=AREAS(INDIRECT("zzz"))` → ★#REF!★（#VALUE! では ない）
           ⇒★中の 誤りを ★自分の #VALUE! で 塗り潰して いました★★ */
      if (a.値 && a.値.型 === '誤') return a.値;
      if (a.種 === '束') {
        /* ★中の どれか 1つでも マスを 指さなければ #VALUE!★ */
        for (var i = 0; i < a.並.length; i++) {
          var b = a.並[i];
          if (!b || (b.種 !== 'マス' && b.種 !== '四角')) return 誤('#VALUE!');
        }
        return 数(a.並.length);
      }
      if (a.種 === 'マス' || a.種 === '四角') return 数(1);
      return 誤('#VALUE!');
    },
    /* ══ ★★誤りは そのまま 通す★★ ══（2026-09-18・経営者1 の 実測）
         ★実Excel★ `=ROWS(FILTERXML(...))` で 中が #VALUE! なら ★#VALUE!★
         ★前の うち★ ... ★1★ を 返して いました（★1マス扱い★）
         ⇒★★中の 誤りを 自分の 数で 塗り潰して いました★★
           ＝今日 `AREAS` で 直したのと ★同じ 形★
         ★誤りで ない 時は 全部 合って います★
           `//b` -> 3 ／ `//c` -> 2 ／ `//c/@id` -> 2 ／ `//b[position()>1]` -> 2
           ＝★溢れの 本数は 数えられて います＝誤りの 時だけ★
         ★★COLUMNS は 測って いません★★（★紙に 誤りの 例が 在りません★）
           ⇒★但し ROWS と 対なので 同じに します★
           ⇒★★合わなければ ここを 見て ください★★ */
    ROWS: function (引数たち) {
      if (!引数たち.length) return 誤('#VALUE!');
      var a = 引数たち[0];
      if (a && a.値 && a.値.型 === '誤') return a.値;
      return 数(a.種 === '四角' ? (a.行数 || 0) : 1);
    },
    COLUMNS: function (引数たち) {
      if (!引数たち.length) return 誤('#VALUE!');
      var a = 引数たち[0];
      if (a && a.値 && a.値.型 === '誤') return a.値;   /* ★ROWS と 同じ（★未測定★）★ */
      return 数(a.種 === '四角' ? (a.列数 || 0) : 1);
    },

    /* ══ ★★字の 一族★★ ══（2026-09-15）
       ★物差し★ `golden-346-2026-09-08.tsv`
       ★字に する 所は 1つに★ … `計.字にする()`（`&` と 同じ 台）
       ★実Excel に 聞いて 分かった 事★
         `=CODE(D1)` → ★52★（D1＝45292 ⇒ "45292" の 頭 `4`）
         `=CODE(TRUE)` → ★84★（`T`）＝★真偽も 字に なる★
         `=LEFT(A1)` → `1`（★型は 字★）／`=LEFT("101",2)` → `10`
         `=EXACT(A1:A5,A1)` → ★TRUE★（暗黙の 交わりで どちらも A1） */
    LEFT: function (引数たち, 手, 所) { return 端を取る(引数たち, 手, 所, true); },
    RIGHT: function (引数たち, 手, 所) { return 端を取る(引数たち, 手, 所, false); },

    /* ══ ★★LENB ／ LEFTB ／ RIGHTB ／ MIDB★★ ══（2026-09-18）
       ★前は お客さんの 画面で #NAME? でした★（★台に 無い 56個の うちの 4個★）

       ★★紙（実Excel の 実測・12本）★★
         `docs/measured/golden-kansuu-7kaime-2026-09-18.tsv`
         ★11 / 12 が 見込み どおり★（★外れた 1本は 下の ★MIDB★★）
           =LENB("ABC")                      → ★3★   （半角 1バイト）
           =LENB(全角ア)                      → ★2★   （★全角 2バイト★）
           =LENB(半角ｱ)                       → ★1★   （★半角カナは 1バイト★）
           =LENB(全角ア&"A")                  → ★3★
           =LENB("")                          → ★0★
           =LEN(LEFTB(全角ア,1))              → ★1★
           =UNICODE(LEFTB(全角ア,1))          → ★★32（空白）★★
           =LEN(LEFTB(全角ア,2))              → ★1★
           =LEN(LEFTB(全角ア&"A",3))          → ★2★
           =LEN(RIGHTB(全角ア&"A",1))         → ★1★
           =LEN(MIDB(全角ア&"A",1,2))         → ★1★
           ★=LEN(MIDB(全角ア&"A",2,2))        → ★2★★ … ★私の 見込みは 1＝外れました★

       ★★決まり（★実測から★）★★
         ★2バイトの 字を ★半分で 切ると 空白 1文字に なる★★
         ⇒`LEFTB(ア,1)` … 前の 半分だけ ⇒ ★空白（32）／1文字★
         ⇒`MIDB(ア&"A",2,2)` … ★ア の 後ろ半分★ ＋ A ⇒ ★空白 ＋ A ＝ 2文字★
           （★私は 1文字と 見込みました＝外れ★）

       ★★1バイトか 2バイトかの 決め方★★
         ★1バイト★ … ASCII（U+0000〜U+007F）と ★半角カナ（U+FF61〜U+FF9F）★
         ★2バイト★ … それ以外
         ★★測ったのは 3つだけ★★（半角A ／ 全角ア ／ 半角ｱ）
         ⇒★漢字・ひらがな・絵文字は ★測って いません★★
         ⇒`docs/measured/kansuu46/lenb-kiku-koto.md` に 名指しで */
    LENB: function (引数たち, 手, 所) {
      var t = 字の引数(引数たち[0], 手, 所);
      if (!引数たち.length || t.誤) return t.誤 || 誤('#VALUE!');
      return 数(バイト数(t.字));
    },
    LEFTB: function (引数たち, 手, 所) { return バイトで端を取る(引数たち, 手, 所, true); },
    RIGHTB: function (引数たち, 手, 所) { return バイトで端を取る(引数たち, 手, 所, false); },
    MIDB: function (引数たち, 手, 所) {
      if (引数たち.length < 3) return 誤('#VALUE!');
      var t = 字の引数(引数たち[0], 手, 所);
      if (t.誤) return t.誤;
      var s = 数の引数(引数たち[1], 手, 所);
      if (s.誤) return s.誤;
      var n = 数の引数(引数たち[2], 手, 所);
      if (n.誤) return n.誤;
      var 始 = 切り捨て(s.数), 長 = 切り捨て(n.数);
      if (始 < 1 || 長 < 0) return 誤('#VALUE!');
      return 計.字(バイトで切る(t.字, 始, 長));
    },
    LEN: function (引数たち, 手, 所) {
      var t = 字の引数(引数たち[0], 手, 所);
      if (!引数たち.length || t.誤) return t.誤 || 誤('#VALUE!');
      return 数(t.字.length);
    },
    MID: function (引数たち, 手, 所) {
      if (引数たち.length < 3) return 誤('#VALUE!');
      var t = 字の引数(引数たち[0], 手, 所);
      if (t.誤) return t.誤;
      var a = 数の引数(引数たち[1], 手, 所);
      if (a.誤) return a.誤;
      var b = 数の引数(引数たち[2], 手, 所);
      if (b.誤) return b.誤;
      var 頭 = 切り捨て(a.数), 長 = 切り捨て(b.数);
      if (頭 < 1 || 長 < 0) return 誤('#VALUE!');
      return 計.字(t.字.substr(頭 - 1, 長));
    },
    UPPER: function (引数たち, 手, 所) { return 字ひとつ(引数たち, 手, 所, function (x) { return x.toUpperCase(); }); },
    LOWER: function (引数たち, 手, 所) { return 字ひとつ(引数たち, 手, 所, function (x) { return x.toLowerCase(); }); },
    TRIM: function (引数たち, 手, 所) {
      /* ★間の 空きは 1つに する★（実Excel と 同じ） */
      return 字ひとつ(引数たち, 手, 所, function (x) { return x.replace(/ +/g, ' ').replace(/^ | $/g, ''); });
    },
    PROPER: function (引数たち, 手, 所) {
      /* ★字でない 物の 次を 大文字に する★（実Excel と 同じ） */
      return 字ひとつ(引数たち, 手, 所, function (x) {
        var 出 = '', 前が字 = false;
        for (var i = 0; i < x.length; i++) {
          var c = x[i], 字か = /[A-Za-zÀ-ɏ]/.test(c);
          出 += 字か ? (前が字 ? c.toLowerCase() : c.toUpperCase()) : c;
          前が字 = 字か;
        }
        return 出;
      });
    },
    CLEAN: function (引数たち, 手, 所) {
      /* eslint-disable-next-line no-control-regex */
      return 字ひとつ(引数たち, 手, 所, function (x) { return x.replace(/[\u0000-\u001f]/g, ''); });
    },
    T: function (引数たち, 手, 所) {
      var v = 場.ひとつに(引数たち[0], 所 && 所.今のマス);
      if (!引数たち.length || !v) return 誤('#VALUE!');
      if (v.型 === '誤') return v;
      return 計.字(v.型 === '字' ? String(v.値) : '');
    },
    REPT: function (引数たち, 手, 所) {
      if (引数たち.length < 2) return 誤('#VALUE!');
      var t = 字の引数(引数たち[0], 手, 所);
      if (t.誤) return t.誤;
      var n = 数の引数(引数たち[1], 手, 所);
      if (n.誤) return n.誤;
      var k = 切り捨て(n.数);
      if (k < 0) return 誤('#VALUE!');
      return 計.字(t.字.repeat(k));
    },
    EXACT: function (引数たち, 手, 所) {
      if (引数たち.length < 2) return 誤('#VALUE!');
      var a = 字の引数(引数たち[0], 手, 所);
      if (a.誤) return a.誤;
      var b = 字の引数(引数たち[1], 手, 所);
      if (b.誤) return b.誤;
      return 計.真偽(a.字 === b.字);
    },
    CODE: function (引数たち, 手, 所) {
      var t = 字の引数(引数たち[0], 手, 所);
      if (!引数たち.length || t.誤) return t.誤 || 誤('#VALUE!');
      if (!t.字.length) return 誤('#VALUE!');
      return 数(t.字.charCodeAt(0));
    },
    UNICODE: function (引数たち, 手, 所) {
      var t = 字の引数(引数たち[0], 手, 所);
      if (!引数たち.length || t.誤) return t.誤 || 誤('#VALUE!');
      if (!t.字.length) return 誤('#VALUE!');
      return 数(t.字.codePointAt(0));
    },
    CHAR: function (引数たち, 手, 所) {
      var n = 数の引数(引数たち[0], 手, 所);
      if (!引数たち.length || n.誤) return n.誤 || 誤('#VALUE!');
      var k = 切り捨て(n.数);
      if (k < 1 || k > 255) return 誤('#VALUE!');
      return 計.字(String.fromCharCode(k));
    },
    UNICHAR: function (引数たち, 手, 所) {
      var n = 数の引数(引数たち[0], 手, 所);
      if (!引数たち.length || n.誤) return n.誤 || 誤('#VALUE!');
      var k = 切り捨て(n.数);
      if (k < 1 || k > 1114111) return 誤('#VALUE!');
      return 計.字(String.fromCodePoint(k));
    },

    /* ══ ★★ASC ／ DBCS★★ ══（2026-09-18）
       ★前は お客さんの 画面で #NAME? でした★（★台に 無い 58個の うちの 2個★）

       ★★紙（実Excel の 実測・22本）★★
         `docs/measured/golden-kansuu-7kaime-2026-09-18.tsv`
         ★★22本 とも 見込み どおりでした★★（★当て推量 0★）
         ★字は 直に 書かず `UNICHAR` で 作り、答えも コードと 文字数で 受けました★
           ASC  全角Ａ(65313)→★65／1文字★   全角１(65297)→★49／1文字★
                全角ア(12450)→★65393／1文字★
                ★全角ガ(12460)→★65398／2文字★★（★濁点が 分かれる★）
                ★全角空白(12288)→★32／1文字★★
                全角、(12289)→★65380／1文字★
           DBCS 半角A(65)→★65313★  半角1(49)→★65297★  半角ｱ(65393)→★12450★
                ★半角濁点(65438)→★12443★★  半角空白(32)→★12288★

       ★★★字の 表を ★記憶で 書いて いません★★★（★これが 一番 大事★）
         ★訳★ … 半角カナの 表は ★62字 在ります★。
                ★記憶で 書くと 必ず どこか 間違えます★
                （2026-09-08 … 役所の 続柄コードを ★12個 並べて 12個とも 違った★）
         ⇒★★`String.normalize('NFKC'／'NFKD')` を 使います★★
            ＝★node（と ブラウザ）が 持って いる ★Unicode の 表★★
            ＝★私の 記憶では なく 機械の 表★
         ⇒★ASC★ … `NFKD` で ばらして（ガ → カ ＋ 結合濁点）、
                    半角カナの 表を ★逆に 引いて★ 当てる
         ⇒★DBCS★ … 半角カナは `NFKC`／ASCII は ＋0xFEE0／空白は U+3000

       ★★但し ★実測が 1か所 教えて くれました★★★
         ★DBCS(半角濁点 U+FF9E)★ … `NFKC` は ★U+3099（結合用）★を 返します
                                    ★実Excel は U+309B（単独の ゛）★
         ⇒★★ここだけ 別に して います★★（★測った から 書けました★）

       ★★測って いない 事（★名指しで 紙に★）★★
         `docs/measured/kansuu46/asc-dbcs-kiku-koto.md`
         ・★全角ガ の ★2文字目★★（紙は「2文字」までしか 言って いません）
         ・★62字 の うち 測ったのは ★4字★★（ア・ガ・、・空白）
         ・★半濁点（パ）★／★ー（長音）★／★ヴ★／★全角の ゛（U+309B）に ASC★
         ・★ひらがな★（ASC は 何も しない はず＝★未測定★）
         ・★2文字目 以降が 混ざった 字★ */
    ASC: function (引数たち, 手, 所) {
      var t = 字の引数(引数たち[0], 手, 所);
      if (!引数たち.length || t.誤) return t.誤 || 誤('#VALUE!');
      return 計.字(全角を半角に(t.字));
    },
    DBCS: function (引数たち, 手, 所) {
      var t = 字の引数(引数たち[0], 手, 所);
      if (!引数たち.length || t.誤) return t.誤 || 誤('#VALUE!');
      return 計.字(半角を全角に(t.字));
    },
    SUBSTITUTE: function (引数たち, 手, 所) {
      if (引数たち.length < 3) return 誤('#VALUE!');
      var t = 字の引数(引数たち[0], 手, 所);
      if (t.誤) return t.誤;
      var 古 = 字の引数(引数たち[1], 手, 所);
      if (古.誤) return 古.誤;
      var 新 = 字の引数(引数たち[2], 手, 所);
      if (新.誤) return 新.誤;
      if (古.字 === '') return 計.字(t.字);
      if (引数たち.length < 4) return 計.字(t.字.split(古.字).join(新.字));
      var k = 数の引数(引数たち[3], 手, 所);
      if (k.誤) return k.誤;
      var 何番 = 切り捨て(k.数);
      if (何番 < 1) return 誤('#VALUE!');
      var 場所 = -1, 数え = 0, i = 0;
      while ((i = t.字.indexOf(古.字, i)) >= 0) {
        数え++;
        if (数え === 何番) { 場所 = i; break; }
        i += 古.字.length;
      }
      if (場所 < 0) return 計.字(t.字);
      return 計.字(t.字.slice(0, 場所) + 新.字 + t.字.slice(場所 + 古.字.length));
    },
    REPLACE: function (引数たち, 手, 所) {
      if (引数たち.length < 4) return 誤('#VALUE!');
      var t = 字の引数(引数たち[0], 手, 所);
      if (t.誤) return t.誤;
      var a = 数の引数(引数たち[1], 手, 所);
      if (a.誤) return a.誤;
      var b = 数の引数(引数たち[2], 手, 所);
      if (b.誤) return b.誤;
      var 新 = 字の引数(引数たち[3], 手, 所);
      if (新.誤) return 新.誤;
      var 頭 = 切り捨て(a.数), 長 = 切り捨て(b.数);
      if (頭 < 1 || 長 < 0) return 誤('#VALUE!');
      return 計.字(t.字.slice(0, 頭 - 1) + 新.字 + t.字.slice(頭 - 1 + 長));
    },
    FIND: function (引数たち, 手, 所) { return 探す(引数たち, 手, 所, true); },
    SEARCH: function (引数たち, 手, 所) { return 探す(引数たち, 手, 所, false); },

    /* ══ ★★三角と 丸めの 一族★★ ══（2026-09-15）
       ★物差し★ `golden-346-2026-09-08.tsv`
       ★実Excel に 聞いて 分かった 事★
         `=CEILING.MATH(2,3,4)` → 3（★3つ目は 負の 数の 時だけ 効く★）
         `=FLOOR.MATH(2,3)` → ★0★
         `=ROUND(2.5,0)` → ★3★（★半分は 0から 遠い 方へ★＝JS の `Math.round` とは 負で 違う） */
    SIN: function (引数たち, 手, 所) { return 数ひとつ(引数たち, 手, 所, Math.sin); },
    COS: function (引数たち, 手, 所) { return 数ひとつ(引数たち, 手, 所, Math.cos); },
    TAN: function (引数たち, 手, 所) { return 数ひとつ(引数たち, 手, 所, Math.tan); },
    ASIN: function (引数たち, 手, 所) {
      return 数ひとつ(引数たち, 手, 所, function (x) { return (x < -1 || x > 1) ? null : Math.asin(x); });
    },
    ACOS: function (引数たち, 手, 所) {
      return 数ひとつ(引数たち, 手, 所, function (x) { return (x < -1 || x > 1) ? null : Math.acos(x); });
    },
    ATAN: function (引数たち, 手, 所) { return 数ひとつ(引数たち, 手, 所, Math.atan); },
    ATAN2: function (引数たち, 手, 所) {
      /* ★実Excel は (x, y) の 順★＝JS の `Math.atan2(y, x)` と ★逆★ */
      return 数ふたつ(引数たち, 手, 所, function (x, y) {
        return (x === 0 && y === 0) ? null : Math.atan2(y, x);
      });
    },
    SINH: function (引数たち, 手, 所) { return 数ひとつ(引数たち, 手, 所, Math.sinh); },
    COSH: function (引数たち, 手, 所) { return 数ひとつ(引数たち, 手, 所, Math.cosh); },
    TANH: function (引数たち, 手, 所) { return 数ひとつ(引数たち, 手, 所, Math.tanh); },
    ASINH: function (引数たち, 手, 所) { return 数ひとつ(引数たち, 手, 所, Math.asinh); },
    ACOSH: function (引数たち, 手, 所) {
      return 数ひとつ(引数たち, 手, 所, function (x) { return x < 1 ? null : Math.acosh(x); });
    },
    ATANH: function (引数たち, 手, 所) {
      return 数ひとつ(引数たち, 手, 所, function (x) { return (x <= -1 || x >= 1) ? null : Math.atanh(x); });
    },
    COT: function (引数たち, 手, 所) {
      return 数ひとつ(引数たち, 手, 所, function (x) { return x === 0 ? null : 1 / Math.tan(x); });
    },
    SEC: function (引数たち, 手, 所) { return 数ひとつ(引数たち, 手, 所, function (x) { return 1 / Math.cos(x); }); },
    CSC: function (引数たち, 手, 所) {
      return 数ひとつ(引数たち, 手, 所, function (x) { return x === 0 ? null : 1 / Math.sin(x); });
    },

    /* ★ROUND／ROUNDUP／ROUNDDOWN★ … ★半分は 0から 遠い 方へ★ */
    ROUND: function (引数たち, 手, 所) { return 丸める(引数たち, 手, 所, 'R'); },
    ROUNDUP: function (引数たち, 手, 所) { return 丸める(引数たち, 手, 所, 'U'); },
    ROUNDDOWN: function (引数たち, 手, 所) { return 丸める(引数たち, 手, 所, 'D'); },

    /* ★CEILING.MATH／FLOOR.MATH★ … ★3つ目（向き）は ★負の 数の 時だけ★ 効く★ */
    'CEILING.MATH': function (引数たち, 手, 所) { return 数学丸め(引数たち, 手, 所, true); },
    'FLOOR.MATH': function (引数たち, 手, 所) { return 数学丸め(引数たち, 手, 所, false); },

    /* ★GCD／LCM★ … ★切り捨ててから★（実Excel も そう） */
    GCD: function (引数たち, 手) { return 整数の組(引数たち, 手, function (a, b) {
      while (b) { var t = a % b; a = b; b = t; }
      return a;
    }, 0); },
    LCM: function (引数たち, 手) { return 整数の組(引数たち, 手, function (a, b) {
      if (a === 0 || b === 0) return 0;
      var x = a, y = b;
      while (y) { var t = x % y; x = y; y = t; }
      return a / x * b;
    }, 1); },

    ISEVEN: function (引数たち, 手, 所) {
      var t = 数の引数(引数たち[0], 手, 所);
      if (!引数たち.length || t.誤) return t.誤 || 誤('#VALUE!');
      return 計.真偽(Math.abs(切り捨て(t.数)) % 2 === 0);
    },
    ISODD: function (引数たち, 手, 所) {
      var t = 数の引数(引数たち[0], 手, 所);
      if (!引数たち.length || t.誤) return t.誤 || 誤('#VALUE!');
      return 計.真偽(Math.abs(切り捨て(t.数)) % 2 === 1);
    },

    /* ══ ★★数の 一族★★ ══（2026-09-15）
       ★物差し★ `golden-346-2026-09-08.tsv`（実Excel 16.0 build 20326）
       ★引数の 取り方は 1つに★ … `数の引数()`（暗黙の 交わり → 数に する）
       ★実Excel に 聞いて 分かった 事★
         `=EVEN(0.5)` → ★2★（★0 では ない＝必ず 外へ 逃がす★）
         `=EVEN(TRUE)` → 2 ／ `=EVEN(FALSE)` → ★0★
         `=CEILING(2,TRUE)` → 2 ／ `=FLOOR(2,3)` → ★0★
         `=COMBIN(2.5,1)` → 2（★切り捨ててから 数える★）
         `=TRUNC` は ★桁を 省ける★（既定 0） */
    ABS: function (引数たち, 手, 所) { return 数ひとつ(引数たち, 手, 所, Math.abs); },
    SIGN: function (引数たち, 手, 所) { return 数ひとつ(引数たち, 手, 所, Math.sign); },
    INT: function (引数たち, 手, 所) { return 数ひとつ(引数たち, 手, 所, Math.floor); },
    SQRT: function (引数たち, 手, 所) {
      return 数ひとつ(引数たち, 手, 所, function (x) { return x < 0 ? null : Math.sqrt(x); });
    },
    EXP: function (引数たち, 手, 所) { return 数ひとつ(引数たち, 手, 所, Math.exp); },
    LN: function (引数たち, 手, 所) {
      return 数ひとつ(引数たち, 手, 所, function (x) { return x <= 0 ? null : Math.log(x); });
    },
    LOG10: function (引数たち, 手, 所) {
      return 数ひとつ(引数たち, 手, 所, function (x) { return x <= 0 ? null : Math.log10(x); });
    },
    RADIANS: function (引数たち, 手, 所) {
      return 数ひとつ(引数たち, 手, 所, function (x) { return x * Math.PI / 180; });
    },
    DEGREES: function (引数たち, 手, 所) {
      return 数ひとつ(引数たち, 手, 所, function (x) { return x * 180 / Math.PI; });
    },
    PI: function (引数たち) { return 引数たち.length ? 誤('#VALUE!') : 数(Math.PI); },

    /* ★EVEN／ODD★ … ★0 から 外へ 逃がす★（`=EVEN(0.5)` → 2） */
    EVEN: function (引数たち, 手, 所) {
      return 数ひとつ(引数たち, 手, 所, function (x) {
        if (x === 0) return 0;
        var 符 = x < 0 ? -1 : 1, a = Math.abs(x);
        return 符 * Math.ceil(a / 2) * 2;
      });
    },
    ODD: function (引数たち, 手, 所) {
      return 数ひとつ(引数たち, 手, 所, function (x) {
        if (x === 0) return 1;
        var 符 = x < 0 ? -1 : 1, a = Math.abs(x);
        var k = Math.ceil((a + 1) / 2) * 2 - 1;
        return 符 * k;
      });
    },

    /* ★TRUNC★ … 桁を 省ける（既定 0）／★0 の 方へ 切る★ */
    TRUNC: function (引数たち, 手, 所) {
      var t = 数の引数(引数たち[0], 手, 所);
      if (!引数たち.length || t.誤) return t.誤 || 誤('#VALUE!');
      var 桁 = 0;
      if (引数たち.length > 1) {
        var d = 数の引数(引数たち[1], 手, 所);
        if (d.誤) return d.誤;
        桁 = 切り捨て(d.数);
      }
      /* ★★桁が 大きすぎると 10^桁 が 無限に なります★★（2026-09-15 実測）
           `=TRUNC(D1,D2)`（D1=45292 D2=46023）… 前 ★NaN★ ／実Excel ★45292★
           ＝`10^46023` → `Infinity` ⇒ `Infinity/Infinity` → `NaN`
           ⇒★桁が 十分 大きければ ★何も 切らない★★（元の 数を そのまま） */
      var p = Math.pow(10, 桁);
      if (!isFinite(p)) return 数(桁 > 0 ? t.数 : 0);
      if (p === 0) return 数(0);
      var x = t.数 * p;
      if (!isFinite(x)) return 数(t.数);
      return 数((x < 0 ? Math.ceil(x) : Math.floor(x)) / p);
    },

    POWER: function (引数たち, 手, 所) { return 数ふたつ(引数たち, 手, 所, Math.pow); },
    QUOTIENT: function (引数たち, 手, 所) {
      return 数ふたつ(引数たち, 手, 所, function (a, b) {
        if (b === 0) return null;
        var x = a / b;
        return x < 0 ? Math.ceil(x) : Math.floor(x);
      });
    },
    /* ★CEILING／FLOOR★ … ★刻みの 倍数に する★（刻み 0 は 0） */
    CEILING: function (引数たち, 手, 所) {
      return 数ふたつ(引数たち, 手, 所, function (x, k) {
        if (k === 0) return 0;
        return Math.ceil(x / k) * k;
      });
    },
    FLOOR: function (引数たち, 手, 所) {
      return 数ふたつ(引数たち, 手, 所, function (x, k) {
        if (k === 0) return 0;
        return Math.floor(x / k) * k;
      });
    },
    MROUND: function (引数たち, 手, 所) {
      return 数ふたつ(引数たち, 手, 所, function (x, k) {
        if (k === 0) return 0;
        if ((x < 0) !== (k < 0) && x !== 0) return null;   /* 実Excel は 符号違いで #NUM! */
        return Math.round(x / k) * k;
      });
    },
    LOG: function (引数たち, 手, 所) {
      var t = 数の引数(引数たち[0], 手, 所);
      if (!引数たち.length || t.誤) return t.誤 || 誤('#VALUE!');
      var 底 = 10;
      if (引数たち.length > 1) {
        var b = 数の引数(引数たち[1], 手, 所);
        if (b.誤) return b.誤;
        底 = b.数;
      }
      if (t.数 <= 0 || 底 <= 0 || 底 === 1) return 誤('#NUM!');
      /* ★★自然対数の 比では ありません★★（2026-09-15 実測）
           `=LOG(2,3)` … 実Excel `0.63092975357145742`
             `ln2/ln3`   … `0.6309297535714575`  ←★1つ 上★
             `log2/log2` … `0.6309297535714575`  ←★同じく 違う★
             ★`log10(2)/log10(3)` … `0.6309297535714574`★ ＝★実Excel と 同じ double★
           ⇒★実Excel は ★常用対数の 比★で 出して います★
           ⇒★算数では どれも 同じ／機械の 中では 別★（HARMEAN と 同じ 型）
         ★★但し 実Excel 自身が 1つの 形では ありません★★
           `=LOG("101",2)` … 実Excel `6.6582114827517955`＝★自然対数の 比が 合う★
           ⇒★両方 合う 形は 作れない★／★16行 中 15行 合う 常用対数の 比を 採った★
           ⇒残り 1本は ★名指しで まだ★（GEOMEAN と 同じ 形の 話・棚㊲） */
      return 数(Math.log10(t.数) / Math.log10(底));
    },
    FACT: function (引数たち, 手, 所) {
      return 数ひとつ(引数たち, 手, 所, function (x) {
        var n = 切り捨て(x);
        if (n < 0) return null;
        var r = 1;
        for (var i = 2; i <= n; i++) r *= i;
        return r;
      });
    },
    COMBIN: function (引数たち, 手, 所) {
      return 数ふたつ(引数たち, 手, 所, function (a, b) {
        var n = 切り捨て(a), k = 切り捨て(b);
        if (n < 0 || k < 0 || k > n) return null;
        var r = 1;
        for (var i = 1; i <= k; i++) r = r * (n - k + i) / i;
        return Math.round(r);
      });
    },

    /* ══ ★★順列 2つ★★ ══（2026-09-18）
       ★前は お客さんの 画面で #NAME? でした★（台に 無い 62個の うちの 2個）

       ★★紙（実Excel の 実測）★★ … `docs/measured/golden-oddf-to-46ko-2026-09-16.tsv`
         `=PERMUT(5,2)`        → ★20★  （＝5×4）
         `=PERMUTATIONA(5,2)`  → ★25★  （＝5^2）
       ★★紙は この 2本 だけです★★
         ⇒★★中の 計算は 決まって います★★
           PERMUT       ＝ n×(n−1)×…×(n−k+1)   （＝n! ÷ (n−k)!）
           PERMUTATIONA ＝ n^k                   （★同じ 物を 選び直してよい★）
         ⇒★★但し ★境目（負の 数・0・k>n）は 測って いません★★★
           ＝★ここは COMBIN（★実測済み★）と 同じ 決めに 揃えました★
             COMBIN … `n < 0 || k < 0 || k > n` ⇒ #NUM!
           ＝★★当てずっぽうでは なく「既に 測れて いる 隣に 合わせた」★★
           ⇒★次の 枠で 聞く 事に 名指しで 載せて あります★
             `docs/measured/kansuu46/junretsu-kiku-koto.md`
       ★切り捨ては COMBIN と 同じ★（`=COMBIN(2.5,1)` → 2 が 実測） */
    PERMUT: function (引数たち, 手, 所) {
      return 数ふたつ(引数たち, 手, 所, function (a, b) {
        var n = 切り捨て(a), k = 切り捨て(b);
        if (n < 0 || k < 0 || k > n) return null;
        var r = 1;
        for (var i = 0; i < k; i++) r *= (n - i);
        return r;
      });
    },
    PERMUTATIONA: function (引数たち, 手, 所) {
      return 数ふたつ(引数たち, 手, 所, function (a, b) {
        var n = 切り捨て(a), k = 切り捨て(b);
        if (n < 0 || k < 0) return null;
        return Math.pow(n, k);
      });
    },

    /* ══ ★★真偽の 一族★★ ══（2026-09-15）
       ★物差し★ `golden-346-2026-09-08.tsv`
       ★実Excel に 聞いて 初めて 分かった 事★
         `=OR("あ",1)` → ★TRUE★（★数に 読めない 字は ★飛ばす★＝#VALUE! に しない★）
         `=OR("101",2)` → TRUE（数に 読める 字は 数に する）
         `=OR(A1:B5,2,FALSE)` → TRUE
       ★マスの 字・真偽・空は 飛ばす★（四角の 決めと 同じ）
       ★1つも 拾えなければ #VALUE!★ */
    OR: function (引数たち, 手) { return 真偽の一族(引数たち, 手, 'OR'); },
    XOR: function (引数たち, 手) { return 真偽の一族(引数たち, 手, 'XOR'); },
    AND: function (引数たち, 手) { return 真偽の一族(引数たち, 手, 'AND'); },
    NOT: function (引数たち, 手, 所) {
      if (引数たち.length !== 1) return 誤('#VALUE!');
      var 判 = 真偽にする(場.ひとつに(引数たち[0], 所 && 所.今のマス));
      if (判.誤) return 判.誤;
      return 計.真偽(!判.真);
    },

    /* ══ ★★A の 付く 一族★★ ══（2026-09-15）
       ★マスの 字と 真偽を 飛ばさない★＝`Aで数を拾う`（上） */
    AVERAGEA: function (引数たち, 手) {
      var 拾 = Aで数を拾う(引数たち, 手);
      if (拾.誤) return 拾.誤;
      var a = 拾.数たち;
      if (!a.length) return 誤('#DIV/0!');
      var 和 = 0;
      for (var i = 0; i < a.length; i++) 和 += a[i];
      return 数(和 / a.length);
    },
    MAXA: function (引数たち, 手) {
      var 拾 = Aで数を拾う(引数たち, 手);
      if (拾.誤) return 拾.誤;
      return 数(拾.数たち.length ? Math.max.apply(null, 拾.数たち) : 0);
    },
    MINA: function (引数たち, 手) {
      var 拾 = Aで数を拾う(引数たち, 手);
      if (拾.誤) return 拾.誤;
      return 数(拾.数たち.length ? Math.min.apply(null, 拾.数たち) : 0);
    },
    VARA: function (引数たち, 手) { return A散らばりで(引数たち, 手, false, false); },
    VARPA: function (引数たち, 手) { return A散らばりで(引数たち, 手, true, false); },
    STDEVA: function (引数たち, 手) { return A散らばりで(引数たち, 手, false, true); },
    STDEVPA: function (引数たち, 手) { return A散らばりで(引数たち, 手, true, true); },

    /* ★MATCH★ … 並びの 中で 何番目か（1 から）
       ★型★ … 0＝ぴたり ／ ＞0＝以下で 一番 後ろ（二分探索） ／ ＜0＝以上で 一番 後ろ（降順）
       ★見つからなければ #N/A★（★これが MATCH の 本番★＝客は 探して 無い事の 方が 多い） */
    MATCH: function (引数たち, 手, 所) {
      var 今 = 所 && 所.今のマス;
      if (引数たち.length < 2) return 誤('#VALUE!');
      var 探 = 場.ひとつに(引数たち[0], 今);
      if (!探) return 誤('#VALUE!');
      if (探.型 === '誤') return 探;                    /* 実測 `=MATCH(1/0,2)` → #DIV/0! */
      var 型 = 1;
      if (引数たち.length > 2) {
        var t = 数の引数(引数たち[2], 手, 所);
        if (t.誤) return t.誤;                          /* 実測 `=MATCH(3,A1:A5,"x")` → #VALUE! */
        型 = t.数;
      }
      var 向 = 型 > 0 ? 1 : (型 < 0 ? -1 : 0);          /* ★符号だけ★（実測 2→1・-0.5→-1） */
      var 元 = 引数たち[1];
      if (元.種 === '四角') {
        /* ★2次元の 表は 探せません★（★未測定★＝棚。実Excel が 何を 返すか 見て いない） */
        if ((元.行数 || 0) > 1 && (元.列数 || 0) > 1) return 誤('#N/A');
        var 並び = 元.並び;
        if (向 === 0) {
          for (var i = 0; i < 並び.length; i++) if (当たるか(探, 並び[i])) return 数(i + 1);
          return 誤('#N/A');
        }
        var n = (向 > 0) ? 二分で探す(並び, 探) : 降順で探す(並び, 探);
        return n ? 数(n) : 誤('#N/A');
      }
      /* ★式に 直に 書いた 1つの 値を 表として 渡した 時★
         ★実測 10本＝「ぴたり」以外は 当たりません★（型 1 でも 以下で 拾わない）
           `=MATCH(2.5,1)` → #N/A ／ `=MATCH(D1,2)` → #N/A ／ `=MATCH(12,-10,100)` → #N/A
         ★字を 表に して 大小を 比べると #VALUE!★（実測 3本）
           `=MATCH("あ","い")` → #VALUE! ／ `=MATCH(A1:A5,"<3")` → #VALUE!
         ★合わない 1本を 隠しません★ … `=MATCH(A1:A5,1,0)` は 実Excel ★#N/A★／うち ★1★
           （★探す値が 表の 時＝溢れの 話＝土台⑤ の 先★。棚に 書いて 在ります） */
      if (元.種 === '直' && 元.値 && 元.値.型 === '字' && 向 !== 0) return 誤('#VALUE!');
      if (元.種 === '直') return 当たるか(探, 元.値) ? 数(1) : 誤('#N/A');
      /* 1マスを 指した 物（A1）は ★1つだけの 表★として ふつうに 探す */
      var ひとつ = [元.値];
      if (向 === 0) return 当たるか(探, ひとつ[0]) ? 数(1) : 誤('#N/A');
      var m = (向 > 0) ? 二分で探す(ひとつ, 探) : 降順で探す(ひとつ, 探);
      return m ? 数(m) : 誤('#N/A');
    },

    /* ══ ★★XMATCH★★ ══（2026-09-18）
       ★前は お客さんの 画面で #NAME? でした★（台に 無い 60個の うちの 1個）

       ★★紙（実Excel の 実測・12本）★★
         `docs/measured/golden-kansuu-7kaime-2026-09-18.tsv`（★材料 C1:C5 ＝ 1,3,5,7,9★）
         ★★12本 とも 見込み どおりでした★★（★当て推量が 0★）
           =XMATCH(5,C)        → ★3★
           =XMATCH(4,C)        → ★#N/A★   ⇒★既定は ぴたりだけ（0）★
           =XMATCH(4,C,-1)     → ★2★     ⇒★-1 ＝ 次に 小さい★（3）
           =XMATCH(4,C,1)      → ★3★     ⇒★1 ＝ 次に 大きい★（5）
           =XMATCH(4,C,0)      → ★#N/A★
           =XMATCH(0,C,-1)     → ★#N/A★  ⇒★下に 無い★
           =XMATCH(10,C,1)     → ★#N/A★  ⇒★上に 無い★
           =XMATCH(9,C,0,-1)   → ★5★     ⇒★-1 ＝ 後ろから★
           =XMATCH(9,C,0,2)    → ★5★     ⇒★2 ＝ 二分（昇順）★
           =XMATCH(9,C,0,-2)   → ★#N/A★  ⇒★-2 ＝ 二分（降順）を 昇順に 当てた★
           =XMATCH(3,C,2)      → ★2★     ⇒★2 ＝ 型紙。数なら ぴたりと 同じ★
           =XMATCH(4,C,3)      → ★★#N/A★★ ⇒★★無い 型は #VALUE! では なく #N/A★★

       ★★測って いない 事（★書かない 見張りは「全部 守った」と 読まれる★）★★
         ・★字の 型紙★（`=XMATCH("*b*",…,2)`）… ★今回は 数だけ 聞きました★
         ・★2次元の 表★／★空の 並び★／★誤りを 含む 並び★
         ・★検索型に 無い 値★（3 など）… ★一致型に 合わせて #N/A に して います★
         ・★一致型 -1／1 と 検索型 -1／2／-2 の 組み合わせ★
         ⇒★★これらは `docs/measured/kansuu46/xmatch-kiku-koto.md` に 名指しで★★ */
    XMATCH: function (引数たち, 手, 所) {
      var 今 = 所 && 所.今のマス;
      if (引数たち.length < 2) return 誤('#VALUE!');
      var 探 = 場.ひとつに(引数たち[0], 今);
      if (!探) return 誤('#VALUE!');
      if (探.型 === '誤') return 探;
      var 一致 = 0, 検索 = 1;
      if (引数たち.length > 2) {
        var m1 = 数の引数(引数たち[2], 手, 所);
        if (m1.誤) return m1.誤;
        一致 = 切り捨て(m1.数);
      }
      if (引数たち.length > 3) {
        var m2 = 数の引数(引数たち[3], 手, 所);
        if (m2.誤) return m2.誤;
        検索 = 切り捨て(m2.数);
      }
      /* ★無い 型は #N/A★（★実測 `=XMATCH(4,C1:C5,3)` → #N/A★）
         ★検索型の 方は 測って いません★＝★一致型に 合わせて います★ */
      /* ★★一致型と 検索型で ★誤りが 違います★★★（2026-09-18・8枠目で 実測）
           `=XMATCH(4,C1:C5,3)`   -> ★#N/A★   （★一致型が 無い★）
           `=XMATCH(4,C1:C5,0,7)` -> ★#VALUE!★（★検索型が 無い★）
           ⇒★同じに して いました★＝★実測で 分かれました★ */
      if (一致 !== 0 && 一致 !== -1 && 一致 !== 1 && 一致 !== 2) return 誤('#N/A');
      if (検索 !== 1 && 検索 !== -1 && 検索 !== 2 && 検索 !== -2) return 誤('#VALUE!');

      var 元 = 引数たち[1], 並び;
      if (元.種 === '四角') {
        /* ★★2次元は #VALUE!★★（2026-09-18・8枠目で 実測）
             `=XMATCH(3,A1:B5)` -> ★#VALUE!★（★#N/A では ない★）
             ⇒★前は「測って いない」と 書いて #N/A に して いました★ */
        if ((元.行数 || 0) > 1 && (元.列数 || 0) > 1) return 誤('#VALUE!');
        並び = 元.並び;
      } else {
        並び = [元.値];
      }
      if (!並び || !並び.length) return 誤('#N/A');

      /* ★ぴたり（0）と 型紙（2）★ … `当たるか` は 字の 時だけ 型紙を 見ます */
      if (一致 === 0 || 一致 === 2) {
        if (検索 === 2 || 検索 === -2) {
          /* ★二分★ … 並びが その向きに 揃って いなければ 見つかりません（実測 -2 → #N/A） */
          var lo = 1, hi = 並び.length;
          while (lo <= hi) {
            var mid = Math.floor((lo + hi) / 2);
            var c = 場.比べる(並び[mid - 1], 探);
            if (c === 0) return 数(mid);
            if (c === null) return 誤('#N/A');
            var 右へ = (検索 === 2) ? (c < 0) : (c > 0);
            if (右へ) lo = mid + 1; else hi = mid - 1;
          }
          return 誤('#N/A');
        }
        if (検索 === -1) {
          for (var j = 並び.length - 1; j >= 0; j--) if (当たるか(探, 並び[j])) return 数(j + 1);
          return 誤('#N/A');
        }
        for (var i = 0; i < 並び.length; i++) if (当たるか(探, 並び[i])) return 数(i + 1);
        return 誤('#N/A');
      }

      /* ══ ★★二分（±2）の 時は ★二分の 動きが そのまま 出ます★★★ ══（2026-09-18）
           ★実Excel★ `=XMATCH(4,C1:C5,1,-2)` -> ★5★（C ＝ 1,3,5,7,9＝★昇順★）
             ＝★-2 は「降順だと 思って 二分する」★
             ＝★降順で ない 並びに 当てると 端まで 進みます★
           ★前★ 「一番 近い 物」を 前から 探して いた ⇒ 3
           ★後★ ★二分の まま 最後の 見込みを 覚える★ ⇒ 5
           ★これは LOOKUP で 学んだ のと 同じ 型★
             ＝★「揃って いない 時は 決まって いない」は ★私の 見立て★★
             ＝★二分の 動きは 決まって います★ */
      if (検索 === 2 || 検索 === -2) {
        var lo2 = 1, hi2 = 並び.length, 当2 = 0;
        while (lo2 <= hi2) {
          var mid2 = Math.floor((lo2 + hi2) / 2);
          var c2 = 場.比べる(並び[mid2 - 1], 探);
          if (c2 === 0) return 数(mid2);
          if (c2 === null) return 誤('#N/A');
          if (一致 === -1 && c2 < 0) 当2 = mid2;       /* 次に 小さい */
          if (一致 === 1 && c2 > 0) 当2 = mid2;        /* 次に 大きい */
          var 右へ2 = (検索 === 2) ? (c2 < 0) : (c2 > 0);
          if (右へ2) lo2 = mid2 + 1; else hi2 = mid2 - 1;
        }
        return 当2 ? 数(当2) : 誤('#N/A');
      }

      /* ★-1 ＝ 次に 小さい／1 ＝ 次に 大きい★（★前から／後ろから 見る★）
         ★並びの 向きに 関わらず「一番 近い 物」を 選びます★
           （★実測は 昇順の 並びだけ★＝★向きが 違う 時は 測って いません★） */
      var 当番 = 0, 当値 = null;
      var 始 = (検索 === -1) ? 並び.length - 1 : 0;
      var 終 = (検索 === -1) ? -1 : 並び.length;
      var 歩 = (検索 === -1) ? -1 : 1;
      for (var k = 始; k !== 終; k += 歩) {
        var v = 並び[k];
        if (!v || v.型 === '誤' || v.型 === '空') continue;
        var d = 場.比べる(v, 探);
        if (d === null) continue;
        if (d === 0) return 数(k + 1);                 /* ぴたりが 在れば それ */
        if (一致 === -1 && d > 0) continue;            /* 大きいのは 見ない */
        if (一致 === 1 && d < 0) continue;             /* 小さいのは 見ない */
        if (当値 === null) { 当番 = k + 1; 当値 = v; continue; }
        var e = 場.比べる(v, 当値);
        if (e === null) continue;
        /* -1 なら より 大きい 方が 近い ／ 1 なら より 小さい 方が 近い */
        if ((一致 === -1 && e > 0) || (一致 === 1 && e < 0)) { 当番 = k + 1; 当値 = v; }
      }
      return 当番 ? 数(当番) : 誤('#N/A');
    },

    /* ★INDEX★ … ★値では なく 参照を 返します★（実測 `=ISBLANK(INDEX(E1:E5,4))` → TRUE）
       ★0 は「まるごと」★（行 0＝その 列 全部／列 0＝その 行 全部）
       ★転け方は 2種類★ 行/列が 外＝#REF! ／ 負・字＝#VALUE! */
    INDEX: function (引数たち, 手, 所) {
      if (引数たち.length < 2) return 誤('#VALUE!');   /* 実Excel は 式として 受け付けない */
      var 元 = 引数たち[0], R, C, 並び, 番地;
      if (元.種 === '四角') {
        R = 元.行数 || 0; C = 元.列数 || 0; 並び = 元.並び; 番地 = 元.番地 || null;
      } else {
        R = 1; C = 1; 並び = [元.値]; 番地 = 元.番地 || null;
      }
      if (!R || !C) return 誤('#REF!');
      if (R === 1 && C === 1 && 並び[0] && 並び[0].型 === '誤') return 並び[0];

      /* ★数に ならない 引数が 先★（実測 `=INDEX(D1,D2,"D")` → #REF! では なく ★#VALUE!★） */
      var g = 数の引数(引数たち[1], 手, 所); if (g.誤) return g.誤;
      var 行 = 切り捨て(g.数), 列 = null;
      if (引数たち.length > 2) {
        var h = 数の引数(引数たち[2], 手, 所); if (h.誤) return h.誤;
        列 = 切り捨て(h.数);
      }
      if (引数たち.length > 3) {
        var k = 数の引数(引数たち[3], 手, 所); if (k.誤) return k.誤;
        var 区 = 切り捨て(k.数);
        if (区 < 1) return 誤('#VALUE!');              /* 実測 `=INDEX(A1,A1:B5,2,FALSE)` → #VALUE! */
        if (区 > 1) return 誤('#REF!');                /* 実測 `=INDEX(A1:B5,1,1,2)` → #REF! */
      }
      if (列 === null) {
        if (C === 1) 列 = 1;
        else if (R === 1) { 列 = 行; 行 = 1; }         /* 横1本の 表＝番号は 列の 事 */
        else return 誤('#REF!');                       /* 実測 `=INDEX(A1:B5,3)` → #REF! */
      }
      if (行 < 0 || 列 < 0) return 誤('#VALUE!');       /* 実測 `=INDEX(A1:A5,-1)` → #VALUE! */
      if (行 > R || 列 > C) return 誤('#REF!');         /* 実測 `=INDEX(A1:A5,6)` → #REF! */

      var r0 = (行 === 0) ? 1 : 行, 行数 = (行 === 0) ? R : 1;
      var c0 = (列 === 0) ? 1 : 列, 列数 = (列 === 0) ? C : 1;
      var 出 = [];
      for (var i = 0; i < 行数; i++)
        for (var j = 0; j < 列数; j++)
          出.push(並び[(r0 - 1 + i) * C + (c0 - 1 + j)] || 計.空);
      if (!番地) {
        /* ★番地が 無い＝式が 作った 表★（`{1;2;3}` や 直に 書いた 値）
           実測 `=INDEX("あ",1)` → あ ／ `=INDEX(0.5,1)` → 0.5
           ★まるごと 取った 時★（2026-09-15 に 測って 埋めた＝棚 ⑬④）
             `=INDEX({1;2;3},0)` → ★1★（先頭） ／ `=SUM(INDEX({1;2;3},0))` → ★6★（全部）
           ⇒★行・列を 0 に した 参照★＝★1つの 値が 要る 所では 先頭・関数へは まるごと★ */
        if (行数 === 1 && 列数 === 1) return 出[0];
        return 場.参照(0, 0, 行数, 列数, 出, true);   /* ★番地なし★（0 は A1 の 事なので 印で 持つ） */
      }
      return 場.参照(番地.行 + r0 - 1, 番地.列 + c0 - 1, 行数, 列数, 出);
    },

    /* 数が 1つも 無ければ 0（実測 `=SUM(B1:B3)` → 0） */
    SUM: function (引数たち, 手) {
      var 拾 = 数を拾う(引数たち, 手);
      if (拾.誤) return 拾.誤;
      return 足しあげる(拾.数たち, 手);
    },
    /* ★数が 1つも 無ければ #DIV/0!★（SUM は 0 なのに 平均は 誤り＝実測） */
    AVERAGE: function (引数たち, 手) {
      var 拾 = 数を拾う(引数たち, 手);
      if (拾.誤) return 拾.誤;
      if (拾.数たち.length === 0) return 誤('#DIV/0!');
      var 合 = 足しあげる(拾.数たち, 手);
      if (合.型 === '誤') return 合;
      return 計.つなぎ('/', 合, 数(拾.数たち.length), 手);
    },
    PRODUCT: function (引数たち, 手) {
      var 拾 = 数を拾う(引数たち, 手);
      if (拾.誤) return 拾.誤;
      return 掛けあげる(拾.数たち, 手);
    },
    /* ★数が 1つも 無ければ 0★。★在れば その中の 一番★（負だけでも 0は 混ざらない＝実測） */
    MAX: function (引数たち, 手) {
      var 拾 = 数を拾う(引数たち, 手);
      if (拾.誤) return 拾.誤;
      if (拾.数たち.length === 0) return 数(0);
      var m = 拾.数たち[0];
      for (var i = 1; i < 拾.数たち.length; i++) if (拾.数たち[i] > m) m = 拾.数たち[i];
      return 数(m);
    },
    MIN: function (引数たち, 手) {
      var 拾 = 数を拾う(引数たち, 手);
      if (拾.誤) return 拾.誤;
      if (拾.数たち.length === 0) return 数(0);
      var m = 拾.数たち[0];
      for (var i = 1; i < 拾.数たち.length; i++) if (拾.数たち[i] < m) m = 拾.数たち[i];
      return 数(m);
    },
    /* ★COUNT は 誤りを 数えない★（実測 `=COUNT(A1:A6)` → 2）
       ★誤りが 在っても 誤りを 返さない★＝数を 拾う 道とは 別（だから ここは 自分で 回す） */
    COUNT: function (引数たち, 手) {
      var n = 0;
      ほどく(引数たち, function (v, 元) {
        if (v.型 === '誤') return true;               /* ★飛ばす（伝えない）★ */
        if (元 === 'マス') { if (v.型 === '数') n++; return true; }
        var x = 計.数にする(v, 手);                    /* 直に 書いた 物は 数に して みる */
        if (x.型 === '数') n++;
        return true;
      });
      return 数(n);
    },
    /* ★COUNTA は 誤りを 数える★（実測 `=COUNTA(A1:A6)` → 5＝COUNT と 違う） */
    COUNTA: function (引数たち) {
      var n = 0;
      ほどく(引数たち, function (v) {
        if (v.型 !== '空') n++;
        return true;
      });
      return 数(n);
    }
  };

  /* ★知らない 名前は 呼ばない★＝★半分 合う 計算を 出さない★ */
  /* ══ ★★範囲を 渡されたら 1マスずつ 呼び直す 関数★★ ══（2026-09-21）
       ★なぜ 名簿か★
         ★範囲を 受ける 決まりの 関数★（`SUM` `AVERAGE` 等）は 溢れては いけません。
         ★１つの 値を 受ける 決まりの 関数★（`DAY` `LEN` 等）だけ 溢れます。
         ⇒★どちらかは 実Excel が 決めて います★
         ⇒★だから 実Excel に 聞いた 紙に 「溢れる」と 在る 名前だけ★ 包みます
       ★物差し★ `docs/measured/golden-jitsu-excel-hanni-zenbu-2026-09-20.tsv`
         （実Excel 16.0 build 20326 が 2マス 以上 返した 式）
       ★漏れたら どう なるか★
         ★門 `tests/kobore-nakami.test.mjs` が 数で 見て います★
         ＝合う 数が 減っても ★増えても★ 赤に なります
       ★ここで 包まない 物★
         ・既に 溢れて 戻って くる 物（`見分け` `数ひとつ` `字ひとつ` で 直した 分）
         ・台が 知らない 物（`GROUPBY` `PIVOTBY` ＝ピボットが まだ 無い） */
  var 溢れる名簿 = {
    ARABIC: 1, CHAR: 1, CODE: 1, DATEVALUE: 1, DAY: 1, HOUR: 1, HYPERLINK: 1,
    ISOWEEKNUM: 1, LEFT: 1, LEN: 1, LOG: 1, MINUTE: 1, MONTH: 1,
    NORMSINV: 1, NOT: 1, RIGHT: 1, ROMAN: 1, SECOND: 1, TIMEVALUE: 1,
    TRUNC: 1, UNICHAR: 1, UNICODE: 1, VALUE: 1, WEEKDAY: 1, YEAR: 1
  };
  /* ══ ★★値では なく ★場所★を 渡す 名簿★★ ══（2026-09-21）
       ★上の 名簿と 何が 違うか★
         上 ＝ `{ 種:'直', 値: 中身 }`＝★中身だけ★ 渡す（`LEN` `DAY` 等）
         下 ＝ `{ 種:'マス', 番地, 打った字 }`＝★どのマスか★ 渡す
       ★`ISFORMULA` は 値を 1度も 見ません★＝★打って ある 字★を 見ます
         ⇒上の 名簿に 入れると ★3マス 全部 `#VALUE!`★ に なります（2026-09-21 実測）
       ★`ROW` `COLUMN` は ここには 入れません★
         ＝★マスごとでは なく ★形の ぶんだけ★ 溢れる★ ので
           `番地の番号()` の 中で 見て います（そこの 注を 読んで ください）
       ★★`FORMULATEXT` も ここです★★（2026-09-21 経営者1 が 実Excel で 測りました）
         `=FORMULATEXT(A1:A3)` ⇒ ★3マス★ `#N/A` / `=1+1` / `#N/A`
           （A2 だけ 式＝`=1+1`。★式で ない マスは `#N/A`★＝`#VALUE!` では ない）
         ⇒★`ISFORMULA` と 同じ 形★です。
       ★★`CELL` は 入れません★★（★入れると 逆に 壊れます★）
         `=CELL("row",A1:A3)` ⇒ ★1マス 1★＝★溢れません★（同じ 日に 測りました）
         ＝★「場所を 見る 関数だから 溢れる」では ありません★
           ＝★1つずつ 実Excel に 聞くしか ありません★
       ★物差し★ `docs/measured/golden-jitsu-excel-basho-kansuu-2026-09-21.tsv` */
  var 場所で溢れる名簿 = { ISFORMULA: 1, FORMULATEXT: 1 };

  function 呼ぶ(名前, 引数たち, 手, 所) {
    var 大 = String(名前).toUpperCase();
    var f = 表[大];
    if (!f) return null;              /* ★null＝この 台は 知らない（借り物に 回す）★ */
    /* ★名簿に 在り、引数が 1つで、それが 範囲なら 1マスずつ 呼び直す★ */
    if (溢れる名簿[大] && 引数たち && 引数たち.length === 1 && 範囲か(引数たち[0])) {
      var 溢 = マスごとに溢れさせる(引数たち[0], 所 || {}, function (v) {
        return 出口の門(f([{ 種: '直', 値: v }], 手, 所 || {}), 名前);
      });
      if (溢) return 溢;
    }
    /* ★中身では なく ★場所★を 1つずつ 渡す 物★（`ISFORMULA`） */
    if (場所で溢れる名簿[大] && 引数たち && 引数たち.length === 1 && 範囲か(引数たち[0])) {
      var 溢場 = 場所ごとに溢れさせる(引数たち[0], 所 || {}, function (m) {
        return 出口の門(f([m], 手, 所 || {}), 名前);
      });
      if (溢場) return 溢場;
    }
    return 出口の門(f(引数たち, 手, 所 || {}), 名前);
  }

  /** ★★出口の 門★★（2026-09-16・経営者1 の 注文）
   *   ★★`NaN` と 無限を 1マスも 通さない★★
   *
   *   ★なぜ 1か所に するか★
   *     ★`NaN` は ★誤りに ならない★★＝お客さんは 気づきません
   *       ⇒★そのまま 足し算に 入り 合計が 全部 `NaN` に なる★
   *       ⇒★気づくのは 締めの 後★＝`#VALUE!` より ずっと 悪い
   *     ★今日 3つ 出ました★ … `TRUNC`／`PMT`／`IMSECH`（どれも 大きい 数の `Infinity` から）
   *     ★208個 書いて 門は 20か所★＝★1個ずつ 覚えて 書くのは もたない★
   *   ⇒★★決まりを 1か所で 守る★★（★見張りは 決まりを 見る★と 同じ 形）
   *
   *   ★先に 測りました★ … ★手元の 紙 全部（60枚）で 実Excel が 無限・`NaN` を 返した 行は ★0行★★
   *     ⇒★★門で 正しい 答えを 潰しません★★
   *
   *   ★何に するか★ … `#NUM!`（★実Excel も 数に ならない 時は これ★）
   */
  function 出口の門(出, 名前) {
    if (出 && 出.型 === '数' && !isFinite(出.値)) return 誤('#NUM!');
    if (出 && 出.溢れ === true && 出.並び) {
      for (var r = 0; r < 出.並び.length; r++) {
        var 段 = 出.並び[r];
        for (var c = 0; c < 段.length; c++) {
          if (段[c] && 段[c].型 === '数' && !isFinite(段[c].値)) 段[c] = 誤('#NUM!');
        }
      }
    }
    return 出;
  }

  return {
    呼ぶ: 呼ぶ, 表: 表,
    数を拾う: 数を拾う, ほどく: ほどく,
    足しあげる: 足しあげる, 掛けあげる: 掛けあげる
  };
});
