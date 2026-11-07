/* shiki-hyou.js — ★表を 持ち・依存を 追い・変わった所だけ 計算し直す★（2026-09-13）
 *
 *  ★★土台を 自分で 作る ⑥枚目＝★本体★★★
 *    土台は ①字に切る ②形にする ③頼りの地図 ④順番と計算 ⑤溢れ、
 *    そして ★この 本体（表を 持つ・変わったら 直す）★です。
 *
 *  ★★なぜ ここを 先に やるか（2026-09-13 実測）★★
 *    借り物（HyperFormula・944,635字・GPLv3）の 中身を 2人で 別々に 測ったら：
 *      ★関数の 実装 … 423,496字 ＝ 44.8%★
 *      ★本体       … 521,139字 ＝ 55.2%★
 *    ⇒★関数を 519個 全部 作っても まだ 半分★。★本体の 方が 大きい★。
 *    ⇒ だから ★本体の 一番 小さい 一周★ を 先に 作って ★規模を 測ります★。
 *
 *  ★★この 台が やる 事（★見ていない範囲も 先に 書く★）★★
 *    やる  … ①マスに 値／式を 置く ②式を 計算する ③★誰が 誰を 見ているか 覚える★
 *            ④★1マス 変えたら その マスを 見ている 式だけ★ 計算し直す
 *            ⑤★輪（A1→A2→A1）を 見つけて #REF! に する★
 *    ★やらない★ … 溢れ（土台⑤）／行の 挿入で 式を 追従させる／元に 戻す／
 *                  複数の 板（シート）／関数は ★shiki-kansuu.js が 知っている 7個だけ★
 *    ⇒★出来て いない 物を 出来た 顔で 混ぜない★
 *
 *  ★値の 形★ `shiki-keisan.js` と 同じ { 型:'数'|'字'|'真偽'|'誤'|'空', 値 }
 */
/* ★どちらが 答えたかの 数★（台ごと 1つ。`答えた数()` で 見る） */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(
    typeof require === 'function' ? require('./shiki-kiru.js') : root.ShikiKiru,
    typeof require === 'function' ? require('./shiki-katachi.js') : root.ShikiKatachi,
    typeof require === 'function' ? require('./shiki-sansho.js') : root.ShikiSansho,
    typeof require === 'function' ? require('./shiki-keisan.js') : root.ShikiKeisan,
    typeof require === 'function' ? require('./shiki-kansuu.js') : root.ShikiKansuu,
    typeof require === 'function' ? require('./shiki-tsunagi.js') : root.ShikiTsunagi,
    typeof require === 'function' ? require('./shiki-afure.js') : root.ShikiAfure,
    typeof require === 'function' ? require('./shiki-basho.js') : root.ShikiBasho);
  else root.ShikiHyou = factory(root.ShikiKiru, root.ShikiKatachi, root.ShikiSansho,
    root.ShikiKeisan, root.ShikiKansuu, root.ShikiTsunagi, root.ShikiAfure, root.ShikiBasho);
})(typeof self !== 'undefined' ? self : this, function (切, 形, 参, 計, 関, 皮, 溢, 場) {
  'use strict';

  /* ★どちらが 答えたかを 1本ずつ 数える★（2026-09-14・経営者1 と 決めた 見張り①）
     ★訳★＝377個を 書いて いく 間、★借り物の 出番が 減る様子を 数で 見る★為。
       数が 減らなければ ★「書いたのに 誰も 呼んで いない」★と すぐ 分かる。 */
  var 答えた = {};
  /* ★名前を わざと 長く して います★（2026-09-14）
     ＝`数える` という 名前は ★repo に 既に 4つ 在ります★
       （ribbon.js／ribbon-keytips.js／ribbon-launcher.js／ribbon-context-spec.js）
     ★`tools/unused-param.mjs` は「名前で 呼ぶ」形も 見ます★ので、
     ここで `数える(誰)` と 呼ぶと ★向こうの 4つが「渡されている」ように 見え★、
     ★見張りの 候補 15個 → 11個★に なって いました＝★穴を 隠した★。
     ⇒ ★repo に 在る 名前を そのまま 使わない★ */
  function 答えた者を数える(誰) { 答えた[誰] = (答えた[誰] || 0) + 1; }

  /* ══ ★マスの 名前（A1）と 番地（行,列）を 行き来する★ ══ */
  /* ★★板の 名前を 切り出す★★（2026-09-15・㍸）
     ★`!` は ★後ろから★ 探す★＝板の 名前に `!` は 入れられない（Excel）ので
       ★`#REF!` を 含む 形でも 最後の `!` が 区切り★。
     ★`#REF!`（後ろが 空）は 番地に ならない★ので 下で null に なります。
     ★`#REF!A1`（Excel で 板を 消した 形）は 板「#REF」★
       ⇒★知らない 板★なので `読む` が #REF! を 返します（★実物には 0本★） */
  function 板を切る(名) {
    var s = String(名);
    var i = s.lastIndexOf('!');
    if (i < 0) return { 板: null, 芯: s };
    var 板 = s.slice(0, i);
    /* ★★引用符で 包んだ 板の 名前を 外す★★（2026-09-18・★お客さんに 出て いた 欠陥★）
         ★実Excel は ★数で 始まる 名前・空白を 含む 名前★を `'` で 包みます★
           例 `='5月'!E14` ／ `='売上 一覧'!A1`
         ⇒★包んだ ままだと ★知らない 板★に なり #REF! を 返して いました★
         ⇒★`tests/cross-sheet.test.mjs` が 捕まえました★（★34万円 少ない★ 形）
         ★Excel の 決まり★ ... 名前の 中の `'` は `''` と 2つ 書く ⇒ 1つに 戻す */
    if (板.length >= 2 && 板.charAt(0) === "'" && 板.charAt(板.length - 1) === "'") {
      板 = 板.slice(1, -1).replace(/''/g, "'");
    }
    return { 板: 板, 芯: s.slice(i + 1) };
  }
  /* ══ ★★1度 読んだ 名は 覚えて おく★★ ══（2026-09-25）
       ★★なぜ★★（経営者1 の 実測・`docs/measured/hiraku-made-2fun-no-shoutai.md`）
         司さんの 実物（板 15枚・マス 35,760・式 15,799）を 開くと:
           CPU の 記録 ... ★87.4% が この 台 1本★
           `名から番地` ... ★141,920,112 回★ 呼ばれて
                            ★別々の 字は 53,365 種類★
           ＝★同じ 字を 平均 2,659 回 読み直して いました★
       ★★効き（chromium・各 3回・中ほどで 比べる）★★
         そのまま ....... ★59,471 ms★（一番速い 55,979 ／ 一番遅い 76,919）
         覚えて おく ..... ★32,078 ms★（54%）
         ⇒差 ★27,393 ms★ ＞ 揺れ幅 ★20,940 ms★ ＝★本物の 差★
       ★★答えは 1つも 変わりません★★（9回 とも 完全一致）
         マス 35,760 ／ 式 15,799 ／ 誤 0 ／ 数 3,857 ／ 字 31,903
         数の 足し算 ★54241061.144673★

       ★★同じ 物を 返して 大丈夫か（★私が 数え直しました★）★★
         `名から番地` を 呼ぶ 所 ... ★17か所★（`lib/` `js/` `book.html` `exally-formula.js`）
         返り物を 書き換えて いる 所 ... ★0件★（`.板=` `.行=` `.列=` で 探して 0）
         ⇒★同じ 物を 返しても 誰も 壊しません★
         ⇒★★但し これは 今の 字の 話です★★＝
           後で 誰かが 書き換えると ★覚えた 物が 汚れて 別の マスが 狂います★。
           ⇒★`tests/oboeru-nakara-banchi.test.mjs` が 増えたら 赤に します★

       ★★捨てなくて よい 訳★★
         `名から番地` は ★字 → 番地 の 変換だけ★ です。
         ★板の 中身にも 板の 有る無しにも 依りません★（`'5月'!E14` → `{板:'5月',行:13,列:4}`）
         ⇒★板を 消しても 足しても 答えは 変わりません★＝★捨てる 必要が ありません★
       ★★際限なく 増えないように 上限を 付けます★★
         実物で ★53,365 種類★。★別の 本なら 別の 数★ なので、
         ★20万で 1度 空に します★（★答えは 変わりません★＝また 数え直すだけ) */
  var _覚えた番地 = null;
  var 覚えの上限 = 200000;
  function 名から番地(名) {
    var 鍵 = String(名);
    if (_覚えた番地 === null) _覚えた番地 = new Map();
    var 覚 = _覚えた番地.get(鍵);
    if (覚 !== undefined) return 覚;
    var 出 = _名から番地の元(鍵);
    if (_覚えた番地.size >= 覚えの上限) _覚えた番地.clear();
    _覚えた番地.set(鍵, 出);
    return 出;
  }
  /* ★元の 中身は 1文字も 変えて いません★（名前だけ 変えました） */
  function _名から番地の元(名) {
    var 切 = 板を切る(名);
    if (切.板 !== null && !切.板.length) return null;   /* `!A1` は 番地で は ない */
    var m = /^\$?([A-Z]+)\$?([0-9]+)$/.exec(切.芯.toUpperCase());
    if (!m) return null;
    var 列 = 0;
    for (var i = 0; i < m[1].length; i++) 列 = 列 * 26 + (m[1].charCodeAt(i) - 64);
    /* ★板の 名前は Excel と 同じく 大文字小文字を 区別しない★ */
    return { 板: (切.板 === null) ? null : 切.板.toUpperCase(), 行: +m[2] - 1, 列: 列 - 1 };
  }
  /* ★★列の 字は 覚えて おきます★★（2026-09-25）
       経営者1 の 実測（`49ca1fb`・3回）... `番地から名` は ★15.4%★・★64,777,815回★
       ⇒でも 出る 字は ★列の 番号だけで 決まります★（A B C ･･･ Z AA AB ･･･）
       ⇒★6,477万回 とも 同じ 輪を 回して 同じ 字を 作り直して いました★
       ★覚えるのは 列の 字だけ★＝板の 名と 行の 数は 毎回 繋ぎます
         （鍵を 作るのに また 字を 繋いだら 元も 子も ありません）
       ★際限なく 増えません★＝Excel の 列は 16,384 まで（`XFD`）
       ★元の 出し方は 1文字も 変えて いません★＝覚えが 無い 時に 同じ 輪を 回します */
  var _列の字 = [];
  function 列の字(列) {
    var 覚 = _列の字[列];
    if (覚 !== undefined) return 覚;
    var s = '', c = 列 + 1;
    while (c > 0) { var r = (c - 1) % 26; s = String.fromCharCode(65 + r) + s; c = ((c - r) / 26) | 0; }
    _列の字[列] = s;
    return s;
  }
  function 番地から名(行, 列, 板) {
    return (板 ? (板 + '!') : '') + 列の字(列) + (行 + 1);
  }
  /* ★この 台の 既定の 板★（★板を 書かない 名は ここ★）
     ★今までの 書き方（`打つ('A1', …)`）を 1文字も 変えない 為★ */
  var 既定の板 = '板1';

  function 表() {
    /* ★マスの 中身★ … 名前 → { 打った字, 値, 形 } */
    var 中身 = {};
    /* ══ ★★在る マスの 索引（板 → 行 → 列）★★ ══（2026-09-26）
         ★★なぜ 要るか（経営者1 の 実測）★★
           1打ち目 15.1秒の うち ★91.4%（5,700万回）が 「中身に 無い マスを 訊いて いる」★
           ＝`四角を読む` は 四角の 全部の マスを 回ります。
           ＝空の マスでも 毎回 ★⑴字を 作り ⑵中身を 引き ⑶覆いを 引く★
           ⇒★その 3つを 全部 やめれば 上限 6.7秒★（経営者1 が 測った・答えを 捨てた 数）
         ★★だから 「この マスは 空だ」を 字を 作らずに 言える ように します★★
         ★★消す 所が 無いのが 効いて います★★
           `delete 中身[･･･]` は ★1件も 在りません★（字で 数えた）
           ⇒★索引は 増えるだけ＝「在るのに 無い」に なりません★
           ⇒★覆い（溢れ）は 消えます★ ので そこだけ 引きます
         ★★入れる 道は 2本 在ります★★＝`打つ` と `置く`
           ＝★1回 `置く` を 落として `tests/shiki-sagasu.test.mjs` が 赤に しました★
           ＝記憶「★作る道が 2本 在る 時は 両方 直せ★」
         ★★一番 怖いのは 更新漏れ★★＝★「在るのに 空」と 答える★
           ＝★誤りに ならない／画面も 崩れない／数だけ 違う★
           ⇒★だから `索引を照らす()` を 先に 作りました★
         ★見張り★ tests/shiki-hyou-sakuin.test.mjs */
    var 在る = {};          /* 板U → 行 → { 列: true } */
    var 空と答えた = 0;     /* ★字を 作らずに 空を 返した 回数（効いた 回数）★ */
    /* ══ ★★入口ごとの 回数★★（2026-10-05・1打ち目の 固まりを 割る 為）══
         ★割合（CPU の 記録）だけでは 直し方は 決まらない★＝回数と「別々の 種類」まで 数える
         ・四角を読む ･･･ 呼ばれた 回数／中の マス／索引に 在った マス
         ・読む ･･･ 1マスの 読み（外から・式の 名）
         ・揃った名で ･･･ 呼ばれた 回数／中身に 無かった（空 か 覆い）回数
         ・計算する ･･･ 式を 計算した 回数
         ★四角の 種類★（同じ 四角を 何回 読んだか）は 重いので `数えを始める(true)` の 時だけ 取る
         ★加算 だけ★＝答えは 1つも 変えない */
    var 数え = { 四角を読む: 0, 四角のマス: 0, 四角で在った: 0, 読む: 0, 揃った名で: 0, 揃った名で無かった: 0, 計算する: 0, 四角の控えで: 0, 帯を見た: 0 };
    var 四角の種類 = null;
    /* ★まとめて 打つ 間★（2026-10-05）＝打った 名を 覚えるだけで 1つずつは 直さない／終わりに 1回 直す
         `まとめの覆い` ･･･ その 間に 溢れで 埋まった マス（終わりに それを 見る 式を もう1回 直す） */
    var まとめ中 = null, まとめの覆い = null;
    /* ★塞がれた 溢れ★（2026-10-05・経営者の 物差し 137 の 組 J・M）＝{ 式の 名: { 元: 番地, 行数, 列数 } }
         溢れる 先が 塞がって #SPILL! に なった 式と、溢れる はずだった 範囲。
         ⇒ その 範囲に 打ったら（塞いだ 物を 消したら）式を 計算し直す＝★消せば 戻る★（実Excel J） */
    var 塞がれた = {};
    function 索引に足す(n) {
      端を変える();   /* ★帯の 端の 覚えを 捨てる★（打つ・置く・溢れ は 皆 ここを 通る） */
      var ば = 名から番地(n);
      if (!ば) return;
      var 板U = ば.板 ? String(ば.板).toUpperCase() : '';
      var 行ら = 在る[板U] || (在る[板U] = {});
      var 列ら = 行ら[ば.行] || (行ら[ば.行] = {});
      列ら[ば.列] = true;
    }
    /* ══ ★★帯の 端★★（2026-11-04・案X・設計前 taiketsu 止め3・4 で 直した）══
         ★帯の 列（行）の 中で 空で ない 一番 下（右）の マス★まで 切って 読む。
         ★板で 1つに しない★ ... 関係ない 長い 列（Z1:Z50000）で 1打ち 7.2秒（書いた 範囲 0.16秒）に なった。
         ★消した マスは 数えない★ ... 一度 Q200000 に 打って 消すと ずっと 重いまま だった。
         ★端は 帯の 列の 中身だけで 決まる★＝その 列に 打てば 帯の 見張りで 式は 必ず 直る
           （端が 帯の 外の マスで 変わると、直されない 式が 古い 答えの まま 残った＝MATCH 型1 の 二分探索）。
         ★覚え★ 中身が 変わる たびに 版を 上げ、同じ 版の 間は 同じ 帯の 端を 1回だけ 数える。 */
    var 端の版 = 0, 端の覚え = {};
    /* ★値だけの 四角の 控え★（2026-11-07・本番前 taiketsu 叩き直し 止め1）＝式も 溢れも 無い 四角は 端の版 が 変わるまで 同じ 並び
         ＝列まるごと `=VLOOKUP(Ai,$A:$B,2,0)` を 4000行 打つと 1本ごとに 8000マスを 読み直して いた（開く 18.5秒・1打ち 9.1秒）
         ★使わない 時★ ④の 残りの 間（前の覆い を 見る 為）・知らない 板・式 か 溢れ が 1つでも 在る 四角 */
    var 四角の控え = new Map();   /* ★満杯なら 一番 長く 当たって いない 物 1つだけ 捨てる★（叩き直し3 止め1＝全部 捨てると 1本の 式で 同じ 範囲を 2回 読む 形（`SUM(r)+COUNT(r)`）が 64行ごとに $A:$B を 消し 2乗の 1/64 が 残った）＝当たったら 後ろへ 回す */   /* ★同じ 版で 2回目に 読んだ 四角だけ 控える★（叩き直し 任意2＝1回きりの 広い 四角を 写すと 遅く なる） */
    /* ★1回目の 印は 別の 入れ物★（叩き直し2 止め1＝印を 同じ 64件に 数えると 行ごとに 違う 四角（`SUM(C(i-1):Ci)`）が
         64件を 埋めて $A:$B の 控えまで 捨て、2乗に 戻った）＝印は 字と 数だけで 軽い⇒上限 4096・満杯なら 印だけ 捨てる */
    var 四角の印 = {}, 四角の印の数 = 0;
    /* ★控えの 上限は 件数で なく マスの 数★（2026-11-07・wide）＝64件では 70列の INDEX/MATCH×500行（違う 四角 71個を 行ごとに 回す）が 全部 外れ 開く 15.8秒
         ＝控えた マスの 合計 200万まで（1件が それより 大きければ 控えない）・件数は 4096まで・満杯なら 一番 古い 物から 捨てる */
    var 控えのマス = 0, 控えのマスの上限 = 2000000;
    /* ★版が 変われば 控えも 印も 二度と 当たらない★（設計前 taiketsu 任意4＝古い 版の 控えが 上限いっぱい 約100MB 残った）＝その場で 空に する */
    function 端を変える() { 端の版++; if (控えのマス || 四角の控え.size) { 四角の控え = new Map(); 控えのマス = 0; } if (四角の印の数) { 四角の印 = {}; 四角の印の数 = 0; } }
    function 空でないか(n) {
      var m = 中身[n];
      return !!((m && (m.形 || String(m.打った字 || '') !== '')) || 覆い[n]);
    }
    function 帯の端(帯) {
      var 鍵 = 帯.板U + '|' + (帯.列全体 ? 'C' : 'R') + '|' + (帯.列全体 ? 帯.左 + ':' + 帯.右 : 帯.上 + ':' + 帯.下);
      var 覚 = 端の覚え[鍵];
      if (覚 && 覚.版 === 端の版) return 覚.端;
      var 行ら = 在る[帯.板U] || {}, 端 = -1, r, c;
      for (r in 行ら) {
        if (!Object.prototype.hasOwnProperty.call(行ら, r)) continue;
        var 列ら = 行ら[r], ri = +r;
        if (帯.列全体) {
          if (ri <= 端) continue;
          for (c in 列ら) {
            var ci = +c;
            if (ci >= 帯.左 && ci <= 帯.右 && 空でないか(番地から名(ri, ci, 帯.板U))) { 端 = ri; break; }
          }
        } else if (ri >= 帯.上 && ri <= 帯.下) {
          for (c in 列ら) { var cj = +c; if (cj > 端 && 空でないか(番地から名(ri, cj, 帯.板U))) 端 = cj; }
        }
      }
      端の覚え[鍵] = { 版: 端の版, 端: 端 };
      return 端;
    }
    function 索引から引く(n) {
      if (中身[n]) return;   /* ★中身に 残って いるなら 引きません★ */
      var ば = 名から番地(n);
      if (!ば) return;
      var 板U = ば.板 ? String(ば.板).toUpperCase() : '';
      if (在る[板U] && 在る[板U][ば.行]) delete 在る[板U][ば.行][ば.列];
    }
    function 在るか(板U, r, c) {
      var 行ら = 在る[板U || ''];
      if (!行ら) return false;
      var 列ら = 行ら[r];
      return !!(列ら && 列ら[c]);
    }
    /* ★★索引を 照らす★★＝★中身と 覆いから 作り直して 今の 索引と 突き合わせる★ */
    function 索引を照らす() {
      var 真 = {}, 名;
      function 入れる(nn) {
        var ば = 名から番地(nn);
        if (!ば) return;
        var 板U = ば.板 ? String(ば.板).toUpperCase() : '';
        真[板U + '|' + ば.行 + ',' + ば.列] = true;
      }
      for (名 in 中身) if (Object.prototype.hasOwnProperty.call(中身, 名)) 入れる(名);
      for (名 in 覆い) if (Object.prototype.hasOwnProperty.call(覆い, 名)) 入れる(名);
      var 今 = {}, 板U2, r2, c2;
      for (板U2 in 在る) {
        if (!Object.prototype.hasOwnProperty.call(在る, 板U2)) continue;
        for (r2 in 在る[板U2]) {
          if (!Object.prototype.hasOwnProperty.call(在る[板U2], r2)) continue;
          for (c2 in 在る[板U2][r2]) {
            if (Object.prototype.hasOwnProperty.call(在る[板U2][r2], c2)) 今[板U2 + '|' + r2 + ',' + c2] = true;
          }
        }
      }
      var 真にだけ = [], 索引にだけ = [], k;
      for (k in 真) if (!今[k]) 真にだけ.push(k);
      for (k in 今) if (!真[k]) 索引にだけ.push(k);
      return {
        合う: 真にだけ.length === 0 && 索引にだけ.length === 0,
        真の数: Object.keys(真).length, 索引の数: Object.keys(今).length,
        真にだけ: 真にだけ.slice(0, 20), 索引にだけ: 索引にだけ.slice(0, 20),
      };
    }
    /* ★頼り の 向き★ … 「A2 は A1 を 見ている」を 2つの 向きで 持つ
       ★なぜ 2つ 持つか★＝A1 が 変わった時に ★A1 を 見ている 者★を すぐ 引ける様に */
    var 見ている = {};   /* A2 → [A1, B1]（A2 が 見ている 先） */
    /* ★A1 → { A2:true }（A1 を 見ている 者）★
       ★並びでは なく 名札の 束★＝2026-09-14 に 直した。
         前は 並びで `indexOf` して いて、皆が 同じ マスを 見ると ★N の 2乗★に なった
         （実測 `=A1+1` を N マスに 打つ … N=2000 33ms ／ 4000 125ms ／ ★8000 512ms★）。 */
    var 見られている = {};
    /* ★帯の 見張り★（2026-11-04・案X）＝{ '帯|板|C|列' 等: { 式の 名: true } } ／ 見ている帯＝{ 式の 名: [鍵] } */
    var 帯の見張り = {}, 見ている帯 = {};
    /* ★帯の 見張りの 生きた 数★（2026-11-07・本番前 taiketsu 止め1）＝0件の 本（列まるごとを 使わない 本）は 見る者たち を 前と 同じ 速い 道で 返す
         ＝全部の 節で 帯の 鍵を 作って いて、式の 多い 本が 開く・1打ち とも 本番の 2.2〜2.5倍 遅かった（鎖 20000本 開く 121→300ms） */
    var 帯の見張りの数 = 0, 帯鍵の控え = {};
    /* ★動く 参照の 式★（2026-11-07）＝OFFSET・INDIRECT は 読む マスを 式の 中で 作る＝頼りの 線では 届かない
       ⇒★Excel と 同じく 揮発★＝どの マスが 変わっても 直す（{ 式の 名: true }） */
    var 揮発な式 = {};
    var 揮発の関数 = { OFFSET: true, INDIRECT: true };
    function 揮発を含む(木) {
      if (!木 || typeof 木 !== 'object') return false;
      if (木.種 === '呼' && 揮発の関数[String(木.名 || '').toUpperCase()]) return true;
      for (var k in 木) {
        if (!Object.prototype.hasOwnProperty.call(木, k)) continue;
        var v = 木[k];
        if (Array.isArray(v)) { for (var i = 0; i < v.length; i++) if (揮発を含む(v[i])) return true; }
        else if (v && typeof v === 'object' && 揮発を含む(v)) return true;
      }
      return false;
    }

    /* ★★知って いる 板★★（2026-09-15・㍸）
       ★知らない 板を 指したら #REF!★（★空の 板と 無い 板は 別物★）
       ★打った／置いた 時点で 名簿に 入ります★／`板を足す()` でも 入れられます
       （★実物を 押す 時は 先に 板を 全部 足して おく★＝★中身が 空の 板も 在る★） */
    var 板たち = {};
    板たち[既定の板] = true;
    var 今の板 = 既定の板;

    function 板を覚える(n) {
      var ば = 名から番地(n);
      if (ば && ば.板) 板たち[ば.板] = true;
    }

    /* ★板を 書かない 名を 「今の 板」で 揃える★（★鍵は 必ず 板つき★） */
    function 揃える(名) {
      var ば = 名から番地(名);
      if (!ば) return String(名).toUpperCase();
      return 番地から名(ば.行, ば.列, ば.板 || 今の板);
    }
    /* ★四角の 両端の 板を 決める★（Excel は `板2!A1:A3` の 右端も 板2） */
    function 四角の板(a, b) { return (a && a.板) || (b && b.板) || 今の板; }

    function 空に() { return 計.空; }

    /* ★式の 木から「見ている マス」を 全部 拾う★（四角は ほどいて 1つずつ） */
    function 頼りを拾う(木, 出) {
      if (!木 || typeof 木 !== 'object') return out(出);
      if (木.種 === '名') {
        var n = String(木.値).toUpperCase();
        if (n !== 'TRUE' && n !== 'FALSE' && 名から番地(n)) 出.push(揃える(n));
        /* ★`E1#` は E1 を 頼る★（2026-10-05） */
        else if (n.length > 1 && n.charAt(n.length - 1) === '#' && 名から番地(n.slice(0, -1))) 出.push(揃える(n.slice(0, -1)));
        return 出;
      }
      if (木.種 === '二' && 木.記 === ':') {
        /* ★列まるごと・行まるごとは ほどかず ★帯の 見張りの 鍵★ を 出す（`帯|` で 始まる＝マスの 名とは 混ざらない・案X）★ */
        var 帯 = 帯の形(木);
        if (帯) { 帯の鍵たち(帯).forEach(function (k) { 出.push(k); }); return 出; }
        var a = 名から番地(String(木.左.値).toUpperCase());
        var b = 名から番地(String(木.右.値).toUpperCase());
        if (a && b) {
          var 板2 = 四角の板(a, b);      /* ★両端は 同じ 板★ */
          for (var r = Math.min(a.行, b.行); r <= Math.max(a.行, b.行); r++)
            for (var c = Math.min(a.列, b.列); c <= Math.max(a.列, b.列); c++)
              出.push(番地から名(r, c, 板2));
          return 出;
        }
        /* ★★端が 名で ない（`INDEX(A1:A9,2):A3` 等）★★（2026-11-07・taiketsu 別件）
           ★前は 何も 拾わず★＝A2 を 打ち直しても 5 の まま（正しくは 203）＝★黙って 古い 数★
           ⇒★両端の 中で 拾った マスの「板ごとの 外枠」を 全部 頼る★
             （INDEX の 答えは 第1引数の 中＝`:` の 答えは 外枠の 中に 必ず 収まる） */
        var 子の = [];
        頼りを拾う(木.左, 子の);
        頼りを拾う(木.右, 子の);
        外枠で拾う(子の, 出);
        return 出;
      }
      /* ★★SUMIF・AVERAGEIF の 取り直した 足す 範囲も 頼る★★（2026-11-07）
           shiki-kansuu の 条件で選ぶ は 足す 範囲を ★左上から 範囲の 形に 取り直して 読む★（`SUMIF(A1:A5,">2",B1)` は B1:B5 を 読む）
           ⇒ 書いた B1 だけ 頼ると B3 を 打ち直しても 古い 数（実測 120 の まま・正しくは 91）＝取り直した 所を 全部 頼る */
      if (木.種 === '呼') {
        var 呼U = String(木.名 || '').toUpperCase().replace(/^_XLFN.|^_XLWS./, '');
        if ((呼U === 'SUMIF' || 呼U === 'AVERAGEIF') && 木.引数 && 木.引数.length >= 3) 取り直した足す所を拾う(木.引数, 出);
      }
      for (var k in 木) {
        if (!Object.prototype.hasOwnProperty.call(木, k)) continue;
        var v = 木[k];
        if (Array.isArray(v)) { for (var i = 0; i < v.length; i++) 頼りを拾う(v[i], 出); }
        else if (v && typeof v === 'object') 頼りを拾う(v, 出);
      }
      return 出;
      function out(x) { return x; }
    }
    /* ★書いた 四角の 大きさ★（1マス・A1:B2 の 形だけ）＝{ 行数, 列数 }／分からなければ null */
    function 書いた四角の大きさ(木) {
      while (木 && 木.種 === '括') 木 = 木.子;
      if (!木) return null;
      if (木.種 === '名') return 名から番地(String(木.値).toUpperCase()) ? { 行数: 1, 列数: 1 } : null;
      if (木.種 === '二' && 木.記 === ':' && 木.左 && 木.右 && 木.左.種 === '名' && 木.右.種 === '名') {
        var a = 名から番地(String(木.左.値).toUpperCase()), b = 名から番地(String(木.右.値).toUpperCase());
        return (a && b) ? { 行数: Math.abs(a.行 - b.行) + 1, 列数: Math.abs(a.列 - b.列) + 1 } : null;
      }
      return null;
    }
    /* ★取り直した 足す 所を 頼りを拾う が 知る 形か★（帯なら 足す帯 が 取れる・帯で なければ 書いた 大きさと 左上が 分かる） */
    function 取り直しを頼れるか(引) {
      if (帯の形(引[0])) return !!SUMIFの足す帯(引);
      return !!(書いた四角の大きさ(引[0]) && 左上の番地(引[2]));
    }
    function 引数の大きさ(a) {
      if (!a) return null;
      if (a.種 === '四角') return { 行数: a.行数, 列数: a.列数 };
      if (a.種 === 'マス') return { 行数: 1, 列数: 1 };
      return null;
    }
    function 取り直した足す所を拾う(引, 出) {
      if (帯の形(引[0])) {
        var 足す帯 = SUMIFの足す帯(引);
        if (足す帯) 帯の鍵たち(足す帯).forEach(function (k) { 出.push(k); });
        return;
      }
      var 大 = 書いた四角の大きさ(引[0]), 角 = 左上の番地(引[2]);
      /* ★大きさか 左上が 分からない 形（INDEX(A1:A9,2):A3・名前の 付いた 範囲 等）は 前の まま★＝書いた 物だけ 頼る（未測定） */
      if (!大 || !角) return;
      for (var r = 角.行; r < 角.行 + 大.行数 && r <= 板の最後の行; r++)
        for (var c = 角.列; c < 角.列 + 大.列数 && c <= 板の最後の列; c++) 出.push(番地から名(r, c, 角.板U));
    }
    /* ★拾った 名を 板ごとの 外枠に 広げて 足す★（帯の 鍵は そのまま） */
    function 外枠で拾う(拾った, 出) {
      var 枠 = {}, 板順 = [], 帯 = {};
      for (var i = 0; i < 拾った.length; i++) {
        var k = 拾った[i];
        if (k.indexOf('帯|') === 0) {
          出.push(k);
          /* ★帯の 向きと 番号を 覚える★（taiketsu 止め1）＝`INDEX(A:A,5):C3` は 列A〜C の 帯 */
          var 片 = k.split('|');
          if (片.length === 4) {
            var bu = 帯[片[1]] || (帯[片[1]] = {});
            var no = Number(片[3]), 向 = 片[2];
            if (!bu[向]) bu[向] = { 小: no, 大: no };
            else { if (no < bu[向].小) bu[向].小 = no; if (no > bu[向].大) bu[向].大 = no; }
          }
          continue;
        }
        var ば = 名から番地(k);
        if (!ば) continue;
        var 板 = ば.板 || 今の板;
        var w = 枠[板];
        if (!w) { 枠[板] = { 行0: ば.行, 行1: ば.行, 列0: ば.列, 列1: ば.列 }; 板順.push(板); continue; }
        if (ば.行 < w.行0) w.行0 = ば.行; if (ば.行 > w.行1) w.行1 = ば.行;
        if (ば.列 < w.列0) w.列0 = ば.列; if (ば.列 > w.列1) w.列1 = ば.列;
      }
      for (var j = 0; j < 板順.length; j++) {
        var f = 枠[板順[j]];
        for (var r = f.行0; r <= f.行1; r++)
          for (var c = f.列0; c <= f.列1; c++)
            出.push(番地から名(r, c, 板順[j]));
      }
      /* ★帯が 在る 板は 外枠の 列（行）まで 帯の 鍵を 広げる★ */
      for (var bU in 帯) {
        if (!Object.prototype.hasOwnProperty.call(帯, bU)) continue;
        /* ★列の 帯と 行の 帯が 両方★＝答えは 板の どこにでも 広がり得る⇒★板まるごと★（`INDEX(A:A,5):INDEX(1:1,3)`） */
        if (帯[bU].C && 帯[bU].R) { 出.push('帯|' + bU + '|*'); continue; }
        var 枠U = null;
        for (var q = 0; q < 板順.length; q++) if (String(板順[q]).toUpperCase() === bU) 枠U = 枠[板順[q]];
        [['C', '列0', '列1'], ['R', '行0', '行1']].forEach(function (t) {
          var g = 帯[bU][t[0]];
          if (!g) return;
          var 小 = 枠U ? Math.min(g.小, 枠U[t[1]]) : g.小, 大 = 枠U ? Math.max(g.大, 枠U[t[2]]) : g.大;
          if (大 - 小 + 1 > 帯の鍵の上限) { 出.push('帯|' + bU + '|*'); return; }
          for (var x = 小; x <= 大; x++) 出.push('帯|' + bU + '|' + t[0] + '|' + x);
        });
      }
      return 出;
    }

    /* ══ ★★列まるごと・行まるごと（帯）★★ ══（2026-11-04・案X）
         ★読み口は shiki-sansho の 参.読む★（名から番地 は 変えない＝設計前 taiketsu S5）
         ★出す 形★ { 板U, 上, 下, 左, 右（本当の 大きさ・0から）, 列全体, 行全体 } ／ 帯で なければ null */
    function 帯の形(木) {
      while (木 && 木.種 === '括') 木 = 木.子;
      if (!木 || 木.種 !== '二' || 木.記 !== ':' || !木.左 || !木.右) return null;
      if ((木.左.種 !== '名' && 木.左.種 !== '数') || (木.右.種 !== '名' && 木.右.種 !== '数')) return null;
      var 右字 = String(木.右.値);
      if (右字.indexOf('!') >= 0) return null;
      var r = 参 && 参.読む ? 参.読む(String(木.左.値) + ':' + 右字) : null;
      if (!r || !r.ok || !(r.列全体 || r.行全体)) return null;
      return { 板U: (r.板 !== null && r.板 !== undefined) ? String(r.板).toUpperCase() : 今の板,
        上: r.上, 下: r.下, 左: r.左, 右: r.右, 列全体: !!r.列全体, 行全体: !!r.行全体 };
    }
    /* ★帯の 見張りの 鍵★＝列ごと `帯|板|C|列`・行ごと `帯|板|R|行`／256 を 越える 帯は 板まるごと `帯|板|*` */
    var 帯の鍵の上限 = 256;
    function 帯の鍵たち(帯) {
      var 出 = [], i;
      if (帯.列全体) {
        if (帯.右 - 帯.左 + 1 > 帯の鍵の上限) return ['帯|' + 帯.板U + '|*'];
        for (i = 帯.左; i <= 帯.右; i++) 出.push('帯|' + 帯.板U + '|C|' + i);
      } else {
        if (帯.下 - 帯.上 + 1 > 帯の鍵の上限) return ['帯|' + 帯.板U + '|*'];
        for (i = 帯.上; i <= 帯.下; i++) 出.push('帯|' + 帯.板U + '|R|' + i);
      }
      return 出;
    }
    /* ★この マスを 見る 帯の 鍵★（見る者たち が 使う） */
    /* ★名ごとに 1回だけ 作って 控える★（名→番地 は 変わらない）＝返す 並びは 読むだけ（書き換えるな） */
    function マスの帯の鍵たち(n) {
      var 控 = 帯鍵の控え[n];
      if (控) return 控;
      var ば = 名から番地(n);
      if (!ば) return (帯鍵の控え[n] = []);
      var 板U = ば.板 ? String(ば.板).toUpperCase() : '';
      return (帯鍵の控え[n] = ['帯|' + 板U + '|C|' + ば.列, '帯|' + 板U + '|R|' + ば.行, '帯|' + 板U + '|*']);
    }
    /* ★帯を 切って 読む★＝板の 端（＋空の 1行・1列）まで。
         ★空の 1行を 足す 訳★＝「最初の 空の マス」に 当たる 探し方（XLOOKUP 等）と UNIQUE の 様な 「空も 1つ」を 切っても 同じに する。
         `伸ばす行`・`伸ばす列`（0から）＝INDEX(A:A,k) で k が 端を 越える 時に そこまで 読む。
         ★出す 形★ 四角（引数にする と 同じ 顔）＋ `帯: { 本当の行数, 本当の列数 }` */
    function 帯を読む(帯, 通り道, 伸ばす行, 伸ばす列) {
      var r0 = 帯.上, r1 = 帯.下, c0 = 帯.左, c1 = 帯.右;
      if (帯.列全体) r1 = Math.min(帯.下, Math.max(帯の端(帯) + 1, 伸ばす行 || 0));
      if (帯.行全体) c1 = Math.min(帯.右, Math.max(帯の端(帯) + 1, 伸ばす列 || 0));
      var 左名 = 番地から名(r0, c0, 帯.板U), 右名 = 番地から名(r1, c1, 帯.板U);
      var 四 = 四角を読む(左名, 右名, 通り道);
      return {
        種: '四角', 並び: 四.並び, 行数: 四.行数, 列数: 四.列数,
        番地: { 行: r0, 列: c0, 板: 帯.板U },
        打った字: function (i) {
          var m = 中身[番地から名(r0 + Math.floor(i / 四.列数), c0 + (i % 四.列数), 帯.板U)];
          return m ? String(m.打った字 || '') : '';
        },
        帯: { 本当の行数: 帯.下 - 帯.上 + 1, 本当の列数: 帯.右 - 帯.左 + 1 },
      };
    }

    /* ══ ★★帯を 置いて よい 所★★ ══（2026-11-04・案X）
         ★(a) 切っても 答えが 変わらない 所★（尻尾は 全部 空のマス）＝SUM COUNT COUNTA AVERAGE MAX MIN SUMPRODUCT・SUMIF 等の 範囲・
           完全一致の VLOOKUP/HLOOKUP・MATCH（0 と 1）・既定の XLOOKUP・AREAS・ISREF
         ★(b) 大きさで 変わる 所★＝ROWS COLUMNS COUNTBLANK COUNTIF COUNTIFS INDEX ＝★本当の 大きさで 足し直す★（帯の余りを足す・INDEX は 伸ばして 読む）
         ★それ 以外★（演算子の 中・裸の =A:A・ROW・UNIQUE・SORT・FILTER・OFFSET・INDIRECT・LET・LAMBDA・名簿に 無い 関数）＝★前と 同じ #REF!★
         ★物差し★ tests/fixtures/retsu-marugoto-excel.json（実Excel 16.0） */
    function 字か数の値(木) {
      while (木 && 木.種 === '括') 木 = 木.子;
      if (!木) return null;
      if (木.種 === '数') return { 数: Number(木.値) };
      if (木.種 === '前' && 木.記 === '-' && 木.子 && 木.子.種 === '数') return { 数: -Number(木.子.値) };
      if (木.種 === '名' && /^FALSE$/i.test(String(木.値))) return { 数: 0 };
      if (木.種 === '名' && /^TRUE$/i.test(String(木.値))) return { 数: 1 };
      return null;
    }
    function 数が(木, 許す) { var x = 字か数の値(木); return !!(x && x.数 !== undefined && 許す.indexOf(x.数) >= 0); }
    function 一本の帯か(木) { var b = 帯の形(木); return !!(b && (b.列全体 ? b.左 === b.右 : b.上 === b.下)); }
    /* ★出す 形★ { よい: function(i) 真偽, 伸ばす行, 伸ばす列 } */
    function 帯の置き場(名, 引, 手, 通り道) {
      var n = 引.length, 出 = { 伸ばす行: 0, 伸ばす列: 0, よい: function () { return false; } };
      switch (名) {
        case 'SUM': case 'COUNT': case 'COUNTA': case 'AVERAGE': case 'MAX': case 'MIN': case 'SUMPRODUCT': case 'PRODUCT':
          出.よい = function () { return true; }; break;
        case 'SUMIF': 出.よい = function (i) { return i === 0 || i === 2; }; break;
        case 'COUNTIFS': 出.よい = function (i) { return i % 2 === 0; }; break;
        case 'SUMIFS': 出.よい = function (i) { return i === 0 || i % 2 === 1; }; break;
        case 'COUNTIF': case 'COUNTBLANK': case 'ROWS': case 'COLUMNS': case 'ISREF': case 'AREAS':
          出.よい = function (i) { return i === 0; }; break;
        case 'VLOOKUP': case 'HLOOKUP': {
          var 帯V = 帯の形(引[1]), 縦 = 名 === 'VLOOKUP';
          var よV = n === 4 && 数が(引[3], [0]) && !!帯V && (縦 ? 帯V.列全体 : 帯V.行全体);
          出.よい = function (i) { return よV && i === 1; }; break;
        }
        case 'MATCH': {
          var よM = 一本の帯か(引[1]) && (n === 2 || 数が(引[2], [0, 1]));
          出.よい = function (i) { return よM && i === 1; }; break;
        }
        case 'XLOOKUP': {
          var よX = 一本の帯か(引[1]) && (n < 5 || 数が(引[4], [0])) && (n < 6 || 数が(引[5], [1]));
          出.よい = function (i) { return よX && (i === 1 || i === 2); }; break;
        }
        case 'INDEX': {
          /* ★行・列の 番号が 1 以上の 数の 時だけ★（0＝まるごとの 参照は 大きさで 変わる＝置かない）
             ★端を 越える 番号は そこまで 伸ばして 読む★（実Excel `=INDEX(A:A,10)` → 0） */
          var 帯I = 帯の形(引[0]);
          if (!帯I || n < 2 || n > 3) break;
          var 番 = function (k) {
            var v = 値に(歩く(引[k], 手, 通り道));
            return (v && v.型 === '数' && isFinite(v.値)) ? Math.floor(v.値) : null;
          };
          var k1 = 番(1), k2 = n === 3 ? 番(2) : null;
          if (k1 === null || k1 < 1 || (n === 3 && (k2 === null || k2 < 1))) break;
          if (帯I.列全体) { 出.伸ばす行 = k1; if (n === 2 && 帯I.左 !== 帯I.右) break; }
          else { if (n === 2) { if (帯I.上 !== 帯I.下) break; 出.伸ばす列 = k1; } else 出.伸ばす列 = k2; }
          出.よい = function (i) { return i === 0; }; break;
        }
      }
      return 出;
    }
    /* ★SUMIF の 足す 範囲の 左上★（帯・1マス・A1:B2 の 形だけ）＝{ 行, 列, 板U }／取れなければ null */
    function 左上の番地(木) {
      while (木 && 木.種 === '括') 木 = 木.子;
      if (!木) return null;
      var 帯 = 帯の形(木);
      if (帯) return { 行: 帯.上, 列: 帯.左, 板U: 帯.板U };
      var 板の字 = function (ば) { return ば.板 ? String(ば.板).toUpperCase() : 今の板; };
      if (木.種 === '名') {
        var ば = 名から番地(String(木.値).toUpperCase());
        return ば ? { 行: ば.行, 列: ば.列, 板U: 板の字(ば) } : null;
      }
      if (木.種 === '二' && 木.記 === ':' && 木.左 && 木.右 && 木.左.種 === '名' && 木.右.種 === '名') {
        var a = 名から番地(String(木.左.値).toUpperCase());
        var 右字 = String(木.右.値).toUpperCase();
        var b = 名から番地(右字.indexOf('!') >= 0 ? 右字 : (a && a.板 ? a.板 + '!' + 右字 : 右字));
        if (!a || !b) return null;
        return { 行: Math.min(a.行, b.行), 列: Math.min(a.列, b.列), 板U: 板の字(a) };
      }
      return null;
    }
    /* ★SUMIF(帯, 条件, 足す) の 本当に 足す 所★＝足す 範囲の 左上から 範囲（帯）と 同じ 形の 帯／取れなければ null
         ★左上が 帯の 始まり（列の 帯なら 1行目・行の 帯なら A列）で ない 時は null★（`SUMIF(A:A,">2",B2)` は 板の 外へ はみ出す＝未測定） */
    var 板の最後の行 = 1048575, 板の最後の列 = 16383;
    function SUMIFの足す帯(引) {
      if (!引 || 引.length < 3) return null;
      var 範 = 帯の形(引[0]);
      if (!範) return null;
      var 角 = 左上の番地(引[2]);
      if (!角) return null;
      if (範.列全体) {
        if (角.行 !== 範.上) return null;
        var 右 = 角.列 + (範.右 - 範.左);
        if (右 > 板の最後の列) return null;
        return { 板U: 角.板U, 上: 範.上, 下: 範.下, 左: 角.列, 右: 右, 列全体: true, 行全体: false };
      }
      if (角.列 !== 範.左) return null;
      var 下 = 角.行 + (範.下 - 範.上);
      if (下 > 板の最後の行) return null;
      return { 板U: 角.板U, 上: 角.行, 下: 下, 左: 範.左, 右: 範.右, 列全体: false, 行全体: true };
    }
    function 帯の形を合わせる(名, 引数たち, 足す帯) {
      /* ★形を 突き合わせる 関数★（実Excel `=SUMPRODUCT(A:A,B1:B5)`・`=XLOOKUP(4,A:A,B1:B5)` → #VALUE!）
           本当の 形（帯は 本当の 大きさ・ふつうの 四角は その 大きさ）が 違えば #VALUE!／
           同じでも 帯で ない 範囲が 混ざれば 前と 同じ #REF!（切った 帯と 1,048,576 行の 四角は 台の 中で 形が 合わない） */
      function 形(a) {
        if (!a) return null;
        if (a.帯) return { 行: a.帯.本当の行数, 列: a.帯.本当の列数, 帯: true };
        if (a.種 === '四角') return { 行: a.行数, 列: a.列数, 帯: false };
        return { 行: 1, 列: 1, 帯: false };
      }
      var i, 位置 = [];
      /* ★★SUMIF の 足す 範囲は ★左上から 範囲の 形に 取り直す★★（2026-11-07・shiki-kansuu の 条件で選ぶ が 四角を取る で 取り直す）
           実Excel（retsu-x-mae4-excel.json）`SUMIF(A:A,">2",B:C)` → 120・`SUMIF(A:B,">2",C:C)` → 1212
           ⇒ 範囲が 帯の 時は ★取り直した 足す 範囲も 帯として 端まで 切る★（足す帯＝SUMIFの足す帯）＝尻尾の 先に 数が 在っても 落とさない
           ★足す帯 が 取れない（左上が 1行目／A列で ない・番地の 無い 物）時だけ 前と 同じ #REF!★（未測定） */
      if (名 === 'SUMIF') {
        if (引数たち.length > 2 && 引数たち[0] && 引数たち[0].帯 && !足す帯) return 計.誤('#REF!');
        return null;
      }
      if (名 === 'XLOOKUP') {
        var 探 = 形(引数たち[1]), 返 = 形(引数たち[2]);
        if (!探 || !返) return null;
        var 縦 = 探.列 === 1;
        if (縦 ? 返.行 !== 探.行 : 返.列 !== 探.列) return 計.誤('#VALUE!');
        if (!探.帯 || !返.帯) return 計.誤('#REF!');
        /* ★返す 範囲が 2列（2行）以上は 溢れる＝台は 1つしか 返さない（`XLOOKUP(3,A:A,B:C)` 台 20・テスト線前 taiketsu 止め2）⇒ #REF! */
        if (返.行 !== 探.行 || 返.列 !== 探.列) return 計.誤('#REF!');
        return null;
      }
      if (名 === 'SUMPRODUCT') { for (i = 0; i < 引数たち.length; i++) 位置.push(i); }
      else if (名 === 'COUNTIFS') { for (i = 0; i < 引数たち.length; i += 2) 位置.push(i); }
      else if (名 === 'SUMIFS') { 位置.push(0); for (i = 1; i < 引数たち.length; i += 2) 位置.push(i); }
      else return null;
      var 見本 = null, 混ざる = false;
      for (i = 0; i < 位置.length; i++) {
        var f = 形(引数たち[位置[i]]);
        if (!f) continue;
        if (!f.帯) 混ざる = true;
        if (見本 === null) 見本 = f;
        else if (f.行 !== 見本.行 || f.列 !== 見本.列) return 計.誤('#VALUE!');
      }
      return 混ざる ? 計.誤('#REF!') : null;
    }
    /* ★2026-11-07 直した★＝空の マスの ずれは shiki-kansuu の 当てはまるか で 直した（実Excel と 合う）
         ⇒ 帯でも ★1つの 値の 条件★は 受ける（COUNTIF・COUNTIFS は 余りを 帯の余りを足す で 足す・SUMIF・SUMIFS の 余りは 空＝0）
         ★条件が 四角・表（1つの 値で ない）の 時だけ 前と 同じ #REF!★（帯での 形は 未測定） */
    /* ★★条件の 字に * ? ~ が 在る 時は 帯では 前と 同じ #REF!★★（2026-11-07・本番前 taiketsu 止め3）
         当てはまるか は * ? を 見ない（前から）＝帯の 余りを 足すと `COUNTIF(A:A,"<>*")` が 1048576（Excel 1048575）の 黙った 数
         ⇒ 写し字（ワイルドカード）を 作るまで 帯では 印に する */
    function 一つの条件か(a) {
      if (!(a && (a.種 === '直' || a.種 === 'マス'))) return false;
      var v = a.値;
      if (v && v.型 === '字' && /[*?~]/.test(String(v.値))) return false;
      return true;
    }
    function 帯の条件を見る(名, 引数たち) {
      var i;
      if (名 === 'COUNTIF' || 名 === 'SUMIF') return 一つの条件か(引数たち[1]) ? null : 計.誤('#REF!');
      if (名 === 'COUNTIFS') {
        for (i = 1; i < 引数たち.length; i += 2) if (!一つの条件か(引数たち[i])) return 計.誤('#REF!');
      }
      if (名 === 'SUMIFS') {
        for (i = 2; i < 引数たち.length; i += 2) if (!一つの条件か(引数たち[i])) return 計.誤('#REF!');
      }
      /* ★MATCH の 近い 方（型1・省略）は 二分探索★＝間に 空の マスが 在ると 答えが 長さで 変わる（設計前 taiketsu 止め3・
           実Excel が 空を 挟んだ 時の 答えは まだ 測って いない）⇒ 帯の 中に 空を 挟めば 前と 同じ #REF! */
      if (名 === 'MATCH' && 引数たち[1] && 引数たち[1].帯) {
        var 型 = 引数たち.length < 3 ? 1 : (引数たち[2] && 引数たち[2].値 && 引数たち[2].値.型 === '数' ? 引数たち[2].値.値 : null);
        if (型 !== 0) {
          var 並 = 引数たち[1].並び, 最後 = -1;
          for (i = 0; i < 並.length; i++) if (!(並[i] && 並[i].型 === '空')) 最後 = i;
          for (i = 0; i < 最後; i++) if (並[i] && 並[i].型 === '空') return 計.誤('#REF!');
        }
      }
      return null;
    }
    /* ★(b) 大きさで 変わる 物を 本当の 大きさで 足し直す★＝余り（本当の マス − 切った マス）は ★全部 空のマス★
         ★COUNTIF・COUNTIFS★＝条件が ★全部 空の マスに 当たる★ 時だけ 余りを 足す（空に当たる＝1マスの 空で COUNTIF を 呼んで 見る＝同じ 当てはまるか）
           実Excel（retsu-marugoto-excel.json・A1:A5=1..5）`COUNTIF(A:A,"")` → 1048571・`COUNTIF(A:A,"<>3")` → 1048575
         ★COUNTIFS の 範囲は 全部 同じ 本当の 形・全部 帯（帯の形を合わせる）＝余りの 行は どの 範囲でも 空★ */
    function 余りのマス(a) { return a && a.帯 ? a.帯.本当の行数 * a.帯.本当の列数 - a.行数 * a.列数 : 0; }
    function 帯の余りを足す(名, 引数たち, 出, 空に当たる) {
      if (!出 || 出.型 !== '数') return 出;
      if (名 === 'COUNTIF' && 引数たち.length === 2 && 引数たち[0] && 引数たち[0].帯) {
        return 空に当たる(引数たち[1]) ? 計.数(出.値 + 余りのマス(引数たち[0])) : 出;
      }
      if (名 === 'COUNTIFS' && 引数たち.length >= 2 && 引数たち.length % 2 === 0 && 引数たち[0] && 引数たち[0].帯) {
        for (var k = 1; k < 引数たち.length; k += 2) if (!空に当たる(引数たち[k])) return 出;
        return 計.数(出.値 + 余りのマス(引数たち[0]));
      }
      if (名 === 'COUNTBLANK') {
        var 足 = 0;
        for (var i = 0; i < 引数たち.length; i++) {
          var a = 引数たち[i];
          if (a && a.帯) 足 += a.帯.本当の行数 * a.帯.本当の列数 - a.行数 * a.列数;
        }
        return 計.数(出.値 + 足);
      }
      return 出;
    }
    /* ══ ★参照（INDEX が 返す 物）の 受け口★ ══（2026-09-15）
       ★参照は「値」では ありません★＝★どのマスか★を 持って います。
       ★3つの 出口★しか 在りません。★ここ 以外に 漏らさない★のが 決めです。
         ①★1つの 値が 要る 所★（＋ − 大小・マスに 入れる）… ★暗黙の 交わり★
         ②★関数の 引数★                                  … ★四角と 同じ 顔★に する
         ③★`:` の 端★                                    … ★番地として 使う★
       ★実測★ `=SUM(INDEX(A1:B5,0,2))` → 30（②）／`=SUM(INDEX(A1:A5,2):A5)` → 14（③）
              `=INDEX(A1:B5,0,2)` を J1 に 打つと 2・J9 に 打つと #VALUE!（①） */
    function 値に(v) { return 場.参照か(v) ? 場.交わる(v, 今のマス) : v; }

    /** ★★LET★★ ... 名前に 値を 付けて 後ろで 使う（2026-09-18）
          ★形★ `=LET(名1, 値1, [名2, 値2, ...], 式)`
            ・★引数は 必ず 奇数★（名・値 の 対 ＋ 最後の 式）
            ・★最後は 式★（★名では ない★）
          ★決め（★実Excel に まだ 聞いて いません★＝9枠目 #37〜#40）★
            ・★前の 名前を 後ろの 値で 使えます★（`=LET(a,1,b,a+1,b)` → 2 の はず）
            ・★名が 名で ない（数・字）なら #VALUE!★
            ・★引数が 偶数なら #VALUE!★
          ★★溢れは 詰めません★★＝最後の 式が 表を 返せば そのまま 返します */
    function LETを計算する(木, 手, 通り道) {
      var 引 = 木.引数 || [];
      if (引.length < 3) return 計.誤('#VALUE!');
      if (引.length % 2 === 0) return 計.誤('#VALUE!');
      var 表 = {};
      var 前 = 名の台;
      名の台 = { 表: 表, 親: 前 };
      try {
        for (var i = 0; i + 2 < 引.length; i += 2) {
          var な = 引[i];
          if (!な || な.種 !== '名') return 計.誤('#VALUE!');
          var n = String(な.値).toUpperCase();
          /* ★TRUE／FALSE は 名に 出来ません★（★予約された 名前★・未測定） */
          if (n === 'TRUE' || n === 'FALSE') return 計.誤('#VALUE!');
          var v = 歩く(引[i + 1], 手, 通り道);
          if (v && v.型 === '誤') return v;        /* ★値の 誤りは そのまま 返す★ */
          表[n] = 値に(v);                          /* ★参照は 1つの 値に 詰める★ */
        }
        return 歩く(引[引.length - 1], 手, 通り道);
      } finally {
        名の台 = 前;                                /* ★必ず 戻す★ */
      }
    }

    /* ★木を 歩いて 計算する★（★輪は 呼ぶ側が 見つける★） */
    function 歩く(木, 手, 通り道) {
      switch (木.種) {
        case '数': return 計.数(計.打った数(木.値));
        case '字': return 計.字(String(木.値).slice(1, -1).split('""').join('"'));
        case '誤': return 計.誤(木.値);
        /* ★省いた 引数★（`=IF(TRUE,,2)`）＝★空の 値★
           実測 `=IF(TRUE,,2)` → ★0★（空は 0 として 見える） */
        case '空': return 計.空;
        /* ★括弧を 通ると ★足し引きの 印★ が 落ちます★（2026-09-16 実測）
             `=(0.1+0.2-0.3)+0` … 実Excel ★5.55e-17★（0 では 無い） */
        case '括': return 計.印を落とす(歩く(木.子, 手, 通り道));
        /* ★表を 直に 書いた 物★ `{3;5;7}` `{1,2;3,4}`（2026-09-15 に 入れた）
           ＝形に する 台は 前から `配` に して いましたが ★本体が 知らず #VALUE!★でした。
             実測 … `=TREND({3;5;7;9;11;13})` → ★2.9999999999999987★
                    `=PERCENTOF(2,{1;2;3;4})` → ★0.2★
           ⇒ ★溢れ（行×列）に する★＝四角と 同じ 扱いに なる */
        case '配': {
          var 並3 = 木.行.map(function (段) {
            return 段.map(function (子) { return 歩く(子, 手, 通り道); });
          });
          return 溢.溢れ(並3);
        }
        /* ★溢れ（土台⑤）★＝四角や 溢れに 演算子を 掛けると ★1つずつ★ 効く
           実測 `=-A1:A3` → -1／-2／-3 ／ `=A1:A3&"x"` → 1x／2x／3x */
        case '前': return 溢.ひとつずつ(値に(歩く(木.子, 手, 通り道)), function (x) { return 計.前置き(木.記, x, 手); });
        case '後': return 溢.ひとつずつ(値に(歩く(木.子, 手, 通り道)), function (x) { return 計.後置き(木.記, x, 手); });
        case '二':
          if (木.記 === ' ') {
            var 角Q = 参照の角(木, 手, 通り道);
            if (角Q && 角Q.型 === '誤') return 角Q;
            if (!角Q) return 計.誤('#VALUE!');
            return 歩く(角から木(角Q), 手, 通り道);
          }
          /* ★`A1:A3` を そのまま 打つと 四角が 値に なる★（実測＝下へ 広がる）
             ★`@` を 付けると 1マスに なる★（暗黙の交差）＝`=@A1:A3` → 1
               （字に 切る 台は `@A1` を 1つの 名前に します＝ここで 外す） */
          if (木.記 === ':') {
            /* ★端は「名前」とは 限りません★（2026-09-15）
               ★実測★ `=SUM(INDEX(A1:A5,2):A5)` → ★14★（A2:A5）
               ⇒★INDEX が 返した 参照も `:` の 端に なれます★＝AREAS／OFFSET も ここに 乗ります */
            var 対 = 端の対(木.左, 木.右, 手, 通り道);
            if (対.誤) return 計.誤(対.誤);
            var 左字 = 対.左, 右字 = 対.右;
            var 交差 = 左字.charAt(0) === '@';
            var 四 = 四角を読む(交差 ? 左字.slice(1) : 左字, 右字, 通り道);
            if (!四.行数 || !四.列数) return 計.誤('#REF!');
            var 並 = [];
            for (var r2 = 0; r2 < 四.行数; r2++) {
              var 段2 = [];
              for (var c2 = 0; c2 < 四.列数; c2++) 段2.push(四.並び[r2 * 四.列数 + c2]);
              並.push(段2);
            }
            var 出2 = 溢.溢れ(並);
            return 交差 ? 溢.先頭(出2) : 出2;
          }
          /* ★形が 違う 四角どうしは 足りない 所だけ #N/A★（実測 `=A1:A3+B1:B2` → 11／22／#N/A） */
          return 溢.重ねる(値に(歩く(木.左, 手, 通り道)), 値に(歩く(木.右, 手, 通り道)),
            function (a, b) { return 計.つなぎ(木.記, a, b, 手); });
        case '名': {
          var n = String(木.値).toUpperCase();
          /* ★溢れの 参照 `E1#`★（2026-10-05） */
          if (角で読む木か(木)) {
            var 角N = 参照の角(木, 手, 通り道);
            if (角N && 角N.型 === '誤') return 角N;
            if (!角N) return 計.誤('#REF!');
            return 歩く(角から木(角N), 手, 通り道);
          }
          if (n === 'TRUE') return 計.真偽(true);
          if (n === 'FALSE') return 計.真偽(false);
          /* ★★名の台が ★マスの 名前より 先★★★（2026-09-18）
               ＝`=LET(A1,5,A1)` は ★5★ の はず（★実Excel に 聞きます＝9枠目 #38★）
               ＝★TRUE／FALSE より 後★＝★予約された 名前は 上書きしない★
               ★ここが 「まだ 聞いて いない 決め」です★ */
          var 束 = 名を引く(n);
          if (束 !== undefined) return 束;   /* ★省略の 印も そのまま 返します★（ISOMITTED が 見る） */
          /* ★1マスに `@` が 付いた 形★（`=@A1`）＝1マスなので そのまま */
          if (n.charAt(0) === '@' && 名から番地(n.slice(1))) return 読む(揃える(n.slice(1)), 通り道);
          if (名から番地(n)) return 読む(揃える(n), 通り道);
          return 計.誤('#NAME?');
        }
        /* ══ ★★`式(引数)` の 形★★ ══（2026-09-18・★繋ぐ 前の 関門の 最後★）
             `=LAMBDA(x,x*2)(3)` ... ★実Excel は 6★（9枠目 36行目で 実測）
             ★読む 側★ `lib/shiki-katachi.js` が `{種:'呼び出し', 関, 引数}` を 作ります
             ★ここでは★ ... ★関を 計算して ★ラムダなら 呼ぶ★★
             ★ラムダで なければ #VALUE!★（★測って いない 形を 増やさない★） */
        case '呼び出し': {
          var 関の値 = 歩く(木.関, 手, 通り道);
          if (関の値 && 関の値.型 === '誤') return 関の値;
          if (!ラムダか(関の値)) return 計.誤('#VALUE!');
          /* ★★「省かれた」と「空」を 分けます★★（2026-09-18・実測）
               `=LAMBDA(x,y,ISOMITTED(y))(1,)` ... 実Excel ★True★（★台は FALSE でした★）
               `=LAMBDA(x,y,IF(ISOMITTED(y),..,y))(1,)` ... 実Excel ★ナシ★（台は 0）
               ⇒★読む 側は `{種:'空'}` を 作って います★（`=IF(TRUE,,2)` と 同じ 形）
               ⇒★でも ★ラムダに 渡す 時だけ★ 「省かれた」と して 覚えます★
               ⇒★それ 以外（IF の 枝など）は 今まで通り 0★＝★触りません★ */
          var 引値Y = (木.引数 || []).map(function (子) {
            if (子 && 子.種 === '空') return { 型: '省略' };
            return 値に(歩く(子, 手, 通り道));
          });
          return ラムダを呼ぶ(関の値, 引値Y, 手, 通り道);
        }
        case '呼': {
          /* ══ ★★先に 計算しない 関数★★ ══（2026-09-18・ラムダの 一族）
               ★なぜ 要るか★ ... この 下で `木.引数.map(引数にする)` を すると
                 ★引数の 中身が 全部 その場で 計算されます★。
                 `=LET(x,1,x+1)` の `x+1` は ★x が 決まる 前★に 計算されて #NAME? に なります。
               ⇒★だから LET だけ ★木の まま★ 受け取ります★
               ★ここに 置く 訳★ ... `lib/shiki-kansuu.js` は ★木を 知りません★
                 ＝値しか 受け取れないので ★板の 側に しか 置けません★ */
          var 呼び名 = String(木.名).toUpperCase();
          if (呼び名 === 'LET') return LETを計算する(木, 手, 通り道);
          /* ★★LAMBDA も 先に 計算しません★★（中身は ★呼ばれた 時★に 計算します） */
          if (呼び名 === 'LAMBDA') {
            var 引L = 木.引数 || [];
            if (引L.length < 2) return 計.誤('#VALUE!');
            var 名たち = [];
            for (var iL = 0; iL < 引L.length - 1; iL++) {
              if (!引L[iL] || 引L[iL].種 !== '名') return 計.誤('#VALUE!');
              var nL = String(引L[iL].値).toUpperCase();
              if (nL === 'TRUE' || nL === 'FALSE') return 計.誤('#VALUE!');
              名たち.push(nL);
            }
            return { 型: 'ラムダ', 名たち: 名たち, 体: 引L[引L.length - 1], 台: 名の台 };
          }
          /* ★★名に 入れた ラムダを 呼ぶ★★（`=LET(f,LAMBDA(x,x*2),f(3))`）
               ＝★読む 段では `f(3)` も ただの「呼」の 木です★
               ＝★だから 名の台を 先に 見ます★／★実Excel は 6 を 返します★（9枠目 57行目） */
          var 束L = 名を引く(呼び名);
          if (ラムダか(束L)) {
            var 引値L = (木.引数 || []).map(function (子) { return 値に(歩く(子, 手, 通り道)); });
            return ラムダを呼ぶ(束L, 引値L, 手, 通り道);
          }
          /* ★名前を 付けます★＝★束ねの 中で もう一度 呼べる 為★（2026-09-16） */
          var 引数にする = function (子, 帯よし) {
            /* ★列まるごと・行まるごと★（2026-11-04・案X）＝★置いて よい 所★だけ 切って 読む／それ以外は 前と 同じ #REF! */
            var 帯 = 帯の形(子);
            if (帯) {
              if (!帯よし) return { 種: '直', 値: 計.誤('#REF!') };
              帯たち.push(帯);
              return 帯を読む(帯, 通り道, 帯の伸ばし.行, 帯の伸ばし.列);
            }
            /* ★交わり・溢れの 参照は 四角に 直して `:` の 道へ★（2026-10-05） */
            if (角で読む木か(子)) {
              var 角A = 参照の角(子, 手, 通り道);
              if (!角A || 角A.型 === '誤') return { 種: '直', 値: 角A || 計.誤('#VALUE!') };
              return 引数にする(角から木(角A));
            }
            /* ★★束ね（括弧の 中の カンマ）★★
                 `=AREAS((B2:D4,E5))` の ような ★とびとびの 参照★
                 ★中身を 見ず ★一つずつ 引数に して 並べるだけ★
                   （AREAS は ★数だけ★ 見ます） */
            if (子.種 === '束') {
              return { 種: '束', 並: 子.並.map(function (x) { return 引数にする(x, 帯よし); }) };
            }
            if (子.種 === '二' && 子.記 === ':') {
              /* ★端は 名前とは 限らない★（`SUM(INDEX(…):A5)`）＝`端の名` に 任せる */
              var 対2 = 端の対(子.左, 子.右, 手, 通り道);
              if (対2.誤) return { 種: '直', 値: 計.誤(対2.誤) };
              var 左名 = 対2.左, 右名 = 対2.右;
              if (左名.charAt(0) === '@') 左名 = 左名.slice(1);
              var 四 = 四角を読む(左名, 右名, 通り道);
              /* ★四角の 左上の 番地も 渡す★（2026-09-15）
                 ＝`=CELL("row",A1:C3)` は ★左上（A1）の 事★（実測 … 1）。
                   前は 値だけ 渡して いたので ★#VALUE!★ に なって いた。 */
              var ば左 = 名から番地(左名), ば右 = 名から番地(右名);
              var 角 = (ば左 && ば右)
                ? { 行: Math.min(ば左.行, ば右.行), 列: Math.min(ば左.列, ば右.列),
                    板: 四角の板(ば左, ば右) } : null;
              return {
                種: '四角', 並び: 四.並び, 行数: 四.行数, 列数: 四.列数, 番地: 角,
                /* ★★何番目の マスに 何と 打って 在るか★★（2026-09-15）
                   ＝SUBTOTAL は ★入れ子の SUBTOTAL を 数えません★（実測）
                     A6 に `=SUBTOTAL(9,A1:A5)` が 在る 時
                       `=SUBTOTAL(9,A1:A6)` → ★15★ ／ `=SUM(A1:A6)` → ★30★
                   ⇒★値だけでは 分かりません★＝★元の 字★が 要る。
                   ★呼ばれた 時だけ 作ります★＝★四角を 読む たびに 名前を 並べない★
                     （実物は 12,383回 INDEX/MATCH を 呼びます＝ここを 太らせない） */
                打った字: 角 ? function (i) {
                  var m = 中身[番地から名(角.行 + Math.floor(i / 四.列数), 角.列 + (i % 四.列数), 角.板)];
                  return m ? String(m.打った字 || '') : '';
                } : null,
              };
            }
            if (子.種 === '名') {
              var nn = String(子.値).toUpperCase();
              var ば = 名から番地(nn);
              if (nn !== 'TRUE' && nn !== 'FALSE' && ば) {
                /* ★どの マスかも 渡す★（2026-09-14）
                   ＝`=CELL("row",B7)` は ★中身では なく 番地★が 要ります。
                     前は 値しか 渡さず、自前の 皮の CELL が ★#VALUE!★ を 返して いました。
                   ★足すだけ★＝今まで 読んで いた `値` は そのまま。 */
                var n2 = 揃える(nn);
                return {
                  種: 'マス', 値: 読む(n2, 通り道), 名前: n2, 番地: { 行: ば.行, 列: ば.列, 板: ば.板 || 今の板 },
                  打った字: function () { return 中身[n2] ? String(中身[n2].打った字 || '') : ''; },
                };
              }
            }
            var 値 = 歩く(子, 手, 通り道);
            /* ★参照は 四角と 同じ 顔で 渡す★（2026-09-15）
               ＝`=SUM(INDEX(A1:B5,0,2))` → ★30★（実測）。★番地も 付いて いる★ので
                 中で もう1回 `暗黙の 交わり` が 要る 時にも 使えます。 */
            if (場.参照か(値)) return 場.引数にする(値);
            /* ★表を 直に 書いた 物は 四角として 渡す★（`{1;2}` は 1列の 四角） */
            if (溢.溢れか(値)) {
              return { 種: '四角', 並び: 値.並び.reduce(function (a, 段) { return a.concat(段); }, []),
                行数: 値.行数, 列数: 値.列数 };
            }
            return { 種: '直', 値: 値 };
          };
          /* ══ ★★列まるごと・行まるごとを 引数に 持つ 呼び★★ ══（2026-11-04・案X）
               ★帯は 板の 端まで 切って 読む★。1つの 呼びの 中の 帯は ★同じ 所まで★ 切る（板が 違っても 形が ずれない）。
               ★置いて よい 所★（帯を置いてよいか）の 外は 前と 同じ #REF!。 */
          var 呼名U = 呼び名.replace(/^_XLFN\.|^_XLWS\./, '');
          var 帯たち = [], 帯の伸ばし = { 行: 0, 列: 0 };
          (木.引数 || []).forEach(function (子) {
            var 並 = (子 && 子.種 === '束') ? 子.並 : [子];
            for (var b = 0; b < 並.length; b++) {
              var 帯b = 帯の形(並[b]);
              if (!帯b) continue;
              var 端b = 帯の端(帯b) + 1;   /* ★1つの 呼びの 中の 帯は 同じ 所まで 切る（形を 揃える）★ */
              if (帯b.列全体 && 端b > 帯の伸ばし.行) 帯の伸ばし.行 = 端b;
              if (帯b.行全体 && 端b > 帯の伸ばし.列) 帯の伸ばし.列 = 端b;
              if (!帯の伸ばし.行 && !帯の伸ばし.列) 帯の伸ばし.行 = 1;
            }
          });
          /* ★SUMIF の 取り直した 足す 範囲も 同じ 所まで 切る★（足す 範囲の 尻尾の 数を 落とさない） */
          var 足す帯 = (呼名U === 'SUMIF' && (帯の伸ばし.行 || 帯の伸ばし.列)) ? SUMIFの足す帯(木.引数 || []) : null;
          if (足す帯) {
            var 端S = 帯の端(足す帯) + 1;
            if (足す帯.列全体 && 端S > 帯の伸ばし.行) 帯の伸ばし.行 = 端S;
            if (足す帯.行全体 && 端S > 帯の伸ばし.列) 帯の伸ばし.列 = 端S;
          }
          var 帯よし = function () { return false; };
          if (帯の伸ばし.行 || 帯の伸ばし.列) {   /* ★帯が 1つでも 在れば 1 以上★ */
            var 置き = 帯の置き場(呼名U, 木.引数, 手, 通り道);
            if (置き.伸ばす行 > 帯の伸ばし.行) 帯の伸ばし.行 = 置き.伸ばす行;
            if (置き.伸ばす列 > 帯の伸ばし.列) 帯の伸ばし.列 = 置き.伸ばす列;
            帯よし = 置き.よい;
          }
          var 引数たち = 木.引数.map(function (子, i) {
            /* ★★SUMIF・AVERAGEIF の 足す 範囲（3つ目）★★（2026-11-07・本番前 taiketsu 止め2）
                 ①★省いたかは 式の 木で 決める★（`SUMIF(A1:A5,">2",)` の 3つ目は 木が 空）＝前は 値の 型（空）で 決めて いたので
                   空の マスを 指す 式（`(D1)`）も 省いた 扱い＝★条件の 範囲を 黙って 足した★（12・Excel は 0）
                 ②★括弧は はがす★＝門（左上の番地）と 同じ 形＝`SUMIF(A1:A5,">2",(D1))` も 番地の 付いた 参照で 届く */
            if ((呼名U === 'SUMIF' || 呼名U === 'AVERAGEIF') && i === 2) {
              if (子 && 子.種 === '空') return { 種: '直', 値: 計.空, 省いた: true };
              while (子 && 子.種 === '括') 子 = 子.子;
            }
            return 引数にする(子, 帯よし(i));
          });
          /* ★★SUMIF・AVERAGEIF の 足す 範囲を 取り直す 時は ★頼り（頼りを拾う）が 取り直した 所を 知る 形★だけ 通す★★
               （2026-11-07・設計前 taiketsu 止め1）＝`SUMIF(INDEX(A1:A5,0),">2",B1)`・交わり・LET の 名 等は
               打った 時は 正しくても 打ち直すと 古い 数（B3=1 で 120 の まま・正しくは 91）⇒ 形が 違えば #REF!（黙った 数に しない） */
          if ((呼名U === 'SUMIF' || 呼名U === 'AVERAGEIF') && 引数たち.length >= 3 && !取り直しを頼れるか(木.引数)) {
            var 大0 = 引数の大きさ(引数たち[0]), 大2 = 引数の大きさ(引数たち[2]);
            if (大0 && 大2 && (大0.行数 !== 大2.行数 || 大0.列数 !== 大2.列数)) return 計.誤('#REF!');
          }
          if (帯たち.length) {
            var 形の誤 = 帯の形を合わせる(呼名U, 引数たち, 足す帯) || 帯の条件を見る(呼名U, 引数たち);
            if (形の誤) return 形の誤;
            if (呼名U === 'ROWS' && 引数たち[0] && 引数たち[0].帯) return 計.数(引数たち[0].帯.本当の行数);
            if (呼名U === 'COLUMNS' && 引数たち[0] && 引数たち[0].帯) return 計.数(引数たち[0].帯.本当の列数);
          }
          /* ★今 いる マスも 渡す★（★暗黙の 交わり★に 要る＝INDEX/MATCH が 使う） */
          /* ★★2026-09-18 ... ★マスを 読む 口★も 渡します★★
               ★なぜ★ ... `INDIRECT` と `OFFSET` は ★式の 中で 場所を 作って 読みます★
                 ＝引数として 渡された 四角では 届きません
               ★★通り道を そのまま 渡します★★＝★輪（循環）は 今まで通り 見つかります★
               ★口は 1つだけ★（`四角を取る`）＝★台に 板の 中を 触らせません★
               ★渡すのは 読む 事だけ★＝★書く 口は 渡しません★ */
          var 所を作る = function () {
            return {
              今のマス: 今のマス,
              /* ★★ラムダの 口★★（2026-09-18）＝MAP／REDUCE／SCAN／BYROW／BYCOL／MAKEARRAY が 使う
                   ★`lib/shiki-kansuu.js` は 木を 知りません★ので ★呼ぶ 口だけ 渡します★ */
              ラムダか: ラムダか,
              ラムダを呼ぶ: function (L, 引値) { return ラムダを呼ぶ(L, 引値, 手, 通り道); },
              /** ★行・列（0から）から 四角を 取る★ ... 返しは `場.参照`（無ければ null）
                    ★板を またぎません★（★今の 板だけ★＝`INDIRECT("別板!A1")` は まだ） */
              /* ★板（5つ目）を 渡せば その 板を 読む★（2026-11-07・OFFSET の 板またぎ）＝無ければ 今の 板 */
              四角を取る: function (行, 列, 行数, 列数, 板) {
                if (!(行 >= 0) || !(列 >= 0) || !(行数 >= 1) || !(列数 >= 1)) return null;
                var 読む板 = 板 ? String(板).toUpperCase() : 今の板;
                if (板 && !板たち[読む板]) return null;
                var 並び = [];
                for (var r = 行; r < 行 + 行数; r++) {
                  for (var c = 列; c < 列 + 列数; c++) {
                    var なまえ = 番地から名(r, c, 読む板);
                    if (!なまえ) return null;
                    並び.push(読む(なまえ, 通り道));
                  }
                }
                var 返す = 場.参照(行, 列, 行数, 列数, 並び, false);
                if (板) 返す.板 = 読む板;      /* ★INDEX と 同じ＝`:` の 端でも 板を 落とさない★ */
                return 返す;
              },
              /** ★字（"A1" や "A1:B2"）から 四角を 取る★ ... 返しは `場.参照`（読めなければ null） */
              字から取る: function (字) {
                var s = String(字 == null ? '' : 字).replace(/^\s+|\s+$/g, '');
                if (!s.length) return null;
                if (s.indexOf('!') >= 0) return null;      /* ★板またぎは まだ★ */
                var 割 = s.split(':');
                if (割.length > 2) return null;
                var a = 名から番地(揃える(割[0]));
                if (!a) return null;
                var b = (割.length === 2) ? 名から番地(揃える(割[1])) : a;
                if (!b) return null;
                var r0 = Math.min(a.行, b.行), r1 = Math.max(a.行, b.行);
                var c0 = Math.min(a.列, b.列), c1 = Math.max(a.列, b.列);
                return this.四角を取る(r0, c0, r1 - r0 + 1, c1 - c0 + 1);
              },
            };
          };
          var 出 = 関.呼ぶ(木.名, 引数たち, 手, 所を作る());
          if (出 !== null && 帯たち.length) {
            出 = 帯の余りを足す(呼名U, 引数たち, 出, function (条) {
              var 一 = 関.呼ぶ('COUNTIF', [{ 種: '四角', 並び: [計.空], 行数: 1, 列数: 1 }, 条], 手, 所を作る());
              return !!(一 && 一.型 === '数' && 一.値 === 1);
            });
          }
          if (出 !== null) { 答えた者を数える('土台'); return 出; }
          /* ★土台が 知らない 物は 皮に 回す★（2026-09-14／B-3）
             ＝`lib/shiki-tsunagi.js` が 持つ ★61個★（借り物に 無い 物を 自前で 書いた 分）
             ★皮も 知らなければ null★ ⇒ ★#NAME?★（★半分 合う 答えを 出さない★） */
          if (皮 && typeof 皮.呼ぶ === 'function') {
            /* ★今 いる マスも 渡す★（引数なしの CELL 等が 使う） */
            var 皮の出 = 皮.呼ぶ(木.名, 引数たち, 所を作る());
            if (皮の出 !== null && 皮の出 !== undefined) {
              /* ★溢れは まだ 置けません（土台⑤）★
                 ★左上だけ 返して 誤魔化す事は しません★＝答えは 溢れの まま 持たせ、
                 ★並べるのは ⑤の 仕事★。数だけ 別に 取って「出来て いない」を 隠さない。 */
              答えた者を数える(皮の出 && 皮の出.溢れ === true ? '皮（溢れ待ち）' : '皮');
              return 皮の出;
            }
          }
          答えた者を数える('知らない');
          /* ★知らない 関数は ★#NAME?★ と 言う＝半分 合う 答えを 出さない★ */
          return 計.誤('#NAME?');
        }
        default: return 計.誤('#VALUE!');
      }
    }

    /* ★`:` の 端を 1つの マスの 名前に する★（2026-09-15）
       ★名前なら そのまま（`@` も 残す）／参照なら その 角★
       ★それ以外は null★＝`#REF!`（★半分 合う 答えを 出さない★） */
    /* ══ ★★参照の 四角（交わり・溢れの 参照）★★ ══（2026-10-05・経営者の 物差し 137 の 組 G・O・P・Q）
         ★交わり★（空白の 演算子）`=SUM(A1:B2 B1:C2)` → 重なった B1:B2 ／ 重ならなければ ★#NULL!★
         ★溢れの 参照★ `E1#` → E1 の 溢れの 範囲 ／ E1 が #SPILL! なら ★その 誤り★（実Excel #スピル!）／溢れて いなければ #REF!
         ★出す 形★ { 板, 行0, 列0, 行1, 列1 } ／ 計.誤(誤りの字) ／ null（参照では ない）
         ★求めた 四角は `:` の 道（四角を読む）に 乗せる★＝読み方を 2つに しない */
    function 参照の角(木, 手, 通り道) {
      if (!木) return null;
      if (木.種 === '括') return 参照の角(木.子, 手, 通り道);
      if (木.種 === '名') {
        var n = String(木.値).toUpperCase();
        if (n.length > 1 && n.charAt(n.length - 1) === '#') {
          var 元名 = n.slice(0, -1), ば元 = 名から番地(元名);
          if (!ば元) return 計.誤('#REF!');
          var 元n = 揃える(元名), m元 = 中身[元n];
          if (!m元 || !m元.形) return 計.誤('#REF!');
          if (m元.値 === undefined) {
            if (通り道 && 通り道.indexOf(元n) >= 0) { if (残りに在る(元n)) 残り.疑い = true; return 計.誤('#REF!'); }   /* ★輪★ */
            /* ★④の 残りの 間は「まだ」の 道も 深さで 切る★（2026-11-07・taiketsu 叩き直し2 止め2）＝溢れの 鎖 E1# が 700段で RangeError だった */
            var 直角 = まだを直すか(元n, 通り道);
            if (直角 === 2) return 計.誤('#REF!');
            if (直角 === 1) 計算する(元n, (通り道 || []).concat([元n]));
            if (m元.値 === undefined && 残りに在る(元n)) { 残り.疑い = true; return 計.誤('#REF!'); }
          }
          var v元 = m元.値;
          if (v元 && v元.型 === '誤') return v元;
          if (!溢.溢れか(v元)) return 計.誤('#REF!');
          var ば元2 = 名から番地(元n);
          return { 板: ば元2.板 || 今の板, 行0: ば元2.行, 列0: ば元2.列, 行1: ば元2.行 + v元.行数 - 1, 列1: ば元2.列 + v元.列数 - 1 };
        }
        var ば = 名から番地(n);
        if (!ば) return null;
        return { 板: ば.板 || 今の板, 行0: ば.行, 列0: ば.列, 行1: ば.行, 列1: ば.列 };
      }
      if (木.種 === '二' && 木.記 === ':') {
        var 対3 = 端の対(木.左, 木.右, 手, 通り道);   /* ★板の 違う 端は #VALUE!★（端の対 の 注） */
        if (対3.誤) return 計.誤(対3.誤);
        var 左字 = 対3.左, 右字 = 対3.右;
        var a = 名から番地(左字.charAt(0) === '@' ? 左字.slice(1) : 左字), b = 名から番地(右字);
        if (!a || !b) return 計.誤('#REF!');
        return { 板: 四角の板(a, b), 行0: Math.min(a.行, b.行), 列0: Math.min(a.列, b.列), 行1: Math.max(a.行, b.行), 列1: Math.max(a.列, b.列) };
      }
      if (木.種 === '二' && 木.記 === ' ') {
        var L = 参照の角(木.左, 手, 通り道), R = 参照の角(木.右, 手, 通り道);
        if (L && L.型 === '誤') return L;
        if (R && R.型 === '誤') return R;
        if (!L || !R) return 計.誤('#VALUE!');
        if (String(L.板).toUpperCase() !== String(R.板).toUpperCase()) return 計.誤('#VALUE!');
        var 角 = { 板: L.板, 行0: Math.max(L.行0, R.行0), 列0: Math.max(L.列0, R.列0), 行1: Math.min(L.行1, R.行1), 列1: Math.min(L.列1, R.列1) };
        if (角.行0 > 角.行1 || 角.列0 > 角.列1) return 計.誤('#NULL!');
        return 角;
      }
      return null;
    }
    /* ★四角を `:` の 木に 戻す★（★今まで 通りの 道で 読む★） */
    function 角から木(角) {
      return { 種: '二', 記: ':', 左: { 種: '名', 値: 番地から名(角.行0, 角.列0, 角.板) }, 右: { 種: '名', 値: 番地から名(角.行1, 角.列1, 角.板) } };
    }
    /* ★交わり か 溢れの 参照 か★ */
    function 角で読む木か(木) {
      return !!木 && ((木.種 === '二' && 木.記 === ' ') || (木.種 === '名' && /.#$/.test(String(木.値))));
    }

    function 端の名(子, 手, 通り道, 右か) {
      var s = String(子 && 子.値 != null ? 子.値 : '');
      if (子 && 子.種 === '名') {
        var t = s.charAt(0) === '@' ? s.slice(1) : s;
        if (名から番地(t)) return s;
      }
      var v = 歩く(子, 手, 通り道);
      if (場.参照か(v) && !v.番地なし) {      /* ★番地の 無い 参照は 端に なれません★ */
        var r = 右か ? v.行 + v.行数 - 1 : v.行;
        var c = 右か ? v.列 + v.列数 - 1 : v.列;
        /* ★参照の 値は 元の 板を 持ちます★（INDEX が 付ける・2026-11-07）＝★無い 時だけ 今の 板★
           （★実物に `SUM(INDEX(他板!…):A5)` の 形は 0本★） */
        return 番地から名(r, c, v.板 || 今の板);
      }
      return null;
    }

    /* ★★`:` の 両端を 揃える★★（2026-11-07）
       ★端の 片方でも 参照（INDEX 等）から 来て いて、両端の 板が 違えば #VALUE!★
       ＝実Excel 16.0 実測：`=SUM(INDEX(S2!A:A,2):A3)`・`=INDEX(S2!A1:A100,2):Sheet1!A3` → ★#VALUE!★
         （前は 端を 今の 板に 読み替え ★今の 板の A2:A3＝5 を 黙って 出して いた★）
       ★両端とも ふつうの 名前の 時は 前と 同じ★（`Sheet1!A1:A3` の 右は 左の 板＝四角の板 が 決める） */
    function 名の端か(子) {
      if (!子 || 子.種 !== '名') return false;
      var s = String(子.値 != null ? 子.値 : '');
      return !!名から番地(s.charAt(0) === '@' ? s.slice(1) : s);
    }
    function 端の対(左子, 右子, 手, 通り道) {
      var 左 = 端の名(左子, 手, 通り道, false);
      var 右 = 端の名(右子, 手, 通り道, true);
      if (左 === null || 右 === null) return { 誤: '#REF!' };
      if (!名の端か(左子) || !名の端か(右子)) {
        var ば左 = 名から番地(左.charAt(0) === '@' ? 左.slice(1) : 左), ば右 = 名から番地(右);
        var 板左 = (ば左 && ば左.板) || String(今の板).toUpperCase();
        var 板右 = (ば右 && ば右.板) || String(今の板).toUpperCase();
        if (板左 !== 板右) return { 誤: '#VALUE!' };
      }
      return { 左: 左, 右: 右 };
    }

    /* ★四角は「形（行数×列数）」も 返す★（2026-09-14 に 足した）
       ★並びは 平らな まま★＝`shiki-kansuu.js` の `ほどく` が 平らに 舐めるので 壊さない。
       ★形が 要る 訳★＝横に 広がる 関数が 在る。実物で 出した（材料 A1:B3 = 1,2/3,4/5,6）：
         `=ARRAYTOTEXT(A1:B3,1)` … 実Excel `{1,2;3,4;5,6}`
            形が 無いと ★`{1;2;3;4;5;6}`★（★縦一列に 潰れる★）
         `=MINVERSE(D1:E2)` … 形が 無いと ★#VALUE!★（正方行列に ならない）
         `=LINEST(A1:A5,B1:B5)` … ★縦1列なので 形が 無くても 当たる★
       ⇒★「今 動いて いる」は「形が 要らない」では ない★ */
    /* ══ ★★四角は 番地の まま 回します★★ ══（2026-09-25）
         ★★何が 在ったか（経営者1 の CPU の 記録・`7b43ee6` の 木で 3回）★★
           `名から番地` ... ★41.4%★（覚えさせた 後も まだ 1位）
           `番地から名` ... ★16.4%★
           `四角を読む` ... 7.2% ／ `揃える` 0.8%
           ⇒★この 4つ だけで 65.8%★
         ★★なぜ 覚えても 減らないか★★
           覚えたのは ★1回ごとの 中身★ です（正規表現が 消えた＝3.3倍に なった）。
           でも ★呼ばれる 回数 1億4,192万回 は 1回も 減って いません★。
           ⇒★次は 「呼ぶ 回数を 減らす」★（経営者1 の 言）
         ★★前は 何を して いたか★★
           マス 1つ 読む たびに
             ①`番地から名(r,c,板)` で ★字を 作り★
             ②`読む` の 中の `揃える` が ★その 字を 番地に 戻し★
             ③★もう一度 字に 作り直し★
             ④`#REF!` を 見る 為に ★また 番地に 戻して★ いました
           ＝★1マスに 名から番地 2回・番地から名 2回★
           ＝経営者1 の 数と 合います（名から番地 は 読む の 2.24倍／番地から名 は 2.02倍）
         ★★今は どう するか★★
           ・板の 名を ★輪の 外で 1回だけ★ 大文字に する
           ・`#REF!`（知らない 板）も ★輪の 外で 1回だけ★ 見る
           ・輪の 中は ★字を 1回 作るだけ★（`中身` は 字で 引く 台なので 字は 要る）
           ⇒★1マスあたり 名から番地 2回 → 0回 ／ 番地から名 2回 → 1回★
         ★★揃え方を 変えて いません★★（★ここを 間違えると 別の マスを 読みます★）
           `揃える(名)` ＝ `番地から名(ば.行, ば.列, ば.板 || 今の板)`
           `ば.板` は `名から番地` が ★大文字に した 板の 名★ です。
           ⇒だから ここでも ★`板3.toUpperCase()`★ を 使います。
           ⇒★`Sheet1` と `SHEET1` は 別の 鍵に なります★＝★揃えないと 空を 読みます★
         ★`読む` その物は 1文字も 変えて いません★（外から 呼ぶ 道は そのまま） */
    function 四角を読む(左, 右, 通り道) {
      var a = 名から番地(左), b = 名から番地(右), 並び = [];
      if (!a || !b) return { 並び: 並び, 行数: 0, 列数: 0 };
      var r0 = Math.min(a.行, b.行), r1 = Math.max(a.行, b.行);
      var c0 = Math.min(a.列, b.列), c1 = Math.max(a.列, b.列);
      var 板3 = 四角の板(a, b);          /* ★両端は 同じ 板★ */
      var 板U = 板3 ? String(板3).toUpperCase() : 板3;
      /* ★知らない 板は この 四角 まるごと `#REF!`★（前は 1マスずつ 見て いた） */
      var 知らない板 = !!(板U && !板たち[板U]);
      数え.四角を読む++; 数え.四角のマス += (r1 - r0 + 1) * (c1 - c0 + 1);
      if (四角の種類) 四角の種類[String(板U) + '!' + r0 + ',' + c0 + ':' + r1 + ',' + c1] = (四角の種類[String(板U) + '!' + r0 + ',' + c0 + ':' + r1 + ',' + c1] || 0) + 1;
      var 前の行表 = (残り && 残り.前の所) || null;
      var 控え鍵 = (!残り && !知らない板) ? String(板U) + '!' + r0 + ',' + c0 + ':' + r1 + ',' + c1 : null;
      if (控え鍵) {
        var 控 = 四角の控え.get(控え鍵), 印 = 四角の印[控え鍵];
        if (控 && 控.版 === 端の版) { 四角の控え.delete(控え鍵); 四角の控え.set(控え鍵, 控); 数え.四角の控えで++; return { 並び: 控.並び.slice(), 行数: r1 - r0 + 1, 列数: c1 - c0 + 1 }; }
      }
      var 見た版 = !!(控え鍵 && 印 && 印.版 === 端の版);   /* ★1回目は 印だけ・2回目に 値だけかを 見て 控える★ */
      var 値だけ = 見た版 && !印.式あり;
      for (var r = r0; r <= r1; r++) {
        for (var c = c0; c <= c1; c++) {
          if (知らない板) { 並び.push(計.誤('#REF!')); continue; }
          /* ══ ★★索引に 無い マスは 字を 作らずに 空を 返します★★ ══（2026-09-26）
               ＝経営者1 の 実測 ★1打ち目 15.1秒の 91.4%（5,700万回）が ここ★
               ⇒★索引（中身＋覆い）に 無ければ 字も 引き出しも やめます★
               ★★輪の 話に なりません★★＝通り道に 入るのは 形を 持つ マスだけ
                 ＝★中身に 無い マスは 形を 持ちません★（経営者1 の 実測でも 輪は 0回）
               ★覆い（溢れ）も 索引に 入って います★
               ★効いた 回数は `空と答えた` で 出します★ */
          if (!在るか(板U, r, c)) {
            /* ★④の 残りで 計算中の 元の 溢れ（覆いを 消して 索引から 抜けた マス）は 読む揃った名で へ★（2026-11-07・叩き直し3 止め1） */
            if (前の行表 && 前の行表[r] && 前の行表[r][c]) {   /* ★行ごとの 表＝字を 作らずに 引く★（叩き直し4 任意＝広い 範囲で 6倍 遅く なった） */
              var 名前 = 番地から名(r, c, 板U);
              if (残り.前の覆い[名前]) { 並び.push(読む揃った名で(名前, 通り道)); continue; }
            }
            空と答えた++; 並び.push(空に()); continue;
          }
          数え.四角で在った++;
          var 名q = 番地から名(r, c, 板U);
          if (値だけ) { var mq = 中身[名q]; if (!mq || mq.形 || 覆い[名q]) 値だけ = false; }
          並び.push(読む揃った名で(名q, 通り道));
        }
      }
      if (控え鍵 && !見た版) {
        if (四角の印の数 >= 4096) { 四角の印 = {}; 四角の印の数 = 0; }
        if (!四角の印[控え鍵]) 四角の印の数++;
        四角の印[控え鍵] = { 版: 端の版, 式あり: false };
      } else if (値だけ) {
        var 古 = 四角の控え.get(控え鍵); if (古) { 四角の控え.delete(控え鍵); 控えのマス -= 古.並び.length; }
        if (並び.length <= 控えのマスの上限) {
          while (四角の控え.size >= 4096 || 控えのマス + 並び.length > 控えのマスの上限) { var 先 = 四角の控え.keys().next().value; 控えのマス -= 四角の控え.get(先).並び.length; 四角の控え.delete(先); }
          四角の控え.set(控え鍵, { 版: 端の版, 並び: 並び.slice() }); 控えのマス += 並び.length;
        }
      } else if (控え鍵) 印.式あり = true;
      return { 並び: 並び, 行数: r1 - r0 + 1, 列数: c1 - c0 + 1 };
    }

    /* ★1つの マスの 値を 読む★（★輪を ここで 見つける★） */
    function 読む(名, 通り道) {
      数え.読む++;
      var n = 揃える(名);
      /* ★★知らない 板を 指したら #REF!★★（2026-09-15・㍸）
         ★空の 板（在るけど 中身が 無い）は 0★／★無い 板は #REF!★＝★別物★
         （Excel も 同じ＝板を 消すと 式が `=#REF!A1` に なる） */
      var ば板 = 名から番地(n);
      if (ば板 && ば板.板 && !板たち[ば板.板]) return 計.誤('#REF!');
      return 読む揃った名で(n, 通り道);
    }

    /* ★★揃った 名で 読む★★（2026-09-25 に `読む` から 切り出しました）
         ★中身は 1文字も 変えて いません★＝`読む` の 後半 その物です。
         ★切り出した 訳★ ... 四角を 回す 時は
           ・名は ★こちらで 作って いる★＝もう 揃って いる（揃え直す 必要が 無い）
           ・板が 在るかは ★輪の 外で 1回 見れば よい★
         ⇒★1マスごとの 字の 出し入れが 減ります★ */
    function 読む揃った名で(n, 通り道) {
      数え.揃った名で++;
      var m = 中身[n];
      /* ★消した マス（打った字 '' の 入れ物）が 溢れで 埋まって いれば 溢れを 読む★（2026-10-05・組 J／N）
           ＝塞いだ 物を 消すと 中身に 空の 入れ物が 残る。前は それが 先に 空を 返し、E2 だけ 空に なった */
      if (m && !m.形 && m.打った字 === '' && 覆い[n]) m = null;
      if (!m) {
        数え.揃った名で無かった++;
        /* ★溢れで 埋まった マスか★（土台⑤）＝式は 無いが 値は 在る */
        var 覆 = 覆い[n];
        /* ★④の 残りで 計算中の 元の 溢れ★（2026-11-07・taiketsu 叩き直し3 止め1）＝計算する() が 先に 覆いを 消すので
             元が 通り道に 在る 間に 溢れの マス（F1）を 読むと 空＝0 だった（U5=`=SUM(E1#)` が 30・正 61）。
             マスの 輪（読む揃った名で の 通り道）と 同じく 1回目は 毒・2回目からは 前の 回の 値 */
        if (!覆 && 残り && 残り.前の覆い && 残り.前の覆い[n]) {
          var 前覆 = 残り.前の覆い[n];
          if (通り道 && 通り道.indexOf(前覆.元) >= 0 && 残りに在る(前覆.元)) {
            残り.疑い = true;
            var 前元m = 中身[前覆.元];
            if (残り.回 === 0 || !前元m || !溢.溢れか(前元m.値)) return 計.誤('#REF!');
            var 前行 = 前元m.値.並び[前覆.行];
            return (前行 && 前行[前覆.列] !== undefined) ? 前行[前覆.列] : 計.誤('#REF!');
          }
        }
        if (覆) {
          var 元m = 中身[覆.元];
          /* ★元の 式が「まだ」なら 先に 計算する★（2026-11-07・taiketsu 止め2）
               ＝揮発の 直しで 元が undefined の 間に 読まれると 空（0）を 返して いた */
          if (元m && 元m.形 && 元m.値 !== undefined) {
            var 直0 = 残りで直すか(覆.元, 通り道);
            if (直0 === 2) return 計.誤('#REF!');
            if (直0 === 1) 計算する(覆.元, (通り道 || []).concat([覆.元]));
          }
          if (元m && 元m.形 && 元m.値 === undefined) {
            if (通り道 && 通り道.indexOf(覆.元) >= 0) { if (残りに在る(覆.元)) 残り.疑い = true; return 計.誤('#REF!'); }  /* ★輪★ */
            var 直0u = まだを直すか(覆.元, 通り道);   /* ★④の 残りの 間は 深さで 切る★（叩き直し2 止め2） */
            if (直0u === 2) return 計.誤('#REF!');
            if (直0u === 1) {
              計算する(覆.元, (通り道 || []).concat([覆.元]));
              if (覆い[n] !== 覆) return 読む揃った名で(n, 通り道);   /* ★計算で 覆いが 変わった★ */
            }
            if (元m.値 === undefined && 残りに在る(覆.元)) { 残り.疑い = true; return 計.誤('#REF!'); }
          }
          if (元m && 溢.溢れか(元m.値)) return 元m.値.並び[覆.行][覆.列];
        }
        return 空に();
      }
      if (m.形) {
        if (通り道 && 通り道.indexOf(n) >= 0) {
          /* ★④の 残りでは 前の 値★（2026-11-07・案C）＝偽の 輪（VLOOKUP の 範囲が 自分を 頼る 式を 含む 等）は
               実Excel では 輪で ない（使わない マス）。★1回目は #REF!（毒）★・2回目からは ★前の 回の 値★を 返し、回を 重ねて 動かなく なるまで 回す
               ＝毒を 本当に 使う 式（`=SUM(U1:U4)+0*SUM(U1:U5)`・U5=`=X1+1` の 様な 本当の 輪＝実Excel も 輪の 警告）は #REF! の まま 残り、使わない 式は 戻る */
          if (残り && 残り.在る[n]) { 残り.疑い = true; return (残り.回 === 0 || m.値 === undefined) ? 計.誤('#REF!') : (溢.溢れか(m.値) ? 溢.先頭(m.値) : m.値); }
          return 計.誤('#REF!');  /* ★輪★ */
        }
        if (m.値 === undefined) {
          var 直u = まだを直すか(n, 通り道);   /* ★④の 残りの 間は「まだ」の 道も 深さで 切る★（叩き直し2 止め2＝まとめて打つ で 鎖 700段 RangeError） */
          if (直u === 2) return 計.誤('#REF!');
          if (直u === 1) 計算する(n, (通り道 || []).concat([n]));
          if (m.値 === undefined && 残りに在る(n)) { 残り.疑い = true; return 計.誤('#REF!'); }   /* ★まだの まま＝空(0)に しない★ */
        } else {
          var 直1 = 残りで直すか(n, 通り道);
          if (直1 === 2) return 計.誤('#REF!');
          if (直1 === 1) 計算する(n, (通り道 || []).concat([n]));
        }
      }
      if (m.値 === undefined) return 空に();
      /* ★1マスとして 読む時は 溢れの 左上★（実Excel の 暗黙の交差） */
      return 溢.溢れか(m.値) ? 溢.先頭(m.値) : m.値;
    }

    /* ★溢れの 覆い★＝どの マスが どの 式の 溢れで 埋まって いるか
       { 埋まった マス名: { 元: 式の マス名, 行: i, 列: j } } */
    /* ★今 計算して いる マス★（`=CELL("row")` の ような 引数なしの 物が 使う） */
    var 今のマス = null;

    /* ★★④の 残り（順番で 解けなかった 物）を 直す 間の 印★★（2026-11-07・案C）
         在る＝残りの 名／済＝この 回に 計算した 名／疑い＝この 回に 前の 値を 使った */
    var 残り = null;
    var 残りの深さ = 32;
    /* ★輪の 塊 1つを 直す 読みの 予算★（2026-11-07・本番前 taiketsu 止め2）＝自分の 列を 含む 範囲を 引く 表（C=`VLOOKUP(Ai,$A:$C,2,0)`）は
         塊を 2回 回して 毎回 範囲 全部を 読む＝2000行で 読み 2400万回・1打ち 16秒。越えたら 投げる＝外（book.html の catch）は 借り物へ
         ・★塊ごと★に 数える（設計前 taiketsu 止め2＝④全体で 数えると 100行の 表 50個 の 本が 開けない）
         ・300万＝700行で 約1秒・800行で 投げる（本番は 700〜800行で 深さが 切れて 借り物）。今の 試験の 重い 形は 3万でも 緑（taiketsu 実測） */
    var 残りの予算 = 3000000;
    /* ★壊れた 台★（2026-11-07・設計前 taiketsu 止め1）＝直す群 の 途中で 投げると 下流の 式に 古い 数が 残り、
         次の 打つ・値・字 が 黙って 違う 数を 返した（800行で 投げた 後 D1 を 見る F1 が 在ると D1・F1・G1 が 640800・正 640894）
         ⇒★投げたら 壊れに し、以後 打つ・置く・値・字・まとめて打つ を 全部 投げる★（外は 借り物へ。作り直せば 戻る） */
    var 壊れた = null;
    function 壊れを見る() { if (壊れた) throw new Error('台は 壊れて います（' + 壊れた + '）＝借り物へ'); }   /* ★呼び込みの 深さ★＝これより 深い 所は 前の 値で 次の 回へ（深い 鎖で スタックが 溢れない） */
    /* ★0＝今の 値を 読む／1＝先に 計算する／2＝毒（#REF!）★
         ＝深さで 切った 所は ★1回目は 毒★（2026-11-07・taiketsu 叩き直し 止め2）。前の 値を 読ませると
           32段より 長い 本当の 輪が 打つ 前の 数を 黙って 残した（A1=A2〜A33=IF(C1=1,7,A1) で C1=2 → 7 の まま） */
    function 残りで直すか(n, 通り道) {
      if (!残り || !残り.在る[n] || 残り.済[n]) return 0;
      if ((通り道 ? 通り道.length : 0) >= 残りの深さ) { 残り.疑い = true; return 残り.回 === 0 ? 2 : 0; }
      残り.済[n] = true;
      return 1;
    }
    /* ★「まだ」（値 undefined）の 式を 読む 時★＝④の 残りに 在る 間は 残りで直すか（深さ・毒）に 通す。外は 今まで どおり 先に 計算（1）
         （2026-11-07・taiketsu 叩き直し2 止め2＝まだの 道 3か所に 深さの 守りが 無かった） */
    function 残りに在る(n) { return !!(残り && 残り.在る[n]); }
    function まだを直すか(n, 通り道) {
      if (!残りに在る(n)) return 1;
      if (残り.済[n]) return 0;
      return 残りで直すか(n, 通り道);
    }
    function 値の鍵(v) { if (v === undefined) return 'u'; try { return JSON.stringify(v); } catch (e) { return String(v); } }

    /* ══ ★★名の台（束縛）★★ ══（2026-09-18・ラムダの 一族の 段取り ⑴）
         ★何の 為か★ ... `=LET(x,1,x+1)` の `x` を 読める ように する
           ＝今までは `case '名'` が TRUE／FALSE／★マスの 名前★の 3つしか 知らず
             それ以外は ★#NAME?★ でした。
         ★形★ ... { 表: { X: <値> }, 親: <外側の 台> }
           ＝★親を 辿ります★（`=LET(a,1,LET(b,2,a+b))` が 動く為）
         ★★なぜ 引数で 回さず 1つの 変数か★★
           ＝`歩く()` の 呼び出しは ★40か所 以上★在ります。
           ＝★`今のマス` `今の板` と 同じ 手★（★前の 物を 覚えて 必ず 戻す★）
           ＝★1本の 糸で 順に 計算して います★（並びでは ない）ので これで 足ります。
         ★★別の マスを 読む 時は 必ず 空に します★★
           ＝`計算する()` で `null` に 戻します
           ＝★でないと A1 の LET の 名前が A2 の 式から 見えます★（★他人の 名前★） */
    var 名の台 = null;

    /* ══ ★★ラムダという 値★★ ══（2026-09-18・段取り ⑵）
         ★形★ { 型: 'ラムダ', 名たち: ['X'], 体: <木>, 台: <作られた 時の 名の台> }
         ★★作られた 時の 台を 一緒に 持ちます★★＝★閉じ込め★
           ＝`=LET(a,2,SUM(MAP(A1:A5,LAMBDA(x,x*a))))` が 動く 為
           ＝★実Excel は 30 を 返します★（2026-09-18・9枠目 58行目で 実測）
         ★★マスに 落ちたら #CALC!★★
           ＝`=LAMBDA(x,x*2)` を そのまま 打つと 実Excel は ★#CALC!★（9枠目 35行目） */
    function ラムダか(v) { return !!(v && typeof v === 'object' && v.型 === 'ラムダ'); }

    /** ★ラムダを 呼ぶ★
          ★引数の 数が 合わなければ #VALUE!★（★実Excel と 同じ★・9枠目 55/56行目）
          ★★呼ぶ 時の 台は ★作られた 時の 台★の 上に 載せます★★
            ＝★呼んだ 場所の 名前は 見えません★（＝閉じ込め） */
    function ラムダを呼ぶ(L, 引値, 手, 通り道) {
      if (!ラムダか(L)) return 計.誤('#VALUE!');
      if (!Array.isArray(引値) || 引値.length !== L.名たち.length) return 計.誤('#VALUE!');
      var 表 = {};
      for (var i = 0; i < L.名たち.length; i++) 表[L.名たち[i]] = 引値[i];
      var 前 = 名の台;
      名の台 = { 表: 表, 親: L.台 || null };
      try {
        return 値に(歩く(L.体, 手, 通り道));
      } finally {
        名の台 = 前;
      }
    }

    /** ★名の台から 引く★ ... 無ければ `undefined`（★null は 値として 有り得ます★） */
    function 名を引く(名) {
      for (var t = 名の台; t; t = t.親) {
        if (Object.prototype.hasOwnProperty.call(t.表, 名)) return t.表[名];
      }
      return undefined;
    }

    /* ★★字の 日付・時刻を 数に する★★（2026-09-15 に 足した）
       `lib/shiki-keisan.js` は ★口だけ 開けて 待って いました★
         「★日付の 字（"2026/1/1"→46023）は この 台では 読みません★」（139行）
       ★実測★（golden-346-2026-09-08.tsv）
         `=ABS("12:30")`      … ★0.52083333333333337★（1日を 1 と する 割合）
         `=ABS("2024-01-15")` … ★45306★（1900年からの 通し番号）
         `=AVERAGE("12:30")` `=MIN(…)` `=PRODUCT(…)` … ★0.5208333333333333★
       ★読む 形は 測った 物だけ★＝それ以外は null（＝今までどおり #VALUE!）
         ★`2024-01-15 12:30` の ような 組み合わせは 測って いません★＝読みません
       ★通し番号の 出し方は `lib/formula-soto.js` の `日から数` と 同じ 形★です。
         ★2か所に 在ります★＝`tests/shiki-hyou-hidzuke.test.mjs` が
         ★両方が 同じ 数を 返す事★を 見張ります（ずれたら 赤）。 */
    function 日付の通し番号(y, m, d) {
      var n = Math.round((Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000);
      /* ★★1900年の うるう日の 分を ★逆に★ 足して いました★★（2026-09-16 実測）
           `=AVERAGE("1900-01-01")` … 実Excel ★1★ ／ うち ★3★
           `=DATE(1900,2,28)`       … 実Excel ★59★ ／ うち ★61★
         ＝1899-12-30 から 数えた 日数に ★+1★ して いた（正しくは ★-1★）
           通し 1～59 は ★1899-12-31 起点★／61～ は ★1899-12-30 起点★
         ★正しい 形は `lib/shiki-kansuu.js` の `日から通し` に 在りました★
         ⇒★★同じ 数え方が ★ 3か所 ★に 在り、見張りは 2か所しか 見て いなかった★★
           （shiki-hyou.js ／ formula-soto.js ／ shiki-kansuu.js） */
      return (n < 61) ? n - 1 : n;
    }
    var 既定の手 = {
      字を日付に: function (字) {
        var t = String(字).trim();
        /* ★★分・秒が 1桁でも 読みます★★（2026-09-16 実測）
             `=AVERAGE("12:3")` … 実Excel ★0.502083333333333★ ／ うち ★#VALUE!★
           ★時が 24 を 越えたら 日に 繰り上がる★（"24:00"→1 ／ "25:00"→1.0416…）
             ＝★TIMEVALUE とは 別の 道★（あちらは 日を 切り捨てる） */
        var 時 = /^(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?$/.exec(t);
        if (時) {
          var h = +時[1], mi = +時[2], se = 時[3] ? +時[3] : 0;
          if (mi > 59 || se > 59) return null;
          return (h * 3600 + mi * 60 + se) / 86400;
        }
        var 日 = /^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})$/.exec(t);
        if (日) {
          var y = +日[1], mo = +日[2], da = +日[3];
          if (mo < 1 || mo > 12 || da < 1 || da > 31) return null;
          /* ★★1900年より 前は 日に なりません★★（2026-09-16 実測）
               `=AVERAGE("0000-01-01")` … 実Excel ★#VALUE!★ ／ うち ★3★
               ＝JS の `Date.UTC(0,…)` は ★ 0〜99 を 1900年代と 見ます★
               ⇒★年を 先に 弾く★（通し番号の 世界は 1900年から） */
          if (y < 1900 || y > 9999) return null;
          /* ★★1900-02-29 は 実Excelに 在る 日★★（本当の 暦には 無い）
               `=AVERAGE("1900-02-29")` … 実Excel ★60★ ／ うち ★#VALUE!★ */
          if (y === 1900 && mo === 2 && da === 29) return 60;
          var v = new Date(Date.UTC(y, mo - 1, da));
          if (v.getUTCMonth() !== mo - 1 || v.getUTCDate() !== da) return null;  /* 2月30日 等 */
          return 日付の通し番号(y, mo, da);
        }
        return null;                         /* ★測って いない 形は 読まない★ */
      },
    };

    var 覆い = {};
    /* ★元ごとに「どこを 覆ったか」も 覚える★
       ＝前は 消す たびに ★覆い 全部を なめて★ いました。
         ★実測（2026-09-14）★ 溢れる 式を N 個 打つ … N=1000 ★198.5ms★／N=2000 ★911.3ms★
           ＝★N の 2乗より 悪い★（2倍の N で 4.6倍）
         ⇒ 元ごとの 一覧を 持てば ★覆った 数だけ★で 消せる */
    var 覆った先 = {};

    function 覆いを消す(元) {
      var 先 = 覆った先[元];
      if (!先) return;
      端を変える();
      for (var i = 0; i < 先.length; i++) {
        if (覆い[先[i]] && 覆い[先[i]].元 === 元) {
          if (残り) {   /* ★④の 残りの 間は 消した 覆いを 控える★（叩き直し3 止め1） */
            if (!残り.前の覆い) { 残り.前の覆い = {}; 残り.前の所 = {}; }
            残り.前の覆い[先[i]] = 覆い[先[i]];
            var ば前 = 名から番地(先[i]);
            if (ば前) (残り.前の所[ば前.行] || (残り.前の所[ば前.行] = {}))[ば前.列] = true;   /* 行 → 列 の 表 */
          }
          delete 覆い[先[i]];
          if (まとめの覆い) まとめの覆い.push(先[i]);   /* ★溢れが 消えた マスも 見る 式を 直す★（2026-10-05） */
          索引から引く(先[i]);   /* ★中身に 残って いれば 引きません★ */
        }
      }
      delete 覆った先[元];
    }

    function 計算する(名, 通り道) {
      var m = 中身[名];
      if (!m || !m.形) return;
      数え.計算する++;
      if (残り && 数え.揃った名で - 残り.読みの始め > 残りの予算) throw new RangeError('台：輪の 塊の 読みが 予算（' + 残りの予算 + '）を 越えた');
      覆いを消す(名);
      /* ★今 どの マスを 計算して いるかを 皮に 教える★（2026-09-15）
         ＝`=CELL("row")`（引数 1つ）は ★今 いる マス★の 事（実測 … 1）。
         ★入れ子で 別の マスを 読む 事が 在る★（読む → 計算する）ので
         ★前の 物を 覚えて おいて 必ず 戻します★。 */
      var 前のマス = 今のマス;
      今のマス = 名から番地(名);
      /* ★★別の マスを 計算する 時は 名の台を 空に します★★（2026-09-18）
           ＝でないと A1 の `=LET(x,1,...)` の `x` が ★A2 の 式からも 見えます★
           ＝★他人の 名前★＝★一番 見つけにくい 食い違い★に なります
           ★前の 物を 覚えて 必ず 戻します★（`今のマス` `今の板` と 同じ 手） */
      var 前の名の台 = 名の台;
      名の台 = null;
      /* ★★裸の `A1` は 「この 式が 住んで いる 板」の A1★★（2026-09-15・㍸）
         ★ここを 入れないと 何が 起きるか（実物で 見ました）★
           板を またぐ 式が 入ると ★別の 板を 押して いる 最中に
           この 式が 直されます★。その 時 「今の 板」は ★押して いる 板★なので
           ★裸の `A1` が 全く 別の 板の A1 に なります★。
           ★実測（2026-09-15・司さんの 実物）★
             `=INDEX(他板!…, MATCH(A9, 他板!…, 0))` が ★#N/A★（3,462本）
             ＝★MATCH の 探す 値が 別の 板の 空の マスに なって いた★
         ★入れ子で 別の 板の 式を 読む 事が 在る★ので
           ★前の 物を 覚えて おいて 必ず 戻します★（`今のマス` と 同じ 手） */
      var 前の板 = 今の板;
      if (今のマス && 今のマス.板) 今の板 = 今のマス.板;
      var v;
      try {
        /* ★マスに 入れる 前に 参照を 1つの 値に する★（★暗黙の 交わり★）
           ＝`=INDEX(A1:B5,0,2)` は 打った マスの 行と 交わる（実測 J1→2・J3→6・J9→#VALUE!） */
        v = 値に(歩く(m.形, m.手 || 既定の手, 通り道 || [名]));
        /* ★★式の 答えが「空」なら 0★★（2026-08-29 に 実物 2,918本で 決まって いた 事）
           ＝`tests/empty-ref-zero.test.mjs` … `=C2`（C2 は 空）→ ★0★
             ★`=IF(C2="","",1)` の "" は 空の まま★（字なので ここに 来ません）
           ★2026-09-15 の 実測でも 同じ★
             `=IFERROR(A9,"x")` → ★0★ ／ `=IF(TRUE,,2)` → ★0★ ／ `=INDEX(E1:E5,4)` → ★0★
           ⇒★土台は ここを 入れて いませんでした★（借り物の 側だけ 入って いた＝道が 2本） */
        /* ★★ラムダは マスに 置けません★★（2026-09-18・9枠目 35行目で 実測）
             `=LAMBDA(x,x*2)` を そのまま 打つと 実Excel は ★#CALC!★ */
        if (ラムダか(v)) v = 計.誤('#CALC!');
        /* ★省略の 印が そのまま マスに 出ないように★（★空と 同じ 扱い★） */
        if (v && v.型 === '省略') v = 計.空;
        if (v && v.型 === '空') v = 計.数(0);
        /* ★★消え残りを 消すのは ここ★★＝★式の 一番 最後の 演算が ＋－の 時だけ★
             （前は ＋－の その場で 消して いて ★関数の 中まで 漏れて★ いました） */
        v = 計.足し引きを締める(v);
      } finally {
        今のマス = 前のマス;
        今の板 = 前の板;
        名の台 = 前の名の台;
      }
      /* ★溢れ（土台⑤）★＝広がる 先が 空いて いるかを 見る
         ★実測★ 溢れる先に 物が 在れば ★#SPILL!★（`=""` の ★見た目は 空★ でも 駄目） */
      if (溢.溢れか(v)) {
        var 元 = 名から番地(名);
        if (!元) { m.値 = 計.誤('#REF!'); return; }
        var 判 = 溢.溢れられるか(中身, 名, v, function (i, j) {
          return 番地から名(元.行 + i, 元.列 + j, 元.板);   /* ★溢れも 同じ 板★ */
        });
        if (!判.よい) { 塞がれた[名] = { 元: 元, 行数: v.行数, 列数: v.列数 }; m.値 = 計.誤('#SPILL!'); return; }
        delete 塞がれた[名];
        var 先一覧 = [];
        for (var i = 0; i < v.行数; i++) {
          for (var j = 0; j < v.列数; j++) {
            if (i === 0 && j === 0) continue;
            var なま = 番地から名(元.行 + i, 元.列 + j, 元.板);
            覆い[なま] = { 元: 名, 行: i, 列: j };
            if (まとめの覆い) まとめの覆い.push(なま);
            索引に足す(なま);   /* ★溢れで 埋まった マスも 「在る」★ */
            先一覧.push(なま);
          }
        }
        覆った先[名] = 先一覧;
      }
      if (!溢.溢れか(v)) delete 塞がれた[名];
      m.値 = v;
    }

    /* ★★このマスに 打つと 溢れが 変わる 式★★（2026-10-05）
         ① このマスを 溢れで 埋めて いる 式（溢れた 後に 打つと 実Excel は #スピル!＝組 M）
         ② このマスが 溢れる はずの 範囲に 在って 塞がれて いる 式（塞いだ 物を 消すと 戻る＝組 J） */
    function 溢れが変わる式(n) {
      var 出 = [];
      if (覆い[n] && 覆い[n].元 !== n) 出.push(覆い[n].元);
      var ば = 名から番地(n);
      if (ば) {
        for (var f in 塞がれた) {
          if (f === n) continue;
          var 塞 = 塞がれた[f], 元 = 塞.元;
          if (String(元.板 || '') !== String(ば.板 || '')) continue;
          if (ば.行 >= 元.行 && ば.行 < 元.行 + 塞.行数 && ば.列 >= 元.列 && ば.列 < 元.列 + 塞.列数) 出.push(f);
        }
      }
      return 出;
    }

    /* ══ ★外に 出す 物★ ══ */
    return {
      /* ★マスに 打つ★（`=` で 始まれば 式） */
      /* ══ ★★まとめて 打つ★★ ══（2026-10-05・1打ち目の 固まり）
           ★何が 在ったか★ 最初の 1打ちで `ShikiItaAwase.合わせる` が 本の 全部の マス（約35,000）を 1つずつ 打ち、
             打つ たびに そのマスを 見る 式を 直して いた ⇒ 式 15,800本を ★58万回★ 計算（1本 平均 約37回）
           ★直し★ 打つ 間は 名を 覚えるだけ／`まとめて打つ終わり` で 変わった 物と それを 見る 式を 頼りの 順に 1回ずつ
           ★溢れ★ その 間に 埋まった マスは、それを 見る 式を もう1回 直す（溢れの 線は 頼りの 線に 無い 為） */
      まとめて打つ始め: function () { 壊れを見る(); if (!まとめ中) { まとめ中 = {}; まとめの覆い = []; } },
      まとめて打つ終わり: function () {
        壊れを見る();
        if (!まとめ中) return 0;
        var 名ら = Object.keys(まとめ中);
        まとめ中 = null;
        try { 直す群(名ら, true); var 覆 = まとめの覆い; まとめの覆い = null; if (覆 && 覆.length) 直す群(覆, false); }
        finally { まとめの覆い = null; }
        return 名ら.length;
      },
      打つ: function (名, 字) {
        壊れを見る();
        var n = 揃える(名);
        板を覚える(n);                 /* ★打った 板は 知って いる 板★ */
        /* ★★頼りを 拾う 間も「この 式が 住む 板」★★（2026-09-15・㍸）
           ★入れないと 何が 起きるか★
             `板2!B1` に `=A1*10` を 打つ と、頼りの 線だけ ★板1!A1★ に 付く。
             ⇒★答えは 合って いるのに 板2!A1 を 打ち直しても 直らない★
             ＝★㆝（直に 書き換える）と 同じ 型の ★黙った 嘘★
           ★実物では 漏れて いました★（★材料を 打ち直さない★から）
             ⇒★見張り（⓪）で 見つけました★ */
        var 前の板2 = 今の板;
        var ばn = 名から番地(n);
        if (ばn && ばn.板) 今の板 = ばn.板;
        try {
        var s = String(字 == null ? '' : 字);
        var 起こす = 溢れが変わる式(n);   /* ★打つ 前に 見る★（打つと 覆いが 変わる 為・2026-10-05） */
        頡りを外す(n);          /* ★前に 見ていた 先から 自分を 外す★ */
        /* ★★索引に 足すのは ここ 1か所で 足ります★★
             ＝この 下の どの 道を 通っても ★必ず 中身[n] に 入ります★（4通り とも）
             ＝★だから 枝ごとに 書きません（書き忘れる 道を 作らない）★ */
        索引に足す(n);

        if (s.charAt(0) === '=') {
          var か = 切.切る(s.slice(1));
          if (!か.ok) { 中身[n] = { 打った字: s, 形: null, 値: 計.誤('#VALUE!') }; return; }
          var y = 形.形にする(か.出);
          if (!y.ok) { 中身[n] = { 打った字: s, 形: null, 値: 計.誤('#VALUE!') }; return; }
          中身[n] = { 打った字: s, 形: y.形, 値: undefined };
          var 拾った = 頼りを拾う(y.形, []), 先 = [], 帯鍵 = [];
          for (var j0 = 0; j0 < 拾った.length; j0++) (拾った[j0].indexOf('帯|') === 0 ? 帯鍵 : 先).push(拾った[j0]);
          見ている[n] = 先;
          if (揮発を含む(y.形)) 揮発な式[n] = true;
          for (var j = 0; j < 先.length; j++) {
            if (!見られている[先[j]]) 見られている[先[j]] = {};
            見られている[先[j]][n] = true;
          }
          /* ★列まるごと・行まるごとは 帯の 見張りへ★（2026-11-04・案X） */
          見ている帯[n] = 帯鍵;
          for (var j2 = 0; j2 < 帯鍵.length; j2++) {
            if (!帯の見張り[帯鍵[j2]]) 帯の見張り[帯鍵[j2]] = {};
            if (!帯の見張り[帯鍵[j2]][n]) { 帯の見張り[帯鍵[j2]][n] = true; 帯の見張りの数++; }
          }
        } else if (s === '') {
          中身[n] = { 打った字: '', 形: null, 値: 計.空 };
        } else {
          var d = 計.字を数に(s, {});
          中身[n] = { 打った字: s, 形: null,
            値: (d === null) ? 計.字(s) : 計.数(計.打った数(String(d))) };
        }
        /* ★打った マスの 覆いを 消す★（前の 溢れが 残らない ように・土台⑤） */
        覆いを消す(n);
        /* ★溢れる 式は 打った その場で 計算する★（2026-09-14／土台⑤）
           ＝溢れが どこまで 広がるかは ★計算しないと 分からない★。
             後回しに すると `=E2+E3` の ような 式が ★まだ 覆いが 無い★ので 0 に なる
             （実測 2026-09-14＝`=A1:A3` の 後に `=E2+E3` を 打つと ★0★ に なって いた）。 */
        /* ★このマスを 見ている 式だけ★ 計算し直す */
        /* ★まとめて 打つ 間は 直さない★（終わりに `直す群` で 1回だけ・2026-10-05） */
        if (まとめ中) { まとめ中[n] = true; for (var 起i = 0; 起i < 起こす.length; 起i++) まとめ中[起こす[起i]] = true; return; }
        /* ══ ★★1つずつ 打つ 時も 溢れで 埋まった マスを 見る 式を 直す★★ ══（2026-10-05）
             ★前★ F1 `=E2+E3` を 先に 打ち、後から E1 `=A1:A3` を 打つと、E2・E3 は 溢れで 埋まるが
               ★F1 は 直されず 0 の まま★（実Excel は 5・経営者の 物差し 137 の 組 F）
             ★直し★ この 打つの 間に 出来た 覆いを 覚え、終わりに それを 見る 式を `直す群` で 直す */
        var 覆前 = まとめの覆い;
        まとめの覆い = [];
        try {
        直す(n);
        /* ★順番が 大事★＝★`直す` の 後★に 読む。
             先に 計算すると ★輪（A1→A2→A1）が 見つからなく なります★
             （2026-09-14 実測＝先に 計算したら `#REF!` では なく `3` に なった）。
           `読む` は ★まだ 計算して いない時だけ★ 計算します。 */
        if (中身[n] && 中身[n].形) 読む(n, []);
          /* ★溢れが 変わる 式を 計算し直す★（その 先を 見る 式も 一緒に） */
          if (起こす.length) {
            for (var 起j = 0; 起j < 起こす.length; 起j++) { var 起m = 中身[起こす[起j]]; if (起m && 起m.形) 起m.値 = undefined; }
            直す群(起こす, true);
          }
          var 出来た覆い = まとめの覆い;
          まとめの覆い = 覆前;
          if (出来た覆い && 出来た覆い.length) 直す群(出来た覆い, false);
        } finally { まとめの覆い = 覆前; }
        } finally { 今の板 = 前の板2; }
      },

      /* ★★マスに 「型の まま」 置く★★（2026-09-15・土台⑲）
         ★なぜ 要るか★ … `打つ` は ★字しか 受けません★。
           `打つ('A2','2')` は ★数の 2★ に なるので、
           ★字の "2"★ や ★真偽★ を 置く 口が ありませんでした。
         ★それで 何が 起きて いたか★ … 呼ぶ側が
           ★`表.中身[マス].値` を 直に 書き換えて いました★。
           ★それでは その マスを 見て いる 式に 報せが 行きません★
           ＝★★先に 式を 打つ と その 式は 古い 答えの まま★★。
           （今の 測り道具は「材料 先・式 後」なので 当たって いません。
             ★順を 変えた 途端に 黙って 嘘に なる★ので 口を 作りました。）
         ★受ける 形★
           ・`{ 型:'字'|'数'|'真偽'|'誤'|'空', 値 }` … ★台の 値そのもの★
           ・数        … 数
           ・真偽      … 真偽
           ・字        … ★必ず 字★（★ここが `打つ` との 違い★）
           ・ null/無し … 空
         ★`打った字`★ は 3つ目で 渡せます（省くと 値から 作ります） */
      置く: function (名, 値, 打った字) {
        壊れを見る();
        var n = 揃える(名);
        板を覚える(n);
        var v = 値の形に(値);
        頡りを外す(n);
        /* ══ ★★索引に 足すのを ここでも★★ ══（2026-09-26 ★1回 落としました★）
             ＝中身に 入れる 道は ★`打つ` と ここ（`置く`）の 2本★
             ＝★`打つ` にだけ 足して ここを 落として いました★
             ⇒`tests/shiki-sagasu.test.mjs` が ★MATCH/INDEX が 全部 #N/A★ で 捕まえました
             ⇒記憶「★作る道が 2本 在る 時は 両方 直せ★」 */
        索引に足す(n);
        中身[n] = {
          打った字: (打った字 == null) ? 既定の字(v) : String(打った字),
          形: null,
          値: v
        };
        覆いを消す(n);
        直す(n);                  /* ★★この 1行が 無いのが ⑲の 欠陥でした★★ */
      },

      /* ★読む★（値の 形の まま） */
      /* ★マスが 持って いる 物を そのまま★ 返す（溢れなら 溢れの まま）
         ＝`読む()` は ★1マスとして 読む時に 左上★を 返すので、
           外から 広がりを 見たい 時に 使えません（2026-09-14 実測）。 */
      値: function (名) {
        壊れを見る();
        var n = 揃える(名);
        var m = 中身[n];
        if (!m) { 読む(n, []); m = 中身[n]; }
        if (m && m.形 && m.値 === undefined) 読む(n, []);
        if (m && m.値 !== undefined) return m.値;
        return 読む(n, []);
      },

      /* ★画面に 出す 字★ */
      字: function (名) {
        壊れを見る();
        var v = 読む(名, []);
        if (v.型 === '誤') return v.値;
        if (v.型 === '真偽') return v.値 ? 'TRUE' : 'FALSE';
        if (v.型 === '数') return 計.数を字に(v.値);
        if (v.型 === '空') return '';
        return v.値;
      },

      /* ★誰が 誰を 見ているか（測る為に 外へ 出す）★ */
      見ている: function (名) { return (見ている[揃える(名)] || []).slice(); },
      /* ★外に 出す 形は 今までどおり 並び★（中の 持ち方だけ 変えた） */
      見られている: function (名) { return Object.keys(見られている[揃える(名)] || {}); },

      /* ★★板の 口★★（2026-09-15・㍸）
         `板('名')`      … ★今の 板を 変える★（引数なしで 今の 板を 返す）
         `板を足す('名')` … ★中身が 空の 板も 「在る」に する★
                          （★無い 板は #REF!／空の 板は 0★＝別物）
         `板一覧()`    … ★知って いる 板★ */
      板: function (名) {
        if (名 === undefined) return 今の板;
        今の板 = String(名).toUpperCase();
        板たち[今の板] = true;
        return 今の板;
      },
      板を足す: function (名) { 板たち[String(名).toUpperCase()] = true; },

      /* ★その マスに 何と 打って 在るか★（2026-09-15・㍸）
         ★`中身` を 直に 引かせない 為★＝★鍵は `板名!A1` に なりました★ので
           `表.中身['A1']` は ★黙って undefined★ に なります（★一番 怖い 壊れ方★）。 */
      壊れたか: function () { return 壊れた; },
      /* ★輪の 塊の 読みの 予算を 置く★（試験と 測りの 為・既定 300万） */
      読みの予算を置く: function (n) { var 前 = 残りの予算; 残りの予算 = +n; return 前; },
      打った字: function (名) {
        var m = 中身[揃える(名)];
        return m ? String(m.打った字 || '') : '';
      },
      板一覧: function () { return Object.keys(板たち); },

      /* ★数えた 回数（★変わった所だけ 直したか★を 測る為）★ */
      直した回数: 0,
      /* ★★索引の 口★★（2026-09-26・経営者1 の 注文 ⑷） */
      索引を照らす: 索引を照らす,
      空と答えた数: function () { return 空と答えた; },
      空と答えたを消す: function () { 空と答えた = 0; },
      索引の数: function () { return 索引を照らす().索引の数; },
      /* ★★覆い（溢れで 埋まった マス）の 数★★（2026-09-26・経営者1 の 注文 ④）
           ＝★`覆い` は `表()` の 中の 閉じた 物＝外から 引けません★
           ⇒★口を 出さないと 相手が 別の 道で 数える 事に なります★
           ⇒★それは 「どの 段で 見たか」が 違う 数に なります★（09-25 の 159個 と 1個 の 型） */
      覆いの数: function () { return Object.keys(覆い).length; },
      /* ★★入口ごとの 回数の 口★★（2026-10-05）＝`数え` の 写しを 返す／四角の 種類は 始めた 時だけ */
      数えを出す: function () {
        var o = {}; for (var k in 数え) o[k] = 数え[k];
        if (四角の種類) {
          var 鍵 = Object.keys(四角の種類), 二回以上 = 0, 余分 = 0;
          for (var i = 0; i < 鍵.length; i++) { var 回 = 四角の種類[鍵[i]]; if (回 > 1) { 二回以上++; 余分 += 回 - 1; } }
          o.四角の種類 = 鍵.length; o.二回以上読んだ四角 = 二回以上; o.読み直しの回数 = 余分;
        }
        return o;
      },
      数えを始める: function (種類も) { for (var k in 数え) 数え[k] = 0; 四角の種類 = 種類も ? {} : null; },
      中身: 中身
    };

    /* ★前に 見ていた 先から 自分を 外す★（`打つ` と `置く` の 両方が 使う） */
    function 頡りを外す(n) {
      var 前 = 見ている[n] || [];
      for (var i = 0; i < 前.length; i++) {
        var 者 = 見られている[前[i]];
        if (者) delete 者[n];
      }
      見ている[n] = [];
      var 前帯 = 見ている帯[n] || [];
      for (var k = 0; k < 前帯.length; k++) {
        var 帯者 = 帯の見張り[前帯[k]];
        if (帯者 && 帯者[n]) { delete 帯者[n]; 帯の見張りの数--; }
      }
      delete 見ている帯[n];
      delete 揮発な式[n];
    }
    /* ★★この マスを 見る 式 全部★★（2026-11-04・案X）＝見られている ＋ この マスの 列・行・板の 帯の 見張り
         ★直す群 は ここ だけを 使う★（帯の 式が 古い 答えで 残らない 為・設計前 taiketsu S1） */
    function 見る者たち(x, 帯を足すか) {
      var 直 = 見られている[x];
      if (帯の見張りの数 === 0) return 直 ? Object.keys(直) : [];
      数え.帯を見た++;
      var 出 = {}, k, 帯鍵 = マスの帯の鍵たち(x);
      if (直) for (k in 直) if (Object.prototype.hasOwnProperty.call(直, k)) 出[k] = true;
      for (var i = 0; i < 帯鍵.length; i++) {
        if (帯を足すか && !帯を足すか(帯鍵[i])) continue;
        var 帯者 = 帯の見張り[帯鍵[i]];
        if (帯者) for (k in 帯者) if (Object.prototype.hasOwnProperty.call(帯者, k)) 出[k] = true;
      }
      return Object.keys(出);
    }

    /* ★外から 渡された 物を 台の 値の 形に 直す★
       ★分からない 物は 黙って 通さない★（★黙った 変換が 一番 怖い★） */
    function 値の形に(x) {
      if (x === null || x === undefined) return 計.空;
      if (typeof x === 'number') return 計.数(x);
      if (typeof x === 'boolean') return { 型: '真偽', 値: x };
      if (typeof x === 'string') return 計.字(x);
      if (typeof x === 'object' && x.型) {
        if (x.型 === '数')   return 計.数(Number(x.値));
        if (x.型 === '字')   return 計.字(String(x.値));
        if (x.型 === '真偽') return { 型: '真偽', 値: !!x.値 };
        if (x.型 === '誤')   return 計.誤(String(x.値));
        if (x.型 === '空')   return 計.空;
      }
      throw new Error('置く：分からない 形 … ' + JSON.stringify(x));
    }

    /* ★`打った字` を 省いた 時の 既定★ */
    function 既定の字(v) {
      if (v.型 === '空')   return '';
      if (v.型 === '真偽') return v.値 ? 'TRUE' : 'FALSE';
      if (v.型 === '数')   return 計.数を字に(v.値);
      return String(v.値);
    }

    /* ★変わった マスを 見ている 式を たどって 計算し直す★
       ★全部 計算し直さない★＝ここが 本体の 肝 */
    function 直す(変わった) { 直す群([変わった], false); }

    /* ★★まとめて 直す★★（2026-10-05）＝`まとめて打つ` の 終わりに 1回だけ 呼ぶ
         ★変わったたち★ を 見ている 式 全部（★自分も★ が 真なら 変わった 式 自身も）を 頼りの 順に 1回ずつ */
    function 直す群(変わったたち, 自分も) {
      try { 直す群の中(変わったたち, 自分も); }
      catch (e) { 壊れた = String((e && e.message) || e); throw e; }
    }
    function 直す群の中(変わったたち, 自分も) {
      /* ══ ★★頼りの 順に 直す★★ ══（2026-09-15 に 直した）
         ★★前は「先に 着いた 順」で 直して いました★★＝★菱形で 古い 答えが 残ります★
           K1=10 ／ Q1=`=K1*2` ／ M1=`=K1-Q1`（★M も Q も K を 見る／M は Q も 見る★）
           ⇒ K が 変わると ★M と Q の 両方★に 報せが 行く
           ⇒ ★M を 先に 直すと 古い Q を 使う★／その後 Q を 直しても
             ★M は「済」印が 付いて いて もう 直らない★
           ★実測（2026-09-15・司さんの実物）★
             `=IFERROR((K14-Q14)/L14,…)` が ★古い 答えの まま★
             ⇒ そこから V／X／Z／AB → AN14 → SUBTOTAL と ★鎖で 広がって いた★
             ⇒ 押す順で 合わない マスが 変わる（順 37本／逆 176本／共通 6本）
         ★直し★＝★自分が 見て いる 物を 全部 直してから 自分を 直す★（頼りの 順）
           ・★数え上げで 回します★（O(マス＋線)）＝★2乗に しません★
             （`tests/shiki-hyou-omosa.test.mjs` が 2乗を 赤に します）
           ・★輪は 順番では 解けません★⇒★残った 物は そのまま 直す★
             （輪の 見つけ方は `計算する` の 通り道が 持って います） */
      /* ①★届く 先を 全部 集める★ */
      var 要る = {}, 積 = [], i;
      /* ★同じ 帯の 見張りは ★最後に 出る 所★だけ 積む★（2026-11-07・本番前 taiketsu 叩き直し 止め1）
           ＝4000マスが 同じ A:A を 見る 2000本を 4000回 積んで いた（開く 1.8秒）。積は 上から 取る＝前に 積んだ 同じ 名は
             後の 所が 先に 取られて 要る に 入り 飛ばされる＝★取る 順は 前と 同じ★（最初の 所だけ 残すと 名ら の 順が 変わり
             本当の 輪の 答えが 変わった＝叩き直し 任意1） */
      var 帯の最後 = {}, 帯の今 = 0;
      if (帯の見張りの数 > 0) for (i = 0; i < 変わったたち.length; i++) { var 鍵ら0 = マスの帯の鍵たち(変わったたち[i]); for (var b0 = 0; b0 < 鍵ら0.length; b0++) 帯の最後[鍵ら0[b0]] = i; }
      var 最後だけ = function (鍵) { return 帯の最後[鍵] === 帯の今; };
      for (i = 0; i < 変わったたち.length; i++) {
        帯の今 = i;
        var 見る者 = 見る者たち(変わったたち[i], 最後だけ);
        for (var q = 0; q < 見る者.length; q++) 積.push(見る者[q]);
        if (自分も && 中身[変わったたち[i]] && 中身[変わったたち[i]].形) 積.push(変わったたち[i]);
      }
      /* ★揮発な 式は 何が 変わっても 直す★（2026-11-07）＝変わった マス 自身が 揮発な 式の 時も 入れる */
      var 揮発あり = false;
      if (変わったたち.length) for (var vn in 揮発な式) if (Object.prototype.hasOwnProperty.call(揮発な式, vn) && 中身[vn] && 中身[vn].形) { 積.push(vn); 揮発あり = true; }
      while (積.length) {
        var n = 積.pop();
        if (要る[n]) continue;
        要る[n] = true;
        var 次 = 見る者たち(n);
        for (i = 0; i < 次.length; i++) if (!要る[次[i]]) 積.push(次[i]);
      }
      var 名ら = Object.keys(要る);
      if (!名ら.length) return;
      /* ②★自分より 先に 直す物が いくつ 在るか 数える★ */
      /* ★待ち数は ③で 減らす 辺（見る者たち）と ★同じ 辺★で 数える★（2026-11-04・案X）
           ＝帯の 頼り（A:A を 見る 式）も 順番に 入る。前の 数え方（見ている の 並び）は 帯を 知らず、
             `=A1+A1` の 様な 重なりでも ③と 数が 合わなかった（④の 落ち穂に 回って いた） */
      var 待ち数 = {}, 順番 = [], 辺 = {};
      for (i = 0; i < 名ら.length; i++) 待ち数[名ら[i]] = 0;
      for (i = 0; i < 名ら.length; i++) {
        var 者ら = 見る者たち(名ら[i]);
        辺[名ら[i]] = 者ら;
        for (var k = 0; k < 者ら.length; k++) if (要る[者ら[k]] && 者ら[k] !== 名ら[i]) 待ち数[者ら[k]]++;
      }
      for (i = 0; i < 名ら.length; i++) if (待ち数[名ら[i]] === 0) 順番.push(名ら[i]);
      /* ★揮発が 混じる 時は 直す 物 全部を 先に「まだ」に する★（2026-11-07）
           ＝揮発な 式が 読む マスは 線に 無い＝順番では 守れない⇒★読まれた 時に その場で 計算★（読む の 値 undefined の 道） */
      if (揮発あり) for (i = 0; i < 名ら.length; i++) { var mv = 中身[名ら[i]]; if (mv && mv.形) mv.値 = undefined; }
      /* ③★0 に なった 物から 直す★ */
      var 直した = 0;
      while (順番.length) {
        var y = 順番.pop();
        直した++;
        var m = 中身[y];
        if (m && m.形) { m.値 = undefined; 計算する(y, [y]); }
        var 見て = 辺[y] || 見る者たち(y);
        for (i = 0; i < 見て.length; i++) {
          var w = 見て[i];
          if (!要る[w] || w === y) continue;
          if (--待ち数[w] === 0) 順番.push(w);
        }
      }
      /* ④★輪で 残った 物★＝順番では 解けない ⇒ そのまま 直す */
      /* ★回を 重ねて 動かなく なるまで★＝下の `塊を直す` に 書いて あります（2026-11-07・案C） */
      if (直した < 名ら.length) {
        var 残り名 = [];
        for (i = 0; i < 名ら.length; i++) {
          var z = 名ら[i];
          if (待ち数[z] === 0) continue;         /* もう 直した */
          待ち数[z] = 0;
          var m2 = 中身[z];
          if (m2 && m2.形) 残り名.push(z);
        }
        /* ★輪の 塊（強く 繋がった 組）に 分けて 上流から★（2026-11-07・taiketsu 叩き直し 止め1）
             ＝前は 残り 全部を 1つの 塊と して 回した。偽の 輪の 組を つなぐと 1回に 1組しか 進まず、
               上限の 回で 届かずに #REF!（組 15個から）。下流の COUNT・IFERROR まで #REF! を 直に 入れて いた
             ＝塊ごとに 回せば 上の 塊は もう 決まって いる＝下の 塊は 1〜2回で 決まる。下流の 式は 自分の 塊で 式どおりに 計算
             ＝塊の 分け方は 辺（見る者たち）の 強連結（Tarjan・積で 回す＝深い 鎖で スタックが 溢れない・O(マス＋線)） */
        var 塊ら = 強く繋がった組(残り名, 辺);
        for (var 塊i = 塊ら.length - 1; 塊i >= 0; 塊i--) 塊を直す(塊ら[塊i]);
      }
    }

    /* ★強連結の 組（Tarjan を 積で）★＝返す 順は ★下流が 先★（呼ぶ 側で 逆に 回す） */
    function 強く繋がった組(名ら, 辺) {
      var 在る = {}, 番 = {}, 低 = {}, 積に = {}, 積 = [], 組ら = [], 次の番 = 0, i;
      for (i = 0; i < 名ら.length; i++) 在る[名ら[i]] = true;
      for (i = 0; i < 名ら.length; i++) {
        var 根 = 名ら[i];
        if (番[根] !== undefined) continue;
        var 呼び = [[根, 0]];
        番[根] = 低[根] = 次の番++; 積.push(根); 積に[根] = true;
        while (呼び.length) {
          var 上 = 呼び[呼び.length - 1], v = 上[0], 先 = 辺[v] || [];
          if (上[1] < 先.length) {
            var w = 先[上[1]++];
            if (!在る[w]) continue;
            if (番[w] === undefined) { 番[w] = 低[w] = 次の番++; 積.push(w); 積に[w] = true; 呼び.push([w, 0]); }
            else if (積に[w] && 番[w] < 低[v]) 低[v] = 番[w];
            continue;
          }
          呼び.pop();
          if (呼び.length) { var p = 呼び[呼び.length - 1][0]; if (低[v] < 低[p]) 低[p] = 低[v]; }
          if (低[v] === 番[v]) {
            var 組 = [], x;
            do { x = 積.pop(); 積に[x] = false; 組.push(x); } while (x !== v);
            組ら.push(組.length > 1 ? 塊の中の順(v, 組, 辺) : 組);
          }
        }
      }
      return 組ら;
    }

    /* ★塊の 中の 順★＝根から 辺を 深さ優先で たどった 帰りがけの 逆（輪で ない 所は 頼りの 順）
         ＝順を 守らないと 計算が 読まれた 先へ 深く 潜り、深さで 切った 所が 毒に なって 回が 増える
           （広く 頼る H3 の 下の 鎖 1000段を 打つと 1打ちごとに 鎖 全部が 1つの 塊＝名の 順では 作るのに 27秒） */
    function 塊の中の順(根, 組, 辺) {
      var 中 = {}, 見た = {}, 帰り = [], i;
      for (i = 0; i < 組.length; i++) 中[組[i]] = true;
      var 呼び = [[根, 0]]; 見た[根] = true;
      while (呼び.length) {
        var 上 = 呼び[呼び.length - 1], 先 = 辺[上[0]] || [];
        if (上[1] < 先.length) {
          var w = 先[上[1]++];
          if (中[w] && !見た[w]) { 見た[w] = true; 呼び.push([w, 0]); }
          continue;
        }
        帰り.push(呼び.pop()[0]);
      }
      for (i = 0; i < 組.length; i++) if (!見た[組[i]]) 帰り.push(組[i]);   /* 念の 為（強連結なら 全部 届く） */
      return 帰り.reverse();
    }

    /* ★1つの 塊を 回を 重ねて 動かなく なるまで★（2026-11-07・案C）＝前は 名の 順に 1回だけ 計算し、
         まだ 直して いない 物の ★古い 数★を 読んで いた（H3=`SUM(INDEX(A:A,5):INDEX(1:1,3))`・J1=`=H3*2`・K1=`同じ式+J1` で K1 4972・正 4926）。
         ・読まれた 塊の 物は その場で 先に 計算（頼りの 順）・深さ 残りの深さ 段まで（越えたら 1回目は 毒・2回目から 前の 回の 値）
         ・輪の 通り道に 当たったら 1回目は 毒・2回目から 前の 値（偽の 輪＝使わない マスなら 答えは 変わらない）
         ・前の 値を 使わずに 済んだ 回 か、値が 1つも 動かない 回で 終わり
         ・上限の 回までに 動き続ける 物は ★本当の 輪★＝#REF!（その後 下流の 塊が 式どおりに 読む）
         ★言えない 事★＝動きが 止まる 本当の 輪（A1=`=ISERROR(B1)`・B1=`=A1` 等）は 輪の 印に ならず 止まった 値を 出す（前の 版も 同じく 黙る） */
    function 塊を直す(塊) {
      var i;
      残り = { 在る: {}, 済: {}, 疑い: false, 回: 0, 読みの始め: 数え.揃った名で };
      for (i = 0; i < 塊.length; i++) 残り.在る[塊[i]] = true;
      var 上限 = Math.ceil(塊.length / 残りの深さ) * 2 + 10, 動いた = [];
      try {
        for (var 回 = 0; 回 < 上限; 回++) {
          var 前の鍵 = {};
          for (i = 0; i < 塊.length; i++) 前の鍵[塊[i]] = 値の鍵(中身[塊[i]].値);
          残り.済 = {}; 残り.疑い = false; 残り.回 = 回;
          for (i = 0; i < 塊.length; i++) {
            var z2 = 塊[i];
            if (残り.済[z2] || !中身[z2] || !中身[z2].形) continue;
            残り.済[z2] = true;
            計算する(z2, [z2]);
          }
          動いた = [];
          for (i = 0; i < 塊.length; i++) if (中身[塊[i]] && 値の鍵(中身[塊[i]].値) !== 前の鍵[塊[i]]) 動いた.push(塊[i]);
          if (!残り.疑い || (!動いた.length && 回 > 0)) { 動いた = []; break; }   /* ★1回目に 毒を 使った 回は 動かなくても もう1回★（叩き直し2 止め1） */
        }
      } finally { 残り = null; }
      for (i = 0; i < 動いた.length; i++) if (中身[動いた[i]] && 中身[動いた[i]].形) 中身[動いた[i]].値 = 計.誤('#REF!');
    }
  }

  /* ══ ★★先に 計算しない 関数（特別な 形）の 名簿★★ ══（2026-09-18）
       ★なぜ 表に 出すか★
         ＝`docs/measured/kami-ga-nai-kansuu.mjs` と `tests/tsunagu-mon.test.mjs` は
           ★台が 出せる 関数★を `shiki-kansuu.js` の `表` と `shiki-tsunagi.js` から 数えます。
         ⇒★★ここ（板）に 書いた 物は ★数から 落ちます★★
         ⇒★LET を 書いたのに「あと 35個」が 減りませんでした★（実測・2026-09-18）
         ⇒★★「書いたのに 数に 出ない」は 数の 側の 欠陥です★★
       ★だから 名簿を 出します★＝★数える 側が 足せる ように★ */
  var 特別な形 = ['LET', 'LAMBDA'];

  return {
    表: 表, 名から番地: 名から番地, 番地から名: 番地から名,
    特別な形: 特別な形,
    /* ★どちらが 答えたか★（経営者1 と 決めた 見張り①・2026-09-14）
       ＝★借り物の 出番が 減る様子を 数で 見る★為。
         `土台` … `lib/shiki-kansuu.js` の 7個
         `皮`   … `lib/shiki-tsunagi.js` の 61個
         `皮（溢れ待ち）` … 皮が 答えたが ★溢れなので まだ マスに 並べられない★（土台⑤ 待ち）
         `知らない` … どちらも 知らない ⇒ #NAME?（★今は 借り物が 出して いる 物★） */
    答えた数: function () { return JSON.parse(JSON.stringify(答えた)); },
    答えた数を消す: function () { for (var k in 答えた) if (Object.prototype.hasOwnProperty.call(答えた, k)) delete 答えた[k]; }
  };
});
