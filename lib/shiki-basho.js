/* shiki-basho.js — ★参照を 返す／参照を 見る★（2026-09-15）
 *
 *  ★★何の 台か★★
 *    INDEX は ★値では なく 参照（どのマスか）を 返します★。
 *    ★実測で 決まりました★（`docs/measured/kansuu46/golden-index-match-kimari-2026-09-15.tsv`）
 *      `=ISBLANK(INDEX(E1:E5,4))`          → ★TRUE★   （空の マスを 指して いる）
 *      `=CELL("address",INDEX(A1:B5,2,2))` → ★$B$2★
 *      `=ROW(INDEX(A1:B5,3,2))`            → ★3★
 *      `=SUM(INDEX(A1:A5,2):A5)`           → ★14★     （★答えに `:` が 使える★）
 *      `=SUM(INDEX(A1:B5,0,2))`            → ★30★     （★0＝まるごと★）
 *    ⇒★値だけ 返す 作りでは この 5本が 全部 外れます★。
 *
 *  ★★この 台を 1つに する 訳★★（指示役1 2026-09-15）
 *    ★AREAS・OFFSET・INDIRECT も 同じ 物を 使います★＝★後で 乗せるだけ★に する。
 *    ⇒ INDEX/MATCH の 為だけの 作りに しない。
 *
 *  ★★暗黙の 交わり★★（★実Excel の 決まり／INDEX の 決まりでは 無い★）
 *    ★四角を「1つの 値」に 詰める 時、実Excel は ★式の 行／列と 交わる 所★を 採ります★。
 *    ★実測（同じ 式を 別の マスに 打った）★
 *      `=INDEX(A1:B5,0,2)` … J1 → ★2★（B1） ／ J3 → ★6★（B3） ／ J9 → ★#VALUE!★（行9 は 外）
 *      `=INDEX(A1:B5,2,0)` … A8 → ★2★（A2） ／ B8 → ★4★（B2）
 *    ⇒★同じ 式でも 打つ マスで 答えが 変わります★＝★紙の 頭に 打った マスを 書く★決まりに した。
 *
 *  ★値の 形★ `shiki-keisan.js` と 同じ { 型, 値 }
 *  ★参照の 形★ { 種:'参照', 行, 列, 行数, 列数, 並び }
 *      ・`行`／`列` … ★左上の 番地★（1 から）
 *      ・`並び`     … ★平ら・行から★（`四角を読む` と 同じ 並べ方）
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(
    typeof require === 'function' ? require('./shiki-keisan.js') : root.ShikiKeisan);
  else root.ShikiBasho = factory(root.ShikiKeisan);
})(typeof self !== 'undefined' ? self : this, function (計) {
  'use strict';

  /* ══ ★参照を 作る★ ══ */
  function 参照(行, 列, 行数, 列数, 並び, 番地なし) {
    /* ★★番地は 0 から★★（この 台の 決め＝A1 は {行:0, 列:0}）
       ⇒★「番地が 無い」を 0 で 表しては いけません★（2026-09-15 に 踏んだ）
         ＝A 列・1 行目の 参照が 全部「番地なし」に なり ★18本 外しました★。
       ⇒★別の 印を 持ちます★ */
    return { 種: '参照', 行: 行, 列: 列, 行数: 行数, 列数: 列数, 並び: 並び, 番地なし: !!番地なし };
  }
  function 参照か(v) { return !!(v && typeof v === 'object' && v.種 === '参照'); }

  /* ══ ★暗黙の 交わり★＝四角を 1つの 値に 詰める ══
     ★1本しか 無い 向きは そのまま／2本以上 在る 向きは 式の 行（列）と 交わる★
     ★交わらなければ #VALUE!★（実測 `=INDEX(A1:B5,0,2)` を J9 に 打つと #VALUE!） */
  function 交わる(参, 今のマス) {
    /* ★★番地の 無い 参照＝式が 作った 表★★（2026-09-15 実測で 決まった）
       `=INDEX({1;2;3},0)`      → ★1★（先頭）   ／ `=SUM(INDEX({1;2;3},0))`      → ★6★（全部）
       `=INDEX({1,2;3,4},0,2)`  → ★2★（先頭）   ／ `=SUM(INDEX({1,2;3,4},0,2))`  → ★6★（全部）
       ⇒★1つの 値が 要る 所では 先頭★／★関数へは まるごと★
       ＝★交わる 相手の 行・列が 無い★ので ★暗黙の 交わりの 代わりに 先頭★を 採ります。 */
    if (参.番地なし) return 参.並び[0] || 計.空;
    var r = 参.行, c = 参.列;
    if (参.行数 !== 1) {
      if (!今のマス) return 計.誤('#VALUE!');
      r = 今のマス.行;
      if (r < 参.行 || r > 参.行 + 参.行数 - 1) return 計.誤('#VALUE!');
    }
    if (参.列数 !== 1) {
      if (!今のマス) return 計.誤('#VALUE!');
      c = 今のマス.列;
      if (c < 参.列 || c > 参.列 + 参.列数 - 1) return 計.誤('#VALUE!');
    }
    var v = 参.並び[(r - 参.行) * 参.列数 + (c - 参.列)];
    return v || 計.空;
  }

  /* ══ ★関数へ 渡す 形に する★（四角と 同じ 顔に する） ══ */
  function 引数にする(参) {
    return {
      種: '四角', 並び: 参.並び, 行数: 参.行数, 列数: 参.列数,
      /* ★番地の 無い 物に 嘘の 番地を 付けない★（式が 作った 表） */
      番地: 参.番地なし ? null : { 行: 参.行, 列: 参.列 },
    };
  }

  /* ══ ★引数を「1つの 値」に する★ ══
     ★四角を 渡された 所で 1つの 値が 要る 時★は ★暗黙の 交わり★。
     ★実測で 決まった 物★（これを 知らないと 9本 外します）
       `=MATCH(A1:B5,2,2)` → ★#VALUE!★（2次元は どの 行・列とも 1本に ならない）
       `=INDEX(A1,A1:B5,2)` → ★#VALUE!★（同じ 訳）
       `=INDEX(A1:A5,B1:B5)` → ★2★（B1:B5 の 縦 → 式の 行1 → B1）
     ★番地を 持たない 物（式が 作った 表）は 先頭★＝★暗黙の 交わりの 相手が 居ない★
       （`=INDEX(A1:A5,A1:A5>2)` → 行=FALSE=0 → まるごと → 1／実測 1） */
  function ひとつに(引数, 今のマス) {
    if (!引数) return null;
    if (引数.種 !== '四角') return 引数.値;
    var R = 引数.行数 || 0, C = 引数.列数 || 0;
    if (R <= 1 && C <= 1) return 引数.並び[0] || 計.空;
    if (!引数.番地) return 引数.並び[0] || 計.空;        /* ★式が 作った 表＝先頭★ */
    return 交わる(参照(引数.番地.行, 引数.番地.列, R, C, 引数.並び), 今のマス);
  }

  /* ══ ★大小を 比べる★（MATCH が 使う） ══
     ★実Excel の 並び順★ … 数 ＜ 字 ＜ FALSE ＜ TRUE
     ★字は 大文字小文字を 区別しません★（実測 `=MATCH("BANANA",D1:D5,0)` → 2）
     ★空は 比べられません★（null を 返す＝呼んだ側が 決める） */
  function 位(v) {
    if (!v) return null;
    if (v.型 === '数') return 0;
    if (v.型 === '字') return 1;
    if (v.型 === '真偽') return 2;
    return null;                                   /* 空・誤り は 比べない */
  }
  function 比べる(a, b) {
    var pa = 位(a), pb = 位(b);
    if (pa === null || pb === null) return null;
    if (pa !== pb) return pa < pb ? -1 : 1;
    if (pa === 0) return a.値 < b.値 ? -1 : (a.値 > b.値 ? 1 : 0);
    if (pa === 1) {
      var x = String(a.値).toUpperCase(), y = String(b.値).toUpperCase();
      return x < y ? -1 : (x > y ? 1 : 0);
    }
    var ba = a.値 ? 1 : 0, bb = b.値 ? 1 : 0;
    return ba < bb ? -1 : (ba > bb ? 1 : 0);
  }

  /* ══ ★ワイルドカード★（MATCH の 型 0 が 使う） ══
     ★実測★ `=MATCH("b*",D1:D5,0)` → ★2★（Banana）／`=MATCH("?anana",D1:D5,0)` → ★2★
             `=MATCH("z*",D1:D5,0)` → ★#N/A★
     ★`~` は 次の 1字を そのまま★（実Excel の 決まり／★ここは 未測定★＝棚） */
  function 型紙か(s) {
    /* ★`~` が 在るだけでも 型紙として 見ます★（2026-09-15 に 踏んだ）
       ＝`=MATCH("a~*b",D1:D3,0)` は ★"a*b" の 字そのもの★を 探す 事（実測 → 1）。
         `~` を 外さずに そのまま くらべると ★当たりません★でした。 */
    return /[*?~]/.test(String(s));
  }
  function 型紙に合うか(型紙, 字) {
    var p = String(型紙).toUpperCase(), s = String(字).toUpperCase();
    /* ★後戻りする 突き合わせ★（`*` は 何文字でも） */
    var pi = 0, si = 0, 星 = -1, 印 = 0;
    while (si < s.length) {
      var c = p.charAt(pi);
      if (pi < p.length && c === '~' && /[*?~]/.test(p.charAt(pi + 1))) {
        /* ★`~` の 次が ＊ ？ ~ の 時だけ 逃がし★
           ★それ以外の `~` は そのままの 字★（★未測定★＝実Excel に 打って いません／棚） */
        if (p.charAt(pi + 1) === s.charAt(si)) { pi += 2; si++; continue; }
      } else if (pi < p.length && (c === '?' || c === s.charAt(si))) {
        pi++; si++; continue;
      } else if (pi < p.length && c === '*') {
        星 = pi++; 印 = si; continue;
      }
      if (星 >= 0) { pi = 星 + 1; si = ++印; continue; }
      return false;
    }
    while (p.charAt(pi) === '*') pi++;
    return pi === p.length;
  }

  return {
    参照: 参照, 参照か: 参照か, 交わる: 交わる,
    引数にする: 引数にする, ひとつに: ひとつに,
    比べる: 比べる, 型紙か: 型紙か, 型紙に合うか: 型紙に合うか,
  };
});
