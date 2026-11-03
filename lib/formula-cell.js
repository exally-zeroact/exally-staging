/* formula-cell.js — ★CELL（セルの 事を 聞く）★（2026-09-07）
 *
 *  ★★答えは 全部 実Excel に 打たせて 取った★★（合わせて 411本）
 *    `docs/measured/kansuu46/golden-cell-2026-09-07.tsv` … 91本（種類 × 場所）
 *    `docs/measured/kansuu46/golden-cell2-2026-09-07.tsv` … 111本（表示形式・そろえ・変わり種）
 *    `docs/measured/kansuu46/golden-cell3-2026-09-07.tsv` … 102本（色・かっこ・円・うちの 表示形式）
 *    `docs/measured/kansuu46/golden-cell4-2026-09-07.tsv` … 10本（別シート・隠し列・鍵・保存後）
 *    `docs/measured/kansuu46/golden-cell5-2026-09-07.tsv` … 81本（日付/時刻の 合図）
 *    `docs/measured/kansuu46/golden-cell6-2026-09-07.tsv` … 16本（残った あいまいな 形）
 *    Excel 16.0 build 20326（UI 1041＝日本語）
 *
 *  ★★1回 測っただけでは 決まらなかった 物★★（＝★どちらでも 合う 組で 決めない★）
 *    ・`[Red]` は ★日本語の Excel では 付かない★（`[赤]`）
 *      ⇒ 2回目は color が ★1度も 1に ならなかった＝物差しが 空洞★だった
 *    ・`$#,##0` は 日本語の Excel では ★通貨に ならない★（,0）／★円（¥）が 通貨★（C0）
 *    ・かっこは ★" " の 中の ( を 数えない★
 *      （`"("#,##0")"` は ★0★／`(#,##0);(#,##0)` は 1）
 *      ⇒★「( が 在れば 1」に していたら 間違えていた★
 *    ・日付は ★年/月/日の うち 2つ 以上★ 要る（`yyyy` だけ は G）
 *      しかも ★年+日（月が 無い）は G★（`yyyy/d` → G）
 *    ・時刻は ★時（h）と 分が 揃わないと G★（`h:ss` も `mm:ss` も `[h]:mm` も G）
 *
 *  ★★実測で 1本も 出なかった 物★★
 *    ・`D4` `D5` … 34通りの 日付/時刻を 打っても 1本も 出なかった ⇒★出しません★
 *
 *  ★★うちが 返せない 物★★（★無い 物は 作らない★＝INFO と 同じ）
 *    ・`filename` … うちは ★ファイルの 置き場を 持たない★（ブラウザの 中）
 *      ⇒★★保存していない Excel と 同じ ""（空）を 返す★★＝実測の 形と 同じ
 *    ・`protect` … うちに ★セルの 鍵★は まだ 無い ⇒ 実Excel の 既定と 同じ 1
 *
 *  見張り: tests/formula-cell.test.mjs
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.FormulaCell = factory();
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var 誤り = function (種) { return { 誤り: 種 }; };

  /* ── 表示形式を 読む 下ごしらえ ───────────────────────────── */

  /** ★`;` で 節に 分ける★（" " の 中・`\x` `!x` `_x` `*x`・`[..]` の 中の `;` は 分けない） */
  function 節に分ける(fmt) {
    var 節 = [], 今 = '';
    for (var i = 0; i < fmt.length; i++) {
      var c = fmt.charAt(i);
      if (c === '"') {                       /* 字の かたまり */
        今 += c; i++;
        while (i < fmt.length && fmt.charAt(i) !== '"') { 今 += fmt.charAt(i); i++; }
        if (i < fmt.length) 今 += fmt.charAt(i);
        continue;
      }
      if (c === '\\' || c === '!' || c === '_' || c === '*') { 今 += c; i++; if (i < fmt.length) 今 += fmt.charAt(i); continue; }
      if (c === '[') {                       /* [赤] [>100] [h] */
        var j = fmt.indexOf(']', i);
        if (j < 0) { 今 += fmt.slice(i); break; }
        今 += fmt.slice(i, j + 1); i = j; continue;
      }
      if (c === ';') { 節.push(今); 今 = ''; continue; }
      今 += c;
    }
    節.push(今);
    return 節;
  }

  /** ★飾りを 落とす★＝`"…"` の 中／`\x` `!x` `_x` `*x`／`[..]` を 取り除く
   *  返す物 … { 字: 残った 記号だけ, 角: [..] の 中身 } */
  function 飾りを落とす(節) {
    var 出 = '', 角 = [];
    for (var i = 0; i < 節.length; i++) {
      var c = 節.charAt(i);
      if (c === '"') { i++; while (i < 節.length && 節.charAt(i) !== '"') i++; continue; }
      if (c === '\\' || c === '!' || c === '_' || c === '*') { i++; continue; }
      if (c === '[') { var j = 節.indexOf(']', i); if (j < 0) break; 角.push(節.slice(i + 1, j)); i = j; continue; }
      出 += c;
    }
    return { 字: 出, 角: 角 };
  }

  /* ★色の 名前★（日本語の Excel と 英語の Excel の 両方）＋ [色5] [Color5] */
  var 色の名 = ['黒', '青', '水', '緑', '紫', '赤', '白', '黄',
    'black', 'blue', 'cyan', 'green', 'magenta', 'red', 'white', 'yellow'];
  function 色の指定か(中) {
    var s = String(中).trim().toLowerCase();
    if (色の名.indexOf(s) >= 0) return true;
    return /^(色|color)\s*\d+$/.test(s);
  }

  /* ── 日付／時刻の 合図（D1〜D9） ─────────────────────────── */

  /** ★字の かたまりごとに 分ける★（mmm → 1かたまり）＝
   *  `m` が ★月★か ★分★かは ★前後の かたまり★で 決まる（実Excel と 同じ） */
  function かたまり(字) {
    var 出 = [], i = 0;
    while (i < 字.length) {
      var c = 字.charAt(i);
      if (/[a-z]/.test(c)) {
        var j = i; while (j < 字.length && 字.charAt(j) === c) j++;
        出.push({ 字: c, 数: j - i, 記号: false }); i = j;
      } else { 出.push({ 字: c, 数: 1, 記号: true }); i++; }
    }
    return 出;
  }

  /** ★日付/時刻の 合図を 出す★ 出なければ null（＝数の 方で 決める） */
  function 日付の合図(節) {
    var 落 = 飾りを落とす(節);
    var s = 落.字.toLowerCase();
    /* ★[h] [mm] [ss] は ★経過時間★＝実測 G（日付/時刻の 合図に しない） */
    for (var a = 0; a < 落.角.length; a++) {
      if (/^[hms]+$/.test(String(落.角[a]).trim().toLowerCase())) return 'G';
    }
    /* ★午前/午後★を 先に 見つけて 取り除く（a は 曜日の aaa と 混ざる） */
    var 午 = /am\/pm|a\/p/.test(s);
    s = s.replace(/am\/pm|a\/p/g, ' ');
    /* ★曜日（aaa/aaaa/ddd…）は 日付の 部品に 数えない★（実測 `aaa` だけ は G） */
    s = s.replace(/a+/g, ' ');

    var 並 = かたまり(s);
    var 年 = false, 月 = false, 日 = false, 時 = false, 分 = false, 秒 = false;
    for (var i = 0; i < 並.length; i++) {
      var t = 並[i];
      if (t.記号) continue;
      if (t.字 === 'y') { 年 = true; continue; }
      if (t.字 === 'd') { if (t.数 <= 2) 日 = true; continue; }  /* ddd/dddd は 曜日 */
      if (t.字 === 'h') { 時 = true; continue; }
      if (t.字 === 's') { 秒 = true; continue; }
      if (t.字 === 'm') {
        /* ★前の かたまりが h／後ろの かたまりが s なら 分★ それ以外は 月 */
        var 前 = null, 後 = null, k;
        for (k = i - 1; k >= 0; k--) { if (!並[k].記号) { 前 = 並[k].字; break; } }
        for (k = i + 1; k < 並.length; k++) { if (!並[k].記号) { 後 = 並[k].字; break; } }
        if (前 === 'h' || 後 === 's') 分 = true; else 月 = true;
        continue;
      }
    }

    /* ★日付が 先★（年月日+時分 の 形は 実測 D1） */
    var 数 = (年 ? 1 : 0) + (月 ? 1 : 0) + (日 ? 1 : 0);
    if (数 >= 2) {
      if (年 && 月 && 日) return 'D1';
      if (年 && 月) return 'D2';
      if (月 && 日) return 'D3';
      return 'G';                       /* ★年+日（月が 無い）は 実測 G★ */
    }
    /* ★時刻は 時と 分が 揃った 時だけ★（実測 h:ss も h だけ も G） */
    if (時 && 分) {
      if (秒 && 午) return 'D6';
      if (午) return 'D7';
      if (秒) return 'D8';
      return 'D9';
    }
    /* ★★日付/時刻の 部品が 在るのに 揃わない 形は G★★
       ⇒★数の 合図へ 落としては いけない★
         実測 `mm:ss.0` は ★G★（落とすと 小数1桁と 見て F1 に なる＝間違い） */
    if (年 || 月 || 日 || 時 || 分 || 秒) return 'G';
    return null;                        /* 日付/時刻では ない */
  }

  /* ── 数の 合図（G / F / , / C / P / S） ──────────────────── */

  function 小数の桁(字) {
    var s = 字.split(/[eE]/)[0];
    var i = s.indexOf('.');
    if (i < 0) return 0;
    var n = 0;
    for (var k = i + 1; k < s.length; k++) {
      var c = s.charAt(k);
      if (c === '0' || c === '#' || c === '?') n++;
      else break;
    }
    return n;
  }

  /* ★通貨は ¥ だけ★（実測 golden-cell7）
     ・`¥#,##0`（裸）／`"¥"#,##0`（字の かたまり）／`¥¥#,##0`（逃がし）／`¥-0`
       … ★4つとも C0★ ⇒★節の どこかに ¥ が 在れば 通貨★
     ・`$#,##0` `"$"#,##0` `"円"#,##0` `#,##0"円"` … ★4つとも ,0★
       ⇒★日本語の Excel では $ も 「円」の 字も 通貨では ない★
     ⇒ だから ここだけは ★飾りを 落とす 前の 字★を 見る
       （落としてしまうと `"¥"` も `¥¥` も 消えて ★半分 合う 計算★に なる） */
  var 通貨の記号 = ['¥', '￥'];

  function 数の合図(節) {
    var 字 = 飾りを落とす(節).字;
    var 桁 = 小数の桁(字);
    if (字.indexOf('%') >= 0) return 'P' + 桁;
    if (/[eE][+-]/.test(字)) return 'S' + 桁;
    for (var i = 0; i < 通貨の記号.length; i++) {
      if (String(節).indexOf(通貨の記号[i]) >= 0) return 'C' + 桁;
    }
    if (/[0#?]\s*\/\s*[0#?]/.test(字)) return 'G';        /* 分数（# ?/?）は 実測 G */
    if (/,/.test(字) && /[0#]/.test(字)) return ',' + 桁;
    if (/[0#]/.test(字)) return 'F' + 桁;
    return 'G';
  }

  /** ★CELL("format") の 答え★ */
  function 書式の合図(表示形式) {
    var fmt = (表示形式 === null || 表示形式 === undefined) ? '' : String(表示形式);
    if (fmt === '' || /^(general|g\/標準)$/i.test(fmt.trim())) return 'G';
    var 節 = 節に分ける(fmt);
    var 正 = 節[0] || '', 負 = 節.length > 1 ? 節[1] : '';
    var 本 = 日付の合図(正);
    if (本 === null) 本 = 数の合図(正);
    /* ★負の 数を 色で 出す 形なら 後ろに `-`★（実測 `0;[赤]0` → F0-） */
    if (色がつくか(fmt)) 本 += '-';
    /* ★正の 数に かっこが 付く 形なら 後ろに `()`★（実測 `(#,##0);(#,##0)` → ,0()） */
    if (かっこがつくか(fmt)) 本 += '()';
    return 本;
  }

  /** ★CELL("color") の 答え★＝負の 節に 色の 指定が 在れば 1 */
  function 色がつくか(表示形式) {
    var fmt = (表示形式 === null || 表示形式 === undefined) ? '' : String(表示形式);
    var 節 = 節に分ける(fmt);
    if (節.length < 2) return false;
    var 角 = 飾りを落とす(節[1]).角;
    for (var i = 0; i < 角.length; i++) if (色の指定か(角[i])) return true;
    return false;
  }

  /** ★CELL("parentheses") の 答え★＝正の 節に ★" " の 外の `(`★ が 在れば 1
   *  （実測 `"("#,##0")"` は ★0★＝字の かたまりの 中の かっこは 数えない） */
  function かっこがつくか(表示形式) {
    var fmt = (表示形式 === null || 表示形式 === undefined) ? '' : String(表示形式);
    var 節 = 節に分ける(fmt);
    return 飾りを落とす(節[0] || '').字.indexOf('(') >= 0;
  }

  /* ── そろえ方（prefix） ─────────────────────────────────── */

  /** ★実測★ 字＝左/既定 `'`／中央 `^`／右 `"`／繰り返し `\`／★数と 空は ""★ */
  function そろえの印(そろえ, 字か) {
    if (!字か) return '';
    var s = String(そろえ || '').toLowerCase();
    if (s === 'center' || s === '中央') return '^';
    if (s === 'right' || s === '右') return '"';
    if (s === 'fill' || s === '繰り返し') return '\\';
    return "'";                                  /* 左・既定・その他 */
  }

  /* ── 中身の 種類（type） ───────────────────────────────── */

  /** ★実測★ 空＝`b`／字（`=""` も 字）＝`l`／それ以外（数・真偽・誤り）＝`v` */
  function 中身の種類(値, 空か) {
    if (空か) return 'b';
    if (typeof 値 === 'string') return 'l';
    return 'v';
  }

  /* ── 幅（width） ───────────────────────────────────────── */

  /** ★うちの 画面は 点（px）で 持ち、実Excel は ★字の 数★で 返す★
   *  ⇒ 既定の 幅を ★実Excel の 既定 8.43文字★と 見て 割る
   *    （実測 … 何も していない 列は 8／広げた 列は そのままの 数）
   *  ★隠した 列は 0★（実測） */
  function 幅を文字にする(点, 既定の点, 隠れ) {
    if (隠れ) return 0;
    var 既定 = Number(既定の点) || 80;
    var p = (点 === null || 点 === undefined || 点 === '') ? 既定 : Number(点);
    if (!isFinite(p) || p < 0) return 0;
    var 一文字 = 既定 / 8.43;
    return Math.round(p / 一文字);
  }

  /* ── 番地（address） ───────────────────────────────────── */

  function 列の名(c) {                       /* 0 → A */
    var s = '', n = Number(c);
    if (!isFinite(n) || n < 0) return '';
    n = Math.floor(n) + 1;
    while (n > 0) { var r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
    return s;
  }

  /** ★実測★ 同じ シート … `$B$3`／別の シート … `[ブック名]シート名!$B$3`
   *  ★うちに ブック名が 無い 時は `[ ]` を 付けない★（空の かっこを 出さない） */
  function 番地(行, 列, 別のシートか, シート名, ブック名) {
    var 芯 = '$' + 列の名(列) + '$' + (Number(行) + 1);
    if (!別のシートか) return 芯;
    var 頭 = ブック名 ? ('[' + ブック名 + ']') : '';
    return 頭 + String(シート名 || '') + '!' + 芯;
  }

  /* ── 出口 ─────────────────────────────────────────────── */

  /** ★CELL を 出す★
   *  @param 何を   'address' 'col' 'row' … （大文字小文字は 問わない＝実測）
   *  @param 見た目 {
   *    行, 列, 別のシートか, シート名, ブック名,
   *    中身, 空か, 表示形式, そろえ, 幅の点, 既定の点, 隠れ
   *  }
   *  @returns 値 または { 誤り: 'VALUE' } */
  /* ★既定の 見た目★＝★実Excel の まっさらな シートと 同じ 答えに なる 値★
     ★2026-09-14 に ここへ 移しました★＝前は `lib/formula-cell-plug.js` の 中だけに 在り、
       ★自前の 皮が 同じ物を 持たず #VALUE! を 返して いました★（＝作る道が 2本）。
     ★既定の点 80★ … 実Excel の 既定の 幅 8.43文字。⇒ `=CELL("width",D1)` → ★8★
     ★ロックなし false★ … うちに 鍵は 無い ⇒ `=CELL("protect",…)` → ★1★（実Excel の 既定） */
  var 既定の見た目 = {
    表示形式: '', そろえ: '', 幅の点: null, 既定の点: 80,
    隠れ: false, シート名: '', ブック名: '', ロックなし: false
  };
  /* ★足りない 所だけ 既定で 埋める★（在る 物は そのまま） */
  function 見た目を埋める(v) {
    var 出 = {};
    for (var k in 既定の見た目) {
      if (Object.prototype.hasOwnProperty.call(既定の見た目, k)) {
        出[k] = (v && v[k] !== undefined && v[k] !== null) ? v[k] : 既定の見た目[k];
      }
    }
    if (v) for (var j in v) if (Object.prototype.hasOwnProperty.call(v, j) && !(j in 出)) 出[j] = v[j];
    return 出;
  }

  function セルの事(何を, 見た目) {
    var k = String(何を === null || 何を === undefined ? '' : 何を).trim().toLowerCase();
    var e = 見た目 || {};
    switch (k) {
      case 'address':     return 番地(e.行, e.列, !!e.別のシートか, e.シート名, e.ブック名);
      case 'col':         return Number(e.列) + 1;
      case 'row':         return Number(e.行) + 1;
      case 'contents':    return e.空か ? 0 : (e.中身 === null || e.中身 === undefined ? 0 : e.中身);
      case 'type':        return 中身の種類(e.中身, !!e.空か);
      case 'width':       return 幅を文字にする(e.幅の点, e.既定の点, !!e.隠れ);
      case 'prefix':      return そろえの印(e.そろえ, typeof e.中身 === 'string' && !e.空か);
      case 'format':      return 書式の合図(e.表示形式);
      case 'color':       return 色がつくか(e.表示形式) ? 1 : 0;
      case 'parentheses': return かっこがつくか(e.表示形式) ? 1 : 0;
      /* ★うちに 鍵は まだ 無い★＝実Excel の 既定と 同じ 1 */
      case 'protect':     return e.ロックなし ? 0 : 1;
      /* ★うちは ファイルの 置き場を 持たない★
         ⇒★保存していない Excel と 同じ ""★（実測 golden-cell4 の「保存する 前」） */
      case 'filename':    return '';
      default:            return 誤り('VALUE');     /* 知らない 種類は #VALUE!（実測） */
    }
  }

  /* ★数える（見張り用）★手で 並べない＝ここが 正本★ */
  function 足した名前() { return ['CELL']; }
  /** ★出せる 種類★＝知らない 字を 足した 時に 気づけるように 並べる */
  function 種類の名前() {
    return ['address', 'col', 'row', 'contents', 'type', 'width',
      'prefix', 'format', 'color', 'parentheses', 'protect', 'filename'];
  }

  return {
    既定の見た目: 既定の見た目, 見た目を埋める: 見た目を埋める,
    セルの事: セルの事,
    書式の合図: 書式の合図, 色がつくか: 色がつくか, かっこがつくか: かっこがつくか,
    そろえの印: そろえの印, 中身の種類: 中身の種類, 幅を文字にする: 幅を文字にする,
    番地: 番地, 列の名: 列の名, 節に分ける: 節に分ける, 飾りを落とす: 飾りを落とす,
    足した名前: 足した名前, 種類の名前: 種類の名前
  };
}));
