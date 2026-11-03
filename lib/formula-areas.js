/* formula-areas.js — ★AREAS（範囲が 幾つに 分かれているか）★（2026-09-07）
 *
 *  ★★答えは 全部 実Excel に 打たせて 取った★★（4回・合わせて 89本）
 *    `docs/measured/kansuu46/golden-areas-2026-09-07.tsv` … 36本（基本）
 *    `docs/measured/kansuu46/golden-areas2-2026-09-07.tsv` … 22本（範囲で ない 物）
 *    `docs/measured/kansuu46/golden-areas3-2026-09-07.tsv` … 16本（★答えが 割れる 組★）
 *    `docs/measured/kansuu46/golden-areas4-2026-09-07.tsv` … 15本（残していた 穴）
 *    Excel 16.0 build 20326
 *
 *  ★★実Excel の 決まり（測って 分かった）★★
 *    ★AREAS は マスの ★中身★を 1度も 見ません★
 *    見ているのは ★「その 式が マスを 指しているか」★だけ。
 *      `AREAS(IF(TRUE,A1,A2))` … A1 が ★字★でも ★1★
 *      `AREAS(IF(TRUE,D1,B1))` … D1 が ★#DIV/0!★でも ★1★（★誤りは 伝わらない★）
 *      `AREAS(IF(TRUE,1,2))` … ★#VALUE!★（値そのもの）
 *
 *  ★数え方★
 *    ・とびとび（カンマ）… ★その 数だけ★（`(B2:D4,E5,F6:I9)` → 3）
 *    ・★同じ 物を 2回 書いても 2★（`(A1,A1)` → 2）
 *    ・重なり（空白）……… ★重なれば 1★／★重ならなければ #NULL!★
 *    ・名前 …………………… ★画面が 先に 参照に 開く★ので ここへは 来ない
 *      （`=AREAS(とびとび)` → `=AREAS(($B$2:$D$4,$F$6:$I$9))` → ★2★）
 *      ⇒★開いた 結果は ★入れ子の かっこ★に なる★
 *        `(とびとび,A1)` → `(($B$2:$D$4,$F$6:$I$9),A1)` → ★3★（2+1）
 *      ⇒★だから ★入れ子の とびとびを 数える★（下の 数える）
 *    ・★別の シートを またぐ とびとびは #VALUE!★（同じ シート同士なら 数える）
 *    ・空白や `$` は 数に 関係しない（`( A1 , B2 )` も `($A$1,$B$2)` も 2）
 *
 *  ★★出さない 形（★半分 合う 答えを 出さない★）★★
 *    ★中に 関数が 入る 形★（`AREAS(INDEX(...))` など）は ★出しません★
 *    ⇒ 実測 … `AREAS(INDEX(D1:D2,1))` は ★1★／`AREAS(INDEX(A1:C3,99,1))` は ★#REF!★
 *      ⇒★どちらも「計算したら 誤り」＝★値では 見分けられません★★
 *      ⇒ 片方を 通せば もう片方が 間違う
 *    ⇒★★今の 作りでは 出せません★（道具の 話／★Excel が 出来ない のでは ない★）
 *
 *  見張り: tests/formula-areas.test.mjs
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.FormulaAreas = factory();
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var 誤り = function (種) { return { 誤り: 種 }; };

  /** ★一番 外の かっこを 1枚 剥がす★（`(A1,A2)` → `A1,A2`） */
  function 皮をむく(s) {
    var t = s.trim();
    while (t.charAt(0) === '(' && t.charAt(t.length - 1) === ')') {
      /* ★釣り合っている 時だけ 剥がす★（`(A1),(A2)` を 剥がしては いけない） */
      var 深 = 0, 端まで = true, 字中 = false;
      for (var i = 0; i < t.length; i++) {
        var c = t.charAt(i);
        if (字中) { if (c === "'") 字中 = false; continue; }
        if (c === "'") { 字中 = true; continue; }
        if (c === '(') 深++;
        else if (c === ')') { 深--; if (!深 && i < t.length - 1) { 端まで = false; break; } }
      }
      if (!端まで || 深) break;
      t = t.slice(1, -1).trim();
    }
    return t;
  }

  /** ★一番 外の カンマで 割る★（`'a,b'!A1` の 中は 割らない） */
  function カンマで割る(s) {
    var 出 = [], 今 = '', 深 = 0, 字中 = false;
    for (var i = 0; i < s.length; i++) {
      var c = s.charAt(i);
      if (字中) { 今 += c; if (c === "'") 字中 = false; continue; }
      if (c === "'") { 字中 = true; 今 += c; continue; }
      if (c === '(') { 深++; 今 += c; continue; }
      if (c === ')') { 深--; 今 += c; continue; }
      if (c === ',' && !深) { 出.push(今); 今 = ''; continue; }
      今 += c;
    }
    出.push(今);
    return 出;
  }

  /** ★空白で 割る★（重なりの 印）＝★`'二 枚目'!A1` の 中の 空白では 割らない★
   *  ⇒ 2026-09-07 に 踏んだ … ふつうに `split(/\s+/)` したら
   *    `'二 枚目'!A1` が 2つに 割れて ★実Excel が 1と 答える 形を 出せなく なっていた★
   *  ⇒★シート名に 空白を 入れる 人は 居ます★（実際に 測って 気づいた） */
  function 空白で割る(s) {
    var 出 = [], 今 = '', 字中 = false;
    for (var i = 0; i < s.length; i++) {
      var c = s.charAt(i);
      if (字中) { 今 += c; if (c === "'") 字中 = false; continue; }
      if (c === "'") { 字中 = true; 今 += c; continue; }
      if (/\s/.test(c)) { if (今.trim()) 出.push(今.trim()); 今 = ''; continue; }
      今 += c;
    }
    if (今.trim()) 出.push(今.trim());
    return 出;
  }

  /** ★1つの かたまりを 読む★（シート名・$・空白を 落として 形を 見る） */
  function ひとつを読む(s) {
    var t = s.trim();
    if (!t) return null;
    if (/[()]/.test(t)) return null;                 /* ★関数が 入る＝出さない★ */
    var シート = null;
    var m = /^'([^']*)'\s*!\s*(.+)$/.exec(t);
    if (m) { シート = m[1]; t = m[2]; }
    else {
      var m2 = /^([^!'\s]+)\s*!\s*(.+)$/.exec(t);
      if (m2) { シート = m2[1]; t = m2[2]; }
    }
    t = t.replace(/\$/g, '').trim();
    /* ★A1 / A1:B2★ */
    var m3 = /^([A-Za-z]{1,3})([0-9]{1,7})(?::([A-Za-z]{1,3})([0-9]{1,7}))?$/.exec(t);
    if (m3) {
      var c1 = 列番号(m3[1]), r1 = +m3[2];
      var c2 = m3[3] ? 列番号(m3[3]) : c1, r2 = m3[4] ? +m3[4] : r1;
      return { 種: 'マス', シート: シート, 左: Math.min(c1, c2), 右: Math.max(c1, c2),
        上: Math.min(r1, r2), 下: Math.max(r1, r2) };
    }
    /* ★A:A（列ぜんぶ）★ */
    var m4 = /^([A-Za-z]{1,3}):([A-Za-z]{1,3})$/.exec(t);
    if (m4) {
      var a = 列番号(m4[1]), b = 列番号(m4[2]);
      return { 種: 'マス', シート: シート, 左: Math.min(a, b), 右: Math.max(a, b), 上: 1, 下: 1048576 };
    }
    /* ★1:1（行ぜんぶ）★ */
    var m5 = /^([0-9]{1,7}):([0-9]{1,7})$/.exec(t);
    if (m5) {
      var x = +m5[1], y = +m5[2];
      return { 種: 'マス', シート: シート, 左: 1, 右: 16384, 上: Math.min(x, y), 下: Math.max(x, y) };
    }
    /* ★値そのもの★ */
    if (/^"[^"]*"$/.test(t)) return { 種: '値' };
    if (/^-?[0-9]+(\.[0-9]+)?$/.test(t)) return { 種: '値' };
    if (/^(TRUE|FALSE)(\(\))?$/i.test(t)) return { 種: '値' };
    /* ★名前★＝外から 教えて もらう（画面が 持っている） */
    if (/^[A-Za-z_À-￿][A-Za-z0-9_.À-￿]*$/.test(t)) return { 種: '名前', 名: t };
    return null;                                     /* 分からない＝出さない */
  }

  function 列番号(s) {
    var n = 0;
    for (var i = 0; i < s.length; i++) n = n * 26 + (s.toUpperCase().charCodeAt(i) - 64);
    return n;
  }

  function 重なるか(a, b) {
    if ((a.シート || '') !== (b.シート || '')) return false;
    return a.左 <= b.右 && b.左 <= a.右 && a.上 <= b.下 && b.上 <= a.下;
  }

  /** ★AREAS の 答え★
   *  @param 中 かっこの 中の 字（`(A1,A2)` でも `A1,A2` でも よい）
   *  @returns 数 ／ { 誤り:'NULL'|'VALUE' } ／ ★null＝出さない★
   *
   *  ★★名前の 一覧は 受け取りません★★（2026-09-08 に 口を 1つ 消した）
   *    はじめ「名前 → か所の 数」の 表を 受け取る 作りに していましたが、
   *    ★本番の 順で 押したら ★1度も 来ない★事が 分かりました★
   *    ⇒ 画面では `convertFormula` の 頭で ★`名前の箱.開く` が 先に 走る★
   *      `=AREAS(とびとび)` → `=AREAS(($B$2:$D$4,$F$6:$I$9))`
   *    ⇒★名前は ここへ 届く 前に 参照に なっている★
   *    ⇒★★要らない 口を 作り、その 口の 試験まで 書いていました★★
   *      （見張り `unused-param` が「使っていない 口」と 数えて 教えてくれた）
   *    ⇒ 代わりに ★入れ子の とびとび★を 数えられる ように した（それが 本当に 要る 物） */
  /* ★★名前は `数える` に しません★★（2026-09-08 に 踏んだ）
     この repo には `数える()` という 関数が ★他に 4つ★ 在ります
       lib/ribbon.js ／ ribbon-keytips.js ／ ribbon-launcher.js ／ ribbon-context-spec.js
     見張り `tools/unused-param.mjs` は ★名前で 呼び出しを 数えます★。
     ⇒ 私が `数える(部[0])` と ★1つの 引数で★ 呼んだ せいで
       ★上の 4つが「引数つきで 呼ばれている」と 見えて 候補から 消えました★
     ⇒★★見張りの 数が 15個 → 11個に「減った」＝★私が 4つ 隠した★★★
     ⇒★『減った』を 喜ばずに 中身を 見たから 気づけた★
     ⇒ だから ★ぶつからない 名前★に します */
  function 区画を数える(中) {
    if (typeof 中 !== 'string') return null;
    var s = 皮をむく(中);
    if (!s) return null;
    var 組 = カンマで割る(s);
    var 合 = 0;
    var シートたち = {};
    for (var i = 0; i < 組.length; i++) {
      /* ★重なり（空白）★＝1か所に なる（重ならなければ #NULL!） */
      var 部 = 空白で割る(組[i]);
      if (!部.length) return null;
      /* ★★入れ子の とびとび★★（2026-09-08 に 本番の 順で 押して 見つけた）
         ⇒ 画面では ★名前は 先に 参照に 開かれる★（`名前の箱.開く`）
         ⇒ とびとびの 名前が union の 中に 在ると ★かっこが 入れ子に なる★
           `=AREAS((とびとび,A1))` → `=AREAS((($B$2:$D$4,$F$6:$I$9),A1))`
         ⇒★実Excel は 3★（2 + 1）
         ⇒★★だから かたまりが まるごと かっこなら 中を 数え直す★★
         ⇒★これが 無いと ★名前を 使った お客さんだけ 黙って #NAME?★に なる★ */
      if (部.length === 1 && 部[0].charAt(0) === '(' && 皮をむく(部[0]) !== 部[0]) {
        var 中の数 = 区画を数える(部[0]);
        if (typeof 中の数 !== 'number') return 中の数 === null ? null : 中の数;
        合 += 中の数;
        シートたち[''] = 1;
        continue;
      }
      var 面 = [];
      for (var j = 0; j < 部.length; j++) {
        var v = ひとつを読む(部[j]);
        if (!v) return null;                          /* 分からない＝出さない */
        if (v.種 === '値') return 誤り('VALUE');
        if (v.種 === '名前') return null;   /* ★開けなかった 名前＝実Excel も #NAME?★ */
        面.push(v);
      }
      if (面.length > 1) {
        /* ★重なるか★＝重ならなければ #NULL!（実測 `AREAS(B2:D4 A1)`） */
        for (var k = 1; k < 面.length; k++) {
          if (!重なるか(面[k - 1], 面[k])) return 誤り('NULL');
        }
      }
      合 += 1;
      シートたち[面[0].シート || ''] = 1;
    }
    /* ★別の シートを またぐ とびとびは #VALUE!★（実測）
       ⇒ ただし 1つだけの 時は シートが 違っても よい */
    if (組.length > 1 && Object.keys(シートたち).length > 1) return 誤り('VALUE');
    return 合;
  }

  function 足した名前() { return ['AREAS']; }

  return {
    区画を数える: 区画を数える, 皮をむく: 皮をむく, カンマで割る: カンマで割る,
    ひとつを読む: ひとつを読む, 空白で割る: 空白で割る, 重なるか: 重なるか, 列番号: 列番号,
    足した名前: 足した名前
  };
}));
