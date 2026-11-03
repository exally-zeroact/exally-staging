/* shiki-sashikomi.js — ★行/列を 入れたり 消したりした時、式の 参照を 追従させる★（2026-09-13）
 *
 *  ★★借り物の 中で 一番 手が 込んでいる所★★
 *    借り物（HyperFormula・GPLv3・944,635字）では
 *    `LazilyTransformingAstService` が やって いる 仕事です。
 *    ★司さん 2026-09-13「借り物が あったら 商用など いろいろ 引っかかる。全部 やれや」★
 *
 *  ★★答えは 全部 実Excel に 打って 測りました★★
 *    （2026-09-13・★17通り★／道具 `docs/measured/toru-sashikomi.ps1`
 *      紙 `docs/measured/golden-sashikomi-2026-09-13.tsv`／Excel 16.0 build 20326）
 *
 *  ★★当て推量なら 外して いた 物★★
 *    `=$A$2+1` に 行を 入れる → ★`=$A$3+1`★
 *      ＝★「絶対参照」も 追従する★（★絶対 という 名前なのに 動く★）
 *    `=SUM(A1:A3)` の ★途中★(2行目)に 入れる → ★`=SUM(A1:A4)`★（★広がる★）
 *    `=SUM(A2:A4)` の ★先頭★(2行目)に 入れる → ★`=SUM(A3:A5)`★（★広がらず ずれる★）
 *      ＝★1行 違うだけで 答えが 変わる★
 *    `=SUM(A2:A2)` の 中を 消す → ★`=SUM(#REF!)`★（四角 丸ごと 消えた）
 *    `=A1+A2` の 1行目に 入れる → ★`=A2+A3`★（★式 自身が 動いても 参照は 追う★）
 *
 *  ★★この 台が 見て いない 範囲★★
 *    ・複数の 板（シート）を またぐ 参照（`板!A1`）
 *    ・名前を 付けた 範囲／表（Table[列名]）
 *    ・溢れ（土台⑤）
 *    ⇒★出来て いない 物を 出来た 顔で 混ぜない★
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(
    typeof require === 'function' ? require('./shiki-kiru.js') : root.ShikiKiru);
  else root.ShikiSashikomi = factory(root.ShikiKiru);
})(typeof self !== 'undefined' ? self : this, function (切) {
  'use strict';

  var 参照の形 = /^(\$?)([A-Z]{1,3})(\$?)([0-9]{1,7})$/;

  function 列を数に(s) {
    var n = 0;
    for (var i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64);
    return n;
  }
  function 数を列に(n) {
    var s = '';
    while (n > 0) { var r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = ((n - r) / 26) | 0; }
    return s;
  }

  /* ★1つの 参照（A1・$A$1）を ほどく★ */
  function ほどく(字) {
    var m = 参照の形.exec(String(字).toUpperCase());
    if (!m) return null;
    return { 列固定: m[1] === '$', 列: 列を数に(m[2]), 行固定: m[3] === '$', 行: +m[4] };
  }
  function 組む(p) {
    if (p.こわれた) return '#REF!';
    return (p.列固定 ? '$' : '') + 数を列に(p.列) + (p.行固定 ? '$' : '') + p.行;
  }

  /* ══ ★1つの 番号を ずらす★ ══
     ★実測の 決まり★
       入れる … その 番号 ★以上★ なら ＋台数（★$ が 付いて いても 動く★）
       消す   … 消した 範囲の 中なら ★こわれる★／後ろなら −台数 */
  function ずらす(番号, どこ, 台数, 消すか) {
    if (!消すか) return (番号 >= どこ) ? { 値: 番号 + 台数 } : { 値: 番号 };
    if (番号 >= どこ && 番号 < どこ + 台数) return { こわれた: true };
    if (番号 >= どこ + 台数) return { 値: 番号 - 台数 };
    return { 値: 番号 };
  }

  /* ══ ★四角（A1:A3）の 端を ずらす★ ══
     ★実測の 決まり★
       途中に 入れる … ★広がる★（A1:A3 に 2行目 → A1:A4）
       先頭に 入れる … ★広がらず ずれる★（A2:A4 に 2行目 → A3:A5）
       中を 消す     … ★縮む★（A1:A3 から 2行目 → A1:A2）
       丸ごと 消える … ★#REF!★ */
  function 四角をずらす(上, 下, どこ, 台数, 消すか) {
    if (!消すか) {
      /* ★下の端は「どこ 以上」で 動く／上の端も 同じ★
         ⇒ 上が どこ 以上なら 上も 下も 動く＝★ずれる★
            上が どこ 未満で 下が どこ 以上なら 下だけ 動く＝★広がる★ */
      var 新上 = (上 >= どこ) ? 上 + 台数 : 上;
      var 新下 = (下 >= どこ) ? 下 + 台数 : 下;
      return { 上: 新上, 下: 新下 };
    }
    /* ★「丸ごと 消えた」の 判定は 下の `b < a` が 拾います★
       ＝2026-09-13 に ここに 同じ 判定を 2つ 書いて いました。
       ★片方を 壊しても 緑の まま★＝★死にコード★だったので 消しました
       （[[feedback_mihon_no_michi_ga_futatsu_aru_toki_katahou_dake_naosu_na]] の 裏返し）。 */
    var 消し上 = どこ, 消し下 = どこ + 台数 - 1;
    var a = 上, b = 下;
    if (上 > 消し下) a = 上 - 台数;
    else if (上 >= 消し上) a = 消し上;
    if (下 > 消し下) b = 下 - 台数;
    else if (下 >= 消し上) b = 消し上 - 1;
    if (b < a) return { こわれた: true };
    return { 上: a, 下: b };
  }

  /* ══ ★式の 字を 書き換える★ ══
     ★字の 中（"A2" の ような 文字列）は 触らない★＝`shiki-kiru.js` で 切ってから 見る */
  function 式を直す(式, わざ) {
    var s = String(式 == null ? '' : 式);
    if (s.charAt(0) !== '=') return s;
    var か = 切.切る(s.slice(1));
    if (!か.ok) return s;
    var 出 = '=';
    var かたまり = か.出;
    for (var i = 0; i < かたまり.length; i++) {
      var t = かたまり[i];
      if (t.型 !== '名') { 出 += t.字; continue; }
      /* ★四角（A1:A3）は 記号 ':' で 繋がった 2つの 名前★ */
      var 次 = かたまり[i + 1], 次々 = かたまり[i + 2];
      if (次 && 次.型 === '記' && 次.字 === ':' && 次々 && 次々.型 === '名'
          && ほどく(t.字) && ほどく(次々.字)) {
        出 += 四角の字を直す(t.字, 次々.字, わざ);
        i += 2;
        continue;
      }
      出 += 参照の字を直す(t.字, わざ);
    }
    return 出;
  }

  function 参照の字を直す(字, わざ) {
    var p = ほどく(字);
    if (!p) return 字;                                  /* TRUE・関数名 など */
    var 縦 = (わざ.向き === '行');
    var r = ずらす(縦 ? p.行 : p.列, わざ.どこ, わざ.台数 || 1, わざ.消すか);
    if (r.こわれた) return '#REF!';
    if (縦) p.行 = r.値; else p.列 = r.値;
    return 組む(p);
  }

  function 四角の字を直す(左字, 右字, わざ) {
    var a = ほどく(左字), b = ほどく(右字);
    if (!a || !b) return 左字 + ':' + 右字;
    var 縦 = (わざ.向き === '行');
    var 上 = 縦 ? Math.min(a.行, b.行) : Math.min(a.列, b.列);
    var 下 = 縦 ? Math.max(a.行, b.行) : Math.max(a.列, b.列);
    var r = 四角をずらす(上, 下, わざ.どこ, わざ.台数 || 1, わざ.消すか);
    if (r.こわれた) return '#REF!';
    if (縦) { a.行 = r.上; b.行 = r.下; } else { a.列 = r.上; b.列 = r.下; }
    /* ★もう 片方の 向きは そのまま★（行を 入れた時 列は 動かない） */
    return 組む(a) + ':' + 組む(b);
  }

  return {
    式を直す: 式を直す,
    ほどく: ほどく, 組む: 組む,
    ずらす: ずらす, 四角をずらす: 四角をずらす,
    列を数に: 列を数に, 数を列に: 数を列に
  };
});
