/* vba-tejun.js — ★マクロから「手順」を 取り出す★（③レシピにして 会社に残す）
 *
 *  ★なぜ在るか（設計の正本 features.md の順番）★
 *    ①VBAの中身を読んで「何をしているか」を 日本語で言う …………… 出来た（lib/vba.js）
 *    ②同じ事を Exally で どう済ませるかを 1本ずつ出す ……………… 出来た（lib/vba-mikata.js）
 *    ③★レシピにして 会社に残す★＝★人ではなく 会社に残る★ ……… ここ
 *
 *  ★言葉は レシピ（lib/recipe.js）の 物を そのまま使う★＝2つ作らない
 *    使える種類 … 式の列を足す ／ 列の名前を変える ／ 列を消す ／ 並べ替え ／ 切り出す
 *
 *  ★一番 大事な決まり★
 *    ・★取り出せた分だけ 手順にする★。取り出せない行は ★「取り出せません」と 数で言う★。
 *      （黙って 少ない手順を 返すと ★出来たように 見えて 中身が 減る★＝合計が静かに小さくなる型）
 *    ・★当てずっぽうで 埋めない★。列の名前が 分からない式は ★手順にしない★（理由を残す）。
 *    ・★VBAは 動かさない★。読んで 写すだけ。
 *
 *  ★実物での 実測（2026-08-28・司さんの150本の中の マクロ入り5本）★
 *    手続き 14本のうち ★手順が 取り出せたのは 0本★。
 *    ＝この5本は「自分で作った入力画面」と「打ったら 別のシートへ 写す きっかけ」が中心で、
 *      並べ替え・列を消す・式を入れる を していない（★母数が 5本と 小さい事も 併せて 見る★）。
 *    ⇒ ★取り出せない時は ボタンを 出さない★（出来ていない物のボタンを 見せない）。
 */
(function (root) {
  'use strict';

  /** 列の字 → 番号（A=0） */
  function 列の番号(字) {
    var s = String(字 || '').toUpperCase().replace(/[^A-Z]/g, '');
    if (!s) return -1;
    var n = 0;
    for (var i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64);
    return n - 1;
  }
  /** 番号 → 列の字（0=A） */
  function 列の字(c) {
    var s = '';
    c = Number(c);
    if (!(c >= 0)) return '';
    while (c >= 0) { s = String.fromCharCode(65 + (c % 26)) + s; c = Math.floor(c / 26) - 1; }
    return s;
  }

  /* ══ ①1行が「表を ★変えている★ 行」か ══════════════════════
     ★分かった／分からなかった を 数えるための 母数★＝ここを 間違えると 数が 嘘になる。

     ★2026-08-28 実物で 測って 見つけた★（読み取れない42行の 形を 数えた）
       ・`Set 集計シート = Sheets("集計")` ……………………………… 5行
       ・`For i = 2 To …Cells(Rows.Count,1).End(xlUp).Row` ……… 3行
       ・`最終行 = …Cells(Rows.Count, 4).End(xlUp).Row` ………… 何行も
       これらは ★表を 読んでいるだけ★で 1セルも 変えていない。
       それを「読み取れない所」に 数えて ★母数を 膨らませていた★（42行の 半分ちかく）。
       ⇒ ★母数は「表を 変える行」だけ★にする。
     ★日本語の変数名★（Set 集計シート = …）は \w に 当たらないので
       ★字の種類に 頼らない見方★にする（Set で 始まれば 読み飛ばす）。 */
  var 変える呼び = /\.(Delete|Insert|Clear|ClearContents|ClearFormats|Sort|AutoFilter|AdvancedFilter|RemoveDuplicates|Copy|Cut|PasteSpecial|Paste|Merge|UnMerge|AutoFit|TextToColumns)\b/i;
  var 場所 = /(Range|Cells|Columns|Rows|Selection|ActiveCell|Target|UsedRange|CurrentRegion|EntireRow|EntireColumn)\s*[.(]/i;
  var 足す物 = /(Sheets|Worksheets|Names|SortFields|FormatConditions)\s*\.\s*Add\b/i;
  var 読み飛ばす = /^\s*(Attribute\b|Option\b|Dim\b|Const\b|Set\b|'|Rem\b|End\b|If\b|Else|ElseIf\b|For\b|Next\b|Do\b|Loop\b|While\b|Wend\b|With\b|Exit\b|On\s+Error|Select\s+Case|Case\b|GoTo\b|Application\.(ScreenUpdating|EnableEvents|Calculation|DisplayAlerts)|\s*$)/i;

  /** ★その行は 表を 変えているか★（読んでいるだけの行は 数えない） */
  function 変える行(行) {
    var l = String(行 || '');
    if (読み飛ばす.test(l)) return false;
    /* ★条件の 続きの行★＝「… = … Then」で 終わる行は ★比べているだけ★（代入ではない）。
       実物で 2行 これを「書いている」と 数えていた（2026-08-28 実測）。
       ＝If が 何行かに 分かれて 書いてあると 頭の If が 前の行に 在る。 */
    if (/\bThen\s*$/i.test(l)) return false;
    if (足す物.test(l)) return true;
    if (変える呼び.test(l) && 場所.test(l)) return true;
    /* 代入の ★左側★に 場所が 在れば「変える行」。右側だけなら「読むだけ」 */
    var i = l.indexOf('=');
    while (i > 0 && /[<>!]/.test(l.charAt(i - 1))) i = l.indexOf('=', i + 1);
    if (i < 0) return false;
    return 場所.test(l.slice(0, i));
  }

  /* ══ ②1行を 手順に する ═════════════════════════════════ */

  /** 並べ替え … .Sort Key1:=Range("B2"), Order1:=xlDescending ／ SortFields.Add Key:=Range("B2:B9") */
  function 並べ替えを読む(行) {
    if (!/\.Sort\b|SortFields\.Add/i.test(行)) return null;
    var m = /Key1?\s*:=\s*(?:\w+\.)*Range\("([A-Z]+)\d*(?::[A-Z]+\d*)?"\)/i.exec(行);
    if (!m) {
      var c = /Key1?\s*:=\s*(?:\w+\.)*Cells\(\s*\d+\s*,\s*(\d+)\s*\)/i.exec(行);
      if (!c) return { だめ: '並べ替える列が 書かれていません' };
      m = [null, 列の字(Number(c[1]) - 1)];
    }
    var 降順 = /Order1?\s*:=\s*xlDescending/i.test(行);
    return { 手順: { 種類: '並べ替え', 列: String(m[1]).toUpperCase(), 向き: 降順 ? '降順' : '昇順' } };
  }

  /** 列を消す … Columns("C").Delete ／ Columns(3).Delete ／ Range("C:C").Delete */
  function 列を消すを読む(行) {
    if (!/\.Delete\b/i.test(行)) return null;
    if (/Rows\(|EntireRow/i.test(行)) return { だめ: '行を消す作業は まだ 手順に 出来ません' };
    var m = /Columns\(\s*"([A-Z]+)"\s*\)/i.exec(行);
    if (m) return { 手順: { 種類: '列を消す', 列: m[1].toUpperCase() } };
    var n = /Columns\(\s*(\d+)\s*\)/i.exec(行);
    if (n) return { 手順: { 種類: '列を消す', 列: 列の字(Number(n[1]) - 1) } };
    var r = /Range\("([A-Z]+):[A-Z]+"\)/i.exec(行);
    if (r) return { 手順: { 種類: '列を消す', 列: r[1].toUpperCase() } };
    if (/EntireColumn/i.test(行)) return { だめ: '消す列が どこかを 書いていません' };
    return null;
  }

  /** 見出しを書く … Range("E1").Value = "税込" ／ Cells(1, 5).Value = "税込" */
  function 見出しを読む(行) {
    var m = /Range\("([A-Z]+)1"\)\s*\.\s*Value\s*=\s*"([^"]+)"/i.exec(行);
    if (m) return { 列: m[1].toUpperCase(), 名: m[2] };
    var c = /Cells\(\s*1\s*,\s*(\d+)\s*\)\s*\.\s*Value\s*=\s*"([^"]+)"/i.exec(行);
    if (c) return { 列: 列の字(Number(c[1]) - 1), 名: c[2] };
    return null;
  }

  /** 式を入れる … Range("E2:E100").Formula = "=D2*1.1" */
  function 式を読む(行) {
    var m = /Range\("([A-Z]+)(\d+)(?::([A-Z]+)\d+)?"\)\s*\.\s*Formula(?:R1C1)?\s*=\s*"(=[^"]*)"/i.exec(行);
    if (!m) return null;
    var 列 = m[1].toUpperCase();
    if (m[3] && m[3].toUpperCase() !== 列) return { だめ: '式を入れる先が 1列では ありません' };
    var 始まり = Number(m[2]);
    var 式 = m[4];
    if (/R1C1/i.test(行)) return { だめ: 'R1C1の書き方の式は まだ 読めません' };
    /* ★行の番号を {行} に 置き換える★＝レシピは 1行ずつ 当てる形で 持つ。
       ★置き換えるのは「その式が 始まる行の番号」だけ★（他の数字を 触ると 式が 変わる）。 */
    var 直した = 式.replace(new RegExp('(\\$?[A-Z]{1,3}\\$?)' + 始まり + '(?![0-9])', 'g'), '$1{行}');
    return { 列: 列, 式: 直した, 動かした: 直した !== 式 };
  }

  /** ★写して貼る★
   *    Range("A1:D9").Copy Destination:=Range("F1")
   *    Range("F1:I9").Value = Range("A1:D9").Value
   *  ★別のシートが 出てくる物は 断る★＝今の当て方は ★同じシートの中だけ★。
   *  勝手に 同じシートへ 当てると ★別の所を 書き換える★事になる。 */
  var よそのシート = /(Sheets|Worksheets|Workbooks)\s*[.(]/i;
  function 写すを読む(行) {
    var l = String(行 || '');
    /* ★写している行かどうかを 先に 見分ける★＝そうしないと
       「別のシートへ 写している」のに ★「知らない書き方」という 中身の無い理由★が出る。 */
    var 写しの形 = (/\.Copy\b/i.test(l) && /Destination\s*:=/i.test(l))
      || /Range\("[A-Z]+\d+[^"]*"\)\s*\.\s*Value2?\s*=\s*[^=]*Range\("/i.test(l);
    if (!写しの形) return null;
    if (よそのシート.test(l)) return { だめ: '別のシートへ 写す物は まだ 手順に 出来ません' };

    var m = /Range\("([A-Z]+\d+(?::[A-Z]+\d+)?)"\)\s*\.\s*Copy\s+Destination\s*:=\s*(?:\w+\.)?Range\("([A-Z]+\d+)/i.exec(l);
    if (m) return { 手順: { 種類: '写す', 元: m[1].toUpperCase(), 先: m[2].toUpperCase(), 値だけ: false } };

    var v = /Range\("([A-Z]+\d+(?::[A-Z]+\d+)?)"\)\s*\.\s*Value2?\s*=\s*(?:\w+\.)?Range\("([A-Z]+\d+(?::[A-Z]+\d+)?)"\)\s*\.\s*Value2?/i.exec(l);
    /* ★左が 貼り先・右が 元★（VBAの書き方の向き） */
    if (v) return { 手順: { 種類: '写す', 元: v[2].toUpperCase(), 先: String(v[1]).toUpperCase().split(':')[0], 値だけ: true } };

    return { だめ: '写す先が どこかを 読めません' };
  }

  /* ══ ③手続き1本を まるごと 読む ═══════════════════════════
     @param 手続き … vba-mikata の 手続きに切る() が返す 1本
     @returns {手順, 取り出せなかった, 数:{効く行, 分かった, 分からない}} */
  function 取り出す(手続き) {
    var 行たち = String((手続き && 手続き.中身) || '').split(/\r\n|\r|\n/);
    var 手順 = [], 取り出せなかった = [];
    var 効いた = 0;
    var 見出し = {};      /* 列 → 名前（式より 先に 書かれる事が 多い） */
    var 式待ち = [];      /* 見出しが まだ 見つかっていない式 */

    for (var i = 0; i < 行たち.length; i++) {
      var 行 = 行たち[i];
      if (!変える行(行)) continue;
      効いた++;

      var h = 見出しを読む(行);
      if (h) { 見出し[h.列] = h.名; 手順.push({ 種類: '列の名前を変える', 元: h.列, 新: h.名, 行: i + 1 }); continue; }

      var f = 式を読む(行);
      if (f) {
        if (f.だめ) { 取り出せなかった.push({ 行: i + 1, なぜ: f.だめ }); continue; }
        式待ち.push({ 列: f.列, 式: f.式, 行: i + 1 });
        continue;
      }

      var w = 写すを読む(行);
      if (w) {
        if (w.だめ) { 取り出せなかった.push({ 行: i + 1, なぜ: w.だめ }); continue; }
        w.手順.行 = i + 1;
        手順.push(w.手順);
        continue;
      }

      var s = 並べ替えを読む(行);
      if (s) {
        if (s.だめ) { 取り出せなかった.push({ 行: i + 1, なぜ: s.だめ }); continue; }
        s.手順.行 = i + 1;
        手順.push(s.手順);
        continue;
      }

      var d = 列を消すを読む(行);
      if (d) {
        if (d.だめ) { 取り出せなかった.push({ 行: i + 1, なぜ: d.だめ }); continue; }
        d.手順.行 = i + 1;
        手順.push(d.手順);
        continue;
      }

      取り出せなかった.push({ 行: i + 1, なぜ: 'この書き方は まだ 手順に 出来ません' });
    }

    /* ★式は 列の名前と 組にして はじめて 手順になる★
       ＝名前が 分からない式を「列E」などと 勝手に 名づけない（当てずっぽうで 埋めない）。 */
    for (var k = 0; k < 式待ち.length; k++) {
      var w = 式待ち[k];
      var 名 = 見出し[w.列];
      if (!名) { 取り出せなかった.push({ 行: w.行, なぜ: 'この式の 列の名前が どこにも 書かれていません' }); continue; }
      /* 名前だけの手順は 式の手順に まとめる（同じ列に 2本 出さない） */
      for (var q = 手順.length - 1; q >= 0; q--) {
        if (手順[q].種類 === '列の名前を変える' && 手順[q].元 === w.列) { 手順.splice(q, 1); break; }
      }
      手順.push({ 種類: '式の列を足す', 見出し: 名, 式: w.式, 行: w.行 });
    }
    手順.sort(function (a, b) { return (a.行 || 0) - (b.行 || 0); });

    return {
      手順: 手順,
      取り出せなかった: 取り出せなかった,
      /* ★行の数と 手順の数は 別★（列の名前と 式は 2行で 1つの手順になる）。
         ★同じ数のように 書くと どちらかが 嘘になる★ので 分けて 持つ。 */
      数: {
        変える行: 効いた,          /* ★表を 変えている行だけ★（読むだけの行は 数えない） */
        読めた行: Math.max(0, 効いた - 取り出せなかった.length),
        読めない行: 取り出せなかった.length,
        手順: 手順.length,
      },
    };
  }

  /* ══ ④客に見せる 1行（★減った事を 隠さない★） ═══════════════ */
  function 知らせの字(出) {
    if (!出 || !出.数 || !出.数.変える行) return 'この中に 表を 変える所が 見つかりません。';
    if (!出.手順.length) {
      return '表を 変える所が ' + 出.数.変える行 + 'か所 ありますが、'
        + 'まだ 手順として 覚えられる書き方では ありません。';
    }
    var s = '表を 変える ' + 出.数.変える行 + 'か所のうち ' + 出.数.読めた行 + 'か所を 読めました。'
      + '手順は ' + 出.数.手順 + 'つです。';
    if (出.数.読めない行) {
      s += '残り ' + 出.数.読めない行 + 'か所は 読み取れていません（覚えるのは この ' + 出.数.手順 + 'つだけです）。';
    }
    return s;
  }

  /** 手順を 日本語の1行にする（見せる時だけ・当てるのは lib/recipe.js） */
  function 手順の字(t) {
    if (!t) return '';
    if (t.種類 === '並べ替え') return '列 ' + t.列 + ' で ' + (t.向き || '昇順') + 'に 並べ替える';
    if (t.種類 === '列を消す') return '列 ' + t.列 + ' を 空にする';
    if (t.種類 === '列の名前を変える') return '列 ' + t.元 + ' の名前を「' + t.新 + '」にする';
    if (t.種類 === '式の列を足す') return '「' + t.見出し + '」の列を 足す（' + t.式 + '）';
    if (t.種類 === '写す') return t.元 + ' を ' + t.先 + ' へ 写す' + (t.値だけ ? '（値だけ）' : '');
    if (t.種類 === '切り出す') return '1人分・何月分を 切り出す';
    return String(t.種類 || '');
  }

  var api = {
    取り出す: 取り出す,
    変える行: 変える行,
    知らせの字: 知らせの字,
    手順の字: 手順の字,
    列の字: 列の字,
    列の番号: 列の番号,
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.VbaTejun = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
