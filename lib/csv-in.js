/* csv-in.js — ★テキストまたは CSV から（データ→データの取得と変換）★ 2026-08-30
 *
 *  ★実Excel を 測って 合わせた所★ … tools/measure-csv.ps1
 *    ・区切りは ★カンマ／タブ★（Excelの「テキスト ファイル ウィザード」も この2つが 既定）
 *    ・「"」で 囲んだ 中の カンマは ★区切りに しない★
 *    ・「""」は ★1つの "★ に なる
 *    ・改行は CRLF も LF も 読む（★中の 改行も "…" の 中なら 1つの セル★）
 *
 *  ★文字の 種類（エンコード）★
 *    日本の CSV は ★Shift_JIS（cp932）★が 多い。ブラウザは UTF-8 で 読むので、
 *    ★化けたら 分かるように★ 読んだ 後に 確かめて、直せる時は cp932 で 読み直す。
 *    （読み直しは 画面の 側で TextDecoder を 使う＝この 部品は 字を 受け取るだけ）
 *
 *  ★黙って 落とさない★＝行の 長さが ばらばらでも ★一番 長い 行に 合わせて★ 空で 埋める。
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.CsvIn = factory();
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  /** 区切りを 当てる（★カンマと タブの どちらが 多いか★・"…" の 中は 数えない） */
  function 区切りを当てる(字) {
    var s = String(字 || '');
    var 中 = false, カンマ = 0, タブ = 0, セミコロン = 0;
    for (var i = 0; i < s.length; i++) {
      var c = s.charAt(i);
      if (c === '"') { 中 = !中; continue; }
      if (中) continue;
      if (c === ',') カンマ++;
      else if (c === '\t') タブ++;
      else if (c === ';') セミコロン++;
    }
    if (タブ > カンマ && タブ >= セミコロン) return '\t';
    if (セミコロン > カンマ && セミコロン > タブ) return ';';
    return ',';
  }

  /**
   * CSV の 字 → 表（[[…],…]）
   * @param 区切り … 省くと 当てる
   */
  function 読む(字, 区切り) {
    var s = String(字 === null || 字 === undefined ? '' : 字);
    if (s.charCodeAt(0) === 0xFEFF) s = s.slice(1);      /* ★BOM を 落とす★ */
    var d = 区切り || 区切りを当てる(s);
    var 表 = [], 行 = [], 今 = '', 中 = false;
    for (var i = 0; i < s.length; i++) {
      var c = s.charAt(i);
      if (中) {
        if (c === '"') {
          if (s.charAt(i + 1) === '"') { 今 += '"'; i++; }   /* ★"" は 1つの "★ */
          else 中 = false;
        } else if (c === '\r') {
          /* ★"…" の 中の 改行は 1つの セル★＝実Excel は 中を 改行1つで 見せる */
          if (s.charAt(i + 1) === '\n') i++;
          今 += '\n';
        } else 今 += c;
        continue;
      }
      if (c === '"') { 中 = true; continue; }
      if (c === d) { 行.push(今); 今 = ''; continue; }
      if (c === '\r') {
        if (s.charAt(i + 1) === '\n') i++;
        行.push(今); 今 = ''; 表.push(行); 行 = [];
        continue;
      }
      if (c === '\n') { 行.push(今); 今 = ''; 表.push(行); 行 = []; continue; }
      今 += c;
    }
    if (今 !== '' || 行.length) { 行.push(今); 表.push(行); }
    /* ★最後の 空行は 落とす（末尾の 改行）★ */
    while (表.length && 表[表.length - 1].length === 1 && 表[表.length - 1][0] === '') 表.pop();
    /* ★行の 長さを そろえる（黙って 落とさない）★ */
    var 幅 = 0;
    for (var j = 0; j < 表.length; j++) if (表[j].length > 幅) 幅 = 表[j].length;
    for (var k = 0; k < 表.length; k++) while (表[k].length < 幅) 表[k].push('');
    return { 表: 表, 区切り: d, 行: 表.length, 列: 幅 };
  }

  /** ★化けているか★＝置き換え字（U+FFFD）が 混ざっていたら 化けている */
  function 化けているか(字) {
    return String(字 || '').indexOf('�') >= 0;
  }

  return { 読む: 読む, 区切りを当てる: 区切りを当てる, 化けているか: 化けているか };
}));
