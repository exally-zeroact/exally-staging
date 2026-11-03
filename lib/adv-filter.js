/* adv-filter.js — ★詳細設定（フィルター オプション）★ 2026-08-30
 *
 *  ★真値（実Excel 16.0 で 実測）★ … tools/measure-data3.ps1
 *    元の 表 A1:A5 ＝ 見出し「名前」／あ・い・あ・う
 *
 *    ① 条件（C1:C2＝「名前」「あ」）で ★別の所へ 出す★（xlFilterCopy=2・Unique=False）
 *         → E1「名前」／E2「あ」／E3「あ」＝★見出し ＋ 2行 の 3行★
 *    ② 同じ 条件で ★重複を 除く★（Unique=True）
 *         → G2「あ」＝★見出し ＋ 1行 の 2行★
 *    ③ ★条件なし★＋重複を 除く
 *         → C1「名前」C2「あ」C3「い」C4「う」＝★見出し ＋ 3行 の 4行★
 *    ④ ★条件なし★＋重複を 残す
 *         → E1「名前」E2「あ」E3「い」E4「あ」E5「う」＝★5行★
 *
 *    ★測り直した★＝はじめ ③を 2行と 読んだが、それは
 *      ★前の 条件の 表が 残っていた★せい。条件の 表を 消して 測り直したら 4行だった。
 *      （「前の 物が 残ったまま 測る」と 嘘の 数が 出る）
 *
 *  ★条件の 書き方（実Excelと 同じ）★
 *    ・条件の 表の ★1行目は 見出し★（元の 表の 見出しと 同じ 字）
 *    ・★同じ 行に 並べた 条件は「かつ」★／★行を 変えた 条件は「または」★
 *    ・`>100` `<=5` `<>あ` `=あ` の ような 書き方が 出来る
 *    ・空の セルは ★条件なし★（何でも 通る）
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.AdvFilter = factory();
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  function 字(v) {
    if (v === null || v === undefined) return '';
    return String(v);
  }

  /** 1つの 条件に 合うか（`>100` `<=5` `<>あ` `=あ` と ふつうの 字） */
  function 合うか(値, 条件) {
    var c = 字(条件).trim();
    if (c === '') return true;                     /* 空＝何でも 通る */
    var v = 値;
    var m = /^(<=|>=|<>|=|<|>)\s*(.*)$/.exec(c);
    if (m) {
      var 記号 = m[1], 右 = m[2].trim();
      var 数v = Number(v), 数右 = Number(右);
      var 数で = isFinite(数v) && 右 !== '' && isFinite(数右);
      if (記号 === '=') return 数で ? 数v === 数右 : 字(v) === 右;
      if (記号 === '<>') return 数で ? 数v !== 数右 : 字(v) !== 右;
      if (!数で) {
        /* 数で 比べられない時は ★字の 並び順★で 比べる（実Excel も 字は 字で 比べる） */
        var a = 字(v), b = 右;
        if (記号 === '<') return a < b;
        if (記号 === '>') return a > b;
        if (記号 === '<=') return a <= b;
        return a >= b;
      }
      if (記号 === '<') return 数v < 数右;
      if (記号 === '>') return 数v > 数右;
      if (記号 === '<=') return 数v <= 数右;
      return 数v >= 数右;
    }
    /* ふつうの 字＝★前が 合えば 通る★（実Excel も 「あ」で「あい」が 通る） */
    var s = 字(v);
    if (c.indexOf('*') >= 0 || c.indexOf('?') >= 0) {
      var re = new RegExp('^' + c.replace(/[.+^${}()|[\]\\]/g, '\\$&')
        .replace(/\*/g, '.*').replace(/\?/g, '.') + '$');
      return re.test(s);
    }
    return s.indexOf(c) === 0;
  }

  /** 条件の 表（1行目＝見出し／2行目から 条件）を 読む
   *  @returns [{見出し名: 条件, …}, …]（★行ごと＝または★・★同じ行の中＝かつ★）
   */
  function 条件を読む(条件表) {
    if (!条件表 || !条件表.length) return [];
    var 見出し = 条件表[0].map(字);
    var 出 = [];
    for (var r = 1; r < 条件表.length; r++) {
      var 組 = {};
      var 何か = false;
      for (var c = 0; c < 見出し.length; c++) {
        var v = 字(条件表[r][c]).trim();
        if (v === '') continue;
        組[見出し[c]] = v;
        何か = true;
      }
      if (何か) 出.push(組);
    }
    return 出;
  }

  /** ★詳細設定★＝元の 表を 条件で 絞り、重複を 除くか どうかも 決める
   *  @param 元   [[見出し…],[値…],…]（★1行目は 見出し★）
   *  @param 条件表 同じ形（無ければ null）
   *  @param 重複を除く true で ★同じ行を 1つに する★
   *  @returns [[見出し…],[値…],…]（★見出しは いつも 付く★＝実測と 同じ）
   */
  function 絞る(元, 条件表, 重複を除く) {
    if (!元 || !元.length) return [];
    var 見出し = 元[0].map(字);
    var 組たち = 条件を読む(条件表);
    var 出 = [元[0].slice()];
    var 見た = {};
    for (var r = 1; r < 元.length; r++) {
      var 行 = 元[r];
      var 通る = 組たち.length === 0;
      for (var k = 0; k < 組たち.length && !通る; k++) {
        var 組 = 組たち[k];
        var 全部合う = true;
        for (var 名 in 組) {
          if (!Object.prototype.hasOwnProperty.call(組, 名)) continue;
          var i = 見出し.indexOf(名);
          if (i < 0) { 全部合う = false; break; }   /* 見出しが 無い＝合わない */
          if (!合うか(行[i], 組[名])) { 全部合う = false; break; }
        }
        if (全部合う) 通る = true;
      }
      if (!通る) continue;
      if (重複を除く) {
        var キー = 行.map(字).join('');
        if (見た[キー]) continue;
        見た[キー] = true;
      }
      出.push(行.slice());
    }
    return 出;
  }

  return { 合うか: 合うか, 条件を読む: 条件を読む, 絞る: 絞る };
}));
