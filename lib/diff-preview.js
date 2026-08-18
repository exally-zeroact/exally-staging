/* diff-preview.js — ★直す前に「何を書き込むか」を見せる物を作る（画面は持たない純関数）★
 *
 * ═══ なぜ要るのか（2026-08-18 実物で測った）═══════════════════════════
 *   司さんの実物（14シート・式15,126本）を本番と同じ経路で開いて測った:
 *     開いただけ（何も触らない）        … 書き込むセル ★0本★
 *     計算!C10 を 5600 → 9999 に直した … ★18本★（売上表6 / 計算6 / 月別6）
 *   ★人は「1つ直した」つもりでも 3シート18本が書き換わる★。
 *   今までは保存を押すと その場で書き込んでいた＝何が変わるか見る場所が無かった。
 *   方針ver.6「絶対に守る3つ」の②＝★直す前に必ず見せる★。
 *
 * ═══ ★数える場所は1か所★ ═════════════════════════════════════════
 *   ここは ★書き込む側が使うのと同じ物（BookOpen.changedCells）★を渡してもらう。
 *   自分で数え直さない。★見せた数と実際に書いた数が違う★のが この機能で
 *   いちばんやってはいけない壊れ方なので、数える口を2つにしない。
 *
 * ═══ 見せ方の決まり（指示役 2026-08-18 の裁定）═══════════════════════
 *   ・★件数は省略しない★（シートごとの件数は全部 出す）／中身だけ先頭3行
 *   ・★画面と同じ書式で見せる★（961827.2727… ではなく 961,827）
 *   ・★「あなたが直した所」と「波及した所」を分ける★
 *   ・★式を直した時は 値ではなく式で見せる（前の式 → 後の式）★
 *   ・★前が分からない時は「分かりません」と出す★（黙って0や空にしない）
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.DiffPreview = factory();
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var UNKNOWN = '分かりません';

  function colName(i) {
    var s = '';
    i++;
    while (i > 0) { var m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = (i - 1 - m) / 26; }
    return s;
  }

  /**
   * 見せる物を作る。
   * @param {Object} o
   *   o.sheets        グリッドのシート配列 [{name, data}]
   *   o.changedCells  ★書き込む側と同じ関数★ (sheet) => {'r,c': 値}
   *   o.base          開いた時の控え {'シート名|r,c': 値}
   *   o.edited        人が直したセル {'シート名|r,c': {beforeF:(式|null)}}（無くてよい）
   *   o.format        (値, cell) => 画面に出す字（無ければそのまま）
   *   o.maxRows       1シートに見せる行数（既定3）
   * @returns {{total:number, userCount:number, spreadCount:number,
   *            sheets:Array<{name:string,count:number,rows:Array,more:number}>}}
   */
  function build(o) {
    var sheets = o.sheets || [];
    var changed = o.changedCells;
    var base = o.base || {};
    var edited = o.edited || {};
    var fmt = o.format || function (v) { return v === null || v === undefined ? '' : String(v); };
    var maxRows = (o.maxRows === undefined || o.maxRows === null) ? 3 : o.maxRows;

    var out = { total: 0, userCount: 0, spreadCount: 0, sheets: [] };
    for (var i = 0; i < sheets.length; i++) {
      var sh = sheets[i];
      var ch = changed(sh) || {};
      var keys = Object.keys(ch);
      if (!keys.length) continue;
      /* 番地の順に並べる（行→列）。人が読む順にする */
      keys.sort(function (a, b) {
        var pa = a.split(','), pb = b.split(',');
        return (+pa[0] - +pb[0]) || (+pa[1] - +pb[1]);
      });
      var rows = [];
      for (var k = 0; k < keys.length; k++) {
        var key = keys[k];
        var p = key.split(','), r = +p[0], c = +p[1];
        var cell = (sh.data && sh.data[key]) || {};
        var mark = edited[sh.name + '|' + key];
        var byUser = !!mark;
        out.total++;
        if (byUser) out.userCount++; else out.spreadCount++;
        if (rows.length >= maxRows) continue;         // ★件数は数えたうえで、中身だけ絞る★

        /* ★式で見せるのは「人が式を直した時」だけ★。
           波及しただけの式セルは ★値が変わった★のであって 式は変わっていない。
           そこを式で見せると「分かりません → =IFERROR(…)」という読めない行が並ぶ
           （2026-08-18 実ブラウザで実物を開いて見つけた）。 */
        var isF = typeof cell.f === 'string' && cell.f.charAt(0) === '=' && byUser;
        var row = { addr: colName(c) + (r + 1), byUser: byUser, kind: isF ? 'formula' : 'value' };
        if (isF) {
          /* ★式を直した時は 式で見せる★。前の式を控えていなければ ★分かりません★ */
          row.before = (mark && typeof mark.beforeF === 'string') ? mark.beforeF : UNKNOWN;
          row.beforeUnknown = !(mark && typeof mark.beforeF === 'string');
          row.after = cell.f;
        } else {
          var b = base[sh.name + '|' + key];
          var known = !(b === undefined);
          row.before = known ? fmt(b, cell) : UNKNOWN;
          row.beforeUnknown = !known;
          row.after = fmt(ch[key], cell);
        }
        rows.push(row);
      }
      out.sheets.push({
        name: sh.name, count: keys.length, rows: rows,
        more: Math.max(0, keys.length - rows.length),
        hasUser: rows.some(function (x) { return x.byUser; }) || keysHaveUser(sh, keys, edited),
      });
    }
    /* ★人が直したシートを一番上に出す★（2026-08-18 実物で撮って気づいた）
       シートの並び順のままだと、14シートの11枚目を直した人は
       ★自分がやった1つを探さないといけない★。自分の手より先に 波及が並ぶのはおかしい。
       ★並べ替えるだけ。件数も中身も言葉も変えない★ */
    out.sheets.sort(function (a, b) { return (b.hasUser ? 1 : 0) - (a.hasUser ? 1 : 0); });
    /* 同じシートの中でも、人が直した行を先に出す */
    for (var s2 = 0; s2 < out.sheets.length; s2++) {
      out.sheets[s2].rows.sort(function (x, y) { return (y.byUser ? 1 : 0) - (x.byUser ? 1 : 0); });
    }
    return out;
  }

  /* 見せていない行も含めて「人が直した所が在るシートか」を見る
     （先頭3行に入らなかった所を直した時も、そのシートを上に出す） */
  function keysHaveUser(sh, keys, edited) {
    for (var i = 0; i < keys.length; i++) if (edited[sh.name + '|' + keys[i]]) return true;
    return false;
  }

  /** 窓に出す1行目の言葉（★数を必ず入れる★）
   *  ★言い方は「書き出す ↔ 読み込む」で全アプリ統一（2026-08-18 決定）★
   *    「保存する／落とす／Excelにする」は2通り目の言い方になるので使わない。
   *  ★「◯◯に 書き込みます」と言わない★＝元のファイルに上書きされると読める。
   *    実際は ★元のファイルは1バイトも変わらず、同じ名前で1本 書き出す★だけ。 */
  function headline(plan, fileName) {
    if (!plan || !plan.total) return '';
    var n = plan.sheets.length;
    return (fileName ? fileName + ' を直した物を 書き出します' : '直した物を 書き出します')
      + '（' + plan.total + 'か所'
      + (n > 1 ? '・' + n + 'つのシートに広がっています' : '') + '）';
  }
  /** ボタンの言葉（★見せた数と同じ数を入れる★） */
  function goLabel(plan) {
    return 'この' + ((plan && plan.total) || 0) + 'か所を直して 書き出す';
  }

  return { build: build, headline: headline, goLabel: goLabel, colName: colName, UNKNOWN: UNKNOWN };
}));
