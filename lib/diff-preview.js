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

    var out = { total: 0, userCount: 0, spreadCount: 0, otherSheetCount: 0, sheets: [] };
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
      var rows = [], 板の人 = 0, 板のつられ = 0;
      for (var k = 0; k < keys.length; k++) {
        var key = keys[k];
        var p = key.split(','), r = +p[0], c = +p[1];
        var cell = (sh.data && sh.data[key]) || {};
        var mark = edited[sh.name + '|' + key];
        var byUser = !!mark;
        if (byUser) 板の人 += 1; else 板のつられ += 1;
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
        /* ★★板ごとの 数は 先頭3行では なく 全部を 数えます★★（2026-09-22）
             `rows` は 先頭3行だけ ⇒★行から 数えると 嘘に なります★ */
        userCount: 板の人, spreadCount: 板のつられ,
      });
      /* ★★人が 1マスも 触って いない 板★★＝★つられて 動いた だけの 板★
           ＝★お客さんが 一番 気づかない 所★（2026-09-22 実Excel で 割れた） */
      if (!板の人) out.otherSheetCount += 板のつられ;
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

  /* ══ ★★どこを 触ったかを 1文で 言う★★ ══（2026-09-22）
       ★★司さんの 注文★★（ア）
         「★どこを 触ったか どんな 関数や マクロが 組まれてるかは
           簡潔に 文に して ドロップダウンで 詳しく★」
       ⇒★詳しい 所は もう 在ります★（`build()` の `sheets[].rows`）
       ⇒★★足りないのは 1文の 方でした★★
       ★★形は 経営者1 の 注文で 決めました★★（2026-09-21）
         「★私が 実Excel で 割れる 形に して ください★」
         ＝「★A1 を 3 から 42 に した★」の ように
           ★マスと 前後の 値★ が 在れば ★向こうが 開いて 突き合わせられます★
       ★★あなたが 直した 所を 先に 言います★★
         ＝`build()` が 「あなたが 直した 所」と 「つられて 変わった 所」を
           前から 分けて います（2026-08-18 の 裁定）
         ＝★実物では 1つ 直すと 3板 18マスが 変わります★
         ⇒★先に 「あなたが 直した 所」を 言わないと 何を したか 分かりません★
       ★★前が 分からない 時は 「分かりません」と 出します★★
         ＝この 台の 元からの 決め（★黙って 0や 空に しない★）
       ★3つ 以上は 「ほか ◯か所」に します★＝★1文に する のが 注文★ */
  function 文(plan) {
    if (!plan || !plan.total) return '変える所はありません';
    /* ★★口を 増やしません★★（2026-09-22 `unused-param` が 捕まえました）
         最初は `maxRows` を 受ける 形に しましたが ★誰も 渡して いません★。
         ⇒★渡されない 口は 「そのうち 使う」では なく 「今は 嘘」★ です。
         ⇒★★要る 日に 足します★★（門の 上限を 上げて 通さない） */
    var 上限 = 2;
    var 私 = [], 全 = 0;
    for (var i = 0; i < (plan.sheets || []).length; i++) {
      var sh = plan.sheets[i];
      for (var j = 0; j < (sh.rows || []).length; j++) {
        var r = sh.rows[j];
        if (!r.byUser) continue;
        全++;
        if (私.length >= 上限) continue;
        私.push({ 板: sh.name, 行: r });
      }
    }
    var 言 = function (x) {
      var r = x.行;
      var 前 = r.beforeUnknown ? UNKNOWN : String(r.before);
      var 何 = (r.kind === 'formula') ? 'の式' : '';
      return x.板 + '!' + r.addr + 何 + ' を ' + 前 + ' から ' + String(r.after) + ' に しました';
    };
    var 頭;
    if (!私.length) {
      /* ★あなたが 直した 所が 出て いない★（先頭 3行に 入らなかった 等）
           ⇒★無い事に しません★＝数だけ 言います */
      頭 = plan.userCount
        ? ('あなたが 直した所 ' + plan.userCount + 'か所')
        : 'あなたが 直した所は ありません';
    } else {
      頭 = 私.map(言).join('／');
      /* ★★残りは  で 数えます★★（2026-09-22 ここで 1回 踏みました）
            は ★先頭 3行だけ★ です（この 台の 元からの 決め）。
           ⇒見えて いる 分だけ 数えると ★4つ 直したのに 「ほか 1か所」★ に なりました。
           ⇒★数は 必ず 数えた 側（）から 取ります★ */
      var 残 = (plan.userCount || 0) - 私.length;
      if (残 > 0) 頭 += '（ほか ' + 残 + 'か所）';
    }
    /* ★★「うち 別のシートが N か所」まで 言います★★（2026-09-22）
         ★★なぜ★★ ... 経営者1 が 実Excel で 割った 中身（`f876f43`）
           `4月!C4` を 1 ⇒ 99 に すると つられて 4か所。その ★半分の 2か所が 別の 板★。
             同じ 板 ... 4月!E4 ／ 4月!E14
             別の 板 ... まとめ!B4 ／ まとめ!B7
           ⇒★同じ 板だけ 見て いたら 「2か所」と 言う 所でした★
         ⇒★お客さんは 「そこだけ 直した つもり」で 別の 板の 合計が 動きます★
         ⇒★★そこが この 1文の 値打ち★★（2026-08-18 の 裁定＝1つ 直すと 3板 18本）
         ★0の 時は 言いません★＝★字を 長くする 分の 値打ちが 無い★ */
    var 尻 = '';
    if (plan.spreadCount) {
      尻 = '。つられて ' + plan.spreadCount + 'か所が 変わります';
      if (plan.otherSheetCount) 尻 += '（うち 別のシートが ' + plan.otherSheetCount + 'か所）';
    }
    return 頭 + 尻;
  }

  return { build: build, headline: headline, goLabel: goLabel, colName: colName, UNKNOWN: UNKNOWN, 文: 文 };
}));
