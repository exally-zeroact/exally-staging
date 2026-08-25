/* shindan.js — ★5 E2診断：うちから「ここが危ない」と言う★（純関数・0円）
 *
 *  ★なぜ在るか（司さん／指示役 2026-08-25）★
 *    ・相手（AIの競合）は ★聞かれた事に答える★＝人が「どこが変か」を知っている前提。
 *    ・★人が気づけない物は 一生 聞かれない★。だから ★うちから 出す★。
 *    ・★機械が全セルを見るのは 0円★（AIを1回も呼ばない）。ここが うちの本体。
 *
 *  ★1本目＝「消えた参照(#REF!)が IFERROR で隠れている式」★
 *    ・画面には ★空しか出ない★。エラーの印も出ない。★客は 一生 気づけない★。
 *    ・司さんの実物で ★122本★（2026-08-25 実測）。
 *    ・★競合が見つけられるのは 69個（Excel自身がエラーを出している物）だけ★
 *      ＝この122本は ★Excelが「空」を返すので エラーとして見えない★。
 *
 *  ★言い方の決まり★
 *    ・★客に見せる字に ★ を書かない★（★は うちの覚え書きの印）
 *    ・★「壊れています」と言わない★＝「見えない所で 空になっています」と 何が起きているかを言う
 *    ・★直しは まだ出さない★（順番8＝提案機能）。ここは ★見つけて 場所を出すだけ★。
 *
 *  ★数え方は2通り。混ぜない★
 *    のべ … 1つの式に 2か所 隠れていたら 2
 *    式の本数 … 上を1本と数える ←★客に見せるのは こちら★
 */
(function (root) {
  'use strict';

  /** 列番号 → A, B, …, AA */
  function 列の字(c) {
    var s = '';
    c = Number(c);
    while (c >= 0) { s = String.fromCharCode(65 + (c % 26)) + s; c = Math.floor(c / 26) - 1; }
    return s;
  }
  function セルの名(rc) {
    var p = String(rc).split(',');
    return 列の字(Number(p[1])) + (Number(p[0]) + 1);
  }

  /* ★包まれているかを 本当に見る★
     ＝ /IFERROR/.test(f) && /#REF!/.test(f) では
       「=IFERROR(A1,0) + INDEX(#REF!,1)」のように ★包まれていない物まで 数えてしまう★。
     ⇒ IFERROR( の中身を ★括弧を数えて★ 切り出し、★1つ目の引数の中★に #REF! が在るかを見る。
       （2つ目＝「だめだった時に出す値」の側に #REF! が在っても それは別の話） */
  function 隠している所(f) {
    var 出 = [];
    var s = String(f == null ? '' : f);
    var re = /(IFERROR|IFNA)\s*\(/gi;
    var m;
    while ((m = re.exec(s))) {
      var 名 = m[1].toUpperCase();
      var i = m.index + m[0].length;      // '(' の次
      /* ★字の中の #REF! は 参照ではない★ =IFERROR("#REF!","") は 壊れていない。
         ⇒ ★"…" の中身は 数える側に入れない★（見た目のために 引用符だけ残す） */
      var 深さ = 1, 引数 = '', 文字の中 = false;
      for (; i < s.length; i++) {
        var ch = s.charAt(i);
        if (ch === '"') { 文字の中 = !文字の中; 引数 += ch; continue; }
        if (文字の中) { continue; }
        if (ch === '(') 深さ++;
        else if (ch === ')') { 深さ--; if (!深さ) break; }
        else if (ch === ',' && 深さ === 1) break;   /* ★1つ目の引数の終わり★ */
        引数 += ch;
      }
      var 数 = (引数.match(/#REF!/g) || []).length;
      if (数) 出.push({ 関数: 名, 中身: 引数, 数: 数 });
    }
    return 出;
  }

  /** その式が どこを見に行っていたか（残っている参照）＝直す手掛かり */
  function 手掛かり(f) {
    var s = String(f == null ? '' : f).replace(/#REF!/g, '');
    var 出 = [];
    var re = /(?:'((?:[^']|'')+)'|([^\s!'"(),;:+\-*/^&=<>%]+))!(\$?[A-Z]{1,3}\$?[0-9]{1,7}(?::\$?[A-Z]{1,3}\$?[0-9]{1,7})?)/g;
    var m;
    while ((m = re.exec(s))) {
      出.push((m[1] !== undefined ? m[1].replace(/''/g, "'") : m[2]) + '!' + m[3]);
      if (出.length >= 3) break;
    }
    return 出;
  }

  /**
   * ★調べる（0円・AIを1回も呼ばない）★
   * @param {Array} sheets [{name, data:{'r,c':{v,f,t}}}]
   * @param {Object} opt {一度に:number}
   * @returns {{見つけた:Array, のべ:number, 式の本数:number, 見たセル:number, かかった秒:number}}
   */
  function 調べる(sheets, opt) {
    var it = 調べる途中(sheets, opt);
    var r;
    do { r = it.next(); } while (!r.done);
    return r.value;
  }

  /** ★小分けにする（画面を固めない）★＝ref-graph と同じ流儀 */
  function 調べる途中(sheets, opt) {
    opt = opt || {};
    var 一度に = opt.一度に || 3000;
    /* ★2026-08-25 実測で分かった事★
       小分けは「次の描き直し」を待って動くので、★開いた直後は 画面が別の仕事で忙しい★。
       ブラウザで測ったら ★仕事は 12ミリ秒／出るまでは 5.6秒★＝★待っていた時間★だった。
       ⇒ ★客に言うのは「調べるのにかかった時間（仕事の時間）」★。
          ★出るまでの時間も 別に持つ★（混ぜると どちらも嘘になる）。 */
    var t0 = (typeof performance !== 'undefined' && performance.now) ? performance.now() : Date.now();
    var いま = function () { return (typeof performance !== 'undefined' && performance.now) ? performance.now() : Date.now(); };
    var 仕事 = 0;
    var 見つけた = [], のべ = 0, 見たセル = 0;
    var si = 0, keys = null, ki = 0;
    var done = false;
    function 結果() {
      var t1 = いま();
      var 本 = {};
      for (var i = 0; i < 見つけた.length; i++) 本[見つけた[i].シート + '!' + 見つけた[i].セル] = 1;
      return {
        見つけた: 見つけた,
        のべ: のべ,
        式の本数: Object.keys(本).length,
        見たセル: 見たセル,
        かかった秒: Math.max(0.01, Math.round(仕事 / 10) / 100),   /* ★実際に調べた時間★ */
        出るまでの秒: Math.round((t1 - t0) / 10) / 100,            /* ★待っていた時間も含む★ */
      };
    }
    return {
      next: function () {
        if (done) return { done: true, value: 結果() };
        var 始 = いま();
        var 数 = 0;
        while (数 < 一度に) {
          if (!sheets || si >= sheets.length) { done = true; 仕事 += いま() - 始; return { done: true, value: 結果() }; }
          if (!keys) { keys = Object.keys((sheets[si] && sheets[si].data) || {}); ki = 0; }
          if (ki >= keys.length) { si++; keys = null; continue; }
          var rc = keys[ki++]; 数++; 見たセル++;
          var cell = sheets[si].data[rc];
          var f = cell && cell.f;
          if (!f || String(f).charAt(0) !== '=') continue;   /* ★式だけ見る（打った字は式ではない）★ */
          var 所 = 隠している所(f);
          if (!所.length) continue;
          var n = 0;
          for (var i = 0; i < 所.length; i++) n += 所[i].数;
          のべ += n;
          見つけた.push({
            種類: 'kakureta_ref',
            シート: sheets[si].name,
            セル: セルの名(rc),
            r: Number(String(rc).split(',')[0]),
            c: Number(String(rc).split(',')[1]),
            式: String(f),
            包んでいる関数: 所[0].関数,
            隠している数: n,
            いま出ている物: (cell.v === undefined || cell.v === null || cell.v === '') ? '（空）' : String(cell.v),
            手掛かり: 手掛かり(f),
          });
          continue;
        }
        仕事 += いま() - 始;
        return { done: false, value: null };
      },
    };
  }

  /* ★客に見せる言葉は ここ1か所で作る★（画面に散らさない） */
  var 言い方 = {
    kakureta_ref: {
      題: '見えない所で 空になっています',
      なに: function (n) {
        return '見に行く先が 消えている式が ' + n + 'か所 あります。'
          + 'エラーの印が出ない書き方（IFERROR）で包まれているので、画面には 空が出るだけです。';
      },
      なぜ: 'そのまま合計すると、合計だけが 静かに 小さくなります。',
      つぎ: '場所を出します。中身を見て、消えた先を どこに向けるかを 決めてください。',
    },
  };
  function 言葉(種類, n) {
    var w = 言い方[種類];
    if (!w) return null;
    return { 題: w.題, 本文: w.なに(n) + w.なぜ, つぎ: w.つぎ };
  }

  var api = {
    調べる: 調べる,
    調べる途中: 調べる途中,
    隠している所: 隠している所,
    手掛かり: 手掛かり,
    言葉: 言葉,
    列の字: 列の字,
    セルの名: セルの名,
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.Shindan = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
