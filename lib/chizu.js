/* chizu.js — ★7 地図＝AIに毎回 渡す「ブック全体の見取り図」★（純関数・0円）
 *
 *  ★正本（2026-08-22 司さん決定・zeroact-memory/projects/exally/decisions.md）★
 *    1.★地図を毎回 渡す（数千文字）★ … シート一覧と役割／表の名前／
 *      ★どのシートが どこを 何本 参照しているか★／大きさ
 *    2.★参照の網で辿る★（別シート・別ファイルを跨いでも）
 *    3.★AIが「もっと見せて」と言える★（→ lib/horu.js）
 *    4.★地図とよく使う所は 置いたまま使い回す★（prompt caching・導入済み）
 *
 *  ★動かせない事実★
 *    ★司さんの実物1冊で 約142万文字＝AIには入らない★（100万トークンでも足りない）。
 *    ⇒ ★「ブック全部を毎回渡す」は 誰にも出来ない★。勝負は ★要る所を どれだけ確実に見つけるか★。
 *
 *  ★核心＝気づくのは機械。AIは説明するだけ★
 *    参照の網も 診断も ★0円・AIを1回も呼ばずに★ 出来ている。
 *    地図は ★その結果の要約★＝★数千文字★。だから 大きいファイルほど うちが強い。
 *
 *  ★お金の数字は 地図に書かない★（合計・請求額・給与）
 *    ＝AIに 金額を覚えさせない。金額は ★うちの計算（HyperFormula）の値だけ★を使う。
 */
(function (root) {
  'use strict';

  var LF = String.fromCharCode(10);
  var 既定の上限 = 4000;   /* ★数千文字★（司さんの決定）。超えたら 削って「出していない」と書く */

  function 列の字(c) {
    var s = '';
    c = Number(c);
    while (c >= 0) { s = String.fromCharCode(65 + (c % 26)) + s; c = Math.floor(c / 26) - 1; }
    return s;
  }

  /** そのシートの 大きさ・式の本数・見出し（1行目）
   *  ★数えている物の名前★（2026-08-27 指示役の指摘で 1行 書く）
   *    ★値か式が在る所★ … data に入っているセルの いちばん端（うちが数えた物）
   *    ★表の枠★         … Excel が持っている「使っている範囲」(!ref)。★こちらの方が 広い事が在る★
   *      実物 … 計算シート 値か式 400行×72列 ／ ★表の枠 404行×152列★
   *              給料表     値か式 468行×9列  ／ ★表の枠 468行×131列★
   *    ⇒ ★AIは「どこまで在るか」で 掘る範囲を決める★ので ★両方 出す★（片方だけだと 見落とす） */
  function シートの姿(sheet, 見出しの数) {
    var data = (sheet && sheet.data) || {};
    var 最大行 = -1, 最大列 = -1, 式 = 0;
    for (var k in data) {
      if (!Object.prototype.hasOwnProperty.call(data, k)) continue;
      var p = k.split(',');
      var r = Number(p[0]), c = Number(p[1]);
      if (r > 最大行) 最大行 = r;
      if (c > 最大列) 最大列 = c;
      var f = data[k] && data[k].f;
      if (f && String(f).charAt(0) === '=') 式++;
    }
    var 見出し = [];
    for (var ci = 0; ci <= 最大列 && 見出し.length < (見出しの数 || 12); ci++) {
      var cell = data['0,' + ci];
      var v = (cell && cell.v !== undefined && cell.v !== null) ? String(cell.v) : '';
      if (v === '') continue;
      /* ★見出しに 数が入っていたら 値を書かない★
         ＝1行目が 見出しとは限らない（金額が入っている本が 実際に在った）。
         ★お金の数字を AIに覚えさせない★（合計・請求額・給与） */
      if (typeof (cell && cell.v) === 'number' || /^[\d,.\-¥￥]+$/.test(v)) {
        見出し.push(列の字(ci) + '=（数）');
        continue;
      }
      見出し.push(列の字(ci) + '=' + v);
    }
    var 枠 = (sheet && sheet.枠) || null;
    return {
      名: (sheet && sheet.name) || '',
      行数: 最大行 + 1, 列数: 最大列 + 1,      /* 値か式が在る所 */
      枠: 枠,                                  /* 表の枠(!ref)・無ければ null＝未測定 */
      式: 式, 見出し: 見出し,
    };
  }

  /** 長い字を 端で切る（★何を切ったかは 呼ぶ側が まとめて書く★） */
  function 短く(s, n) {
    s = String(s || '');
    return (s.length <= n) ? s : (s.slice(0, n) + '…');
  }

  /** ★どのシートが どこを 何本 参照しているか★（網から数える・0円） */
  function 参照のまとめ(網, sheets) {
    var 出 = {};
    var 別ファイル = {};
    if (!網) return { 行: [], 別ファイル: [] };
    var 一覧 = 網.別シート参照 || [];
    for (var i = 0; i < 一覧.length; i++) {
      var x = 一覧[i];
      var 元 = (sheets[x.from] && sheets[x.from].name) || String(x.from);
      /* ★to は シートの番号★（名前ではない）。
         ★自分で地図を読んで見つけた（2026-08-27）★＝「給料表→10」と 番号が出ていて、
         ★地図の いちばん大事な行が AIに読めない★状態だった。⇒ 名前に直す。 */
      var 先 = (typeof x.to === 'number' && sheets[x.to]) ? sheets[x.to].name : String(x.to);
      var k = 元 + '→' + 先;
      出[k] = (出[k] || 0) + 1;
    }
    var f一覧 = 網.別ファイル参照 || [];
    for (var j = 0; j < f一覧.length; j++) {
      var b = f一覧[j].book || f一覧[j].to || '';
      if (b) 別ファイル[b] = (別ファイル[b] || 0) + 1;
    }
    var 行 = Object.keys(出).map(function (k2) { return { 道: k2, 本数: 出[k2] }; });
    行.sort(function (a, b2) { return b2.本数 - a.本数; });
    return { 行: 行, 別ファイル: Object.keys(別ファイル).map(function (n) { return { 名: n, 本数: 別ファイル[n] }; }) };
  }

  /**
   * ★地図を作る（0円・AIを1回も呼ばない）★
   * @param {Array} sheets
   * @param {Object} 網      lib/ref-graph.js の結果（無くてもよい＝その時は「未測定」と書く）
   * @param {Object} 診断    lib/shindan.js の結果（無くてもよい）
   * @param {{上限:number, 今のシート:string}} opt
   * @returns {{字:string, 文字数:number, 削った:number, 中身:Object}}
   */
  function 作る(sheets, 網, 診断, opt) {
    opt = opt || {};
    var 上限 = opt.上限 || 既定の上限;
    sheets = sheets || [];
    var 姿 = sheets.map(function (s) { return シートの姿(s); });
    var まとめ = 参照のまとめ(網, sheets);
    var 行 = [];
    /* ★地図に ★ を書かない★＝★は うちの覚え書きの印（客にも AIにも 出さない） */
    行.push('# このブックの地図（機械が0円で作った物。AIは これを見て 場所を決める）');
    行.push('シート数 ' + sheets.length
      + ' ／ 式 ' + 姿.reduce(function (a, x) { return a + x.式; }, 0) + '本'
      + (opt.今のシート ? ' ／ 今 見ているシート: ' + opt.今のシート : ''));
    行.push('');
    行.push('## シート');
    for (var i = 0; i < 姿.length; i++) {
      var s = 姿[i];
      /* ★枠の方が広い時は 必ず両方 出す★（AIが 73列目から先を 一生 掘らない事故を止める） */
      var 大きさ = s.行数 + '行×' + s.列数 + '列';
      if (s.枠 && (s.枠.行数 > s.行数 || s.枠.列数 > s.列数)) {
        大きさ = '値か式 ' + 大きさ + '／表の枠 ' + s.枠.行数 + '行×' + s.枠.列数 + '列';
      }
      行.push('- ' + s.名 + '（' + 大きさ + '・式' + s.式 + '本）'
        + (s.見出し.length ? ' 見出し: ' + s.見出し.join(' / ') : ' 見出し: （無し）'));
    }
    行.push('');
    行.push('## どのシートが どこを 何本 見ているか');
    if (!網) {
      行.push('（未測定 … まだ調べ終わっていません）');
    } else if (!まとめ.行.length) {
      行.push('（別のシートを見ている式は 0本）');
    } else {
      for (var j = 0; j < まとめ.行.length; j++) {
        行.push('- ' + まとめ.行[j].道 + ' … ' + まとめ.行[j].本数 + '本');
      }
    }
    if (まとめ.別ファイル.length) {
      行.push('');
      行.push('## 別のファイルを見ている所');
      for (var k = 0; k < まとめ.別ファイル.length; k++) {
        行.push('- ' + まとめ.別ファイル[k].名 + ' … ' + まとめ.別ファイル[k].本数 + '本');
      }
    }
    if (網 && (網.解けない || []).length) {
      行.push('');
      行.push('## その場で決まる参照（INDIRECT/OFFSET など・行き先は 開くまで分からない）');
      行.push('- ' + 網.解けない.length + '本');
    }
    if (診断 && 診断.式の本数) {
      行.push('');
      行.push('## 機械が見つけた 危ない所（AIは呼んでいない）');
      行.push('- 見に行く先が消えていて IFERROR で空になっている式 … ' + 診断.式の本数 + 'か所');
      var 先頭 = (診断.見つけた || []).slice(0, 5).map(function (x) { return x.シート + '!' + x.セル; });
      if (先頭.length) 行.push('  例) ' + 先頭.join(' / ')
        + ((診断.見つけた || []).length > 先頭.length ? ' ほか ' + ((診断.見つけた || []).length - 先頭.length) + 'か所' : ''));
    }
    行.push('');
    行.push('## 決まり（守ってください）');
    行.push('- 中身が要る時は 自分で書かず「もっと見せて」で 範囲を頼んでください。');
    行.push('- 金額・合計は 自分で計算して書かないでください。{{シート名!セル}} と書けば こちらの計算結果を入れます。');
    行.push('- 答えには どのセルを見て言ったか を必ず書いてください。');

    /* ★削る順（2026-08-27 指示役の指摘で 直した）★
       前は 後ろから切っていた＝★決まり（守ってください）が いちばん先に消える★。
       ★決まりが消えると AIは 金額を自分で書き、見た所も書かない★＝いちばん危ない削り方だった。
       ⇒ ★先に消してよいのは シートの見出し★／★決まりは 最後まで 必ず残す★。 */
    var LF = String.fromCharCode(10);
    var 決まりから = 行.indexOf('## 決まり（守ってください）');
    var 決まり = (決まりから >= 0) ? 行.slice(決まりから) : [];
    var 本体 = (決まりから >= 0) ? 行.slice(0, 決まりから) : 行.slice();
    var 決まりの字 = 決まり.join(LF);
    var 元の長さ = 行.join(LF).length;
    var 一言 = '';
    var 使える = 上限 - 決まりの字.length - 2;
    for (var 段 = 0; 段 < 3; 段++) {
      if (本体.join(LF).length + 一言.length <= 使える) break;
      if (段 < 2) {
        for (var n = 0; n < 本体.length; n++) {
          var l = 本体[n];
          var pp = l.indexOf(' 見出し: ');
          if (pp > 0) 本体[n] = (段 === 0) ? (l.slice(0, pp + 6) + 短く(l.slice(pp + 6), 40)) : l.slice(0, pp);
        }
      } else {
        while (本体.length > 4 && 本体.join(LF).length + 一言.length > 使える) {
          var 消す = -1;
          for (var mm = 本体.length - 1; mm >= 0; mm--) if (本体[mm].charAt(0) === '-') { 消す = mm; break; }
          if (消す < 0) break;
          本体.splice(消す, 1);
        }
      }
      一言 = LF + '…（地図が長いので 一部を 出していません。要る所は「もっと見せて」で頼んでください）';
    }
    var 字 = 本体.join(LF) + 一言 + (決まりの字 ? (LF + 決まりの字) : '');
    /* ★最後の砦★＝それでも超えるなら 本体だけ 端で切る（★決まりは 必ず残す★） */
    if (字.length > 上限) {
      var 残り = 上限 - 決まりの字.length - 40;
      if (残り < 0) 残り = 0;
      字 = 本体.join(LF).slice(0, 残り) + LF + '…（ここから先は 出していません）' + (決まりの字 ? (LF + 決まりの字) : '');
    }
    var 削った = Math.max(0, 元の長さ - 字.length);
    return {
      字: 字,
      文字数: 字.length,
      削った: 削った,
      中身: { シート: 姿, 参照: まとめ.行, 別ファイル: まとめ.別ファイル },
    };
  }

  var api = { 作る: 作る, シートの姿: シートの姿, 参照のまとめ: 参照のまとめ, 列の字: 列の字, 既定の上限: 既定の上限 };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.Chizu = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
