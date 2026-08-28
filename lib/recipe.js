/* recipe.js — ★6 レシピ＝手順を覚えて、2回目からは AIを呼ばない★（純関数・0円）
 *
 *  ★決まり（司さん／指示役 2026-08-25〜26）★
 *    ・★AIに送るのは 列名＋見本3行＋機械が作った構造の要約★（客の全データは送らない）
 *    ・★返ってくるのは「手順」★（答えではない）
 *    ・★実行は Exally の中で 全行に 決定論で★／★実行に上限を置かない★
 *    ・★2回目からは AIを呼ばない★＝同じ形の表なら 覚えた手順を そのまま当てる
 *    ・★「前と同じ事」は AIを呼ぶ前に 機械が当てる★
 *      当てられなければ ★普通にAIを呼ぶ★（★黙って外さない★＝当てられなかった事を出す）
 *
 *  ★手順は 決められた形だけ★（AIが何を返しても、この一覧に無い物は 断る）
 *    式の列を足す ／ 列の名前を変える ／ 列を消す ／ 並べ替え
 *  ★断るのは 安全のため★＝AIの返事を そのまま実行しない。
 */
(function (root) {
  'use strict';

  var 列の字 = function (c) {
    var s = '';
    c = Number(c);
    while (c >= 0) { s = String.fromCharCode(65 + (c % 26)) + s; c = Math.floor(c / 26) - 1; }
    return s;
  };

  /* ══ ①機械が 0円で作る「要約」══════════════════════════════
     ★これが AIへ渡す物の全部★（客の中身は 見本3行だけ） */
  /* ★見出しの見つけ方は 借りる（自分で書かない）★
     ＝2つ在ると、片方だけ直った時に ★覚えた物が 当たらなくなる★。 */
  var CH = (typeof require === 'function' && typeof module === 'object')
    ? require('./chizu.js')
    : root.Chizu;
  /* ★並べ替えの順は lib/grid-sort.js が正本★（実Excelを動かして測った真値）＝ここで書かない。
     ★期間の決め方は lib/kikan.js が正本★（Timeally の締め期間を 借りた物）＝ここで書かない。 */
  var GS = (typeof require === 'function' && typeof module === 'object')
    ? require('./grid-sort.js')
    : root.GridSort;
  var KI = (typeof require === 'function' && typeof module === 'object')
    ? require('./kikan.js')
    : root.Kikan;

  function 要約を作る(sheet, opt) {
    opt = opt || {};
    var 見本の行数 = opt.見本の行数 || 3;
    var data = (sheet && sheet.data) || {};
    var 最大行 = -1, 最大列 = -1;
    for (var k in data) {
      if (!Object.prototype.hasOwnProperty.call(data, k)) continue;
      var p = k.split(',');
      var r = Number(p[0]), c = Number(p[1]);
      if (r > 最大行) 最大行 = r;
      if (c > 最大列) 最大列 = c;
    }
    /* ★見出しは 1行目とは限らない★（実物の給料表は 2行目・売上表は 4行目）。
       ★日付や数を 見出しにしない★＝来月には 変わる字を 覚えの鍵にすると 当たらなくなる。 */
    var み = (CH && CH.見出しの行を探す) ? CH.見出しの行を探す(data, 最大列, 5) : { 行: 0, 数: 0 };
    var 見出しの行 = (み.行 >= 0 && み.数 > 0) ? み.行 : 0;
    var 言葉か = (CH && CH.言葉らしいか) ? CH.言葉らしいか : function () { return true; };
    var 列 = [];
    for (var ci = 0; ci <= 最大列; ci++) {
      var セル = data[見出しの行 + ',' + ci];
      var 見出し = 言葉か(セル) ? 値(data, 見出しの行, ci) : '';
      列.push({
        字: 列の字(ci),
        名: (見出し === '' ? '（見出し無し）' : String(見出し)),
        型: 列の型(data, ci, 最大行),
      });
    }
    var 見本 = [];
    for (var ri = 見出しの行 + 1; ri <= Math.min(最大行, 見出しの行 + 見本の行数); ri++) {
      var 行 = [];
      for (var cj = 0; cj <= 最大列; cj++) 行.push(String(値(data, ri, cj)));
      見本.push(行);
    }
    return {
      シート: (sheet && sheet.name) || '',
      行数: 最大行 + 1,
      見出しの行: 見出しの行,
      列数: 最大列 + 1,
      列: 列,
      見本: 見本,
      式の数: 式の数(data),
    };
  }

  function 値(data, r, c) {
    var cell = data[r + ',' + c];
    if (!cell) return '';
    if (cell.v === undefined || cell.v === null) return '';
    return cell.v;
  }
  function 列の型(data, c, 最大行) {
    var 数 = 0, 字 = 0, 日 = 0;
    for (var r = 1; r <= 最大行 && r <= 200; r++) {
      var v = 値(data, r, c);
      if (v === '') continue;
      if (typeof v === 'number') 数++;
      else if (/^\d{4}[-/]\d{1,2}[-/]\d{1,2}/.test(String(v))) 日++;
      else 字++;
    }
    if (数 > 字 && 数 > 日) return '数';
    if (日 > 字 && 日 > 数) return '日付';
    if (字) return '文字';
    return '空';
  }
  function 式の数(data) {
    var n = 0;
    for (var k in data) {
      if (!Object.prototype.hasOwnProperty.call(data, k)) continue;
      var f = data[k] && data[k].f;
      if (f && String(f).charAt(0) === '=') n++;
    }
    return n;
  }

  /* ══ ②「前と同じ事」を AIを呼ぶ前に 当てる ══════════════════
     ★形が同じ＝列の名前と並びが同じ★（行数や中身は 変わってよい＝毎月 中身は変わる） */
  function 指紋(要約) {
    if (!要約 || !要約.列) return '';
    var 名 = [];
    for (var i = 0; i < 要約.列.length; i++) 名.push(要約.列[i].名);
    return 名.join('');
  }
  function 同じ形か(A, B) {
    return !!A && !!B && 指紋(A) === 指紋(B) && 指紋(A) !== '';
  }
  /**
   * ★覚えた手順の中から 今の表に当たる物を探す★
   * @returns {{見つかった:boolean, レシピ:Object|null, なぜ:string}}
   *   ★見つからなかった理由を 必ず返す（黙って外さない）★
   */
  function 当ててみる(要約, 覚えた物) {
    var 一覧 = 覚えた物 || [];
    if (!一覧.length) return { 見つかった: false, レシピ: null, なぜ: 'まだ 覚えた手順が ありません' };
    for (var i = 一覧.length - 1; i >= 0; i--) {
      if (同じ形か(要約, 一覧[i].要約)) {
        return { 見つかった: true, レシピ: 一覧[i], なぜ: '前と 同じ形の表です（列の名前と並びが 同じ）' };
      }
    }
    return { 見つかった: false, レシピ: null, なぜ: '列の名前か 並びが 前と違います' };
  }

  /* ══ ③AIへ渡す物（★小さい★）══════════════════════════════ */
  function AIへ渡す物(要約, 頼み) {
    var 行 = [];
    行.push('次の表について、やる事の「手順」だけを JSON で返してください。答えの値は返さないでください。');
    行.push('');
    行.push('# 表');
    行.push('シート: ' + 要約.シート + '（' + 要約.行数 + '行 × ' + 要約.列数 + '列・式 ' + 要約.式の数 + '本）');
    行.push('列: ' + 要約.列.map(function (c) { return c.字 + '=' + c.名 + '(' + c.型 + ')'; }).join(' / '));
    行.push('見本(先頭' + 要約.見本.length + '行): ');
    for (var i = 0; i < 要約.見本.length; i++) 行.push('  ' + 要約.見本[i].join(' | '));
    行.push('');
    行.push('# やりたい事');
    行.push(String(頼み || ''));
    行.push('');
    行.push('# 返し方（この形以外は 使いません）');
    行.push('{"手順":[ {"種類":"式の列を足す","見出し":"…","式":"=A{行}*2"},');
    行.push('          {"種類":"列の名前を変える","元":"A","新":"…"},');
    行.push('          {"種類":"列を消す","列":"C"},');
    行.push('          {"種類":"並べ替え","列":"B","向き":"昇順"},');
    行.push('          {"種類":"切り出す","人":"（見出しの名前）","月":"2026-01","始まりの日":1} ]}');
    行.push('切り出す … その人の列と 日付の行だけを 新しいシートに出します。');
    行.push('始まりの日 … 1なら その月まるごと。10なら 前の月の10日〜その月の9日。');
    行.push('式の {行} は その行の番号に置き換えます。');
    return 行.join('\n');
  }

  /* ══ ④AIの返事を「手順」として読む（★決めた形以外は 断る★）══ */
  var 使える種類 = ['式の列を足す', '列の名前を変える', '列を消す', '並べ替え', '切り出す'];
  function 手順を読む(text) {
    var s = String(text == null ? '' : text);
    var m = s.match(/\{[\s\S]*\}/);
    if (!m) return { ok: false, 手順: [], なぜ: '手順が 返ってきませんでした' };
    var o;
    try { o = JSON.parse(m[0]); } catch (e) { return { ok: false, 手順: [], なぜ: '手順の形が 読めませんでした' }; }
    var 生 = (o && o.手順) || [];
    if (!生.length) return { ok: false, 手順: [], なぜ: '手順が 1つも ありませんでした' };
    var 出 = [], 断った = [];
    for (var i = 0; i < 生.length; i++) {
      var t = 生[i] || {};
      if (使える種類.indexOf(t.種類) < 0) { 断った.push(String(t.種類)); continue; }
      if (t.種類 === '式の列を足す') {
        if (!t.見出し || !t.式 || String(t.式).charAt(0) !== '=') { 断った.push('式の列を足す（中身が足りない）'); continue; }
      }
      if (t.種類 === '列の名前を変える' && (!t.元 || !t.新)) { 断った.push('列の名前を変える（中身が足りない）'); continue; }
      if (t.種類 === '列を消す' && !t.列) { 断った.push('列を消す（中身が足りない）'); continue; }
      if (t.種類 === '並べ替え' && !t.列) { 断った.push('並べ替え（中身が足りない）'); continue; }
      出.push(t);
    }
    if (!出.length) return { ok: false, 手順: [], なぜ: '使える手順が ありませんでした（' + 断った.join('・') + '）', 断った: 断った };
    return { ok: true, 手順: 出, なぜ: '', 断った: 断った };
  }

  /* ══ ⑤手順を 全行に 決定論で 当てる ══════════════════════
     ★変える所を返すだけ★＝実際に書き換えるのは 画面（差分プレビューを通す） */
  function 手順を当てる(sheet, 手順, opt) {
    opt = opt || {};
    var data = (sheet && sheet.data) || {};
    /* ★見出しの行は 1行目とは限らない★（実物の給料表は 2行目）。
       ★実際に押して 見つけた（2026-08-27）★＝見出しを 1行目（表の題）に書き、
         本当の見出しの行に 式を書いていた。⇒ 見出しの行を 探してから 当てる。 */
    var 見出しの行 = opt.見出しの行;
    if (見出しの行 === undefined || 見出しの行 === null) {
      var 端 = -1;
      for (var k0 in data) {
        if (!Object.prototype.hasOwnProperty.call(data, k0)) continue;
        var c0 = Number(k0.split(',')[1]);
        if (c0 > 端) 端 = c0;
      }
      var み = (CH && CH.見出しの行を探す) ? CH.見出しの行を探す(data, 端, 5) : { 行: 0, 数: 0 };
      見出しの行 = (み.行 >= 0 && み.数 > 0) ? み.行 : 0;
    }
    var 最大行 = -1, 最大列 = -1;
    for (var k in data) {
      if (!Object.prototype.hasOwnProperty.call(data, k)) continue;
      var p = k.split(',');
      if (Number(p[0]) > 最大行) 最大行 = Number(p[0]);
      if (Number(p[1]) > 最大列) 最大列 = Number(p[1]);
    }
    var 変える = {};
    var なに = [];
    var 出すシート = [];
    var 次の列 = 最大列 + 1;
    for (var i = 0; i < 手順.length; i++) {
      var t = 手順[i];
      if (t.種類 === '式の列を足す') {
        var c = 次の列++;
        変える[見出しの行 + ',' + c] = { v: String(t.見出し) };
        for (var r = 見出しの行 + 1; r <= 最大行; r++) {
          変える[r + ',' + c] = { f: String(t.式).replace(/\{行\}/g, String(r + 1)) };
        }
        なに.push('列 ' + 列の字(c) + ' に「' + t.見出し + '」を足す（' + (最大行 - 見出しの行) + '行）');
      } else if (t.種類 === '列の名前を変える') {
        var ci = 列の番号(t.元);
        if (ci < 0) continue;
        変える[見出しの行 + ',' + ci] = { v: String(t.新) };
        なに.push('列 ' + t.元 + ' の名前を「' + t.新 + '」にする');
      } else if (t.種類 === '列を消す') {
        var cj = 列の番号(t.列);
        if (cj < 0) continue;
        for (var r2 = 0; r2 <= 最大行; r2++) 変える[r2 + ',' + cj] = { v: '' };
        なに.push('列 ' + t.列 + ' を空にする（' + (最大行 + 1) + '行）');
      } else if (t.種類 === '並べ替え') {
        /* ★2026-08-27 指示役の指摘★＝ここは 名前を出すだけで ★1セルも 動かしていなかった★
           （「何も見ていないのに緑」に近い形）。⇒ ★実際に 並べ替える★。
           ★順は lib/grid-sort.js が正本★（実Excelを動かして測った真値）＝ここで くらべ方を書かない。 */
        var ck = 列の番号(t.列);
        if (ck < 0 || !GS || !GS.order) continue;
        var 頭 = 見出しの行 + 1;
        if (最大行 < 頭) continue;
        var 元の行 = [];
        for (var r3 = 頭; r3 <= 最大行; r3++) 元の行.push(r3);
        var 鍵 = 元の行.map(function (rr) { return data[rr + ',' + ck] || { v: '' }; });
        var 順 = GS.order(鍵, (t.向き === '降順') ? 'desc' : 'asc');
        var 動いた = 0;
        for (var q = 0; q < 順.length; q++) {
          var 元 = 元の行[順[q]], 先 = 頭 + q;
          if (元 === 先) continue;
          動いた++;
          for (var cc = 0; cc <= 最大列; cc++) {
            var もと = data[元 + ',' + cc];
            /* ★式は そのまま運ぶ（値に潰さない）★＝運んだ先で 計算し直す */
            変える[先 + ',' + cc] = もと
              ? (もと.f ? { f: もと.f } : { v: (もと.v === undefined ? '' : もと.v) })
              : { v: '' };
          }
        }
        なに.push('列 ' + t.列 + ' で ' + (t.向き || '昇順') + 'に並べ替える（'
          + 元の行.length + '行のうち ' + 動いた + '行が 動く）');
      } else if (t.種類 === '切り出す') {
        /* ★切り出しは 別のシートに出す★＝元の表は 1セルも 触らない（原本を触らない）。
           ここでは 何を出すかを 決めるだけ。作るのは 画面（見せてから）。 */
        var 出 = 切り出す(sheet, t, { 見出しの行: 見出しの行 });
        if (!出) continue;
        出すシート.push(出);
        なに.push(出.なに);
      }
    }
    return { 変える: 変える, なに: なに, 何行: 最大行 - 見出しの行, 見出しの行: 見出しの行,
      出すシート: 出すシート };
  }
  /* ══ ★切り出す（1人分 × 何月分）★ ══════════════════════════
     ★司さんの言葉（2026-08-27）★
       「今 同じシートに複数人 あるやろ？ ワンクリックで 1人分と 何月分って 分けて表示もさせれるん？」
     ★実物の形（読んで 数えた）★
       給料表 … 1行目=表の題／2行目=見出し（日付・白石正人・長野孝…）／3行目〜=日付と金額
       ⇒ ★1人分＝列を1本 抜く★（行を絞るのではない）
       ⇒ ★何月分＝日付の行を 期間で絞る★
     ★期間の決め方は lib/kikan.js（Timeally が正本）★／★見出しは lib/chizu.js の道具★
     ★元の表は 1セルも 触らない★＝新しいシートの中身を 返すだけ。 */
  /** ★セルの本当の値★（★実際に押して 見つけた 2026-08-27★）
   *  画面では ★式のセルは v が空★で、計算した値は d に入る（例 {v:'', f:'=D1', d:'46023'}）。
   *  v だけを見ていたので ★実物の給料表で「日付の列が 見つかりません」と言って 外れた★。
   *  ⇒ ★v が空なら d を見る★（部品の試験だけでは 出なかった＝画面で押して 出た穴）。 */
  function セルの値(cell) {
    if (!cell) return '';
    if (cell.v !== undefined && cell.v !== null && cell.v !== '') return cell.v;
    if (cell.d !== undefined && cell.d !== null && cell.d !== '') return cell.d;
    return '';
  }

  /** ★日付らしい値か★（★小さい数を 日付にしない★）
   *  ＝Excelの通し番号は 1 でも「1900-01-01」になる。1・2・3 が並ぶ ★点数や 個数の列★を
   *    日付の列と 間違えた（実測 2026-08-28）。
   *  ⇒ ★20000（1954年）より 前の通し番号は 日付として 数えない★。
   *    字で書かれた日付（'2026-01-05'）は そのまま 日付として 数える。 */
  function 日付らしいか(v) {
    if (!KI) return false;
    var ymd = KI.日付に直す(v);
    if (!ymd) return false;
    if (typeof v === 'number' || /^\d+(\.\d+)?$/.test(String(v).trim())) {
      return Number(v) >= 20000;
    }
    return true;
  }

  /** ★日付の列は「近い日が 並んでいる」★（★金額の列を 日付にしない★）
   *  ★実測 2026-08-28★＝給料明細のシートで 56,647円 や 7,475円 が
   *    「2055年」「1920年」の日付として 読まれ、★別のシートに 切り替えの帯が 出た★。
   *  ⇒ ★一番古い日と 一番新しい日の 差が 3年より広い列は 日付の列にしない★
   *     （台帳・勤怠・売上は どれも 数年の中に 収まる）。 */
  var 日付の列の決まり = { 何個以上: 3, 何日まで: 1100 };

  function 日付の列を探す(data, 見出しの行, 最大行, 最大列) {
    var 一番 = -1, 一番の数 = 0;
    var 見る = Math.min(最大行, 見出しの行 + 30);
    for (var c = 0; c <= 最大列; c++) {
      var 数 = 0, 小 = null, 大 = null;
      for (var r = 見出しの行 + 1; r <= 見る; r++) {
        var cell = data[r + ',' + c];
        if (!cell) continue;
        var v = セルの値(cell);
        if (!日付らしいか(v)) continue;
        数++;
        var n = Number(v);
        if (!isFinite(n)) {
          var ymd = KI && KI.日付に直す(v);
          n = ymd ? Math.round(Date.parse(ymd + 'T00:00:00Z') / 86400000) + 25569 : null;
        }
        if (n === null || !isFinite(n)) continue;
        if (小 === null || n < 小) 小 = n;
        if (大 === null || n > 大) 大 = n;
      }
      /* ★散らばりが 広い列は 日付ではない（金額が 並んでいる）★ */
      if (小 !== null && (大 - 小) > 日付の列の決まり.何日まで) continue;
      if (数 > 一番の数) { 一番の数 = 数; 一番 = c; }
    }
    /* ★1つ2つ 当たっただけの列を 日付の列にしない★（3つ以上 並んでいる事） */
    if (一番の数 < 日付の列の決まり.何個以上) return { 列: -1, 数: 0 };
    return { 列: 一番, 数: 一番の数 };
  }

  /**
   * @param {Object} sheet   元のシート（★触らない★）
   * @param {Object} t       {種類:'切り出す', 人:'白石正人', 月:'2026-01', 始まりの日:1, 出す先:'…'}
   * @returns {{name,data,何行,何列,なに,元の行数,出した行数,期間,人の列,なぜ}|null}
   */
  function 切り出す(sheet, t, opt) {
    opt = opt || {};
    var data = (sheet && sheet.data) || {};
    var 最大行 = -1, 最大列 = -1;
    for (var k in data) {
      if (!Object.prototype.hasOwnProperty.call(data, k)) continue;
      var p = k.split(',');
      if (Number(p[0]) > 最大行) 最大行 = Number(p[0]);
      if (Number(p[1]) > 最大列) 最大列 = Number(p[1]);
    }
    var 見出しの行 = opt.見出しの行;
    if (見出しの行 === undefined || 見出しの行 === null) {
      var み = (CH && CH.見出しの行を探す) ? CH.見出しの行を探す(data, 最大列, 5) : { 行: 0, 数: 0 };
      見出しの行 = (み.行 >= 0 && み.数 > 0) ? み.行 : 0;
    }
    /* ★人の列は 見出しの字で 探す★（★同じ名前が2人 居たら 断る＝どちらか 分からない） */
    var 当たり = [];
    var 名 = String(t.人 == null ? '' : t.人).trim();
    for (var c2 = 0; c2 <= 最大列; c2++) {
      var h = data[見出しの行 + ',' + c2];
      var hv = h ? String(h.v === undefined ? '' : h.v).trim() : '';
      if (名 && hv === 名) 当たり.push(c2);
    }
    if (!名) return { なぜ: '誰の分かが 決まっていません' };
    if (!当たり.length) return { なぜ: '「' + 名 + '」という見出しの列が 見つかりません' };
    if (当たり.length > 1) {
      return { なぜ: '「' + 名 + '」が ' + 当たり.length + '列 あります（どちらの分か 決められません）' };
    }
    var 人の列 = 当たり[0];

    /* ★日付の列★（頼まれた列が在れば それ・無ければ 探す） */
    var 日付の列 = (t.日付の列 !== undefined && t.日付の列 !== null)
      ? 列の番号(String(t.日付の列)) : -1;
    if (日付の列 < 0) {
      var ひ = 日付の列を探す(data, 見出しの行, 最大行, 最大列);
      日付の列 = ひ.列;
      if (日付の列 < 0 || ひ.数 === 0) return { なぜ: '日付の列が 見つかりません' };
    }

    var 期間 = (t.月 && KI) ? KI.期間を決める(t.月, t.始まりの日 === undefined ? 1 : t.始まりの日) : null;
    if (t.月 && !期間) return { なぜ: 'いつの分かが 読めません' };

    /* ★行を 期間で絞る（上限を置かない＝何行でも 全部 見る）★ */
    var 出す行 = [], 元の行数 = 0;
    for (var r2 = 見出しの行 + 1; r2 <= 最大行; r2++) {
      var dc = data[r2 + ',' + 日付の列];
      var ymd = dc ? KI.日付に直す(セルの値(dc)) : null;
      if (!ymd) continue;                       /* 日付でない行（合計の行など）は 出さない */
      元の行数++;
      if (期間 && !KI.期間の中か(ymd, 期間)) continue;
      出す行.push({ 行: r2, 日付: ymd });
    }
    if (!出す行.length) {
      return { なぜ: (期間 ? 期間.言い方 + 'に 当たる日が 1日も ありません' : '出せる行が ありません') };
    }

    /* ★新しいシートの中身★（日付の列 と その人の列 だけ） */
    var 出 = {};
    var 元の見出し = data[見出しの行 + ',' + 日付の列];
    出['0,0'] = { v: 元の見出し ? セルの値(元の見出し) : '日付' };
    出['0,1'] = { v: 名 };
    for (var i2 = 0; i2 < 出す行.length; i2++) {
      var src = 出す行[i2];
      var a = data[src.行 + ',' + 日付の列];
      var b = data[src.行 + ',' + 人の列];
      /* ★式は 運ばない（切り出した先で 指す所が 変わる）★＝計算した値を そのまま置く。
         ★金額は うちが計算した値★（d が在れば d・無ければ v）。 */
      出[(i2 + 1) + ',0'] = { v: 数なら数に(セルの値(a)), numFmt: a && a.numFmt };
      出[(i2 + 1) + ',1'] = { v: 数なら数に(セルの値(b)), numFmt: b && b.numFmt };
    }
    /* ★シートの名前は 客が読む字★（01月分 ではなく 1月分）。Excelの上限は31文字。
       ★「1日から」と「10日から」で 同じ名前にしない★＝実物で 2つとも「白石正人 1月分」に
       なっていた（2026-08-27 実測）。まるごと1か月でない時は 始まりの日を 付ける。
       ★シート名に使えない字（/ \ ? * [ ] :）は 入れない★ */
    var 期間の字 = '';
    if (期間) {
      期間の字 = ' ' + (+期間.ym.slice(5, 7)) + '月分'
        + (期間.締め日 >= 31 ? '' : '（' + ((期間.締め日 % 31) + 1) + '日から）');
    }
    var 名前 = String(t.出す先 || (名 + 期間の字)).replace(/[\\/?*\[\]:]/g, '').slice(0, 31);
    return {
      name: 名前,
      data: 出,
      何行: 出す行.length,
      何列: 2,
      期間: 期間,
      人の列: 列の字(人の列),
      元の行数: 元の行数,
      出した行数: 出す行.length,
      なに: '「' + 名 + '」の ' + (期間 ? 期間.言い方 : 'ぜんぶ') + ' を 新しいシートに出す（'
        + 出す行.length + '行／元は ' + 元の行数 + '行）',
    };
  }

  /** ★数の字は 数に戻す★＝画面の d は 字（'46023'）。字のまま置くと 合計できない。 */
  function 数なら数に(v) {
    if (typeof v === 'number') return v;
    var s = String(v == null ? '' : v).trim();
    if (s === '') return '';
    if (/^-?\d+(\.\d+)?$/.test(s)) return Number(s);
    return v;
  }

  function 列の番号(字) {
    var s = String(字 || '').toUpperCase();
    if (!/^[A-Z]{1,3}$/.test(s)) return -1;
    var c = 0;
    for (var i = 0; i < s.length; i++) c = c * 26 + (s.charCodeAt(i) - 64);
    return c - 1;
  }

  /* ══ ⑥覚える形（倉庫にも 画面にも 同じ形で置く）══ */
  /* ══ ★前と同じ事か（AIを呼ぶ前に 機械が当てる）★ ══════════
     ★形が同じ だけでは 足りない★＝同じ表に 別の事を頼まれる方が 多い。
     ★頼みの字も 同じ時だけ★「前と同じ事」と言う。違えば ★黙って外さず 理由を言って★ 普通に呼ぶ。 */
  function 頼みの形(s) {
    return String(s == null ? '' : s)
      .replace(/[\s　]/g, '')
      .replace(/[。．、，!！?？]/g, '')
      .toLowerCase();
  }
  /**
   * @returns {{やる:boolean, レシピ:Object|null, なぜ:string}}
   *   ★なぜ は いつも 返す（当てられなかった事を 客に言うため）★
   */
  function 前と同じか(頼み, 要約, 覚えた物) {
    var 一覧 = 覚えた物 || [];
    if (!一覧.length) return { やる: false, レシピ: null, なぜ: 'まだ 覚えた手順が ありません' };
    var t = 頼みの形(頼み);
    if (!t) return { やる: false, レシピ: null, なぜ: '頼みが 空です' };
    var 形は同じ = false;
    for (var i = 一覧.length - 1; i >= 0; i--) {
      var R = 一覧[i];
      var 同形 = 同じ形か(要約, R.要約);
      if (同形) 形は同じ = true;
      if (同形 && 頼みの形(R.頼み) === t) {
        return { やる: true, レシピ: R, なぜ: '前と 同じ頼み・同じ形の表です' };
      }
    }
    return {
      やる: false, レシピ: null,
      なぜ: 形は同じ ? '同じ形の表ですが、前と違う頼みです' : '列の名前か 並びが 覚えた表と違います',
    };
  }

  function レシピを作る(名, 頼み, 要約, 手順) {
    return {
      名: String(名 || 頼み || '手順'),
      頼み: String(頼み || ''),
      要約: { シート: 要約.シート, 列: 要約.列, 列数: 要約.列数 },   /* ★見本(客の中身)は 覚えない★ */
      手順: 手順,
      指紋: 指紋(要約),
    };
  }

  var api = {
    要約を作る: 要約を作る,
    指紋: 指紋,
    同じ形か: 同じ形か,
    当ててみる: 当ててみる,
    前と同じか: 前と同じか,
    頼みの形: 頼みの形,
    AIへ渡す物: AIへ渡す物,
    手順を読む: 手順を読む,
    手順を当てる: 手順を当てる,
    切り出す: 切り出す,
    /* ★切り替え（lib/kirikae.js）にも 同じ物を使わせる★＝2つ作らない */
    セルの値: セルの値,
    日付の列を探す: 日付の列を探す,
    日付らしいか: 日付らしいか,
    日付の列の決まり: 日付の列の決まり,
    レシピを作る: レシピを作る,
    列の字: 列の字,
    列の番号: 列の番号,
    使える種類: 使える種類,
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.Recipe = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
