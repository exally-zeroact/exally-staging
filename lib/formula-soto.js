/* formula-soto.js — ★外へ 出る 関数の 計算だけ★（2026-09-07）
 *
 *  ★★外へは ここからは 出ません★★
 *    行くのは `lib/formula-soto-plug.js` ＋ `api/soto.js`（★うちの サーバ★）だけ。
 *    ここは ★住所の 組み立て★と ★返って きた 字の 読み取り★だけを する。
 *    ⇒ node で そのまま 試験できる（外へ 1回も 出さずに）
 *
 *  ★★なぜ「うちの サーバを 通す」のか★★
 *    実測（docs/measured/soto/）
 *      ・ブラウザから 撃つと ★4回／4回とも 相手に 届く★（CORS は 読む 方だけ 止める）
 *      ・もらった `.xlsx` の 式は ★式のまま★ 入る＝★開いた だけで 走る★
 *    ⇒★★画面から 直に 外へ 出さない／行き先は うちが 決める★★
 *
 *  見張り: tests/formula-soto.test.mjs
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.FormulaSoto = factory();
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var 誤り = function (種) { return { 誤り: 種 }; };

  /* ════════ 日付（Excel の 数 ↔ 年月日） ════════ */
  function 数から日(n) {
    n = Math.floor(Number(n));
    /* ★通し 60 は 1900-02-29★（`lib/shoshiki.js` の `数から日` と 同じ 決め） */
    if (n === 60) return { y: 1900, m: 2, d: 29 };
    var 基 = (n < 61) ? Date.UTC(1899, 11, 31) : Date.UTC(1899, 11, 30);
    var d = new Date(基 + n * 86400000);
    return { y: d.getUTCFullYear(), m: d.getUTCMonth() + 1, d: d.getUTCDate() };
  }
  function 日から数(y, m, d) {
    /* ★うるう日の 分は ★引く★（足すのは 逆）… 2026-09-16 実測★
         `=DATE(1900,1,1)` … 実Excel ★1★ ／ 足すと ★3★
       ★通し 1～59 は 1899-12-31 起点 ／ 61～ は 1899-12-30 起点★ */
    var n = Math.round((Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000);
    return (n < 61) ? n - 1 : n;
  }
  function 二桁(n) { return (n < 10 ? '0' : '') + n; }
  function 年月日(n) { var p = 数から日(n); return p.y + '-' + 二桁(p.m) + '-' + 二桁(p.d); }

  /* ════════ STOCKHISTORY ════════
     ★行き先は 1つだけ★（stooq の 公開 CSV・鍵も 契約も 要らない）
     ⇒ 住所は ★うちが 組み立てる★＝お客さんの 打った 字を そのまま 渡さない */
  function 銘柄を整える(s) {
    /* ★通すのは 英数字・点・ハイフンだけ★（それ以外は 落とす＝住所に 記号を 載せない）
       ★★長さも 12字までに 切る★★（2026-09-07 指示役の 問いで 足した）
         ⇒ 銘柄の 記号は 長くても 10字ほど
         ⇒★長い 字を そのまま 通すと ★セルの 中身を 少しずつ 外へ 運べる★★
       ★★『塞いだ』のでは なく『★狭めた★』です★★（指示役 2026-09-07）
         ⇒★12字までなら ★12字は 運べます★／式を 並べれば ★12字 × 本数★★
         ⇒★★でも 行き先が stooq に 固定なので
           ★仕込んだ 人が そこの 記録を 読めません★＝★運べても 受け取れない★★★
         ⇒★★だから 行き先を 足す 時が 危ない★★
           ＝★セルから 字が 渡る 関数の 行き先は
             ★仕込んだ 人が 記録を 読めない 相手★に 限る★（api/soto.js の 決まり） */
    var v = String(s == null ? '' : s).trim().toLowerCase().replace(/[^a-z0-9.\-]/g, '');
    return v.slice(0, 12);
  }
  function 株の住所(銘柄, 始, 終) {
    var s = 銘柄を整える(銘柄);
    if (!s) return null;
    /* stooq は 米国株に `.us` を 付ける（付いていなければ 付ける） */
    if (s.indexOf('.') < 0) s += '.us';
    var q = 'https://stooq.com/q/d/l?s=' + s + '&i=d';
    if (始) q += '&d1=' + 年月日(始).replace(/-/g, '');
    if (終) q += '&d2=' + 年月日(終).replace(/-/g, '');
    return q;
  }
  /** ★CSV を 表に する★（stooq の 形＝Date,Open,High,Low,Close,Volume）
   *  @param 何を 0=日付と終値 / 1=+始値 / 2=+高値安値 / 3=+出来高（Excel の properties に 近い 形）
   *  @param 見出し 0=無し 1=有り */
  function 株の表(csv, 何を, 見出し) {
    var 行 = String(csv || '').trim().split(/\r?\n/).filter(function (x) { return x; });
    if (行.length < 2) return 誤り('NA');
    var 頭 = 行[0].split(',').map(function (x) { return x.trim().toLowerCase(); });
    var i = {
      date: 頭.indexOf('date'), open: 頭.indexOf('open'), high: 頭.indexOf('high'),
      low: 頭.indexOf('low'), close: 頭.indexOf('close'), volume: 頭.indexOf('volume'),
    };
    if (i.date < 0 || i.close < 0) return 誤り('NA');
    var 欲 = [['date', '日付'], ['close', '終値']];
    if (何を >= 1) 欲.splice(1, 0, ['open', '始値']);
    if (何を >= 2) { 欲.push(['high', '高値']); 欲.push(['low', '安値']); }
    if (何を >= 3) 欲.push(['volume', '出来高']);
    var 出 = [];
    if (見出し) 出.push(欲.map(function (x) { return x[1]; }));
    for (var r = 1; r < 行.length; r++) {
      var c = 行[r].split(',');
      if (c.length < 2) continue;
      var 一行 = [];
      for (var k = 0; k < 欲.length; k++) {
        var 名 = 欲[k][0], v = c[i[名]];
        if (名 === 'date') {
          var m = String(v || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
          一行.push(m ? 日から数(+m[1], +m[2], +m[3]) : v);
        } else {
          var n = Number(v);
          一行.push(isFinite(n) ? n : '');
        }
      }
      出.push(一行);
    }
    if (!出.length) return 誤り('NA');
    return 出;
  }

  /* ════════ TRANSLATE / DETECTLANGUAGE ════════
     ★行き先は うちの `/api/claude`★＝★外の 会社と 新しい 契約は していません★
     （すでに AI に 聞く 口が 在り、そこを 使う）
     ⇒ ここでは ★頼み文★を 組み立てるだけ */
  function 訳す頼み(文, 元, 先) {
    var f = String(元 || '').trim();
    var t = String(先 || '').trim();
    return '次の 文を ' + (t || '日本語') + ' に 訳してください。'
      + (f ? '（元の 言語は ' + f + ' です）' : '')
      + '★訳した 文だけ★を 返してください。前置きも 説明も 引用符も 付けないでください。\n\n'
      + String(文 == null ? '' : 文);
  }
  function 何語か頼み(文) {
    return '次の 文が 何語で 書かれているかを 見て、'
      + '★言語コードだけ★（例 ja / en / zh-Hans / ko）を 返してください。'
      + '説明は 要りません。\n\n' + String(文 == null ? '' : 文);
  }
  /** ★返事から 言語コードだけ 取り出す★（AI が 前置きを 付けても 落とす） */
  function 言語コードを拾う(返事) {
    var s = String(返事 == null ? '' : 返事).trim();
    var m = s.match(/[a-z]{2,3}(?:-[A-Za-z]{2,8})*/);
    return m ? m[0] : 誤り('VALUE');
  }
  /** ★訳した 文を きれいにする★（前後の 引用符・前置きを 落とす） */
  function 訳を整える(返事) {
    var s = String(返事 == null ? '' : 返事).trim();
    s = s.replace(/^["'「『]+/, '').replace(/["'」』]+$/, '');
    return s;
  }

  /* ════════ WEBSERVICE ════════
     ★住所は お客さんが 打つ★／★でも 出るのは うちが 許した 相手だけ★
     ⇒ 許していない 相手は `api/soto.js` が 断る（ここでは 形だけ 見る） */
  function 住所の形か(u) {
    var s = String(u == null ? '' : u).trim();
    return /^https:\/\/[^\s]+$/i.test(s);
  }

  function 足した名前() {
    return ['WEBSERVICE', 'STOCKHISTORY', 'TRANSLATE', 'DETECTLANGUAGE'];
  }
  function 保留の名前() {
    return ['IMAGE', 'RTD'];
  }

  return {
    数から日: 数から日, 日から数: 日から数, 年月日: 年月日,
    銘柄を整える: 銘柄を整える, 株の住所: 株の住所, 株の表: 株の表,
    訳す頼み: 訳す頼み, 何語か頼み: 何語か頼み,
    言語コードを拾う: 言語コードを拾う, 訳を整える: 訳を整える,
    住所の形か: 住所の形か,
    足した名前: 足した名前, 保留の名前: 保留の名前
  };
}));
