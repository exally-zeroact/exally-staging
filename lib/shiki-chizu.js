/* shiki-chizu.js — ★頼りの 地図（どの 式が どの 式に 頼って いるか）★（2026-09-11）
 *
 *  ★★土台を 自分で 作る ③枚目★★
 *    土台は 5つ … ①字に 切る ②形に する ③★頼りの 地図★ ④順番と 計算 ⑤溢れ
 *    ①=`lib/shiki-kiru.js` ②=`lib/shiki-katachi.js` 参照=`lib/shiki-sansho.js`
 *
 *  ★★何の 為か★★
 *    A1 が B1 を 見て いて、B1 が C1 を 見て いるなら
 *    ★C1 → B1 → A1 の 順に 計算しないと 古い 答えが 出ます★
 *    この 台は ★その 順番を 出します★（計算は まだ しません＝④）
 *
 *  ★★借り物の 中は 1文字も 読んで いません★★
 *
 *  ★★大事な 決め事★★
 *    ★頼りは「式の マス」だけ 結びます★
 *      A1:A3 の 中に 式が 無ければ そこは ただの 数＝★順番を 考える 必要が ありません★
 *      ⇒だから `A:A`（104万マス）でも ★広げずに 済みます★
 *
 *  ★輪（ぐるぐる回り）は 勝手に 直しません★＝名指しで 返します
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(
    typeof require === 'function' ? require('./shiki-kiru.js') : root.ShikiKiru,
    typeof require === 'function' ? require('./shiki-katachi.js') : root.ShikiKatachi,
    typeof require === 'function' ? require('./shiki-sansho.js') : root.ShikiSansho);
  else root.ShikiChizu = factory(root.ShikiKiru, root.ShikiKatachi, root.ShikiSansho);
})(typeof self !== 'undefined' ? self : this, function (切, 形, 参) {
  'use strict';

  function 札(板, 行, 列) { return 板 + '!' + 行 + ',' + 列; }

  /* ══ ★式の マスの 置き場（列ごとに 行を 並べて おく）★ ══
     ★なぜ 列ごとか★＝`A:A` の ような 広い 参照でも
     ★式が 在る 列・在る 行だけ 見れば 済む★（104万マスを 数えない） */
  function 棚を作る(板たち) {
    var 棚 = {};
    for (var i = 0; i < 板たち.length; i++) {
      var b = 板たち[i], 列ごと = {};
      for (var k in b.式) {
        if (!Object.prototype.hasOwnProperty.call(b.式, k)) continue;
        var t = k.split(','), r = +t[0], c = +t[1];
        (列ごと[c] || (列ごと[c] = [])).push(r);
      }
      for (var c2 in 列ごと) 列ごと[c2].sort(function (a, b2) { return a - b2; });
      棚[b.名] = { 列ごと: 列ごと, 列の札: Object.keys(列ごと).map(Number) };
    }
    return 棚;
  }

  /* ★並んだ 行から「上〜下」に 入る 物を 取り出す★（頭から 数えない＝二分探索） */
  function 挟まれた行(並, 上, 下) {
    var lo = 0, hi = 並.length;
    while (lo < hi) { var m = (lo + hi) >> 1; if (並[m] < 上) lo = m + 1; else hi = m; }
    var 出 = [];
    for (var i = lo; i < 並.length && 並[i] <= 下; i++) 出.push(並[i]);
    return 出;
  }

  /* ★四角の 中に 居る「式の マス」を 拾う★ */
  function 四角の中の式(棚, 板名, 四) {
    var t = 棚[板名];
    if (!t) return [];
    var 出 = [];
    /* ★式の 在る 列だけ 回る★＝`A:A` でも `1:1` でも 端から 端まで 数えません */
    for (var j = 0; j < t.列の札.length; j++) {
      var c = t.列の札[j];
      if (c < 四.左 || c > 四.右) continue;
      /* ★行は 二分探索★＝104万行の 参照でも 行を 1本ずつ 見ません（ここが 一番 効く） */
      var rs = 挟まれた行(t.列ごと[c], 四.上, 四.下);
      for (var i = 0; i < rs.length; i++) 出.push(札(板名, rs[i], c));
    }
    return 出;
  }

  /* ══ ★形の 木を 歩いて 参照を 拾う★ ══
     ★`:` で 繋がった 物は まとめて 1つの 四角★（A1:B2 を A1 と B2 に 割らない） */
  function 参照を拾う(木, 既定の板, 出, 名前) {
    if (!木 || typeof 木 !== 'object') return;
    if (木.種 === '名' || (木.種 === '二' && 木.記 === ':')) {
      var 字 = 形.字に戻す(木);
      var r = 参.読む(字 == null ? '' : String(字).trim());
      if (r && r.ok) {
        出.push({ 板: r.板 || 既定の板, 上: r.上, 左: r.左, 下: r.下, 右: r.右 });
        return;
      }
      if (木.種 === '名') { 名前.push(字); return; }
      /* ★`:` なのに 読めない★＝中を 1つずつ 見る（`A1:INDEX(...)` の ような 物） */
    }
    if (木.種 === '呼') { for (var i = 0; i < 木.引数.length; i++) 参照を拾う(木.引数[i], 既定の板, 出, 名前); return; }
    if (木.種 === '配') { for (var a = 0; a < 木.行.length; a++) for (var b = 0; b < 木.行[a].length; b++) 参照を拾う(木.行[a][b], 既定の板, 出, 名前); return; }
    if (木.左) 参照を拾う(木.左, 既定の板, 出, 名前);
    if (木.右) 参照を拾う(木.右, 既定の板, 出, 名前);
    if (木.子) 参照を拾う(木.子, 既定の板, 出, 名前);
  }

  /**
   * ★地図を 作る★
   * @param 板たち [{ 名:'Sheet1', 式:{ '行,列': '=A2+1', … } }]（行・列は ★0から★）
   * @returns { ok, 頼り, 逆, 順番, 輪, 読めず, 名前 }
   *   頼り … その マスが ★見て いる★ 式の マス
   *   逆   … その マスを ★見て いる★ 式の マス
   *   順番 … ★この 順に 計算すれば 古い 答えが 出ない★
   *   輪   … ぐるぐる 回って いて 順番が 決まらない マス（★勝手に 直しません★）
   */
  function 地図を作る(板たち) {
    var 棚 = 棚を作る(板たち);
    var 頼り = {}, 逆 = {}, 読めず = [], 名前 = {};
    var 全部 = [];

    for (var i = 0; i < 板たち.length; i++) {
      var b = 板たち[i];
      for (var k in b.式) {
        if (!Object.prototype.hasOwnProperty.call(b.式, k)) continue;
        var t = k.split(','), 私 = 札(b.名, +t[0], +t[1]);
        全部.push(私);
        頼り[私] = [];
        var 字 = String(b.式[k] || '');
        if (字.charAt(0) === '=') 字 = 字.slice(1);
        var か = 切.切る(字);
        if (!か.ok) { 読めず.push({ マス: 私, 訳: か.なぜ || '切れない' }); continue; }
        var 木 = 形.形にする(か.出);
        if (!木.ok) { 読めず.push({ マス: 私, 訳: 木.なぜ || '形に ならない' }); continue; }
        var 四たち = [], 名 = [];
        参照を拾う(木.形, b.名, 四たち, 名);
        if (名.length) 名前[私] = 名;
        var 済 = {};
        for (var j = 0; j < 四たち.length; j++) {
          var 先 = 四角の中の式(棚, 四たち[j].板, 四たち[j]);
          for (var m = 0; m < 先.length; m++) {
            if (済[先[m]]) continue;
            /* ★自分が 自分を 見て いる 時も そのまま 結びます★
               ⇒順番が 決まらない＝★輪として 名指しで 返る★（実Excel も 断ります） */
            済[先[m]] = 1;
            頼り[私].push(先[m]);
            (逆[先[m]] || (逆[先[m]] = [])).push(私);
          }
        }
      }
    }

    /* ══ ★順番を 決める★ ══
       ★頼って いる 先が 全部 済んだ 物から 並べる★（残った 物＝輪） */
    var 残り = {}, 順番 = [], 待ち = [];
    for (var n = 0; n < 全部.length; n++) 残り[全部[n]] = 頼り[全部[n]].length;
    for (var n2 = 0; n2 < 全部.length; n2++) if (残り[全部[n2]] === 0) 待ち.push(全部[n2]);
    while (待ち.length) {
      var 今 = 待ち.shift();
      順番.push(今);
      var 見て = 逆[今] || [];
      for (var p = 0; p < 見て.length; p++) {
        if (--残り[見て[p]] === 0) 待ち.push(見て[p]);
      }
    }
    var 輪 = [];
    if (順番.length !== 全部.length) {
      for (var q = 0; q < 全部.length; q++) if (残り[全部[q]] > 0) 輪.push(全部[q]);
    }

    return { ok: 読めず.length === 0, 頼り: 頼り, 逆: 逆, 順番: 順番, 輪: 輪,
      読めず: 読めず, 名前: 名前, 札: 札 };
  }

  /** ★1マス 変えた時に 計算し直す 物だけ 出す★（全部 やり直さない）
   *  ★返る 順は 地図の 順番と 同じ★＝古い 答えが 出ない */
  function 変えたら(地図, 変えたマス) {
    var 要る = {}, 待ち = [].concat(変えたマス);
    while (待ち.length) {
      var 今 = 待ち.shift();
      var 見て = 地図.逆[今] || [];
      for (var i = 0; i < 見て.length; i++) {
        if (要る[見て[i]]) continue;
        要る[見て[i]] = 1; 待ち.push(見て[i]);
      }
    }
    var 出 = [];
    for (var j = 0; j < 地図.順番.length; j++) if (要る[地図.順番[j]]) 出.push(地図.順番[j]);
    return 出;
  }

  return { 地図を作る: 地図を作る, 変えたら: 変えたら, 札: 札 };
});
