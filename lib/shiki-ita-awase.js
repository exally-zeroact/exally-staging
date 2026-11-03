/* shiki-ita-awase.js — ★台の 板と 画面の `data` が 合って いるかを 数える★（2026-09-18）
 *
 *  ★★なぜ 要るか★★
 *    ㋑（計算を 台に 回す）では ★台が 板を 1つ 持ちます★。
 *    ＝★マスを 打つ たびに 台にも 打つ★（`book.html` の 5か所）
 *    ⇒★★1つ 忘れると ★黙って 古い 値★が 出ます★★
 *      ＝★誤りに ならない／画面も 崩れない／数だけ 違う★
 *      ＝★一番 見つけにくい 形★
 *    ⇒★★だから 先に 門を 作ります★★（★書く 前に★）
 *
 *  ★★3つとも 出します★★（経営者1 の 注文）
 *    ・★式の 数★（`data` の 側 ／ 台の 側）
 *    ・★値の 数★（同上）
 *    ・★★食い違った マスの 数★★
 *
 *  ★★この 台が 見て いない 事★★
 *    ・★答え（`d`）は 見て いません★＝★打った 物（`f` / `v`）だけ★
 *      ＝★答えが 違うのは 別の 話★（計算の 話／ここは ★写しの 話★）
 *    ・★書式・罫線・判子は 見て いません★
 *    ・★溢れで 埋まった マスは 数えません★（★打った 物では ない★）
 *
 *  ★★呼び方★★
 *    const 出 = ShikiItaAwase.写しを数える(板, data);
 *    ⇒ { 式: {紙, 台}, 値: {紙, 台}, 食い違い: 0, 中身: [...] }
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.ShikiItaAwase = factory();
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  /** ★行・列（0から）を マスの 名前に★（A1 の 形） */
  function 番地から名(行, 列) {
    var s = '', n = 列 + 1;
    while (n > 0) { var t = (n - 1) % 26; s = String.fromCharCode(65 + t) + s; n = (n - t - 1) / 26; }
    return s + (行 + 1);
  }

  /** ★打った 物を 字に する★（★答えでは なく 打った 物★）
   *    ★式は `f`／値は `v`★（`book.html` の `data` の 形）
   *    ★空は 空の 字★ */
  function 打った物(セル) {
    if (!セル) return '';
    if (セル.f !== undefined && セル.f !== null && セル.f !== '') return String(セル.f);
    if (セル.v !== undefined && セル.v !== null) return String(セル.v);
    return '';
  }

  /**
   * ★★台の 板と 画面の `data` を 数える★★
   *  ★★名前を わざと 長く して います★★（2026-09-18）
   *    ＝`数える` は ★repo に 既に 5つ 在ります★
   *      （ctx-menu.js／ribbon.js／ribbon-keytips.js／ribbon-launcher.js／ribbon-context-spec.js）
   *    ＝`tools/unused-param.mjs` は ★名前で 呼ぶ 形★も 見ますので、
   *      ここで `数える(...)` と 書くと ★向こうの 5つが 「渡されている」ように 見え★、
   *      ★見張りの 候補を 狂わせます★（★私は 実際に 狂わせました★）
   *    ⇒★repo に 在る 名前を そのまま 使わない★
   *      （★`lib/shiki-hyou.js` に 同じ 覚書が 在り、私は それを 読んで いませんでした★）
   * @param 板 `ShikiHyou.表()` が 返した 物（`打った字` を 持つ）
   * @param data `book.html` の `sheets[i].data`（鍵は "行,列"）
   * @returns {{式:{紙:number,台:number}, 値:{紙:number,台:number}, 食い違い:number, 中身:Array}}
   */
  /** ★板の 名前を 頭に 付ける★（★1冊の 板に 何枚も 入れる 為★・2026-09-18）
   *    ★なぜ 要るか★ ... `='5月'!E14` の ような ★板またぎの 式★は
   *      ★板が 1枚ずつだと 読めません★（#REF! に なる）
   *      ⇒★`tests/cross-sheet.test.mjs` が 34万円 少ない 形で 捕まえました★
   *    ⇒★1冊の 板に 全部 入れ、名前を `板名!A1` に します★ */
  function 冠をつける(板名, 名) {
    return 板名 ? (String(板名) + '!' + 名) : 名;
  }
  /** ★台の 鍵から 板の 名前と 芯を 割る★（鍵は `板1!A1` の 形） */
  function 鍵を割る(鍵) {
    var i = String(鍵).lastIndexOf('!');
    if (i < 0) return { 板: '', 芯: String(鍵) };
    return { 板: String(鍵).slice(0, i), 芯: String(鍵).slice(i + 1) };
  }
  function 同じ板か(a, b) {
    return String(a || '').toUpperCase() === String(b || '').toUpperCase();
  }

  function 写しを数える(板, data, 板名) {
    var 出 = { 式: { 紙: 0, 台: 0 }, 値: { 紙: 0, 台: 0 }, 食い違い: 0, 中身: [] };
    if (!板 || !data) { 出.読めない = true; return 出; }
    var 見た = {};

    /* ── ①画面の 側から 見る ── */
    for (var k in data) {
      if (!Object.prototype.hasOwnProperty.call(data, k)) continue;
      var 番 = k.split(',');
      var r = Number(番[0]), c = Number(番[1]);
      if (!(r >= 0) || !(c >= 0)) continue;
      var 名 = 番地から名(r, c);
      見た[名.toUpperCase()] = true;
      var 紙字 = 打った物(data[k]);
      if (紙字 === '') continue;
      if (紙字.charAt(0) === '=') 出.式.紙++; else 出.値.紙++;

      var 台字 = '';
      try { 台字 = String(板.打った字 ? (板.打った字(冠をつける(板名, 名)) || '') : ''); }
      catch (e) { 台字 = '読めません'; }
      if (台字.charAt(0) === '=') 出.式.台++; else if (台字 !== '') 出.値.台++;

      if (台字 !== 紙字) {
        出.食い違い++;
        if (出.中身.length < 20) 出.中身.push({ マス: 名, 画面: 紙字, 台: 台字 });
      }
    }

    /* ── ②台の 側に だけ 在る 物 ──
         ★★台の 鍵は `板1!A1` の 形です★★（★板の 名前が 頭に 付きます★）
           ＝`h.中身` の 鍵を 見て 確かめました（2026-09-18）
           ＝★`打った字('A1')` は 板の 名前 無しで 引けます★
         ⇒★鍵から 板の 名前を 外して 比べます★
         ★1枚の 板だけ 見ます★（★`book.html` は 板ごとに 板を 1つ 持つ 積もり★） */
    var 台の名 = [];
    try {
      var 中 = 板.中身 || {};
      for (var kk in 中) {
        if (!Object.prototype.hasOwnProperty.call(中, kk)) continue;
        var 割 = 鍵を割る(kk);
        /* ★他の 板の マスは 見ません★（★1冊に 何枚も 入って います★） */
        if (板名 ? !同じ板か(割.板, 板名) : false) continue;
        台の名.push(割.芯);
      }
    } catch (e) { 台の名 = []; }
    for (var i = 0; i < 台の名.length; i++) {
      var n2 = 台の名[i];
      if (見た[n2.toUpperCase()]) continue;
      var t2 = '';
      try { t2 = String(板.打った字 ? (板.打った字(冠をつける(板名, n2)) || '') : ''); } catch (e) { t2 = ''; }
      if (t2 === '') continue;
      if (t2.charAt(0) === '=') 出.式.台++; else 出.値.台++;
      出.食い違い++;
      if (出.中身.length < 20) 出.中身.push({ マス: n2, 画面: '(無い)', 台: t2 });
    }
    return 出;
  }

  /**
   * ★★画面の `data` に 合わせて 板を 直す★★（★食い違った マスだけ★）
   *
   *  ★★なぜ この 形か★★（★測ってから 決めました★）
   *    ㋐★打つ たびに 台にも 打つ★ ... `book.html` の 書く所は ★37か所★（実測）
   *       ⇒★1つ 忘れると 黙って 古い 値★＝★一番 見つけにくい 形★
   *    ㋑★聞く たびに 板を 建て直す★ ... ★16,000式で 637ms★（`ita-wa-omoi-ka.mjs`）
   *       ⇒★1打鍵ごとには 使えません★（司さんの 実物は 15,799式）
   *    ⇒★★㋒＝1か所で 写しを 合わせる（★食い違った マスだけ 打ち直す★）★★
   *       ＝比べるのは ★字と 字★（★読み直しは 違う マスだけ★）
   *       ＝★忘れる 所が 37 → 1 に なります★
   *
   *  @returns {{打ち直した:number, 消した:number, 見た:number}}
   */
  function 合わせる(板, data, 板名) {
    var 出 = { 打ち直した: 0, 消した: 0, 見た: 0 };
    if (!板 || !data) { 出.読めない = true; return 出; }
    var 見た = {};

    for (var k in data) {
      if (!Object.prototype.hasOwnProperty.call(data, k)) continue;
      var 番 = k.split(',');
      var r = Number(番[0]), c = Number(番[1]);
      if (!(r >= 0) || !(c >= 0)) continue;
      var 名 = 冠をつける(板名, 番地から名(r, c));
      見た[名.toUpperCase()] = true;
      出.見た++;
      var 紙字 = 打った物(data[k]);
      var 台字 = '';
      try { 台字 = String(板.打った字 ? (板.打った字(名) || '') : ''); } catch (e) { 台字 = ''; }
      if (台字 === 紙字) continue;              /* ★同じ 字なら 読み直しません★ */
      板.打つ(名, 紙字);
      if (紙字 === '') 出.消した++; else 出.打ち直した++;
    }

    /* ★台にだけ 残った 物を 消す★（★画面から 消えた マス★） */
    var 中 = 板.中身 || {};
    var 消す = [];
    for (var kk in 中) {
      if (!Object.prototype.hasOwnProperty.call(中, kk)) continue;
      var 割 = 鍵を割る(kk);
      /* ★他の 板の マスは 消しません★（★1冊に 何枚も 入って います★） */
      if (板名 ? !同じ板か(割.板, 板名) : false) continue;
      var n2 = 冠をつける(板名, 割.芯);
      if (見た[n2.toUpperCase()]) continue;
      var t2 = '';
      try { t2 = String(板.打った字 ? (板.打った字(n2) || '') : ''); } catch (e) { t2 = ''; }
      if (t2 === '') continue;
      消す.push(n2);
    }
    for (var j = 0; j < 消す.length; j++) { 板.打つ(消す[j], ''); 出.消した++; }
    return 出;
  }

  return { 写しを数える: 写しを数える, 合わせる: 合わせる, 番地から名: 番地から名, 打った物: 打った物,
    冠をつける: 冠をつける };
}));
