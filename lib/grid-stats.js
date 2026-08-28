/* grid-stats.js — ★選んだ所の 合計・平均・個数（ステータスバー）★（純関数）
 *
 *  ★実Excel 16.0.20228 を COM で押して測った（2026-08-21）★
 *    置いた物 → SUM / AVERAGE / COUNT / COUNTA / MIN / MAX を そのまま読んだ。
 *
 *    ① 10,20,30            … 合計60 平均20 数値の個数3 データの個数3 最小10 最大30
 *    ② ★文字だけ あ,い,う★ … 合計0 ★平均は出ない★ 数値の個数0 データの個数3
 *    ③ 10,あ,20            … 合計30 ★平均15（文字は 分母に入らない＝÷2）★ 数値2 データ3
 *    ④ 10,(空),20          … データの個数2 ★空は数えない★
 *    ⑤ ★TRUE,FALSE,10★     … 合計10 平均10 ★数値の個数1★ データの個数3
 *                            ＝★セルの中の TRUE/FALSE は 数として扱わない★
 *    ⑥ ★'10,'20（数字の形の文字）★ … ★合計0★ 平均出ない 数値の個数0 データの個数2
 *                            ＝★見た目が数でも 文字なら 足されない★
 *    ⑦ ★10,#DIV/0!,20★     … ★合計も平均も 出ない（エラーになる）★ 数値2 データ3
 *
 *  ★未測定★
 *    ステータスバーに ★既定でどの3つが出るか★ は COM からもレジストリからも読めなかった。
 *    （HKCU\…\Excel\StatusBar には 変えた物しか残らない＝既定は残らない）
 *    うちは ★平均・データの個数・数値の個数・合計★ を出す事にした（理由は下）。
 *
 *  ★うちの方が良い所（Excel は黙る）★
 *    ⑥ が実務で一番 刺さる＝★見た目が数なのに 合計に入らない★。
 *    Excel は 黙って小さい合計を出す。うちは ★「数の形の文字が◯個 在ります（足していません）」★ と出す。
 *    ⑦ も同じで、★エラーが◯個 在るので 合計は出せません★ と ★理由を出す★。
 */
(function (root) {
  'use strict';

  /** 画面に出ている字（式なら答え）＝印刷と同じ見方 */
  function shown(cell) {
    if (!cell) return '';
    var v = (cell.d !== undefined && cell.d !== null && cell.d !== '') ? cell.d : cell.v;
    if (v === undefined || v === null) return '';
    return String(v);
  }

  function isError(text) {
    return /^#(DIV\/0!|N\/A|NAME\?|NULL!|NUM!|REF!|VALUE!|SPILL!|CALC!|GETTING_DATA)$/.test(String(text));
  }

  /* ★計算できていない値★
     うちの計算は `=1/0` や `=1/A99` で ★Infinity★ を返す（2026-08-21 実測）。
     実Excel は #DIV/0!。★これを 黙って合計に入れない★（入れると合計が Infinity になる）。 */
  function isNotComputed(text) {
    return /^(-?Infinity|NaN)$/.test(String(text).trim());
  }

  /** ★先頭の ' は「これは文字」の印（Excel と同じ）★
      うちの SUM も '20 を足さない（実測）ので 帯も足さない。
      ★同じ物を 2通りで数えない★ */
  function isForcedText(cell) {
    return !!(cell && typeof cell.f === 'string' && cell.f.charAt(0) === "'");
  }

  /** ★セルが「数」かどうか★＝生の値が数の時だけ（文字の '10 は数ではない＝実測⑥） */
  function numberOf(cell) {
    if (!cell) return null;
    if (isForcedText(cell)) return null;             /* ★'20 は文字（うちの SUM と揃える）★ */
    var raw = cell.v;
    if (typeof raw === 'number' && isFinite(raw)) return raw;
    /* 式のセルは 答え(d) を見る。答えが数の形なら 数として足す（Excel も足す） */
    if (cell.f !== undefined && cell.f !== null && cell.f !== '') {
      var d = cell.d;
      if (typeof d === 'number' && isFinite(d)) return d;
      if (typeof d === 'string' && d !== '' && !isError(d) && !isNotComputed(d)) {
        var n = Number(d.replace(/,/g, ''));
        if (!isNaN(n) && /^[\s\-+]*[\d.,]+\s*$/.test(d)) return n;
      }
    }
    return null;
  }

  /** ★見た目は数なのに 文字として入っている★（合計に入らない＝黙って小さくなる元） */
  function looksLikeNumberButText(cell) {
    if (!cell) return false;
    if (numberOf(cell) !== null) return false;
    var raw = cell.v;
    if (typeof raw !== 'string') return false;
    var s = raw.trim();
    if (s.charAt(0) === "'") s = s.slice(1);          /* ★'20 も 「数の形の文字」★ */
    if (s === '' || s.charAt(0) === '=') return false;
    if (isError(s)) return false;
    return /^[+-]?[\d,]+(\.\d+)?$/.test(s);
  }

  /**
   * 選んだ所を まとめる
   * @param {Array} cells セルの配列（空マスは null / undefined で渡してよい）
   * @returns {Object}
   */
  function summarize(cells) {
    cells = cells || [];
    var データの個数 = 0, 数値の個数 = 0, 合計 = 0, 最小 = null, 最大 = null;
    var 数の形の文字 = 0, エラー = 0, 計算できていない = 0;

    for (var i = 0; i < cells.length; i++) {
      var cell = cells[i];
      var t = shown(cell);
      var 空 = (t === '') && (!cell || cell.v === undefined || cell.v === null || cell.v === '');
      if (空) continue;                              /* ★空は数えない（実測④）★ */
      データの個数++;
      if (isError(t)) { エラー++; continue; }
      if (isNotComputed(t)) { 計算できていない++; continue; }   /* ★Infinity/NaN を 合計に入れない★ */
      if (looksLikeNumberButText(cell)) 数の形の文字++;
      var n = numberOf(cell);
      if (n === null) continue;                      /* 文字・論理値は 数として扱わない（実測⑤⑥） */
      数値の個数++;
      合計 += n;
      if (最小 === null || n < 最小) 最小 = n;
      if (最大 === null || n > 最大) 最大 = n;
    }

    /* ★エラーが1つでも在れば 合計も平均も 出さない（実測⑦）★ */
    var 数を出せる = (エラー === 0) && (計算できていない === 0) && (数値の個数 > 0);
    return {
      データの個数: データの個数,
      数値の個数: 数値の個数,
      合計: 数を出せる ? 合計 : null,
      平均: 数を出せる ? (合計 / 数値の個数) : null,
      最小: 数を出せる ? 最小 : null,
      最大: 数を出せる ? 最大 : null,
      数の形の文字: 数の形の文字,
      エラー: エラー,
      計算できていない: 計算できていない,
      出せない理由: エラー > 0 ? ('エラーが ' + エラー + '個 在るので 合計は出せません')
        : 計算できていない > 0 ? ('計算できていない値が ' + 計算できていない + '個 在るので 合計は出せません')
        /* ★1つのセルだけ 選んでいる時は 言わない★（実Excelも 言わない）
           ＝表を切り替えた画面で「数が1つも在りません」が ずっと出ていて、
             ★見えている物と 合っていない★ と 読めた（2026-08-28 指示役）。 */
        : (データの個数 > 1 && 数値の個数 === 0 ? '数が1つも在りません' : ''),
    };
  }

  /** 数を 見せる形にする（けた区切り／小数は多くしない） */
  function fmt(n) {
    if (n === null || n === undefined) return '';
    var s = (Math.abs(n) < 1e15 && Math.round(n) !== n) ? String(Math.round(n * 1e10) / 1e10) : String(n);
    var p = s.split('.');
    p[0] = p[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    return p.join('.');
  }

  /** ★平均は割り切れない事が多い★… 小数は4桁までにし、
      ★丸めた時は 丸めたと分かるよう ≒ を付ける★（黙って丸めない）。
      ★未測定★：実Excel の帯が 何桁まで出すかは COM から読めなかった。 */
  function fmtAvg(n) {
    if (n === null || n === undefined) return '';
    var r = Math.round(n * 10000) / 10000;
    return (r === n ? '' : '≒') + fmt(r);
  }

  /**
   * ステータスバーに出す並び（★出せる物だけ★を返す＝空の枠を出さない）
   * @returns {Array<{名:string, 値:string, 注意:boolean}>}
   */
  function items(s) {
    var 普通 = [], 注意 = [];
    if (!s || s.データの個数 === 0) return [];
    /* ★黙って小さくならないように 理由を必ず出す★ */
    if (s.数の形の文字 > 0) {
      注意.push({ 名: '⚠️ 数の形の文字', 値: s.数の形の文字 + '個（足していません）', 注意: true });
    }
    if (s.出せない理由) 注意.push({ 名: '⚠️', 値: s.出せない理由, 注意: true });

    if (s.平均 !== null) 普通.push({ 名: '平均', 値: fmtAvg(s.平均), 注意: false });
    普通.push({ 名: 'データの個数', 値: fmt(s.データの個数), 注意: false });
    if (s.数値の個数 !== s.データの個数) 普通.push({ 名: '数値の個数', 値: fmt(s.数値の個数), 注意: false });
    if (s.合計 !== null) 普通.push({ 名: '合計', 値: fmt(s.合計), 注意: false });

    /* ★注意を 先に出す★
       375px の実機で測ったら 帯が 679px になり、後ろに置いた注意が
       ★画面の外で 読めなかった★（2026-08-21 実測）。★DOMに在る≠読める★ */
    return 注意.concat(普通);
  }


  var api = { summarize: summarize, items: items, fmt: fmt, fmtAvg: fmtAvg, shown: shown, isError: isError, numberOf: numberOf };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.GridStats = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
