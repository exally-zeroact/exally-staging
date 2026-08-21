/* grid-sort.js — ★並べ替え（純関数）★
 *
 *  ★実Excel 16.0.20228 を COM で動かして測った真値（2026-08-21）★
 *
 *  ① 並ぶ順（昇順）… 10/あ/2/空/TRUE/#DIV!0/B/1 を並べると
 *        ★1 | 2 | 10 | B | あ | TRUE | #DIV/0! | （空）★
 *     ＝ ★数 → 文字 → 論理値 → エラー → ★空白は いちばん最後★★
 *  ② 降順 … ★#DIV/0! | TRUE | あ | B | 10 | 2 | 1 | （空）★
 *     ＝ 逆順にするが ★空白は 降順でも いちばん最後★
 *  ③ 表の広がり … B3 で 1セルだけ選んで測ると ★A1:C4★（空の行と列で囲まれた かたまり）。
 *     離れて1つだけ在る E1 は ★E1 だけ★。
 *  ④ 見出しの自動判定（Excel の xlGuess）
 *        先頭が文字・下が数 → ★見出しとみなす★（「なまえ」が動かない）
 *        全部 数            → ★見出しではない★（30 が動いた）
 *        全部 文字          → ★見出しではない★（「う」が動いた）
 *     ＝ ★先頭行が文字で、2行目から下に数が在る時だけ 見出し★
 *  ⑤ 行はセットで動く … 2列目をキーに A1:B4 を並べ替えると
 *        名前|ろ|は|い ／ 数|1|2|3 ＝ ★名前と数がバラバラにならない★
 *  ⑥ ★式は行ごと運ばれる★
 *        =A1*10 のような ★相対参照だけの式★ は 並べ替えても 各行で正しいまま
 *        =$A$1 のような ★$付きの式★ は ★指す先が変わる（値が変わる）★
 *        （実測: =$A$1|=$A$2|=$A$3 → =$A$2|=$A$3|=$A$1 ／ 答えが 3|1|2 → 2|3|1）
 *        ＝ Excel 自身がそうなる。うちも同じにする。ただし ★黙って値が変わる★ ので
 *          $付きの式が在った時だけ 一言 知らせる（画面側の仕事）。
 */
(function (root) {
  'use strict';

  var 数 = 0, 文字 = 1, 論理 = 2, エラー = 3, 空 = 4;

  /** そのセルが「何の仲間」か＋比べる値（★実Excelの並び順に合わせた★） */
  function 型と値(cell) {
    if (!cell) return { t: 空, v: null };
    var raw = (cell.d !== undefined && cell.d !== null && cell.d !== '') ? cell.d
      : (cell.v !== undefined && cell.v !== null && cell.v !== '') ? cell.v : null;
    if (raw === null || String(raw).trim() === '') return { t: 空, v: null };
    var s = String(raw).trim();
    if (s.charAt(0) === '#') return { t: エラー, v: s };
    var u = s.toUpperCase();
    if (u === 'TRUE' || u === 'FALSE') return { t: 論理, v: u === 'TRUE' ? 1 : 0 };
    var n = Number(s.replace(/,/g, ''));
    if (s !== '' && !isNaN(n) && isFinite(n)) return { t: 数, v: n };
    return { t: 文字, v: s };
  }

  function くらべる(a, b) {
    if (a.t !== b.t) return a.t - b.t;
    if (a.t === 数 || a.t === 論理) return a.v - b.v;
    return a.v < b.v ? -1 : a.v > b.v ? 1 : 0;
  }

  /**
   * 並べる順を決める（★中身は動かさない。元の番号の並びを返すだけ★）
   * @param {Array<Object>} cells キーの列のセル（上から順）
   * @param {'asc'|'desc'} dir
   * @returns {number[]} 新しい並び（元のindex）
   */
  function order(cells, dir) {
    var 印 = cells.map(function (c, i) { return { i: i, k: 型と値(c) }; });
    var 中身 = 印.filter(function (x) { return x.k.t !== 空; });
    var 空の物 = 印.filter(function (x) { return x.k.t === 空; });   /* ★空白は常に最後★ */
    中身.sort(function (a, b) {
      var d = くらべる(a.k, b.k);
      if (d !== 0) return dir === 'desc' ? -d : d;
      return a.i - b.i;                                             /* 同じなら元の順を保つ */
    });
    return 中身.map(function (x) { return x.i; }).concat(空の物.map(function (x) { return x.i; }));
  }

  /**
   * 1セルだけ選んでいる時の「表」の広がり（★Excel の CurrentRegion と同じ考え方★）
   * @param {function(number,number):boolean} 中身がある (r,c) → 空でなければ true
   * @param {number} r,c 起点
   * @param {number} maxR,maxC 見に行く上限
   */
  function region(中身がある, r, c, maxR, maxC) {
    if (!中身がある(r, c)) return { r1: r, c1: c, r2: r, c2: c };
    var r1 = r, r2 = r, c1 = c, c2 = c;
    var 行に何かある = function (rr, a, b) { for (var i = a; i <= b; i++) if (中身がある(rr, i)) return true; return false; };
    var 列に何かある = function (cc, a, b) { for (var i = a; i <= b; i++) if (中身がある(i, cc)) return true; return false; };
    var 広がった = true;
    while (広がった) {
      広がった = false;
      while (r1 > 0 && 行に何かある(r1 - 1, c1, c2)) { r1--; 広がった = true; }
      while (r2 < maxR && 行に何かある(r2 + 1, c1, c2)) { r2++; 広がった = true; }
      while (c1 > 0 && 列に何かある(c1 - 1, r1, r2)) { c1--; 広がった = true; }
      while (c2 < maxC && 列に何かある(c2 + 1, r1, r2)) { c2++; 広がった = true; }
    }
    return { r1: r1, c1: c1, r2: r2, c2: c2 };
  }

  /**
   * 先頭行を見出しとみなすか（★実Excelの xlGuess と同じ判定★）
   *   ★先頭行が文字で、2行目から下に数が在る時だけ 見出し★
   */
  function guessHeader(get, r1, c1, r2, c2) {
    if (r2 <= r1) return false;
    var 先頭に文字 = false;
    for (var c = c1; c <= c2; c++) { if (型と値(get(r1, c)).t === 文字) { 先頭に文字 = true; break; } }
    if (!先頭に文字) return false;
    for (var r = r1 + 1; r <= r2; r++) {
      for (var c2i = c1; c2i <= c2; c2i++) if (型と値(get(r, c2i)).t === 数) return true;
    }
    return false;
  }

  /** その式が ★$付きの参照★ を持っているか（並べ替えで指す先が変わる物） */
  function hasAbsRef(f) {
    return typeof f === 'string' && f.charAt(0) === '=' && /\$[A-Za-z]{1,3}|\$\d/.test(f);
  }

  var api = { order: order, region: region, guessHeader: guessHeader, hasAbsRef: hasAbsRef, 型と値: 型と値 };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.GridSort = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
