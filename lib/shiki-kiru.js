/* shiki-kiru.js — ★式を「字の かたまり」に 切る★（2026-09-11）
 *
 *  ★★なぜ★★
 *    今 ★式を 読む 所は 借り物（HyperFormula）★です。
 *    司さん「借りんで ええように」＝★土台も 自分で 作る★
 *    土台は 5つ … ①読む ②頼りの 地図 ③計算し直す 順番 ④計算 ⑤溢れ
 *    ⇒★これは ①の 一番 下＝字に 切る 所★
 *
 *  ★★借り物の 中は 1文字も 読んで いません★★
 *    形は ★実Excel が 受ける 字★から 決めます（実物 15,799本で 確かめる）
 *
 *  ★★測り方★★
 *    ★切って 繋ぎ直したら 元の 字と 1バイトも 違わない事★
 *    ＝実Excel を 呼ばずに 測れます（取りこぼしが 0 の 証し）
 *
 *  ★かたまりの 種類★
 *    数 ／ 字（"..."） ／ 誤り（#N/A など） ／ 名前（関数名・TRUE・参照）
 *    記号（+ - * / ^ & = <> < > <= >= % : , ; 空白） ／ 括弧 ／ 中括弧
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.ShikiKiru = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var 数字 = '0123456789';
  var 記号2 = ['<=', '>=', '<>'];
  /* ★改行（\n・\r）も 空白と 同じ 記号★（2026-10-05・実Excel は 改行入りの 式を 受ける＝Alt+Enter・経営者の 物差し 137 の 組 R） */
  var 記号1 = '+-*/^&=<>%:,;()[]{} \n\r';

  function 数字か(c) { return 数字.indexOf(c) >= 0; }
  function 名前の字か(c) {
    /* ★名前に 使える 字★＝英数・下線・点・ドル・感嘆符・日本語 など
       ★記号と 括弧 以外★（当て推量で 白名簿を 作らない）
       ★★ただし「読めない 字」は 名前に しません★★（2026-09-11 見張りが 捕まえた）
         前は ★NUL などを 名前として 飲み込んで★ いました。
         ⇒★黙って 飲み込むのが 一番 危ない★＝落ちた 字が そのまま 画面に 出ます。
         ⇒★見えない 字（32未満）と 削除(127)は 断ります★ */
    if (c === undefined) return false;
    var n = c.charCodeAt(0);
    if (n < 32 || n === 127) return false;
    return 記号1.indexOf(c) < 0 && c !== '"' && c !== "'" && c !== '#';
  }

  /**
   * ★式を かたまりに 切る★
   * @param {string} 式 … '=' で 始まっても よい（'=' も 1かたまりに する）
   * @returns {{ok:boolean, 出:Array, なぜ:string}}
   *   出 … [{型, 字}]／型 … '数' '字' '誤' '名' '記' '＝'
   *   ★ok:false の 時は 何も しない★（壊すより 断る）
   */
  function 切る(式) {
    var s = String(式 == null ? '' : 式);
    var 出 = [], i = 0;
    if (s.charAt(0) === '=') { 出.push({ 型: '＝', 字: '=' }); i = 1; }
    while (i < s.length) {
      var c = s.charAt(i);

      /* ★字（"..."）★＝中の "" は 1つの " */
      if (c === '"') {
        var j = i + 1, 閉じた = false;
        while (j < s.length) {
          if (s.charAt(j) === '"') {
            if (s.charAt(j + 1) === '"') { j += 2; continue; }
            閉じた = true; j++; break;
          }
          j++;
        }
        if (!閉じた) return { ok: false, 出: 出, なぜ: '閉じて いない 字（位置 ' + i + '）' };
        出.push({ 型: '字', 字: s.slice(i, j) });
        i = j; continue;
      }

      /* ★板の 名前（'...'）★＝中の '' は 1つの ' */
      if (c === "'") {
        var k = i + 1, 閉 = false;
        while (k < s.length) {
          if (s.charAt(k) === "'") {
            if (s.charAt(k + 1) === "'") { k += 2; continue; }
            閉 = true; k++; break;
          }
          k++;
        }
        if (!閉) return { ok: false, 出: 出, なぜ: '閉じて いない 板の 名前（位置 ' + i + '）' };
        /* ★後ろに 続く 名前も 1かたまり★＝`'歩合 表'!A1` */
        while (k < s.length && (名前の字か(s.charAt(k)) || s.charAt(k) === '!')) k++;
        出.push({ 型: '名', 字: s.slice(i, k) });
        i = k; continue;
      }

      /* ★誤り（#N/A #REF! #DIV/0! …）★ */
      /* ★溢れの 参照 `E1#`★（2026-10-05）＝名前の すぐ 後ろ（空白 無し）の `#` は 名前に 付ける
           （誤りの 字 `#N/A` などは 名前の 後ろに 来ない＝前は 誤りの `#` に 切って 読めなかった） */
      if (c === '#' && 出.length && 出[出.length - 1].型 === '名' && i > 0 && s.charAt(i - 1) !== ' '
          && !名前の字か(s.charAt(i + 1)) && s.charAt(i + 1) !== '/' && s.charAt(i + 1) !== '!') {
        出[出.length - 1].字 += '#';
        i++; continue;
      }
      if (c === '#') {
        var m = i + 1;
        while (m < s.length && (名前の字か(s.charAt(m)) || s.charAt(m) === '/' || s.charAt(m) === '!')) m++;
        出.push({ 型: '誤', 字: s.slice(i, m) });
        i = m; continue;
      }

      /* ★★数で 始まる 板の 名前★★（2026-09-18・★お客さんに 出て いた 欠陥★）
           ★`=5月!E14` の `5月` を ★数★と 読んで #VALUE! に して いました★
             ＝Exally は 開いた 時に `'5月'!E14` の 引用符を 外します（実測）
             ＝`tests/cross-sheet.test.mjs` が ★合計が 34万円 少ない★ 形で 捕まえました
             ＝★誤りに ならない／画面も 崩れない／数だけ 違う★＝★一番 見つけにくい 形★
           ★見分け方★ ... ★数の 後ろに 名前の 字が 続き、その先に `!` が 在る★
             ＝`5月!E14` は 板の 名前／`5E14` は 数／`1.5` は 数
           ★`!` まで 見ます★＝★`5月` だけでは 名前に しません★（数と 見分けが 付かない） */
      if (数字か(c)) {
        var j = i;
        /* ★`!` は `名前の字か` が 通します★ので ここでは ★手前で 止めます★
             ＝止めないと `5月!E14` を 丸ごと 飲み込んで `!` が 見つかりません */
        while (j < s.length && 名前の字か(s.charAt(j)) && s.charAt(j) !== '!') j++;
        if (j > i && s.charAt(j) === '!') {
          var 板末 = j + 1;
          while (板末 < s.length && (名前の字か(s.charAt(板末)) || s.charAt(板末) === '!')) 板末++;
          出.push({ 型: '名', 字: s.slice(i, 板末) });
          i = 板末; continue;
        }
      }

      /* ★数★＝1.5 / .5 / 1.64E-14 */
      if (数字か(c) || (c === '.' && 数字か(s.charAt(i + 1)))) {
        var n = i;
        while (n < s.length && (数字か(s.charAt(n)) || s.charAt(n) === '.')) n++;
        if (s.charAt(n) === 'e' || s.charAt(n) === 'E') {
          var p = n + 1;
          if (s.charAt(p) === '+' || s.charAt(p) === '-') p++;
          if (数字か(s.charAt(p))) { p++; while (p < s.length && 数字か(s.charAt(p))) p++; n = p; }
        }
        出.push({ 型: '数', 字: s.slice(i, n) });
        i = n; continue;
      }

      /* ★2字の 記号★ */
      var 二 = s.substr(i, 2);
      if (記号2.indexOf(二) >= 0) { 出.push({ 型: '記', 字: 二 }); i += 2; continue; }

      /* ★1字の 記号★ */
      if (記号1.indexOf(c) >= 0) { 出.push({ 型: '記', 字: c }); i++; continue; }

      /* ★名前★＝関数名・参照・TRUE／表の 参照（`表[列]`）も 1かたまり */
      var q = i;
      while (q < s.length && 名前の字か(s.charAt(q))) q++;
      if (q === i) return { ok: false, 出: 出, なぜ: '読めない 字「' + c + '」（位置 ' + i + '）' };
      /* ★後ろに [ が 来たら 表の 参照★＝括弧の 深さを 数えて まとめる */
      if (s.charAt(q) === '[') {
        var 深 = 0, r = q;
        while (r < s.length) {
          if (s.charAt(r) === '[') 深++;
          else if (s.charAt(r) === ']') { 深--; if (!深) { r++; break; } }
          r++;
        }
        if (深 !== 0) return { ok: false, 出: 出, なぜ: '閉じて いない 表の 参照（位置 ' + q + '）' };
        q = r;
      }
      出.push({ 型: '名', 字: s.slice(i, q) });
      i = q;
    }
    return { ok: true, 出: 出, なぜ: '' };
  }

  /** ★繋ぎ直す★＝元の 字に 戻るか を 測る為 */
  function 繋ぐ(出) {
    var s = '';
    for (var i = 0; i < 出.length; i++) s += 出[i].字;
    return s;
  }

  return { 切る: 切る, 繋ぐ: 繋ぐ };
});
