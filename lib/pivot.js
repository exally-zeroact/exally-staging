/* pivot.js — ★ピボットテーブル（挿入→ピボットテーブル）★ 2026-08-30
 *
 *  ★真値（実Excel 16.0 で 実測）★ … tools/measure-pivot.ps1
 *    元の表（A1:D6）
 *      店 品 数 金 ／ 東 A 1 100 ／ 東 B 2 200 ／ 西 A 3 300 ／ 西 B 4 400 ／ 東 A 5 500
 *    行＝店・列＝品・値＝金（合計）で 作ると:
 *
 *      3行目: 「合計 / 金」  「列ラベル」
 *      4行目: 「行ラベル」    A     B     総計
 *      5行目:  西            300   400    700
 *      6行目:  東            600   200    800
 *      7行目:  総計          900   600   1500
 *
 *    ⇒ ①左上は ★「合計 / 金」★（＝「集計の名前 / 元の列名」）
 *      ②その 右に ★「列ラベル」★・下に ★「行ラベル」★
 *      ③総計は ★行も 列も 出す★（RowGrand=True / ColumnGrand=True）
 *      ④空セルは ★空のまま★（NullString=''）
 *      ⑤小計は True（1つの 列だけの 時は 出ない）
 *
 *  ★実Excelと 違う所（黙らない）★
 *    実測で 行の 並びが ★西 → 東★ に なった。字の 番号だと 東(U+6771) の方が 先。
 *    ＝★実Excel は 日本語を「読み」で 並べている★（にし → ひがし）。
 *    うちは 読みの 辞書を 持っていないので ★字の 順★で 並べる。
 *    （読みで 並べるには かなの 辞書が 要る＝持っていない物を 持っているふりに しない）
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.Pivot = factory();
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  function 字(v) { return (v === null || v === undefined) ? '' : String(v); }

  function 数にする(v) {
    if (v === null || v === undefined || v === '') return null;
    if (typeof v === 'number') return isFinite(v) ? v : null;
    var s = String(v).replace(/[,¥￥\s]/g, '');
    if (s === '') return null;
    var n = Number(s);
    return isFinite(n) ? n : null;
  }

  /* 集計の やり方（実Excel と 同じ 呼び名） */
  var 集計たち = {
    '合計': function (値たち) {
      var s = 0, あり = false;
      for (var i = 0; i < 値たち.length; i++) { var n = 数にする(値たち[i]); if (n !== null) { s += n; あり = true; } }
      return あり ? s : null;
    },
    '個数': function (値たち) {
      var n = 0;
      for (var i = 0; i < 値たち.length; i++) if (字(値たち[i]) !== '') n++;
      return n;      /* ★実Excelの「データの個数」＝空でない物を 数える★ */
    },
    '数値の個数': function (値たち) {
      var n = 0;
      for (var i = 0; i < 値たち.length; i++) if (数にする(値たち[i]) !== null) n++;
      return n;
    },
    '平均': function (値たち) {
      var s = 0, k = 0;
      for (var i = 0; i < 値たち.length; i++) { var n = 数にする(値たち[i]); if (n !== null) { s += n; k++; } }
      return k ? s / k : null;
    },
    '最大': function (値たち) {
      var m = null;
      for (var i = 0; i < 値たち.length; i++) { var n = 数にする(値たち[i]); if (n !== null && (m === null || n > m)) m = n; }
      return m;
    },
    '最小': function (値たち) {
      var m = null;
      for (var i = 0; i < 値たち.length; i++) { var n = 数にする(値たち[i]); if (n !== null && (m === null || n < m)) m = n; }
      return m;
    },
  };

  /** ★値の 列が 数なら 合計・字なら 個数★（実Excel の 当て方と 同じ 考え） */
  function 集計を当てる(列の値たち) {
    var 数 = 0, 字あり = 0;
    for (var i = 0; i < 列の値たち.length; i++) {
      if (字(列の値たち[i]) === '') continue;
      if (数にする(列の値たち[i]) !== null) 数++; else 字あり++;
    }
    return (数 > 字あり) ? '合計' : '個数';
  }

  /**
   * ピボットを 作る
   * @param 表    … [[見出し…],[値…],…]（1行目が 見出し）
   * @param 決め  … { 行: '店', 列: '品', 値: '金', 集計: '合計' }（列と 集計は 無くてよい）
   * @returns { ok, なぜ, 表: [[…]], 行の名, 列の名 }
   *
   * 出す 形は ★実測どおり★（左上＝「合計 / 金」・「列ラベル」・「行ラベル」・総計）
   */
  function 作る(表, 決め) {
    if (!表 || 表.length < 2 || !表[0] || !表[0].length) {
      return { ok: false, なぜ: '見出しの 行と 中身が 1行 以上 要ります' };
    }
    var 見出し = 表[0].map(字);
    var 行の列 = 見出し.indexOf(字(決め && 決め.行));
    if (行の列 < 0) return { ok: false, なぜ: '行に する 列が 見つかりません' };
    var 列の列 = (決め && 決め.列) ? 見出し.indexOf(字(決め.列)) : -1;
    if (決め && 決め.列 && 列の列 < 0) return { ok: false, なぜ: '列に する 列が 見つかりません' };
    var 値の列 = 見出し.indexOf(字(決め && 決め.値));
    if (値の列 < 0) return { ok: false, なぜ: '値に する 列が 見つかりません' };

    var 値たち全部 = [];
    for (var r = 1; r < 表.length; r++) 値たち全部.push(表[r][値の列]);
    var 集計名 = (決め && 決め.集計 && 集計たち[決め.集計]) ? 決め.集計 : 集計を当てる(値たち全部);
    var 集計 = 集計たち[集計名];

    /* 行と 列の 見出しを 集める（★出てきた順では なく 並べる★。
       ★実Excel は 読みで 並べる（実測 西→東）／うちは 字の 順★＝上の 見出しに 書いた） */
    var 行名 = [], 列名 = [], 見た行 = {}, 見た列 = {};
    var 箱 = {};
    for (var r2 = 1; r2 < 表.length; r2++) {
      var 行 = 字(表[r2][行の列]);
      var 列 = 列の列 >= 0 ? 字(表[r2][列の列]) : '';
      if (!見た行[行]) { 見た行[行] = true; 行名.push(行); }
      if (列の列 >= 0 && !見た列[列]) { 見た列[列] = true; 列名.push(列); }
      var k = 行 + '\u0000' + 列;
      if (!箱[k]) 箱[k] = [];
      箱[k].push(表[r2][値の列]);
    }
    行名.sort();
    列名.sort();
    if (列の列 < 0) 列名 = [''];

    /* 形を 組む（★実測どおり★）
       ・列を 使う時   … 1行目「合計 / 金」「列ラベル」／2行目「行ラベル」列名… 総計
       ・列を 使わない時 … ★「行ラベル」「合計 / 金」の 2列だけ★
         （「列ラベル」の 行も 総計の 列も 出ない＝2026-08-30 実測） */
    var 列を使う = (列の列 >= 0);
    var 値の題 = 集計名 + ' / ' + 見出し[値の列];
    var 出 = [];
    if (列を使う) {
      出.push([値の題, '列ラベル']);
      var 見出し行 = ['行ラベル'];
      for (var i = 0; i < 列名.length; i++) 見出し行.push(列名[i]);
      見出し行.push('総計');
      出.push(見出し行);
    } else {
      出.push(['行ラベル', 値の題]);
    }

    for (var j = 0; j < 行名.length; j++) {
      var 一行 = [行名[j]];
      var まとめ = [];
      for (var k2 = 0; k2 < 列名.length; k2++) {
        var v = 箱[行名[j] + '\u0000' + 列名[k2]] || [];
        for (var m = 0; m < v.length; m++) まとめ.push(v[m]);
        var 答 = v.length ? 集計(v) : null;
        一行.push(答 === null ? '' : 答);      /* ★空セルは 空のまま（実測 NullString=''）★ */
      }
      if (列を使う) 一行.push(まとめ.length ? 集計(まとめ) : '');
      出.push(一行);
    }

    /* 総計の 行（★実測＝行も 列も 出す★） */
    var 総 = ['総計'];
    var 全部 = [];
    for (var k3 = 0; k3 < 列名.length; k3++) {
      var 縦 = [];
      for (var j2 = 0; j2 < 行名.length; j2++) {
        var v2 = 箱[行名[j2] + '\u0000' + 列名[k3]] || [];
        for (var m2 = 0; m2 < v2.length; m2++) { 縦.push(v2[m2]); 全部.push(v2[m2]); }
      }
      総.push(縦.length ? 集計(縦) : '');
    }
    if (列を使う) 総.push(全部.length ? 集計(全部) : '');
    出.push(総);

    return { ok: true, 表: 出, 行の名: 行名, 列の名: 列名, 集計: 集計名 };
  }

  /** ★おすすめ★＝どの 列を 行・列・値に するかを 当てる
   *  ・値 … 数が 一番 多く 入っている 列
   *  ・行 … 字の 列のうち ★種類が 一番 少ない★物（多すぎると 表に ならない）
   *  ・列 … 次に 種類が 少ない 字の 列（3種類 以下の 時だけ。多いと 横に 伸びすぎる） */
  function おすすめ(表) {
    if (!表 || 表.length < 2) return null;
    var 見出し = 表[0].map(字);
    var 数の多さ = [], 種類 = [];
    for (var c = 0; c < 見出し.length; c++) {
      var 数 = 0, 見た = {}, 個 = 0;
      for (var r = 1; r < 表.length; r++) {
        var v = 表[r][c];
        if (数にする(v) !== null) 数++;
        var s = 字(v);
        if (s !== '' && !見た[s]) { 見た[s] = true; 個++; }
      }
      数の多さ.push(数); 種類.push(個);
    }
    var 値 = -1;
    for (var i = 0; i < 見出し.length; i++) if (値 < 0 || 数の多さ[i] > 数の多さ[値]) 値 = i;
    if (値 < 0 || !数の多さ[値]) return null;
    var 字の列 = [];
    for (var j = 0; j < 見出し.length; j++) {
      if (j === 値) continue;
      if (数の多さ[j] > (表.length - 1) / 2) continue;      /* 数ばかりの 列は 行に しない */
      if (種類[j] < 1) continue;
      字の列.push({ i: j, 種類: 種類[j] });
    }
    if (!字の列.length) return null;
    字の列.sort(function (a, b) { return a.種類 - b.種類; });
    var 決め = { 行: 見出し[字の列[0].i], 値: 見出し[値] };
    if (字の列.length > 1 && 字の列[1].種類 <= 3) 決め.列 = 見出し[字の列[1].i];
    return 決め;
  }

  return { 作る: 作る, おすすめ: おすすめ, 集計たち: 集計たち, 集計を当てる: 集計を当てる, _数にする: 数にする };
}));
