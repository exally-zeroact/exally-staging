/* cell-styles.js — ★セルのスタイル（ホーム→スタイル）★ 自動生成
 *
 *  ★作り方★  node scripts/make-cell-styles.mjs
 *  ★元★      docs/excel-cellstyles-2026-08-31.tsv
 *
 *  ★2026-08-29★ 実Excel を COM で 読んで ★よく使う 9つ★を 作った。
 *  ★2026-08-31★ ★47個 全部★に 広げた（司さん「細胞レベルまで 網羅して 持ち込む」）。
 *
 *  ★測り方（実測）★
 *    新規の 空ブックで ★スタイルを 1つずつ ★別のセル★に 当てて 読む★。
 *    ★同じ セルを 使い回すと 前の 書式が 残る★（2026-08-31 実際に 踏んだ＝
 *      「メモ」の 字の色を 22428 と 読み違えた。別のセルで 測り直して 0 と 判った）。
 *    ★Style を そのまま 読むのと 当てた 結果は 別物★＝★当てた 結果が 真値★。
 *
 *  ★COM の Color は BGR★（0xBBGGRR）なので RGB に 直してある。
 *
 *  ★色は 実Excelの 値を そのまま 使う★＝ここは「見た目を 借りる」ではなく
 *  ★同じ書類を 開いた時に 同じに 見える★ための 相互運用。
 *  （うちの緑は うちの画面の色。セルの中身の 色は ファイルの 中身）
 *
 *  ★手で 書かない★＝tsv を 測り直して 道具を 走らせる。
 *  見張り: tests/cell-styles.test.mjs
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.CellStyles = factory();
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  /* BGR → RGB（COM が 返す 形を 直す） */
  function bgr(n) {
    var b = (n >> 16) & 0xff, g = (n >> 8) & 0xff, r = n & 0xff;
    return '#' + [r, g, b].map(function (x) {
      var s = x.toString(16).toUpperCase();
      return s.length === 1 ? '0' + s : s;
    }).join('');
  }

  var 一覧 = [
    { 名: '標準',           形: {}, 英: 'Normal', },
    { 名: '20% - アクセント 1',形: { fontSize: 11, color: '#000000', bg: '#C0E6F5' }, },
    { 名: '20% - アクセント 2',形: { fontSize: 11, color: '#000000', bg: '#FBE2D5' }, },
    { 名: '20% - アクセント 3',形: { fontSize: 11, color: '#000000', bg: '#C1F0C8' }, },
    { 名: '20% - アクセント 4',形: { fontSize: 11, color: '#000000', bg: '#CAEDFB' }, },
    { 名: '20% - アクセント 5',形: { fontSize: 11, color: '#000000', bg: '#F2CEEF' }, },
    { 名: '20% - アクセント 6',形: { fontSize: 11, color: '#000000', bg: '#DAF2D0' }, },
    { 名: '40% - アクセント 1',形: { fontSize: 11, color: '#000000', bg: '#83CCEB' }, },
    { 名: '40% - アクセント 2',形: { fontSize: 11, color: '#000000', bg: '#F7C7AC' }, },
    { 名: '40% - アクセント 3',形: { fontSize: 11, color: '#000000', bg: '#83E28E' }, },
    { 名: '40% - アクセント 4',形: { fontSize: 11, color: '#000000', bg: '#94DCF8' }, },
    { 名: '40% - アクセント 5',形: { fontSize: 11, color: '#000000', bg: '#E49EDD' }, },
    { 名: '40% - アクセント 6',形: { fontSize: 11, color: '#000000', bg: '#B5E6A2' }, },
    { 名: '60% - アクセント 1',形: { fontSize: 11, color: '#000000', bg: '#44B3E1' }, },
    { 名: '60% - アクセント 2',形: { fontSize: 11, color: '#000000', bg: '#F1A983' }, },
    { 名: '60% - アクセント 3',形: { fontSize: 11, color: '#000000', bg: '#47D359' }, },
    { 名: '60% - アクセント 4',形: { fontSize: 11, color: '#000000', bg: '#61CBF3' }, },
    { 名: '60% - アクセント 5',形: { fontSize: 11, color: '#000000', bg: '#D86DCD' }, },
    { 名: '60% - アクセント 6',形: { fontSize: 11, color: '#000000', bg: '#8ED973' }, },
    { 名: 'アクセント 1',      形: { fontSize: 11, color: '#FFFFFF', bg: '#156082' }, },
    { 名: 'アクセント 2',      形: { fontSize: 11, color: '#FFFFFF', bg: '#E97132' }, },
    { 名: 'アクセント 3',      形: { fontSize: 11, color: '#FFFFFF', bg: '#196B24' }, },
    { 名: 'アクセント 4',      形: { fontSize: 11, color: '#FFFFFF', bg: '#0F9ED5' }, },
    { 名: 'アクセント 5',      形: { fontSize: 11, color: '#FFFFFF', bg: '#A02B93' }, },
    { 名: 'アクセント 6',      形: { fontSize: 11, color: '#FFFFFF', bg: '#4EA72E' }, },
    { 名: 'タイトル',         形: { fontSize: 18, color: '#0E2841' }, },
    { 名: 'チェック セル',      形: { bold: true, fontSize: 11, color: '#FFFFFF', bg: '#A5A5A5' }, },
    { 名: 'どちらでもない',      形: { fontSize: 11, color: '#9C5700', bg: '#FFEB9C' }, },
    { 名: 'パーセント',        形: { fontSize: 11, color: '#000000' }, 英: 'Percent', },
    { 名: 'メモ',           形: { fontSize: 11, color: '#000000', bg: '#FFFFCC' }, },
    { 名: 'リンク セル',       形: { fontSize: 11, color: '#FA7D00' }, },
    { 名: '悪い',           形: { fontSize: 11, color: '#9C0006', bg: '#FFC7CE' }, },
    { 名: '計算',           形: { bold: true, fontSize: 11, color: '#FA7D00', bg: '#F2F2F2' }, },
    { 名: '警告文',          形: { fontSize: 11, color: '#FF0000' }, },
    { 名: '桁区切り [0]',     形: { fontSize: 11, color: '#000000' }, 英: 'Comma [0]', },
    { 名: '桁区切り',         形: { fontSize: 11, color: '#000000' }, 英: 'Comma', },
    { 名: '見出し 1',        形: { bold: true, fontSize: 15, color: '#0E2841' }, },
    { 名: '見出し 2',        形: { bold: true, fontSize: 13, color: '#0E2841' }, },
    { 名: '見出し 3',        形: { bold: true, fontSize: 11, color: '#0E2841' }, },
    { 名: '見出し 4',        形: { bold: true, fontSize: 11, color: '#0E2841' }, },
    { 名: '集計',           形: { bold: true, fontSize: 11, color: '#000000' }, },
    { 名: '出力',           形: { bold: true, fontSize: 11, color: '#3F3F3F', bg: '#F2F2F2' }, },
    { 名: '説明文',          形: { fontSize: 11, color: '#7F7F7F' }, },
    { 名: '通貨 [0]',       形: { fontSize: 11, color: '#000000' }, 英: 'Currency [0]', },
    { 名: '通貨',           形: { fontSize: 11, color: '#000000' }, 英: 'Currency', },
    { 名: '入力',           形: { fontSize: 11, color: '#3F3F76', bg: '#FFCC99' }, },
    { 名: '良い',           形: { fontSize: 11, color: '#006100', bg: '#C6EFCE' }, },
  ];

  /* ★スタイルを 当てた時に 触る所★＝ここに 無い書式は 残す（Excelと 同じ）
     「標準」は ★この4つを 外す★＝元に戻す */
  var 触るキー = ['bold', 'fontSize', 'color', 'bg'];

  function 探す(名) {
    for (var i = 0; i < 一覧.length; i++) if (一覧[i].名 === 名) return 一覧[i];
    return null;
  }

  /** セル1つに 当てる（当てた物の 一覧を 返す＝元に戻す為） */
  function 当てる(cell, 名) {
    var st = 探す(名);
    if (!st || !cell) return null;
    var 控え = [];
    for (var i = 0; i < 触るキー.length; i++) {
      var k = 触るキー[i];
      var 新 = st.形[k];
      if (cell[k] === 新) continue;
      控え.push({ prop: k, old: cell[k], nw: 新 });
      if (新 === undefined) delete cell[k];
      else cell[k] = 新;
    }
    return 控え;
  }

  return { 一覧: 一覧, 探す: 探す, 当てる: 当てる, 触るキー: 触るキー, _bgr: bgr };
}));
