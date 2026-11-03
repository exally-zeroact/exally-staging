/* make-cell-styles.mjs — ★実Excel の 組み込みスタイル 47個から lib/cell-styles.js を 起こす★ 2026-08-31
 *
 *  ★元★ docs/excel-cellstyles-2026-08-31.tsv
 *        （新規の 空ブックで ★スタイルを 1つずつ ★別のセル★に 当てて 読んだ★ 実測）
 *
 *  ★★前からの 物を 捨てない★★
 *    lib/cell-styles.js は 2026-08-29 に ★9個★を 測って 作ってあった（画面も 繋がっている）。
 *    ★形（一覧／形／探す／当てる／触るキー／_bgr）は そのまま★にして
 *    ★中身だけ 47個に 広げる★。
 *
 *  ★色の 並び★ Excel は ★BGR★（R=c&255 / G=(c>>8)&255 / B=(c>>16)&255）
 *  ★手で 書かない★＝tsv を 測り直して この 道具を 走らせる。
 *
 *  走らせ方: node scripts/make-cell-styles.mjs [--check]
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const NL = String.fromCharCode(10), TAB = String.fromCharCode(9);

const 生 = fs.readFileSync(path.join(ROOT, 'docs/excel-cellstyles-2026-08-31.tsv'), 'utf8')
  .split(NL).filter((l) => l.trim() && !l.startsWith('#'));
const 見出し = 生[0].split(TAB);
const 行 = 生.slice(1).map((l) => {
  const c = l.split(TAB), o = {};
  見出し.forEach((h, i) => { o[h] = c[i] === undefined ? '' : c[i]; });
  return o;
});
if (行.length !== 47) { console.error('★元が 読めていない★ ' + 行.length + '行'); process.exit(1); }

/** ★BGR の 数 → #RRGGBB★（空なら null） */
function 色(v) {
  if (v === '' || v === undefined || v === null) return null;
  const n = Number(v);
  if (!isFinite(n)) return null;
  const r = n & 255, g = (n >> 8) & 255, b = (n >> 16) & 255;
  const 二 = (x) => (x < 16 ? '0' : '') + x.toString(16).toUpperCase();
  return '#' + 二(r) + 二(g) + 二(b);
}
const 真 = (v) => String(v) === 'True';

/* ★前からの 9個と 同じ 書き方に する★
   （前は 色が 黒でも color を 書いていた＝「メモ」。★同じに なる★ように そろえる） */
/* ★英語の ままの 名前を 画面の 字に そろえる★
   実Excel の COM は この 6つを ★英語で★ 返す（実測）。
   画面（ホーム→スタイル）に 出るのは 日本語なので ★画面の 字に 合わせる★。
   ★元の 名前も 残す★＝ファイルに 書く時は 英語が 要る（styles.xml の builtinId）。 */
const 日本語名 = {
  'Normal': '標準',
  'Currency': '通貨',
  'Currency [0]': '通貨 [0]',
  'Percent': 'パーセント',
  'Comma': '桁区切り',
  'Comma [0]': '桁区切り [0]'
};

const 出 = [];
for (const r of 行) {
  const 形 = {};
  if (真(r.bold)) 形.bold = true;
  const s = Number(r.size);
  形.fontSize = isFinite(s) ? s : 11;
  const f = 色(r.fontColor);
  if (f) 形.color = f;
  const 塗 = 色(r.fill);
  if (塗) 形.bg = 塗;
  出.push({ 名: 日本語名[r.name] || r.name, 英: r.name, 形: 形 });
}
/* ★「標準」は 先頭★＝形は 空（当てると 4つを 外す＝元に戻る） */
/* ★「標準」を 先頭に★（当てると 4つを 外す＝元に 戻る） */
const 標準 = 出.find((v) => v.名 === '標準');
const 並び = [{ 名: '標準', 英: '標準' === (標準 && 標準.英) ? undefined : (標準 && 標準.英), 形: {} }]
  .concat(出.filter((v) => v.名 !== '標準'));

const 幅 = 並び.reduce((m, v) => Math.max(m, v.名.length), 0);
const 一覧字 = 並び.map((v) => {
  const 名 = "'" + v.名 + "',";
  const 埋 = ' '.repeat(Math.max(0, 幅 + 3 - 名.length));
  const 中 = Object.keys(v.形).map((k) => k + ': '
    + (typeof v.形[k] === 'string' ? "'" + v.形[k] + "'" : v.形[k])).join(', ');
  const 英 = (v.英 && v.英 !== v.名) ? (" 英: '" + v.英 + "',") : '';
  return '    { 名: ' + 名 + 埋 + '形: {' + (中 ? ' ' + 中 + ' ' : '') + '},' + 英 + ' },';
}).join(NL);

const 中身 = `/* cell-styles.js — ★セルのスタイル（ホーム→スタイル）★ 自動生成
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
${一覧字}
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
`;

const OUT = path.join(ROOT, 'lib/cell-styles.js');
if (!process.argv.includes('--check')) fs.writeFileSync(OUT, 中身, 'utf8');
console.log('[make-cell-styles]');
console.log('  元の 行 … ' + 行.length);
console.log('  ★一覧 … ' + 並び.length + '個★'
  + (process.argv.includes('--check') ? '  ※--check＝書き込んでいない' : ''));
console.log('  塗りが 在る … ' + 並び.filter((v) => v.形.bg).length + '個');
console.log('  太字 … ' + 並び.filter((v) => v.形.bold).length + '個');
