/* make-ribbon-spec.mjs — ★リボンの並び（正本）から 部品の骨を 起こす★ 2026-08-29
 *
 *  ★なぜ在るか（司さん）★
 *    「リボンは前から言うてるけど ★配置なども真似しろ★」
 *    「★Excel全機能全能力が Exallyに 入って Excelの最上級に なる★」
 *
 *  ★並びの正本★ = docs/excel-ribbon-flat.tsv
 *    実Excel 16.0 build 20326（日本語UI）を UI Automation で 並び順のまま 取った物。
 *    タブ12 / グループ67 / 部品288。
 *
 *  ★出す物★ = lib/ribbon-spec.js
 *    ・並び（タブ→グループ→部品）は ★正本 そのまま★
 *    ・各部品に「うちの どの働きか」を 結ぶ欄（action）を 付ける
 *      … 既に 結んである物は ★消さずに 引き継ぐ★（★手で書いた所を 潰さない★）
 *
 *  ★見せ方（司さんの指示）★
 *    「★訴えられんような見せ方で 同じように★」
 *    ⇒ ★並びと 言葉は 同じ★（人が 探せる為に 要る）
 *    ⇒ ★絵は うちの物★（Excelの絵は 1つも 写さない）＝印は 字か 自作の形
 *
 *  使い方: node scripts/make-ribbon-spec.mjs
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const TSV = path.join(ROOT, 'docs/excel-ribbon-flat.tsv');
const OUT = path.join(ROOT, 'lib/ribbon-spec.js');

if (!fs.existsSync(TSV)) {
  console.error('★正本が 無い★: docs/excel-ribbon-flat.tsv');
  process.exit(1);
}

/* ── 並び（正本）を 読む ─────────────────────────────── */
const 三つ組 = Array.from(new Set(
  fs.readFileSync(TSV, 'utf8').replace(/^\uFEFF/, '').split(/\r?\n/).filter((l) => l.trim())
)).map((l) => l.split('\t')).filter((c) => c.length >= 3);

/* ── 今 結んである物を 引き継ぐ（★手で書いた所を 潰さない★）── */
const 前 = {};
if (fs.existsSync(OUT)) {
  const src = fs.readFileSync(OUT, 'utf8');
  const re = /\{\s*t:\s*'((?:[^'\\]|\\.)*)'\s*,\s*g:\s*'((?:[^'\\]|\\.)*)'\s*,\s*p:\s*'((?:[^'\\]|\\.)*)'\s*,\s*(?:s:\s*'((?:[^'\\]|\\.)*)'\s*,)?([\s\S]*?)\}\s*,?\s*(?=\{|\];)/g;
  let m;
  while ((m = re.exec(src)) !== null) {
    const 鍵 = m[1] + '\u0000' + m[2] + '\u0000' + m[3] + '\u0000' + (m[4] || '');
    前[鍵] = m[5].trim().replace(/,$/, '');
    /* ★親の字が 付いた／外れた だけの 引っ越しでも 結び目を 落とさない★
       （4列目を 足した 2026-09-03 に、組を 移した 4行が `a: null` に 戻る所だった） */
    const 親なし鍵 = m[1] + '\u0000' + m[2] + '\u0000' + m[3] + '\u0000';
    if (前[親なし鍵] === undefined) 前[親なし鍵] = 前[鍵];
    const 品だけ鍵 = m[1] + '\u0000\u0000' + m[3] + '\u0000' + (m[4] || '');
    if (前[品だけ鍵] === undefined) 前[品だけ鍵] = 前[鍵];
  }
}

const q = (s) => "'" + String(s).replace(/\\/g, '\\\\').replace(/'/g, "\\'") + "'";
const 行 = [];
let 結んだ = 0;
for (const [t, g, p, s] of 三つ組) {
  const 親 = (s || '').trim();
  /* ★組が 変わっても（親の字の 箱へ 引っ越しても）結び目を 引き継ぐ★ */
  const 鍵 = t + '\u0000' + g + '\u0000' + p + '\u0000' + 親;
  const 品だけ鍵 = t + '\u0000\u0000' + p + '\u0000' + 親;
  const 旧の組鍵 = t + '\u0000' + 親 + '\u0000' + p + '\u0000';     /* 親の字が 組だった 頃 */
  const 中身 = 前[鍵] !== undefined ? 前[鍵]
             : 前[旧の組鍵] !== undefined ? 前[旧の組鍵]
             : 前[品だけ鍵] !== undefined ? 前[品だけ鍵] : 'a: null';
  if (!/a:\s*null/.test(中身)) 結んだ++;
  行.push('  { t: ' + q(t) + ', g: ' + q(g) + ', p: ' + q(p) + ', '
    + (親 ? 's: ' + q(親) + ', ' : '') + 中身 + ' },');
}

const 頭 = `/* ribbon-spec.js — ★リボンの並び（Excelと同じ配置）★ 自動生成 + 手で結ぶ
 *
 *  ★作り方★  node scripts/make-ribbon-spec.mjs
 *    ・並び（t=タブ / g=グループ / p=部品）は ★docs/excel-ribbon-flat.tsv そのまま★
 *      ＝実Excel 16.0 build 20326（日本語UI）を UI Automation で 取った物。
 *    ・a（action）＝うちの どの働きに 結ぶか。★ここだけ 手で 書く★
 *      作り直しても ★手で書いた a は 引き継がれる★（潰さない）。
 *
 *  ★a の書き方★
 *      a: null                         … まだ 作っていない（★ボタンを 出さない★）
 *      a: { fn: 'doUndo' }             … その関数を 呼ぶ
 *      a: { fn: 'toggleFormat', arg: 'bold' }
 *      a: { fn: 'setAlign', arg: 'left', icon: '≡' }
 *    icon … ★うちの印★（字か 自作の形）。★Excelの絵は 1つも 写さない★（司さんの指示）
 *
 *  ★見せ方の決まり（司さん 2026-08-29）★
 *    「★訴えられんような見せ方で 同じように★」
 *      ⇒ 並びと 言葉は 同じ（人が 探せる為に 要る）
 *      ⇒ ★絵は うちの物★・色も うちの物
 *
 *  ★出来ていない物は ボタンを 出さない。ただし ★この一覧には 残す★（数えられる形に）
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.RibbonSpec = factory();
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var ITEMS = [
`;
const 尾 = `  ];

  /* タブ→グループ→部品 の 木に する（並び順は 上の一覧の 通り） */
  function ツリー() {
    var tabs = [], byTab = {};
    for (var i = 0; i < ITEMS.length; i++) {
      var it = ITEMS[i];
      if (!byTab[it.t]) { byTab[it.t] = { name: it.t, groups: [], byG: {} }; tabs.push(byTab[it.t]); }
      var T = byTab[it.t];
      if (!T.byG[it.g]) { T.byG[it.g] = { name: it.g, items: [] }; T.groups.push(T.byG[it.g]); }
      T.byG[it.g].items.push(it);
    }
    return tabs;
  }

  function 数える() {
    var 全 = ITEMS.length, 有 = 0;
    for (var i = 0; i < ITEMS.length; i++) if (ITEMS[i].a) 有++;
    return { 全: 全, 有: 有, 無: 全 - 有 };
  }

  return { ITEMS: ITEMS, ツリー: ツリー, 数える: 数える };
}));
`;

fs.writeFileSync(OUT, 頭 + 行.join('\n') + '\n' + 尾, 'utf8');
console.log('[make-ribbon-spec] lib/ribbon-spec.js を 作りました');
console.log('  部品 … ' + 三つ組.length + '個（タブ ' + new Set(三つ組.map((x) => x[0])).size
  + ' / グループ ' + new Set(三つ組.map((x) => x[0] + '\u0000' + x[1])).size + '）');
console.log('  ★うちに 結んである … ' + 結んだ + '個 ／ まだ … ' + (三つ組.length - 結んだ) + '個★');
