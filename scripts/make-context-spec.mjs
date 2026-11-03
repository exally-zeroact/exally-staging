/* make-context-spec.mjs — ★コンテキストタブ（物を 選んだ時だけ 出る タブ）の 正本を 起こす★ 2026-08-31
 *
 *  ★元★ docs/excel-ribbon-context-2026-08-30.tsv
 *        （実Excel 16.0 build 20326・UI Automation・新規ブックだけ・保存なし・VBAは 動かさない）
 *        ★グラフの2タブは 1回目 失敗（ホームの中身が 出た）→ IsSelected を 確かめる形で 取り直した★
 *
 *  ★実測★ 8タブ／235部品
 *      図形の書式 55 ／ 書式 47 ／ ピボットテーブル分析 28 ／ テーブル デザイン 26
 *      スライサー 24 ／ タイムライン 23 ／ デザイン 16 ／ グラフのデザイン 16
 *
 *  ★何を 出すか★
 *    ・並び（タブ→組→部品）は ★実Excel そのまま★
 *    ・★働きが 在る 物だけ ボタンに する★（司さんの 決まり）
 *    ・働きが 無い 組は ★箱だけ 出して「これから」★＝並びの 穴として 数えられる形に 残す
 *
 *  ★どの 物を 選んだ時に どのタブが 出るか★は 実Excel と 同じ:
 *      表      → テーブル デザイン
 *      グラフ  → グラフのデザイン ／ 書式
 *      図形    → 図形の書式
 *      ピボット→ ピボットテーブル分析 ／ デザイン
 *      スライサー → スライサー
 *      タイムライン → タイムライン
 *
 *  ★手で 書かない★＝tsv を 測り直して この 道具を 走らせる。
 *  走らせ方: node scripts/make-context-spec.mjs [--check]
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const NL = String.fromCharCode(10), TAB = String.fromCharCode(9);

const 行 = fs.readFileSync(path.join(ROOT, 'docs/excel-ribbon-context-2026-08-30.tsv'), 'utf8')
  .split(NL).filter((l) => l.trim() && !l.startsWith('#')).map((l) => l.split(TAB));
if (行.length < 200) { console.error('★元が 読めていない★ ' + 行.length + '行'); process.exit(1); }

/* ★何を 選んだら 出るか★（実Excel と 同じ。★推測では なく 出し方の 記録どおり★） */
const 出る条件 = {
  'テーブル デザイン':     'テーブル',
  'グラフのデザイン':      'グラフ',
  '書式':                  'グラフ',
  '図形の書式':            '図形',
  'ピボットテーブル分析':  'ピボット',
  'デザイン':              'ピボット',
  'スライサー':            'スライサー',
  'タイムライン':          'タイムライン'
};

/* タブ→組→部品 に まとめる（★並びは 出てきた順のまま★） */
const タブ = [];
const タブ索 = {};
for (const c of 行) {
  const t = (c[0] || '').trim(), g = (c[1] || '').trim(), p = (c[2] || '').trim();
  if (!t || !g || !p) continue;
  if (!タブ索[t]) { タブ索[t] = { name: t, 出る: 出る条件[t] || null, groups: [], 組索: {} }; タブ.push(タブ索[t]); }
  const T = タブ索[t];
  if (!T.組索[g]) { T.組索[g] = { name: g, items: [] }; T.groups.push(T.組索[g]); }
  T.組索[g].items.push({ t: t, g: g, p: p });
}
for (const T of タブ) delete T.組索;

const 部品数 = タブ.reduce((s, T) => s + T.groups.reduce((x, g) => x + g.items.length, 0), 0);

const 中身 = `/* ribbon-context-spec.js — ★コンテキストタブ（物を 選んだ時だけ 出る タブ）★ 自動生成
 *
 *  ★作り方★  node scripts/make-context-spec.mjs
 *  ★元★      docs/excel-ribbon-context-2026-08-30.tsv
 *             （実Excel 16.0 build 20326・UI Automation
 *               ★新規ブックだけ・保存なし・VBAは 動かしていない★）
 *
 *  ★なぜ 要るか（司さん 2026-08-30）★
 *    「Excel を 細胞分解レベルまで 網羅して 把握した上で 持ち込み パクる」
 *    ★新規の 空ブックだけ 見ていると この 8タブ 235部品は 1つも 見えない★。
 *    実測＝実Excel ${タブ.length}タブ／${部品数}部品 vs うち ★0個★ だった。
 *
 *  ★取る時に 踏んだ 罠（記録）★
 *    ★グラフの 2タブは 1回目 失敗した★＝ホームの 中身が 出た。
 *    ★タブを 選べたか（IsSelected）を 確かめてから 中身を 取る★形に 直して 取り直した。
 *
 *  ★手で 書かない★＝tsv を 測り直して 道具を 走らせる。
 *  見張り: tests/ribbon-context.test.mjs
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.RibbonContextSpec = factory();
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  /* ★出る＝何を 選んだ時に 出るか★（実Excel と 同じ） */
  var ツリー = ${JSON.stringify(タブ, null, 2).split(NL).map((l, i) => (i ? '  ' + l : l)).join(NL)};

  /** その「選んだ物」で 出る タブたち */
  function 出るタブ(何) {
    var 出 = [];
    for (var i = 0; i < ツリー.length; i++) if (ツリー[i].出る === 何) 出.push(ツリー[i]);
    return 出;
  }

  /** 平らな 一覧（数える為） */
  function 全部() {
    var 出 = [];
    for (var i = 0; i < ツリー.length; i++) {
      for (var j = 0; j < ツリー[i].groups.length; j++) {
        var g = ツリー[i].groups[j];
        for (var k = 0; k < g.items.length; k++) 出.push(g.items[k]);
      }
    }
    return 出;
  }

  /** 数える（つないだ数は 働きを 渡すと 分かる） */
  function 数える(働き) {
    var 全 = 全部();
    var 有 = 0;
    if (働き) {
      for (var i = 0; i < 全.length; i++) {
        var n = String(全[i].p).replace(/\\.\\.\\.$/, '').replace(/…$/, '').trim();
        if (typeof 働き[n] === 'function') 有++;
      }
    }
    return { タブ: ツリー.length, 部品: 全.length, つないだ: 有, まだ: 全.length - 有 };
  }

  return { ツリー: ツリー, 出るタブ: 出るタブ, 全部: 全部, 数える: 数える };
}));
`;

const OUT = path.join(ROOT, 'lib/ribbon-context-spec.js');
if (!process.argv.includes('--check')) fs.writeFileSync(OUT, 中身, 'utf8');
console.log('[make-context-spec]');
console.log('  元の 行 … ' + 行.length);
console.log('  ★タブ … ' + タブ.length + '個★'
  + (process.argv.includes('--check') ? '  ※--check＝書き込んでいない' : ''));
console.log('  ★部品 … ' + 部品数 + '個★');
for (const T of タブ) {
  console.log('    ' + T.name.padEnd(22) + T.groups.length + '組  '
    + T.groups.reduce((x, g) => x + g.items.length, 0) + '部品  ← ' + (T.出る || '★条件が 無い★'));
}
