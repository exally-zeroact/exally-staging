/* break-check.mjs — ★測った 数を わざと 1つ 壊して 試験が 赤に なるか★ 2026-08-30
 *
 *  ★何の 為か★
 *    「試験が 緑」は「試験が 見ている」の 証拠に ならない。
 *    ★実Excel から 測った 数★を 1つずつ 壊して、★その 試験が 本当に 赤に なるか★を 数える。
 *
 *  ★2026-08-30 の 結果＝12/12 全部 赤に なった★
 *
 *  ★大事★
 *    ・この 道具は ★repo の 部品を 一時的に 書き換える★。
 *      必ず `finally` で 元に 戻し、最後に `git status` で ★残り 0件★を 見せる。
 *    ・だから ★tests/run.js には 入れない★（毎回 走らせる物では ない）。
 *      ★数を 増やした 時・見張りを 疑う 時★に 手で 走らせる。
 *    ・途中で 止めると 部品が 壊れたまま に なる。★止めたら git status を 見る事★。
 *
 *  走らせ方: node tools/break-check.mjs
 */
import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';

/* ★どこから 走らせても 同じ★＝この ファイルの 1つ 上を repo と する */
const ROOT = path.join(path.dirname(new URL(import.meta.url).pathname.replace(/^\/([A-Za-z]:)/, '$1')), '..');

/* [部品, 壊す前, 壊した後, 赤に なってほしい 試験] */
const 例 = [
  ['lib/theme.js',
    'L = (t > 0) ? Math.floor(L * (1 - t) + HLSMAX * t) : Math.floor(L * (1 + t));',
    'L = (t > 0) ? Math.round(L * (1 - t) + HLSMAX * t) : Math.round(L * (1 + t));',
    'tests/theme.test.mjs', '色の濃淡を 切り捨て→四捨五入'],

  ['lib/timeline.js',
    'return new Date(Date.UTC(1899, 11, 31) + n * 86400000);',
    'return new Date(Date.UTC(1899, 11, 30) + n * 86400000);',
    'tests/shape-map-time.test.mjs', '日付の通し番号を 1日 ずらす'],

  ['lib/adv-filter.js',
    'var 出 = [元[0].slice()];',
    'var 出 = [];',
    'tests/data-conn.test.mjs', '詳細設定で 見出しを 付けない'],

  ['lib/ink-shape.js',
    'if (閉じ && 出.length >= 2) {',
    'if (false && 出.length >= 2) {',
    'tests/shape-map-time.test.mjs', '手書きの 角を 2回 数える'],

  ['lib/jp-map.js',
    "{ 名: '沖縄',   行: 9, 列: 0,  読: 'おきなわ' },",
    "{ 名: '沖縄',   行: 0, 列: 10, 読: 'おきなわ' },",
    'tests/shape-map-time.test.mjs', '沖縄を 北海道と 同じ 升に 置く'],

  ['lib/xml-map.js',
    'if (!出せるか(結び)) return null;                 /* ★実Excel も 断る★ */',
    'if (false) return null;                 /* ★実Excel も 断る★ */',
    'tests/xml-dev.test.mjs', '結ぶ前でも XML を 出せてしまう'],

  ['lib/data-model.js',
    '        訳: \'★右の「\' + 右列名 + \'」は 同じ 値が 2回 以上 出ます★\'',
    '        訳: \'ちがう字\'',
    'tests/model-web-data.test.mjs', '1つずつで ない 時の 断り方を 変える'],

  ['lib/review.js',
    "recieve: 'receive', teh: 'the',",
    "teh: 'the',",
    'tests/review2.test.mjs', 'recieve を 一覧から 外す'],

  ['lib/model3d.js',
    'var 幅 = Math.sqrt(dx * dx + dy * dy + dz * dz) || 1;',
    'var 幅 = Math.max(dx, dy, dz) || 1;',
    'tests/model-web-data.test.mjs', '3Dの 大きさを 対角線→一番長い辺'],

  ['lib/diagram.js',
    "  var 実Excelの節の数 = 5;          /* ★置いた 時 5個★ */",
    "  var 実Excelの節の数 = 4;          /* ★置いた 時 5個★ */",
    'tests/insert-diagram.test.mjs', 'SmartArt の 節を 5→4'],

  ['lib/view-mode.js',
    '  var 既定の余白 = { 上: 1.905, 下: 1.905, 左: 1.778, 右: 1.778 };   /* ★実測★ */',
    '  var 既定の余白 = { 上: 2.0, 下: 2.0, 左: 2.0, 右: 2.0 };   /* ★実測★ */',
    'tests/view-window.test.mjs', '紙の余白を 実測から ずらす'],

  ['lib/connections.js',
    "    if (種類たち[i].種 === 種) return !種類たち[i].選び直せる || 種 === 'web';",
    '    if (種類たち[i].種 === 種) return true;',
    'tests/data-conn.test.mjs', 'ファイル由来でも 更新できると 言う'],
];

let 赤になった = 0;
const 素通り = [];

for (const [部品, 前, 後, 試験, 話] of 例) {
  const p = path.join(ROOT, 部品);
  const 元 = fs.readFileSync(p, 'utf8');
  if (元.split(前).length - 1 !== 1) {
    素通り.push('★壊せなかった（目印が ' + (元.split(前).length - 1) + '個）★ ' + 部品 + ' … ' + 話);
    continue;
  }
  try {
    fs.writeFileSync(p, 元.replace(前, 後), 'utf8');
    let 赤 = false;
    try {
      execFileSync('node', [path.join(ROOT, 試験)], { cwd: ROOT, stdio: 'pipe' });
    } catch (e) { 赤 = true; }
    if (赤) { 赤になった++; console.log('  ★赤★ ' + 話 + '  → ' + 試験); }
    else { 素通り.push('★素通り★ ' + 話 + ' → ' + 試験 + ' が 緑の まま'); }
  } finally {
    fs.writeFileSync(p, 元, 'utf8');           /* ★必ず 戻す★ */
  }
}

console.log('\n★' + 赤になった + '/' + 例.length + ' 通りで 赤に なった★');
if (素通り.length) {
  console.log('★素通り ' + 素通り.length + '件★');
  素通り.forEach((v) => console.log('  ' + v));
}
/* ★戻したか を 最後に 確かめる★ */
try {
  const 差 = execFileSync('git', ['status', '--porcelain'], { cwd: ROOT, encoding: 'utf8' });
  /* ★この 道具 自身が まだ commit されていない 分は 数えない★（?? の 行） */
  const 残り = 差.split(String.fromCharCode(10))
    .filter((l) => l.trim() && l.indexOf('?? ') !== 0);
  console.log('★repo に 残った 変更＝'
    + (残り.length ? (String.fromCharCode(10) + 残り.join(String.fromCharCode(10)))
                    : '0件（ちゃんと 戻した）') + '★');
  if (残り.length) process.exitCode = 1;
} catch (e) { console.log('git status を 読めなかった'); }
process.exit(素通り.length ? 1 : 0);
