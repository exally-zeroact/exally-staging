/* make-keys.mjs — ★実Excel の キー割り当てから lib/ribbon-keys.js を 起こす★ 2026-08-30
 *
 *  ★司さんの 方針★「細胞分解レベルまで 網羅して 把握して 持ち込み パクる」
 *  ★事務の 人は 指が 覚えている★（Alt,H,1＝太字／Alt,H,A L＝左揃え）
 *  ⇒ ★Excel の 順番を そのまま 写す★（うちで 勝手に 決めない）
 *
 *  ★元★ docs/excel-keys-2026-08-30.tsv（UI Automation の AccessKey・実測 462行）
 *  ★出す物★ lib/ribbon-keys.js
 *
 *  ★引き方★ タブ|部品名 で 引く（同じ名前が 別タブに 在るので タブ込みで 見る）
 *  走らせ方: node scripts/make-keys.mjs [--check]
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const NL = String.fromCharCode(10), TAB = String.fromCharCode(9);
const 尻 = (s) => String(s).replace(/\.\.\.$/, '').replace(/…$/, '').trim();

const 生 = fs.readFileSync(path.join(ROOT, 'docs/excel-keys-2026-08-30.tsv'), 'utf8')
  .split(NL).filter((l) => l.trim() && !l.startsWith('#')).map((l) => l.split(TAB));
if (生.length < 400) { console.error('★元が 読めていない★ ' + 生.length + '行'); process.exit(1); }

/* ── タブの 入口（Alt,H など）── */
const タブ名 = ['ホーム', '挿入', '描画', 'ページ レイアウト', '数式', 'データ',
  '校閲', '表示', '自動化', '開発', 'ヘルプ'];
const タブキー = {};
for (const c of 生) {
  if (c.length < 3) continue;
  const 名 = 尻(c[1]), ak = (c[2] || '').trim();
  if (!タブ名.includes(名) || !ak) continue;
  /* ★一番 短い 物が 本来の 入口★（Alt,H,I のような 二次経路は 採らない） */
  const 押す = ak.replace(/^Alt,\s*/, '').split(/[\s,]+/).filter(Boolean);
  if (!タブキー[名] || 押す.length < タブキー[名].length) タブキー[名] = 押す;
}

/* ── 部品の キー ── */
const 部品キー = {};
const 曖昧 = {};   /* ★名前で 引くと 決まらない 物★ */
for (const c of 生) {
  if (c.length < 3) continue;
  const t = (c[0] || '').trim(), 名 = 尻(c[1]), ak = (c[2] || '').trim();
  if (!t || !名 || !ak) continue;
  if (タブ名.includes(名)) continue;                 /* タブは 上で 拾った */
  /* ★`Alt+=` の形は 「押す順」では ない★（同時に 押す 別の層）
     実測：4行（合計×2・Alt+F8・Ctrl+Alt+L）。
     ★混ぜると 同じ 鍵に 2つ 当たって ぶつかる★（2026-08-31 実測）
     ⇒ book.html の ショートカットの層が 受ける（Alt+= は 既に 動いている） */
  if (!/^Alt,/.test(ak)) continue;
  const k = t + '|' + 名;
  const 押す = ak.replace(/^Alt,\s*/, '').split(/[\s,]+/).filter(Boolean);
  /* ★同じ名前に ★違う★ キーが 2つ 在る時は ★決められない★★
     例：ページ レイアウト|印刷 は P,Y,P G（枠線）と P,Y,P H（見出し）の 2つ。
     ★名前で 引いている★以上、どちらか は 決められない＝★短い方を 選ぶと 嘘になる★
     ⇒ ★両方 捨てる★（推さない）。測り直すなら tsv に 区別できる 印を 入れる。 */
  const 前 
    = 部品キー[k];
  if (前 && 前.join(',') !== 押す.join(',')) { 曖昧[k] = true; continue; }
  部品キー[k] = 押す;
}

/* ★自前で 逃さない★（逃し方を 手で 書いて 2026-08-30 に 壊した） */
const q = (s) => JSON.stringify(String(s));
for (const k of Object.keys(曖昧)) delete 部品キー[k];

const 行 = [];
for (const k of Object.keys(部品キー).sort()) {
  行.push('    ' + q(k) + ': ' + JSON.stringify(部品キー[k]) + ',');
}
const タブ行 = タブ名.filter((n) => タブキー[n])
  .map((n) => '    ' + q(n) + ': ' + JSON.stringify(タブキー[n]) + ',');

const 中身 = `/* ribbon-keys.js — ★Alt の キー割り当て（実Excel そのまま）★ 自動生成
 *
 *  ★作り方★  node scripts/make-keys.mjs
 *  ★元★      docs/excel-keys-2026-08-30.tsv
 *             （実Excel 16.0 build 20326 を UI Automation で 測った 462行）
 *
 *  ★なぜ 要るか（司さん 2026-08-30）★
 *    「Excel を 細胞分解レベルまで 網羅して 把握した上で 持ち込み パクる」
 *    ★毎日 Excel を 使う人ほど Alt の 順番で 打つ★＝ここが 一番 体感に 効く。
 *    実測＝Excel 462個 vs うち ★0個★ だった。
 *
 *  ★勝手に 決めない★＝キーは ★実Excel が 持っている 物だけ★。
 *    直す時は tsv を 測り直して この 道具を 走らせる（手で 書かない）。
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.RibbonKeys = factory();
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  /* ★タブの 入口★（Alt を 押した 次に 打つ 字） */
  var タブ = {
${タブ行.join(NL)}
  };

  /* ★部品の キー★（キー＝'タブ|部品名'／値＝Alt の 後に 打つ 順） */
  var 部品 = {
${行.join(NL)}
  };

  /** その タブの 入口（無ければ null） */
  function タブの鍵(名) { return タブ[名] || null; }
  /** その 部品の 鍵（無ければ null） */
  function 部品の鍵(t, p) { return 部品[t + '|' + p] || null; }
  /** 数える */
  function 数える() {
    return { タブ: Object.keys(タブ).length, 部品: Object.keys(部品).length };
  }

  return { タブ: タブ, 部品: 部品, タブの鍵: タブの鍵, 部品の鍵: 部品の鍵, 数える: 数える };
}));
`;

const OUT = path.join(ROOT, 'lib/ribbon-keys.js');
if (!process.argv.includes('--check')) fs.writeFileSync(OUT, 中身, 'utf8');
console.log('[make-keys]');
console.log('  元の 行 … ' + 生.length);
console.log('  ★タブの 入口 … ' + Object.keys(タブキー).length + '個★');
console.log('  ★名前で 決まらないので 捨てた … ' + Object.keys(曖昧).length + '個★');
console.log('  ★部品の キー … ' + Object.keys(部品キー).length + '個★'
  + (process.argv.includes('--check') ? '  ※--check＝書き込んでいない' : ''));
