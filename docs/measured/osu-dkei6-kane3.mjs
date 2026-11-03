/* osu-dkei6-kane3.mjs — ★実Excel に 聞いた 26行を 本番の 道で 押し返す★（2026-09-15）
 *
 *  ★紙★ … `docs/measured/kansuu46/golden-dkei6-kane3-2026-09-15.tsv`
 *          （★実Excel 16.0 build 20326／2026-09-15／Excel の 枠 1回★）
 *
 *  ★★材料を 自分で 決めません★★
 *    ・A1:B5 ／ D1:D2 ／ H1:H3 … ★紙の `#材料` の 行から★ 読む
 *    ・F1:G3（条件の 置き場）  … ★★紙を 取った `toru-dkei6-kane3.ps1` の 字から★★ 読む
 *      ⇒★手で 写しません★＝★写すと すぐ ずれる★
 *      （2026-09-15 … 材料を 自分で 決めて ★928本の 偽の 負け★を 出した 事が 在る）
 *
 *  ★押し方★ … `honban-no-michi.mjs` の `建てる()`＝★本番の 3段★
 *    ①`convertFormula` → ②材料を 板に 入れる → ③JS層 → エンジン
 *    ★1段でも 飛ばすと「うちが 負けて いる」向きの 偽の 答えが 出ます★
 *
 *  ★見て いない 事★
 *    ・★★この 台の 数は 画面の 数では ない★★（node の 台）
 *      ＝★板ごと `setSheetContent` で 入れて います★（本番は 1マスずつ）
 *      ★2026-09-10 に 踏んだ★ … 板ごと 入れて 裸の `=LINEST` が `#VALUE!` に なり
 *        ★本番が 壊れて いると 報せる 一歩 手前★まで 行った
 *      ⇒★画面の 事を 言いたいなら ★ブラウザで 押す★★
 *    ・★この 紙に 無い 形★は 何も 言えません
 *
 *  使い方: node docs/measured/osu-dkei6-kane3.mjs
 */
import fs from 'node:fs';
import path from 'node:path';
import { 建てる, ROOT } from './honban-no-michi.mjs';

const 紙の道 = path.join(ROOT, 'docs/measured/kansuu46/golden-dkei6-kane3-2026-09-15.tsv');
const 道具の道 = path.join(ROOT, 'docs/measured/kansuu46/toru-dkei6-kane3.ps1');
const 読む = (p) => fs.readFileSync(p, 'utf8').replace(/^﻿/, '');

/* ── ★門は 使う 前に 置く★ ── */
for (const p of [紙の道, 道具の道]) {
  if (!fs.existsSync(p)) { console.error('★★道具が 壊れて います★★ … 無い … ' + p); process.exit(1); }
}

/* ── ①紙から … 材料（`#材料`）と 聞いた 26行 ── */
const 全行 = 読む(紙の道).split(/\r?\n/);
const 材料 = [];
const 問い = [];
for (const l of 全行) {
  if (l.startsWith('#材料')) {
    const c = l.split('\t');
    材料.push({ マス: c[1], 値: c[2], 型: c[3] });
    continue;
  }
  if (!l || l.startsWith('#')) continue;
  const c = l.split('\t');
  if (c.length >= 4) 問い.push({ 札: c[0], 式: c[1], 実: c[2], 型: c[3] });
}

/* ── ②道具の 字から … 条件の 置き場（札 → 置き） ── */
const 置き表 = new Map();
for (const m of 読む(道具の道).matchAll(
  /札 = '([^']*)';\s*置き = \[ordered\]@\{([^}]*)\};/g)) {
  const 置き = [];
  for (const p of m[2].matchAll(/'([A-Z]+\d+)'\s*=\s*('([^']*)'|-?[\d.]+)/g)) {
    置き.push({ マス: p[1], 値: p[3] !== undefined ? p[3] : Number(p[2]), 字か: p[3] !== undefined });
  }
  置き表.set(m[1], 置き);
}

console.log('★読んだ 物★');
console.log('   紙 … ' + path.relative(ROOT, 紙の道) + ' … 聞いた ★' + 問い.length + '行★／材料 ★' + 材料.length + '行★');
console.log('   道具 … ' + path.relative(ROOT, 道具の道) + ' … 条件の 置き ★' + 置き表.size + '件★');
if (!問い.length || !材料.length || !置き表.size) {
  console.error('★★道具が 壊れて います★★ … 読んだ 数が 0 です');
  process.exit(1);
}
/* ★★札が 紙と 道具で 食い違ったら 赤★★（★黙って 条件なしで 押さない★） */
const 無い札 = 問い.filter((q) => !置き表.has(q.札)).map((q) => q.札);
if (無い札.length) {
  console.error('★★道具が 壊れて います★★ … 紙に 在る 札が 道具に 無い … ' + 無い札.join(' '));
  process.exit(1);
}

/* ── ③本番の 道で 押す ── */
const 番地 = (s) => {
  const m = /^([A-Z]+)(\d+)$/.exec(s);
  let c = 0;
  for (const ch of m[1]) c = c * 26 + (ch.charCodeAt(0) - 64);
  return { 行: Number(m[2]) - 1, 列: c - 1 };
};
const 赤の名 = (t) => ({
  VALUE: '#VALUE!', DIV_BY_ZERO: '#DIV/0!', NUM: '#NUM!', NA: '#N/A',
  NAME: '#NAME?', REF: '#REF!', CYCLE: '#CYCLE!', ERROR: '#ERROR!', SPILL: '#SPILL!',
}[t] || ('#' + t));

const 台 = await 建てる();
const { EF, hf, SID } = 台;
const 式の行 = 0, 式の列 = 9;          /* J1 … ★道具と 同じ 置き場★ */

function 板を作る(置き) {
  const 表 = [];
  for (let r = 0; r < 6; r++) 表.push(new Array(12).fill(null));
  for (const m of 材料) {
    const a = 番地(m.マス);
    表[a.行][a.列] = m.型 === '数' ? Number(m.値) : (m.型 === '空' ? null : m.値);
  }
  for (const p of (置き || [])) {
    const a = 番地(p.マス);
    表[a.行][a.列] = p.字か ? String(p.値) : Number(p.値);
  }
  return 表;
}

function 押す(式, 置き) {
  let 後;
  try { 後 = EF.convertFormula(式); } catch (e) { return '★変換で 投げた★'; }
  const 表 = 板を作る(置き);
  表[式の行][式の列] = 後;
  hf.setSheetContent(SID, 表);
  try {
    const js = EF._jsComputeFormula(0, 式);
    if (js !== null && js !== undefined) return String(js);
  } catch (e) { /* エンジンへ */ }
  const v = hf.getCellValue({ sheet: SID, row: 式の行, col: 式の列 });
  if (v && v.type) return 赤の名(v.type);
  if (v === null || v === undefined) return '';
  return String(v);
}

console.log('');
console.log('札\t式\tうち\t実Excel\t合うか');
let 合 = 0;
const 外れ = [];
for (const q of 問い) {
  const 出 = 押す(q.式, 置き表.get(q.札));
  const ok = 出 === String(q.実);
  if (ok) 合++; else 外れ.push({ ...q, 出 });
  console.log(q.札 + '\t' + q.式 + '\t' + 出 + '\t' + q.実 + '\t' + (ok ? '○' : '★×★'));
}
console.log('');
console.log('★★合った ' + 合 + '/' + 問い.length + '★★');
if (外れ.length) {
  console.log('★★外した ' + 外れ.length + '件★★');
  for (const x of 外れ) console.log('   ' + x.札 + '  うち ★' + x.出 + '★ ／ 実Excel ★' + x.実 + '★');
}
