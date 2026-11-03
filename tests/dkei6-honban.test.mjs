/* dkei6-honban.test.mjs — ★棚㉝ の 6か所を 本番の 道で 実Excel と 突き合わせる★（2026-09-15）
 *
 *  ★★この 紙は 直す 前に 書きました★★（★先に 赤★）
 *    ⇒★「動いた」の 後に 書くと ★通る形に 合わせて★ しまいます★
 *
 *  ★★物差し★★ ... ★実Excel★（★Excel の 枠 1回で 26件 聞いた★・2026-09-15）
 *    紙   `docs/measured/kansuu46/golden-dkei6-kane3-2026-09-15.tsv`（16.0 build 20326）
 *    道具 `docs/measured/kansuu46/toru-dkei6-kane3.ps1`
 *
 *  ★★材料を 自分で 決めません★★
 *    ・A1:B5 ／ D1:D2 ／ H1:H3 ... ★紙の `#材料` の 行から★
 *    ・F1:G3（条件）           ... ★紙を 取った `.ps1` の 字から★（★手で 写さない★）
 *
 *  ★★直す 前に 分かって いる 赤（★10件★）★★
 *    ②多列(かつ)片方 外れ 2／0 ・ ②多行(または) 2／5
 *    ④`>3` 0／9 ・ ④`<>2` 0／12 ・ ④`>=3` 0／12
 *    ⑤field 超過 0／`#VALUE!` ・ ⑥DPRODUCT 0件 1／0
 *    ⑦DGET 0件 `#NUM!`／`#VALUE!` ・ 条件の マスが 空 0／14
 *    ＋㉘ `ACCRINT(DATE(...))` `#VALUE!`／16.666666666666664
 *
 *  ★★㉘（お金の `DATE()`）は 2026-09-18 に 直りました★★
 *    ★前は こう 書いて ありました★ ... 「因は 借り物の 中（`NUMBER_DATE`）
 *      ＝借り物外しと 同じ 山／だから 直しません」
 *    ★★その 見立ては 間違って いました★★
 *      実物 ... 借り物は 日付を `{ val: 39763, format: ... }` で 渡します
 *              うちの `数()` が `Number()` を 掛けて ★NaN★ ⇒ `#VALUE!`
 *      ⇒★★包みを 外す 1行で 直りました★★（借り物は 外して いません）
 *      ⇒★お客さんの 道で 数えたら ★729本★ これで 落ちて いました★
 *
 *  ★★この 台の 数は 画面の 数では ない★★（node の 台／板ごと `setSheetContent`）
 *    ⇒★画面の 事を 言いたいなら ブラウザで 押す★
 *
 *  使い方: node tests/dkei6-honban.test.mjs
 *          node tests/dkei6-honban.test.mjs --self-test
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const 読む道 = (p) => import(pathToFileURL(path.join(ROOT, p)).href);
const { 建てる } = await 読む道('docs/measured/honban-no-michi.mjs');

const 紙の道 = path.join(ROOT, 'docs/measured/kansuu46/golden-dkei6-kane3-2026-09-15.tsv');
const 道具の道 = path.join(ROOT, 'docs/measured/kansuu46/toru-dkei6-kane3.ps1');
const 字を読む = (p) => fs.readFileSync(p, 'utf8').replace(/^﻿/, '');

/* ★★「まだ 直さない」と 名指しで 許す 物★★（★数では なく 名前で 書く★）
     ＝★数だけで 許すと ★別の 物が 壊れても 気づけません★★ */
const まだ = new Map([
  /* ★★2026-09-18 ... ㉘（お金の DATE()）を ★外しました★★★
     ＝★直ったからです★（`lib/formula-kane-plug.js` の 1行）
     ＝★★この 紙に 書いて あった 訳は ★間違って いました★★★
        書いて あった 訳 ...「因は ★借り物の 中（NUMBER_DATE）★／借り物外しと 同じ 山」
        ★実物★ ...... 因は ★うちの 包み（`lib/formula-kane-plug.js` の `数()`）★
                    借り物は 日付を `{ val: 39763, format: ... }` で 渡す
                    ⇒`Number()` が NaN ⇒ #VALUE!
        ⇒★★借り物を 外さなくても 直せました★★
     ⇒★「因は 借り物の 中」という 見立てで 棚に 載せた 物は
        ★外から 1回 実物で 確かめる★★ */
]);

let pass = 0, fail = 0, 許し = 0;
const T = (n, f) => {
  try { f(); pass++; console.log('  ✓ ' + n); }
  catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + e.message); }
};

console.log('\n[dkei6-honban] ★棚㉝ の 6か所を 実Excel と 突き合わせる★');

/* ── 紙から ── */
const 材料 = [], 問い = [];
for (const l of 字を読む(紙の道).split(/\r?\n/)) {
  if (l.startsWith('#材料')) { const c = l.split('\t'); 材料.push({ マス: c[1], 値: c[2], 型: c[3] }); continue; }
  if (!l || l.startsWith('#')) continue;
  const c = l.split('\t');
  if (c.length >= 4) 問い.push({ 札: c[0], 式: c[1], 実: c[2] });
}
/* ── 道具の 字から（条件の 置き場） ── */
const 置き表 = new Map();
for (const m of 字を読む(道具の道).matchAll(/札 = '([^']*)';\s*置き = \[ordered\]@\{([^}]*)\};/g)) {
  const 置き = [];
  for (const p of m[2].matchAll(/'([A-Z]+\d+)'\s*=\s*('([^']*)'|-?[\d.]+)/g)) {
    置き.push({ マス: p[1], 値: p[3] !== undefined ? p[3] : Number(p[2]), 字か: p[3] !== undefined });
  }
  置き表.set(m[1], 置き);
}

/* ★★分母の 見張り★★＝★押して いない 行は「合わない」事も 出来ない★ */
const 要る問い = 26, 要る材料 = 15;
T('★分母が 在る★（聞いた ' + 要る問い + '行／材料 ' + 要る材料 + '行）', () => {
  if (問い.length !== 要る問い) throw new Error('聞いた 行が ' + 問い.length + '（' + 要る問い + ' のはず）');
  if (材料.length !== 要る材料) throw new Error('材料の 行が ' + 材料.length + '（' + 要る材料 + ' のはず）');
  if (!置き表.size) throw new Error('★道具から 条件を 1件も 読めて いません★');
  const 無い = 問い.filter((q) => !置き表.has(q.札)).map((q) => q.札);
  if (無い.length) throw new Error('★紙に 在る 札が 道具に 無い★ ... ' + 無い.join(' '));
});
/* ★★「まだ」の 名前が 紙に 在るか★★（★消えた 物を 黙って 許さない★） */
T('★「まだ 直さない」の 名前が 紙に 在る★', () => {
  for (const 名 of まだ.keys()) {
    if (!問い.some((q) => q.札 === 名)) throw new Error('★紙に 無い のに 許して います★ ... ' + 名);
  }
});

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
const 式の行 = 0, 式の列 = 9;          /* J1 ... ★道具と 同じ 置き場★ */

function 押す(式, 置き) {
  let 後;
  try { 後 = EF.convertFormula(式); } catch (e) { return '★変換で 投げた★'; }
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

for (const q of 問い) {
  const 訳 = まだ.get(q.札);
  const 出 = 押す(q.式, 置き表.get(q.札));
  if (訳) {
    /* ★許すのは「まだ 合わない」時だけ★＝★合ったら その行は 許しから 外す（赤に する）★ */
    if (出 === String(q.実)) {
      fail++;
      console.log('  ✗ ★合ったのに「まだ」の ままです★ ' + q.札 + ' ... ★許しを 外して ください★');
    } else {
      許し++;
      console.log('  ... ★まだ★ ' + q.札 + ' ... うち ' + 出 + ' ／実Excel ' + q.実 + '  ' + 訳);
    }
    continue;
  }
  T(q.札 + '  ' + q.式, () => {
    if (出 !== String(q.実)) throw new Error('うち ★' + 出 + '★ ／実Excel ★' + q.実 + '★');
  });
}

console.log('\n' + pass + ' passed, ' + fail + ' failed'
  + ' ／ ★名指しで まだ★ ' + 許し + '件（分母 ' + 問い.length + '行）');

if (process.argv.includes('--self-test')) {
  /* ★この 紙が 何を 見て いないか を 字で 残す★ */
  console.log('\n[self-test] ★見て いない 事★');
  console.log('  ・★画面では ありません★（node の 台／板ごと setSheetContent）');
  console.log('  ・★この 紙に 無い 形★（多列が 3列以上・条件が 3行以上・日付の 条件 ...）');
  console.log('  ・★㉘（お金の DATE()）は 名指しで 「まだ」★＝因は 借り物の 中');
  console.log('[self-test] ★「まだ」は ' + まだ.size + '件・全部 訳つき★');
  process.exit(まだ.size === 許し ? 0 : 1);
}
process.exit(fail ? 1 : 0);
