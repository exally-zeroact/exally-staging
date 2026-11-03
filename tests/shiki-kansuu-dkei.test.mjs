/* shiki-kansuu-dkei.test.mjs — ★D系 12個を 自前の 台に 移す★（2026-09-15）
 *
 *  ★★この 紙は 移す 前に 書きました★★（★先に 赤★）
 *
 *  ★★なぜ D系 12個か★★
 *    ★母数 489個の うち 台が 知るのは 76個★／★知らない 413個★
 *    ★その 413個の 割り★
 *       A 移すだけ（自前の プラグイン） …… 4個
 *       B 移すだけ（JS層・借り物の 表を 読まない） 13個
 *       ★★C 書き直す（JS層だが 借り物の 表を 読む） 11個★★ ←★ここ★
 *       D 書く（借り物 本体） ……………… 385個
 *    ★C が 一番 安い★＝★土台（`D系の値たち`）は DSUM の 時に もう 書いた★
 *    ⇒★1個 10行くらいで 76 → 87個に なります★
 *
 *  ★★物差しは 実Excel★★（★2枚の 紙から 読む／手で 写さない★）
 *    ①`docs/measured/kansuu46/*.tsv` … D系の 行（★21行★・2026-09-08 と 09-14）
 *    ②`docs/measured/kansuu46/golden-dkei6-kane3-2026-09-15.tsv`（★26行★の うち D系）
 *       ＋その 条件は `toru-dkei6-kane3.ps1` の 字から 読む
 *
 *  ★★材料も 紙から★★
 *    A1:A5=1,2,3,4,5 ／ B1:B5=2,4,6,8,10 ／ D1=45292 ／ D2=46023
 *    （★台は `DATE` を まだ 知りません★＝紙が 読み返した ★数★を そのまま 置く）
 *
 *  ★★見て いない 事★★
 *    ・★画面では ありません★（台だけ）
 *    ・★紙に 無い 形★（3列以上の 条件・3行以上・日付の 条件 …）
 *    ・★DSUM は 別の 紙で 見て います★（`tests/shiki-kansuu-dsum.test.mjs`）
 *
 *  使い方: node tests/shiki-kansuu-dkei.test.mjs
 *          node tests/shiki-kansuu-dkei.test.mjs --self-test
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(path.join(ROOT, 'package.json'));
const H = require_(path.join(ROOT, 'lib/shiki-hyou.js'));
const 字を読む = (p) => fs.readFileSync(p, 'utf8').replace(/^﻿/, '');

const D系 = /^(DSUM|DAVERAGE|DCOUNT|DCOUNTA|DGET|DMAX|DMIN|DPRODUCT|DSTDEV|DSTDEVP|DVAR|DVARP)\s*\(/;

/* ── ①紙①＝`kansuu46` の D系（材料は A1:B5 と D1:D2 だけ） ── */
const 紙1 = path.join(ROOT, 'docs/measured/kansuu46');
const 組1 = new Map();
for (const f of fs.readdirSync(紙1).filter((x) => x.endsWith('.tsv'))) {
  if (f === 'golden-dkei6-kane3-2026-09-15.tsv') continue;
  for (const l of 字を読む(path.join(紙1, f)).split(/\r?\n/)) {
    if (!l || l.startsWith('#')) continue;
    const c = l.split('\t');
    if (c[1] && c[1].startsWith('=') && D系.test(c[1].slice(1))) 組1.set(c[1], c[2]);
  }
}

/* ── ②紙②＝26行の うち D系（条件は 道具の 字から） ── */
const 紙2 = path.join(紙1, 'golden-dkei6-kane3-2026-09-15.tsv');
const 道具2 = path.join(紙1, 'toru-dkei6-kane3.ps1');
const 材料 = [];
const 組2 = [];
for (const l of 字を読む(紙2).split(/\r?\n/)) {
  if (l.startsWith('#材料')) { const c = l.split('\t'); 材料.push({ マス: c[1], 値: c[2], 型: c[3] }); continue; }
  if (!l || l.startsWith('#')) continue;
  const c = l.split('\t');
  if (c.length >= 4 && c[1].startsWith('=') && D系.test(c[1].slice(1))) 組2.push({ 札: c[0], 式: c[1], 実: c[2] });
}
const 置き表 = new Map();
for (const m of 字を読む(道具2).matchAll(/札 = '([^']*)';\s*置き = \[ordered\]@\{([^}]*)\};/g)) {
  const 置き = [];
  for (const p of m[2].matchAll(/'([A-Z]+\d+)'\s*=\s*('([^']*)'|-?[\d.]+)/g)) {
    置き.push({ マス: p[1], 値: p[3] !== undefined ? p[3] : p[2], 字か: p[3] !== undefined });
  }
  置き表.set(m[1], 置き);
}

let pass = 0, fail = 0;
const T = (n, f) => {
  try { f(); pass++; console.log('  ✓ ' + n); }
  catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + e.message); }
};

console.log('\n[shiki-kansuu-dkei] ★D系 12個を 自前の 台へ★');

/* ★★分母の 見張り★★ */
const 要る1 = 21, 要る2 = 17, 要る材料 = 15;
T('★分母が 在る★（紙① ' + 要る1 + '行／紙② ' + 要る2 + '行／材料 ' + 要る材料 + '行）', () => {
  if (組1.size !== 要る1) throw new Error('紙① が ' + 組1.size + '行（' + 要る1 + ' のはず）');
  if (組2.length !== 要る2) throw new Error('紙② が ' + 組2.length + '行（' + 要る2 + ' のはず）');
  if (材料.length !== 要る材料) throw new Error('材料が ' + 材料.length + '行（' + 要る材料 + ' のはず）');
  const 無い = 組2.filter((q) => !置き表.has(q.札)).map((q) => q.札);
  if (無い.length) throw new Error('★紙に 在る 札が 道具に 無い★ … ' + 無い.join(' '));
});

/** ★★紙と 同じ 物を 台から 取る★★（2026-09-15）
 *   ★紙は `.Value2` を 'R'（★丸めない 書き方★）で 取って います★
 *   ★台の `字()` は ★実Excel の 画面と 同じ 15桁★に 丸めます★（＝★台の 方が 正しい★）
 *     `=DSTDEV(…)` … 台の 字 `1.29099444873581` ／ 紙 `1.2909944487358056`
 *   ⇒★★丸めた 字と 丸めない 値を 比べると ★偽の 負け★に なります★★
 *   ⇒★紙が 数なら ★台の 生の 値★と 比べる／誤り・字なら `字()` と 比べる★
 *   ★これは「中の 数が 同じ＝同じ」では ありません★
 *     ＝★紙が 持って いるのが ★生の 値★だから 生で 比べる★ という 話です */
function 台から取る(h, 名, 紙の値) {
  if (紙の値 !== '' && isFinite(Number(紙の値))) {
    var v = h.値(名);
    if (v && v.型 === '数') return String(v.値);
  }
  return String(h.字(名));
}

/** ★紙の 材料を 台に 置く★（＋条件） */
function 台を建てる(置き) {
  const h = H.表();
  for (const m of 材料) h.打つ(m.マス, m.型 === '空' ? '' : String(m.値));
  for (const p of (置き || [])) h.打つ(p.マス, String(p.値));
  return h;
}

/* ── 紙①（条件は 紙の 中に 書いて ある A1:A2 や D1:D2） ── */
for (const [式, 実] of [...組1.entries()].sort()) {
  T('①' + 式, () => {
    const h = 台を建てる(null);
    h.打つ('J1', 式);
    const 出 = 台から取る(h, 'J1', 実);
    if (出 !== String(実)) throw new Error('うち ★' + 出 + '★ ／実Excel ★' + 実 + '★');
  });
}
/* ── 紙②（条件は 道具から） ── */
for (const q of 組2) {
  T('②' + q.札 + '  ' + q.式, () => {
    const h = 台を建てる(置き表.get(q.札));
    h.打つ('J1', q.式);
    const 出 = 台から取る(h, 'J1', q.実);
    if (出 !== String(q.実)) throw new Error('うち ★' + 出 + '★ ／実Excel ★' + q.実 + '★');
  });
}

console.log('\n' + pass + ' passed, ' + fail + ' failed');

if (process.argv.includes('--self-test')) {
  /* ★★台が 知って いる D系を 数える★★（★字では なく 呼んで 見る★） */
  const K = require_(path.join(ROOT, 'lib/shiki-kansuu.js'));
  const 名 = ['DSUM', 'DAVERAGE', 'DCOUNT', 'DCOUNTA', 'DGET', 'DMAX', 'DMIN',
    'DPRODUCT', 'DSTDEV', 'DSTDEVP', 'DVAR', 'DVARP'];
  const 知る = 名.filter((n) => typeof (K.表 || {})[n] === 'function');
  console.log('\n[self-test] ★台が 知る D系★ ' + 知る.length + '/12 … ' + 知る.join(' '));
  console.log('[self-test] ★物差しは 実Excel★ 紙① ' + 組1.size + '行 ／ 紙② ' + 組2.length + '行');
  console.log('[self-test] ★見て いない★ 画面／3列以上の 条件／3行以上／日付の 条件');
  process.exit(知る.length === 12 ? 0 : 1);
}
process.exit(fail ? 1 : 0);
