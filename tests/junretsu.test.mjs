/* junretsu.test.mjs — ★順列 2つ（PERMUT／PERMUTATIONA）を 守る★（2026-09-18）
 *
 *  ★★前は お客さんの 画面で #NAME? でした★★（★台に 無い 62個の うちの 2個★）
 *
 *  ★★紙（実Excel の 実測）★★ `docs/measured/golden-oddf-to-46ko-2026-09-16.tsv`
 *    `=PERMUT(5,2)` → ★20★ ／ `=PERMUTATIONA(5,2)` → ★25★
 *    ★★紙は この 2本 だけです★★
 *
 *  ★★この 見張りが 守る 物（★合った 分だけ★）★★
 *    ①★紙の 2本★（★紙から 読みます＝手で 書き写しません★）
 *    ②★中の 計算★（PERMUT ＝ n×(n−1)×… ／ PERMUTATIONA ＝ n^k）
 *    ③★切り捨て★（COMBIN と 同じ＝`=COMBIN(2.5,1)` → 2 が 実測）
 *    ④★隣（COMBIN・FACT）を 壊して いない★
 *
 *  ★★守って いない 物（★書かない 見張りは「全部 守った」と 読まれる★）★★
 *    ・★境目（負の 数・k>n・0）は ★実Excel に 聞いて いません★★
 *      ＝★COMBIN（実測済み）に 合わせた だけ★
 *      ⇒`docs/measured/kansuu46/junretsu-kiku-koto.md` に 名指しで 載せて あります
 *
 *  使い方: node tests/junretsu.test.mjs
 */
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(path.join(ROOT, 'package.json'));
const F = require_(path.join(ROOT, 'lib/shiki-kansuu.js'));
const K = require_(path.join(ROOT, 'lib/shiki-keisan.js'));

let pass = 0, fail = 0;
const T = (n, f) => {
  try { f(); pass++; console.log('  ✓ ' + n); }
  catch (e) { fail++; console.log('  ✗ ' + n + '\n      ' + e.message); }
};
const 直 = (x) => ({ 種: '直', 値: K.数(x) });
const 引 = (n, ...a) => {
  const r = F.呼ぶ(n, a.map(直), {});
  if (r === null) throw new Error('★台が ' + n + ' を 知りません★');
  return r.値;
};

console.log('\njunretsu.test.mjs ★PERMUT ／ PERMUTATIONA★\n');

/* ═══ ① ★紙から 読んで 突き合わせる★（★手で 書き写しません★）═══ */
T('★紙の 2本と 合う★（★紙から 読みます★）', () => {
  const 紙 = path.join(ROOT, 'docs/measured/golden-oddf-to-46ko-2026-09-16.tsv');
  const 行 = fs.readFileSync(紙, 'utf-8').replace(/^﻿/, '').split(/\r?\n/)
    .filter((l) => l && !l.startsWith('#')).map((l) => l.split('\t'))
    .filter((c) => /^=(PERMUT|PERMUTATIONA)\(/.test(String(c[2]).trim()));
  if (行.length !== 2) throw new Error('★紙に 2本 在る はずが ' + 行.length + '本★');
  let 押した = 0;
  for (const c of 行) {
    const m = /^=([A-Z]+)\((-?[\d.]+),(-?[\d.]+)\)$/.exec(String(c[2]).trim());
    if (!m) throw new Error('★紙の 式が 読めません★ ' + c[2]);
    const 出 = 引(m[1], Number(m[2]), Number(m[3]));
    const 正 = Number(c[3]);
    if (出 !== 正) throw new Error('★' + c[2] + '★ 出た ' + 出 + ' ／ 実Excel ' + 正);
    押した++;
  }
  if (押した !== 2) throw new Error('★2本 押す はずが ' + 押した + '本★');
  console.log('      … ★2 / 2★（紙から 読んだ 分母）');
});

/* ═══ ② ★中の 計算★（★紙に 無い 分は「計算の 決まり」として 守る★）═══ */
T('★PERMUT ＝ n×(n−1)×…×(n−k+1)★', () => {
  const 組 = [[5, 0, 1], [5, 1, 5], [5, 2, 20], [5, 5, 120], [10, 3, 720], [1, 1, 1]];
  for (const [n, k, 正] of 組) {
    const 出 = 引('PERMUT', n, k);
    if (出 !== 正) throw new Error('PERMUT(' + n + ',' + k + ') ＝ ' + 出 + '（' + 正 + 'の はず）');
  }
  console.log('      … ★' + 組.length + ' / ' + 組.length + '★');
});

T('★PERMUTATIONA ＝ n^k★', () => {
  const 組 = [[5, 2, 25], [3, 4, 81], [2, 10, 1024], [7, 0, 1], [0, 0, 1], [0, 3, 0]];
  for (const [n, k, 正] of 組) {
    const 出 = 引('PERMUTATIONA', n, k);
    if (出 !== 正) throw new Error('PERMUTATIONA(' + n + ',' + k + ') ＝ ' + 出 + '（' + 正 + 'の はず）');
  }
  console.log('      … ★' + 組.length + ' / ' + 組.length + '★');
});

/* ═══ ③ ★切り捨て★（COMBIN と 同じ・★COMBIN の 方は 実測★）═══ */
T('★切り捨ては COMBIN と 同じ★', () => {
  if (引('COMBIN', 2.5, 1) !== 2) throw new Error('★COMBIN の 切り捨てが 変わりました★');
  if (引('PERMUT', 2.5, 1) !== 2) throw new Error('PERMUT(2.5,1) ＝ 2 の はず');
  if (引('PERMUT', 5.9, 2.9) !== 20) throw new Error('PERMUT(5.9,2.9) ＝ 20 の はず');
  if (引('PERMUTATIONA', 5.9, 2.9) !== 25) throw new Error('PERMUTATIONA(5.9,2.9) ＝ 25 の はず');
});

/* ═══ ④ ★境目は 測って いません（★「合わせた 先」を 名指しで 押さえる★）═══ */
T('★境目は COMBIN に 合わせて ある（★実測では ありません★）★', () => {
  /* ★この 見張りは「正しい」とは 言って いません★
     ★「COMBIN と 同じ 決めに して ある」だけを 押さえます★
     ⇒★実Excel に 聞いたら ★この 見張りごと★ 直して ください★ */
  if (引('PERMUT', 3, 4) !== '#NUM!') throw new Error('PERMUT(3,4) は #NUM!（COMBIN に 合わせた）');
  if (引('COMBIN', 3, 4) !== '#NUM!') throw new Error('★COMBIN の 決めが 変わりました★');
  if (引('PERMUT', -1, 1) !== '#NUM!') throw new Error('PERMUT(-1,1) は #NUM!');
  if (引('PERMUTATIONA', -1, 2) !== '#NUM!') throw new Error('PERMUTATIONA(-1,2) は #NUM!');
});

/* ═══ ⑤ ★隣を 壊して いない★ ═══ */
T('★隣（COMBIN・FACT）が 生きて いる★', () => {
  if (引('COMBIN', 5, 2) !== 10) throw new Error('COMBIN(5,2) ＝ 10');
  if (引('FACT', 5) !== 120) throw new Error('FACT(5) ＝ 120');
});

/* ═══ ⑥ ★台の 数が 増えた★（★分母を 出す★）═══ */
T('★台が 知る 数★（★2個 増えました★）', () => {
  const n = Object.keys(F.表).length;
  console.log('      … ★' + n + '個★（2026-09-18 の 直しで 394 → 396）');
  if (n < 396) throw new Error('★台が 知る 数が ' + n + 'に 減りました★（396以上の はず）');
  if (!F.表.PERMUT || !F.表.PERMUTATIONA) throw new Error('★順列が 表から 消えて います★');
});

console.log('\n' + pass + ' passed, ' + fail + ' failed');
if (fail) process.exit(1);
