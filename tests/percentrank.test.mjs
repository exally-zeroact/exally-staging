/* percentrank.test.mjs — ★PERCENTRANK を 守る★（2026-09-18）
 *
 *  ★★前は お客さんの 画面で #NAME? でした★★（★台に 無い 59個の うちの 1個★）
 *
 *  ★★紙（実Excel の 実測・6本）★★
 *    `docs/measured/golden-kansuu-7kaime-2026-09-18.tsv`（材料 C1:C5 ＝ 1,3,5,7,9）
 *    ★★6本 とも 聞く 前の 見込み どおりでした★★（★当て推量 0★）
 *
 *  ★★この 見張りが 守る 物（★合った 分だけ★）★★
 *    ①★紙の 6本★（★紙から 読みます★）
 *    ②★間は 線で 埋める★（(i ＋ 端の 割合) ÷ (n−1)）
 *    ③★範囲の 外は #N/A★
 *    ④★隣（PERCENTILE・QUARTILE）を 壊して いない★
 *
 *  ★★守って いない 物★★
 *    ・★有効桁の 切り方★（★切り捨てに して います／測って いません★）
 *    ・★有効桁 0・負★／★同じ 数が 並ぶ★／★字・空が 混ざる★
 *    ⇒`docs/measured/kansuu46/percentrank-kiku-koto.md`
 *
 *  使い方: node tests/percentrank.test.mjs
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
const 四角 = (a) => ({ 種: '四角', 行数: a.length, 列数: 1, 並び: a.map((v) => K.数(v)) });
const C = () => 四角([1, 3, 5, 7, 9]);
const 引 = (...a) => {
  const r = F.呼ぶ('PERCENTRANK', a, {});
  if (r === null) throw new Error('★台が PERCENTRANK を 知りません★');
  return r.値;
};

console.log('\npercentrank.test.mjs ★PERCENTRANK★\n');

/* ═══ ① ★紙から 読んで 突き合わせる★ ═══ */
T('★紙の 6本と 合う★（★紙から 読みます★）', () => {
  const 紙 = path.join(ROOT, 'docs/measured/golden-kansuu-7kaime-2026-09-18.tsv');
  const 行 = fs.readFileSync(紙, 'utf-8').replace(/^﻿/, '').split(/\r?\n/)
    .filter((l) => l && !l.startsWith('#')).map((l) => l.split('\t'))
    .filter((c) => /^=PERCENTRANK\(/.test(String(c[1]).trim()));
  if (行.length !== 6) throw new Error('★紙に 6本 在る はずが ' + 行.length + '本★');
  let 押した = 0;
  for (const c of 行) {
    const m = /^=PERCENTRANK\(C1:C5,(-?\d+)(?:,(-?\d+))?\)$/.exec(String(c[1]).trim());
    if (!m) throw new Error('★紙の 式が 読めません★ ' + c[1]);
    const 引数 = [C(), 直(Number(m[1]))];
    if (m[2] !== undefined) 引数.push(直(Number(m[2])));
    const 出 = 引(...引数);
    const 正 = String(c[3]).trim();
    if (String(出) !== 正) throw new Error('★' + c[1] + '★ 出た ' + 出 + ' ／ 実Excel ' + 正);
    押した++;
  }
  if (押した !== 6) throw new Error('★6本 押す はずが ' + 押した + '本★');
  console.log('      … ★6 / 6★（紙から 読んだ 分母）');
});

/* ═══ ② ★間は 線で 埋める★（★紙に 無い 分は 計算の 決まりとして★）═══ */
T('★間は 線で 埋める★', () => {
  const 組 = [[2, 0.125], [3, 0.25], [4, 0.375], [6, 0.625], [7, 0.75], [8, 0.875]];
  for (const [x, 正] of 組) {
    const 出 = 引(C(), 直(x));
    if (Math.abs(出 - 正) > 1e-12) throw new Error('PERCENTRANK(C,' + x + ') ＝ ' + 出 + '（' + 正 + 'の はず）');
  }
  console.log('      … ★' + 組.length + ' / ' + 組.length + '★');
});

/* ═══ ③ ★範囲の 外は #N/A★ ═══ */
T('★範囲の 外は #N/A★', () => {
  if (引(C(), 直(10)) !== '#N/A') throw new Error('上の 外は #N/A');
  if (引(C(), 直(0)) !== '#N/A') throw new Error('下の 外は #N/A');
  /* ★端は 入る★ */
  if (引(C(), 直(1)) !== 0) throw new Error('一番 下は 0');
  if (引(C(), 直(9)) !== 1) throw new Error('一番 上は 1');
});

/* ═══ ④ ★有効桁（★切り捨てに して います／測って いません★）★ ═══ */
T('★有効桁は 切り捨て★（★これは 実測では ありません★）', () => {
  /* {1,2,3,4} で x=3.5 ⇒ (2 ＋ 0.5)/3 ＝ 0.8333… */
  const D = 四角([1, 2, 3, 4]);
  const 三桁 = 引(D, 直(3.5));
  if (三桁 !== 0.833) throw new Error('★既定 3桁で 0.833 の はず★（丸めなら 0.833・切り捨ても 0.833／出た ' + 三桁 + '）');
  const 二桁 = 引(D, 直(3.5), 直(2));
  if (二桁 !== 0.83) throw new Error('★2桁で 0.83 の はず★（出た ' + 二桁 + '）');
  /* ★★丸めと 切り捨てが ★本当に 分かれる★ 組★★
       (2 ＋ 0.9)/3 ＝ 0.96666…  ⇒ 1桁 … ★切り捨て 0.9 ／ 丸め 1★
       ★★これは 実測では ありません★★＝★「切り捨てに して ある」事だけを 押さえます★
       ⇒★実Excel に 聞いたら ★この 見張りごと★ 直して ください★
       ⇒`docs/measured/kansuu46/percentrank-kiku-koto.md` の ③に 名指しで */
  const 分かれる = 引(D, 直(3.9), 直(1));
  if (分かれる !== 0.9) {
    throw new Error('★切り捨てに なって いません★（丸めなら 1／出た ' + 分かれる + '）');
  }
  const 分かれる2 = 引(D, 直(3.9), 直(2));   /* 0.96666… ⇒ 切り捨て 0.96 ／ 丸め 0.97 */
  if (分かれる2 !== 0.96) {
    throw new Error('★2桁でも 切り捨ての はず★（丸めなら 0.97／出た ' + 分かれる2 + '）');
  }
});

/* ═══ ⑤ ★隣を 壊して いない★ ═══ */
T('★隣（PERCENTILE・QUARTILE）が 生きて いる★', () => {
  const p = F.呼ぶ('PERCENTILE', [C(), 直(0.5)], {});
  if (!p || p.値 !== 5) throw new Error('PERCENTILE(C,0.5) ＝ 5（出た ' + (p && p.値) + '）');
  const q = F.呼ぶ('QUARTILE', [C(), 直(2)], {});
  if (!q || q.値 !== 5) throw new Error('QUARTILE(C,2) ＝ 5（出た ' + (q && q.値) + '）');
});

/* ═══ ⑥ ★台が 知る 数★ ═══ */
T('★台が 知る 数★（★1個 増えました★）', () => {
  const n = Object.keys(F.表).length;
  console.log('      … ★' + n + '個★（2026-09-18 の 直しで 397 → 398）');
  if (n < 398) throw new Error('★台が 知る 数が ' + n + 'に 減りました★（398以上の はず）');
  if (!F.表.PERCENTRANK) throw new Error('★PERCENTRANK が 表から 消えて います★');
});

console.log('\n' + pass + ' passed, ' + fail + ' failed');
if (fail) process.exit(1);
