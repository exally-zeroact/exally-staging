/* xmatch.test.mjs — ★XMATCH を 守る★（2026-09-18）
 *
 *  ★★前は お客さんの 画面で #NAME? でした★★（★台に 無い 60個の うちの 1個★）
 *
 *  ★★紙（実Excel の 実測・12本）★★
 *    `docs/measured/golden-kansuu-7kaime-2026-09-18.tsv`（材料 C1:C5 ＝ 1,3,5,7,9）
 *    ★★12本 とも 聞く 前の 見込み どおりでした★★（★当て推量が 0★）
 *
 *  ★★この 見張りが 守る 物（★合った 分だけ★）★★
 *    ①★紙の 12本★（★紙から 読みます＝手で 書き写しません★）
 *    ②★既定は ぴたりだけ（一致 0）★
 *    ③★無い 型は #VALUE! では なく ★#N/A★★（★ここは 間違えやすい★）
 *    ④★隣（MATCH・XLOOKUP）を 壊して いない★
 *
 *  ★★守って いない 物★★
 *    ・★字の 型紙★／★2次元の 表★／★空の 並び★／★誤りを 含む 並び★
 *    ・★一致 -1／1 と 検索 -1／2／-2 の 組み合わせ★
 *    ⇒`docs/measured/kansuu46/xmatch-kiku-koto.md` に 名指しで
 *
 *  使い方: node tests/xmatch.test.mjs
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
/* ★材料★ C1:C5 ＝ 1,3,5,7,9（★紙と 同じ★） */
const 四角 = (a) => ({ 種: '四角', 行数: a.length, 列数: 1, 並び: a.map((v) => K.数(v)) });
const C = () => 四角([1, 3, 5, 7, 9]);
const 引 = (...a) => {
  const r = F.呼ぶ('XMATCH', a, {});
  if (r === null) throw new Error('★台が XMATCH を 知りません★');
  return r.値;
};

console.log('\nxmatch.test.mjs ★XMATCH★\n');

/* ═══ ① ★紙から 読んで 突き合わせる★ ═══ */
T('★紙の 12本と 合う★（★紙から 読みます★）', () => {
  const 紙 = path.join(ROOT, 'docs/measured/golden-kansuu-7kaime-2026-09-18.tsv');
  const 行 = fs.readFileSync(紙, 'utf-8').replace(/^﻿/, '').split(/\r?\n/)
    .filter((l) => l && !l.startsWith('#')).map((l) => l.split('\t'))
    .filter((c) => /^=XMATCH\(/.test(String(c[1]).trim()));
  if (行.length !== 12) throw new Error('★紙に 12本 在る はずが ' + 行.length + '本★');
  let 押した = 0;
  for (const c of 行) {
    const m = /^=XMATCH\((-?\d+),C1:C5(?:,(-?\d+))?(?:,(-?\d+))?\)$/.exec(String(c[1]).trim());
    if (!m) throw new Error('★紙の 式が 読めません★ ' + c[1]);
    const 引数 = [直(Number(m[1])), C()];
    if (m[2] !== undefined) 引数.push(直(Number(m[2])));
    if (m[3] !== undefined) 引数.push(直(Number(m[3])));
    const 出 = 引(...引数);
    const 正 = String(c[3]).trim();          /* ★出る字★（#N/A か 数） */
    if (String(出) !== 正) throw new Error('★' + c[1] + '★ 出た ' + 出 + ' ／ 実Excel ' + 正);
    押した++;
  }
  if (押した !== 12) throw new Error('★12本 押す はずが ' + 押した + '本★');
  console.log('      … ★12 / 12★（紙から 読んだ 分母）');
});

/* ═══ ② ★既定は ぴたりだけ★ ═══ */
T('★引数を 省くと ぴたりだけ（一致 0）★', () => {
  if (引(直(5), C()) !== 3) throw new Error('XMATCH(5,C) ＝ 3');
  if (引(直(4), C()) !== '#N/A') throw new Error('XMATCH(4,C) ＝ #N/A（★既定は ぴたり★）');
});

/* ═══ ③ ★★無い 型は ★一致型と 検索型で 誤りが 違う★★★ ═══
     ★2026-09-18・8枠目で 実測して 分かれました★
       `=XMATCH(4,C1:C5,3)`   -> ★#N/A★   （★一致型が 無い★）
       `=XMATCH(4,C1:C5,0,7)` -> ★#VALUE!★（★検索型が 無い★）
     ★前は ここに「検索型は 測って いない＝一致型に 合わせる」と 書いて いました★
     ⇒★★見立てを 試験に 焼いて いました★★
     ⇒★実測が 出たので 試験の 方を 直します★（★台では なく★） */
T('★無い 型は 一致＝#N/A ／ 検索＝#VALUE!★（★実測で 分かれた★）', () => {
  const a = 引(直(4), C(), 直(3));
  if (a !== '#N/A') throw new Error('一致 3 は #N/A の はず（出た ' + a + '）');
  const b = 引(直(4), C(), 直(0), 直(7));
  if (b !== '#VALUE!') throw new Error('★検索 7 は #VALUE!★（8枠目 実測・出た ' + b + '）');
});

/* ═══ ④ ★二分（2 と -2）が 向きを 見て いる★ ═══ */
T('★二分は 向きを 見る★（昇順に -2 を 当てたら 見つからない）', () => {
  if (引(直(9), C(), 直(0), 直(2)) !== 5) throw new Error('検索 2（昇順 二分）＝ 5');
  if (引(直(9), C(), 直(0), 直(-2)) !== '#N/A') throw new Error('検索 -2 を 昇順に 当てたら #N/A');
});

/* ═══ ⑤ ★隣を 壊して いない★ ═══ */
T('★隣（MATCH）が 生きて いる★', () => {
  const r = F.呼ぶ('MATCH', [直(5), C(), 直(0)], {});
  if (!r || r.値 !== 3) throw new Error('MATCH(5,C,0) ＝ 3（出た ' + (r && r.値) + '）');
});

/* ═══ ⑥ ★台が 知る 数★ ═══ */
T('★台が 知る 数★（★1個 増えました★）', () => {
  const n = Object.keys(F.表).length;
  console.log('      … ★' + n + '個★（2026-09-18 の 直しで 396 → 397）');
  if (n < 397) throw new Error('★台が 知る 数が ' + n + 'に 減りました★（397以上の はず）');
  if (!F.表.XMATCH) throw new Error('★XMATCH が 表から 消えて います★');
});

console.log('\n' + pass + ' passed, ' + fail + ' failed');
if (fail) process.exit(1);
