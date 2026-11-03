/* aggregate.test.mjs — ★AGGREGATE を 守る★（2026-09-18）
 *
 *  ★★前は お客さんの 画面で #NAME? でした★★（★台に 無い 50個の うちの 1個★）
 *
 *  ★★★この 見張りの 一番 の 仕事 ── ★外した 2本★★★★
 *    ①`=AGGREGATE(9,★7★,B1:B5)` → ★12★
 *       ＝★★選択 7 は ★誤りも★ 無視します★★（「隠した 行だけ」では ない）
 *       ⇒私は #DIV/0! と 見込みました＝外れ
 *    ②`=AGGREGATE(★19★,6,B1:B5,1)` → ★1.25★
 *       ＝★★機能 19 は QUARTILE.★EXC★ です★★（INC では ない）
 *       ⇒私は 1.75 と 見込みました＝外れ
 *    ⇒★★次の 人も 必ず 同じ 所で 間違えます★★
 *
 *  ★★紙（実Excel の 実測・12本）★★
 *    `docs/measured/golden-kansuu-7kaime-2026-09-18.tsv`
 *    ★材料★ A1:A5 ＝ 1,2,3,4,5 ／ ★B1:B5 ＝ 1,2,(#DIV/0!),4,5★
 *
 *  ★★守って いない 物★★
 *    ・★選択 1・3・5★（★2・6・7 と 同じに して います／未測定★）
 *    ・★隠した 行★（★行を 隠さないと 4 と 7 の 違いが 出ません★）
 *    ・★入れ子の AGGREGATE／SUBTOTAL★
 *    ⇒`docs/measured/kansuu46/aggregate-kiku-koto.md`
 *
 *  使い方: node tests/aggregate.test.mjs
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
const 数 = (n) => ({ 種: '直', 値: K.数(n) });
const 誤値 = (s) => ({ 型: '誤', 値: s });
const A = () => ({ 種: '四角', 行数: 5, 列数: 1, 並び: [1, 2, 3, 4, 5].map(K.数) });
const B = () => ({ 種: '四角', 行数: 5, 列数: 1, 並び: [K.数(1), K.数(2), 誤値('#DIV/0!'), K.数(4), K.数(5)] });
const 引 = (...a) => {
  const r = F.呼ぶ('AGGREGATE', a, {});
  if (r === null) throw new Error('★台が AGGREGATE を 知りません★');
  return r.値;
};

console.log('\naggregate.test.mjs ★AGGREGATE★\n');

/* ═══ ① ★★外した 2本（一番 上）★★ ═══ */
T('★★選択 7 は 誤りも 無視する★★（★私の 見込みは #DIV/0! で 外れました★）', () => {
  const r = 引(数(9), 数(7), B());
  if (r !== 12) {
    throw new Error('★12 の はず★（★「隠した 行だけ」と 読むと #DIV/0! に なります★／出た ' + r + '）');
  }
});
T('★★選択 4 は 誤りを 無視しない★★（★7 と 違います★）', () => {
  const r = 引(数(9), 数(4), B());
  if (r !== '#DIV/0!') throw new Error('★#DIV/0! の はず★（出た ' + r + '）');
});
T('★★機能 19 は QUARTILE.EXC★★（★私の 見込みは 1.75（INC）で 外れました★）', () => {
  const r = 引(数(19), 数(6), B(), 数(1));
  if (Math.abs(r - 1.25) > 1e-12) {
    throw new Error('★1.25 の はず★（★INC なら 1.75★／出た ' + r + '）');
  }
});

/* ═══ ② ★紙から 読んで 突き合わせる★ ═══ */
T('★紙の 12本と 合う★（★紙から 読みます★）', () => {
  const 紙 = path.join(ROOT, 'docs/measured/golden-kansuu-7kaime-2026-09-18.tsv');
  const 行 = fs.readFileSync(紙, 'utf-8').replace(/^﻿/, '').split(/\r?\n/)
    .filter((l) => l && !l.startsWith('#')).map((l) => l.split('\t'))
    .filter((c) => /^=AGGREGATE\(/.test(String(c[1]).trim()));
  if (行.length !== 12) throw new Error('★紙に 12本 在る はずが ' + 行.length + '本★');
  const 範 = (s) => (s === 'A1:A5' ? A() : B());
  let 押した = 0;
  for (const c of 行) {
    const m = /^=AGGREGATE\((\d+),(\d+),([AB]1:[AB]5)(?:,([AB]1:[AB]5))?(?:,(\d+))?\)$/.exec(String(c[1]).trim());
    if (!m) throw new Error('★紙の 式が 読めません★ ' + c[1]);
    const a = [数(Number(m[1])), 数(Number(m[2])), 範(m[3])];
    if (m[4]) a.push(範(m[4]));
    if (m[5] !== undefined) a.push(数(Number(m[5])));
    const 出 = 引(...a);
    const 正 = String(c[3]).trim();
    const 同 = /^#/.test(正) ? (String(出) === 正) : (Math.abs(Number(出) - Number(正)) < 1e-12);
    if (!同) throw new Error('★' + c[1] + '★ 出た ' + 出 + ' ／ 実Excel ' + 正);
    押した++;
  }
  if (押した !== 12) throw new Error('★12本 押す はずが ' + 押した + '本★');
  console.log('      … ★12 / 12★（紙から 読んだ 分母）');
});

/* ═══ ③ ★k が 要る／要らない★ ═══ */
T('★k が 要るのに 無い／要らないのに 在る は #VALUE!★', () => {
  if (引(数(14), 数(6), B()) !== '#VALUE!') throw new Error('14 に k が 無い ⇒ #VALUE!');
  if (引(数(9), 数(6), B(), 数(2)) !== '#VALUE!') throw new Error('9 に k が 在る ⇒ #VALUE!');
  /* ★範囲 2つは 通ります★（実測） */
  if (引(数(9), 数(0), A(), B()) !== '#DIV/0!') throw new Error('範囲 2つは 通る');
});

/* ═══ ④ ★機能 19個が 全部 呼べる★（★分母を 出す★）═══ */
T('★機能 1〜19 が 全部 答える★（★#VALUE! を 返さない★）', () => {
  const kが要る = [14, 15, 16, 17, 18, 19];
  let 答えた = 0;
  for (let f = 1; f <= 19; f++) {
    const a = [数(f), 数(6), A()];
    if (kが要る.includes(f)) a.push(数(f >= 16 && f <= 18 ? 0.5 : (f === 19 ? 1 : 2)));
    const r = 引(...a);
    if (r === '#VALUE!') throw new Error('★機能 ' + f + ' が #VALUE!★');
    答えた++;
  }
  if (答えた !== 19) throw new Error('★19個 見る はずが ' + 答えた + '個★');
  console.log('      … ★19 / 19★');
});

/* ═══ ⑤ ★隣を 壊して いない★ ═══ */
T('★隣（SUBTOTAL・SUM・LARGE）が 生きて いる★', () => {
  const s = F.呼ぶ('SUBTOTAL', [数(9), A()], {});
  if (!s || s.値 !== 15) throw new Error('SUBTOTAL(9,A) ＝ 15（出た ' + (s && s.値) + '）');
  const u = F.呼ぶ('SUM', [A()], {});
  if (!u || u.値 !== 15) throw new Error('SUM(A) ＝ 15');
  const l = F.呼ぶ('LARGE', [A(), 数(2)], {});
  if (!l || l.値 !== 4) throw new Error('LARGE(A,2) ＝ 4');
});

/* ═══ ⑥ ★台が 知る 数★ ═══ */
T('★台が 知る 数★（★1個 増えました★）', () => {
  const n = Object.keys(F.表).length;
  console.log('      … ★' + n + '個★（2026-09-18 の 直しで 406 → 407）');
  if (n < 407) throw new Error('★台が 知る 数が ' + n + 'に 減りました★（407以上の はず）');
  if (!F.表.AGGREGATE) throw new Error('★AGGREGATE が 表から 消えて います★');
});

console.log('\n' + pass + ' passed, ' + fail + ' failed');
if (fail) process.exit(1);
