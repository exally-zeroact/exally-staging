/* textafter.test.mjs — ★TEXTAFTER ／ TEXTBEFORE を 守る★（2026-09-18）
 *
 *  ★★前は お客さんの 画面で #NAME? でした★★（★台に 無い 52個の うちの 2個★）
 *
 *  ★★★この 見張りの 一番 の 仕事★★★
 *    ★★私たちの 見込みが ★逆★だった 2本を 先に 押さえる★★
 *      `=TEXTAFTER("a-B-c","b")`     … 見込み "-c" ／ ★実は #N/A★
 *      `=TEXTAFTER("a-B-c","b",1,1)` … 見込み #N/A ／ ★実は "-c"★
 *    ⇒★★既定は ★大小を 見る★／`1` が ★大小を 見ない★★★
 *    ⇒★SEARCH（見ない）と FIND（見る）の 感じとは ★別★★
 *    ⇒★★次の 人も 必ず 逆に します★★
 *
 *  ★★紙（実Excel の 実測・12本）★★
 *    `docs/measured/golden-kansuu-7kaime-2026-09-18.tsv`
 *    ★9 / 12 が 聞く 前の 見込み どおり★（★外した 3本を 下で 名指し★）
 *
 *  ★★守って いない 物★★
 *    ・★5つ目（match_end）★／★6つ目（if_not_found）★
 *    ・★区切りが 並び★／★空の 字を 探す TEXTAFTER★／★0 番目★
 *    ⇒`docs/measured/kansuu46/textafter-kiku-koto.md`
 *
 *  使い方: node tests/textafter.test.mjs
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
const 字 = (s) => ({ 種: '直', 値: K.字(s) });
const 数 = (n) => ({ 種: '直', 値: K.数(n) });
const 引 = (名, ...a) => {
  const r = F.呼ぶ(名, a, {});
  if (r === null) throw new Error('★台が ' + 名 + ' を 知りません★');
  return r.値;
};

console.log('\ntextafter.test.mjs ★TEXTAFTER ／ TEXTBEFORE★\n');

/* ═══ ① ★★逆だった 2本（★一番 先に★）★★ ═══ */
T('★★既定は 大小を 見る★★（★私たちの 見込みは 逆でした★）', () => {
  const r = 引('TEXTAFTER', 字('a-B-c'), 字('b'));
  if (r !== '#N/A') {
    throw new Error('★#N/A の はず★（★大小を 見ない に すると "-c" に なります★／出た ' + JSON.stringify(r) + '）');
  }
});
T('★★`1` が 大小を 見ない★★（★私たちの 見込みは 逆でした★）', () => {
  const r = 引('TEXTAFTER', 字('a-B-c'), 字('b'), 数(1), 数(1));
  if (r !== '-c') {
    throw new Error('★"-c" の はず★（★大小を 見る に すると #N/A に なります★／出た ' + JSON.stringify(r) + '）');
  }
});
T('★★`0` は 大小を 見る★★（★既定と 同じ★）', () => {
  const r = 引('TEXTAFTER', 字('a-B-c'), 字('b'), 数(1), 数(0));
  if (r !== '#N/A') throw new Error('★#N/A の はず★（出た ' + JSON.stringify(r) + '）');
});

/* ═══ ② ★空の 区切り★（★#VALUE! では ない★）═══ */
T('★空の 区切りで TEXTBEFORE は 空の 字★（★#VALUE! では ない★）', () => {
  const r = 引('TEXTBEFORE', 字('abc'), 字(''));
  if (r !== '') throw new Error('★空の 字の はず★（出た ' + JSON.stringify(r) + '）');
});

/* ═══ ③ ★紙から 読んで 突き合わせる★ ═══ */
T('★紙の 12本と 合う★（★紙から 読みます★）', () => {
  const 紙 = path.join(ROOT, 'docs/measured/golden-kansuu-7kaime-2026-09-18.tsv');
  const 行 = fs.readFileSync(紙, 'utf-8').replace(/^﻿/, '').split(/\r?\n/)
    .filter((l) => l && !l.startsWith('#')).map((l) => l.split('\t'))
    .filter((c) => /^=TEXT(AFTER|BEFORE)\(/.test(String(c[1]).trim()));
  if (行.length !== 12) throw new Error('★紙に 12本 在る はずが ' + 行.length + '本★');
  let 押した = 0;
  for (const c of 行) {
    const m = /^=TEXT(AFTER|BEFORE)\("([^"]*)","([^"]*)"(?:,(-?\d+))?(?:,(-?\d+))?\)$/.exec(String(c[1]).trim());
    if (!m) throw new Error('★紙の 式が 読めません★ ' + c[1]);
    const a = [字(m[2]), 字(m[3])];
    if (m[4] !== undefined) a.push(数(Number(m[4])));
    if (m[5] !== undefined) a.push(数(Number(m[5])));
    const 出 = 引('TEXT' + m[1], ...a);
    /* ★紙の「出る字」欄と 比べます★（空の 字は 空欄） */
    const 正 = (c[3] === undefined) ? '' : String(c[3]);
    if (String(出) !== 正) {
      throw new Error('★' + c[1] + '★ 出た ' + JSON.stringify(出) + ' ／ 実Excel ' + JSON.stringify(正));
    }
    押した++;
  }
  if (押した !== 12) throw new Error('★12本 押す はずが ' + 押した + '本★');
  console.log('      … ★12 / 12★（紙から 読んだ 分母）');
});

/* ═══ ④ ★番目（後ろから・行き過ぎ）★ ═══ */
T('★番目は 負なら 後ろから／行き過ぎたら #N/A★', () => {
  if (引('TEXTAFTER', 字('a-b-c'), 字('-'), 数(-1)) !== 'c') throw new Error('-1 は 後ろから 1つ目');
  if (引('TEXTBEFORE', 字('a-b-c'), 字('-'), 数(-1)) !== 'a-b') throw new Error('BEFORE の -1');
  if (引('TEXTAFTER', 字('a-b-c'), 字('-'), 数(5)) !== '#N/A') throw new Error('行き過ぎは #N/A');
  if (引('TEXTAFTER', 字('a-b-c'), 字('-'), 数(-5)) !== '#N/A') throw new Error('後ろへ 行き過ぎも #N/A');
});

/* ═══ ⑤ ★隣を 壊して いない★ ═══ */
T('★隣（TEXTJOIN・FIND・SEARCH）が 生きて いる★', () => {
  const j = F.呼ぶ('TEXTJOIN', [字('-'), { 種: '直', 値: K.真偽 ? K.真偽(true) : K.数(1) }, 字('a'), 字('b')], {});
  if (!j || j.値 !== 'a-b') throw new Error('TEXTJOIN ＝ a-b（出た ' + (j && j.値) + '）');
  const f = F.呼ぶ('FIND', [字('b'), 字('abc')], {});
  if (!f || f.値 !== 2) throw new Error('FIND("b","abc") ＝ 2');
  const s = F.呼ぶ('SEARCH', [字('B'), 字('abc')], {});
  if (!s || s.値 !== 2) throw new Error('SEARCH("B","abc") ＝ 2（★大小を 見ない★）');
});

/* ═══ ⑥ ★台が 知る 数★ ═══ */
T('★台が 知る 数★（★2個 増えました★）', () => {
  const n = Object.keys(F.表).length;
  console.log('      … ★' + n + '個★（2026-09-18 の 直しで 404 → 406）');
  if (n < 406) throw new Error('★台が 知る 数が ' + n + 'に 減りました★（406以上の はず）');
  if (!F.表.TEXTAFTER || !F.表.TEXTBEFORE) throw new Error('★表から 消えて います★');
});

console.log('\n' + pass + ' passed, ' + fail + ' failed');
if (fail) process.exit(1);
