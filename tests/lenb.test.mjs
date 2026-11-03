/* lenb.test.mjs — ★LENB ／ LEFTB ／ RIGHTB ／ MIDB を 守る★（2026-09-18）
 *
 *  ★★前は お客さんの 画面で #NAME? でした★★（★台に 無い 56個の うちの 4個★）
 *
 *  ★★紙（実Excel の 実測・12本）★★
 *    `docs/measured/golden-kansuu-7kaime-2026-09-18.tsv`
 *    ★11 / 12 が 聞く 前の 見込み どおり★
 *    ★★外れた 1本＝`MIDB(ア&"A",2,2)`★★ … 私は 1文字と 見込み、実は ★2文字★
 *      ⇒★★この 1本を 名指しで 押さえます★★（★次に 誰かが 1文字に したら 赤★）
 *
 *  ★★この 見張りが 守る 物（★合った 分だけ★）★★
 *    ①★紙の 12本★（★紙から 読みます★）
 *    ②★★2バイトの 字を 半分で 切ると 空白 1文字★★（★実測★）
 *    ③★★外れた 1本（MIDB）★★
 *    ④★隣（LEN・LEFT・RIGHT・MID）を 壊して いない★
 *
 *  ★★守って いない 物★★
 *    ・★1バイトか 2バイトかを 測ったのは ★3つだけ★★（半角A ／ 全角ア ／ 半角ｱ）
 *    ・★漢字・ひらがな・絵文字・サロゲート★
 *    ・★長さを 省いた 時★／★負の 長さ★／★始まりが 字数を 越える★
 *    ⇒`docs/measured/kansuu46/lenb-kiku-koto.md`
 *
 *  使い方: node tests/lenb.test.mjs
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
const ア = String.fromCodePoint(12354);

console.log('\nlenb.test.mjs ★LENB ／ LEFTB ／ RIGHTB ／ MIDB★\n');

/* ═══ ① ★紙から 読んで 突き合わせる★ ═══ */
T('★紙の 12本と 合う★（★紙から 読みます★）', () => {
  const 紙 = path.join(ROOT, 'docs/measured/golden-kansuu-7kaime-2026-09-18.tsv');
  const 行 = fs.readFileSync(紙, 'utf-8').replace(/^﻿/, '').split(/\r?\n/)
    .filter((l) => l && !l.startsWith('#')).map((l) => l.split('\t'))
    .filter((c) => /^=(LENB|LEN|UNICODE)\(/.test(String(c[1]).trim())
      && /(LENB|LEFTB|RIGHTB|MIDB)\(/.test(String(c[1]).trim()));
  if (行.length !== 12) throw new Error('★紙に 12本 在る はずが ' + 行.length + '本★');
  /* ★式を そのまま 読む★（UNICHAR(12354) は ア に 直す） */
  const 直す = (s) => s.replace(/UNICHAR\((\d+)\)/g, (_, n) => JSON.stringify(String.fromCodePoint(Number(n))));
  let 押した = 0;
  for (const c of 行) {
    const 式 = 直す(String(c[1]).trim());
    const 出 = 押す(式);
    const 正 = Number(String(c[3]).trim());
    if (出 !== 正) throw new Error('★' + c[1] + '★ 出た ' + 出 + ' ／ 実Excel ' + 正);
    押した++;
  }
  if (押した !== 12) throw new Error('★12本 押す はずが ' + 押した + '本★');
  console.log('      … ★12 / 12★（紙から 読んだ 分母）');
});

/* ★ごく 小さい 読み取り★（★紙の 12本の 形だけ★） */
function 押す(式) {
  let m;
  if ((m = /^=LENB\((.+)\)$/.exec(式))) return 引('LENB', 字(字にする(m[1])));
  if ((m = /^=LEN\((LEFTB|RIGHTB|MIDB)\((.+)\)\)$/.exec(式))) {
    return String(呼ぶB(m[1], m[2])).length;
  }
  if ((m = /^=UNICODE\((LEFTB|RIGHTB|MIDB)\((.+)\)\)$/.exec(式))) {
    const s = String(呼ぶB(m[1], m[2]));
    if (!s.length) throw new Error('★空の 字★');
    return s.codePointAt(0);
  }
  throw new Error('★式が 読めません★ ' + 式);
}
function 呼ぶB(名, 中) {
  const 並 = 割る(中);
  const a = [字(字にする(並[0]))];
  for (let i = 1; i < 並.length; i++) a.push(数(Number(並[i])));
  return 引(名, ...a);
}
function 割る(s) {
  const 出 = []; let 深 = 0, 今 = '', 引用 = false;
  for (const ch of s) {
    if (ch === '"') 引用 = !引用;
    if (!引用 && ch === '(') 深++;
    if (!引用 && ch === ')') 深--;
    if (!引用 && ch === ',' && 深 === 0) { 出.push(今); 今 = ''; continue; }
    今 += ch;
  }
  出.push(今); return 出;
}
function 字にする(s) {
  /* `"ア"&"A"` の ような 形を ほどく */
  return s.split('&').map((p) => {
    const t = p.trim();
    const m = /^"(.*)"$/.exec(t);
    if (!m) throw new Error('★字が 読めません★ ' + t);
    return m[1];
  }).join('');
}

/* ═══ ② ★半分で 切ると 空白★（★実測★）═══ */
T('★2バイトの 字を 半分で 切ると 空白 1文字★', () => {
  const r = 引('LEFTB', 字(ア), 数(1));
  if (r.length !== 1) throw new Error('1文字の はず（出た ' + r.length + '）');
  if (r.codePointAt(0) !== 32) throw new Error('★空白（32）の はず★（出た ' + r.codePointAt(0) + '）');
});

/* ═══ ③ ★★外れた 1本を 名指しで 押さえる★★ ═══ */
T('★★MIDB(ア&"A",2,2) は 2文字★★（★私の 見込みは 1で 外れました★）', () => {
  const r = 引('MIDB', 字(ア + 'A'), 数(2), 数(2));
  if (r.length !== 2) {
    throw new Error('★2文字の はず★（★後ろ半分の 空白 ＋ A★／出た ' + r.length + '）');
  }
  if (r.codePointAt(0) !== 32) throw new Error('1文字目は 空白（32）の はず');
  if (r.codePointAt(1) !== 65) throw new Error('2文字目は A の はず');
});

/* ═══ ④ ★隣を 壊して いない★ ═══ */
/* ★★RIGHTB が ★右から★ 取って いるか★★
     ★なぜ 要るか★ … 紙の 1本は `LEN(RIGHTB(ア&"A",1))` ＝ 1 だけで
       ★LEFTB に すり替えても 空白 1文字で ★同じ 1★ に なります★
       （★わざと 壊したら 赤に ならなかった★）
     ⇒★半角だけの 字で ★左右を 見分けます★（ASCII は 1バイト＝1文字）
     ★これは 半分の 話を 使わない ので ★未測定に 踏み込みません★ */
T('★RIGHTB は 右から 取る★（★LEFTB と 見分ける★）', () => {
  if (引('RIGHTB', 字('ABC'), 数(1)) !== 'C') throw new Error('RIGHTB("ABC",1) ＝ C');
  if (引('LEFTB', 字('ABC'), 数(1)) !== 'A') throw new Error('LEFTB("ABC",1) ＝ A');
  if (引('RIGHTB', 字('ABC'), 数(2)) !== 'BC') throw new Error('RIGHTB("ABC",2) ＝ BC');
});

T('★隣（LEN・LEFT・RIGHT・MID）が 生きて いる★', () => {
  if (F.呼ぶ('LEN', [字('ABC')], {}).値 !== 3) throw new Error('LEN("ABC") ＝ 3');
  if (F.呼ぶ('LEFT', [字('ABC'), 数(2)], {}).値 !== 'AB') throw new Error('LEFT("ABC",2) ＝ AB');
  if (F.呼ぶ('RIGHT', [字('ABC'), 数(1)], {}).値 !== 'C') throw new Error('RIGHT("ABC",1) ＝ C');
  if (F.呼ぶ('MID', [字('ABC'), 数(2), 数(1)], {}).値 !== 'B') throw new Error('MID("ABC",2,1) ＝ B');
});

/* ═══ ⑤ ★台が 知る 数★ ═══ */
T('★台が 知る 数★（★4個 増えました★）', () => {
  const n = Object.keys(F.表).length;
  console.log('      … ★' + n + '個★（2026-09-18 の 直しで 400 → 404）');
  if (n < 404) throw new Error('★台が 知る 数が ' + n + 'に 減りました★（404以上の はず）');
  for (const k of ['LENB', 'LEFTB', 'RIGHTB', 'MIDB']) {
    if (!F.表[k]) throw new Error('★' + k + ' が 表から 消えて います★');
  }
});

console.log('\n' + pass + ' passed, ' + fail + ' failed');
if (fail) process.exit(1);
