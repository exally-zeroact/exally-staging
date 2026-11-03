/* asc-dbcs.test.mjs — ★ASC ／ DBCS を 守る★（2026-09-18）
 *
 *  ★★前は お客さんの 画面で #NAME? でした★★（★台に 無い 58個の うちの 2個★）
 *
 *  ★★紙（実Excel の 実測・22行＝11字 × 2）★★
 *    `docs/measured/golden-kansuu-7kaime-2026-09-18.tsv`
 *    ★★22行 とも 聞く 前の 見込み どおりでした★★（★当て推量 0★）
 *    ★字は 直に 書かず `UNICHAR` で 作り、答えも ★コードと 文字数★で 受けました★
 *      ⇒★★この 紙も 試験も 字が 化けません★★
 *
 *  ★★この 見張りが 守る 物（★合った 分だけ★）★★
 *    ①★紙の 22行★（★紙から 読みます★）
 *    ②★★字の 表を 記憶で 書いて いない★★
 *        ＝`String.normalize` の 表を 引いて いる
 *        ＝★62字 を 端から 端まで 通して「1字も 落ちない」を 見ます★
 *    ③★濁点が 分かれる★（ガ → ｶ ＋ ﾞ ＝ 2文字）
 *    ④★隣（UNICODE・UNICHAR・LEN）を 壊して いない★
 *
 *  ★★守って いない 物★★
 *    ・★全角ガ の 2文字目★（★紙は「2文字」までしか 言って いません★）
 *    ・★62字の うち 測ったのは 4字★／半濁点・長音・ヴ・ひらがな
 *    ⇒`docs/measured/kansuu46/asc-dbcs-kiku-koto.md`
 *
 *  使い方: node tests/asc-dbcs.test.mjs
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
const 引 = (名, s) => {
  const r = F.呼ぶ(名, [字(s)], {});
  if (r === null) throw new Error('★台が ' + 名 + ' を 知りません★');
  return r.値;
};

console.log('\nasc-dbcs.test.mjs ★ASC ／ DBCS★\n');

/* ═══ ① ★紙から 読んで 突き合わせる★（★コードと 文字数で★）═══ */
T('★紙の 22行と 合う★（★紙から 読みます★）', () => {
  const 紙 = path.join(ROOT, 'docs/measured/golden-kansuu-7kaime-2026-09-18.tsv');
  const 行 = fs.readFileSync(紙, 'utf-8').replace(/^﻿/, '').split(/\r?\n/)
    .filter((l) => l && !l.startsWith('#')).map((l) => l.split('\t'))
    .filter((c) => /^=(UNICODE|LEN)\((ASC|DBCS)\(UNICHAR\(\d+\)\)\)$/.test(String(c[1]).trim()));
  if (行.length !== 22) throw new Error('★紙に 22行 在る はずが ' + 行.length + '行★');
  let 押した = 0;
  for (const c of 行) {
    const m = /^=(UNICODE|LEN)\((ASC|DBCS)\(UNICHAR\((\d+)\)\)\)$/.exec(String(c[1]).trim());
    const 出字 = 引(m[2], String.fromCodePoint(Number(m[3])));
    if (typeof 出字 !== 'string') throw new Error('★' + c[1] + '★ 字が 返って いません（' + 出字 + '）');
    const 出 = (m[1] === 'UNICODE') ? 出字.codePointAt(0) : 出字.length;
    const 正 = Number(String(c[3]).trim());
    if (出 !== 正) throw new Error('★' + c[1] + '★ 出た ' + 出 + ' ／ 実Excel ' + 正);
    押した++;
  }
  if (押した !== 22) throw new Error('★22行 押す はずが ' + 押した + '行★');
  console.log('      … ★22 / 22★（紙から 読んだ 分母）');
});

/* ═══ ② ★★表を 記憶で 書いて いない＝62字 とも 通る★★ ═══ */
/* ★★この 試験は 「正しい」とは 言って いません★★
     ★表を 記憶で 書いて いない（＝機械の 表を 引いて いる）事だけを 押さえます★
     ★実測は 4字だけ★（ア・ガ・、・空白）
     ★U+FF9E／U+FF9F（濁点・半濁点）の 往復は ★実測の 裏返し★★
       ＝`=DBCS(U+FF9E)` → U+309B は 測った／`=ASC(U+309B)` は ★測って いません★ */
T('★半角カナ 63字が 1字も 落ちない★（★表は 機械の 物★）', () => {
  let 数えた = 0, 落ち = [];
  for (let c = 0xFF61; c <= 0xFF9F; c++) {
    const 半 = String.fromCodePoint(c);
    const 全 = 引('DBCS', 半);
    数えた++;
    if (全 === 半) { 落ち.push(c); continue; }       /* ★何も 変わって いない＝落ちた★ */
    /* ★濁点（U+FF9E）と 半濁点（U+FF9F）は ★往復の 片方が 未測定★
       ⇒★往復する 事だけを 見ます（中身を 決めつけません）★ */
    const 戻 = 引('ASC', 全);
    if (戻 !== 半) 落ち.push(c);                      /* ★往復で 戻らない★ */
  }
  if (数えた !== 63) throw new Error('★63字 見る はずが ' + 数えた + '字★（U+FF61〜U+FF9F）');
  if (落ち.length) {
    throw new Error('★' + 落ち.length + '字 落ちました★ … ' + 落ち.slice(0, 6).join(' '));
  }
  console.log('      … ★63 / 63★（★往復で 戻る★）');
});

/* ═══ ③ ★濁点が 分かれる★ ═══ */
T('★ガ は 2文字に なる（ｶ ＋ ﾞ）★', () => {
  const r = 引('ASC', String.fromCodePoint(12460));
  if (r.length !== 2) throw new Error('2文字の はず（出た ' + r.length + '）');
  if (r.codePointAt(0) !== 65398) throw new Error('1文字目は 65398（出た ' + r.codePointAt(0) + '）');
  /* ★2文字目は 紙に 在りません★＝★測って いません★
     ⇒★「何か 1文字 在る」だけを 見ます★（★中身を 決めつけません★） */
  if (!r[1]) throw new Error('2文字目が 在りません');
});

/* ═══ ④ ★実測が 教えた 1か所★ ═══ */
T('★DBCS(半角濁点) は U+309B（単独の ゛）★（★NFKC の U+3099 では ない★）', () => {
  const r = 引('DBCS', String.fromCodePoint(65438));
  if (r.codePointAt(0) !== 12443) {
    throw new Error('★12443 の はず★（NFKC なら 12441／出た ' + r.codePointAt(0) + '）');
  }
});

/* ═══ ⑤ ★隣を 壊して いない★ ═══ */
T('★隣（UNICODE・UNICHAR・LEN）が 生きて いる★', () => {
  const u = F.呼ぶ('UNICODE', [字('A')], {});
  if (!u || u.値 !== 65) throw new Error('UNICODE("A") ＝ 65');
  const c = F.呼ぶ('UNICHAR', [{ 種: '直', 値: K.数(65) }], {});
  if (!c || c.値 !== 'A') throw new Error('UNICHAR(65) ＝ A');
  const l = F.呼ぶ('LEN', [字('ABC')], {});
  if (!l || l.値 !== 3) throw new Error('LEN("ABC") ＝ 3');
});

/* ═══ ⑥ ★台が 知る 数★ ═══ */
T('★台が 知る 数★（★2個 増えました★）', () => {
  const n = Object.keys(F.表).length;
  console.log('      … ★' + n + '個★（2026-09-18 の 直しで 398 → 400）');
  if (n < 400) throw new Error('★台が 知る 数が ' + n + 'に 減りました★（400以上の はず）');
  if (!F.表.ASC || !F.表.DBCS) throw new Error('★ASC／DBCS が 表から 消えて います★');
});

console.log('\n' + pass + ' passed, ' + fail + ' failed');
if (fail) process.exit(1);
