/* ramuda-honban.test.mjs -- ★ラムダの 一族が ★お客さんの 道★で 悪く なって いないか★（2026-09-18）
 *
 *  ★★紙★★ `docs/measured/golden-ramuda-honban-2026-09-18.tsv`
 *    ＝9枠目の 実Excel（24本）を ★お客さんの 道★で 押した 突き合わせ
 *    ＝★2026-09-18 の 実測 ... 合った 8本 ／ 違った 16本★
 *
 *  ★★この 試験が 守る 物★★
 *    ⑴★合って いる 8本が 崩れない★（★名指しで 持ちます★）
 *       ＝「直したら 別の 所が 壊れた」を その場で 止める
 *    ⑵★違って いる 16本が 増えない★（★上限★）
 *    ⑶★★減ったら 赤に します★★＝★紙と 上限を 下げて ください★
 *       ＝★緩い 上限は 門では ありません★（2026-09-18 の 決め）
 *
 *  ★★見て いない 事（★書かない 見張りは「全部 守った」と 読まれる★）★★
 *    ・★材料は 1組だけ★です（A1:A5 / C / D / F1:G2）
 *      ＝★別の 材料で 分かれるかは 見て いません★
 *    ・★溢れの 2つ目 以降は 見て いません★＝★左上だけ★
 *    ・★書き出し・読み込みの 道は 見て いません★
 *
 *  使い方: node tests/ramuda-honban.test.mjs
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const 紙の道 = path.join(ROOT, 'docs/measured/golden-ramuda-honban-2026-09-18.tsv');

let pass = 0, fail = 0;
const T = (n, f) => {
  try { f(); pass += 1; console.log('  ok   ' + n); }
  catch (e) { fail += 1; console.log('  NG   ' + n + '\n       ' + e.message); }
};

console.log('\n[ramuda-honban] ★ラムダの 一族 ... お客さんの 道★');

/* ═══ ★紙を 読む★ ═══ */
const 生 = fs.readFileSync(紙の道, 'utf-8').replace(/^﻿/, '');
const 行 = 生.split(/\r?\n/)
  .filter((l) => l && !l.startsWith('#') && l.includes('\t'))
  .map((l) => l.split('\t'))
  .filter((c) => String(c[1] || '').startsWith('='));

/* ★2026-09-18 の 実測★（★減ったら 下げる★） */
const 紙の本数 = 24;
/* ★★2026-09-18 ... ★取り直しました★★★
     ★訳★ … 道具が ★本番の 道と 1段 違って いました★
       前 … convertFormula -> _jsComputeFormula -> 借り物（★2段★）
       今 … convertFormula -> _jsComputeFormula -> ★台★ -> 借り物（★3段★）
     ⇒★★合った 8 → 21 ／ 違った 16 → 3★★
     ⇒★上限を 下げました★（★緩い 上限は 門では ありません★） */
/* ══ ★★2026-09-19 ... ★また 同じ 道具が 本番と 1段 違って いました★★★ ══
     ★前回（09-18）★ ... 道具が ★台の 段★を 飛ばして いた ⇒ 合った 8 -> 21
     ★今回（09-19）★ ... `honban-no-michi.mjs` の `台に聞く()` が
       ★溢れを null で 返して いた★（＝★昔の `book.html` の 写し★）
       ⇒`=MAP(A1:A5,LAMBDA(x,x*2))` は ★台が 2 を 出して いる★のに
         null ⇒ 借り物へ ⇒ ★借り物は MAP を 知らない ⇒ `#NAME?`★
       ⇒★★「台が 悪い」のでは なく 測り道具が 本番と 違った★★
     ★★同じ 道具で 2回 同じ 型を 踏みました★★
       ＝★本番の `recalcSheet` が 変わった のに 道具が 付いて いかない★
       ⇒記憶「★台が 違うと 言う 前に 測り道具が その式を 押せるかを 見る★」
     ★直した 後★ ... 合った ★21 -> 23★ ／ 違った ★3 -> 1★
       ＝BYROW(F1:G2) が `#NAME?` -> ★3★ ／ BYCOL(F1:G2) が `#NAME?` -> ★11★
       ＝★どちらも 実Excel と 同じ★
     ★上限を 下げました★（★緩い 上限は 門では ありません★） */
const 紙の合った = 23;
const 紙の違った = 1;

T('★紙が ' + 紙の本数 + '本 在る★（★分母★）', () => {
  if (行.length !== 紙の本数) {
    throw new Error('★紙に ' + 紙の本数 + '本 在る はずが ' + 行.length + '本★');
  }
});

/* ═══ ★道具を その場で 走らせて 紙と 比べる★ ═══ */
const { execFileSync } = await import('node:child_process');
let 出し;
T('★道具が 走る★', () => {
  出し = execFileSync(process.execPath,
    [path.join(ROOT, 'docs/measured/osu-ramuda-honban.mjs')],
    { encoding: 'utf-8', stdio: ['ignore', 'pipe', 'ignore'] });
  if (!出し || !出し.includes('ラムダの一族')) throw new Error('★出しが 空です★');
});

const 今 = new Map();
for (const l of String(出し || '').split(/\r?\n/)) {
  const c = l.split('\t');
  if (c.length >= 5 && String(c[1] || '').startsWith('=')) 今.set(c[1].trim(), c[3]);
}

T('★道具が ' + 紙の本数 + '本 出す★（★分母★）', () => {
  if (今.size !== 紙の本数) throw new Error('★' + 紙の本数 + '本 の はずが ' + 今.size + '本★');
});

/* ═══ ★① 紙の 1行 1行と 突き合わせる★ ═══ */
T('★★紙の ' + 紙の本数 + '本と 1行ずつ 合う★★（★紙から 読みます★）', () => {
  let 押した = 0;
  const ずれ = [];
  for (const c of 行) {
    const 式 = String(c[1]).trim();
    const 紙のうち = String(c[3]).trim();
    if (!今.has(式)) throw new Error('★道具が この 式を 出しません★ ' + 式);
    const 今のうち = String(今.get(式)).trim();
    if (今のうち !== 紙のうち) ずれ.push(式 + ' ... 紙 [' + 紙のうち + '] ／ 今 [' + 今のうち + ']');
    押した += 1;
  }
  if (押した !== 紙の本数) throw new Error('★' + 紙の本数 + '本 押す はずが ' + 押した + '本★');
  if (ずれ.length) {
    throw new Error('★台の 答えが 変わりました（' + ずれ.length + '本）★\n       '
      + ずれ.join('\n       ')
      + '\n       ★良く なったなら 紙を 取り直して 上限も 下げて ください★');
  }
  console.log('      ... ★' + 押した + ' / ' + 紙の本数 + '★（紙から 読んだ 分母）');
});

/* ═══ ★② 合って いる 本数が 減らない・違って いる 本数が 増えない★ ═══ */
T('★★合った ' + 紙の合った + '本 ／ 違った ' + 紙の違った + '本★★（★動いたら 赤★）', () => {
  let 合 = 0, 違 = 0;
  for (const c of 行) {
    if (String(c[4]).trim() === 'o') 合 += 1; else 違 += 1;
  }
  if (合 !== 紙の合った || 違 !== 紙の違った) {
    throw new Error('★紙が 合った ' + 合 + ' ／ 違った ' + 違 + ' に なって います★'
      + '\n       ★紙を 取り直したなら この 試験の 数も 直して ください★');
  }
  if (合 + 違 !== 紙の本数) throw new Error('★足しても ' + 紙の本数 + 'に なりません★');
});

/* ═══ ★③ 合って いる 8本を ★名指しで★ 持つ★ ═══ */
const 合っている = 行.filter((c) => String(c[4]).trim() === 'o').map((c) => String(c[1]).trim());
T('★★合って いる ' + 紙の合った + '本が 名指しで 崩れない★★', () => {
  const 崩れ = [];
  for (const 式 of 合っている) {
    const 紙の実 = String((行.find((c) => String(c[1]).trim() === 式) || [])[2]).trim();
    const 今のうち = String(今.get(式)).trim();
    const 数どうし = /^-?[\d.]+$/.test(紙の実) && /^-?[\d.]+$/.test(今のうち);
    const 同 = (今のうち === 紙の実)
      || (数どうし && Math.abs(Number(今のうち) - Number(紙の実)) < 1e-9);
    if (!同) 崩れ.push(式 + ' ... 実Excel [' + 紙の実 + '] ／ 今 [' + 今のうち + ']');
  }
  if (崩れ.length) {
    throw new Error('★合って いた ものが 崩れました（' + 崩れ.length + '本）★\n       ' + 崩れ.join('\n       '));
  }
  console.log('      ... ★' + 合っている.length + ' / ' + 紙の合った + '★');
});

console.log('\nramuda-honban: ' + pass + ' 緑 / ' + fail + ' 赤\n');
process.exit(fail ? 1 : 0);
