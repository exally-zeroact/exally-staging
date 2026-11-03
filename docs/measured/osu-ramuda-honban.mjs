/* osu-ramuda-honban.mjs -- ★ラムダの 一族を ★お客さんの 道★で 押す★（2026-09-18）
 *
 *  ★訳★
 *    9枠目（実Excel）で ラムダの 一族 24本の 正しい 答えが 取れました。
 *    ★でも「正が 在る」は「うちが 合って いる」では ありません★。
 *    ⇒★お客さんが 通る 道で 押して 並べます★。
 *
 *  ★道（★台の 名簿では なく 道★）★
 *    convertFormula → _jsComputeFormula → 借り物
 *    ＝`honban-no-michi.mjs` の `建てる()`（book.html と 同じ 8本を つなぐ）
 *
 *  ★★★この 台の 数は 画面の 数では ない★★★（2026-09-18・`hakaridai-mon` の 門）
 *    ＝★この 道具は `hf.setSheetContent` で ★板ごと★ 材料を 入れます★
 *      （★お客さんの 画面は 1マスずつ 入れます★）
 *    ⇒★★同じ 式でも 答えが 変わる 事が 在ります★★
 *      2026-09-10 ... 板ごと 入れたら 裸の `=LINEST(...)` が `#VALUE!` に なり
 *      ★「本番が 壊れて いる」と 報告する 一歩 手前★まで 行きました
 *      （実配信を ブラウザで 押したら 4本とも 動いて いた＝★台の 産物★）
 *    ⇒★★画面の 事を 言いたいなら ブラウザで 押して ください★★
 *      ＝`docs/measured/osu-kami-webkit.mjs`（★同じ 24本を 画面で 押せます★）
 *
 *  ★見て いない 事★
 *    ・★書き出し（lib/xlsx-io.js）と 読み込みの 道は 見て いません★
 *    ・★溢れ（スピル）の 2つ目 以降は 見て いません★＝★左上だけ★
 *      （実Excel 側も 左上を 取って います＝同じ 見方）
 *
 *  使い方: node docs/measured/osu-ramuda-honban.mjs
 */
import path from 'node:path';
import fs from 'node:fs';
import { fileURLToPath, pathToFileURL } from 'node:url';

const ここ = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(ここ, '..', '..');
const { 建てる, 台に聞く } = await import(pathToFileURL(path.join(ここ, 'honban-no-michi.mjs')).href);
const 台 = await 建てる();
const { EF, hf, SID, 板 } = 台;   /* ★板＝自前の 計算の 台（★本番の 3段目★） */

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

/* ★9枠目の 紙と ★同じ 材料★（紙の 頭に 書いて あります）★ */
const 材料 = {
  A1: 1, A2: 2, A3: 3, A4: 4, A5: 5,
  C1: 1, C2: 3, C3: 5, C4: 7, C5: 9,
  D1: 9, D2: 7, D3: 5, D4: 3, D5: 1,
  F1: 1, G1: 2, F2: 10, G2: 20,
};
/* ★借り物は 9996-9999行目を 下書きに 使います＝板を そこまで 用意します★ */
const 行数 = 10001, 列数 = 12, 式の行 = 0, 式の列 = 9;

function 押す(式) {
  let 後;
  try { 後 = EF.convertFormula(式); } catch (e) { return '★変換で 投げた★'; }
  const 表 = [];
  for (let r = 0; r < 行数; r++) 表.push(new Array(列数).fill(null));
  for (const k of Object.keys(材料)) {
    const a = 番地(k);
    表[a.行][a.列] = 材料[k];
  }
  表[式の行][式の列] = 後;
  hf.setSheetContent(SID, 表);
  try {
    const js = EF._jsComputeFormula(0, 式);
    if (js !== null && js !== undefined) return String(js);
  } catch (e) { /* ★台へ 落とす★ */ }
  /* ★★②台に 聞く★★（2026-09-18・★本番は 3段です★）
       ＝前は ここが 抜けて いて ★JS層 → 借り物★の 2段でした
       ＝★「お客さんの 道」と 名乗りながら 本番と 1段 違って いました★
       ⇒★材料を 台にも 入れてから 聞きます★（★空の 板に 聞くと 0 が 返ります★） */
  if (板) {
    try {
      for (const k of Object.keys(材料)) 板.打つ(k, String(材料[k]));
      const 台答 = 台に聞く(板, 式);
      if (台答 !== null) return 台答;
    } catch (e) { /* ★借り物へ 落とす★ */ }
  }
  const v = hf.getCellValue({ sheet: SID, row: 式の行, col: 式の列 });
  if (v && v.type) return 赤の名(v.type);
  if (v === null || v === undefined) return '';
  return String(v);
}

const 紙の名 = 'golden-kansuu-9kaime-2026-09-18.tsv';
const 紙 = fs.readFileSync(path.join(ここ, 紙の名), 'utf-8').replace(/^﻿/, '');
const 行 = 紙.split(/\r?\n/).filter((l) => l && !l.startsWith('#')).map((l) => l.split('\t'));
const ラムダ = 行.filter((c) => /LAMBDA|LET|MAP|REDUCE|SCAN|BYROW|BYCOL|MAKEARRAY/i.test(c[2] || ''));

/* ★飾りを 外してから 比べます★（★機械が 読む 欄に 飾りを 書かない★の 裏） */
const 裸 = (s) => String(s).replace(/★/g, '').replace(/`/g, '').trim();

let 合 = 0, 違 = 0;
const 出し = ['番\t式\t実Excel\tうち\t判定'];
for (const c of ラムダ) {
  const 式 = String(c[2]).trim();
  const 実 = 裸(c[5]);
  let 出;
  try { 出 = 押す(式); } catch (e) { 出 = '例外:' + e.message; }
  const 数どうし = /^-?[\d.]+$/.test(実) && /^-?[\d.]+$/.test(出);
  const 同 = (出 === 実) || (数どうし && Math.abs(Number(出) - Number(実)) < 1e-9);
  if (同) { 合 += 1; } else { 違 += 1; }
  出し.push(c[0] + '\t' + 式 + '\t' + 実 + '\t' + 出 + '\t' + (同 ? 'o' : 'x'));
}
出し.push('');
出し.push('★ラムダの一族 ' + (合 + 違) + '本 ... 合った ' + 合 + ' ／ 違った ' + 違 + '★');
console.log(出し.join('\n'));
