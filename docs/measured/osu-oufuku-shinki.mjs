/* osu-oufuku-shinki.mjs — ★今回 足した 関数が Exally→Excel で 生き残るか★（2026-09-16）
 *
 *  ★★司さん（2026-09-16）★★
 *    「ルールに Exally と Excel に 引き渡しても ちゃんと 動くか
 *      ★実際に 動作確認しながら★ やれよ」
 *
 *  ★★なぜ 要るか★★
 *    紙（golden-*.tsv）が 全部 緑でも それは ★台の 中だけ★の 話。
 *    お客さんは ★Exally で 作って Excel で 開く★。
 *    ⇒★式が 落ちる／値が 消える★ かもしれない＝★紙は 1本も 赤に ならない★
 *
 *  ★★測る 事（3つとも 見る）★★
 *    ①★式★  … Excel が 持って いる 式（`.Formula`）＝★落ちて いないか★
 *    ②★出る字★ … 画面に 出る 字（`.Text`）
 *    ③★答え★ … 中の 数（`.Value2`）＝★うちの 答えと 合うか★
 *    ⇒★答えだけ 見ない★（★式が 落ちて 値だけ 残る★ 事が 在る）
 *
 *  ★★押す 関数★★ … 棚58〜67 で 足した 物（★手で 並べません★＝下の 一覧が 正本）
 *
 *  使い方:
 *    ① node docs/measured/osu-oufuku-shinki.mjs          … ★うちで 書き出す★
 *    ② pwsh -NoProfile -File docs/measured/toru-oufuku-shinki.ps1  … ★実Excel で 読む★
 *    ③ node docs/measured/osu-oufuku-shinki.mjs --awase  … ★突き合わせる★
 */
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { ROOT } from './honban-no-michi.mjs';

const require_ = createRequire(path.join(ROOT, 'package.json'));
const ファイル = path.join(ROOT, 'docs/measured/oufuku-shinki.xlsx');
const 紙 = path.join(ROOT, 'docs/measured/golden-oufuku-shinki-2026-09-16.tsv');

/* ★★押す 式★★ … ★棚58〜67 で 足した 関数だけ★
     ★材料を 使う 物は A1:A5（1〜5）／B1:B5（2,4,6,8,10）／D1・D2（日付）★ */
const 式たち = [
  /* 棚58 統計 */
  ['C1', '=NORM.DIST(5,3,2,TRUE)'],
  ['C2', '=BINOM.DIST(6,10,0.5,FALSE)'],
  ['C3', '=CORREL(A1:A5,B1:B5)'],
  ['C4', '=SLOPE(B1:B5,A1:A5)'],
  ['C5', '=RANK.EQ(3,A1:A5)'],
  ['C6', '=CHISQ.DIST(2,3,TRUE)'],
  ['C7', '=CONFIDENCE.NORM(0.05,2.5,50)'],
  /* 棚59 日付の字 */
  ['C8', '=DATE(1900,2,28)'],
  ['C9', '=DATEVALUE("2024-01-01")'],
  ['C10', '=TIMEVALUE("12:30")'],
  /* 棚61 お金 */
  ['C11', '=TBILLPRICE(2,3,4)'],
  ['C12', '=SERIESSUM(2,3,4,5)'],
  ['C13', '=DB(0.5,1,2,3,4)'],
  ['C14', '=CUMIPMT(0.05,12,100,1,12,0)'],
  /* 棚62 番地 */
  ['C15', '=ROW(A2)'],
  ['C16', '=COLUMN(B1:B5)'],
  ['C17', '=ISOMITTED(2)'],
  /* 棚65 束ね（★一番 危ない＝式の 形が 特殊★） */
  ['C18', '=AREAS((A1,A2))'],
  ['C19', '=AREAS((A1:A2,B1:B3,A4))'],
  ['C20', '=SUM(AREAS((A1,A2)),1)'],
  /* 棚67 よく 使う */
  ['C21', '=SUMPRODUCT(A1:A5,B1:B5)'],
  ['C22', '=CONCATENATE("a","b")'],
  ['C23', '=FALSE()'],
  /* ★★YEARFRAC basis 1★★ … 2026-09-16 に ★客に 出て いた 欠陥★を 直した 所
       ★分かれ道を 3つとも 往復させます★
         C24 … ②年を またぐが 暦で 1年 以内（★直した 形★）
         C25 … ③同じ 366日でも 平均に なる 形（★直しで 壊して いないか★）
         C26 … ①同じ 年（★前から 合って いた 形★） */
  ['C24', '=YEARFRAC(DATE(2008,10,15),DATE(2009,3,1),1)'],
  ['C25', '=YEARFRAC(DATE(2008,3,1),DATE(2009,3,2),1)'],
  ['C26', '=YEARFRAC(DATE(2008,1,15),DATE(2008,6,30),1)'],
];

const マスを分ける = (a) => {
  const m = /^([A-Z]+)(\d+)$/.exec(a);
  const c = m[1].split('').reduce((s, ch) => s * 26 + (ch.charCodeAt(0) - 64), 0) - 1;
  return { r: Number(m[2]) - 1, c: c };
};

if (process.argv.indexOf('--awase') < 0) {
  /* ══ ★うちで 書き出す★ ══ */
  global.XLSX = require_(path.join(ROOT, 'lib/xlsx.full.min.js'));
  const GX = require_(path.join(ROOT, 'lib/grid-xlsx.js'));
  const IO = require_(path.join(ROOT, 'lib/xlsx-io.js'));
  const H = require_(path.join(ROOT, 'lib/shiki-hyou.js'));

  /* ★うちの 答えも その場で 出す★（★本番の 台で★） */
  const h = H.表();
  [1, 2, 3, 4, 5].forEach((v, i) => h.打つ('A' + (i + 1), String(v)));
  [2, 4, 6, 8, 10].forEach((v, i) => h.打つ('B' + (i + 1), String(v)));
  h.打つ('D1', '45292');
  h.打つ('D2', '46023');
  const うちの答え = [];
  for (const [マス, 式] of 式たち) {
    h.打つ(マス, 式);
    うちの答え.push([マス, 式, h.字(マス)]);
  }

  const 板 = { data: {}, name: 'Sheet1' };
  const 置く = (a, o) => { const p = マスを分ける(a); 板.data[p.r + ',' + p.c] = o; };
  [1, 2, 3, 4, 5].forEach((v, i) => 置く('A' + (i + 1), { v: v }));
  [2, 4, 6, 8, 10].forEach((v, i) => 置く('B' + (i + 1), { v: v }));
  置く('D1', { v: 45292 });
  置く('D2', { v: 46023 });
  for (const [マス, 式] of 式たち) 置く(マス, { v: 式, f: 式, d: '' });

  const buf = IO.writeBook(GX.gridToBook([板]));
  fs.writeFileSync(ファイル, Buffer.from(buf));

  /* ★うちの 答えを 控える★（突き合わせの 時に 使う） */
  fs.writeFileSync(path.join(ROOT, 'docs/measured/oufuku-shinki-uchi.tsv'),
    うちの答え.map((r) => r.join('\t')).join('\n') + '\n', 'utf-8');

  console.log('');
  console.log('★うちで 書き出した★ … ' + ファイル);
  console.log('  ' + Buffer.from(buf).length + ' バイト ／ 式 ' + 式たち.length + '本');
  console.log('');
  console.log('★うちの 答え★');
  for (const [マス, 式, 出] of うちの答え) {
    console.log('  ' + マス.padEnd(4) + 式.padEnd(32) + ' → ' + 出);
  }
  console.log('');
  console.log('★次に これを 走らせて ください★');
  console.log('  pwsh -NoProfile -File docs/measured/toru-oufuku-shinki.ps1');
  console.log('  node docs/measured/osu-oufuku-shinki.mjs --awase');
} else {
  /* ══ ★突き合わせる★ ══ */
  if (!fs.existsSync(紙)) {
    console.error('★先に `pwsh -NoProfile -File docs/measured/toru-oufuku-shinki.ps1` を 走らせて ください★');
    process.exit(2);
  }
  const 実 = new Map();
  for (const l of fs.readFileSync(紙, 'utf-8').replace(/^﻿/, '').split(/\r?\n/)) {
    if (!l || l.startsWith('#')) continue;
    const c = l.split('\t');
    if (c.length < 4) continue;
    実.set(c[0], { 式: c[1], 字: c[2], 答: c[3] });
  }
  const うち = new Map();
  for (const l of fs.readFileSync(path.join(ROOT, 'docs/measured/oufuku-shinki-uchi.tsv'), 'utf-8').split(/\r?\n/)) {
    if (!l) continue;
    const c = l.split('\t');
    うち.set(c[0], { 式: c[1], 答: c[2] });
  }

  let 式合 = 0, 式落 = 0, 答合 = 0, 答違 = 0;
  const 見てない = [];
  const 落ちた = [], 違う = [];
  for (const [マス, u] of うち) {
    const e = 実.get(マス);
    /* ★★実Excel の 紙に 無い マスは ★黙って 飛ばして は いけません★★★
         2026-09-16 … 式を 3本 足したのに PS1 の 名簿が 手書きだったので
         ★26本 押して 23本しか 見ず、それでも 「全部 合って います」と 出ました★
       ⇒★見て いない 分を 数えて ★最後に 赤に します★★ */
    if (!e) { 見てない.push(マス); continue; }
    /* ①★式が 落ちて いないか★ */
    const 実式 = String(e.式 || '').trim();
    if (実式 === u.式) 式合++;
    else { 式落++; 落ちた.push([マス, u.式, 実式]); }
    /* ③★答えが 合うか★（数なら 数で 比べる） */
    /* ★真偽の 字を 揃える★（実Excel の `.Value2` は `False`／台は `FALSE`）
         ＝★どちらが 正しいかの 話では なく 取り方の 話★（紙の 決まりと 同じ） */
    const 真偽を揃える = (x) => (x === 'True' ? 'TRUE' : (x === 'False' ? 'FALSE' : x));
    const uu = 真偽を揃える(String(u.答)), ee = 真偽を揃える(String(e.答));
    const a = Number(uu), b = Number(ee);
    const 同 = (uu === ee) ||
      (isFinite(a) && isFinite(b) && (a === b || Math.abs(a - b) <= Math.abs(b) * 1e-12));
    if (同) 答合++; else { 答違++; 違う.push([マス, u.答, e.答]); }
  }

  console.log('');
  console.log('★★Exally → Excel の 往復★★ … 押した 式 ' + うち.size + '本');
  console.log('');
  console.log('①★式が 生き残った★ … ' + 式合 + '本 ／ ★落ちた★ ' + 式落 + '本');
  for (const [m, u, j] of 落ちた) console.log('    ★' + m + '★ うち `' + u + '` ／ Excel `' + j + '`');
  console.log('③★答えが 合った★ … ' + 答合 + '本 ／ ★違う★ ' + 答違 + '本');
  for (const [m, u, j] of 違う) console.log('    ★' + m + '★ うち `' + u + '` ／ Excel `' + j + '`');
  console.log('');
  /* ★★分母を 出す★★＝★押した 数と 突き合わせた 数が 合わなければ 赤★ */
  console.log('★押した 式★ ' + うち.size + '本 ／ ★突き合わせた★ ' + (式合 + 式落) + '本'
    + ' ／ ★実Excel の 紙に 無かった★ ' + 見てない.length + '本');
  if (見てない.length) {
    console.log('    ★' + 見てない.join(' ') + ' … ★PS1 が 見て いません★★');
  }
  console.log('');
  if (式落 === 0 && 答違 === 0 && 見てない.length === 0) {
    console.log('★★全部 生き残って 答えも 合って います★★');
  } else {
    console.log('★★上の 分が 往復で 壊れて います（または 見て いません）★★');
    process.exit(1);
  }
}
