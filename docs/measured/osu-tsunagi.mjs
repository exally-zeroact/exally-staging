/* osu-tsunagi.mjs — ★取った 紙と 見張りの 中の 162通りを 突き合わせる★（2026-09-11）
 *
 *  ★★何の 為か★★
 *    `lib/shiki-keisan.js` の 答えは 全部 実Excel の 実測です。
 *    ★実Excel の 版が 変わった／測り方が 崩れた 時に ★ここが 赤に なります★
 *
 *  ★使い方★
 *    pwsh -NoProfile -File docs/measured/toru-tsunagi.ps1   … ★実Excel に 打って 紙を 作る★
 *    node docs/measured/osu-tsunagi.mjs                     … ★突き合わせる★
 *
 *  ★★この 台は 見張りでは ありません★★
 *    `tests/run.js` から は 回しません（★Excel が 要る＝この パソコンでしか 動かない★）。
 *    見張り（tests/shiki-keisan.test.mjs）は ★実Excel 無しで★ 162通りを 守ります。
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..', '..');
const 紙 = path.join(ROOT, 'docs/measured/golden-tsunagi-2026-09-11.tsv');
const 見張り = path.join(ROOT, 'tests/shiki-keisan.test.mjs');

if (!fs.existsSync(紙)) {
  console.error('★先に これを 走らせて ください★');
  console.error('  pwsh -NoProfile -File docs/measured/toru-tsunagi.ps1');
  process.exit(2);
}

/* ★見張りの 中の 〔式, 答え〕を 読む★（★2か所に 同じ 表を 書かない★） */
const 中 = fs.readFileSync(見張り, 'utf-8');
const 期待 = new Map();
for (const m
  of 中.matchAll(/\['(=[^']*)',\s*'((?:[^'\\]|\\.)*)'\]/g)) {
  期待.set(m[1], m[2].replace(/\\'/g, "'").replace(/\\\\/g, '\\'));
}
if (期待.size < 100) {
  console.error('★見張りから ' + 期待.size + '通り しか 読めない★（162通り のはず）');
  process.exit(2);
}

/* ★取った 紙を 読む★ */
const 実 = new Map();
for (const l of fs.readFileSync(紙, 'utf-8').split(/\r?\n/)) {
  if (!l || l.startsWith('#')) continue;
  const c = l.split('\t');
  if (c.length < 2) continue;
  実.set(c[0], c[1]);
}

let 合 = 0;
const 違い = [];
const 無い = [];
for (const [式, 欲] of 期待) {
  if (!実.has(式)) { 無い.push(式); continue; }
  const 出 = 実.get(式);
  if (出 === 欲) 合++;
  else 違い.push(式 + ' … 紙 `' + 出 + '` ／ ★見張り `' + 欲 + '`★');
}

console.log('');
console.log('★★締め★★ 見張り ' + 期待.size + '通り ／ 合った ' + 合 + ' ／ ★違う ' + 違い.length + '★'
  + (無い.length ? ' ／ ★紙に 無い ' + 無い.length + '★' : ''));
if (違い.length) {
  console.log('');
  console.log('★合わない 物の 実物★');
  違い.slice(0, 20).forEach((x) => console.log('  ' + x));
  console.log('');
  console.log('★どちらが 正しいかは 人が 決めます★');
  console.log('  ・実Excel の 版が 変わった のかも しれません');
  console.log('  ・★測り方（列の 幅・空マスの 場所）が 崩れた のかも しれません★');
}
if (無い.length) {
  console.log('');
  console.log('★紙に 無い 式★（取り直して ください）');
  無い.slice(0, 10).forEach((x) => console.log('  ' + x));
}
process.exit((違い.length || 無い.length) ? 1 : 0);
