/* osu-sashikomi.mjs — ★取った 紙と 見張りの 中の 17通りを 突き合わせる★（2026-09-13）
 *
 *  ★★何の 為か★★
 *    `lib/shiki-sashikomi.js`（行/列の 入れ消しで 式の 参照が 追従する）の 答えは
 *    全部 実Excel の 実測です。
 *    ★実Excel の 版が 変わった／測り方が 崩れた 時に ここが 赤に なります★
 *
 *  ★使い方★
 *    pwsh -NoProfile -File docs/measured/toru-sashikomi.ps1   … ★実Excel に 打って 紙を 作る★
 *    node docs/measured/osu-sashikomi.mjs                     … ★突き合わせる★
 *
 *  ★★この 台は 見張りでは ありません★★
 *    `tests/run.js` から は 回しません（★Excel が 要る＝この パソコンでしか 動かない★）。
 *    見張り（tests/shiki-sashikomi.test.mjs）は ★実Excel 無しで★ 17通りを 守ります。
 *
 *  ★★列の 名前で 読みます（番号で 読まない）★★
 *    2026-09-13 に 紙の 列が ★6列→8列★ に 増えました（型・本当にゼロか を 足した）。
 *    ★番号で 読んで いたら 黙って ずれて いました★。
 *    ⇒★見出しの 行（`# 置いた式 …`）から 列の 名前を 読んで 位置を 決めます★
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..', '..');
const 紙 = path.join(ROOT, 'docs/measured/golden-sashikomi-2026-09-13.tsv');
const 見張り = path.join(ROOT, 'tests/shiki-sashikomi.test.mjs');

if (!fs.existsSync(紙)) {
  console.error('★先に これを 走らせて ください★');
  console.error('  pwsh -NoProfile -File docs/measured/toru-sashikomi.ps1');
  process.exit(2);
}

/* ★見張りの 中の 〔置いた式, 向き, どこ, 消すか, 後の式〕を 読む★
   （★2か所に 同じ 表を 書かない★） */
const 中 = fs.readFileSync(見張り, 'utf-8');
const 期待 = new Map();
for (const m of 中.matchAll(/\['(=[^']*)',\s*'([行列])',\s*(\d+),\s*(true|false),\s*'((?:[^'\\]|\\.)*)'/g)) {
  const 鍵 = m[1] + '|' + (m[2] === '行' ? 'Rows' : 'Columns') + '(' + m[3] + ').' + (m[4] === 'true' ? 'Delete' : 'Insert');
  期待.set(鍵, m[5].replace(/\\'/g, "'").replace(/\\\\/g, '\\'));
}
if (期待.size < 10) {
  console.error('★見張りから ' + 期待.size + '通り しか 読めない★（17通り のはず）');
  process.exit(2);
}

/* ★取った 紙を 読む★（★列の 名前で 位置を 決める★） */
const 行たち = fs.readFileSync(紙, 'utf-8').split(/\r?\n/);
const 見出し = 行たち.find((l) => l.startsWith('# 置いた式'));
if (!見出し) { console.error('★紙に 見出しの 行が 無い★'); process.exit(2); }
const 列名 = 見出し.replace(/^#\s*/, '').split('\t');
const 位置 = (名) => {
  const i = 列名.indexOf(名);
  if (i < 0) { console.error('★紙に 列「' + 名 + '」が 無い★（' + 列名.join(' / ') + '）'); process.exit(2); }
  return i;
};
const c式 = 位置('置いた式'), cわざ = 位置('わざ'), c後 = 位置('後の式');

const 実 = new Map();
for (const l of 行たち) {
  if (!l || l.startsWith('#')) continue;
  const c = l.split('\t');
  if (c.length <= c後) continue;
  実.set(c[c式] + '|' + c[cわざ], c[c後]);
}

let 合 = 0;
const 違い = [], 無い = [];
for (const [鍵, 欲] of 期待) {
  if (!実.has(鍵)) { 無い.push(鍵); continue; }
  const 出 = 実.get(鍵);
  if (出 === 欲) 合++;
  else 違い.push(鍵 + ' … 紙 `' + 出 + '` ／ ★見張り `' + 欲 + '`★');
}

console.log('');
console.log('★★締め★★ 見張り ' + 期待.size + '通り ／ 合った ' + 合 + ' ／ ★違う ' + 違い.length + '★'
  + (無い.length ? ' ／ ★紙に 無い ' + 無い.length + '★' : ''));
console.log('　紙の 列 … ' + 列名.join(' / '));
if (違い.length) {
  console.log('');
  console.log('★合わない 物の 実物★');
  違い.slice(0, 20).forEach((x) => console.log('  ' + x));
  console.log('');
  console.log('★どちらが 正しいかは 人が 決めます★');
  console.log('  ・実Excel の 版が 変わった のかも しれません');
  console.log('  ・★測り方（材料の 置き場・式を 置く マス）が 崩れた のかも しれません★');
}
if (無い.length) {
  console.log('');
  console.log('★紙に 無い 物★（取り直して ください）');
  無い.slice(0, 10).forEach((x) => console.log('  ' + x));
}
process.exit((違い.length || 無い.length) ? 1 : 0);
