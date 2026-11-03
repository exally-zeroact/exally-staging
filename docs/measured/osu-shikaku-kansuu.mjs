/* osu-shikaku-kansuu.mjs — ★取った 紙と 見張りの 中の 37通りを 突き合わせる★（2026-09-13）
 *
 *  ★★何の 為か★★
 *    `lib/shiki-kansuu.js`（四角と 関数）の 答えは 全部 実Excel の 実測です。
 *    ★実Excel の 版が 変わった／測り方が 崩れた 時に ここが 赤に なります★
 *
 *  ★使い方★
 *    pwsh -NoProfile -File docs/measured/toru-shikaku-kansuu.ps1   … ★実Excel に 打って 紙を 作る★
 *    node docs/measured/osu-shikaku-kansuu.mjs                     … ★突き合わせる★
 *
 *  ★★この 台は 見張りでは ありません★★
 *    `tests/run.js` から は 回しません（★Excel が 要る＝この パソコンでしか 動かない★）。
 *    見張り（tests/shiki-kansuu.test.mjs）は ★実Excel 無しで★ 37通りを 守ります。
 *
 *  ★★紙は 40通り・見張りは 37通り＝この 差は わざと★★
 *    `=SUMPRODUCT(A1:A5,A1:A5)` … 形の 違う 四角の 決まりが ★未測定★
 *    `=A1:A3` `=A1:A3*2`        … ★溢れ＝土台⑤★（まだ 作って いない）
 *    ⇒ ★出来て いない 物を 出来た 顔で 混ぜない★ 為に わざと 外して います。
 *      ★この 3つが 紙から 消えたら 赤に します★（わざと 外した のか 落ちたのか を 見分ける 為）
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..', '..');
const 紙 = path.join(ROOT, 'docs/measured/golden-shikaku-kansuu-2026-09-13.tsv');
const 見張り = path.join(ROOT, 'tests/shiki-kansuu.test.mjs');

/* ★わざと 見張りに 入れて いない 式★（★紙には 在る はず★） */
const わざと外した = ['=SUMPRODUCT(A1:A5,A1:A5)', '=A1:A3', '=A1:A3*2'];

if (!fs.existsSync(紙)) {
  console.error('★先に これを 走らせて ください★');
  console.error('  pwsh -NoProfile -File docs/measured/toru-shikaku-kansuu.ps1');
  process.exit(2);
}

/* ★見張りの 中の 〔式, 答え〕を 読む★（★2か所に 同じ 表を 書かない★） */
const 中 = fs.readFileSync(見張り, 'utf-8');
const 期待 = new Map();
for (const m
  of 中.matchAll(/\['(=[^']*)',\s*'((?:[^'\\]|\\.)*)'\]/g)) {
  期待.set(m[1], m[2].replace(/\\'/g, "'").replace(/\\\\/g, '\\'));
}
if (期待.size < 30) {
  console.error('★見張りから ' + 期待.size + '通り しか 読めない★（37通り のはず）');
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

/* ★わざと 外した 物が 紙から 消えて いないか★
   ＝「わざと 外した」のか「測り忘れた」のかを ★見分けられる ように して おく★ */
const 外した物が紙に無い = わざと外した.filter((式) => !実.has(式));

console.log('');
console.log('★★締め★★ 見張り ' + 期待.size + '通り ／ 合った ' + 合 + ' ／ ★違う ' + 違い.length + '★'
  + (無い.length ? ' ／ ★紙に 無い ' + 無い.length + '★' : ''));
console.log('　紙は ' + 実.size + '通り ／ ★わざと 外した ' + わざと外した.length + '通り★'
  + '（' + わざと外した.join(' ') + '）');
if (違い.length) {
  console.log('');
  console.log('★合わない 物の 実物★');
  違い.slice(0, 20).forEach((x) => console.log('  ' + x));
  console.log('');
  console.log('★どちらが 正しいかは 人が 決めます★');
  console.log('  ・実Excel の 版が 変わった のかも しれません');
  console.log('  ・★測り方（材料の 置き場・空マスの 場所）が 崩れた のかも しれません★');
}
if (無い.length) {
  console.log('');
  console.log('★紙に 無い 式★（取り直して ください）');
  無い.slice(0, 10).forEach((x) => console.log('  ' + x));
}
if (外した物が紙に無い.length) {
  console.log('');
  console.log('★★わざと 外した 式が 紙にも 無い★★（測り忘れと 見分けが 付かなく なります）');
  外した物が紙に無い.forEach((x) => console.log('  ' + x));
}
process.exit((違い.length || 無い.length || 外した物が紙に無い.length) ? 1 : 0);
