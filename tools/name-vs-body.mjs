/* name-vs-body.mjs — ★名前と 中身が 違う 試験★を 機械で 探す 2026-08-30
 *
 *  ★なぜ★（2026-08-30 監査役の 指摘）
 *    見出しに ★「わざと 壊して 赤に なるか」★と 書いてあるのに、
 *    中で ★何も 壊していない★ 試験が 10本 あった。
 *    ⇒ 誰かが「緑だから 壊して 赤を 見たんだな」と ★誤解する★。
 *    ★名前が 嘘の 見張りは、見張りが 無いより 悪い。★
 *
 *  ★数え方★
 *    `--self-test` の かたまり（`if (壊す) {` から 終わりまで）を 取り出し、
 *    ★本当に 壊しているか★を 見る。壊していると 数えるのは 次の どれか：
 *      ・元の 字を 書き換える（`.replace(` / `writeFileSync`）
 *      ・わざと 間違った 物を その場で 作って 通す（`にせ` `事故` `壊し` などの 変数）
 *      ・別の 壊し方の 表を 持っている（`壊し方` `breakers` `わざと`）
 *    1つも 無ければ ★名前と 中身が 違う★。
 *
 *  走らせ方: node tools/name-vs-body.mjs
 */
import fs from 'node:fs';
import path from 'node:path';

const ROOT = path.join(path.dirname(new URL(import.meta.url).pathname.replace(/^\/([A-Za-z]:)/, '$1')), '..');
const TESTS = path.join(ROOT, 'tests');

/** `if (壊す) {` の かたまりを 取り出す（無ければ null） */
function 壊すかたまり(src) {
  /* ★書き方は 1つでは ない★＝`if (壊す)` の他に
     `if (process.argv.includes('--self-test'))` の 形も 在る（08-30 実測）。
     ★取りこぼすと「壊していない」と 嘘の 報告に なる★ので 両方 見る。 */
  const 印 = ['if (壊す) {', 'if(壊す){', 'if (壊す)  {',
    "if (process.argv.includes('--self-test')) {",
    "if(process.argv.includes('--self-test')){"];
  let i = -1;
  for (const p of 印) { const k = src.indexOf(p); if (k >= 0 && (i < 0 || k < i)) i = k; }
  if (i < 0) return null;
  const j = src.indexOf('{', i);
  let d = 0;
  for (let k = j; k < src.length; k++) {
    if (src[k] === '{') d++;
    else if (src[k] === '}') { d--; if (d === 0) return src.slice(i, k + 1); }
  }
  return src.slice(i);
}

/** ★本当に 壊しているか★ */
function 壊しているか(中) {
  if (/\.replace\(/.test(中)) return '元の 字を 書き換えている';
  if (/writeFileSync/.test(中)) return 'わざと 壊した ファイルを 作っている';
  if (/にせ|事故|壊し方|breakers|わざと壊|壊す物|こわし/.test(中)) return 'わざと 間違った 物を 通している';
  /* 作り物の 入力を その場で 組んで 通す 形（純関数の 見張りに 多い） */
  if (/(判定|見る|調べる|数える)\(\s*\{|作り物|ダミー/.test(中)) return '作り物を 通している';
  return null;
}

const 名乗り = /わざと\s*壊して\s*赤に\s*なるか/;
const 出 = [];
const 良い = [];
const 壊す無し = [];

for (const f of fs.readdirSync(TESTS)) {
  if (!/\.(mjs|js)$/.test(f)) continue;
  if (f === 'run.js') continue;      /* 走らせる 側＝試験では ない */
  const p = path.join(TESTS, f);
  let src;
  try { src = fs.readFileSync(p, 'utf8'); } catch (e) { continue; }
  if (!名乗り.test(src)) continue;                 /* その 名前を 名乗っていない */
  const 中 = 壊すかたまり(src);
  if (中 === null) { 壊す無し.push(f); continue; } /* --self-test が 無い のに 名乗っている */
  const 訳 = 壊しているか(中);
  if (訳) 良い.push(f + ' … ' + 訳);
  else 出.push(f);
}

console.log('★「わざと 壊して 赤に なるか」と 名乗る 試験★');
console.log('  ★本当に 壊している = ' + 良い.length + '本★');
良い.forEach((v) => console.log('    ok   ' + v));
if (壊す無し.length) {
  console.log('  ★--self-test の かたまりが 見つからない = ' + 壊す無し.length + '本★');
  壊す無し.forEach((v) => console.log('    ??   ' + v));
}
console.log('\n★名前と 中身が 違う（壊していないのに そう 名乗る）= ' + 出.length + '本★');
出.forEach((v) => console.log('  ★NG★ ' + v));
if (出.length) {
  console.log('\n★直し方★ … どちらか');
  console.log('  (1) 見出しを 本当の 事に 直す（例：「この 見張りが 見ている物を 直に 確かめる」）');
  console.log('  (2) 実際に 壊す（tools/break-check.mjs と 同じ やり方）');
}
process.exit((出.length + 壊す無し.length) ? 1 : 0);
