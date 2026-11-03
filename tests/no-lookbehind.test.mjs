/* no-lookbehind.test.mjs — ★後読み正規表現 `(?<=` `(?<!` を 増やさせない★ 2026-08-29
 *
 *  ★なぜ在るか★
 *    後読みは ★旧iOS Safari(<16.4) が 読めない★。しかも 読めないのは 実行時ではなく
 *    ★読み込んだ瞬間（構文解析）★なので、★その <script> の塊が 丸ごと 動かない★。
 *    lib/xlsx-io.js・lib/excel-version.js・kyuyo/lib/pay-parse.js には
 *    「後読みは使うな」と ★書いてある★。それでも book.html に 4件 入っていた。
 *    ⇒ ★決まりを 書くだけでは 止まらない★。だから 見張りを 置く。
 *
 *  ★★この見張りの 作りで 一番 大事な所★★
 *    ★「(?<」の 字を 探すな＝コメントの中の「後読み(?<=)は使うな」に 当たる★
 *    実際 2026-08-29 に、字で探した人が
 *      「後読みは 4ファイルに 在る（kyuyo/lib/pay-parse.js / lib/excel-version.js / lib/xlsx-io.js / book.html）」
 *    と 報告した。★中3つは 全部 コメント＝「使うな」と 書いてある行★だった。
 *    ⇒ ★コメントを 落としてから 数える★。
 *    （同じ型を うちで 2回 踏んでいる。docs の「人が書いた」を 字で探して
 *      自分の撤回文に 当たった のも これ）
 *    ★文字列は 落とさない★＝`new RegExp('(?<![A-Za-z])'+x)` は ★本物の使用★だから。
 *
 *  走らせ方: node tests/no-lookbehind.test.mjs [--self-test]
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { 注記を外す } from '../scripts/lib/chuki.mjs';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const 壊す = process.argv.includes('--self-test');
let 緑 = 0, 赤 = 0;
const ok = (名, 条件, 添え) => {
  if (条件) { 緑++; console.log('  ok   ' + 名); }
  else { 赤++; console.log('  ★NG★ ' + 名 + (添え !== undefined ? '  … ' + 添え : '')); }
};

/* ★今 在る物の 名簿★（★減ったら この行を 消す★／★増えたら 赤★）
 *  2026-08-29 実測。全部 book.html の 中。
 *  ★直す時は 1ファイルずつ★（指示役 2026-08-29）。 */
export const 名簿 = {
  /* ★2026-08-29 に 4件 全部 直した★（book.html）。
     ⇒ ★名簿は 空★＝以後 1件でも 増えたら 赤。
     直し方は「前を1文字 食べて 置き換える時に 戻す」。
     ★実物 19,321本の式で 前の書き方と 突き合わせ＝答えが違った式 0本★ */
  /* ★2026-10-11★ 禁止の 字の 見張り＝全席 共通の 物を 逐語で 写した（指示役）。node の 中でだけ 走る
     （客の ブラウザには 届かない＝古い Safari の 穴に 当たらない）。字を 替えると 他の 席と blob が 揃わない */
  'tests/kinshi-ji.test.mjs': 2,
};

/* ── コメントを 落とす（★字を 探さない★の 中身）─────────────── */
export function コメントを落とす(src, html) {
  /* ★自前で 書かない＝共有部品 scripts/lib/chuki.mjs を 使う★（指示役 2026-08-26 の決まり）。
     最初 自前で 書いたら tests/chuki.test.mjs が ★赤で 捕まえた★（2026-08-29）。
     共有部品の方が 正しい＝★字の中の `//` を 注記と 読まない★
     （`'https://example.com'` の // を 消すと、その先の 本物の後読みを 見落とす）。 */
  return 注記を外す(String(src), html ? { html: true } : {});
}

export function 後読みを数える(src, html) {
  const 本文 = コメントを落とす(src, html);
  const m = 本文.match(/\(\?<[=!]/g);
  return m ? m.length : 0;
}

/* ── 走査 ─────────────────────────────────────── */
/* ★この見張り自身は 数から 外す★
 *   中に「わざと足した後読み」の 見本を 持っているので、数えると 必ず 赤になる。
 *   ★外すのは この1本だけ★（他の tests/ は 数える＝試験の中で 後読みを 使っても
 *   旧iPhoneには 出ないが、うっかり 本番へ 写す事が あるため）。 */
const 自分 = 'tests/no-lookbehind.test.mjs';
const 除く = /(^|[\\/])(node_modules|\.git|hyperformula\.full\.min\.js|xlsx\.full\.min\.js)([\\/]|$)/;
function 拾う(dir, 出) {
  for (const 名 of fs.readdirSync(dir)) {
    const p = path.join(dir, 名);
    const 相対 = path.relative(ROOT, p).split(path.sep).join('/');
    if (除く.test('/' + 相対)) continue;
    const st = fs.statSync(p);
    if (st.isDirectory()) { 拾う(p, 出); continue; }
    if (!/\.(js|mjs|cjs|html)$/i.test(名)) continue;
    出.push(相対);
  }
  return 出;
}
const ファイル達 = 拾う(ROOT, []).filter((f) => f !== 自分);
const 見つけた = {};
for (const f of ファイル達) {
  const n = 後読みを数える(fs.readFileSync(path.join(ROOT, f), 'utf8'), /\.html$/i.test(f));
  if (n) 見つけた[f] = n;
}

console.log('\n[① 読んだ数（空振りしていないか）]');
ok('★.js/.mjs/.html を 30本以上 読んだ', ファイル達.length >= 30, ファイル達.length + '本');

console.log('\n[② 名簿と 実物が 合っているか]');
const 全部 = Array.from(new Set([...Object.keys(名簿), ...Object.keys(見つけた)])).sort();
for (const f of 全部) {
  const 名 = 名簿[f] || 0, 実 = 見つけた[f] || 0;
  if (実 > 名) ok('★増えている★ ' + f, false, '名簿 ' + 名 + ' → 実物 ' + 実 + '（★後読みを 足さない★）');
  else if (実 < 名) ok('★減った（名簿を 直す）★ ' + f, false, '名簿 ' + 名 + ' → 実物 ' + 実 + '（★名簿の数を ' + 実 + ' に 書き換える★）');
  else ok('合っている ' + f + ' … ' + 実 + '件', true);
}
ok('★名簿に 無いファイルが 増えていない★',
  Object.keys(見つけた).every((f) => 名簿[f] !== undefined),
  Object.keys(見つけた).filter((f) => 名簿[f] === undefined).join(', '));

console.log('\n[③ ★コメントの中の「後読み(?<=)は使うな」を 数えない★]');
const 見本コメント = '// ★後読み(?<=)は使わない＝旧iOS Safari で壊れる\nvar a = 1;';
ok('行コメントの中は 数えない', 後読みを数える(見本コメント) === 0, String(後読みを数える(見本コメント)));
ok('ブロックコメントの中は 数えない', 後読みを数える('/* 後読み(?<!x) は 使うな */\nvar b=2;') === 0);
ok('★文字列の中は 数える（new RegExp は 本物の使用）★',
  後読みを数える("var r = new RegExp('(?<![A-Za-z])'+x,'g');") === 1,
  String(後読みを数える("var r = new RegExp('(?<![A-Za-z])'+x,'g');")));
ok('★正規表現リテラルは 数える★', 後読みを数える('var r = /(?<=a)b/g;') === 1);
ok('HTMLのコメントも 落とす', 後読みを数える('<!-- 後読み(?<=) は 使うな -->\n<script>var a=1;</script>', true) === 0);

/* ★実際に 誤報が 起きた3ファイルで 確かめる★（2026-08-29） */
console.log('\n[④ ★誤報が起きた3ファイル＝0件と 数えられるか★]');
for (const f of ['lib/xlsx-io.js', 'lib/excel-version.js', 'kyuyo/lib/pay-parse.js']) {
  const p = path.join(ROOT, f);
  if (!fs.existsSync(p)) { ok('（無い）' + f, true); continue; }
  const 生 = fs.readFileSync(p, 'utf8');
  const 字で = (生.match(/\(\?</g) || []).length;
  const 本当 = 後読みを数える(生);
  ok(f + ' … 字で探すと ' + 字で + '件・★本当は ' + 本当 + '件★', 本当 === 0, '本当 ' + 本当);
}

console.log('\nno-lookbehind: ' + 緑 + '/' + (緑 + 赤) + ' passed');

if (壊す) {
  console.log('\n★--self-test＝わざと 1件 足して 赤に なるか★');
  let 素通り = 0;
  const 足した = 'var x = /(?<=abc)def/g;';
  if (後読みを数える(足した) === 0) { 素通り++; console.log('  ★素通り★ わざと足した 後読みを 数えられなかった'); }
  /* ★字で探す作り（＝誤報の元）に 戻したら、コメントを 数えてしまう事を 見せる★ */
  const 字で数える = (s) => (String(s).match(/\(\?<[=!]/g) || []).length;
  if (字で数える('// 後読み(?<=) は 使うな') === 0) { 素通り++; console.log('  ★素通り★ 誤報の再現に 失敗（見本が おかしい）'); }
  if (後読みを数える('// 後読み(?<=) は 使うな') !== 0) { 素通り++; console.log('  ★素通り★ コメントを 数えてしまった'); }
  if (素通り) { console.log('★抜け道 ' + 素通り + '件★'); process.exit(1); }
  console.log('  ok   抜け道 0件（わざと足せば 数える／コメントは 数えない）');
}
process.exit(赤 ? 1 : 0);
