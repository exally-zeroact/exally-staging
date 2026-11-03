/* type-to-edit.test.mjs — ★セルを 選んで いきなり 打つと そのまま 入る★ 2026-08-30
 *
 *  ★なぜ★（2026-08-30 実ブラウザで 実測して 見つけた）
 *    Excel の ★一番 基本の 手★＝A1 を 選んで「あ」と 打てば「あ」が 入る。
 *
 *    ★うちは 違う 字が 入っていた★
 *      「あいう」と 打つと ★「いうあ」★（★1文字目が 末尾に 回る★）
 *      英数（ABC）と 数式（=1+2）は ★たまたま 合っていた★
 *      ⇒ ★だから 今まで 誰も 気づかなかった★
 *      ⇒ ★入らないより「黙って 違う字が 入る」方が 悪い★（客は 気づかない）
 *
 *  ★根っこ（実測で 押さえた）★
 *    1文字目で 編集を 始める時に 見た目（left/top/width/height/class）を 変える。
 *    ★その 再描画で カーソルが 先頭（selectionStart=0）に 戻る★。
 *    2文字目以降が ★先頭に 入る★ので 順番が ひっくり返る。
 *    英数は なぜか 1 のままだった＝★たまたま★。
 *
 *  ★直し★ book.html の cell-input の input で
 *    編集を 始めた 直後に ★カーソルを 末尾へ 戻す★。
 *    ★IME の 変換中（_composing）は 触らない★（変換が 壊れる）。
 *
 *  ★この 見張りが 見る物★
 *    ① 直しが ソースに 在る（カーソルを 末尾へ 戻している）
 *    ② ★IME の 変換中は 触らない★（守りが 効いている）
 *    ③ 直しの 場所が ★編集を 始める 所の 中★（別の 場所に 移っていない）
 *
 *  ★本物の 確かめは 実ブラウザ★（node では キーを 打てない）
 *    → tools/type-check.md に ★実測した 表★を 置いてある（直す前／直した後）
 *
 *  走らせ方: node tests/type-to-edit.test.mjs  ／  --self-test
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
let ok = 0, ng = 0;
const 言う = (よい, 文, 添え) => {
  if (よい) { ok++; console.log('  ok   ' + 文); }
  else { ng++; console.log('  NG   ' + 文); if (添え) console.log('       ' + 添え); }
};

const book = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf8');
console.log('★セルを 選んで いきなり 打つ★');
言う(book.length > 100000, '★空振りしていない★（book.html を 読めた ' + book.length + '字）');

/* ★編集を 始める かたまり★を 取り出す（★注記は 外してから 数える★＝
   注記の 字に 当たって「在る」と 誤判定する のを 防ぐ。2026-08-30 実際に 踏んだ） */
/* ★注記外しは 自前で 書かない★＝★共通の 部品を 使う★
   （自前だと URL の // や 字の中の /* を 間違える。
     見張り tests/chuki.test.mjs が ★自前を 赤に する★） */
const { 注記を外す } = await import(
  pathToFileURL(path.join(ROOT, 'scripts/lib/chuki.mjs')).href);
const 頭 = book.indexOf("if(!editingCell && v.length>0){");
言う(頭 >= 0, '★編集を 始める 所が 在る（!editingCell && v.length>0）★');
let 塊 = '';
if (頭 >= 0) {
  let 深 = 0, i = book.indexOf('{', 頭);
  for (let k = i; k < book.length; k++) {
    if (book[k] === '{') 深++;
    else if (book[k] === '}') { 深--; if (深 === 0) { 塊 = book.slice(頭, k + 1); break; } }
  }
}
const 塊素 = 注記を外す(塊);

/* ── ① カーソルを 末尾へ 戻しているか ── */
const 戻す = /setSelectionRange\s*\(\s*v\.length\s*,\s*v\.length\s*\)/.test(塊素);
言う(戻す, '★★編集を 始めた 直後に カーソルを 末尾へ 戻している★★',
  '★戻さないと「あいう」→「いうあ」に なる（2026-08-30 実測）★');

/* ── ② IME の 変換中は 触らないか ── */
const 守り = /if\s*\(\s*!\s*_composing\s*\)/.test(塊素);
言う(守り, '★IME の 変換中は カーソルを 触らない（_composing を 見ている）★',
  '★変換中に 触ると 変換が 壊れる★');

/* ── ③ 直しが 編集を 始める 所の 中に 在るか ── */
言う(戻す && 塊.length > 0 && 塊.length < 3000,
  '★直しは 編集を 始める 所の 中に 在る（別の 場所へ 移していない）★',
  '塊の 長さ ' + 塊.length);

/* ── ④ 実ブラウザの 実測が 紙に 残っているか ── */
const 紙 = path.join(ROOT, 'tools/type-check.md');
言う(fs.existsSync(紙), '★実ブラウザで 測った 表が 残っている（tools/type-check.md）★',
  '★node では キーを 打てない＝実測の 記録が 要る★');
if (fs.existsSync(紙)) {
  const t = fs.readFileSync(紙, 'utf8');
  言う(/いうあ/.test(t), '　その 紙に ★直す前の 症状（いうあ）★が 書いてある');
  言う(/あいうえお/.test(t) && /IME/.test(t), '　その 紙に ★直した後の 6通り★が 書いてある');
}

/* ── わざと 壊して 赤に なるか ── */
if (process.argv.includes('--self-test')) {
  console.log('\n★わざと 壊して 赤に なるか★');
  const にせ1 = 塊素.replace(/setSelectionRange\s*\(\s*v\.length\s*,\s*v\.length\s*\)/, 'noop()');
  言う(!/setSelectionRange\s*\(\s*v\.length\s*,\s*v\.length\s*\)/.test(にせ1),
    '★カーソルを 戻すのを やめたら ①が 落ちる★');
  const にせ2 = 塊素.replace(/if\s*\(\s*!\s*_composing\s*\)/, 'if (true)');
  言う(!/if\s*\(\s*!\s*_composing\s*\)/.test(にせ2),
    '★IME の 守りを 外したら ②が 落ちる★');
  /* ★注記だけに 在る 字で 緑に ならないか★（2026-08-30 に 私が 踏んだ 罠） */
  const 注記だけ = '/* setSelectionRange(v.length, v.length) と 書いてあるだけの 注記 */';
  言う(!/setSelectionRange/.test(注記を外す(注記だけ)),
    '★注記に 同じ字が 在っても 緑に ならない（注記を 外して 数えている）★');
  言う(/setSelectionRange\s*\(\s*v\.length\s*,\s*v\.length\s*\)/.test(塊素),
    '★本物は 壊していない（にせ物を 作っただけ）★');
}

console.log('\ntype-to-edit: ' + ok + '/' + (ok + ng) + ' passed');
process.exit(ng ? 1 : 0);
