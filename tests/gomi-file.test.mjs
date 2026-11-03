/* gomi-file.test.mjs — ★書き損じで 出来た ゴミを repo に 置かない★（2026-09-16）
 *
 *  ★★なぜ 要るか（★同じ 型を 今日 2回 踏みました★）★★
 *    ①`... 2>&1 | tail` を 書き損じて ★`=` という 名前の 空ファイル★
 *    ②`python -c "..."` の 中の ★`>`★ を 貝殻が ★リダイレクト★に 取り、
 *      ★`①★この` `②実Excel` など 5本の 空ファイル★が 出来た
 *    ⇒★どちらも commit まで 通って しまいました★
 *
 *  ★★人の 心がけでは 止まりません★★
 *    ・出来た その場では ★気づきません★（`git status` を 見て 初めて）
 *    ・★名前が 読めないので 目でも 拾いにくい★
 *    ⇒★機械で 塞ぎます★
 *
 *  ★★この 見張りが 見る 物★★
 *    ①★repo の ファイル名は 全部 ASCII★
 *      ＝★日本語の 名前の ファイルは 1本も 在りません★（★今 実際に 0本★）
 *      ⇒★出来たら その場で 赤★
 *    ②★0バイトの ファイルが 在ったら 赤★（★書き損じは たいてい 空★）
 *      ＝★わざと 置く 空ファイルが 要る 時は ここに 名前を 書く★
 *
 *  使い方: node tests/gomi-file.test.mjs
 */
import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
let pass = 0, fail = 0;
const T = (n, f) => {
  try { f(); pass++; console.log('  ✓ ' + n); }
  catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + e.message); }
};

console.log('\n[gomi-file] ★書き損じで 出来た ゴミを repo に 置かない★');

/* ★名簿は git 本人から 読む★（★歩いて 拾うと `.gitignore` の 物まで 見ます★） */
const 一覧 = execFileSync('git', ['-C', ROOT, 'ls-files', '-z'], { maxBuffer: 64 * 1024 * 1024 })
  .toString('utf-8').split('\0').filter(Boolean);

console.log('      … repo の ファイル ' + 一覧.length + '本');

/* ★わざと 置く 空ファイル★（★在るなら ここに 名前を 書く★） */
const 空でよい = [];

T('★名簿が 読めて いる（★空振りして いない★）★', () => {
  if (一覧.length < 100) throw new Error('★' + 一覧.length + '本しか 読めない★（git が 読めて いない）');
  for (const n of ['package.json', 'tests/run.js']) {
    if (一覧.indexOf(n) < 0) throw new Error('★' + n + ' が 名簿に 無い★');
  }
});

T('★★ファイル名は 全部 ASCII（★書き損じの ゴミを 止める★）★★', () => {
  /* ★2026-09-16 に 2回 踏みました★
       ①`=` という 名前の 空ファイル
       ②`python -c` の 中の `>` を 貝殻が リダイレクトに 取り 5本 */
  const 変 = 一覧.filter((n) => !/^[\x20-\x7e]+$/.test(n));
  if (変.length) {
    throw new Error('★' + 変.length + '本 ASCII で ない 名前★ … '
      + 変.slice(0, 8).map((n) => JSON.stringify(n)).join(' ')
      + '／★貝殻の 書き損じで 出来た ゴミの 疑い★（`>` が リダイレクトに なった 等）');
  }
});

T('★★0バイトの ファイルが 無い（★書き損じは たいてい 空★）★★', () => {
  const 空 = [];
  for (const n of 一覧) {
    if (空でよい.indexOf(n) >= 0) continue;
    const p = path.join(ROOT, n);
    let st;
    try { st = fs.statSync(p); } catch (e) { continue; }   /* ★消えた 物は 別の 見張りの 仕事★ */
    if (st.isFile() && st.size === 0) 空.push(n);
  }
  if (空.length) {
    throw new Error('★' + 空.length + '本 0バイト★ … ' + 空.slice(0, 8).join(' ')
      + '／★わざと 置くなら この 見張りの `空でよい` に 名前を 書く★');
  }
});

T('★名前に 危ない 字が 入って いない★', () => {
  /* ★`=` `|` `>` `<` `"` は ★貝殻の 書き損じで 出来た 名前の 印★★ */
  const 危 = 一覧.filter((n) => /[=|><"*?]/.test(path.basename(n)));
  if (危.length) {
    throw new Error('★' + 危.length + '本 名前に 危ない 字★ … ' + 危.slice(0, 8).join(' '));
  }
});

console.log('\n' + pass + ' passed, ' + fail + ' failed');
if (fail) process.exit(1);
