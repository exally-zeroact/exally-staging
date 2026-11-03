/* shiki-sansho.test.mjs — ★参照（A1・$A$1・A1:B2・板!A1）を 読む★（2026-09-11）
 *
 *  ★★土台を 自分で 作る ③枚目の 前半★★
 *    地図（どの 式が どの 式に 頼って いるか）を 作るには
 *    ★まず 参照が 読めないと 始まりません★
 *
 *  ★★借り物の 中は 1文字も 読んで いません★★
 *    答えは ★実Excel に 聞いて 測りました★（2026-09-11）
 *      A1 → 行1 列1 ／ Z9 → 行9 列26 ／ AA1 → 列27 ／ ★XFD1 → 列16384★
 *      A1048576 → 行1048576 ／ `$A$1` `$A1` `A$1` … ★どれも A1 と 同じ 場所★
 *      A:A → 列1・行 1048576本 ／ 1:1 → 行1・列 16384本
 *
 *  ★実物での 実測★ 司さんの 実物 15,799本 から ★参照 21,178本 全部 読めた★（読めず 0）
 *
 *  ★この 台は エンジンを 建てません★＝字を 読むだけ。
 */
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(path.join(ROOT, 'package.json'));
const S = require_(path.join(ROOT, 'lib/shiki-sansho.js'));

let pass = 0, fail = 0;
const T = (n, f) => {
  try { f(); pass++; console.log('  ✓ ' + n); }
  catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + e.message); }
};

console.log('');
console.log('[shiki-sansho] ★参照を 読む★');

/* ★実Excel に 打たせて 測った 場所★（行・列は 0から 数える） */
const 正 = [
  ['A1', 0, 0, 0, 0], ['Z9', 8, 25, 8, 25], ['AA1', 0, 26, 0, 26],
  ['XFD1', 0, 16383, 0, 16383], ['A1048576', 1048575, 0, 1048575, 0],
  ['$A$1', 0, 0, 0, 0], ['$A1', 0, 0, 0, 0], ['A$1', 0, 0, 0, 0],
  ['A1:B2', 0, 0, 1, 1], ['B2:A1', 0, 0, 1, 1],
];

T('★★場所が 実Excel と 合う（' + 正.length + '本）★★', () => {
  const 違い = [];
  for (const [字, 上, 左, 下, 右] of 正) {
    const r = S.読む(字);
    if (!r.ok) { 違い.push(字 + ' … 読めない'); continue; }
    if (r.上 !== 上 || r.左 !== 左 || r.下 !== 下 || r.右 !== 右) {
      違い.push(字 + ' → 上' + r.上 + ' 左' + r.左 + ' 下' + r.下 + ' 右' + r.右);
    }
  }
  if (違い.length) throw new Error('★' + 違い.length + '本 違う★  ' + 違い.slice(0, 3).join(' ／ '));
});

T('★★まるごと（A:A / 1:1）が 実Excel と 合う★★', () => {
  const a = S.読む('A:A');
  if (!a.ok || a.左 !== 0 || (a.下 - a.上 + 1) !== 1048576) {
    throw new Error('★A:A … ' + JSON.stringify(a) + '★（実Excel は 列1・行 1048576本）');
  }
  const b = S.読む('1:1');
  if (!b.ok || b.上 !== 0 || (b.右 - b.左 + 1) !== 16384) {
    throw new Error('★1:1 … ' + JSON.stringify(b) + '★（実Excel は 行1・列 16384本）');
  }
});

T('★★板の 名前が 取れる（空白・引用符 付き も）★★', () => {
  const 見る = [['Sheet1!A1', 'Sheet1'], ["'あ い'!A1:B2", 'あ い'], ["'それ''これ'!A1", "それ'これ"]];
  const 違い = [];
  for (const [字, 期待] of 見る) {
    const r = S.読む(字);
    if (!r.ok || r.板 !== 期待) 違い.push(字 + ' → ' + JSON.stringify(r.板));
  }
  if (違い.length) throw new Error('★' + 違い.length + '本★  ' + 違い.join(' ／ '));
});

T('★★外れた 物は 断る（勝手に 直さない）★★', () => {
  /* ★実Excel の 端を 1つ 超えた 物★＝黙って 通すと 地図が 狂います */
  const 断る = ['XFE1', 'A1048577', 'A0', 'A', '1', 'A1:', ':A1', 'あ1', 'A1:B2:C3', '!A1'];
  const 通った = [];
  for (const t of 断る) if (S.読む(t).ok) 通った.push(t);
  if (通った.length) throw new Error('★' + 通った.length + '本 通した★  ' + 通った.join(' ／ '));
});

T('★★字に 戻せる（A↔0・Z↔25・AA↔26・XFD↔16383）★★', () => {
  const 違い = [];
  for (const [字, 数] of [['A', 0], ['Z', 25], ['AA', 26], ['AB', 27], ['XFD', 16383]]) {
    if (S.列を数に(字) !== 数) 違い.push(字 + ' → ' + S.列を数に(字));
    if (S.数を列に(数) !== 字) 違い.push(数 + ' → ' + S.数を列に(数));
  }
  if (違い.length) throw new Error('★' + 違い.length + '本★  ' + 違い.join(' ／ '));
});

console.log('');
console.log('  ' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
