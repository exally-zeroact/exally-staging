/* shiki-sashikomi.test.mjs — ★行/列を 入れ・消しした時 式の 参照が 追従するか★（2026-09-13）
 *
 *  ★★借り物の 中で 一番 手が 込んでいる所★★（LazilyTransformingAstService に あたる）
 *
 *  ★★下の 17通りは 全部 実Excel に 打って 読んだ 物です★★
 *    （2026-09-13・道具 `docs/measured/toru-sashikomi.ps1`／Excel 16.0 build 20326）
 *    ★1つも 当て推量が 在りません★
 *
 *  ★★当て推量なら 外して いた 物★★
 *    `=$A$2+1` → ★`=$A$3+1`★（★絶対参照も 追従する★）
 *    `=SUM(A1:A3)` の 途中に 入れる → ★広がる★／`=SUM(A2:A4)` の 先頭に 入れる → ★ずれる★
 *
 *  ★★見て いない 範囲★★
 *    板を またぐ 参照（`板!A1`）／名前を 付けた 範囲／表（Table[列名]）／溢れ
 */
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(path.join(ROOT, 'package.json'));
const S = require_(path.join(ROOT, 'lib/shiki-sashikomi.js'));

let pass = 0, fail = 0;
const T = (n, f) => {
  try { f(); pass++; console.log('  ✓ ' + n); }
  catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + e.message); }
};

/* ══ ★★実Excel の 答え★★ ══ 〔置いた式, 向き, どこ, 消すか, 後の式, 何を見ているか〕 */
const 実Excel = [
  ['=A2+1', '行', 2, false, '=A3+1', '★指す先の 上に 行を 入れる（A2→A3 に なるか）★'],
  ['=A2+1', '行', 3, false, '=A2+1', '指す先の 下に 入れる（変わらないはず）'],
  ['=A2+1', '行', 2, true, '=#REF!+1', '★指す先を 消す（#REF! に なるか）★'],
  ['=A3+1', '行', 2, true, '=A2+1', '指す先の 上を 消す（A3→A2 に なるか）'],
  ['=$A$2+1', '行', 2, false, '=$A$3+1', '★絶対参照も 追従するか★'],
  ['=SUM(A1:A3)', '行', 2, false, '=SUM(A1:A4)', '★四角の 途中に 入れる（A1:A4 に 広がるか）★'],
  ['=SUM(A2:A4)', '行', 2, false, '=SUM(A3:A5)', '★四角の 先頭に 入れる（A3:A5 か A2:A5 か）★'],
  ['=SUM(A1:A3)', '行', 4, false, '=SUM(A1:A3)', '四角の 下に 入れる（変わらないはず）'],
  ['=SUM(A1:A3)', '行', 2, true, '=SUM(A1:A2)', '★四角の 中を 消す（A1:A2 に 縮むか）★'],
  ['=SUM(A2:A2)', '行', 2, true, '=SUM(#REF!)', '★四角が 丸ごと 消える（#REF! か）★'],
  ['=A2+1', '列', 1, false, '=B2+1', '★左に 列を 入れる（A2→B2 か）★'],
  ['=B2+1', '列', 1, false, '=C2+1', '列を 入れる（B2→C2 か）'],
  ['=B2+1', '列', 2, true, '=#REF!+1', '★指す列を 消す（#REF! か）★'],
  ['=SUM(A1:C1)', '列', 2, false, '=SUM(A1:D1)', '★横の 四角の 途中に 列を 入れる★'],
  ['=A1+A2', '行', 1, false, '=A2+A3', '★式 自身が 下に ずれる時★'],
  ['=A$2+1', '行', 2, false, '=A$3+1', '★行だけ 絶対★'],
  ['=$A2+1', '行', 2, false, '=$A3+1', '★列だけ 絶対★'],
];

console.log('\n[shiki-sashikomi] ★行/列の 入れ消しで 式が 追従するか★');

T('★★実Excel の ' + 実Excel.length + '通りと 1つも 違わない★★', () => {
  const 違い = [];
  for (const [式, 向き, どこ, 消すか, 欲, なぜ] of 実Excel) {
    let 得;
    try { 得 = S.式を直す(式, { 向き, どこ, 台数: 1, 消すか }); }
    catch (e) { 違い.push(式 + ' … ★落ちた★ ' + e.message); continue; }
    if (得 !== 欲) 違い.push(式 + '（' + 向き + どこ + (消すか ? '消す' : '入れる') + '）… うち `' + 得 + '` ／ ★実Excel `' + 欲 + '`★');
  }
  if (違い.length) {
    throw new Error('★' + 違い.length + ' / ' + 実Excel.length + ' 違う★' + String.fromCharCode(10) + '      ' + 違い.slice(0, 8).join(String.fromCharCode(10) + '      '));
  }
});

T('★★絶対参照も 追従する（「絶対」という 名前なのに 動く）★★', () => {
  const 得 = S.式を直す('=$A$2+1', { 向き: '行', どこ: 2, 台数: 1, 消すか: false });
  if (得 !== '=$A$3+1') throw new Error('★' + 得 + '★（実Excel は =$A$3+1）');
});

T('★★四角は 途中なら 広がり、先頭なら ずれる（1行 違うだけで 変わる）★★', () => {
  const 途中 = S.式を直す('=SUM(A1:A3)', { 向き: '行', どこ: 2, 台数: 1, 消すか: false });
  const 先頭 = S.式を直す('=SUM(A2:A4)', { 向き: '行', どこ: 2, 台数: 1, 消すか: false });
  if (途中 !== '=SUM(A1:A4)') throw new Error('途中 … ★' + 途中 + '★');
  if (先頭 !== '=SUM(A3:A5)') throw new Error('先頭 … ★' + 先頭 + '★');
  if (途中 === 先頭) throw new Error('★途中と 先頭を 同じに 扱って いる★');
});

T('★★消して 指す先が 無くなったら #REF!★★', () => {
  const 一つ = S.式を直す('=A2+1', { 向き: '行', どこ: 2, 台数: 1, 消すか: true });
  const 四角 = S.式を直す('=SUM(A2:A2)', { 向き: '行', どこ: 2, 台数: 1, 消すか: true });
  if (一つ !== '=#REF!+1') throw new Error('1マス … ★' + 一つ + '★');
  if (四角 !== '=SUM(#REF!)') throw new Error('四角 … ★' + 四角 + '★');
});

T('★字の 中の「A2」は 触らない（式の 参照だけ 直す）★', () => {
  const 得 = S.式を直す('="A2"&A2', { 向き: '行', どこ: 2, 台数: 1, 消すか: false });
  if (得 !== '="A2"&A3') throw new Error('★' + 得 + '★（字の 中まで 書き換えて いる）');
});

T('★行を 入れた時 列は 動かない（もう 片方の 向きは そのまま）★', () => {
  const 得 = S.式を直す('=$B$2+1', { 向き: '行', どこ: 2, 台数: 1, 消すか: false });
  if (得 !== '=$B$3+1') throw new Error('★' + 得 + '★（列まで 動かして いる）');
});

console.log('');
console.log('★締め★ 通った ' + pass + ' ／ ★落ちた ' + fail + '★');
process.exit(fail ? 1 : 0);
