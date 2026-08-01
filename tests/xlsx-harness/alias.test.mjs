/* alias.test.mjs — ★日本語UIの表示名(JIS)を打っても、本名(DBCS)として動くことを見張る★
 *
 * なぜ必要か（2026-08-01 第3波P1で実測して分かった）:
 *   半角→全角の関数、日本語Excelの画面では「JIS」と出るが、
 *   【xlsx の中身も US-English構文も本名は DBCS】。実Excelに `=JIS(A1)` を .Formula で入れると
 *   #NAME? になる（実測）。つまり:
 *     ・JIS のままエンジンへ渡す  → 計算できない
 *     ・JIS のまま xlsx へ書き出す → Excelがその式を #NAME? にする
 *   日本語Excelの癖で JIS と打つ人は普通に居るので、入口(convertFormula)と
 *   出口(xlsx書き出し)の両方で本名 DBCS に寄せる。ここはその2箇所が効いているかの見張り。
 *
 * 使い方: node tests/xlsx-harness/alias.test.mjs
 */
import { createRequire } from 'node:module';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const require = createRequire(import.meta.url);
const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..', '..');
const F = require(path.join(ROOT, 'exally-formula.js'));
const IO = require(path.join(ROOT, 'tests/xlsx-harness/xlsx-io.js'));

let pass = 0, fail = 0;
function T(n, fn) { try { fn(); pass++; console.log('  ✓ ' + n); } catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + (e && e.message)); } }
function eq(got, want, what) { if (got !== want) throw new Error((what || '') + ' 期待=' + JSON.stringify(want) + ' 実際=' + JSON.stringify(got)); }

console.log('\n[alias] 日本語UI名 JIS → 本名 DBCS（入口=エンジン / 出口=xlsx書き出し）');

// ── 入口: convertFormula（エンジンへ渡す前の正規化） ──
T('① convertFormula が JIS( を DBCS( に直す', () => {
  eq(F.convertFormula('=JIS(A1)'), '=DBCS(A1)');
  eq(F.convertFormula('=jis(A1)'), '=DBCS(A1)');
  eq(F.convertFormula('=LEN(JIS(A1))'), '=LEN(DBCS(A1))', '入れ子でも');
});
T('② ASC は本名なのでそのまま', () => {
  eq(F.convertFormula('=ASC(A1)'), '=ASC(A1)');
});
T('③ 文字列の中の「JIS」という語は書き換えない（誤爆しない）', () => {
  // 関数呼び出しでなければ触らない。"JIS規格" のような文字は残す
  eq(F.convertFormula('="JIS規格"'), '="JIS規格"');
});

// ── 変換の中身そのもの ──
T('④ DBCS/ASC の変換結果が実Excelの真値と同じ（濁点の合成/分解を含む）', () => {
  eq(F._jsDbcs('ｱｲｳ'), 'アイウ', 'DBCS(半角カナ)');
  eq(F._jsDbcs('ｶﾞｷﾞ'), 'ガギ', 'DBCS(濁点は合成)');
  eq(F._jsDbcs('ABC123'), 'ＡＢＣ１２３', 'DBCS(半角英数)');
  eq(F._jsAsc('ＡＢＣ１２３'), 'ABC123', 'ASC(全角英数)');
  eq(F._jsAsc('アイウ'), 'ｱｲｳ', 'ASC(全角カナ)');
  eq(F._jsAsc('ガギ'), 'ｶﾞｷﾞ', 'ASC(濁点は分解)');
  eq(F._jsAsc('Ａ亜１'), 'A亜1', 'ASC(漢字は変えない)');
});
T('⑤ ASC と DBCS は往復しても壊れない', () => {
  for (const s of ['ｱｲｳ', 'ｶﾞｷﾞ', 'ABC123', 'ﾊﾟﾋﾟ']) eq(F._jsAsc(F._jsDbcs(s)), s, s);
});

// ── 出口: xlsx 書き出し ──
T('⑥ xlsxへ書き出す時 JIS( が DBCS( になる（JISのまま書くとExcelが #NAME? にする）', () => {
  if (typeof IO.applyXlfn !== 'function') throw new Error('xlsx-io.js が applyXlfn を公開していない');
  eq(IO.applyXlfn('JIS(A1)'), 'DBCS(A1)');
  eq(IO.applyXlfn('LEN(JIS(A1))'), 'LEN(DBCS(A1))', '入れ子でも');
  eq(IO.applyXlfn('"JIS規格"'), '"JIS規格"', '文字列の中は触らない');
});
T('⑦ 既存の _xlfn. 付与を壊していない', () => {
  eq(IO.applyXlfn('XLOOKUP(1,A1:A2,B1:B2)'), '_xlfn.XLOOKUP(1,A1:A2,B1:B2)');
  eq(IO.applyXlfn('SUM(A1:A2)'), 'SUM(A1:A2)', '古い関数には付けない');
});

console.log('\n' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
