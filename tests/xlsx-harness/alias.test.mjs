/* alias.test.mjs — ★日本語UIの表示名を打っても、本名として動くことを見張る★
 *   JIS → DBCS（第3波P1で発見）／ YEN → DOLLAR（第3波P2で発見）
 *
 * なぜ必要か（2026-08-01 に実測して分かった）:
 *   日本語Excelの画面に出る関数名と、【xlsx の中身／US-English構文の本名】が違う物がある。
 *   実Excelに `.Formula` で `=JIS(A1)` / `=YEN(1234.5)` を入れると **#NAME?** になり、
 *   `.FormulaLocal`(日本語名) ↔ `.Formula`(本名) を Excel 自身が相互変換していることも確認した
 *   （`.FormulaLocal="=YEN(1234.5)"` を入れて `.Formula` を読むと `=DOLLAR(1234.5)` が返る）。つまり:
 *     ・表示名のままエンジンへ渡す  → 計算できない
 *     ・表示名のまま xlsx へ書き出す → Excelがその式を #NAME? にする（★こちらが本題）
 *   日本語Excelの癖でこの名前を打つ人は普通に居るので、入口(convertFormula)と
 *   出口(xlsx書き出し)の両方で本名に寄せる。ここはその2箇所が効いているかの見張り。
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

console.log('\n[alias] 日本語UI名 → 本名（JIS→DBCS / YEN→DOLLAR。入口=エンジン / 出口=xlsx書き出し）');

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

// ── YEN → DOLLAR（第3波P2）──
T('⑧ convertFormula が YEN( を DOLLAR( に直す', () => {
  eq(F.convertFormula('=YEN(1234.5)'), '=DOLLAR(1234.5)');
  eq(F.convertFormula('=yen(1234.5)'), '=DOLLAR(1234.5)');
  eq(F.convertFormula('=LEN(YEN(1000))'), '=LEN(DOLLAR(1000))', '入れ子でも');
  eq(F.convertFormula('="YEN建て"'), '="YEN建て"', '文字列の中は触らない');
});
T('⑨ xlsxへ書き出す時 YEN( が DOLLAR( になる', () => {
  eq(IO.applyXlfn('YEN(1234.5)'), 'DOLLAR(1234.5)');
  eq(IO.applyXlfn('LEN(YEN(1000))'), 'LEN(DOLLAR(1000))', '入れ子でも');
});
T('⑩ DOLLAR は地域の通貨書式に従う（この環境=日本語1041 では ¥ ・小数0桁・負はマイナス記号）', () => {
  eq(F._jsDollar(1234.567), '¥1,235', '既定');
  eq(F._jsDollar(1234.567, 2), '¥1,234.57');
  eq(F._jsDollar(-1234.5), '¥-1,235', '負は括弧ではなくマイナス記号');
  eq(F._jsDollar(1234.5, -2), '¥1,200', '負の桁数');
});

console.log('\n' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
