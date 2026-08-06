/* grid-colwidth.test.mjs — ★渡した相手の画面で ######## にならないか★
 *
 * なぜ必要か（2026-08-06・司さんが実際に踏んだ）:
 *   グリッドから書き出した xlsx を司さんが開いたら、締め日と支払期限が
 *   ★######## （列が狭くて入りきらない印）★ になっていた。中身は正しいのに読めない。
 *
 *   Excelは日付を打つと★自分で列を広げる★。ところがグリッドは
 *   ★人が幅を変えていない列を「幅の指定なし」で書き出す★ので、
 *   受け取った側は既定の狭い幅（8.43文字ぶん）で開く。
 *   ＝ ★日付を書き出すと、相手の画面では必ず ######## になる。★
 *
 *   ★これは見せ物の作り直しではなく、製品の欠陥。★
 *   「2026/8/31」は9文字。既定の幅では入らない。
 *
 * ここで固定すること:
 *   ① 日付が入る列には、日付が収まる幅を必ず付けて書き出す
 *   ② ★人が決めた幅は上書きしない★（狭くしたのが本人の意思なら、それを尊重する）
 *   ③ 日付でない列に勝手な幅を付けない（余計なことをしない）
 *   ④ 長い文字（見出し）で列が足りない時も、読める幅にする
 *
 * 使い方: node tests/grid-colwidth.test.mjs
 *         node tests/grid-colwidth.test.mjs --self-test
 */
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath, pathToFileURL } from 'node:url';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(pathToFileURL(path.join(ROOT, 'package.json')));
const G = require_(path.join(ROOT, 'lib/grid-xlsx.js'));

let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ✓ ' + n); } catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + (e && e.message)); } };
const eq = (a, b, m) => { if (a !== b) throw new Error((m ? m + ': ' : '') + 'expected ' + JSON.stringify(b) + ' got ' + JSON.stringify(a)); };

/* Excelの既定の列幅（文字ぶん）。これ以下だと「2026/8/31」(9文字)は ######## になる。 */
const EXCEL_DEFAULT_CH = 8.43;
/* 見た目の幅(px) → だいたい何文字ぶんか。Excelの標準フォントで 1文字 ≒ 7px + 余白5px。 */
const chOf = px => (px - 5) / 7;

function sheetWith(data, colW) { return [{ name: 'S', data: data, colW: colW || {} }]; }

/* ══ self-test ═══════════════════════════════════════════════════════ */
if (process.argv.includes('--self-test')) {
  console.log('\n[grid-colwidth --self-test] 幅の数え方が合っているか');
  T('★既定の幅では「2026/8/31」(9文字)が入らない＝######## になる', () => {
    if (EXCEL_DEFAULT_CH >= 9) throw new Error('既定の幅の見積りが違う');
  });
  T('px→文字ぶんの換算（80px はだいたい10.7文字ぶん）', () => {
    if (Math.round(chOf(80) * 10) / 10 !== 10.7) throw new Error('換算が違う: ' + chOf(80));
  });
  console.log('\n' + pass + ' passed, ' + fail + ' failed');
  process.exit(fail ? 1 : 0);
}

/* ══ 本番 ═══════════════════════════════════════════════════════════ */
console.log('\n[grid-colwidth] 渡した相手の画面で ######## にならないか');

T('★① 日付の列には、日付が収まる幅が付く（司さんが踏んだ ######## の再発防止）', () => {
  const b = G.gridToBook(sheetWith({ '0,0': { v: '締め日' }, '0,1': { v: '2026/8/31' } }));
  const col = b.sheets[0].cols[1];
  if (!col || !col.wpx) throw new Error('日付の列に幅が付いていない＝相手の画面で ######## になる');
  const ch = chOf(col.wpx);
  if (ch < 9) throw new Error('幅が足りない（' + Math.round(ch * 10) / 10 + '文字ぶん）。「2026/8/31」は9文字');
});

T('★① 式の答えが日付の列にも幅が付く（支払期限＝30日後）', () => {
  const b = G.gridToBook(sheetWith({
    '0,0': { v: '支払期限' },
    '0,1': { f: '=J7+30', v: '=J7+30', d: '46295', numFmt: 'yyyy/m/d' },
  }));
  const col = b.sheets[0].cols[1];
  if (!col || !col.wpx || chOf(col.wpx) < 9) throw new Error('式の答えの日付列に幅が付いていない');
});

T('★② 人が決めた幅は上書きしない（狭くしたのが本人の意思なら尊重する）', () => {
  const b = G.gridToBook(sheetWith({ '0,1': { v: '2026/8/31' } }, { 1: 60 }));
  eq(b.sheets[0].cols[1].wpx, 60, '人が決めた幅');
});

T('★③ 日付でない列に勝手な幅を付けない（余計なことをしない）', () => {
  const b = G.gridToBook(sheetWith({ '0,0': { v: '本店' }, '0,1': { v: '15000' } }));
  const c0 = b.sheets[0].cols[0], c1 = b.sheets[0].cols[1];
  if (c0 && c0.wpx) throw new Error('文字の列に幅を付けている');
  if (c1 && c1.wpx) throw new Error('数の列に幅を付けている');
});

T('★④ 長い見出しでも読める幅にする（「支払期限(30日後)」が切れない）', () => {
  const b = G.gridToBook(sheetWith({ '0,0': { v: '支払期限(30日後)' } }));
  const col = b.sheets[0].cols[0];
  if (!col || !col.wpx) throw new Error('長い見出しの列に幅が付いていない');
  // 日本語は1文字で2文字ぶんの幅を食う（全角）
  if (chOf(col.wpx) < 16) throw new Error('幅が足りない（' + Math.round(chOf(col.wpx)) + '文字ぶん）');
});

T('★短い文字では広げない（見出しがあるだけで列が伸びない）', () => {
  const b = G.gridToBook(sheetWith({ '0,0': { v: '店舗' } }));
  const col = b.sheets[0].cols[0];
  if (col && col.wpx) throw new Error('短い文字で広げている');
});

T('検査が空振りしていない（実際に列の情報を作れている）', () => {
  const b = G.gridToBook(sheetWith({ '0,0': { v: '2026/8/31' } }));
  if (!Array.isArray(b.sheets[0].cols)) throw new Error('列の情報が無い');
});

console.log('\n── 実測 ──');
const demo = G.gridToBook(sheetWith({
  '0,0': { v: '締め日' }, '0,1': { v: '2026/8/31' },
  '1,0': { v: '支払期限(30日後)' }, '1,1': { f: '=B1+30', v: '=B1+30', d: '46295', numFmt: 'yyyy/m/d' },
  '2,0': { v: '本店' }, '2,1': { v: '15000' },
}));
demo.sheets[0].cols.forEach((c, i) => {
  const w = c && c.wpx;
  console.log('  ' + String.fromCharCode(65 + i) + '列: ' + (w ? w + 'px（約' + Math.round(chOf(w)) + '文字ぶん）' : '幅の指定なし＝Excelの既定(約8.4文字ぶん)'));
});

console.log('\n' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
