/* grid-date.test.mjs — ★グリッドで日付が日付として扱われるか★
 *
 * なぜ必要か（2026-08-05・グリッドのExcel保存を見せる物を作っている最中に見つけた）:
 *   ① 「2026/8/31」と打っても日付にならない
 *        =その+30 → ★2056★（2026＋30 と、ただの数字の足し算になった）
 *        =TEXT(その,"yyyy年m月d日") → "2026/8/31" のまま（書式が効かない）
 *   ② =DATE(2026,8,31) は中身は正しいのに ★画面に 46265 という数字のまま出る★
 *   ③ ②を「表示形式」で日付にしようとすると
 *        ★アプリ自身が「46265 (日付として認識されません)」と出す★
 *
 *   ★「日付を打ったら日付になる」はExcelの基本★。ここが違うと、締め日の30日後・
 *   支払期限といった ★日付の計算がまるごとできない★。
 *
 * どう直すか（Excelと同じ考え方にそろえる）:
 *   Excelは日付を「1900年から数えて何日目か」という★数字（シリアル値）★で持ち、
 *   見せる時だけ日付の顔にする。グリッドもそれに合わせる。
 *     ・打った日付文字列 → 計算にはシリアル値を渡す（+30 が効く）
 *     ・そのセルには日付の表示形式を自動で付ける（画面は「2026/8/31」のまま）
 *     ・シリアル値も日付として読めるようにする（=DATE の結果が日付として扱える）
 *
 * ★書き出し(xlsx)も同じ規則でないと、画面と落としたファイルが食い違う★ので、
 *   lib/grid-xlsx.js の asValue と book.html の toHFVal が同じ答えを出すことも見る
 *   （tests/grid-xlsx.test.mjs ⑬ と対になる守り）。
 *
 * 使い方: node tests/grid-date.test.mjs
 *         node tests/grid-date.test.mjs --self-test
 */
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(import.meta.url);
const SRC = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf8');
const G = require_(path.join(ROOT, 'lib/grid-xlsx.js'));

let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ✓ ' + n); } catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + (e && e.message)); } };
const eq = (a, b, m) => { if (a !== b) throw new Error((m ? m + ': ' : '') + 'expected ' + JSON.stringify(b) + ' got ' + JSON.stringify(a)); };

/* book.html から関数を1つ取り出して動かす（人が書き写さない） */
function grab(name, argsRe) {
  const re = new RegExp('function ' + name + '\\(' + (argsRe || '[^)]*') + '\\)\\s*\\{[\\s\\S]*?\\n\\}');
  const m = re.exec(SRC);
  if (!m) throw new Error('book.html から ' + name + ' を取り出せない（名前が変わった？）');
  return m[0];
}
/* いくつかの関数をまとめて1つの入れ物で動かす（お互いを呼べるように） */
function build(names) {
  const body = names.map(n => grab(n)).join('\n');
  return new Function(body + '\nreturn {' + names.map(n => n + ':' + n).join(',') + '};')();
}

/* ★純関数: 日付文字列 → Excelのシリアル値。self-testで作り物を通せる。
   Excelは 1899-12-30 を 0 として数える（1900年うるう年の歴史的なズレを含む数え方）。 */
export function serialOf(y, m, d) {
  return Math.round((Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000);
}

/* ══ self-test ═══════════════════════════════════════════════════════ */
if (process.argv.includes('--self-test')) {
  console.log('\n[grid-date --self-test] 数え方そのものが合っているか');
  T('★Excelと同じ数え方（2026/8/31 は 46265日目）', () => eq(serialOf(2026, 8, 31), 46265));
  T('1900/1/1 は 2 日目（Excelの数え方の癖もそのまま）', () => eq(serialOf(1900, 1, 1), 2));
  T('30日後は +30 になる', () => eq(serialOf(2026, 9, 30) - serialOf(2026, 8, 31), 30));
  console.log('\n' + pass + ' passed, ' + fail + ' failed');
  process.exit(fail ? 1 : 0);
}

/* ══ 本番（実物を見る） ═══════════════════════════════════════════════ */
console.log('\n[grid-date] グリッドで日付が日付として扱われるか');

// toHFVal / parseDate は下の2つを呼ぶので、一緒に取り出して同じ入れ物で動かす
const B = build(['dateSerial', 'parseDateStr', 'toHFVal', 'parseDate']);

T('★① 打った日付「2026/8/31」が計算に使える数（シリアル値）になる', () => {
  const v = B.toHFVal('2026/8/31');
  if (typeof v !== 'number') throw new Error('数になっていない（' + JSON.stringify(v) + '）＝ +30 が 2056 になる原因');
  eq(v, serialOf(2026, 8, 31), 'シリアル値');
});

T('★① 30日後がちゃんと30日後になる（2026/8/31 → 2026/9/30）', () => {
  const a = B.toHFVal('2026/8/31');
  eq(a + 30, serialOf(2026, 9, 30), '30日足した結果');
});

T('★②③ シリアル値(46265)も日付として読める（=DATE の結果を日付と認める）', () => {
  const pd = B.parseDate(46265);
  if (!pd) throw new Error('日付として読めない＝「46265 (日付として認識されません)」が出る');
  eq(pd.y + '/' + pd.m + '/' + pd.d, '2026/8/31', '読み取った日付');
});

T('★②③ 文字で来たシリアル値("46265")も同じに読める', () => {
  const pd = B.parseDate('46265');
  if (!pd) throw new Error('日付として読めない');
  eq(pd.y + '/' + pd.m + '/' + pd.d, '2026/8/31');
});

T('今までどおり日付の文字列も読める（直したせいで前が壊れていない）', () => {
  const pd = B.parseDate('2026-08-31');
  if (!pd) throw new Error('読めなくなった');
  eq(pd.y + '/' + pd.m + '/' + pd.d, '2026/8/31');
});

T('★日付でない物を日付にしない（誤検知を出さない）', () => {
  for (const s of ['あ', '007-1234', '2026/13/40', '', 'TRUE']) {
    if (B.parseDate(s)) throw new Error('日付でない物を日付にした: ' + JSON.stringify(s));
  }
});

T('★小さい数は日付にしない（0007 のような番号を日付に化けさせない）', () => {
  if (B.parseDate('0007')) throw new Error('社員番号などが日付に化ける');
  if (B.parseDate(7)) throw new Error('小さい数が日付に化ける');
});

/* ★シリアル値も日付として読めるようにした副作用を、ここで塞ぐ。
   「日付の書式を頼まれた時」だけ数を日付と見なす。それ以外の場所で数を日付にすると、
   売上 15000 が日付に化ける。実際に直している最中に3箇所で踏みかけた。 */
T('★売上の数(15000)を日付に化けさせない — 文字列専用の口が要る', () => {
  if (B.parseDateStr('15000')) throw new Error('parseDateStr が数を日付にしている');
  if (B.parseDateStr(15000)) throw new Error('parseDateStr が数を日付にしている');
  if (!B.parseDate(15000)) throw new Error('parseDate はシリアル値を読めるべき（=DATEの結果のため）');
});

T('★セルに打った時、日付だけを日付にする（売上や件数は数のまま）', () => {
  const SRC2 = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf8');
  const set = /function setCell\(r,c,v,noUndo\)\{[\s\S]*?\n\}/.exec(SRC2);
  if (!set) throw new Error('book.html から setCell を取り出せない＝この検査が空振り');
  if (/parseDate\(v\)/.test(set[0])) {
    throw new Error('setCell が parseDate(v) を使っている＝15000 のような数が日付になる（parseDateStr を使うこと）');
  }
  if (!/parseDateStr\(v\)/.test(set[0])) throw new Error('日付の判定が見当たらない＝この検査が空振り');
  if (!/dateSerial\(/.test(set[0])) throw new Error('日付を数にして渡していない＝+30 が 2056 のままになる');
});

T('★連続コピー(オートフィル)も数を日付にしない', () => {
  const SRC2 = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf8');
  const fn = /function detectDatePattern\(values\)\{[\s\S]*?\n\}/.exec(SRC2);
  if (!fn) throw new Error('detectDatePattern を取り出せない＝この検査が空振り');
  if (/[^S]parseDate\(values/.test(fn[0])) throw new Error('parseDate を使っている＝売上の並びが日付の連番になる');
});

T('★書き出し(xlsx)も画面と同じ規則（ズレると落としたファイルだけ日付が違う）', () => {
  const ng = [];
  for (const s of ['2026/8/31', '2026-08-31', '0007', '1,234', '15000', 'あ', '']) {
    const a = G.asValue(s), b = B.toHFVal(s);
    const same = (a === '' && b === null) || a === b;
    if (!same) ng.push(JSON.stringify(s) + ': 書き出し=' + JSON.stringify(a) + ' / 画面=' + JSON.stringify(b));
  }
  if (ng.length) throw new Error('食い違い:\n      ' + ng.join('\n      '));
});

console.log('\n── 実測 ──');
try { console.log('  toHFVal("2026/8/31") = ' + JSON.stringify(B.toHFVal('2026/8/31')) + '（Excelの数え方だと ' + serialOf(2026, 8, 31) + '）'); } catch (e) { console.log('  ' + e.message); }
try { console.log('  parseDate(46265)     = ' + JSON.stringify(B.parseDate(46265))); } catch (e) { console.log('  ' + e.message); }

console.log('\n' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
