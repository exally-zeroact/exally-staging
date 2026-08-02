/* bare-form.test.mjs — ★「客が最初に書く形」のケースが無い関数を赤にする★
 *
 * なぜ必要か（2026-08-02 R19 で実際に踏んだ）:
 *   SORT / UNIQUE / FILTER のケースは【全部が他の関数で包んだ形】だった。
 *     =SUM(SORT(E1:E6)) / =INDEX(SORT(...),1) / =TEXTJOIN(",",TRUE,SORT(...)) …
 *   素の =SORT(E1:E6) のケースが1件も無く、その素の形が #VALUE! になっていることに
 *   一致362という数字では気づけなかった。★一番自然な書き方が動いていなかった。
 *
 *   原因は「直した所の形にケースが寄る」こと。P1/P2で入れ子ばかり直したので、
 *   ケースも入れ子ばかりになった。数字は増えるのに、客が最初に書く形は抜ける。
 *
 * 決まり:
 *   グリッドが提供していて、かつハーネスでケースを持っている関数は、
 *   【素の形（一番外側がその関数）のケースを最低1本持つ】こと。
 *   持てない理由がある物は EXCEPTIONS に理由つきで載せる（載っていない物は赤）。
 *
 * 使い方: node tests/xlsx-harness/bare-form.test.mjs
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const CASEDIR = path.join(__dirname, 'cases');

/* ★素の形を持てない関数（理由必須）。ここに無い物が抜けていたら赤。 */
const EXCEPTIONS = {
  TODAY: '揮発性。素の形は日付が動くので volatileCheck 付きの専用ケース(TODAY_serial)で見ている',
  NOW: '揮発性。素の形は時刻が動くので volatileCheck 付きの専用ケース(NOW_int_is_today)で見ている',
};

/* 関数ではない物（演算子・分類名）は「素の形」という考え方が当てはまらないので数えない。
   例: '&'（文字列連結の演算子）／'配列演算'（=C1:C6*E1:E6 のような式の分類名） */
function isFunctionName(s) { return /^[A-Z][A-Z0-9._]*$/i.test(s); }

let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ✓ ' + n); } catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + e.message); } };

/* 式の一番外側の関数名を取る。'=' の直後に来る関数呼び出しが、式全体を包んでいるかを見る。 */
function outermost(formula) {
  const f = String(formula).replace(/^=/, '').trim();
  const m = /^([A-Z][A-Z0-9._]*)\s*\(/i.exec(f);
  if (!m) return null;
  // その ( に対応する ) が式の末尾なら、この関数が一番外側
  let depth = 0, i = m[0].length - 1;
  for (; i < f.length; i++) {
    const c = f[i];
    if (c === '"') { i++; while (i < f.length && f[i] !== '"') i++; continue; }
    if (c === '(') depth++;
    else if (c === ')') { depth--; if (depth === 0) break; }
  }
  return (i === f.length - 1) ? m[1].toUpperCase() : null;
}

// ケースを読む
const byFunc = {};       // 関数名 → { any: [id], bare: [id] }
for (const file of fs.readdirSync(CASEDIR)) {
  if (file === '_inputs.json') continue;
  const j = JSON.parse(fs.readFileSync(path.join(CASEDIR, file), 'utf8'));
  for (const c of j.cases) {
    const fn = String(c.func || '').toUpperCase();
    if (!fn || fn === '(入力)' || !isFunctionName(fn)) continue;   // 演算子・分類名は対象外
    byFunc[fn] = byFunc[fn] || { any: [], bare: [] };
    byFunc[fn].any.push(c.id);
    if (outermost(c.f) === fn) byFunc[fn].bare.push(c.id);
  }
}

console.log('\n[bare-form] 「客が最初に書く形」のケースがあるか');

const missing = Object.keys(byFunc).filter(fn => !byFunc[fn].bare.length && !EXCEPTIONS[fn]).sort();

T('★ケースを持っている関数は、素の形のケースも最低1本持っている', () => {
  if (missing.length) {
    throw new Error('素の形のケースが無い関数:\n   - '
      + missing.map(fn => fn + '（今あるのは包んだ形だけ: ' + byFunc[fn].any.slice(0, 3).join(', ') + '）').join('\n   - ')
      + '\n   → その関数を「一番外側に置いた式」のケースを足す。'
      + '\n     包んだ形しか無いと、素の形が壊れていても数字では気づけない（R19で実際に踏んだ）。');
  }
});

T('例外表の各項目に理由が書いてある', () => {
  for (const [fn, why] of Object.entries(EXCEPTIONS)) {
    if (!why || why.length < 10) throw new Error(fn + ': 理由が不十分');
  }
});

T('例外表に「もう使っていない関数」が残っていない', () => {
  const dead = Object.keys(EXCEPTIONS).filter(fn => !byFunc[fn]);
  if (dead.length) throw new Error('ケースが無いのに例外表に残っている: ' + dead.join(', ') + '（消すこと）');
});

T('検査が空振りしていない（関数を実際に数えている）', () => {
  if (Object.keys(byFunc).length < 40) throw new Error('拾えた関数が少なすぎます: ' + Object.keys(byFunc).length);
});

const withBare = Object.keys(byFunc).filter(fn => byFunc[fn].bare.length).length;
console.log('\n── 実測 ──');
console.log('  ケースを持つ関数: ' + Object.keys(byFunc).length + '個');
console.log('  うち素の形あり  : ' + withBare + '個');
console.log('  例外(理由つき)  : ' + Object.keys(EXCEPTIONS).length + '個 … ' + Object.keys(EXCEPTIONS).join(', '));
console.log('  素の形が無い    : ' + missing.length + '個' + (missing.length ? '\n   - ' + missing.join('\n   - ') : ''));

console.log('\n' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
