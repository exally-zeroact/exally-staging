/* grid-refedit.test.mjs — ★書き間違えた式を直せるか★
 *
 * なぜ必要か（2026-08-06・司さんが実機で踏んだ）:
 *   A2 に「=B1+30」と打ってしまった。B1 を A1 に直したくて A1 を触ったら
 *   ★「=B1+30A1」になった★。触るたびに悪くなり、B1 を直す手段が無かった。
 *   ＝式を1文字でも書き間違えたら、そのセルは捨てるしかない状態だった。
 *
 * ここで固定すること:
 *   ① カーソルが参照の上／端 → ★その参照だけを置き換える★（B1 が A1 になる）
 *   ② カーソルの直前が = + - * / ( , : → そこに差し込む（今までの組み立て方は壊さない）
 *   ③ どちらでもない（数字の直後など）→ ★1文字も変えない★（勝手に壊さない）
 *   ④ "文字列" の中の A1 は参照ではない
 *   ⑤ 関数名（LOG10( など）は参照ではない
 *
 * 使い方: node tests/grid-refedit.test.mjs
 *         node tests/grid-refedit.test.mjs --self-test
 */
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath, pathToFileURL } from 'node:url';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(pathToFileURL(path.join(ROOT, 'package.json')));
const R = require_(path.join(ROOT, 'lib/grid-refedit.js'));

let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ✓ ' + n); } catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + (e && e.message)); } };
const eq = (a, b, m) => { if (a !== b) throw new Error((m ? m + ': ' : '') + 'expected ' + JSON.stringify(b) + ' got ' + JSON.stringify(a)); };

console.log('grid-refedit（書き間違えた式を直せるか）');

/* ───────── ★司さんが踏んだ場面そのもの★ ───────── */
T('★=B1+30 の B1 にカーソルを置いて A1 を選ぶ → =A1+30★（直せる）', () => {
  // "=B1+30" の B1 は index 1..3。数式バーで B1 を触った状態＝カーソル2
  const r = R.refEditAt('=B1+30', 2, 'A1');
  eq(r.ok, true, 'ok');
  eq(r.v, '=A1+30', '式');
  eq(r.why, 'replace', 'やったこと');
  eq(r.pos, 3, 'カーソル');
});

T('★=B1+30 の末尾（数字の直後）で A1 を選ぶ → 1文字も変わらない★', () => {
  const r = R.refEditAt('=B1+30', 6, 'A1');
  eq(r.ok, false, 'ok');
  eq(r.v, '=B1+30', '式が壊れていない');   // ← 以前は "=B1+30A1" になっていた
  eq(r.why, 'not-a-ref-position', '理由');
});

T('参照の右端に接していても置き換わる（=B1 の末尾）', () => {
  const r = R.refEditAt('=B1', 3, 'A1');
  eq(r.v, '=A1');
  eq(r.why, 'replace');
});

/* ───────── 今までの組み立て方を壊していないか ───────── */
T('= だけ → セルを選ぶと差し込む', () => {
  const r = R.refEditAt('=', 1, 'A1');
  eq(r.v, '=A1'); eq(r.why, 'insert');
});
T('=A1+ → 次のセルを差し込む', () => {
  const r = R.refEditAt('=A1+', 4, 'B2');
  eq(r.v, '=A1+B2'); eq(r.why, 'insert');
});
T('=SUM( → 差し込む', () => {
  const r = R.refEditAt('=SUM(', 5, 'A1');
  eq(r.v, '=SUM(A1'); eq(r.why, 'insert');
});
T('=SUM(A1: → 範囲の続きを差し込む', () => {
  const r = R.refEditAt('=SUM(A1:', 8, 'B9');
  eq(r.v, '=SUM(A1:B9'); eq(r.why, 'insert');
});
T('範囲(A1:B2)の上にカーソル → 範囲ごと置き換える', () => {
  const r = R.refEditAt('=SUM(A1:B2)', 7, 'C3:D4');
  eq(r.v, '=SUM(C3:D4)'); eq(r.why, 'replace');
});
T('真ん中の参照だけ置き換わる（両隣を壊さない）', () => {
  const r = R.refEditAt('=A1+B2+C3', 5, 'Z9');   // B2 の上
  eq(r.v, '=A1+Z9+C3');
});
T('$付き（$A$1）も置き換えられる', () => {
  const r = R.refEditAt('=$A$1+1', 3, 'B2');
  eq(r.v, '=B2+1');
});

/* ───────── 参照でない物を参照と間違えない ───────── */
T('"文字列" の中の B1 は参照ではない（数字の直後扱いで何もしない）', () => {
  const r = R.refEditAt('=IF(A9>0,"B1","")', 11, 'A1');   // "B1" の中
  eq(r.ok, false);
  eq(r.v, '=IF(A9>0,"B1","")');
});
T('関数名 LOG10( は参照ではない', () => {
  const refs = R.findRefs('=LOG10(A1)');
  eq(refs.length, 1, '拾った数');
  eq(refs[0].text, 'A1');
});
T('数字だけ（30）は参照ではない', () => {
  eq(R.findRefs('=1+30').length, 0);
});

/* ───────── 端 ───────── */
T('空の入力 → 差し込む', () => { eq(R.refEditAt('', 0, 'A1').v, 'A1'); });
T('カーソルが範囲外 → 末尾として扱う（落ちない）', () => {
  const r = R.refEditAt('=A1+', 999, 'B2');
  eq(r.v, '=A1+B2');
});
T('選ぶセルが空 → 何もしない', () => {
  const r = R.refEditAt('=A1+', 4, '');
  eq(r.ok, false); eq(r.v, '=A1+');
});

/* ───────── わざと壊して赤になるか（--self-test） ───────── */
if (process.argv.includes('--self-test')) {
  console.log('\n  ── わざと壊して赤になるか ──');
  const old = R.refEditAt;
  // 直す前の実装（末尾が参照なら置換／でなければ末尾に追加）を再現する
  const broken = (v, pos, addr) => {
    const re = /(\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?)$/;
    const nv = re.test(v) ? v.replace(re, addr) : v + addr;
    return { v: nv, pos: nv.length, ok: true, why: 'legacy' };
  };
  let caught = 0;
  const cases = [
    ['=B1+30 の B1 を直せる', () => { if (broken('=B1+30', 2, 'A1').v !== '=A1+30') throw 0; }],
    ['数字の直後で壊さない', () => { if (broken('=B1+30', 6, 'A1').v !== '=B1+30') throw 0; }],
    ['真ん中だけ置き換える', () => { if (broken('=A1+B2+C3', 5, 'Z9').v !== '=A1+Z9+C3') throw 0; }],
  ];
  for (const [n, fn] of cases) {
    try { fn(); console.log('  ✗ ' + n + ' — ★直す前の実装でも通ってしまう＝この検査は空振り★'); fail++; }
    catch { caught++; console.log('  ✓ ' + n + ' — 直す前の実装なら赤になる'); }
  }
  if (caught !== cases.length) fail++;
  void old;
}

console.log(`\n  ${pass} 通過 / ${fail} 失敗`);
process.exit(fail ? 1 : 0);
