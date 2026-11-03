/* shiki-kansuu-dsum.test.mjs — ★DSUM を 1個だけ 台へ 移す★（2026-09-15）
 *
 *  ★★この 紙は「作る前」に 書きました★★（指示役1 の 注文）
 *    ★「動いた」の 後に 書くと ★通る形に 合わせて★ しまいます★
 *
 *  ★★なぜ DSUM 1個か★★
 *    ★417個の うち 50個を「移すだけ」と 私は 書きました★
 *    ⇒ 指示役1「★それも まだ 測って いない★／★まず 1個 移して 手間を 出せ★」
 *    ⇒ ★この 紙が その 1個★です。
 *
 *  ★★物差しは 実Excel★★（★うちの 今の 答えでは ありません★）
 *    `docs/measured/kansuu46/golden-mada-2026-09-14.tsv` 73-74行
 *      `=DSUM(A1:B5,1,D1:D2)` → ★14★
 *      `=DSUM(A1:B5,1,A1:A2)` → ★2★
 *    ★取った Excel★ … 版 16.0 ／ build 20326（紙の 頭に 書いて ある）
 *
 *  ★★材料は 紙から★★（`docs/measured/kansuu46-no-dodai.mjs` と 同じ）
 *    A1:A5 = 1,2,3,4,5 ／ B1:B5 = 2,4,6,8,10
 *    D1 = `=DATE(2024,1,1)` ／ D2 = `=DATE(2026,1,1)`
 *    ★★但し 台は まだ `DATE` を 知りません★★（実測 … `#NAME?`）
 *    ⇒ ★D1/D2 には ★同じ 値★（45292 / 46023）を 直に 置きます★
 *      ＝★紙の 材料を 自分で 作り替えたのでは ありません★（★値は 同じ★）
 *      ＝★これも「見て いない 事」に 書きます★（下）
 *
 *  ★★見て いない 事★★
 *    ・★`DATE` を 通した 形は 見て いません★（値を 直に 置いた）
 *    ・★DSUM 以外の D系 11個は 見て いません★（★1個だけ★が この 紙の 決め）
 *    ・★今の 製品（JS層 `_jsDbFunc`）が 正しいかは ここでは 見ません★
 *      （★別に 測りました★ … `=DSUM(A1:B5,1,D1:D2)` は ★今の 製品 0／実Excel 14★＝★欠陥★）
 *
 *  使い方: node tests/shiki-kansuu-dsum.test.mjs
 *          node tests/shiki-kansuu-dsum.test.mjs --self-test
 */
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(path.join(ROOT, 'package.json'));
const H = require_(path.join(ROOT, 'lib/shiki-hyou.js'));

let pass = 0, fail = 0;
const T = (n, f) => {
  try { f(); pass++; console.log('  ✓ ' + n); }
  catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + e.message); }
};

console.log('\n[shiki-kansuu-dsum] ★DSUM を 1個だけ 台へ★');

/** ★紙と 同じ 材料を 台に 置く★（★式は H1★＝材料の 外／`#CYCLE` 避け） */
function 台を建てる() {
  const h = H.表();
  const A = [1, 2, 3, 4, 5], B = [2, 4, 6, 8, 10];
  for (let i = 0; i < 5; i++) { h.打つ('A' + (i + 1), String(A[i])); h.打つ('B' + (i + 1), String(B[i])); }
  h.打つ('D1', '45292');   /* ＝`=DATE(2024,1,1)` の 値 */
  h.打つ('D2', '46023');   /* ＝`=DATE(2026,1,1)` の 値 */
  return h;
}

T('★★① `=DSUM(A1:B5,1,D1:D2)` → 14★★（★実Excel★／★見出しが 合わない 条件は 効かない★）', () => {
  const h = 台を建てる();
  h.打つ('H1', '=DSUM(A1:B5,1,D1:D2)');
  const 出 = String(h.字('H1'));
  if (出 !== '14') throw new Error('H1 … ' + 出 + '（★14★ のはず）');
});

T('★★② `=DSUM(A1:B5,1,A1:A2)` → 2★★（★実Excel★／★見出しが 合う 条件は 効く★）', () => {
  const h = 台を建てる();
  h.打つ('H1', '=DSUM(A1:B5,1,A1:A2)');
  const 出 = String(h.字('H1'));
  if (出 !== '2') throw new Error('H1 … ' + 出 + '（★2★ のはず）');
});

T('★③ 2列目も 引ける★（★field の 番号が 効いて いるか★）', () => {
  /* ★★実Excel では 測って いません★★＝★①と 同じ 決めの 別の 列★を 台の 決めとして 縛る
     ＝B2:B5 = 4,6,8,10 ⇒ 28（★条件 D1:D2 は ①と 同じく 効かない★）
     ★最初 ここに「打ち直しても 同じ」と 書いて いました★
       ⇒★#NAME? と #NAME? を 比べて ★緑★に なって いました★（★意味の 無い 試験★） */
  const h = 台を建てる();
  h.打つ('H1', '=DSUM(A1:B5,2,D1:D2)');
  const 出 = String(h.字('H1'));
  if (出 !== '28') throw new Error('H1 … ' + 出 + '（★28★ のはず）');
});

T('★④ 材料を 打ち直すと 答えも 直る★（★頼りが 繋がって いるか★）', () => {
  const h = 台を建てる();
  h.打つ('H1', '=DSUM(A1:B5,1,D1:D2)');
  if (String(h.字('H1')) !== '14') throw new Error('先に 14 に ならない … ' + h.字('H1'));
  h.打つ('A3', '30');                       /* 3 → 30 ＝ 14 → 41 */
  if (String(h.字('H1')) !== '41') throw new Error('H1 … ' + h.字('H1') + '（★41★ のはず）'
    + '＝★材料を 直しても 追いかけて いません★');
});

T('★⑤ 誤りは 素通りする★（★どの 関数でも 同じ 決め★）', () => {
  const h = 台を建てる();
  h.打つ('A3', '=1/0');
  h.打つ('H1', '=DSUM(A1:B5,1,D1:D2)');
  const 出 = String(h.字('H1'));
  if (出 !== '#DIV/0!') throw new Error('H1 … ' + 出 + '（★#DIV/0!★ のはず）');
});

console.log('\n' + pass + ' passed, ' + fail + ' failed');

/* ★★自己試験★★＝★この 紙が「土台が 無い うちは 赤」である 事★を 見る
     ＝★★通る形に 合わせて 書いて いない 証し★★ */
if (process.argv.includes('--self-test')) {
  const h = H.表();
  const 知らない = String(h.打つ('Z1', '=DSUM(A1:B5,1,A1:A2)') || h.字('Z1')) === '#NAME?';
  console.log('\n[self-test] ★台が DSUM を 知らない 間★ … ' + (知らない ? '★#NAME?★' : '★知って いる★'));
  if (知らない && fail < 5) {
    console.log('★★おかしい★★ 土台が 無いのに 赤が ' + fail + '本 しか ありません（5本 のはず）');
    process.exit(1);
  }
  process.exit(0);
}
process.exit(fail ? 1 : 0);
