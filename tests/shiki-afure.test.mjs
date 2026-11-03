/* shiki-afure.test.mjs — ★1つの 式が 何マスにも 広がる（溢れ）★（2026-09-13）
 *
 *  ★★土台の ⑤枚目★★
 *
 *  ★★下の 答えは 全部 実Excel に 打って 読んだ 物です★★
 *    （2026-09-13・★16通り★／道具 `docs/measured/toru-afure.ps1`／Excel 16.0 build 20326）
 *    ★1つも 当て推量が 在りません★
 *
 *  ★★材料★★ A1=1 A2=2 A3=3 ／ B1=10 B2=20 B3=30
 *
 *  ★★見て いない 範囲★★
 *    溢れた 先が また 溢れる／SORT・UNIQUE など 大きさが 中身で 決まる 関数／板を またぐ 四角
 */
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(path.join(ROOT, 'package.json'));
const K = require_(path.join(ROOT, 'lib/shiki-keisan.js'));
const A = require_(path.join(ROOT, 'lib/shiki-afure.js'));

let pass = 0, fail = 0;
const T = (n, f) => {
  try { f(); pass++; console.log('  ✓ ' + n); }
  catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + e.message); }
};
const 字に = (v) => {
  if (A.溢れか(v)) return v.並び.map((段) => 段.map(字に).join('|')).join(' / ');
  if (v.型 === '誤') return v.値;
  if (v.型 === '真偽') return v.値 ? 'TRUE' : 'FALSE';
  if (v.型 === '数') return K.数を字に(v.値);
  if (v.型 === '空') return '';
  return v.値;
};
const 同じ = (得, 欲, なぜ) => {
  if (得 !== 欲) throw new Error((なぜ ? なぜ + ' … ' : '') + 'うち `' + 得 + '` ／ ★実Excel `' + 欲 + '`★');
};

/* ★材料を 四角として 作る★（A1:A3 と B1:B3） */
const A列 = A.溢れ([[K.数(1)], [K.数(2)], [K.数(3)]]);
const B列 = A.溢れ([[K.数(10)], [K.数(20)], [K.数(30)]]);
const B短 = A.溢れ([[K.数(10)], [K.数(20)]]);
const AB = A.溢れ([[K.数(1), K.数(10)], [K.数(2), K.数(20)]]);

console.log('\n[shiki-afure] ★1つの 式が 何マスにも 広がる（溢れ）★');

T('★★=A1:A3 … 下に 広がる（1・2・3）★★', () => {
  同じ(字に(A列), '1 / 2 / 3');
});

T('★★=A1:A3*2 … 1つずつ 掛ける（2・4・6）★★', () => {
  const 出 = A.重ねる(A列, K.数(2), (a, b) => K.つなぎ('*', a, b, {}));
  同じ(字に(出), '2 / 4 / 6');
});

T('★★=A1:B2 … 横にも 広がる★★', () => {
  同じ(字に(AB), '1|10 / 2|20');
});

T('★★=A1:A3+B1:B3 … 1つずつ 足す（11・22・33）★★', () => {
  const 出 = A.重ねる(A列, B列, (a, b) => K.つなぎ('+', a, b, {}));
  同じ(字に(出), '11 / 22 / 33');
});

T('★★形が 違う 四角＝足りない 所だけ #N/A（全体を 誤りに しない）★★', () => {
  const 出 = A.重ねる(A列, B短, (a, b) => K.つなぎ('+', a, b, {}));
  同じ(字に(出), '11 / 22 / #N/A', '★=A1:A3+B1:B2★');
});

T('★★=A1:A3&"x" … つなぎも 1つずつ（1x・2x・3x）★★', () => {
  const 出 = A.重ねる(A列, K.字('x'), (a, b) => K.つなぎ('&', a, b, {}));
  同じ(字に(出), '1x / 2x / 3x');
});

T('★★=-A1:A3 … 前置きの − も 1つずつ（-1・-2・-3）★★', () => {
  const 出 = A.ひとつずつ(A列, (v) => K.前置き('-', v, {}));
  同じ(字に(出), '-1 / -2 / -3');
});

T('★★=A1:A3=1 … 大小くらべも 1つずつ（TRUE・FALSE・FALSE）★★', () => {
  const 出 = A.重ねる(A列, K.数(1), (a, b) => K.つなぎ('=', a, b, {}));
  同じ(字に(出), 'TRUE / FALSE / FALSE');
});

T('★★=@A1:A3 … @ を 付けると 溢れない（1に なる）★★', () => {
  同じ(字に(A.先頭(A列)), '1');
});

T('★★溢れる先に 物が 在れば #SPILL!（数・字・式・=""  全部 駄目）★★', () => {
  const 番地 = (i, j) => 'E' + (i + 1) + (j ? '?' : '');
  for (const [なに, 打った字] of [['数', '9'], ['字', 'あ'], ['空の字', '=""'], ['式', '=1+1']]) {
    const 表 = { E2: { 打った字 } };
    const r = A.溢れられるか(表, 'E1', A列, 番地);
    if (r.よい) throw new Error('★' + なに + '（`' + 打った字 + '`）が 在るのに 溢れようと して いる★');
  }
});

T('★★本当に 空の マスなら 溢れてよい★★', () => {
  const 番地 = (i, j) => 'E' + (i + 1) + (j ? '?' : '');
  const 表 = { E2: { 打った字: '' }, E3: {} };
  const r = A.溢れられるか(表, 'E1', A列, 番地);
  if (!r.よい) throw new Error('★空マスなのに 溢れられないと 言って いる … ' + (r.塞ぐ || []).join(',') + '★');
});

T('★★わざと 壊したら 赤に なるか（見張りが 効いて いるか）★★', () => {
  /* ★形が 違う 四角で「短い 方を 繰り返す」作りなら 11/22/31 に なる＝それを 名指しで 見る★ */
  const 出 = A.重ねる(A列, B短, (a, b) => K.つなぎ('+', a, b, {}));
  if (字に(出).indexOf('#N/A') < 0) throw new Error('★足りない 所を #N/A に して いない★');
});

console.log('');
console.log('★締め★ 通った ' + pass + ' ／ ★落ちた ' + fail + '★');
process.exit(fail ? 1 : 0);
