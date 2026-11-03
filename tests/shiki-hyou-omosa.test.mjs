/* shiki-hyou-omosa.test.mjs — ★表が 大きく なっても 重く なりすぎないか★（2026-09-14）
 *
 *  ★★なぜ 要るか★★
 *    2026-09-14 に 測ったら ★N の 2乗★に なって いました：
 *      `=A1+1` を N マスに 打つ … N=2000 33ms ／ 4000 125ms ／ ★8000 512ms★
 *      `=A1:A3`（溢れ）………………… ★8000 で 2,103ms★
 *    ★元★＝`見られている[A1]` が ★並び★で、打つ たびに `indexOf` で 頭から なめて いた。
 *      ⇒ 皆が 同じ マスを 見ると 並びが N まで 伸び ★N×N★。
 *    ★直した後★ … N=8000 で ★27ms★（溢れ ★52ms★）＝★N に 比例★
 *
 *  ★★時間そのものでは なく「N に どう 効くか」を 見ます★★
 *    ＝機械の 速さは 所に よって 違うので、★1本あたりの 時間が どれだけ 増えるか★で 判じる。
 *      N を 4倍に して ★1本あたりが 4倍を 超えたら 赤★（2乗なら 4倍に なる）。
 *      直した後の 実測は ★0.8倍★（増えて いない）＝★4倍は かなり 甘い 線★です。
 *
 *  ★★見て いない 範囲★★
 *    ・読み出し（`字`／`値`）の 重さ ・大きい 溢れ ・板を またぐ 参照
 *    ・実物の ブックを 開く 時の 重さ（★これは まだ 測って いません★）
 *
 *  使い方: node tests/shiki-hyou-omosa.test.mjs
 *          node tests/shiki-hyou-omosa.test.mjs --self-test
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

const 列字 = (n) => { let s = ''; while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = (n - r - 1) / 26; } return s; };

function 一回(N, 式) {
  const h = H.表();
  h.打つ('A1', '1'); h.打つ('A2', '2'); h.打つ('A3', '3');
  const t0 = process.hrtime.bigint();
  for (let i = 1; i <= N; i++) h.打つ(列字(3 + (i - 1) * 2) + '1', 式);
  return Number(process.hrtime.bigint() - t0) / 1e6;
}
/* ★一番 速かった 回を 採る★＝他の 仕事の 割り込みを 避ける */
function 測る(N, 式) {
  一回(300, 式);
  let m = Infinity;
  for (let k = 0; k < 3; k++) m = Math.min(m, 一回(N, 式));
  return m;
}

/* ★N を 4倍に して 1本あたりが 何倍に なるか★（2乗なら 4倍） */
const 伸び = (式) => {
  const 小 = 測る(1000, 式) / 1000;
  const 大 = 測る(4000, 式) / 4000;
  return { 小, 大, 倍: 大 / 小 };
};
const 許す倍 = 4;

console.log('\n[shiki-hyou-omosa] ★表が 大きく なっても 重く なりすぎないか★');
console.log('  ★N を 4倍に した 時、1本あたりが ' + 許す倍 + '倍を 超えたら 赤★（2乗なら 4倍）');

for (const [名, 式] of [['ただの 式  =A1+1', '=A1+1'], ['溢れる 式  =A1:A3', '=A1:A3']]) {
  T(名, () => {
    const r = 伸び(式);
    console.log('      N=1000 … ' + (r.小 * 1000).toFixed(1) + ' マイクロ秒/本'
      + ' ／ N=4000 … ' + (r.大 * 1000).toFixed(1) + ' マイクロ秒/本'
      + ' ／ ★' + r.倍.toFixed(2) + '倍★');
    if (!(r.倍 <= 許す倍)) {
      throw new Error('★1本あたりが ' + r.倍.toFixed(2) + '倍に 増えた（' + 許す倍 + '倍まで）★'
        + '＝★N の 2乗に なって いないか 見る事★');
    }
  });
}

T('★見られている は 名札の 束★（並びに 戻すと また N の 2乗に なる）', () => {
  const h = H.表();
  h.打つ('A1', '1');
  h.打つ('B1', '=A1+1');
  h.打つ('C1', '=A1+2');
  const 者 = h.見られている('A1');
  if (!Array.isArray(者)) throw new Error('★外に 出す 形は 並びの まま★');
  /* ★名札は 板つき★（㋖・2026-09-15）＝★道連れは 板を またぐ★ので 板が 要る */
  if (者.slice().sort().join(',') !== '板1!B1,板1!C1') throw new Error('中身が 違う … ' + 者.join(','));
  /* 同じ 式を 打ち直しても 増えない（重複が 入らない） */
  h.打つ('B1', '=A1+9');
  if (h.見られている('A1').length !== 2) throw new Error('★打ち直しで 増えた★ … ' + h.見られている('A1').length);
});

T('★見なく なったら 外れる★（消し忘れると 要らない 計算が 増える）', () => {
  const h = H.表();
  h.打つ('A1', '1');
  h.打つ('B1', '=A1+1');
  h.打つ('B1', '=2+2');                     /* もう A1 を 見て いない */
  if (h.見られている('A1').length !== 0) throw new Error('★外れて いない★ … ' + h.見られている('A1').join(','));
});

if (process.argv.includes('--self-test')) {
  console.log('\n[shiki-hyou-omosa --self-test] ★測りが 空振りして いないか★');
  T('★本当に N 本 打てて いる★（0本で 緑に しない）', () => {
    const h = H.表();
    h.打つ('A1', '1');
    for (let i = 1; i <= 50; i++) h.打つ(列字(3 + (i - 1) * 2) + '1', '=A1+1');
    if (h.字(列字(3 + 49 * 2) + '1') !== '2') throw new Error('★最後の マスが 計算されて いない★');
    if (h.見られている('A1').length !== 50) throw new Error('★50本 打てて いない★ … ' + h.見られている('A1').length);
  });
  T('★測りが 0ミリ秒を 返して いない★（速すぎて 判じられない を 見つける）', () => {
    const ms = 測る(1000, '=A1+1');
    if (!(ms > 0)) throw new Error('★0ミリ秒＝測れて いない★');
  });
}

console.log('\n' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
