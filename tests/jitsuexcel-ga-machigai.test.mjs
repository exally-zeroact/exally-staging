/* jitsuexcel-ga-machigai.test.mjs — ★「実Excel の 方が 間違って いる」の 逃げ道を 塞ぐ★（2026-09-16）
 *
 *  ★★なぜ この 見張りが 在るか★★
 *    この 家の 決まりは ★「実Excel が 決める」★です。
 *    2026-09-16 に ★その 決まりから 外れる 初めての 1件★が 出ました。
 *      … `BESSELI` の 大きい x で ★実Excel の 方が 間違って います★
 *
 *  ★★一番 恐いのは 逃げ道です★★（★経営者1 が 条件に しました★）
 *    ＝★次に 合わない 物が 出た 時、免除を 1つ 増やすだけで 静かに 広がる★
 *    ⇒★★増やすなら ★赤を 踏んで 名前を 書く★ 形に する★★
 *    ⇒★人の 心がけでは 止まりません＝機械で 塞ぎます★
 *
 *  ★★この 見張りが 見る 物★★
 *    ①★免除は 名簿で 持つ★（★字の 中に 散らさない★）
 *    ②★★本数を 決め打ちに する★★（★1件★）… 増やしたら ★ここが 赤★
 *    ③★1件ずつ 訳（why）と 戻す 条件（back）と ★出所★ が 在る★
 *    ④★★3つの 別々の 道で 確かめた 手順が 紙に 残って いる★★
 *      ＝★数だけ 残ると 次の 人には「実Excel を 疑った 前例」だけが 残ります★
 *    ⑤★★他所（Rakunally など）の 根拠に しない★★
 *      ＝★Exally の BESSEL 1件だけ★と 紙に 書いて ある
 *
 *  使い方: node tests/jitsuexcel-ga-machigai.test.mjs
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
let pass = 0, fail = 0;
const T = (n, f) => {
  try { f(); pass++; console.log('  ✓ ' + n); }
  catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + e.message); }
};

console.log('\n[jitsuexcel-ga-machigai] ★「実Excel の 方が 間違って いる」の 逃げ道を 塞ぐ★');

/* ★★免除の 名簿★★（★ここ 以外に 書かない★）
     ★増やす 時は ★この 数も 直す★＝★直さないと 赤で 止まります★★ */
const 免除 = [
  {
    どれ: 'BESSELI（大きい x）',
    差: 5.14e-4, 幅: 1e-8, 桁: 17,
    why: '★実Excel の BESSELI は 大きい x で 3〜4桁しか 正しく ありません★。'
      + '`=BESSELI(101,2)` … 実Excel 2.848476711286702e+42 ／ 正しい値 2.847013943032507e+42'
      + '（★相対差 5.14e-4★）。'
      + 'また `=BESSELK(2,3)` の 実Excel の 答えは ★0.6473854＝7桁しか 在りません★'
      + '（他の 関数は 17桁 返ります）＝★実Excel 自身が 「ここは 精度が 低い」と 言って いるのと 同じ★。',
    back: '★実Excel の 答えが 変わったら 外す★（版が 上がって 直る 事が 在ります）',
    出所: '★3つの 別々の 道で 確かめました★ … '
      + '①級数（項の比・Exally1）②漸近形（Exally1）③★積分（経営者1 が 独立に 押した・2026-09-16）★',
    紙: 'docs/measured/bessel-no-kotae.md',
  },
  {
    どれ: 'BESSELK（大きい x）',
    差: 9.58e-8, 幅: 1e-8, 桁: 17,
    why: '★`=BESSELK(101,2)` … 実Excel 1.7385182551907827e-45 ／ 正しい値 1.73851808855756e-45★'
      + '（★相対差 9.58e-8★）。★積分の 道で 確かめました★'
      + '（N=20000/80000/320000 で 動かない）。',
    back: '★実Excel の 答えが 変わったら 外す★',
    出所: '★3つ目の 道（積分）で Exally1 が 押しました（2026-09-16）★'
      + ' K_ν(x) = ∫₀^∞ e^{−x cosh t} cosh(νt) dt',
    紙: 'docs/measured/bessel-no-kotae.md',
  },
  {
    どれ: 'BESSELY(2,3)',
    差: 1.04e-8, 幅: 1e-8, 桁: 17,
    why: '★`=BESSELY(2,3)` … 実Excel -1.1277837651220644 ／ 正しい値 -1.127783776725578★'
      + '（★相対差 1.04e-8★）。★積分の 道で 確かめました★。',
    back: '★実Excel の 答えが 変わったら 外す★',
    出所: '★3つ目の 道（積分）で Exally1 が 押しました（2026-09-16）★'
      + ' Y_n(x) = (1/π)∫₀^π sin(x sinθ − nθ)dθ − (1/π)∫₀^∞ [e^{nt}+(−1)^n e^{−nt}] e^{−x sinh t} dt',
    紙: 'docs/measured/bessel-no-kotae.md',
  },
];

/* ★★本数の 決め打ち★★ … ★増やすなら 赤を 踏んで 名前を 書く★ */
/* ★★2026-09-16 … 1 → ★3件★★
     ★門が 狙いどおり 効きました★
       ＝K・Y を 書いたら 合わない 物が 2本 出、
         ★この 門が 赤に なった★⇒★★3つ目の 道（積分）で 測ってから 名前を 書いた★★
     ★黙って 増やせない★＝これが 狙いです */
const 免除の本数 = 3;

console.log('      … 免除 ' + 免除.length + '件（決め打ち ' + 免除の本数 + '件）');

T('★★免除が 黙って 増えて いない（★逃げ道を 塞ぐ★）★★', () => {
  if (免除.length !== 免除の本数) {
    throw new Error('★免除が ' + 免除.length + '件★（' + 免除の本数 + '件の はず）'
      + '／★増やすなら ここも 直す＝★赤を 踏んで 名前を 書く★ 形に して あります★');
  }
});

T('★1件ずつ 訳（why）・戻す 条件（back）・★出所★ が 在る★', () => {
  const 欠 = [];
  for (const m of 免除) {
    for (const k of ['why', 'back', '出所', '紙']) {
      if (!m[k] || String(m[k]).length < 10) 欠.push(m.どれ + ' の ' + k);
    }
  }
  if (欠.length) throw new Error('★' + 欠.length + '件 欠けて いる★ … ' + 欠.join(' ／ '));
});

T('★★3つの 別々の 道の 手順が 紙に 残って いる★★', () => {
  /* ★数だけ 残ると 次の 人には「実Excel を 疑った 前例」だけが 残ります★ */
  for (const m of 免除) {
    const p = path.join(ROOT, m.紙);
    if (!fs.existsSync(p)) throw new Error('★紙が 無い★ … ' + m.紙);
    const s = fs.readFileSync(p, 'utf-8');
    for (const 要 of ['級数', '漸近', '積分', '2.847013943032507e+42', '5.14e-4', '0.6473854']) {
      if (s.indexOf(要) < 0) {
        throw new Error('★紙に「' + 要 + '」が 無い★（' + m.紙 + '）'
          + '／★手順が 残って いない 免除は 通しません★');
      }
    }
    /* ★出所を 名指しで★＝★次の 人が「うちが 間違って いる」と 読んで 直さない 為★ */
    if (s.indexOf('経営者1') < 0) {
      throw new Error('★紙に ★出所（誰が 独立に 押したか）★ が 書いて ない★（' + m.紙 + '）');
    }
  }
});

T('★★1件ずつ ★差の 大きさ★と ★合う幅の 何倍か★ が 在る★★', () => {
  /* ★★境目の 物を 見える ように します★★（経営者1 の 注文）
       ★なぜ★ 境目の 物が 増えると、ある日 誰かが
         「★合う幅を 1e-7 に すれば きれいに なる★」と 言います。
       ⇒★その時「それは ★免除を ◆件 消す 事だ★」と 見えて いないと 通ります★ */
  const 欠 = [];
  for (const m of 免除) {
    if (!(m.差 > 0) || !(m.幅 > 0) || !(m.桁 > 0)) 欠.push(m.どれ);
  }
  if (欠.length) throw new Error('★' + 欠.length + '件 差・幅・桁 が 欠けて いる★ … ' + 欠.join(' '));
  const はっきり = 免除.filter((m) => m.差 > m.幅 * 10);
  const 境目 = 免除.filter((m) => m.差 <= m.幅 * 10);
  console.log('      … ★はっきり 違う★ ' + はっきり.length + '件 ／ ★境目（幅の 10倍 以内）★ ' + 境目.length + '件');
  for (const m of 境目) {
    console.log('        ・' + m.どれ + ' … 差 ' + m.差.toExponential(2)
      + '＝幅の ' + (m.差 / m.幅).toFixed(1) + '倍 ／ ★幅を 1段 緩めると 消えます★');
  }
});

T('★★境目の 件数が 黙って 増えて いない★★', () => {
  /* ★境目の 物は ★幅を 1段 緩めるだけで 消えます★
       ⇒★★だから ここを 決め打ちに します★★ */
  const 境目の本数 = 2;   /* BESSELK（101,2）／BESSELY(2,3) */
  const 境目 = 免除.filter((m) => m.差 <= m.幅 * 10);
  if (境目.length !== 境目の本数) {
    throw new Error('★境目の 物が ' + 境目.length + '件★（' + 境目の本数 + '件の はず）'
      + '／★境目が 増えたら ★幅を 緩めたく なる★＝それが 一番 危ない★');
  }
});

T('★★他所の 根拠に しないと 紙に 書いて ある★★', () => {
  /* ★これは Exally の BESSEL 1件だけ★
       ＝★他所で 合わない 時に「実Excel が 間違って いる」を 持ち出さない★ */
  const p = path.join(ROOT, 免除[0].紙);
  const s = fs.readFileSync(p, 'utf-8');
  for (const 要 of ['1件だけ', '逃げ道']) {
    if (s.indexOf(要) < 0) {
      throw new Error('★紙に「' + 要 + '」の 断りが 無い★（' + 免除[0].紙 + '）');
    }
  }
});

console.log('\n' + pass + ' passed, ' + fail + ' failed');
if (fail) process.exit(1);
