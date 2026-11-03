/* oddf-basho.test.mjs — ★新しい 数え方を ★どこに 当てるか★ を 守る★（2026-09-17）
 *
 *  ★★なぜ 要るか★★
 *    2026-09-17、実Excel を 測って ★新しい 数え方★が 出ました。
 *      ★1つの 端だけ 割り、まるごとの 期は 1と 数える★
 *    ★これを ★全部に 当てると 壊れます★★。当てる 場所は ★3つの うち 2つ★:
 *      DFC … ★当てる★（方式 'B1'）
 *      DSC … ★当てる★（方式 'B2'・★DFC − A を やめて 独立に★）
 *      A  … ★★当てない★★（方式 'B'・まとめて 割る）
 *
 *    ★私は 1度 全部に 当てて 1本 壊しました★。
 *    ★この 見張りが 無いと 次の 人が 同じ 事を します★
 *    （★「同じ 決まりなら 同じ 呼び方に すれば よい」と 見えるから★）
 *
 *  ★★紙★★
 *    docs/measured/kansuu46/golden-kane-2026-09-07.tsv（★壊れる 1本★）
 *    docs/measured/golden-oddf-5kaime-2026-09-17.tsv （DFC ★12/12★）
 *    docs/measured/golden-oddf-4kaime-2026-09-17.tsv （DSC ★17/17★）
 *
 *  使い方: node tests/oddf-basho.test.mjs
 */
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(path.join(ROOT, 'package.json'));
const K = require_(path.join(ROOT, 'lib/formula-kane.js'));

let pass = 0, fail = 0;
const T = (n, f) => {
  try { f(); pass++; console.log('  ✓ ' + n); }
  catch (e) { fail++; console.log('  ✗ ' + n + '\n      ' + e.message); }
};
const 近い = (出, 正, なに) => {
  if (!(Math.abs(出 - 正) <= Math.max(1e-9, Math.abs(正) * 1e-9))) {
    throw new Error('★' + なに + '★ 出た ' + 出 + ' ／ 実Excel ' + 正);
  }
};
const 数 = (y, m, d) => K.日から数(y, m, d);

console.log('\noddf-basho.test.mjs ★新しい 数え方を どこに 当てるか★\n');

/* ═══ ① ★A に 当てたら 壊れる 1本★ ═══════════════════════
     =ODDFPRICE(2009-07-10, 2013-01-01, 2008-07-01, 2010-01-01, 0.06, 0.05, 100, 2, 3)
     ★決済が 3つ目の 準期間に 在る 形★（他の 行は 1つ目か 2つ目）
     ★A を 'B1' に すると 2.0575（＝184/182.5 + 1 + 9/182.5）に なり ここが 落ちます★
     ★正しい A は 2.0493（＝(184+181+9)/182.5＝まとめて 割る）★ */
T('★A は まとめて 割る★（決済が 3つ目の 期に 在る 1本）', () => {
  const 出 = K.初回端数の価格(数(2009, 7, 10), 数(2013, 1, 1), 数(2008, 7, 1), 数(2010, 1, 1),
    0.06, 0.05, 100, 2, 3);
  近い(出, 103.01518348564699, 'ODDFPRICE basis 3・決済が 3つ目の 期');
});

T('★その 1本の A が 「まとめて 割る」値に なって いる★', () => {
  const r = K.初回端数の中身(数(2009, 7, 10), 数(2013, 1, 1), 数(2008, 7, 1), 数(2010, 1, 1), 2, 3);
  近い(r.A, (184 + 181 + 9) / (365 / 2), 'A（まとめて 割る）');
  /* ★★「最初だけ 割る」に なって いないか を 名指しで 弾く★★ */
  const だめ = 184 / (365 / 2) + 1 + 9 / (365 / 2);
  if (Math.abs(r.A - だめ) < 1e-9) {
    throw new Error('★A に 「まるごとは 1」を 当てて います★＝★当てる 場所が 違います★');
  }
});

/* ═══ ② ★DFC は 実Excel と 同じ（12/12）★ ══════════════════
     紙 golden-oddf-5kaime-2026-09-17.tsv から 解いた 値
     決 2009-03-01 ／ 満 2013-01-01 ／ f=2 */
const DFC組 = [
  ['端数1期', [2009, 1, 1], [2009, 7, 1], { 2: 1.0055555556, 3: 0.9917808219 }],
  ['端数2期', [2009, 1, 1], [2010, 1, 1], { 0: 2, 1: 2, 2: 2.0055555556, 3: 1.9917808219, 4: 2 }],
  ['端数3期', [2008, 7, 1], [2010, 1, 1], { 0: 3, 1: 3, 2: 3.0222222222, 3: 3.0082191781, 4: 3 }],
];
T('★DFC が 実Excel と 同じ★（★12本・分母を 出します★）', () => {
  let 数えた = 0;
  for (const [名, 発, 初, 期待] of DFC組) {
    for (const b of Object.keys(期待)) {
      const r = K.初回端数の中身(数(2009, 3, 1), 数(2013, 1, 1), 数(...発), 数(...初), 2, +b);
      if (typeof r.DFC !== 'number') throw new Error('★' + 名 + ' b' + b + ' の DFC が 数で ない★');
      近い(r.DFC, 期待[b], 'DFC ' + 名 + ' basis ' + b);
      数えた++;
    }
  }
  if (数えた !== 12) throw new Error('★12本 押す はずが ' + 数えた + '本★');
  console.log('      … ★12 / 12★');
});

/* ═══ ③ ★DSC は DFC − A では ない★ ════════════════════════
     ★端数3期 basis 2★ … DFC 3.02222 ／ A 1.35 ／ DFC−A ＝ 1.67222
     ★実Excel の DSC★ ＝ 1 ＋ 122/180 ＝ ★1.677778★（4枠目 17/17） */
T('★DSC は DFC − A では ない★（独立に 出して いる）', () => {
  const r = K.初回端数の中身(数(2009, 3, 1), 数(2013, 1, 1), 数(2008, 7, 1), 数(2010, 1, 1), 2, 2);
  近い(r.DSC補いなし, 1 + 122 / 180, 'DSC（端の 日数 ÷ 長さ ＋ まるごと）');
  if (Math.abs(r.DSC補いなし - (r.DFC - r.A)) < 1e-12) {
    throw new Error('★DSC が DFC − A に 戻って います★');
  }
});

/* ═══ ④ ★補い（DSC += 1）が 残って いる★ ════════════════════
     ★外すと 96通りの 合いが 70 → 58 に 落ちます★（実測） */
T('★補い（DSC += 1）が 残って いる★', () => {
  /* 発行が 準利払日に ★乗って いない★ ＋ 準期間 2つ以上 ⇒ 補い 1 */
  const あり = K.初回端数の中身(数(2008, 1, 15), 数(2012, 6, 30), 数(2007, 11, 20), 数(2008, 6, 30), 2, 0);
  if (あり.準の数 < 2 || あり.発が準日) throw new Error('★この 組は 補いの 形に なって いません★');
  if (あり.補い !== 1) throw new Error('★補いが 0 に なって います★（発行が 途中・準期間 2つ以上）');
  /* 発行が 準利払日に ★乗って いる★ ⇒ 補い 0 */
  const なし = K.初回端数の中身(数(2009, 3, 1), 数(2013, 1, 1), 数(2008, 7, 1), 数(2010, 1, 1), 2, 0);
  if (なし.補い !== 0) throw new Error('★発行が 準利払日なのに 補いが 付いて います★');
});

/* ═══ ⑤ ★対照（測り台が 変わって いない）★ ═══════════════════ */
T('★対照 3本★（紙と 同じ）', () => {
  /* ★紙 golden-oddf-5kaime-2026-09-17.tsv の 対照 3本を そのまま★ */
  近い(K.初回端数の価格(数(2008, 11, 11), 数(2021, 3, 1), 数(2008, 10, 15), 数(2009, 3, 1), 0.0785, 0.0625, 100, 2, 2),
    113.59879960832529, '対照1');
  近い(K.初回端数の価格(数(2009, 3, 1), 数(2013, 1, 1), 数(2009, 1, 1), 数(2010, 1, 1), 0.06, 0.05, 100, 2, 1),
    103.37232293583249, '対照2');
  近い(K.初回端数の価格(数(2009, 7, 1), 数(2013, 1, 1), 数(2008, 7, 1), 数(2010, 1, 1), 0.06, 0.05, 100, 2, 3),
    103.03027970006697, '対照3');
});

/* ═══ ⑥ ★紙 全部で 何本 合うか（★分母を 出す★）★ ════════════ */
T('★紙の 96通りで ★70本以上★ 合う★（★減ったら 赤★）', () => {
  const 紙 = path.join(ROOT, 'docs/measured/kansuu46/golden-kane-2026-09-07.tsv');
  const 行 = fs.readFileSync(紙, 'utf-8').replace(/^﻿/, '').split(/\r?\n/)
    .filter((l) => l && !l.startsWith('#')).map((l) => l.split('\t'))
    .filter((c) => /^ODD[FL](PRICE|YIELD)$/.test(c[0]));
  if (行.length !== 96) throw new Error('★紙の ODD 4個が ' + 行.length + '行★（96行の はず）');
  const 引数 = (式) => {
    const 中 = /^=ODD[FL](?:PRICE|YIELD)\((.*)\)$/.exec(式.trim())[1];
    const 並 = []; let 深 = 0, 今 = '';
    for (const ch of 中) {
      if (ch === '(') { 深++; 今 += ch; } else if (ch === ')') { 深--; 今 += ch; }
      else if (ch === ',' && 深 === 0) { 並.push(今); 今 = ''; } else 今 += ch;
    }
    並.push(今); return 並;
  };
  const 日 = (s) => { const m = /^DATE\((\d+),(\d+),(\d+)\)$/.exec(s.trim()); return m ? 数(+m[1], +m[2], +m[3]) : null; };
  let 合 = 0;
  for (const c of 行) {
    const 名 = c[0], a = 引数(c[1]), Fか = /^ODDF/.test(名);
    const f = Number(Fか ? a[7] : a[6]);
    const bs = Fか ? (a.length > 8 ? Number(a[8]) : 0) : (a.length > 7 ? Number(a[7]) : 0);
    let r;
    try {
      if (名 === 'ODDFPRICE') r = K.初回端数の価格(日(a[0]), 日(a[1]), 日(a[2]), 日(a[3]), +a[4], +a[5], +a[6], f, bs);
      else if (名 === 'ODDFYIELD') r = K.初回端数の利回り(日(a[0]), 日(a[1]), 日(a[2]), 日(a[3]), +a[4], +a[5], +a[6], f, bs);
      else if (名 === 'ODDLPRICE') r = K.最終端数の価格(日(a[0]), 日(a[1]), 日(a[2]), +a[3], +a[4], +a[5], f, bs);
      else r = K.最終端数の利回り(日(a[0]), 日(a[1]), 日(a[2]), +a[3], +a[4], +a[5], f, bs);
    } catch (e) { r = { 誤り: 'VALUE' }; }
    const 実誤 = /^#/.test(String(c[2]).trim());
    const u誤 = (r && r.誤り) ? ('#' + r.誤り + '!') : null;
    if (実誤 ? (u誤 === String(c[2]).trim())
      : (!u誤 && Math.abs(Number(r) - Number(c[2])) <= Math.max(1e-9, Math.abs(Number(c[2])) * 1e-9))) 合++;
  }
  console.log('      … ★' + 合 + ' / 96★（2026-09-17 の 直しで 57 → 70）');
  if (合 < 70) throw new Error('★96通りの 合いが ' + 合 + ' に 減りました★（70本の はず）');
});

console.log('\n' + pass + ' passed, ' + fail + ' failed');
if (fail) process.exit(1);
