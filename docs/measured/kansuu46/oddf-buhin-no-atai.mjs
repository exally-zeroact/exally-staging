/* oddf-buhin-no-atai.mjs — ★ODDFPRICE の 組み立てに 入る 数を 全部 出す★（2026-09-16）
 *
 *  ★★なぜ 要るか★★
 *    ODDFPRICE は 30/48。★部品（日数の 数え方）は 210/210 合って います★。
 *    ⇒★残るのは 式の 組み立て★。
 *    ★但し 組み立てを 直す 前に、★今 何が 入って いるか★を 数で 出します★。
 *    ⇒★「たぶん ここ」で 触ると 棚63 の 繰り返しに なります★
 *      （2026-09-16 に 式を 当てて 6通り 試して ★全部 元より 悪く★ なった）
 *
 *  ★★この 道具は 数を 出すだけ★★
 *    ・★直しません★
 *    ・★実Excel を 叩きません★
 *    ⇒★軽い★
 *
 *  ★出す 数★（`lib/formula-kane.js` の `初回端数の価格` が 使って いる 物）
 *    DFC … 端数期間 まるごと（発行 → 初回利払）を 準期間の 数で 割った 物
 *    A   … 発行 → 決済
 *    DSC … DFC − A（＋「発行が 準利払日に 乗って いない」時の 補い 1）
 *    N   … 初回利払 ＋ その後 の 回数
 *
 *  ★★見る 所★★
 *    ★実Excel は basis 2 と 3 で 大きく 変わります★が
 *    ★うちは ほとんど 変わりません★（30/48 の 落ちが basis 2/3 に 偏る）
 *    ⇒★DFC / A / DSC の どれが 動いて いないかを 見ます★
 *
 *  使い方: node docs/measured/kansuu46/oddf-buhin-no-atai.mjs
 */
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '../../..');
const require_ = createRequire(path.join(ROOT, 'package.json'));
const K = require_(path.join(ROOT, 'lib/formula-kane.js'));

const 紙道 = path.join(ROOT, 'docs/measured/kansuu46/golden-kane-2026-09-07.tsv');
const 行 = fs.readFileSync(紙道, 'utf-8').replace(/^﻿/, '').split(/\r?\n/)
  .filter((l) => l && !l.startsWith('#')).map((l) => l.split('\t'))
  .filter((c) => c[0] === 'ODDFPRICE');

const 引数を取る = (式) => {
  const 中 = /^=ODDFPRICE\((.*)\)$/.exec(式.trim())[1];
  const 並 = []; let 深 = 0, 今 = '';
  for (const ch of 中) {
    if (ch === '(') { 深++; 今 += ch; }
    else if (ch === ')') { 深--; 今 += ch; }
    else if (ch === ',' && 深 === 0) { 並.push(今); 今 = ''; }
    else 今 += ch;
  }
  並.push(今);
  return 並;
};
const 日にする = (s) => {
  const m = /^DATE\((\d+),(\d+),(\d+)\)$/.exec(s.trim());
  return { y: +m[1], m: +m[2], d: +m[3] };
};

/* ★★`比` は 外に 出て いません★★
     ⇒★同じ 考えで ここに 書き写しません★（写すと ★2か所に なって ずれます★）
     ⇒★★代わりに 「価格が 利回りに どう 効くか」から DSC を 逆に 出します★★
       … `初回端数の価格` は 利回り y に対して
         `出 = 償還/割^((N-1)+DSC) + 券*DFC/割^DSC + Σ… − 券*A`
       ⇒★y を 2つ 入れて 比べると 中の 数が 動いたかが 分かります★
     ★但し それは 遠回りです★
     ⇒★★一番 素直なのは 本体に 「中の 数を 返す 口」を 付ける 事★★
       ＝★お客さんの 道は 変えません★（引数を 1つ 足すだけ・既定は 今のまま）
     ⇒★この 道具は その 口を 使います★ */

if (typeof K.初回端数の中身 !== 'function') {
  console.log('');
  console.log('★★`lib/formula-kane.js` に `初回端数の中身` が まだ 在りません★★');
  console.log('  ⇒★中の 数（DFC / A / DSC / N）を 外から 見る 口が 要ります★');
  console.log('  ⇒★お客さんの 道は 変えません★（別の 名前で 出すだけ）');
  process.exit(4);
}

console.log('');
console.log('★★ODDFPRICE の 組み立てに 入る 数★★（★直して いません／実Excel を 叩いて いません★）');
console.log('');
console.log('  組 basis  DFC                  A                    DSC                  N    ★合★');

const 束番号 = new Map();
let 番 = 0;
for (const c of 行) {
  const 式 = c[1];
  const 誤りか = /error/.test(c[3] || '') || /^#/.test(c[2]);
  const a = 引数を取る(式);
  const 決 = 日にする(a[0]), 満 = 日にする(a[1]), 発 = 日にする(a[2]), 初 = 日にする(a[3]);
  const f = Number(a[7]);
  const basis = a.length > 8 ? Number(a[8]) : 0;
  const 束 = a[0] + '|' + a[1] + '|' + a[2] + '|' + a[3];
  if (!束番号.has(束)) 束番号.set(束, ++番);

  const 数 = (dt) => K.日から数(dt.y, dt.m, dt.d);
  const m = K.初回端数の中身(数(決), 数(満), 数(発), 数(初), f, basis);
  const r = K.初回端数の価格(数(決), 数(満), 数(発), 数(初),
    Number(a[4]), Number(a[5]), Number(a[6]), f, basis);
  const u誤 = (r && r.誤り) ? ('#' + r.誤り + '!') : null;
  const 同 = 誤りか ? (u誤 === c[2].trim())
    : (!u誤 && Math.abs(Number(r) - Number(c[2])) <= Math.abs(Number(c[2])) * 1e-9);

  if (!m || m.誤り) {
    console.log('  ' + String(束番号.get(束)).padStart(2) + '  ' + String(basis).padEnd(5)
      + '  ★中身を 出せません★ ' + (u誤 || ''));
    continue;
  }
  console.log('  ' + String(束番号.get(束)).padStart(2)
    + '  ' + String(basis).padEnd(5)
    + '  ' + String(m.DFC).padEnd(20)
    + ' ' + String(m.A).padEnd(20)
    + ' ' + String(m.DSC).padEnd(20)
    + ' ' + String(m.N).padEnd(4)
    + ' ' + (同 ? '○' : '★×★'));
}

console.log('');
console.log('★見る 所★');
console.log('  ★実Excel は basis 2 と 3 で 大きく 変わります★');
console.log('  ★DFC / A / DSC が basis で 動いて いなければ そこが 元です★');
