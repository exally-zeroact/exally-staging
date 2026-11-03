/* tsunagu-to-ochiru-mono.mjs -- ★繋いだら ★落ちる 物★ を 数える★（2026-09-18）
 *
 *  ★★なぜ これが 要るか★★
 *    「台だけ 74.5% ／ 客の道 72.6%」は ★合計の 話★です。
 *    ★合計が 上でも ★今 合って いる 物が 落ちる★ 事が 在ります★。
 *    ⇒★★1本ずつ 突き合わせて「客の道だけ 合う」を 数えます★★
 *    ⇒★それが ★繋いだ 日に お客さんが 失う 物★です★
 *
 *  ★使い方★
 *    node docs/measured/osu-dai-dake.mjs   --出し=<どこか>/dai.tsv
 *    node docs/measured/osu-kami-webkit.mjs --出し=<どこか>/kyaku.tsv [--どこ=<URL>]
 *    node docs/measured/tsunagu-to-ochiru-mono.mjs <どこか>/dai.tsv <どこか>/kyaku.tsv
 *
 *  ★★この 台の 数は 画面の 数では ありません★★
 *    ＝`dai.tsv` は node の 台／`kyaku.tsv` は ★ブラウザで 押した 物★
 *    ⇒★左右で 出どころが 違います★（★それが 狙いです★）
 *
 *  ★見て いない 事★
 *    ・★同じ 式が 2回 出て 来たら 1本に します★（分母が 減ります）
 *    ・★溢れの 2つ目 以降★は 見て いません
 */
import fs from 'node:fs';

const 引数 = process.argv.slice(2);
if (引数.length < 2) {
  console.log('★使い方★ node docs/measured/tsunagu-to-ochiru-mono.mjs <dai.tsv> <kyaku.tsv>');
  process.exit(2);
}
const 読む = (p) => {
  const m = new Map();
  for (const l of fs.readFileSync(p, 'utf-8').split(/\r?\n/)) {
    const c = l.split('\t');
    if (c.length < 3) continue;
    /* ★4列目が 在れば ★測り道具の 限り★＝★台の 欠陥では ない★★
       ＝`osu-dai-dake.mjs` が 名指しで 許した 行（★名簿は あちら 1か所だけ★） */
    m.set(c[0], { 出: c[1], 正: c[2], 限り: !!c[3] });
  }
  return m;
};
const 台 = 読む(引数[0]);
const 客 = 読む(引数[1]);

const 数か = (x) => /^[-+]?[0-9]*[.]?[0-9]+([eE][-+]?[0-9]+)?$/.test(x);
const 合うか = (出, 正) => {
  if (出 === 正) return true;
  if (数か(出) && 数か(正)) {
    const a = Number(出); const b = Number(正);
    return Math.abs(a - b) <= Math.max(1e-9, Math.abs(b) * 1e-9);
  }
  return false;
};

let 両方 = 0; let 台だけ = 0; let 客だけ = 0; let どちらも違う = 0;
const 落ちる = [];
for (const [式, d] of 台.entries()) {
  const k = 客.get(式);
  if (!k) continue;
  /* ★測り道具の 限りで 落ちた 行は ★合った★ として 数えます★
     ＝★本物の 窓では 読めるから★（`xpath-name-honmono-de-osu.mjs` で 裏取り済） */
  const a = d.限り ? true : 合うか(d.出, d.正);
  const b = 合うか(k.出, k.正);
  if (a && b) 両方 += 1;
  else if (a) 台だけ += 1;
  else if (b) { 客だけ += 1; 落ちる.push({ 式, 台: d.出, 客: k.出, 正: d.正 }); }
  else どちらも違う += 1;
}

const 分母 = 両方 + 台だけ + 客だけ + どちらも違う;
console.log('');
console.log('[tsunagu-to-ochiru-mono] ★繋いだら 落ちる 物★');
console.log('  ★合わせた 分母 ... ' + 分母 + '本★（★同じ 式は 1本に して います★）');
console.log('  ★測り道具の 限りとして 許した 行 ... ' + [...台.values()].filter((x) => x.限り).length + '本★');
console.log('');
console.log('    両方 合う .................. ' + 両方);
console.log('    ★台だけ 合う★ ............. ' + 台だけ + '  ←★繋ぐと 増える 物★');
console.log('    ★★客の道だけ 合う★★ ...... ' + 客だけ + '  ←★★繋ぐと 落ちる 物★★');
console.log('    どちらも 違う .............. ' + どちらも違う);
console.log('');
const 名 = (f) => { const m = /^=([A-Z0-9_.]+)\(/.exec(f); return m ? m[1] : '(不明)'; };
const 束 = new Map();
for (const x of 落ちる) 束.set(名(x.式), (束.get(名(x.式)) || 0) + 1);
console.log('  ★★落ちる 物を 関数ごとに★★');
for (const [n, c] of [...束.entries()].sort((a, b) => b[1] - a[1])) {
  console.log('    ' + n.padEnd(14) + ' ... ' + c + '本');
}
console.log('');
console.log('  ★★落ちる 物 全部★★');
for (const x of 落ちる) {
  console.log('    ' + x.式.slice(0, 96));
  console.log('        台 ' + String(x.台).slice(0, 24)
    + ' ／客の道 ' + String(x.客).slice(0, 24) + ' ／実Excel ' + String(x.正).slice(0, 24));
}
