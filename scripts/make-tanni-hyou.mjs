/* make-tanni-hyou.mjs — ★CONVERT の 単位表を ★実Excel の 答えから 機械で 作る★★（2026-09-07）
 *
 *  ★★手で 書かない 理由★★
 *    本番の CONVERT は ★1ポンド＝0.4536kg★ に なっていました（4桁で 丸めていた）。
 *    ⇒★実Excel は 0.45359237★
 *    ⇒★私が 覚えている 数を 書いたら 同じ 事を もう一度 やります★
 *    ⇒★★だから 実Excel に 打たせた 答え（golden-convert-…tsv）だけを 材料に します★★
 *
 *  ★作る 物★ … `lib/tanni-hyou.js`（★数だけ★・計算は lib/formula-tanni.js）
 *  ★確かめ方★ … `--check` で ★作り直した 物と 今 在る 物が 同じか★を 見る
 *                （stamp-build と 同じ やり方）
 *
 *  使い方: node scripts/make-tanni-hyou.mjs [--check]
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ここ = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(ここ, '..');
const 金の道 = path.join(ROOT, 'docs/measured/kansuu46/golden-convert-2026-09-07.tsv');
/* ★接頭辞が 付く 単位は 決めつけない★＝★1つずつ 実Excel に 聞いた★
   （実測 … ft・in・mi・yd・lbm・hr・day・BTU・yr・HP・kn・ha・ton には ★付かない★／
     面積は 1000の2乗・体積は 3乗に なる） */
const 頭可の道 = path.join(ROOT, 'docs/measured/kansuu46/golden-convert3-2026-09-07.tsv');
const 出の道 = path.join(ROOT, 'lib/tanni-hyou.js');

const 行 = fs.readFileSync(金の道, 'utf-8').split('\n').filter((l) => l && !l.startsWith('#'))
  .map((l) => l.split('\t'));

/* ★種類ごとの 基準★（式の 形から 読み取る＝手で 並べ直さない） */
const 係数 = {};      /* 単位 → { 種: 基準, 倍: 数 } */
const 頭 = {};        /* 接頭辞 → 倍 */
for (const [, 式, 答] of 行) {
  const m = 式.match(/^=CONVERT\(1,"([^"]*)","([^"]*)"\)$/);
  if (!m) continue;
  const [, u, b] = m;
  if (答.startsWith('#')) continue;
  const v = Number(答);
  if (!isFinite(v)) continue;
  /* 接頭辞の 行（"km"→"m" のように 基準が 素の 単位で 前が 付いている） */
  if (u !== b && u.endsWith(b) && u.length > b.length) {
    頭[u.slice(0, u.length - b.length)] = v;
    continue;
  }
  係数[u] = { 種: b, 倍: v };
}

/* ★接頭辞が 付くか／面積は 2乗・体積は 3乗か★（実測から） */
const 頭可 = {};
for (const l of fs.readFileSync(頭可の道, 'utf-8').split(/\r?\n/)) {
  if (!l || l.startsWith('#')) continue;
  const p2 = l.split('\t');
  const m = (p2[1] || '').match(/^=CONVERT\(1,"k([^"]+)","([^"]+)"\)$/);
  if (!m || (p2[2] || '').startsWith('#')) continue;
  const v = Number(p2[2]);
  頭可[m[1]] = (v === 1000) ? 1 : (v === 1e6 ? 2 : (v === 1e9 ? 3 : 1));
}

/* ★2進の 接頭辞は 情報の 単位にしか 付かない★（実測 … "Yim" は #N/A） */
const 二進 = {};
for (const [, 式, 答] of 行) {
  const m = 式.match(/^=CONVERT\(1,"([A-Za-z]+)i(bit|byte)","(bit|byte)"\)$/);
  if (!m || 答.startsWith('#')) continue;
  二進[m[1] + 'i'] = Number(答);
}
for (const k of Object.keys(二進)) delete 頭[k];

/* ★温度★＝掛け算では ない ⇒ 2点から 傾きと ずれを 出す */
const 温度 = {};
for (const [, 式, 答] of 行) {
  const m = 式.match(/^=CONVERT\((0|100),"C","([^"]+)"\)$/);
  if (!m || 答.startsWith('#')) continue;
  const t = m[2];
  温度[t] = 温度[t] || {};
  温度[t][m[1]] = Number(答);
}
const 温度表 = {};
for (const t of Object.keys(温度)) {
  const a = 温度[t]['0'], b = 温度[t]['100'];
  if (a === undefined || b === undefined) continue;
  温度表[t] = { 傾き: (b - a) / 100, ずれ: a, 頭: 頭可[t] || 0 };   /* t = 傾き × C + ずれ */
}

for (const u of Object.keys(係数)) 係数[u].頭 = 頭可[u] || 0;

const 数字 = (v) => {
  const s = String(v);
  return s.includes('e') || s.includes('E') ? s : s;
};
const 並べる = (o, f) => Object.keys(o).sort().map((k) => "    '" + k + "': " + f(o[k])).join(',\n');

const 本文 = `/* tanni-hyou.js — ★CONVERT の 単位表★（2026-09-07）
 *
 *  ★★この ファイルは 手で 書きません★★
 *    \`node scripts/make-tanni-hyou.mjs\` が
 *    ★実Excel に 打たせた 答え★（docs/measured/kansuu46/golden-convert-2026-09-07.tsv）
 *    から 作ります。
 *  ★直す 時★ … ★答えの 紙を 測り直して★ この 道具を もう一度 走らせる
 *  ★見張り★ … \`node scripts/make-tanni-hyou.mjs --check\`（CI で 走る）
 *
 *  ★なぜ こうしたか★
 *    本番の CONVERT は ★1ポンド＝0.4536kg★ に なっていました（4桁で 丸めていた）。
 *    ⇒★実Excel は 0.45359237★＝★手で 書いた 数は いつか 必ず 丸まる★
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.TanniHyou = factory();
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';
  /* 単位 → { 種: 同じ 種類の 基準, 倍: 基準に 直す 掛け算,
              頭: 接頭辞が 付くか（0＝付かない／1＝そのまま／2＝2乗／3＝3乗） } */
  var 単位 = {
${並べる(係数, (v) => "{ 種: '" + v.種 + "', 倍: " + 数字(v.倍) + ', 頭: ' + (v.頭 || 0) + ' }')}
  };
  /* 10進の 接頭辞（メートルに 付けて 測った） */
  var 接頭 = {
${並べる(頭, (v) => 数字(v))}
  };
  /* 2進の 接頭辞（★情報の 単位にしか 付かない★＝実測 "Yim" は #N/A） */
  var 二進接頭 = {
${並べる(二進, (v) => 数字(v))}
  };
  /* 温度（掛け算では ない）… その 単位 = 傾き × セ氏 + ずれ */
  var 温度 = {
${並べる(温度表, (v) => '{ 傾き: ' + 数字(v.傾き) + ', ずれ: ' + 数字(v.ずれ) + ', 頭: ' + (v.頭 || 0) + ' }')}
  };
  return { 単位: 単位, 接頭: 接頭, 二進接頭: 二進接頭, 温度: 温度 };
}));
`;

if (process.argv.includes('--check')) {
  const 今 = fs.existsSync(出の道) ? fs.readFileSync(出の道, 'utf-8') : '';
  if (今 !== 本文) {
    console.log('✗ lib/tanni-hyou.js が 答えの 紙と 合っていません（node scripts/make-tanni-hyou.mjs で 作り直す）');
    process.exit(1);
  }
  console.log('✓ lib/tanni-hyou.js は 答えの 紙から 作った ままです（単位 '
    + Object.keys(係数).length + '／接頭 ' + Object.keys(頭).length
    + '／2進 ' + Object.keys(二進).length + '／温度 ' + Object.keys(温度表).length + '）');
} else {
  fs.writeFileSync(出の道, 本文, { encoding: 'utf-8' });
  console.log('作った … lib/tanni-hyou.js（単位 ' + Object.keys(係数).length
    + '／接頭 ' + Object.keys(頭).length + '／2進 ' + Object.keys(二進).length
    + '／温度 ' + Object.keys(温度表).length + '）');
}
