/* oddf-wo-kami-de-osu.mjs -- ★ODDF を 紙 全部で 押す★（2026-09-19）
 *
 *  ★★なぜ 要るか★★
 *    Exally1 が ODDFPRICE を ★47 → 48 / 48★ に しました（★あちらの 分母★）。
 *    ⇒★出して よいかを 決める 前に ★私の 分母★で 数えます★
 *    ⇒★`oddl-wo-kami-de-osu.mjs` の ★双子★です★（★作り方は 写して います★）
 *
 *  ★★この 台の 数は 画面の 数では ありません★★
 *    ＝`lib/formula-kane.js` を ★直に★ 呼びます
 *    ＝★`保留の名前()` に ODDFPRICE / ODDFYIELD が 残って いる 間は
 *       ★皮からは #NAME? しか 出ません★★
 *    ⇒★画面の 事を 言いたいなら ブラウザで 押す★
 *
 *  ★見て いない 事★
 *    ・★ODDFYIELD は 紙に 在る 分だけ★
 *    ・★溢れ／書き出しの 道は 見て いません★
 *
 *  使い方: node docs/measured/oddf-wo-kami-de-osu.mjs
 */
import path from 'node:path';
import fs from 'node:fs';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';

const ここ = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(ここ, '..', '..');
const require_ = createRequire(path.join(ROOT, 'package.json'));
const K = require_(path.join(ROOT, 'lib/formula-kane.js'));

const 裸 = (s) => String(s === undefined ? '' : s).replace(/★/g, '').replace(/`/g, '').trim();
const 誤りの数 = {
  '-2146826281': '#DIV/0!', '-2146826252': '#NUM!', '-2146826246': '#N/A',
  '-2146826273': '#VALUE!', '-2146826265': '#REF!', '-2146826259': '#NAME?',
};
/* ★`osu-kami-webkit.mjs` と ★同じ 紙・同じ 列★★（★名簿を 2つ 持たない★） */
const 紙たち = [
  { 名: 'golden-kansuu-8kaime-2026-09-18.tsv', 式列: 1, 答列: 2 },
  { 名: 'golden-kansuu-9kaime-2026-09-18.tsv', 式列: 2, 答列: 4 },
  { 名: 'golden-kansuu-7kaime-2026-09-18.tsv', 式列: 1, 答列: 2 },
  { 名: 'golden-oddl-6kaime-2026-09-18.tsv', 式列: 1, 答列: 2 },
  { 名: 'golden-oddf-buhin-2026-09-16.tsv', 式列: 3, 答列: 4 },
  { 名: 'golden-oddf-to-46ko-2026-09-16.tsv', 式列: 2, 答列: 3 },
  { 名: 'golden-oddf-2kaime-2026-09-16.tsv', 式列: 1, 答列: 2 },
  { 名: 'golden-oddf-3kaime-2026-09-16.tsv', 式列: 1, 答列: 2 },
  { 名: 'golden-oddf-4kaime-2026-09-17.tsv', 式列: 1, 答列: 2 },
  { 名: 'golden-oddf-5kaime-2026-09-17.tsv', 式列: 1, 答列: 2 },
  { 名: 'kansuu46/golden-kane-2026-09-07.tsv', 式列: 1, 答列: 2 },
  { 名: 'golden-oddf-nokori-2026-09-18.tsv', 式列: 1, 答列: 2 },
  { 名: 'golden-oddf-nokori2-2026-09-18.tsv', 式列: 1, 答列: 2 },
  { 名: 'golden-oddf-nokori3-2026-09-18.tsv', 式列: 1, 答列: 2 },
  { 名: 'golden-oddf-nokori4-2026-09-18.tsv', 式列: 1, 答列: 2 },
  { 名: 'golden-oddf-nokori5-2026-09-18.tsv', 式列: 1, 答列: 2 },
  { 名: 'golden-oddf-nokori6-2026-09-18.tsv', 式列: 1, 答列: 2 },
  { 名: 'golden-oddf-nokori7-2026-09-18.tsv', 式列: 1, 答列: 2 },
  /* ★★下の 4枚は ★門を 足した その場で 出て きました★★（2026-09-19）
     ＝★私は 6枚 入れ忘れて いました★（15・16 と この 4枚） */
  { 名: 'golden-oddf-nokori8-2026-09-18.tsv', 式列: 1, 答列: 2 },
  { 名: 'golden-oddf-nokori9-2026-09-18.tsv', 式列: 1, 答列: 2 },
  { 名: 'golden-oddf-nokori10-2026-09-19.tsv', 式列: 1, 答列: 2 },
  { 名: 'golden-oddf-nokori11-2026-09-19.tsv', 式列: 1, 答列: 2 },
  /* ★2026-09-19 に 私が 取った 紙★ */
  { 名: 'golden-oddf-nokori12-2026-09-19.tsv', 式列: 1, 答列: 2 },
  { 名: 'golden-oddf-nokori13-2026-09-19.tsv', 式列: 1, 答列: 2 },
  { 名: 'golden-oddf-nokori14-2026-09-19.tsv', 式列: 1, 答列: 2 },
  { 名: 'golden-oddf-nokori15-2026-09-19.tsv', 式列: 1, 答列: 2 },
  { 名: 'golden-oddf-nokori16-2026-09-19.tsv', 式列: 1, 答列: 2 },
  { 名: 'golden-oddf-nokori17-2026-09-19.tsv', 式列: 1, 答列: 2 },
  { 名: 'golden-oddf-nokori18-2026-09-19.tsv', 式列: 1, 答列: 2 },
  { 名: 'golden-oddf-nokori19-2026-09-19.tsv', 式列: 1, 答列: 2 },
  { 名: 'golden-oddf-nokori20-2026-09-19.tsv', 式列: 1, 答列: 2 },
  { 名: 'golden-oddf-nokori21-2026-09-19.tsv', 式列: 1, 答列: 2 },
  { 名: 'golden-oddf-nokori22-2026-09-19.tsv', 式列: 1, 答列: 2 },
  { 名: 'golden-oddf-nokori23-2026-09-19.tsv', 式列: 1, 答列: 2 },
  { 名: 'golden-oddf-nokori24-2026-09-19.tsv', 式列: 1, 答列: 2 },
  { 名: 'golden-oddf-nokori25-2026-09-19.tsv', 式列: 1, 答列: 2 },
  { 名: 'golden-oddf-nokori26-2026-09-19.tsv', 式列: 1, 答列: 2 },
  { 名: 'golden-oddf-nokori27-2026-09-19.tsv', 式列: 1, 答列: 2 },
];

/* ══ ★★門 ── ★紙を 足したのに 名簿に 入れ忘れたら 止める★★ ══
   ★2026-09-19★ ... ㉔㉕で 取った `nokori15` `nokori16` を ★名簿に 入れ忘れました★
     ⇒★Exally1 が 「決 2009-05-30 も 外れて いる」と 教えて くれるまで 気づかず★
     ⇒★私の 274 / 277 は ★分母に 穴が 開いた ままの 数★でした★
   ★これで 今日 3人目です★（Exally1 の 48本／私の ODDL 4本／これ）
   ⇒★★覚書に 書くのでは なく 道具に 持たせます★★
   ⇒★`docs/measured` の `golden-oddf*.tsv` が 名簿に 1枚でも 無ければ ★exit 4★ */
{
  const 名簿 = new Set(紙たち.map((p) => p.名));
  const 落ち = fs.readdirSync(ここ)
    .filter((f) => /^golden-oddf.*\.tsv$/.test(f))
    .filter((f) => !名簿.has(f));
  if (落ち.length) {
    console.log('');
    console.log('★★紙が 在るのに 名簿に 入って いません ... ' + 落ち.length + '枚★★');
    for (const f of 落ち) console.log('    ' + f);
    console.log('  ⇒★★分母に 穴が 開いた まま 数えては いけません★★');
    console.log('  ⇒★`紙たち` に 足して ください（★式列 と 答列 を 確かめて★）');
    process.exit(4);
  }
}

const 数 = (y, m, d) => K.日から数(y, m, d);
/* ★basis を 書かない 形も 在ります★（★ODDL で 私の 分母に 4本の 穴が 開いて いました★）
   ＝`=ODDFPRICE(..., 100, 2)` の ように ★9つ目が 無い★
   ＝★実Excel では 省くと 0★（30/360 US） */
const 読む式 = new RegExp(
  '^=(ODDFPRICE|ODDFYIELD)\\(' +
  'DATE\\((\\d+),(\\d+),(\\d+)\\),\\s*DATE\\((\\d+),(\\d+),(\\d+)\\),\\s*' +
  'DATE\\((\\d+),(\\d+),(\\d+)\\),\\s*DATE\\((\\d+),(\\d+),(\\d+)\\),\\s*' +
  '([\\d.]+),\\s*([\\d.]+),\\s*([\\d.]+),\\s*(\\d+)(?:,\\s*(\\d+))?\\)$');

const 見た = new Set();
const 問い = [];
for (const p of 紙たち) {
  const 道 = path.join(ここ, p.名);
  if (!fs.existsSync(道)) continue;
  for (const l of fs.readFileSync(道, 'utf-8').replace(/^﻿/, '').split(/\r?\n/)) {
    if (!l || l.startsWith('#') || !l.includes('\t')) continue;
    const c = l.split('\t');
    const 式 = 裸(c[p.式列]).replace(/\s+/g, ' ');
    let 答 = 裸(c[p.答列]);
    const m = 読む式.exec(式);
    if (!m) continue;
    if (見た.has(式)) continue;
    見た.add(式);
    if (誤りの数[答]) 答 = 誤りの数[答];
    if (答 === '' || /打てません|受け付けません|HRESULT/.test(答)) continue;
    const n = m.slice(2).map((x) => (x === undefined ? 0 : Number(x)));   /* ★basis 省略は 0★ */
    問い.push({
      名: m[1], 式, 正: 答,
      決済: 数(n[0], n[1], n[2]), 満期: 数(n[3], n[4], n[5]),
      発行: 数(n[6], n[7], n[8]), 初回: 数(n[9], n[10], n[11]),
      利率: n[12], 二番: n[13], 償還: n[14], 頻度: n[15], basis: n[16],
    });
  }
}

console.log('');
console.log('[oddf-wo-kami-de-osu] ★ODDF を 紙 全部で 押す★');
console.log('  ★読んだ 紙 ... ' + 紙たち.filter((p) => fs.existsSync(path.join(ここ, p.名))).length
  + ' / ' + 紙たち.length + '枚★');
console.log('  ★★拾った 式 ... ' + 問い.length + '本★★（★同じ 式は 1回だけ★）');
const 種 = new Map();
for (const q of 問い) 種.set(q.名, (種.get(q.名) || 0) + 1);
for (const [n, c] of 種) console.log('    ' + n + ' ... ' + c + '本');

/* ══ ★★わざと 壊す 口★★（2026-09-19）══
   ★訳★ ... ★この 道具が 本当に 見て いるか★を 確かめる 為
     ＝`--わざと=1e-6` で ★台の 答えに 下駄を 履かせます★
     ＝★門が 効いて いれば 一気に 赤に なる はず★
   ⇒★★記憶「壊したのに 赤に ならない は まず 壊れて いるか」★★
   ⇒★★lib は 1字も 触りません★★（★Exally1 が 同じ 木で 総なめを 回して います★） */
const わざと = (() => {
  const a = process.argv.find((x) => x.startsWith('--わざと='));
  return a ? Number(a.slice('--わざと='.length)) : 0;
})();
if (わざと) {
  console.log('');
  console.log('  ★★わざと 壊して います ... 台の 答えに ' + わざと + ' を 足します★★');
}

const 数か = (x) => /^[-+]?[0-9]*[.]?[0-9]+([eE][-+]?[0-9]+)?$/.test(x);
let 合 = 0;
const 外れ = [];
for (const q of 問い) {
  let 出;
  try {
    出 = (q.名 === 'ODDFPRICE')
      ? K.初回端数の価格(q.決済, q.満期, q.発行, q.初回, q.利率, q.二番, q.償還, q.頻度, q.basis)
      : K.初回端数の利回り(q.決済, q.満期, q.発行, q.初回, q.利率, q.二番, q.償還, q.頻度, q.basis);
  } catch (e) { 出 = { 誤り: 'EX' }; }
  if (わざと && 出 !== null && typeof 出 === 'number') 出 += わざと;
  const 字 = (出 && 出.誤り) ? ('#' + 出.誤り + '!') : String(出);
  const 両方数 = 数か(q.正) && 数か(字);
  const 同 = (字 === q.正)
    || (両方数 && Math.abs(Number(字) - Number(q.正)) <= Math.max(1e-9, Math.abs(Number(q.正)) * 1e-9));
  if (同) { 合 += 1; continue; }
  外れ.push({ q, 字, 差: 両方数 ? (Number(字) - Number(q.正)) : null });
}

console.log('');
console.log('  ★★合った ' + 合 + ' / ' + 問い.length + '★★ ／ 外れ ' + 外れ.length + '本');

/* ★★外れを 種類で 分ける★★＝★数が 違う のか 誤りの 出方が 違う のか★ */
const 種類 = new Map();
for (const x of 外れ) {
  const う = String(x.字).startsWith('#');
  const 実 = String(x.q.正).startsWith('#');
  const t = (!う && 実) ? 'うちが数・実Excelが誤り'
    : (う && !実) ? 'うちが誤り・実Excelが数'
      : (う && 実) ? '誤りの種類が違う' : '数が違う';
  種類.set(t, (種類.get(t) || 0) + 1);
}
if (外れ.length) {
  console.log('');
  console.log('  ★★外れの 種類★★');
  for (const [t, c] of [...種類.entries()].sort((a, b) => b[1] - a[1])) {
    console.log('    ' + t + ' ... ' + c + '本');
  }
  console.log('');
  console.log('  ★★外れ 全部★★');
  for (const x of 外れ) {
    console.log('    ' + x.q.式.slice(0, 104));
    console.log('        うち ' + String(x.字).slice(0, 22) + ' ／実Excel ' + String(x.q.正).slice(0, 22)
      + (x.差 === null ? '' : ' ／差 ' + x.差.toExponential(3)));
  }
}

/* ★★この 数は 画面の 数では ありません★★（★皮は まだ 保留★） */
console.log('');
console.log('  ★★断り★★');
console.log('    ・★`lib/formula-kane.js` を 直に 呼んで います★');
console.log('    ・★`保留の名前()` に ODDFPRICE / ODDFYIELD が 在る 間は');
console.log('      ★皮（画面）からは #NAME? しか 出ません★');
console.log('    ・★画面の 数は `osu-kami-webkit.mjs` で 別に 測って ください★');
