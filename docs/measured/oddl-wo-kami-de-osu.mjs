/* oddl-wo-kami-de-osu.mjs -- ★ODDL を 紙 全部で 押す★（2026-09-18）
 *
 *  ★★なぜ 要るか★★
 *    Exally1 が ODDL を ★6/12 → 12/12★ に しました（★あちらの 分母★）。
 *    ⇒★出して よいかを 決める 前に ★私の 分母★で 数えます★
 *    ⇒★私の 紙は 16枚／ODDLPRICE と ODDLYIELD が ★199本 以上★ 在ります★
 *    ⇒★★「12/12」は あちらの 48本の 話です★★
 *
 *  ★★この 台の 数は 画面の 数では ありません★★
 *    ＝`lib/formula-kane.js` を ★直に★ 呼びます
 *    ＝★皮（`shiki-tsunagi`）にも `book.html` にも まだ 載って いません★
 *    ⇒★画面の 事を 言いたいなら ブラウザで 押す★
 *
 *  ★見て いない 事★
 *    ・★ODDLYIELD は 紙に 在る 分だけ★
 *    ・★溢れ／書き出しの 道は 見て いません★
 *
 *  使い方: node docs/measured/oddl-wo-kami-de-osu.mjs
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
  /* ★今日 私が 取った 紙★ */
  { 名: 'golden-oddf-nokori-2026-09-18.tsv', 式列: 1, 答列: 2 },
  { 名: 'golden-oddf-nokori2-2026-09-18.tsv', 式列: 1, 答列: 2 },
  { 名: 'golden-oddf-nokori3-2026-09-18.tsv', 式列: 1, 答列: 2 },
  { 名: 'golden-oddf-nokori4-2026-09-18.tsv', 式列: 1, 答列: 2 },
  { 名: 'golden-oddf-nokori5-2026-09-18.tsv', 式列: 1, 答列: 2 },
  { 名: 'golden-oddf-nokori6-2026-09-18.tsv', 式列: 1, 答列: 2 },
  { 名: 'golden-oddf-nokori7-2026-09-18.tsv', 式列: 1, 答列: 2 },
];

const 数 = (y, m, d) => K.日から数(y, m, d);
/* ★★basis を 書かない 形も 在ります★★（2026-09-19 に 気づきました）
   ＝`=ODDLPRICE(..., 100, 2)` の ように ★8つ目が 無い★
   ＝★実Excel では 省くと 0★（30/360 US）
   ⇒★前は この 4本を 押して いませんでした★＝★分母が 271 では なく 275★
   ⇒★★自分の 分母に 穴が 開いて いました★★（★人に 言って いた のと 同じ 型★） */
const 読む式 = /^=(ODDLPRICE|ODDLYIELD)\(DATE\((\d+),(\d+),(\d+)\),\s*DATE\((\d+),(\d+),(\d+)\),\s*DATE\((\d+),(\d+),(\d+)\),\s*([\d.]+),\s*([\d.]+),\s*([\d.]+),\s*(\d+)(?:,\s*(\d+))?\)$/;

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
    const n = m.slice(2).map((x) => (x === undefined ? 0 : Number(x)));   /* ★basis を 書かない 時は 0★ */
    問い.push({
      名: m[1], 式, 正: 答,
      決済: 数(n[0], n[1], n[2]), 満期: 数(n[3], n[4], n[5]), 最終: 数(n[6], n[7], n[8]),
      利率: n[9], 二番: n[10], 償還: n[11], 頻度: n[12], basis: n[13],
    });
  }
}

console.log('');
console.log('[oddl-wo-kami-de-osu] ★ODDL を 紙 全部で 押す★');
console.log('  ★★拾った 式 ... ' + 問い.length + '本★★（★同じ 式は 1回だけ★）');
const 種 = new Map();
for (const q of 問い) 種.set(q.名, (種.get(q.名) || 0) + 1);
for (const [n, c] of 種) console.log('    ' + n + ' ... ' + c + '本');

const 数か = (x) => /^[-+]?[0-9]*[.]?[0-9]+([eE][-+]?[0-9]+)?$/.test(x);
let 合 = 0;
const 外れ = [];
for (const q of 問い) {
  let 出;
  try {
    出 = (q.名 === 'ODDLPRICE')
      ? K.最終端数の価格(q.決済, q.満期, q.最終, q.利率, q.二番, q.償還, q.頻度, q.basis)
      : K.最終端数の利回り(q.決済, q.満期, q.最終, q.利率, q.二番, q.償還, q.頻度, q.basis);
  } catch (e) { 出 = { 誤り: 'EX' }; }
  const 字 = (出 && 出.誤り) ? ('#' + 出.誤り + '!') : String(出);
  const 両方数 = 数か(q.正) && 数か(字);
  const 同 = (字 === q.正) || (両方数 && Math.abs(Number(字) - Number(q.正)) <= Math.max(1e-9, Math.abs(Number(q.正)) * 1e-9));
  if (同) { 合 += 1; continue; }
  外れ.push({ q, 字, 差: 両方数 ? (Number(字) - Number(q.正)) : null });
}

console.log('');
console.log('  ★★合った ' + 合 + ' / ' + 問い.length + '★★ ／ 外れ ' + 外れ.length + '本');

/* ★★外れを 種類で 分ける★★＝★数が 違う のか 誤りの 出方が 違う のか★ */
const 種類 = new Map();
const 決済が最終 = [];
for (const x of 外れ) {
  const う = String(x.字).startsWith('#');
  const 実 = String(x.q.正).startsWith('#');
  const t = (!う && 実) ? 'うちが数・実Excelが誤り'
    : (う && !実) ? 'うちが誤り・実Excelが数'
      : (う && 実) ? '誤りの種類が違う' : '数が違う';
  種類.set(t, (種類.get(t) || 0) + 1);
  if (x.q.決済 === x.q.最終) 決済が最終.push(x);
}
console.log('');
console.log('  ★★外れの 種類★★');
for (const [t, c] of [...種類.entries()].sort((a, b) => b[1] - a[1])) {
  console.log('    ' + t + ' ... ' + c + '本');
}
console.log('  ★そのうち ★決済 と 最終利払日が 同じ★ ... ' + 決済が最終.length + '本★');
if (外れ.length) {
  console.log('');
  console.log('  ★★外れ 全部★★');
  for (const x of 外れ) {
    console.log('    ' + x.q.式.slice(0, 96));
    console.log('        うち ' + String(x.字).slice(0, 22) + ' ／実Excel ' + String(x.q.正).slice(0, 22)
      + (x.差 === null ? '' : ' ／差 ' + x.差.toExponential(3)));
  }
}
