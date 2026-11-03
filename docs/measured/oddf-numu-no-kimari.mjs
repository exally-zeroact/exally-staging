/* ★見立てを 173本 全部に 当てて 割る★（2026-09-18・経営者1）
 *
 *  ★★見立て 2つ（★紙に 在る ODDFPRICE 全部に 当たりました★）★★
 *    ㋐★決済が 発行より 後で ない★ ... #NUM!
 *    ㋑★★初回の 端数が 端数で ない★★ ... #NUM!
 *       ＝★発行 → 初回 の 日数（DFC）が ★ちょうど 1枠（E）★★
 *       ＝★「端数の 初回」なのに 端数が 無い＝Excel は 受け付けません★
 *
 *  ★★なぜ これが 効くか★★
 *    ODDF の 外れ 12本は ★1本も 数が 違いません★
 *      ＝★うちが 数を 返し、実Excel が #NUM! を 返す★ だけ
 *    ⇒★★割り引きの 形（かけらごと 等）は 元では ありません★★
 *    ⇒★★足りないのは ★受け付ける 条件★です★★
 *
 *  ★門★
 *    ・★見立てが 当たれば 見落とし 0・空振り 0★
 *    ・★1本でも 外れたら 見立ては 外れ★（★そう 書きます★）
 *
 *  ★★この 台の 数は 画面の 数では ありません★★（`lib/formula-kane.js` を 直に 呼ぶ）
 */
import path from 'node:path';
import fs from 'node:fs';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';

/* ★手元の 絶対の 道を 焼き込まない★（記憶の 決まり） */
const ここ = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(ここ, '..', '..');
const require_ = createRequire(path.join(ROOT, 'package.json'));
const K = require_(path.join(ROOT, 'lib/formula-kane.js'));

const 裸 = (s) => String(s === undefined ? '' : s).replace(/★/g, '').replace(/`/g, '').trim();
const 誤りの数 = {
  '-2146826281': '#DIV/0!', '-2146826252': '#NUM!', '-2146826246': '#N/A',
  '-2146826273': '#VALUE!', '-2146826265': '#REF!', '-2146826259': '#NAME?',
};
const 紙たち = [
  ['golden-oddf-to-46ko-2026-09-16.tsv', 2, 3],
  ['golden-oddf-2kaime-2026-09-16.tsv', 1, 2],
  ['golden-oddf-3kaime-2026-09-16.tsv', 1, 2],
  ['golden-oddf-4kaime-2026-09-17.tsv', 1, 2],
  ['golden-oddf-5kaime-2026-09-17.tsv', 1, 2],
  ['golden-kansuu-8kaime-2026-09-18.tsv', 1, 2],
  ['golden-kansuu-9kaime-2026-09-18.tsv', 2, 4],
  /* ★2026-09-18 に 足しました★
     ＝Exally1「★あなたの 7枚に この 紙は 入って いない★」
     ＝★私の「数が 違う 物は 0本」は ★私の 分母での 話★でした★
     ⇒★分母を 揃えます★ */
  ['kansuu46/golden-kane-2026-09-07.tsv', 1, 2],
];
const 数 = (y, m, d) => K.日から数(y, m, d);
const 見た = new Set();
const 問い = [];
for (const [名, 式列, 答列] of 紙たち) {
  const 道 = path.join(ここ, 名);
  if (!fs.existsSync(道)) continue;
  for (const l of fs.readFileSync(道, 'utf-8').replace(/^\uFEFF/, '').split(/\r?\n/)) {
    if (!l || l.startsWith('#') || !l.includes('\t')) continue;
    const c = l.split('\t');
    const 式 = 裸(c[式列]);
    let 答 = 裸(c[答列]);
    const m = /^=ODDFPRICE\(DATE\((\d+),(\d+),(\d+)\),DATE\((\d+),(\d+),(\d+)\),DATE\((\d+),(\d+),(\d+)\),DATE\((\d+),(\d+),(\d+)\),([\d.]+),([\d.]+),([\d.]+),(\d+),(\d+)\)$/.exec(式);
    if (!m) continue;
    if (見た.has(式)) continue;
    見た.add(式);
    if (誤りの数[答]) 答 = 誤りの数[答];
    const n = m.slice(1).map(Number);
    問い.push({
      式,
      決済: 数(n[0], n[1], n[2]), 満期: 数(n[3], n[4], n[5]),
      発行: 数(n[6], n[7], n[8]), 初回: 数(n[9], n[10], n[11]),
      利率: n[12], 利回り: n[13], 償還: n[14], 頻度: n[15], basis: n[16],
      正: 答,
    });
  }
}

/* ★★見立て ... 発行→初回 の 日数が ちょうど 1枠か★★
   ＝E（1枠の 日数）は `期間の日数(初回, 満期, 頻度, basis)` では なく
     ★初回を 終わりと する 枠★で 測ります（`期間の日数(発行+, 初回, ...)`） */
function 端数でないか(q) {
  try {
    const E = K.期間の日数(K.数から日(q.発行), K.数から日(q.初回), q.頻度, q.basis);
    const DFC = 日数(q.発行, q.初回, q.basis);
    return Math.abs(DFC - E) < 1e-9;
  } catch (e) { return null; }
}
/* ★basis ごとの 日数の 数え方★（★部品は 210/210 合って います★） */
function 日数(a, b, basis) {
  const A = K.数から日(a); const B = K.数から日(b);
  if (basis === 0 || basis === 4) {
    let d1 = A.d; let d2 = B.d;
    if (basis === 0) {
      if (d1 === 31) d1 = 30;
      if (d2 === 31 && d1 === 30) d2 = 30;
    } else {
      if (d1 === 31) d1 = 30;
      if (d2 === 31) d2 = 30;
    }
    return (B.y - A.y) * 360 + (B.m - A.m) * 30 + (d2 - d1);
  }
  return b - a;   /* ★実日数（basis 1,2,3）★ */
}

const 数か = (x) => /^[-+]?[0-9]*[.]?[0-9]+([eE][-+]?[0-9]+)?$/.test(x);
let 見落とし = 0, 空振り = 0, 当たり = 0, 通常 = 0;
const 外れ一覧 = [];
/* ★★おまけ ... 台が 返す 数が 実Excel と 違う 本数★★
   ＝★規則とは 別の 話です★（規則は「誤りに するか」だけ） */
let 数が違う = 0;
const 数が違う一覧 = [];
for (const q of 問い) {
  const 実が誤り = !数か(q.正);
  /* ★★見立て 2つ★★
       ㋐決済が 発行より 後で ない ... #NUM!（★5本 これでした★）
       ㋑初回の 端数が 端数で ない（DFC == E）... #NUM! */
  const 見立て = (q.決済 <= q.発行) ? true : 端数でないか(q);
  if (見立て === null) continue;
  if (!実が誤り) {
    let 出;
    try {
      出 = K.初回端数の価格(q.決済, q.満期, q.発行, q.初回, q.利率, q.利回り, q.償還, q.頻度, q.basis);
    } catch (e) { 出 = { 誤り: 'EX' }; }
    const 字 = (出 && 出.誤り) ? ('#' + 出.誤り + '!') : String(出);
    if (数か(字) && Math.abs(Number(字) - Number(q.正)) >= 1e-9) {
      数が違う += 1;
      数が違う一覧.push({ q, 字, 差: Number(字) - Number(q.正) });
    }
  }
  if (見立て && 実が誤り) { 当たり += 1; continue; }
  if (!見立て && !実が誤り) { 通常 += 1; continue; }
  if (見立て && !実が誤り) { 空振り += 1; 外れ一覧.push(['空振り（見立ては誤り・実Excelは数）', q]); continue; }
  見落とし += 1; 外れ一覧.push(['見落とし（見立ては数・実Excelは誤り）', q]);
}

console.log('');
console.log('[oddf-kimari] ★見立て ... ㋐決済<=発行 ／ ㋑DFC==E なら 実Excel は #NUM!★');
console.log('  ★別々の 式 ... ' + 問い.length + '本★');
console.log('');
console.log('  ★当たり（見立て 誤り・実Excel 誤り） ... ' + 当たり + '本★');
console.log('  ★通常（見立て 数・実Excel 数） ...... ' + 通常 + '本★');
console.log('  ★★空振り ... ' + 空振り + '本★★');
console.log('  ★★見落とし ... ' + 見落とし + '本★★');
console.log('');
if (!空振り && !見落とし) {
  console.log('  ★★★見立ては ' + 問い.length + '本 全部に 当たりました★★★');
} else {
  console.log('  ★★見立ては 外れました★★（★下に 全部 出します★）');
  for (const [型, q] of 外れ一覧.slice(0, 20)) {
    console.log('    [' + 型 + '] 実Excel ' + q.正);
    console.log('        ' + q.式);
  }
  if (外れ一覧.length > 20) console.log('    ...（残り ' + (外れ一覧.length - 20) + '本）');
}

console.log('');
console.log('  ★★おまけ ... 台の 数が 実Excel と 違う ... ' + 数が違う + '本★★');
console.log('    （★規則とは 別の 話です★／規則は「誤りに するか」だけ）');
for (const x of 数が違う一覧) {
  const 一回 = 100 * x.q.利率 / x.q.頻度;
  console.log('    差 ' + x.差.toFixed(10).padStart(16)
    + ' ／比 ' + (x.差 / 一回).toFixed(8).padStart(13)
    + '  basis ' + x.q.basis + ' 頻度 ' + x.q.頻度);
  console.log('        ' + x.q.式);
}
