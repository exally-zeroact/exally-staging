/* jsso-to-dai-wo-awaseru.mjs — ★JS層と 台を 突き合わせる★（2026-09-18・㋑の ⑴）
 *
 *  ★★なぜ★★
 *    ㋑（計算を 台に 回す）で ★JS層と 台の どちらが 先か★を 決めます。
 *    ★経営者1 の 見立て★ ... JS層 -> 台 -> 借り物
 *    ★★但し 2人とも 測って いません★★
 *    ⇒★★測らずに 順を 決めません★★（★2026-09-18 に 3回 踏んだ 形★）
 *
 *  ★★JS層が 受ける 名前は 27個★★（`exally-formula.js` の `_jsSet` から ★機械で 拾います★）
 *    ★手で 書きません★＝★足された 日に 黙って ずれます★
 *
 *  ★★測る 事★★
 *    ・★JS層が 答える 行は 何本か★（★分母★）
 *    ・★JS層と 台が ★同じ★ 答えを 出した 本数★（★これも 数えます★）
 *    ・★違った 本数★／★どちらが 実Excel と 合ったか★
 *
 *  ★★この 道具が 見て いない 事★★
 *    ・★画面では ありません★（node の 台）
 *    ・★`convertFormula` は 通して いません★
 *      ＝★JS層は `convertFormula` の 前に 呼ばれます★（`book.html:2614` で 確かめた）
 *      ＝★台も 同じ 所に 入れる 積もりです★＝★同じ 入口で 比べて います★
 *    ・★紙に 在る 式だけ★＝★お客さんが 打つ 式 全部では ありません★
 *
 *  ★★★出た 物（2026-09-18）★★★
 *    ★JS層が 実際に 答えた 行★ ... ★28行★（★分母★）
 *      ISOMITTED 16 ／ REDUCE 4 ／ MAP 3 ／ OFFSET 2 ／ MAKEARRAY 2 ／ SCAN 1
 *    ★同じ 答え★ ... ★27行★ ／ ★違う 答え★ ... ★1行★
 *      `=MAP(A1:A5,LAMBDA(x,y,x))` ... JS層 `#ERROR!` ／ ★台 `#VALUE!`★ ／ 実Excel `#VALUE!`
 *      ⇒★JS層だけ 合った 0行 ／ ★台だけ 合った 1行★★
 *
 *  ★★決め ... `JS層 -> 台 -> 借り物`★★（★経営者1 の 見立て通り★）
 *    ★訳★
 *      ・★JS層が 答える 28行の うち 27行は どちらでも 同じ★
 *      ・★残り 1行は 台が 勝ちます★＝★JS層を 先に すると ★1行だけ★ 損を します★
 *      ・★★JS層が 受ける 27個の うち 21個は 紙に 1行も 在りません★★
 *        （DSUM / PERCENTILE / XIRR / MIRR / XNPV / FREQUENCY ...）
 *        ＝★その 21個は ★測って いません★★
 *        ⇒★★測って いない 物の 順を 変えるのは 危ない★★
 *      ⇒★★損 1行と 引き換えに ★測って いない 21個を 動かさない★★★
 *    ★1行の 損は 別に 直せます★（JS層の `MAP` の 引数の 数違いを `#VALUE!` に）
 *
 *  ★★最初 分母が 違って いました（★私の 道具の 欠陥★）★★
 *    ★前★ 同じ 20行 ／ 違う 8行
 *    ★後★ 同じ 27行 ／ 違う 1行
 *    ★訳★ ... ★材料を 板に 入れる 前に `_jsComputeFormula` を 呼んで いました★
 *           ＝JS層が ★空の 板★を 読み `0` や `#NAME?` を 返して いた
 *    ⇒★`kansuu46-no-dodai.mjs` の 頭に ★同じ 事が 書いて あります★★
 *      「前は JS層を 呼んでから 材料を 入れて いた／D系で 合うのが 0本 -> 5本」
 *    ⇒★★読んで いたのに 同じ 所を 踏みました★★
 *
 *  使い方: node docs/measured/jsso-to-dai-wo-awaseru.mjs
 */
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';
import { 建てる } from './honban-no-michi.mjs';

const ここ = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(ここ, '..', '..');
const require_ = createRequire(import.meta.url);
const H = require_(path.join(ROOT, 'lib/shiki-hyou.js'));

/* ══ ★JS層が 受ける 名前★（★機械で 拾う★） ══ */
const EFの字 = fs.readFileSync(path.join(ROOT, 'exally-formula.js'), 'utf-8');
const 頭 = EFの字.indexOf('var _jsSet = {');
const 尻 = EFの字.indexOf('};', 頭);
const JS層の名 = new Set(
  [...new Set((EFの字.slice(頭, 尻).match(/(^|[\s{,])([A-Z][A-Z0-9._]*)\s*:\s*1/g) || [])
    .map((x) => x.replace(/[\s{,]/g, '').replace(':1', '')))]
);
if (JS層の名.size < 20) {
  console.log('★★`_jsSet` が ' + JS層の名.size + '個しか 読めません＝読み方が 壊れて います★★');
  process.exit(3);
}

/* ══ ★紙★（★`osu-dai-dake.mjs` と 同じ 並び★＝★分母を 揃える★） ══ */
const 紙たち = [
  { 名: 'golden-kansuu-8kaime-2026-09-18.tsv', 式列: 1, 答列: 2 },
  { 名: 'golden-kansuu-9kaime-2026-09-18.tsv', 式列: 2, 答列: 4 },
  { 名: 'golden-kansuu-7kaime-2026-09-18.tsv', 式列: 1, 答列: 2 },
  { 名: 'golden-oddl-6kaime-2026-09-18.tsv', 式列: 1, 答列: 2 },
  { 名: 'golden-oddf-buhin-2026-09-16.tsv', 式列: 3, 答列: 4 },
  { 名: 'golden-oddf-to-46ko-2026-09-16.tsv', 式列: 2, 答列: 3,
    材料: { B1: 2, B2: 4, B3: 6, B4: 8, B5: 10 } },
  { 名: 'golden-oddf-2kaime-2026-09-16.tsv', 式列: 1, 答列: 2 },
  { 名: 'golden-oddf-3kaime-2026-09-16.tsv', 式列: 1, 答列: 2 },
  { 名: 'golden-oddf-4kaime-2026-09-17.tsv', 式列: 1, 答列: 2 },
  { 名: 'golden-oddf-5kaime-2026-09-17.tsv', 式列: 1, 答列: 2 },
  { 名: 'kansuu46/golden-kane-2026-09-07.tsv', 式列: 1, 答列: 2 },
  { 名: 'kansuu46/golden-346-2026-09-08.tsv', 式列: 1, 答列: 2 },
];

/* ★材料★（★`kansuu46-no-dodai.mjs` と 同じ★） */
const 既定の材料 = { A1: 1, A2: 2, A3: 3, A4: 4, A5: 5, B1: 1, B2: 2, B4: 4, B5: 5,
  C1: 1, C2: 3, C3: 5, C4: 7, C5: 9, D1: 9, D2: 7, D3: 5, D4: 3, D5: 1,
  F1: 1, G1: 2, F2: 10, G2: 20 };

function 板を作る(足す) {
  const h = H.表();
  const m = Object.assign({}, 既定の材料, 足す || {});
  for (const k of Object.keys(m)) h.打つ(k, String(m[k]));
  if (!足す || 足す.B3 === undefined) h.打つ('B3', '=1/0');
  return h;
}

const 台 = await 建てる();
const { EF, hf, SID } = 台;

let 分母 = 0, 同じ = 0, 違う = 0, 台だけ誤 = 0;
const 内訳 = {};
const 違い = [];
let 読んだ紙 = 0, 行合計 = 0;

for (const 紙 of 紙たち) {
  const p = path.join(ROOT, 'docs/measured', 紙.名);
  if (!fs.existsSync(p)) { console.log('  ★紙が 在りません★ ' + 紙.名); continue; }
  読んだ紙++;
  const 行 = fs.readFileSync(p, 'utf-8').split(/\r?\n/).filter((l) => l && !l.startsWith('#'));
  for (const l of 行) {
    const c = l.split('\t');
    const 式 = c[紙.式列], 正 = c[紙.答列];
    if (!式 || !式.startsWith('=')) continue;
    行合計++;
    const m = /^=([A-Z][A-Z0-9._]*)\(/.exec(式);
    if (!m || !JS層の名.has(m[1])) continue;   /* ★JS層が 受けない 物は 見ません★ */
    分母++;
    内訳[m[1]] = (内訳[m[1]] || 0) + 1;

    /* ★★材料を 先に 板へ 入れる★★（★2026-09-18 に ここで 1回 踏みました★）
         ★前は 入れずに `_jsComputeFormula` を 呼んで いました★
         ⇒★JS層が ★空の 板★を 読み ★DSUM などが 全部「答えない」★に なって いました★
         ⇒★分母が 28行しか 出ませんでした★
         ★同じ 事が `kansuu46-no-dodai.mjs` の 頭に 書いて あります★
           「★前は JS層を 呼んでから 材料を 入れて いました／9本の D系で 合うのが 0本 -> 5本★」
         ⇒★★読んで いたのに 同じ 所を 踏みました★★ */
    {
      const m2 = Object.assign({}, 既定の材料, 紙.材料 || {});
      const 表 = [];
      for (let r = 0; r < 8; r++) 表.push([null, null, null, null, null, null, null, null]);
      const 場所 = (k) => ({ c: k.charCodeAt(0) - 65, r: Number(k.slice(1)) - 1 });
      for (const k of Object.keys(m2)) {
        const b = 場所(k);
        if (b.r >= 0 && b.r < 8 && b.c >= 0 && b.c < 8) 表[b.r][b.c] = m2[k];
      }
      if (!紙.材料) 表[2][1] = '=1/0';      /* B3 */
      try { hf.setSheetContent(SID, 表); } catch (e) { /* 気に しない */ }
    }
    /* ★JS層★（★`convertFormula` の 前＝本番と 同じ 入口★） */
    let js = null;
    try { js = EF._jsComputeFormula(0, 式); } catch (e) { js = '★投げた★'; }
    /* ★台★ */
    const h = 板を作る(紙.材料);
    let dai;
    try { h.打つ('BZ1', 式); dai = String(h.字('BZ1')); } catch (e) { dai = '★投げた★'; }

    const jsの字 = (js === null || js === undefined) ? '(答えない)' : String(js);
    if (jsの字 === '(答えない)') { 分母--; 内訳[m[1]]--; continue; }  /* ★JS層が 実際に 答えた 物だけ★ */

    /* ★★紙の 答えの 列は `.Value2`★★＝★誤りは ★負の 番号★で 入って います★
         ＝`#VALUE!` は -2146826273 ／ `#NUM!` は -2146826252 ...
         ⇒★番号の まま 比べると ★合って いるのに 違うと 出ます★★
         ⇒★名前に 直してから 比べます★（★2026-09-18 に この 道具で 1本 踏みました★） */
    const 誤りの番号 = {
      '-2146826281': '#DIV/0!', '-2146826246': '#N/A', '-2146826259': '#NAME?',
      '-2146826288': '#NULL!', '-2146826252': '#NUM!', '-2146826265': '#REF!',
      '-2146826273': '#VALUE!',
    };
    const 直す = (s) => (誤りの番号[String(s).trim()] || String(s));
    const 近い = (a0, b0) => {
      const a = 直す(a0), b = 直す(b0);
      if (a === b) return true;
      const x = Number(a), y = Number(b);
      return isFinite(x) && isFinite(y) && Math.abs(x - y) <= Math.max(1e-9, Math.abs(y) * 1e-9);
    };
    if (近い(jsの字, dai)) { 同じ++; continue; }
    違う++;
    const JSが正 = 近い(jsの字, 正), 台が正 = 近い(dai, 正);
    if (台が正 && !JSが正) 台だけ誤++;
    違い.push({ 式: 式, JS: jsの字, 台: dai, 正: 正, JSが正: JSが正, 台が正: 台が正 });
  }
}

console.log('');
console.log('★★JS層と 台を 突き合わせる★★');
console.log('');
console.log('  ★JS層が 受ける 名前★ ..... ' + JS層の名.size + '個（★`_jsSet` から 機械で 拾いました★）');
console.log('  ★読んだ 紙★ ............. ' + 読んだ紙 + '枚 ／ ★式の 行★ ' + 行合計 + '行');
console.log('  ★★JS層が 実際に 答えた 行★★ ... ★' + 分母 + '行★（★これが 分母★）');
console.log('');
console.log('  ★★同じ 答え★★ ........... ★' + 同じ + '行★');
console.log('  ★★違う 答え★★ ........... ★' + 違う + '行★');
console.log('');
if (Object.keys(内訳).length) {
  console.log('  ★関数ごと★');
  for (const k of Object.keys(内訳).sort((a, b) => 内訳[b] - 内訳[a])) {
    if (内訳[k] > 0) console.log('    ' + k.padEnd(14) + String(内訳[k]).padStart(5) + '行');
  }
}
if (違い.length) {
  console.log('');
  console.log('★★違った 行（★どちらが 実Excel と 合ったか★）★★');
  let JS勝 = 0, 台勝 = 0, 両方違 = 0;
  for (const x of 違い) {
    if (x.JSが正 && !x.台が正) JS勝++;
    else if (x.台が正 && !x.JSが正) 台勝++;
    else 両方違++;
  }
  console.log('  ★JS層だけ 合った★ ... ' + JS勝 + '行');
  console.log('  ★★台だけ 合った★★ ... ★' + 台勝 + '行★');
  console.log('  ★どちらも 違う★ ..... ' + 両方違 + '行');
  console.log('');
  for (const x of 違い.slice(0, 20)) {
    console.log('    ' + x.式.slice(0, 66));
    console.log('      JS層「' + x.JS.slice(0, 30) + '」' + (x.JSが正 ? '★○★' : '×')
      + ' ／ 台「' + x.台.slice(0, 30) + '」' + (x.台が正 ? '★○★' : '×')
      + ' ／ 実Excel「' + String(x.正).slice(0, 30) + '」');
  }
  if (違い.length > 20) console.log('    （★あと ' + (違い.length - 20) + '行★）');
}
console.log('');
console.log('★言えない 事★');
console.log('  ・★画面では ありません★（node の 台）');
console.log('  ・★`convertFormula` は 通して いません★＝★JS層と 同じ 入口で 比べて います★');
console.log('  ・★紙に 在る 式だけ★＝★お客さんが 打つ 式 全部では ありません★');
console.log('  ・★JS層が「答えない」を 返した 行は 分母から 外して います★');
console.log('  ・★★JS層が 受ける 27個の うち 21個は 紙に 1行も 在りません★★（★測って いません★）');
try { hf.destroy(); } catch (e) { /* 気に しない */ }
