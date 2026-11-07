/* ★条件の 溢れの 索引（等しい 条件・数えるだけの > <）が 全部 なめた 答えと 同じか・重く ならないか★（2026-11-07・値の直し3・叩き直し 止め1＝重さ）
   ★なぜ★ 条件に 範囲を 渡す 式（`SUMPRODUCT(1/COUNTIF(A1:A4000,A1:A4000))`・`COUNTIF(A,">"&A)`）が
     条件ごとに 全マスを なめて N×N＝N=4000 で 2.2〜3.8秒（この 木で 測った）
     ⇒ lib/shiki-kansuu.js の 索引で引く・索引で数える（溢れの 間だけ 作る）
   見る 所 2つ
     ①★同じ 答え★＝溢れ（索引の 道）の 1マスずつ と、同じ 条件を 1つずつ 打った 式（索引を 使わない 道）が 全部 同じ
       盤は 型を 混ぜる（数・数と 読める 字・%・時刻・日付の 字・大小の 字・写し字・真偽・空・="" ・誤り）
       ★物差しは 実Excel では なく 前からの 道★（実Excel との 合いは utsushiji・jouken-hairetsu の 紙が 見る）
     ②★重さ★＝N=4000 で ★条件を 1マスずつ 見た 回数★（lib の 条件を見た数）が 20×N を 超えたら 赤（叩き直し2 止めB）
       ★時間では 測らない★＝前の「N を 4倍で 1本あたり 4倍まで」は N×N でも ちょうど 4倍で 見分けられず（索引を 切った 写しで 3.71・3.94倍 緑）
         2倍の 線に 下げても 溢れの 土台の 揺れ（1本 数µs）で N×N で ない 物まで 赤（2.0〜4.1倍）＝時間は 書くだけ
       ★回数は 揺れない★＝N×N なら 1600万回・直って いれば 数万回
       ★集まりを 作った 回数も 見る★（叩き直し2 止めB 残り）＝1式で 2回まで（毎回 作り直すと N 回＝4000）
       ★小数の 盤（B＝i＋0.5）★＝和・平均は 前と 端数まで 同じに する 為 軽い なめ方（N×N 回 だが 当てはまるか は 0回）
         ⇒ 当てはまるか 20×N まで・作った 2回まで・軽く なめた 行は N×N まで（それより 多ければ 覚えが 効いて いない）
       ★盤は 違う 値が 1種類・5種類・全部 違う の 3つ★＝前の i%97 は 速い 側だけで 遅い 型が 隠れた
       空振りの 確かめ（叩き直し2 で 測った・わざと 壊した 写し）
         lib の 索引の下限 を 1e9 ⇒ 全部 違う 盤の 9本が 赤（1600万回）・1種類／5種類は 覚えで 緑
         条件ごとに並べる の 覚えを 切る ⇒ 1種類・5種類 の 等しい 条件 4本が 赤（1600万・320万回）
   使い方: node tests/jouken-sakuin.test.mjs */
import { createRequire } from 'module';
import path from 'path';
import { fileURLToPath } from 'url';
const require = createRequire(import.meta.url);
const 根 = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
require(path.join(根, 'lib/shiki-tsunagi.js'));
const K関 = require(path.join(根, 'lib/shiki-kansuu.js'));
const H = require(path.join(根, 'lib/shiki-hyou.js'));

let pass = 0, fail = 0;
const ng = (s) => { fail++; console.log('  ✗ ' + s); };

/* ★盤★ A1:A80（条件の 範囲）・B1:B80（足す 範囲＝1..80） */
const 種 = [5, '="5"', '=" 5"', 0.5, '="50%"', '="12:00"', '="2025/1/1"', 45658, '="1,000"', 1000,
  'apple', 'APPLE', '="a*b"', '=TRUE', '=FALSE', null, '=""', '=NA()', '=1/0', -3, 2.5, '="2.5"', '="TRUE"', 'x', 12];
const 盤 = (h) => {
  for (let i = 0; i < 80; i++) {
    const v = 種[(i * 7) % 種.length];
    if (v !== null) h.打つ('A' + (i + 1), v);
    h.打つ('B' + (i + 1), i + 1);
    /* D＝整数・誤り・字・真偽 混ぜ（集まりで答える の 誤りの 番号・数で ない 物を 飛ばす を 見る） */
    h.打つ('D' + (i + 1), i % 11 === 3 ? '=NA()' : i % 13 === 5 ? 'x' : i % 17 === 7 ? '=TRUE' : i % 19 === 2 ? '=1/0' : i * 3 - 100);
  }
};
/* ★条件★ C1:C40 */
const 条件 = [5, '="5"', '="=5"', '="<>5"', '=">4"', '="<=0.5"', '=">=2025/1/1"', '="<12:00"', '="50%"', '="12:00"',
  '="apple"', '="=APPLE"', '="a*"', '="*"', '="a~*b"', '=""', '="="', '="<>"', '="#N/A"', '="<>#N/A"',
  '="TRUE"', '=TRUE', '=">apple"', '="<b"', '="1,000"', '=">-3"', '="<2.5"', '="2.5"', 2.5, 1000,
  '="x"', '="=x"', '="<>x"', '=">=12"', '="< 5"', '="= 5"', '=" 5"', '="?"', 12, null];
const 置く条件 = (h) => { 条件.forEach((v, i) => { if (v !== null) h.打つ('C' + (i + 1), v); }); };
const 字 = (h, a) => String(h.字(a));
const 式たち = [
  ['COUNTIF', (c) => `=COUNTIF(A1:A80,${c})`],
  ['SUMIF', (c) => `=SUMIF(A1:A80,${c},B1:B80)`],
  ['AVERAGEIF', (c) => `=AVERAGEIF(A1:A80,${c},B1:B80)`],
  ['COUNTIFS', (c) => `=COUNTIFS(A1:A80,${c},B1:B80,">10")`],
  ['COUNTIFS 後ろ', (c) => `=COUNTIFS(B1:B80,">10",A1:A80,${c})`],
  ['SUMIFS', (c) => `=SUMIFS(B1:B80,A1:A80,${c})`],
  ['AVERAGEIFS', (c) => `=AVERAGEIFS(B1:B80,A1:A80,${c},B1:B80,"<70")`],
  ['MAXIFS', (c) => `=MAXIFS(B1:B80,A1:A80,${c})`],
  ['MINIFS', (c) => `=MINIFS(B1:B80,A1:A80,${c})`],
  ['SUMIF 誤り混ぜ', (c) => `=SUMIF(A1:A80,${c},D1:D80)`],
  ['AVERAGEIF 誤り混ぜ', (c) => `=AVERAGEIF(A1:A80,${c},D1:D80)`],
  ['SUMIFS 誤り混ぜ', (c) => `=SUMIFS(D1:D80,A1:A80,${c})`],
  ['AVERAGEIFS 誤り混ぜ', (c) => `=AVERAGEIFS(D1:D80,A1:A80,${c})`],
  ['MAXIFS 誤り混ぜ', (c) => `=MAXIFS(D1:D80,A1:A80,${c})`],
  ['MINIFS 誤り混ぜ', (c) => `=MINIFS(D1:D80,A1:A80,${c})`],
  ['SUMIF 足す＝条件の範囲（端数）', (c) => `=SUMIF(A1:A80,${c})`],
  ['MAXIFS 足す＝条件の範囲', (c) => `=MAXIFS(A1:A80,A1:A80,${c})`],
];
console.log('\n[jouken-sakuin] ①溢れ（索引）と 1つずつ（全部 なめる）が 同じ 答えか');
for (const [名, 式] of 式たち) {
  const h = new H.表(); 盤(h); 置く条件(h);
  h.打つ('E1', 式('C1:C40'));
  let 同じ = 0;
  for (let k = 0; k < 条件.length; k++) {
    h.打つ('G' + (k + 1), 式('C' + (k + 1)));
    const 溢 = 字(h, 'E' + (k + 1)), 一 = 字(h, 'G' + (k + 1));
    if (溢 === 一) 同じ++; else ng(名 + ' 条件 C' + (k + 1) + ' 溢れ ' + 溢 + ' ／ 1つ ' + 一);
  }
  if (同じ === 条件.length) { pass++; console.log('  ✓ ' + 名 + ' ' + 同じ + '/' + 条件.length); }
}
/* ★盤の 置き方が 効いて いるか★（字の "5"・真偽・誤り・="" が 化けて いれば 比べても 意味が 無い） */
{
  const h = new H.表(); 盤(h);
  const 全 = new Set(); let 字の数 = 0;
  for (let i = 1; i <= 80; i++) { const v = h.値('A' + i); 全.add(v.型); if (v.型 === '字' && v.値 === '5') 字の数++; }
  if (['数', '字', '真偽', '空', '誤'].every((t) => 全.has(t)) && 字の数 > 0) { pass++; console.log('  ✓ 盤の 型が 混ざって いる ' + [...全].join('・') + '（字の "5" ' + 字の数 + '）'); }
  else ng('盤の 置き方 ' + [...全].join() + ' 字の "5" ' + 字の数);
}

console.log('[jouken-sakuin] ②重さ＝N=4000 で 条件を 見た 回数が 20×N を 超えたら 赤（盤 3つ）');
let 盤の種 = 1, 盤の小数 = false;
const 一回 = (N, 式) => {
  const h = new H.表();
  const K = 盤の種 === 0 ? N : 盤の種;
  for (let i = 1; i <= N; i++) { h.打つ('A' + i, i % K); h.打つ('B' + i, 盤の小数 ? i + 0.5 : i); }
  const t0 = process.hrtime.bigint();
  h.打つ('D1', 式(N));
  return Number(process.hrtime.bigint() - t0) / 1e6;
};
const N = 4000;
for (const [盤名, 種k, 小数] of [['1種類', 1, false], ['5種類', 5, false], ['全部 違う', 0, false], ['全部 違う・B 小数', 0, true]]) for (const [名0, 式] of [
  ['SUMPRODUCT(1/COUNTIF(A,A))', (n) => `=SUMPRODUCT(1/COUNTIF(A1:A${n},A1:A${n}))`],
  ['SUM(SUMIF(A,A,B))', (n) => `=SUM(SUMIF(A1:A${n},A1:A${n},B1:B${n}))`],
  ['SUM(COUNTIF(A,">"&A))', (n) => `=SUM(COUNTIF(A1:A${n},">"&A1:A${n}))`],
  ['SUM(SUMIF(A,"<"&A,B))', (n) => `=SUM(SUMIF(A1:A${n},"<"&A1:A${n},B1:B${n}))`],
  ['SUM(COUNTIFS(A,">"&A))', (n) => `=SUM(COUNTIFS(A1:A${n},">"&A1:A${n}))`],
  ['SUM(COUNTIF(A,"<>"&A))', (n) => `=SUM(COUNTIF(A1:A${n},"<>"&A1:A${n}))`],
  ['SUM(MAXIFS(B,A,"<"&A))', (n) => `=SUM(MAXIFS(B1:B${n},A1:A${n},"<"&A1:A${n}))`],
  ['SUM(AVERAGEIF(A,">="&A,B))', (n) => `=SUM(AVERAGEIF(A1:A${n},">="&A1:A${n},B1:B${n}))`],
  ['SUM(SUMIFS(B,A,"<="&A))', (n) => `=SUM(SUMIFS(B1:B${n},A1:A${n},"<="&A1:A${n}))`],
]) {
  盤の種 = 種k; 盤の小数 = 小数;
  const 名 = 盤名 + ' ' + 名0;
  const 前 = K関.条件を見た数(), 前集 = K関.集まりの数();
  const ms = 一回(N, 式);
  const 回 = K関.条件を見た数() - 前, 後集 = K関.集まりの数();
  const 作 = 後集.作った - 前集.作った, 軽 = 後集.軽くなめた - 前集.軽くなめた;
  const 書 = '      N=' + N + ' … 条件を 見た 回数 ' + 回 + '（' + (回 / N).toFixed(1) + '×N）・集まりを 作った ' + 作 + '・軽く なめた ' + 軽 + '／' + ms.toFixed(0) + 'ms（時間は 書くだけ）';
  const 悪 = [];
  if (!(ms > 0)) 悪.push('測りが 0（空振り）');
  if (回 > 20 * N) 悪.push('条件を 見た 回数が ' + 回 + '（20×N＝' + 20 * N + ' まで）＝N の 2乗に 戻って いないか');
  if (作 > 2) 悪.push('集まりを ' + 作 + '回 作った（2回まで）＝覚えが 効いて いない');
  if (軽 > N * N) 悪.push('軽く なめた 行が ' + 軽 + '（N×N まで）');
  if (!小数 && 軽 > 0) 悪.push('整数の 盤で 軽く なめた（累計の 道を 通って いない）');
  if (!悪.length) { pass++; console.log('  ✓ ' + 名); console.log(書); }
  else { ng(名 + ' ' + 悪.join('／')); console.log(書); }
}
console.log('[jouken-sakuin] pass ' + pass + ' / fail ' + fail);
process.exit(fail ? 1 : 0);
