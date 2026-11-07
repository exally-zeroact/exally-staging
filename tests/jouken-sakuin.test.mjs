/* ★条件の 溢れの 索引（等しい 条件・数えるだけの > <）が 全部 なめた 答えと 同じか・重く ならないか★（2026-11-07・値の直し3・叩き直し 止め1＝重さ）
   ★なぜ★ 条件に 範囲を 渡す 式（`SUMPRODUCT(1/COUNTIF(A1:A4000,A1:A4000))`・`COUNTIF(A,">"&A)`）が
     条件ごとに 全マスを なめて N×N＝N=4000 で 2.2〜3.8秒（この 木で 測った）
     ⇒ lib/shiki-kansuu.js の 索引で引く・索引で数える（溢れの 間だけ 作る）
   見る 所 2つ
     ①★同じ 答え★＝溢れ（索引の 道）の 1マスずつ と、同じ 条件を 1つずつ 打った 式（索引を 使わない 道）が 全部 同じ
       盤は 型を 混ぜる（数・数と 読める 字・%・時刻・日付の 字・大小の 字・写し字・真偽・空・="" ・誤り）
       ★物差しは 実Excel では なく 前からの 道★（実Excel との 合いは utsushiji・jouken-hairetsu の 紙が 見る）
     ②★重さ★＝N を 4倍に して 1本あたりが 4倍を 超えたら 赤（shiki-hyou-omosa と 同じ 線）
   使い方: node tests/jouken-sakuin.test.mjs */
import { createRequire } from 'module';
import path from 'path';
import { fileURLToPath } from 'url';
const require = createRequire(import.meta.url);
const 根 = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
require(path.join(根, 'lib/shiki-tsunagi.js'));
require(path.join(根, 'lib/shiki-kansuu.js'));
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

console.log('[jouken-sakuin] ②重さ＝N を 4倍に して 1本あたりが 4倍を 超えたら 赤');
const 一回 = (N, 式) => {
  const h = new H.表();
  for (let i = 1; i <= N; i++) { h.打つ('A' + i, i % 97); h.打つ('B' + i, i); }
  const t0 = process.hrtime.bigint();
  h.打つ('D1', 式(N));
  return Number(process.hrtime.bigint() - t0) / 1e6;
};
const 測る = (N, 式) => { 一回(200, 式); let m = Infinity; for (let k = 0; k < 3; k++) m = Math.min(m, 一回(N, 式)); return m; };
for (const [名, 式] of [
  ['SUMPRODUCT(1/COUNTIF(A,A))', (n) => `=SUMPRODUCT(1/COUNTIF(A1:A${n},A1:A${n}))`],
  ['SUM(SUMIF(A,A,B))', (n) => `=SUM(SUMIF(A1:A${n},A1:A${n},B1:B${n}))`],
  ['SUM(COUNTIF(A,">"&A))', (n) => `=SUM(COUNTIF(A1:A${n},">"&A1:A${n}))`],
]) {
  const 小 = 測る(1000, 式) / 1000, 大 = 測る(4000, 式) / 4000, 倍 = 大 / 小;
  const 書 = '      N=1000 … ' + (小 * 1000).toFixed(1) + 'µs/本 ／ N=4000 … ' + (大 * 1000).toFixed(1) + 'µs/本 ／ ★' + 倍.toFixed(2) + '倍★';
  if (!(大 > 0 && 小 > 0)) ng(名 + ' 測りが 0（空振り）');
  else if (倍 <= 4) { pass++; console.log('  ✓ ' + 名); console.log(書); }
  else { ng(名 + ' 1本あたりが ' + 倍.toFixed(2) + '倍（4倍まで）＝N の 2乗に 戻って いないか'); console.log(書); }
}
console.log('[jouken-sakuin] pass ' + pass + ' / fail ' + fail);
process.exit(fail ? 1 : 0);
