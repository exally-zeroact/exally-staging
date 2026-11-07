/* ★条件で 集める 一族（COUNTIF SUMIF AVERAGEIF COUNTIFS SUMIFS AVERAGEIFS MAXIFS MINIFS）を 実Excel と 同じに★（2026-11-07・値の直し3）
   物差し（実Excel 16.0 build 20527・Formula2＝今の 道）
     紙9 docs/measured/utsushiji-excel-9-2026-11-07.tsv（34本・測る 道具 toru-utsushiji-9.ps1）
     盤12 A1:A5=1..5・B1:B5=10..50・C1:C5=4,=NA(),TRUE,="5",'7・D1:D5=1,=NA(),3,4,5・E1:E2='>0,'>2・A7:E7=1..5
   見る 所 3つ
     ①★条件に 並び・範囲＝条件ごとに 答えて 溢れる★（右＝migi・下＝shita も 比べる）
     ②★範囲の 形（行と 列）が 違えば #VALUE!★（長さが 同じでも）
     ③★選ばれた マスの 誤りは 伝わる・真偽と 字は 飛ばす★／条件の 範囲の 誤りは 当たらない
   ★紙を 読んで 押す＝紙に 無い 式は 押さない★ */
import { createRequire } from 'module';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
const require = createRequire(import.meta.url);
const 根 = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
require(path.join(根, 'lib/shiki-tsunagi.js'));
require(path.join(根, 'lib/shiki-kansuu.js'));
const H = require(path.join(根, 'lib/shiki-hyou.js'));

const 名 = 'utsushiji-excel-9-2026-11-07.tsv', 本数 = 34;
let pass = 0, fail = 0, 窓の崩れ = 0;
const 紙 = fs.readFileSync(path.join(根, 'docs/measured', 名), 'utf8').split(/\r?\n/).filter((l) => l && l[0] !== '#').map((l) => {
  const 欄 = l.split('\t'), 取る = (頭) => { const c = 欄.find((x) => x.startsWith(頭)); return c === undefined ? undefined : c.slice(頭.length); };
  const 行 = { 式: 欄[0], 値: 取る('F2='), 右: 取る('migi='), 下: 取る('shita='), z0: 取る('z0='), 型: 取る('kata=') };
  /* ★2つ目の 窓★（monosashi-mado の 決まり）＝窓が 無い・"0" なのに z0=False は 紙の 崩れ */
  if (行.z0 === undefined || 行.型 === undefined || (行.型 === 'Double' && (行.値 === '0') !== (行.z0 === 'True'))) { 窓の崩れ++; console.log('✗ 窓の 崩れ ' + l); }
  return 行;
});
if (紙.length !== 本数) { fail++; console.log('✗ ' + 名 + ' の 行が ' + 本数 + '本で ない: ' + 紙.length); }

const 盤12 = (h) => {
  const c5 = [4, '=NA()', '=TRUE', '="5"', '="7"'], d5 = [1, '=NA()', 3, 4, 5];
  for (let i = 0; i < 5; i++) {
    h.打つ('A' + (i + 1), i + 1); h.打つ('B' + (i + 1), 10 * (i + 1));
    h.打つ('C' + (i + 1), c5[i]); h.打つ('D' + (i + 1), d5[i]);
    h.打つ(String.fromCharCode(65 + i) + '7', i + 1);
  }
  h.打つ('E1', '=">0"'); h.打つ('E2', '=">2"');
};
const 字 = (h, a) => { const s = String(h.字(a)); return s === '' ? '(sora)' : s; };
const 合う = (出, 見) => 出 === 見 || (出 !== '(sora)' && 見 !== '(sora)' && isFinite(+出) && isFinite(+見) && Math.abs(+出 - +見) < 1e-9);
for (const 行 of 紙) {
  const h = new H.表(); 盤12(h); h.打つ('BZ1', 行.式);
  const 出 = [字(h, 'BZ1'), 字(h, 'CA1'), 字(h, 'BZ2')], 見 = [行.値, 行.右, 行.下];
  if (出.every((x, i) => 合う(x, 見[i]))) pass++;
  else { fail++; console.log('✗ ' + 行.式 + ' 台 ' + 出.join(' / ') + ' Excel ' + 見.join(' / ')); }
}
/* ★盤の 置き方が 効いて いるか★（字の "5"・真偽・誤りが 別の 型に 化けて いれば 紙と 比べても 意味が 無い） */
{
  const h = new H.表(); 盤12(h);
  const 型 = ['C2', 'C3', 'C4', 'C5', 'D2', 'E1'].map((a) => { const v = h.値(a); return v.型 + ':' + v.値; });
  const 見 = ['誤:#N/A', '真偽:true', '字:5', '字:7', '誤:#N/A', '字:>0'];
  if (JSON.stringify(型) === JSON.stringify(見)) pass++; else { fail++; console.log('✗ 盤12 の 置き方 ' + JSON.stringify(型)); }
}
fail += 窓の崩れ;
console.log('[jouken-hairetsu] pass ' + pass + ' / fail ' + fail + '（窓の 崩れ ' + 窓の崩れ + '）');
process.exit(fail ? 1 : 0);
