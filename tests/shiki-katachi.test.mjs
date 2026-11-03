/* shiki-katachi.test.mjs — ★切った かたまりを「形」に する★（2026-09-11）
 *
 *  ★★土台を 自分で 作る ②枚目★★（①は `lib/shiki-kiru.js`＝字に 切る）
 *    土台は 5つ … ①読む ②★形に する★ ③頼りの 地図 ④順番と 計算 ⑤溢れ
 *
 *  ★★測り方は 2つ★★（1つでは 足りません）
 *    ①★形から 字に 戻したら 元と 1バイトも 違わない★
 *       ⇒ 取りこぼしも 足しすぎも 0（実Excel を 呼ばずに 測れる）
 *       ★でも これだけでは「正しく 読めた」証しに なりません★
 *         `1+2*3` を 足し算 先に 読んでも、字に 戻せば 同じ 字です。
 *    ②★実Excel の 答えと 合う★（結ぶ 順番の 証し）
 *       ⇒ 形を そのまま 計算して 実Excel が 出した 答えと 突き合わせる
 *
 *  ★★実Excel に 打たせて 測った 順番★★（2026-09-11）
 *    `=2^3^2` → ★64★（左から）  `=-2^2` → ★4★（前の - が 先）
 *    `=10%%` → ★0.001★          `=-3%` → ★-0.03★
 *    `="a"&1+2` → ★a3★（+ が & より 先）
 *    `=1+2<4` → ★TRUE★（比べるのが 一番 後）
 *
 *  ★この 台は エンジンを 建てません★＝字を 読んで 形に するだけ。
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(path.join(ROOT, 'package.json'));
const K = require_(path.join(ROOT, 'lib/shiki-kiru.js'));
const C = require_(path.join(ROOT, 'lib/shiki-katachi.js'));

let pass = 0, fail = 0;
const T = (n, f) => {
  try { f(); pass++; console.log('  ✓ ' + n); }
  catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + e.message); }
};

console.log('');
console.log('[shiki-katachi] ★切った かたまりを 形に する★');

const 形に = (式) => {
  const k = K.切る(式);
  if (!k.ok) throw new Error('切れない … ' + k.なぜ);
  const c = C.形にする(k.出);
  if (!c.ok) throw new Error('形に ならない … ' + c.なぜ);
  return c.形;
};

/* ══ ①字に 戻る ══ */
const 戻り試 = [
  '=1+2*3', '=(1+2)*3', '=-5', '=10%', '=2^3', '="あ"&"い"', '="引用""符"',
  '=A1', '=$A$1', '=A1:B2', '=Sheet1!A1', "='歩合 表'!A1",
  '=SUM(A1:A3)', '=SUM(A1:A3,B1)', '=IF(A1<0,"負","正")', '=A1<>B1',
  '={1,2;3,4}', '=#N/A', '=IFERROR(A1,#N/A)', '=表[列]',
  '=INDEX(Table1[#Data], MATCH(B4, Table1[#Data], 0))',   /* ★飾りの 空白★ */
  '=SUM( 1 , 2 )', '=SUM(A1:A3 B1:B3)',                   /* ★交わりの 空白★ */
];

T('★①形から 字に 戻すと 元に 戻る（' + 戻り試.length + '通り）★', () => {
  const 違い = [];
  for (const t of 戻り試) {
    let 形;
    try { 形 = 形に(t); } catch (e) { 違い.push(t + ' … ' + e.message); continue; }
    const 戻 = '=' + C.字に戻す(形);
    if (戻 !== t) 違い.push(t + ' → ' + 戻);
  }
  if (違い.length) throw new Error('★' + 違い.length + '本 違う★  ' + 違い.slice(0, 3).join(' ／ '));
});

/* ══ ②実Excel の 答えと 合う ══ */
const 数に = (v) => { if (typeof v === 'number') return v; if (v === true) return 1; if (v === false) return 0;
  const n = Number(v); return Number.isFinite(n) ? n : NaN; };
const 字に = (v) => (typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : String(v));
function 計算(f) {
  switch (f.種) {
    case '数': return Number(f.値);
    case '字': return f.値.slice(1, -1).split('""').join('"');
    case '名': if (f.値 === 'TRUE') return true; if (f.値 === 'FALSE') return false; throw new Error('名 ' + f.値);
    case '括': return 計算(f.子);
    case '前': { const v = 計算(f.子); return f.記 === '-' ? -数に(v) : 数に(v); }
    case '後': return 数に(計算(f.子)) / 100;
    case '呼': if (f.名 === 'SUM') return f.引数.map(計算).reduce((a, b) => 数に(a) + 数に(b), 0); throw new Error('呼 ' + f.名);
    case '二': {
      const a = 計算(f.左), b = 計算(f.右);
      if (f.記 === '+') return 数に(a) + 数に(b);
      if (f.記 === '-') return 数に(a) - 数に(b);
      if (f.記 === '*') return 数に(a) * 数に(b);
      if (f.記 === '/') return 数に(a) / 数に(b);
      if (f.記 === '^') return Math.pow(数に(a), 数に(b));
      if (f.記 === '&') return 字に(a) + 字に(b);
      if (f.記 === '=') return a === b;
      if (f.記 === '<>') return a !== b;
      if (f.記 === '<') return 数に(a) < 数に(b);
      if (f.記 === '>') return 数に(a) > 数に(b);
      if (f.記 === '<=') return 数に(a) <= 数に(b);
      if (f.記 === '>=') return 数に(a) >= 数に(b);
      throw new Error('記 ' + f.記);
    }
  }
  throw new Error('種 ' + f.種);
}

/* ★実Excel に 打たせて 測った 答え★（2026-09-11） */
const 正 = [
  ['=1+2*3', '7'], ['=2*3+1', '7'], ['=(1+2)*3', '9'], ['=2^3^2', '64'], ['=-2^2', '4'],
  ['=1-2-3', '-4'], ['=100/10/2', '5'], ['=1&2&3', '123'], ['=1<2', 'TRUE'], ['=2%*100', '2'],
  ['=-3%', '-0.03'], ['="a"&1+2', 'a3'], ['=1+2<4', 'TRUE'], ['=TRUE', 'TRUE'], ['=1=1', 'TRUE'],
  ['=2<>3', 'TRUE'], ['=3-(-2)', '5'], ['=--3', '3'], ['=+-3', '-3'], ['=1*-2', '-2'],
  ['=10%%', '0.001'], ['=(1+2)*(3+4)', '21'], ['=SUM(1,2)*3', '9'],
  /* ★★べき乗が 掛け算より 先か★★（2026-09-11 足した）
     ★`=2^3^2` だけでは 捕まりませんでした★＝力を 下げても 赤に ならなかった。
     ⇒★掛け算・割り算・足し算・つなぎ と 並べた 時に 初めて 分かります★ */
  ['=2*3^2', '18'], ['=2+3^2', '11'], ['=-2^2*3', '12'], ['=2^2*3', '12'],
  ['=12/2^2', '3'], ['=2^3*2', '16'], ['=1+2^3-4', '5'], ['=2&3^2', '29'],
];

T('★★②実Excel の 答えと 合う（結ぶ 順番の 証し・' + 正.length + '本）★★', () => {
  const 違い = [];
  for (const [式, 期待] of 正) {
    let 出;
    try { 出 = 字に(計算(形に(式))); } catch (e) { 出 = '★' + e.message + '★'; }
    if (出 !== 期待) 違い.push(式 + ' 実Excel「' + 期待 + '」／うち「' + 出 + '」');
  }
  if (違い.length) throw new Error('★' + 違い.length + '本 違う★  ' + 違い.slice(0, 3).join(' ／ '));
  console.log('      … 2^3^2=64／-2^2=4／10%%=0.001／"a"&1+2=a3 も 合う');
});

T('★★読めない 式は 断る（勝手に 直さない）★★', () => {
  const 壊 = ['=1+', '=(1+2', '=SUM(1,', '=)1(', '={1,2'];
  const 通った = [];
  for (const t of 壊) {
    const k = K.切る(t);
    if (!k.ok) continue;                      /* 切る 段で 断れば よい */
    if (C.形にする(k.出).ok) 通った.push(t);
  }
  if (通った.length) throw new Error('★' + 通った.length + '本 通した★  ' + 通った.join(' ／ '));
});

console.log('');
console.log('  ' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
