/* ★条件の 写し字（* ? ~）★＝COUNTIF／SUMIF／AVERAGEIF／COUNTIFS が 型紙で 当てるか
   物差し（実Excel 16.0 build 20527・2026-11-07）
     紙1 docs/measured/utsushiji-excel-2026-11-07.tsv（27本）盤 A1:A8=apple,Apple,ab,a*b,5(数),空,a?c,15(数)／B1:B8=1..8
     紙2 docs/measured/utsushiji-excel-2-2026-11-07.tsv（39本・設計前 taiketsu 止め1 の 測り足し）
         盤2 A1:A10=apple,#N/A,TRUE,="",a~b," x",x,5(数),#DIV/0!,空／B1:B10=1..10
         盤3（retsu-x の 盤）H1:H5=1..5・I1:I5=10..50・L1="*"
     紙3 docs/measured/utsushiji-excel-3-2026-11-07.tsv（39本・叩き直し 止め1＝数の 形と D系）
     紙4 docs/measured/utsushiji-excel-4-2026-11-07.tsv（9本・字の マスを 数と 読む 範囲）
     紙5 docs/measured/utsushiji-excel-5-2026-11-07.tsv（19本・千の カンマ・> < は 同じ 型だけ）
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
/* ★紙の 行の 後ろ 2欄 = 2つ目の 窓★（monosashi-mado の 決まり）＝ z0=（=(式)=0 の 真偽）・kata=（Value2 の 型）
   ★窓が 無い 行・窓と 値が 食い違う 行（"0" なのに z0=False 等）は 紙の 崩れ＝落とす★ */
let 窓の崩れ = 0;
function 読む(名) {
  return fs.readFileSync(path.join(根, 'docs/measured', 名), 'utf8')
    .split(/\r?\n/).filter((l) => l && l[0] !== '#').map((l) => {
      const 欄 = l.split('\t');
      const z0 = 欄.find((c) => c.startsWith('z0=')), 型 = 欄.find((c) => c.startsWith('kata='));
      const 残 = 欄.filter((c) => !/^(z0|kata)=/.test(c));
      const 値 = 残[残.length - 1];
      if (!z0 || !型 || (型 === 'kata=Double' && (値 === '0') !== (z0 === 'z0=True'))) { 窓の崩れ++; console.log('✗ ' + 名 + ' 窓の 崩れ ' + l); }
      return 残;
    });
}
const 盤1 = (h) => ['apple', 'Apple', 'ab', 'a*b', 5, null, 'a?c', 15].forEach((v, i) => { if (v !== null) h.打つ('A' + (i + 1), v); h.打つ('B' + (i + 1), i + 1); });
const 盤2 = (h) => {
  ['apple', '=NA()', '=TRUE', '=""', 'a~b', ' x', 'x', 5, '=1/0', null].forEach((v, i) => { if (v !== null) h.打つ('A' + (i + 1), v); h.打つ('B' + (i + 1), i + 1); });
  for (let i = 1; i <= 5; i++) { h.打つ('H' + i, i); h.打つ('I' + i, i * 10); }
  h.打つ('L1', '*');
};
/* 盤4・盤5（紙3）... ★字の '5' は ="5" で 置く（台の 打つ('5') は 数に なる）★
   D系の 行は 3列（式・J2=条件・値）＝J2 に ="条件" を 置く */
const 盤3 = (h, 条) => {
  ['0', '=" "', '16', 'x', '5', '="5"'].forEach((v, i) => h.打つ('A' + (i + 1), v));
  h.打つ('F1', '名'); h.打つ('G1', '数');
  ['apple', '=NA()', '=TRUE', '=""', 'a~b', ' x', 'x', 5, '=1/0', null].forEach((v, i) => { if (v !== null) h.打つ('F' + (i + 2), v); h.打つ('G' + (i + 2), i + 1); });
  h.打つ('J1', '名');
  if (条 !== undefined) h.打つ('J2', '="' + 条 + '"');
};
/* 盤8（紙6）A1:A6=0," ",16,x,5(数),'5(字)・B1:B6=1..6・Z1=5・Z2=TRUE／C1:C3=45000,46000,0.5／E1:E10=盤2 の A と 同じ／F1='a~b */
const 盤8 = (h) => {
  ['0', '=" "', '16', 'x', 5, '="5"'].forEach((v, i) => { h.打つ('A' + (i + 1), v); h.打つ('B' + (i + 1), i + 1); });
  h.打つ('Z1', 5); h.打つ('Z2', '=TRUE');
  [45000, 46000, 0.5].forEach((v, i) => { h.打つ('C' + (i + 1), v); h.打つ('D' + (i + 1), i + 1); });
  ['apple', '=NA()', '=TRUE', '=""', 'a~b', ' x', 'x', 5, '=1/0', null].forEach((v, i) => { if (v !== null) h.打つ('E' + (i + 1), v); });
  h.打つ('F1', 'a~b');
};
/* 盤6（紙4）A1:A7='5,' 5','5.','1E3','1,000',5(数),'-0・B1:B7=1..7（紙4 の 頭の '5, は 区切りの 書き違い＝道具では '5）
   盤7（紙5・千の カンマと > < の 型）A1:A10='1,00,'1,2,3,',5,'1,000.5,'-1,000,'1,0000,TRUE,5(数),'b,'5,・B1:B10=1..10 */
const 字に = (v) => (typeof v === 'string' && v[0] !== '=' ? '="' + v + '"' : v);
const 盤6 = (h) => ['5', ' 5', '5.', '1E3', '1,000', 5, '-0'].forEach((v, i) => { h.打つ('A' + (i + 1), 字に(v)); h.打つ('B' + (i + 1), i + 1); });
const 盤7 = (h) => ['1,00', '1,2,3', ',5', '1,000.5', '-1,000', '1,0000', '=TRUE', 5, 'b', '5,'].forEach((v, i) => { h.打つ('A' + (i + 1), 字に(v)); h.打つ('B' + (i + 1), i + 1); });
/* 盤11（紙7）G1:G10=45658,'2025/1/1,0.5,'12:00,'50%,0,=25/24,=1/24,'25:00,0.05・H1:H10=1,2,4,...,512（SUMIF で どの マスか 分かる） */
const 盤11 = (h) => [45658, '2025/1/1', 0.5, '12:00', '50%', 0, '=25/24', '=1/24', '25:00', 0.05].forEach((v, i) => { h.打つ('G' + (i + 1), 字に(v)); h.打つ('H' + (i + 1), 2 ** i); });
const 紙たち = [['utsushiji-excel-2026-11-07.tsv', 27, 盤1], ['utsushiji-excel-2-2026-11-07.tsv', 39, 盤2], ['utsushiji-excel-3-2026-11-07.tsv', 39, 盤3],
  ['utsushiji-excel-4-2026-11-07.tsv', 9, 盤6], ['utsushiji-excel-5-2026-11-07.tsv', 19, 盤7], ['utsushiji-excel-6-2026-11-07.tsv', 18, 盤8], ['utsushiji-excel-7-2026-11-07.tsv', 48, 盤11], ['utsushiji-excel-8-2026-11-07.tsv', 14, 盤11]];
/* ★「まだ」の 口は 残す（直して いない 行を 黙って 外さず 数えて 出す 為）★ */
const まだ = /^$^/;   /* ★紙6 の 日付・%・時刻の 右 7本と MATCH("a~b") 1本は 直した（11-07）＝まだ 0★ */
let まだの数 = 0;
let pass = 0, fail = 0;
for (const [名, 本数, 盤] of 紙たち) {
  const 紙 = 読む(名);
  if (紙.length !== 本数) { console.log('✗ ' + 名 + ' の 行が ' + 本数 + '本で ない: ' + 紙.length); fail++; }
  for (const 行 of 紙) {
    if (名 === 'utsushiji-excel-6-2026-11-07.tsv' && まだ.test(行[0])) { まだの数++; continue; }
    const 式 = 行[0], 条 = 行.length === 3 ? 行[1].replace(/^J2=/, '') : undefined, 見 = 行[行.length - 1];
    const h = new H.表(); 盤(h, 条); h.打つ('BZ1', 式);
    const 出 = String(h.字('BZ1'));
    const ok = 出 === 見 || (出 !== '' && isFinite(+出) && Math.abs(+出 - +見) < 1e-9);
    if (ok) pass++; else { fail++; console.log('✗ ' + 名 + ' ' + 式 + (条 !== undefined ? ' J2=' + 条 : '') + ' 台 ' + 出 + ' Excel ' + 見); }
  }
}
/* ★盤の 置き方が 効いて いるか★（打つ が 字や 真偽を 別の 型に 変えて いれば 紙と 比べても 意味が 無い）
   ★台の 打つ('TRUE') は 字の まま＝真偽の マスは =TRUE で 置く★（値は Excel の 盤と 同じ 真偽） */
{
  const h = new H.表(); 盤2(h);
  const 型 = ['A2', 'A3', 'A4', 'A6', 'A9'].map((a) => String(h.字(a)));
  const 見 = ['#N/A', 'TRUE', '', ' x', '#DIV/0!'];
  if (JSON.stringify(型) === JSON.stringify(見)) pass++; else { fail++; console.log('✗ 盤2 の 置き方 ' + JSON.stringify(型)); }
}
/* ★MATCH も 同じ 型紙＝`*` を 字より 先に 見る（前は "a*" が "a*b" に 当たらなかった）★ */
{
  const h = new H.表(); h.打つ('A1', 'a*b'); h.打つ('BZ1', '=MATCH("a*",A1:A1,0)');
  if (String(h.字('BZ1')) === '1') pass++; else { fail++; console.log('✗ MATCH("a*",{"a*b"},0) 台 ' + h.字('BZ1') + ' 見 1'); }
}
fail += 窓の崩れ;
console.log('[utsushiji] pass ' + pass + ' / fail ' + fail + '（窓の 崩れ ' + 窓の崩れ + '・まだ ' + まだの数 + '本）');
process.exit(fail ? 1 : 0);
