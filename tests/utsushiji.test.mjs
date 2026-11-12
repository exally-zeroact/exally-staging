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
/* 盤13（紙10・範囲の 引数に #REF!）A1:A5=1..5・B1:B5=10..50 */
const 盤13 = (h) => [1, 2, 3, 4, 5].forEach((v, i) => { h.打つ('A' + (i + 1), v); h.打つ('B' + (i + 1), v * 10); });
/* 盤14（紙11・読めない 参照）A1:A5=1..5・B1:B5=10,20,1,1,50（retsu-x の 材料＋B3=1・B4=1） */
const 盤14 = (h) => [10, 20, 1, 1, 50].forEach((v, i) => { h.打つ('A' + (i + 1), i + 1); h.打つ('B' + (i + 1), v); });
/* 盤15（紙13・関数が 返す 参照）＝盤14＋E1 ="5"（式の 字）・E2=7・C2=TRUE（★台の 打つ('TRUE') は 字＝=TRUE で 置く★） */
const 盤15 = (h) => { 盤14(h); h.打つ('E1', '="5"'); h.打つ('E2', 7); h.打つ('C2', '=TRUE'); };
const 紙たち = [['utsushiji-excel-2026-11-07.tsv', 27, 盤1], ['utsushiji-excel-2-2026-11-07.tsv', 39, 盤2], ['utsushiji-excel-3-2026-11-07.tsv', 39, 盤3],
  ['utsushiji-excel-4-2026-11-07.tsv', 9, 盤6], ['utsushiji-excel-5-2026-11-07.tsv', 19, 盤7], ['utsushiji-excel-6-2026-11-07.tsv', 18, 盤8], ['utsushiji-excel-7-2026-11-07.tsv', 48, 盤11], ['utsushiji-excel-8-2026-11-07.tsv', 14, 盤11],
  ['utsushiji-excel-10-2026-11-07.tsv', 16, 盤13], ['utsushiji-excel-11-2026-11-07.tsv', 14, 盤14],
  ['utsushiji-excel-12-2026-11-08.tsv', 22, 盤14], ['utsushiji-excel-13-2026-11-08.tsv', 103, 盤15],
  ['utsushiji-excel-14-2026-11-08.tsv', 35, 盤15], ['utsushiji-excel-15-2026-11-08.tsv', 27, 盤15], ['utsushiji-excel-16-2026-11-08.tsv', 23, 盤15],
  /* ★件4c（2026-11-08）＝束（とびとび）・端の外枠・帯の交わり・INDEX の区・束の門の名簿★ */
  ['utsushiji-excel-17-2026-11-08.tsv', 30, 盤15], ['utsushiji-excel-18-2026-11-08.tsv', 46, 盤15],
  ['utsushiji-excel-19-2026-11-08.tsv', 124, 盤15], ['utsushiji-excel-20-2026-11-08.tsv', 58, 盤15],
  /* ★叩き直し2（2026-11-08）＝束の 関数の 残り（GCD・IM系・COVAR・SUMX系は #VALUE!／FREQUENCY・ZTEST・HSTACK・VSTACK・TOCOL・TOROW は 受ける）・束と 帯の 交わり★ */
  ['utsushiji-excel-21-2026-11-08.tsv', 30, 盤15],
  /* ★叩き直し3（2026-11-08）＝TOCOL/TOROW の 列の 順・区が 1つ 残る 交わり（22）／束の 交わりを MATCH・XLOOKUP・VLOOKUP・HLOOKUP・INDEX・SUMIF に（23）★ */
  ['utsushiji-excel-22-2026-11-08.tsv', 33, 盤15], ['utsushiji-excel-23-2026-11-08.tsv', 48, 盤15],
  /* ★叩き直し3 の 後（2026-11-08）＝TOCOL/TOROW の 3つ目の 字（「TRUE」だけ 真・他の 字は #VALUE!）／IF・CHOOSE・IFS・SWITCH を 通る 帯（件4d）★ */
  ['utsushiji-excel-24-2026-11-08.tsv', 19, 盤15],
  /* ★件4d（2026-11-08）＝IF・CHOOSE・IFS・SWITCH を 通る 帯・AGGREGATE の 帯（紙25・58式・F/F2 違い 10行は 数えない）★ */
  ['utsushiji-excel-25-2026-11-08.tsv', 48, 盤15]];
/* ★「まだ」の 口は 残す（直して いない 行を 黙って 外さず 数えて 出す 為）★ */
const まだ = /^$^/;   /* ★紙6 の 日付・%・時刻の 右 7本と MATCH("a~b") 1本は 直した（11-07）＝まだ 0★ */
/* ★紙11 の まだ★（2026-11-07）＝参照を 返す IF・CHOOSE・LET を 足す 範囲に 置く 形＝Excel 52
     ★IF・CHOOSE は 値の直し4 件4a（2026-11-08）で 直した★＝残りは LET（件4b＝LET・ラムダの 名が 参照を 持つ） */
const まだ11 = new Set([]);   /* ★LET は 件4b（2026-11-08）で 直した★ */
/* ★紙13 の まだ★（2026-11-08・件4a の 後）＝名指しで 件を 書く（黙って 外さず 数えて 出す）
     件4b＝LET・ラムダの 名が 参照を 持たない（値に 詰める）／件4c＝XLOOKUP が 参照を 返さない
     件5＝式が 作った 表を IF が 1つに 詰める（Formula2 の 溢れ）／件6＝帯を AVERAGEIF(S)・MAXIFS・MINIFS に 置くと #REF! */
const まだ13 = new Map([
  ['=SUMIF(A1:A5,">2",XLOOKUP(10,B1:B5,B1:B5))', '件4c'], ['=SUM(XLOOKUP(10,B1:B5,A1:A5):A3)', '件4c'], ['=SUM(XLOOKUP(20,B1:B5,A1:A5):A5)', '件4c'],
  ['=SUM(IF(TRUE,{1,2,3}))', '件5'], ['=ROWS(IF(TRUE,{1;2;3}))', '件5'],
  ['=AVERAGEIF(A:A,">2",B:B)', '件6'], ['=MAXIFS(B:B,A:A,">2")', '件6'], ['=MINIFS(B:B,A:A,">2")', '件6'], ['=AVERAGEIFS(B:B,A:A,">2")', '件6'],
]);
/* ★紙14 の まだ★（2026-11-08 件4b の 後）＝番号が 並び（{1,2}）の CHOOSE は 並びの 数だけ 選んで 溢れる＝件5 */
const まだ14 = new Map([['=SUM(CHOOSE({1,2},B1,B2))', '件5']]);
/* ★紙17・紙19 の まだ★（2026-11-08 件4c の 後）＝CHOOSE({1,2},B1,B2) は 件5・CELL("address",束) は Excel が "$A$1,$B$1"（台は #VALUE! に 止めて ある＝黙って 1つ目を 出さない） */
const まだ17 = new Map([['=SUM(CHOOSE({1,2},B1,B2))', '件5']]);
const まだ19 = new Map([['=CELL("address",(A1,B1))', '束の 番地の 字']]);
/* ★紙22・紙23 の まだ★＝★素の 帯でも 外れる 形★（MEDIAN(A:A)・SUM(INDEX(B:B,0))・近似の VLOOKUP(3,A:B,2) も 台は #REF!/#VALUE!）＝帯の 置き場（案X）の 外・件4c で ない */
const まだ22 = new Map([['=MEDIAN((A:A,B:B) A:A)', '案X MEDIAN の 帯']]);
const まだ23 = new Map([['=SUM(INDEX((A:A,B:B) B:B,0))', '案X INDEX の 0'], ['=VLOOKUP(3,(A:B,D:D) A:B,2)', '案X 近似の VLOOKUP'],
  ['=SUM(INDEX((A:A,B:B),0,1,2))', '案X INDEX の 0'], ['=MEDIAN((A:A,B:B) A:B)', '案X MEDIAN の 帯'], ['=INDEX((A:A,B:B) A:A,0)', '案X INDEX の 0（溢れ）']]);
/* ★紙24 の まだ★＝★帯が IF・CHOOSE・IFS・SWITCH を 通ると COUNT 等が 黙って 0★（ff1aef5 から 在る・件4c の 叩き直し3 で 見つけた）＝件4d */
/*   ★件4d（2026-11-08）で 10本 全部 直した＝まだ24 は 0★ */
const まだ24 = new Map([]);
/* ★紙25 の まだ★（件4d の 後）＝★どれも 誤りを 出す（黙った 数で ない）★
     並びの 判（{TRUE}・{TRUE,FALSE}）は 件5（演算の 溢れ）・LET は 案X の 外・INDEX(A:A,0) は 案X の 0・1:1 は 置き場 BZ1 が 1行目で 輪（K10 なら 11） */
const まだ25 = new Map([['=SUM(IF({TRUE},A:A))', '件5 並びの 判'], ['=COUNT(IF({TRUE,FALSE},A:A))', '件5 並びの 判'],
  ['=SUM(LET(x,IF(TRUE,A:A),x))', '案X LET'], ['=SUM(INDEX(A:A,0))', '案X INDEX の 0'], ['=SUM(IF(TRUE,1:1))', '置き場 BZ1 が 1行目＝輪']]);
let まだの数 = 0;
let pass = 0, fail = 0;
for (const [名, 本数, 盤] of 紙たち) {
  const 紙 = 読む(名);
  if (紙.length !== 本数) { console.log('✗ ' + 名 + ' の 行が ' + 本数 + '本で ない: ' + 紙.length); fail++; }
  for (const 行 of 紙) {
    if (名 === 'utsushiji-excel-6-2026-11-07.tsv' && まだ.test(行[0])) { まだの数++; continue; }
    if (名 === 'utsushiji-excel-11-2026-11-07.tsv' && まだ11.has(行[0])) { まだの数++; console.log('… まだ ' + 名 + ' ' + 行[0] + ' Excel ' + 行[1]); continue; }
    if (名 === 'utsushiji-excel-13-2026-11-08.tsv' && まだ13.has(行[0])) { まだの数++; console.log('... まだ（' + まだ13.get(行[0]) + '） ' + 名 + ' ' + 行[0] + ' Excel ' + 行[1]); continue; }
    for (const [紙名, ま] of [['utsushiji-excel-17-2026-11-08.tsv', まだ17], ['utsushiji-excel-19-2026-11-08.tsv', まだ19], ['utsushiji-excel-22-2026-11-08.tsv', まだ22], ['utsushiji-excel-23-2026-11-08.tsv', まだ23], ['utsushiji-excel-24-2026-11-08.tsv', まだ24], ['utsushiji-excel-25-2026-11-08.tsv', まだ25]]) {
      if (名 === 紙名 && ま.has(行[0])) { まだの数++; console.log('... まだ（' + ま.get(行[0]) + '） ' + 名 + ' ' + 行[0] + ' Excel ' + 行[1]); 行.まだ = true; }
    }
    if (行.まだ) continue;
    if (名 === 'utsushiji-excel-14-2026-11-08.tsv' && まだ14.has(行[0])) { まだの数++; console.log('... まだ（' + まだ14.get(行[0]) + '） ' + 名 + ' ' + 行[0] + ' Excel ' + 行[1]); continue; }
    const 式 = 行[0], 条 = 行.length === 3 ? 行[1].replace(/^J2=/, '') : undefined;
    /* ★Excel の 真偽は Value2 の 字で True／False（紙13 から）＝台の TRUE／FALSE と 揃えて 比べる★ */
    const 見 = ({ True: 'TRUE', False: 'FALSE' })[行[行.length - 1]] || 行[行.length - 1];
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
/* ★★「# (F to F2 ga chigau)」の 行も 数える★★（件4c 叩き直し2・止め1＝紙19 の FREQUENCY((A1:A2,B1:B2),{2}) が 2→#VALUE! に 崩れたのに 「#」の 行は 飛ばして いた）
     ＝溢れる 答え（F2 は 左上）。台の 字('BZ1') は 左上なので F2 と 比べる。★紙17〜21 だけ★（紙13 以前の F/F2 は 演算子の 溢れ＝件5） */
{
  let 見た = 0;
  for (const k of [17, 18, 19, 20, 21, 22, 23]) {
    const 名 = 'utsushiji-excel-' + k + '-2026-11-08.tsv';
    for (const l of fs.readFileSync(path.join(根, 'docs/measured', 名), 'utf8').split(/\r?\n/)) {
      const m = /^# \(F to F2 ga chigau\) (=.*) F=(.*) F2=(.*)$/.exec(l);
      if (!m) continue;
      見た++;
      if (k === 23 && まだ23.has(m[1])) { まだの数++; console.log('... まだ（' + まだ23.get(m[1]) + '） ' + 名 + ' ' + m[1]); continue; }
      const h = new H.表(); 盤15(h); h.打つ('BZ1', m[1]);
      const 出 = String(h.字('BZ1'));
      if (出 === m[3] || (出 !== '' && isFinite(+出) && Math.abs(+出 - +m[3]) < 1e-9)) pass++; else { fail++; console.log('✗ ' + 名 + ' ' + m[1] + '（溢れの 左上） 台 ' + 出 + ' Excel F2 ' + m[3]); }
    }
  }
  if (見た !== 5) { fail++; console.log('✗ 紙17〜23 の F/F2 の 行が 5本で ない: ' + 見た); }
}
/* ★★紙26（件4f・2026-11-09）＝入れ子の 小計（A20）を 外の SUBTOTAL・AGGREGATE が 外すか★★
     実Excel 16.0 build 20527・toru-utsushiji-26*.ps1・盤 A1:A5=1..5・A20=中（式 または 字のマス TEXTMARK）・K1=外
     ★行の 形が 違う★＝「A20 中<TAB>外<TAB>値<TAB>z0=<TAB>kata=」（窓は ここで 見る＝読む() を 通さない・0 の 行 20本）
     直す前の lib（04af44c）は 155/311・直した後 303/311 */
/* ★まだ は「中<TAB>外」で ちょうど 8本だけ（通る 行は 飛ばさない）★ */
const まだ26 = new Map([
  ...['=SUBTOTAL(9,A1:A30)', '=AGGREGATE(3,0,A1:A30)', '=AGGREGATE(9,0,A1:A30)'].map((外) => ['=ROWS(SUBTOTAL)\t' + 外,
    '名 SUBTOTAL（関数で ない 名）を 台は #NAME?・Excel は 1（範囲外）']),
  ...['=SUBTOTAL(9,A1:A5)', '=IF(FALSE,SUBTOTAL(9,A1:A5),7)', '=CHOOSE(2,SUBTOTAL(9,A1:A5),7)', '=IF(FALSE,AGGREGATE(9,6,A1:A5),7)', 'TEXTMARK']
    .map((中) => [中 + '\t=SUBTOTAL(9,Sheet1!A1:A30)', '板の 名 Sheet1 が 台に 無い＝#REF!（範囲外）'])]);
{
  let 行数 = 0, まだ26の数 = 0;
  const 見本数 = { '': 120, b: 64, c: 25, d: 32, e: 70 };
  for (const x of ['', 'b', 'c', 'd', 'e']) {
    const 名 = 'utsushiji-excel-26' + x + '-2026-11-09.tsv';
    const 行たち = fs.readFileSync(path.join(根, 'docs/measured', 名), 'utf8').split(/\r?\n/).filter((l) => l && l[0] !== '#');
    if (行たち.length !== 見本数[x]) { fail++; console.log('✗ ' + 名 + ' の 行が ' + 見本数[x] + '本で ない: ' + 行たち.length); }
    for (const l of 行たち) {
      行数++;
      const [a, 外, 見, z0, 型] = l.split('\t'); const 中 = a.replace(/^A20 /, '');
      /* ★窓（z0= kata=）が 無い・値と 食い違う 行は 紙の 崩れ★（monosashi-mado の 決まり・読む() と 同じ 判） */
      if (!/^z0=/.test(z0 || '') || !/^kata=/.test(型 || '') || (型 === 'kata=Double' && (!/^z0=(True|False)$/.test(z0) || (見 === '0') !== (z0 === 'z0=True')))) { 窓の崩れ++; console.log('✗ ' + 名 + ' 窓の 崩れ ' + l); continue; }
      const 訳 = まだ26.get(中 + '\t' + 外);
      if (訳) { まだ26の数++; まだの数++; console.log('... まだ（' + 訳 + '） ' + 名 + ' ' + 中 + ' ' + 外 + ' Excel ' + 見); continue; }
      const h = new H.表();
      for (let i = 1; i <= 5; i++) h.打つ('A' + i, String(i));
      h.打つ('A20', 中 === 'TEXTMARK' ? 'shoukei SUBTOTAL(9) zumi' : 中); h.打つ('K1', 外);
      const 出 = String(h.字('K1'));
      if (出 === 見) pass++; else { fail++; console.log('✗ ' + 名 + ' A20 ' + 中 + ' ' + 外 + ' 台 ' + 出 + ' Excel ' + 見); }
    }
  }
  if (行数 !== 311) { fail++; console.log('✗ 紙26 の 行が 311本で ない: ' + 行数); }
  if (まだ26の数 !== まだ26.size) { fail++; console.log('✗ 紙26 の まだ が ' + まだ26.size + '本 当たらない: ' + まだ26の数); }
}
/* ★★紙27（件4g・2026-11-09）★★ 実Excel 16.0 build 20527・toru-utsushiji-27*.ps1
     27  ＝外が 1マスだけ（A20・A20:A20・A1,A20）の SUBTOTAL・AGGREGATE（行の 形は 紙26 と 同じ・618本）
           直す前の lib（36567c4）606/618＝外れ 12本は 全部 AGGREGATE(9,o,A1,A20)
     27b〜27e＝AGGREGATE の 引数の 種（盤 B1:B5=1,2,=1/0,4,5・C1='moji'・D1=TRUE・A1 空・行の 形「-<TAB>式<TAB>値<TAB>z0=<TAB>kata=」）
           27b マス・束・参照を 返す 関数・直の 値（222）／27c 1つ目の 引数と 式の 誤り（58）／27d SUBTOTAL の 同じ 形（105）／27e COUNT・COUNTA と 誤り（110）／27f 縦の 定数の 表・直の #REF!・LET・IF({TRUE})（17＋板を 消した 手順の 注 1行）
           直す前の lib（36567c4）27b 36/222・27c 30/58・27d 60/105・27e 70/110
     ★(ukenai)＝Excel が 式を 受け付けない＝比べない（数だけ 見る）★ */
const まだ27 = new Map([
  ['=AGGREGATE(9,6,B1:B5,Sheet1!B2)', '板の 名 Sheet1 が 台に 無い＝範囲外'], ['=AGGREGATE(9,0,B1,Sheet1!B2)', '板の 名 Sheet1 が 台に 無い＝範囲外'],
  ['=AGGREGATE(9,0,B1,B:B)', '帯の 選び 0〜3 は 台が #REF!（閉じる 側・別の 件）'],
  ...['9', '109', '1', '2', '4'].map((n) => ['=SUBTOTAL(' + n + ',Sheet1!B1)', '板の 名 Sheet1 が 台に 無い＝範囲外'])]);
{
  let まだ27の数 = 0;
  const 見本 = { '': [618, 0], b: [222, 6], c: [58, 4], d: [105, 40], e: [110, 0], f: [17, 0] };
  for (const x of ['', 'b', 'c', 'd', 'e', 'f']) {
    const 名 = 'utsushiji-excel-27' + x + '-2026-11-09.tsv';
    const 行たち = fs.readFileSync(path.join(根, 'docs/measured', 名), 'utf8').split(/\r?\n/).filter((l) => l && l[0] !== '#');
    let 受けない = 0;
    for (const l of 行たち) {
      const [a, 外, 見, z0, 型] = l.split('\t');
      if (見 === '(ukenai)') { 受けない++; continue; }
      if (!/^z0=/.test(z0 || '') || !/^kata=/.test(型 || '') || (型 === 'kata=Double' && (!/^z0=(True|False)$/.test(z0) || (見 === '0') !== (z0 === 'z0=True')))) { 窓の崩れ++; console.log('✗ ' + 名 + ' 窓の 崩れ ' + l); continue; }
      const 訳 = x === '' ? null : まだ27.get(外);
      if (訳) { まだ27の数++; まだの数++; console.log('... まだ（' + 訳 + '） ' + 名 + ' ' + 外 + ' Excel ' + 見); continue; }
      const h = new H.表();
      if (x === '') { for (let i = 1; i <= 5; i++) h.打つ('A' + i, String(i)); h.打つ('A20', a.replace(/^A20 /, '')); }
      else { ['1', '2', '=1/0', '4', '5'].forEach((v, i) => h.打つ('B' + (i + 1), v)); h.打つ('C1', 'moji'); h.打つ('D1', '=TRUE'); if (x === 'f') h.打つ('A20', '=SUBTOTAL(9,B1:B2)'); }
      h.打つ('K1', 外);
      const 出 = String(h.字('K1'));
      const ok = 出 === 見 || (出 !== '' && isFinite(+出) && Math.abs(+出 - +見) < 1e-9);
      if (ok) pass++; else { fail++; console.log('✗ ' + 名 + ' ' + (x === '' ? a + ' ' : '') + 外 + ' 台 ' + 出 + ' Excel ' + 見); }
    }
    if (行たち.length !== 見本[x][0] || 受けない !== 見本[x][1]) { fail++; console.log('✗ ' + 名 + ' の 行 ' + 行たち.length + '（受けない ' + 受けない + '）が 見本 ' + 見本[x] + ' で ない'); }
  }
  /* ★知らない 板（Nai）を 指す 参照は #REF!★（taiketsu 止め S1・実測 紙27f＝板を 消すと 式は #REF!B2 に なり Excel は 選び 6 でも #REF!）
       直す前の lib は (9,6,B1:B5,Nai!B2) 12・(9,6,Nai!B2:B3) 0＝黙った 数 */
  for (const 式 of ['=AGGREGATE(9,6,B1:B5,Nai!B2)', '=AGGREGATE(9,6,(B1,Nai!B1))', '=AGGREGATE(9,6,B1:B5,Nai!B2:B3)', '=AGGREGATE(9,6,Nai!B2)', '=AGGREGATE(14,6,Nai!B2:B3,1)', '=AGGREGATE(2,6,B1:B5,Nai!B2)']) {
    const h = new H.表(); ['1', '2', '=1/0', '4', '5'].forEach((v, i) => h.打つ('B' + (i + 1), v)); h.打つ('K1', 式);
    const 出 = String(h.字('K1'));
    if (出 === '#REF!') pass++; else { fail++; console.log('✗ 知らない 板 ' + 式 + ' 台 ' + 出 + ' 見 #REF!'); }
  }
  if (まだ27の数 !== まだ27.size) { fail++; console.log('✗ 紙27 の まだ が ' + まだ27.size + '本 当たらない: ' + まだ27の数); }
}
/* ★窓の Q10 が 溢れを 塞いだ 行＝z0 が #SPILL!（叩き直し3 で 窓の門を 締めて 見つけた）・答えは 紙28j で 塞がずに 測り直した★ */
const 塞がれた窓 = new Set(['=XLOOKUP(2,B1:B5,C1:D5)']);
/* ★★紙27g・27h（件4g 叩き直し 止め1・2026-11-09）★★ 実Excel 16.0 build 20527・toru-utsushiji-27g/27h.ps1
     盤 B1:B5=1,2,=1/0,4,5・C1:C2=7,1。参照を 返す 関数（XLOOKUP・TAKE・DROP・OFFSET・INDIRECT・SWITCH・IFS・IF・CHOOSE・INDEX）を AGGREGATE に 渡す・
     XLOOKUP の 大きさ違い（#VALUE!）・2列 返す 形・見つからない 時の 4つ目・知らない 板（Nai を 消した 後の 式）と 直の 誤りの 順
     直す前の lib（d71a836）27g 17/33・27h 10/22（ISREF の 3本は 字の 大小だけ＝True／TRUE）。土台（9e5c8d2）27g 19/33
     ★紙27i（叩き直し 止め1）★＝探す 所が 2次元の XLOOKUP は #VALUE!（前は 投げて 表 全体が 読めなかった）・TAKE・DROP の 数を マス・TRUE で 渡しても 参照・4つ目を 省いた XLOOKUP は #N/A。f7d0b59 の lib で 27i 4/21
     ★「(Nai wo kesu mae ...)」の 行★＝板 Nai を 作らずに 元の 式を 打つ（台では 知らない 板）＝Excel の 消した 後の 答えと 比べる */
const まだ27g = new Map([
  ['=AGGREGATE(9,6,FILTER(B1:B5,B1:B5<>2))', 'FILTER の 条件に 誤りが 在る 時 Excel は FILTER が #DIV/0!（台は 誤りを 出さない＝AGGREGATE の 外の 別の 件）'],
  ['=AGGREGATE(9,6,TRIMRANGE(B1:B9))', 'TRIMRANGE は 台に 無い（#NAME?）＝別の 件'],
  ['=AGGREGATE(9,6,B1:B9.)', '切り詰めの 参照 B1:B9. を 台は 読まない＝別の 件']]);
{
  let まだ27gの数 = 0;
  const 見本 = { g: 33, h: 22, i: 21 };
  for (const x of ['g', 'h', 'i']) {
    const 名 = 'utsushiji-excel-27' + x + '-2026-11-09.tsv';
    const 行たち = fs.readFileSync(path.join(根, 'docs/measured', 名), 'utf8').split(/\r?\n/).filter((l) => l && l[0] !== '#');
    for (const l of 行たち) {
      const [, 外0, 見, z0, 型] = l.split('\t');
      if (塞がれた窓.has(外0)) { まだの数++; console.log('... まだ（窓の Q10 が 溢れを 塞いだ＝z0 #SPILL!・紙28j で 塞がずに 測り直した） ' + 名 + ' ' + 外0); continue; }
      if (!/^z0=/.test(z0 || '') || !/^kata=/.test(型 || '') || (型 === 'kata=Double' && (!/^z0=(True|False)$/.test(z0) || (見 === '0') !== (z0 === 'z0=True')))) { 窓の崩れ++; console.log('✗ ' + 名 + ' 窓の 崩れ ' + l); continue; }
      const m = 外0.match(/^(\S+) \(Nai wo kesu mae /);
      const 外 = m ? m[1] : 外0;
      const 訳 = まだ27g.get(外);
      if (訳) { まだ27gの数++; まだの数++; console.log('... まだ（' + 訳 + '） ' + 名 + ' ' + 外 + ' Excel ' + 見); continue; }
      const h = new H.表();
      ['1', '2', '=1/0', '4', '5'].forEach((v, i) => h.打つ('B' + (i + 1), v)); h.打つ('C1', '7'); h.打つ('C2', '1');
      h.打つ('K1', 外);
      const 出 = String(h.字('K1'));
      const 見U = 型 === 'kata=Boolean' ? 見.toUpperCase() : 見;
      const ok = 出 === 見U || (出 !== '' && isFinite(+出) && Math.abs(+出 - +見) < 1e-9);
      if (ok) pass++; else { fail++; console.log('✗ ' + 名 + ' ' + 外 + ' 台 ' + 出 + ' Excel ' + 見); }
    }
    if (行たち.length !== 見本[x]) { fail++; console.log('✗ ' + 名 + ' の 行 ' + 行たち.length + ' が 見本 ' + 見本[x] + ' で ない'); }
  }
  if (まだ27gの数 !== まだ27g.size) { fail++; console.log('✗ 紙27g の まだ が ' + まだ27g.size + '本 当たらない: ' + まだ27gの数); }
}
/* ★紙28g の まだ＝台は 閉じる 側の 誤りを 出す（黙った 数・#N/A に しない）＝[訳, 台の 答え]★ */
const まだ28 = new Map([
  ['=XMATCH(4,INDEX(B:C,0,1))', ['INDEX(帯,0,k) は 帯を 置かない＝別の 件 4j', '#REF!']],
  ['=XLOOKUP(4,INDEX(B:C,0,1),C:C,"n")', ['INDEX(帯,0,k) は 帯を 置かない＝別の 件 4j', '#REF!']],
  ['=XMATCH("(?>ab)c",G1:G3,3)', ['JS に 無い (?> ＝戻りで 答えが 変わり得る', '#VALUE!']],
  ['=XMATCH("ab*+",G1:G3,3)', ['JS に 無い 所有の *+ ＝戻りで 答えが 変わり得る', '#VALUE!']],
  ['=XLOOKUP(1,B1,B:B,"n")', ['Excel は 返す 所を 丸ごと 溢れさせ #SPILL!（z0=-2146826243・紙の 字は 空）＝台は 丸ごとを 溢れさせ られない', '#REF!']],
  ['=XLOOKUP(1,B1,C:C,"n")', ['Excel は 返す 所を 丸ごと 溢れさせ #SPILL!（z0=-2146826243・紙の 字は 空）＝台は 丸ごとを 溢れさせ られない', '#REF!']],
  ['=XLOOKUP(1,B1,1:1,"n")', ['Excel は 返す 所を 丸ごと 溢れさせ #SPILL!（z0=-2146826243・紙の 字は 空）＝台は 丸ごとを 溢れさせ られない', '#REF!']],
  ['=XLOOKUP(1,1,B:B,"n")', ['Excel は 返す 所を 丸ごと 溢れさせ #SPILL!（z0=-2146826243・紙の 字は 空）＝台は 丸ごとを 溢れさせ られない', '#REF!']],
  ['=XLOOKUP(7,C1,B:B,"n")', ['Excel は 返す 所を 丸ごと 溢れさせ #SPILL!（z0=-2146826243・紙の 字は 空）＝台は 丸ごとを 溢れさせ られない', '#REF!']],
  ['=XMATCH(2,IF(B1:B2>0,B1:B2))', ['IF が 返す 並びを 探す 所に できない＝9e769fdc から #VALUE!・別の 件（紙28h）', '#VALUE!']]]);
let まだ28の数 = 0;
/* ★kata=Double なら z0 は True／False だけ★（叩き直し3＝前は 紙28i の C1:D1・{5,6} の z0 #SPILL!（窓の Q10 が 溢れを 塞いだ）を 素通しして いた・その 2行は 28i から 外し 28j で 測った） */
/* ★★紙28・28b・28c・28d・28e・28f・28g・28h（件4h・2026-11-09）★★ 実Excel 16.0 build 20527・toru-utsushiji-28〜28h.ps1
     ★紙28i（叩き直し2）★＝探す 所が 1マスなら 大きさを 見ない（見つからなければ n）・返す 所が 2次元の 帯は #VALUE!
     ★紙28h（叩き直し2 止め1）★＝探す 所が 1マスの 参照（S1・(S1)・B3・S1:S1・OFFSET・INDIRECT）なら 誤りを データと して 探す（n・#N/A）・計算した 値（1/0・S1*1）だけ その 誤り
     XLOOKUP の 5つ目（1・-1・2・3＝正規表現）と 6つ目（-1・±2）・XMATCH の 同じ 道・並びの 位（数＜字＜真偽＜誤り・二分では 空も 一番 上）・
     丸ごとの 2次元（B:C・1:2）の XLOOKUP は #VALUE!・TAKE／DROP の 数が TRUE の マス。9e769fdc の lib で 赤を 見てから 直した
     ★盤の 真偽は =TRUE／=FALSE で 置く★（台の 打つ('TRUE') は 字＝件4i の 別の 穴）・式は P1 に 打つ（盤が K 列まで 在る）・★紙28f・28g は 式を P10・盤に S1:S2=#DIV/0!,#N/A を 足す★（丸ごとの 列・行 1:2 を 読むので 1行目に 置くと 輪） */
{
  const 見本 = { '28': 22, '28b': 20, '28c': 33, '28d': 24, '28e': 10, '28f': 27, '28g': 27, '28h': 18, '28i': 16 };
  for (const x of Object.keys(見本)) {
    const 名 = 'utsushiji-excel-' + x + '-2026-11-09.tsv';
    const 行たち = fs.readFileSync(path.join(根, 'docs/measured', 名), 'utf8').split(/\r?\n/).filter((l) => l && l[0] !== '#');
    for (const l of 行たち) {
      const [, 外, 見, z0, 型] = l.split('\t');
      if (!/^z0=/.test(z0 || '') || !/^kata=/.test(型 || '') || (型 === 'kata=Double' && (!/^z0=(True|False)$/.test(z0) || (見 === '0') !== (z0 === 'z0=True')))) { 窓の崩れ++; console.log('✗ ' + 名 + ' 窓の 崩れ ' + l); continue; }
      const h = new H.表();
      ['1', '2', '=1/0', '4', '5'].forEach((v, i) => h.打つ('B' + (i + 1), v)); h.打つ('C1', '7'); h.打つ('C2', '1');
      h.打つ('F1', '=TRUE'); ['x', 'abc', 'ab'].forEach((v, i) => h.打つ('G' + (i + 1), v));
      ['2', 'a', '4', '2', 'c', '=TRUE'].forEach((v, i) => { h.打つ('H' + (i + 1), v); h.打つ('I' + (i + 1), String(10 * (i + 1))); });
      ['2', '=NA()', 'b', '=TRUE'].forEach((v, i) => { h.打つ('K' + (i + 1), v); h.打つ('L' + (i + 1), String(10 * (i + 1))); });
      h.打つ('M1', '1'); h.打つ('M3', '5'); [1, 2, 3].forEach((i) => h.打つ('N' + i, String(10 * i)));
      h.打つ('O1', 'a'); h.打つ('O3', 'c'); h.打つ('R1', '=FALSE'); h.打つ('R3', '=TRUE');
      const 後の紙 = x === '28f' || x === '28g' || x === '28h' || x === '28i';
      const 式の所 = 後の紙 ? 'P10' : 'P1';
      if (後の紙) { h.打つ('S1', '=1/0'); h.打つ('S2', '=NA()'); }
      let 出;
      try { h.打つ(式の所, 外); 出 = String(h.字(式の所)); } catch (e) { 出 = '投げた ' + e.message; }
      const ま = まだ28.get(外);
      if (ま) { まだ28の数++; まだの数++; if (出 === ま[1]) { pass++; console.log('... まだ（' + ま[0] + '） ' + 名 + ' ' + 外 + ' 台 ' + 出 + ' Excel ' + 見); } else { fail++; console.log('✗ まだ の 台が ' + ま[1] + ' で ない ' + 外 + ' 台 ' + 出); } continue; }
      const 見U = 型 === 'kata=Boolean' ? 見.toUpperCase() : 見;
      const ok = 出 === 見U || (出 !== '' && isFinite(+出) && Math.abs(+出 - +見) < 1e-9);
      if (ok) pass++; else { fail++; console.log('✗ ' + 名 + ' ' + 外 + ' 台 ' + 出 + ' Excel ' + 見); }
    }
    if (行たち.length !== 見本[x]) { fail++; console.log('✗ ' + 名 + ' の 行 ' + 行たち.length + ' が 見本 ' + 見本[x] + ' で ない'); }
  }
}
/* ★★紙28j（件4h 叩き直し3・2026-11-09）★★ 実Excel 16.0 build 20527・toru-utsushiji-28j.ps1
     探す 所が 1マスの XLOOKUP＝式は 1本ずつ 新しい 板の P10（隣を 塞がない）・HasSpill の 範囲と P11・Q10 まで 取った
     返す 所が 2次元は #VALUE!・1次元は 当たれば 丸ごと 溢れ（P11／Q10 の 数まで 比べる）・盤の 端を 越えるのは #SPILL!（台は #REF!＝閉じる 側） */
{
  const 名 = 'utsushiji-excel-28j-2026-11-09.tsv';
  const 行たち = fs.readFileSync(path.join(根, 'docs/measured', 名), 'utf8').split(/\r?\n/).filter((l) => l && l[0] !== '#');
  const 溢れ越え = new Set(['=XLOOKUP(1,B1,B1:B1048576,"n")', '=XLOOKUP(1,B1,B:B,"n")']);
  let 越えの数 = 0, 空の数 = 0;
  for (const l of 行たち) {
    const [, 外, 見, 溢, p11, q10, 型] = l.split('\t');
    if (!/^afure=/.test(溢 || '') || !/^P11=/.test(p11 || '') || !/^Q10=/.test(q10 || '') || !/^kata=/.test(型 || '')) { 窓の崩れ++; console.log('✗ ' + 名 + ' 窓の 崩れ ' + l); continue; }
    const h = new H.表();
    ['1', '2', '=1/0', '4', '5'].forEach((v, i) => h.打つ('B' + (i + 1), v)); h.打つ('C1', '7'); h.打つ('C2', '1'); h.打つ('S1', '=1/0');
    let 出, 出11 = '', 出Q = '';
    try { h.打つ('P10', 外); 出 = String(h.字('P10')); 出11 = String(h.字('P11')); 出Q = String(h.字('Q10')); } catch (e) { 出 = '投げた ' + e.message; }
    if (溢れ越え.has(外)) {
      越えの数++; まだの数++;
      if (見 === '#SPILL!' && 出 === '#REF!') { pass++; console.log('... まだ（盤の 端を 越える 溢れ＝Excel #SPILL!・台は 閉じる 側 #REF!） ' + 名 + ' ' + 外); } else { fail++; console.log('✗ まだ の 台が #REF! で ない ' + 外 + ' 台 ' + 出 + ' Excel ' + 見); }
      continue;
    }
    const 同じ = (a, b) => a === b || (a !== '' && b !== '' && isFinite(+a) && isFinite(+b) && Math.abs(+a - +b) < 1e-9);
    const 隣 = (v) => (v === '(sora)' ? '' : v);
    const ok = 同じ(出, 見) && 同じ(出11, 隣(p11.slice(4))) && 同じ(出Q, 隣(q10.slice(4)));
    /* ★溢れの 空★＝Excel は 空の マスを 0 で 溢れさせ、台は 空で 出す（=C1:D1 も 同じ＝XLOOKUP の 外の 別の 件）＝その 1つ だけが 違う 時は まだ（名指し） */
    const 空だけ = !ok && 同じ(出, 見) && (同じ(出11, 隣(p11.slice(4))) || (出11 === '' && p11 === 'P11=0')) && (同じ(出Q, 隣(q10.slice(4))) || (出Q === '' && q10 === 'Q10=0'));
    /* ★空だけ の 時は 溢れが 本当に 届いて いるかを 見る★（叩き直し4 止め2）＝空の 元の マスに 99 を 置いた 盤で その 隣に 99 が 出る（溢れが 丸ごと 消える 壊れを 飲まない） */
    const 空の元 = { '=XLOOKUP(1,B1,C1:D1,"n")': ['D1', 'Q10'], '=XLOOKUP(2,B1:B5,C1:D5)': ['D2', 'Q10'] }[外];
    if (空だけ && 空の元) { const h2 = new H.表(); ['1', '2', '=1/0', '4', '5'].forEach((v, i) => h2.打つ('B' + (i + 1), v)); h2.打つ('C1', '7'); h2.打つ('C2', '1'); h2.打つ(空の元[0], '99'); h2.打つ('P10', 外); if (String(h2.字(空の元[1])) !== '99') { fail++; console.log('✗ 紙28j 溢れが ' + 空の元[1] + ' に 届かない ' + 外 + ' 台 ' + h2.字(空の元[1])); continue; } }
    else if (空だけ) { fail++; console.log('✗ 紙28j 空だけ だが 空の元 が 名指し されて ない ' + 外); continue; }
    if (空だけ) { まだの数++; 空の数++; pass++; console.log('... まだ（溢れの 空は Excel 0・台 空＝別の 件） ' + 名 + ' ' + 外); continue; }
    if (ok) pass++; else { fail++; console.log('✗ ' + 名 + ' ' + 外 + ' 台 ' + 出 + '・P11 ' + 出11 + '・Q10 ' + 出Q + ' Excel ' + 見 + '・' + p11 + '・' + q10); }
  }
  if (行たち.length !== 19) { fail++; console.log('✗ ' + 名 + ' の 行 ' + 行たち.length + ' が 見本 19 で ない'); }
  if (空の数 !== 2) { fail++; console.log('✗ 紙28j の 溢れの 空 が 2本 で ない: ' + 空の数); }
  if (越えの数 !== 溢れ越え.size) { fail++; console.log('✗ 紙28j の まだ が ' + 溢れ越え.size + '本 当たらない: ' + 越えの数); }
}
/* ★★紙28l・28m（件4h 設計を 戻した 後の 設計前 taiketsu 止め1・2／叩き直し1 止め1・2・2026-11-09）★★ 実Excel 16.0 build 20527・toru-utsushiji-28l.ps1・28m.ps1
     式は 1本ずつ 新しい 板の 行の 1つ目の マス（P10・E2・E3）・★置いた 所＋下＋右★まで 比べる（P10 の 値だけだと {8;9}・{8,9} の 3本が 空振り＝叩き直し1 止め2）
     見つからない 時の 4つ目は 丸ごと（範囲は 参照・配列は 溢れ・形不問・置いた 行に 関わらない）・盤の 端を 越える 溢れは Excel #SPILL!＝台は #REF!（閉じる 側・SUM の 中は Excel 値＝別件 4n） */
{
  const 越え = new Set(['=XLOOKUP(2,B1:B5,C1:XFD5)', '=SUM(XLOOKUP(2,B1:B5,C1:XFD5))', '=XLOOKUP(2,B1:B5,C1:XFC5)',
    '=XLOOKUP(99,B1:B5,C1:C5,B1:XFD1)', '=SUM(XLOOKUP(99,B1:B5,C1:C5,B1:XFD1))', '=XLOOKUP(99,B1:B5,C1:C5,B1:B1048576)', '=SUM(XLOOKUP(99,B1:B5,C1:C5,B1:B1048576))']);
  const 列名 = (n) => { let s = ''; n++; while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s; };
  const 列番 = (s) => { let n = 0; for (const c of s) n = n * 26 + (c.charCodeAt(0) - 64); return n - 1; };
  const ずらす = (a, dr, dc) => { const m = /^([A-Z]+)(\d+)$/.exec(a); return 列名(列番(m[1]) + dc) + (+m[2] + dr); };
  const 同じ = (a, b) => a === b || (/^(TRUE|FALSE)$/i.test(a) && a.toUpperCase() === String(b).toUpperCase()) || (a !== '' && b !== '' && isFinite(+a) && isFinite(+b) && Math.abs(+a - +b) < 1e-9);
  const 隣 = (v) => (v === '(sora)' ? '' : v);
  /* ★紙28n（staging 前 taiketsu 止め1・2026-11-09・toru-utsushiji-28n.ps1・盤に H1:H6=2,a,4,2,c,TRUE を 足す）★
       XLOOKUP の 5つ目・6つ目／XMATCH の 3つ目・4つ目を 省いた 時は 既定（前は 6つ目の 空＝0 で #VALUE!）
       ★まだ★ IFERROR(XLOOKUP(99,...,B1:XFD1),"x") は Excel #SPILL!（IFERROR は 越えの #SPILL! を 拾わない）＝台は "x"（閉じる 側の #REF! を IFERROR が 拾う・別件 4n） */
  const IFERRORのまだ = '=IFERROR(XLOOKUP(99,B1:B5,C1:C5,B1:XFD1),"x")';
  let 越えの数 = 0, IFERRORの数 = 0;
  for (const [名, 本] of [['utsushiji-excel-28l-2026-11-09.tsv', 12], ['utsushiji-excel-28m-2026-11-09.tsv', 6], ['utsushiji-excel-28n-2026-11-09.tsv', 10]]) {
    const 行たち = fs.readFileSync(path.join(根, 'docs/measured', 名), 'utf8').split(/\r?\n/).filter((l) => l && l[0] !== '#');
    for (const l of 行たち) {
      const [置, 外, 見, 溢, 下, 右, 型] = l.split('\t');
      if (!/^[A-Z]+\d+$/.test(置 || '') || !/^afure=/.test(溢 || '') || !/^shita=/.test(下 || '') || !/^migi=/.test(右 || '') || !/^kata=/.test(型 || '')) { 窓の崩れ++; console.log('✗ ' + 名 + ' 窓の 崩れ ' + l); continue; }
      const h = new H.表();
      ['1', '2', '=1/0', '4', '5'].forEach((v, i) => h.打つ('B' + (i + 1), v)); h.打つ('C1', '7'); h.打つ('C2', '1'); h.打つ('XFD2', '9');
      if (名.includes('-28n-')) ['2', 'a', '4', '2', 'c', 'TRUE'].forEach((v, i) => h.打つ('H' + (i + 1), v));
      let 出, 出下 = '', 出右 = '';
      try { h.打つ(置, 外); 出 = String(h.字(置)); 出下 = String(h.字(ずらす(置, 1, 0))); 出右 = String(h.字(ずらす(置, 0, 1))); } catch (e) { 出 = '投げた ' + e.message; }
      if (越え.has(外)) {
        越えの数++; まだの数++;
        if (出 === '#REF!') { pass++; console.log('... まだ（盤の 端を 越える＝Excel ' + 見 + '・台は 閉じる 側 #REF!・別件 4n） ' + 名 + ' ' + 外); } else { fail++; console.log('✗ まだ の 台が #REF! で ない ' + 名 + ' ' + 外 + ' 台 ' + 出 + ' Excel ' + 見); }
        continue;
      }
      if (外 === IFERRORのまだ) {
        IFERRORの数++; まだの数++;
        if (出 === 'x') { pass++; console.log('... まだ（IFERROR が 越えを 拾う＝Excel ' + 見 + '・台 x・別件 4n） ' + 名 + ' ' + 外); } else { fail++; console.log('✗ まだ の 台が x で ない ' + 名 + ' ' + 外 + ' 台 ' + 出 + ' Excel ' + 見); }
        continue;
      }
      const ok = 同じ(出, 見) && 同じ(出下, 隣(下.slice(6))) && 同じ(出右, 隣(右.slice(5)));
      if (ok) pass++; else { fail++; console.log('✗ ' + 名 + ' ' + 置 + ' ' + 外 + ' 台 ' + 出 + '・下 ' + 出下 + '・右 ' + 出右 + ' Excel ' + 見 + '・' + 下 + '・' + 右); }
    }
    if (行たち.length !== 本) { fail++; console.log('✗ ' + 名 + ' の 行 ' + 行たち.length + ' が 見本 ' + 本 + ' で ない'); }
  }
  if (越えの数 !== 越え.size) { fail++; console.log('✗ 紙28l・28m の まだ が ' + 越え.size + '本 当たらない: ' + 越えの数); }
  if (IFERRORの数 !== 1) { fail++; console.log('✗ 紙28n の IFERROR の まだ が 1本 当たらない: ' + IFERRORの数); }
}
/* ★★紙28k（件4h 叩き直し4 止め1・2026-11-09）★★ 実Excel 16.0 build 20527・toru-utsushiji-28k.ps1（28j と 同じ 置き方）
     1マスの 探す 所で 丸ごと 溢れた 物は 参照（ISREF TRUE・AGGREGATE が 範囲で 受ける・ROW 3,4,5・SUM(...:C5) 8）・配列定数は 参照で ない */
{
  const 名 = 'utsushiji-excel-28k-2026-11-09.tsv';
  let まだ28kの数 = 0;
  const 行たち = fs.readFileSync(path.join(根, 'docs/measured', 名), 'utf8').split(/\r?\n/).filter((l) => l && l[0] !== '#');
  for (const l of 行たち) {
    const [, 外, 見, 溢, p11, q10, 型] = l.split('\t');
    if (!/^afure=/.test(溢 || '') || !/^P11=/.test(p11 || '') || !/^Q10=/.test(q10 || '') || !/^kata=/.test(型 || '')) { 窓の崩れ++; console.log('✗ ' + 名 + ' 窓の 崩れ ' + l); continue; }
    const h = new H.表();
    ['1', '2', '=1/0', '4', '5'].forEach((v, i) => h.打つ('B' + (i + 1), v)); h.打つ('C1', '7'); h.打つ('C2', '1'); h.打つ('S1', '=1/0');
    let 出, 出11 = '', 出Q = '';
    try { h.打つ('P10', 外); 出 = String(h.字('P10')); 出11 = String(h.字('P11')); 出Q = String(h.字('Q10')); } catch (e) { 出 = '投げた ' + e.message; }
    const 大 = (v) => (型 === 'kata=Boolean' ? String(v).toUpperCase() : v);
    /* ★まだ（名指し）★＝範囲の「:」が 溢れの 参照の 大きさを 見ない（左上 だけ 見る 作りで 無く #REF!）＝閉じる 側・別の 件（4m 候補）・HEAD は 1マスを 返して 8 だった */
    if (外 === '=SUM(XLOOKUP(1,B1,C1:C2):C5)') { まだの数++; まだ28kの数++; if (出 === '#REF!') { pass++; console.log('... まだ（「:」が 溢れの 参照の 大きさを 見ない＝別の 件） ' + 名 + ' ' + 外 + ' Excel ' + 見); } else { fail++; console.log('✗ まだ の 台が #REF! で ない ' + 外 + ' 台 ' + 出); } continue; }
    const 同じ = (a, b) => a === b || (a !== '' && b !== '' && isFinite(+a) && isFinite(+b) && Math.abs(+a - +b) < 1e-9);
    const 隣 = (v) => (v === '(sora)' ? '' : v);
    if (同じ(出, 大(見)) && 同じ(出11, 隣(p11.slice(4))) && 同じ(出Q, 隣(q10.slice(4)))) pass++; else { fail++; console.log('✗ ' + 名 + ' ' + 外 + ' 台 ' + 出 + '・P11 ' + 出11 + '・Q10 ' + 出Q + ' Excel ' + 見 + '・' + p11 + '・' + q10); }
  }
  if (まだ28kの数 !== 1) { fail++; console.log('✗ 紙28k の まだ が 1本 で ない: ' + まだ28kの数); }
  if (行たち.length !== 10) { fail++; console.log('✗ ' + 名 + ' の 行 ' + 行たち.length + ' が 見本 10 で ない'); }
}
if (まだ28の数 !== まだ28.size) { fail++; console.log('✗ 紙28 の まだ が ' + まだ28.size + '本 当たらない: ' + まだ28の数); }
fail += 窓の崩れ;
console.log('[utsushiji] pass ' + pass + ' / fail ' + fail + '（窓の 崩れ ' + 窓の崩れ + '・まだ ' + まだの数 + '本）');
process.exit(fail ? 1 : 0);
