/* shiki-wo-tsukuru2.mjs — ★予測・統計／表・情報 の 式を 作る★（2026-09-07）
 *
 *  ★★まず「本当に 動いていないか」を 疑う★★
 *    `exally-missing-2026-09-07.txt` に 載っている 16個の うち
 *    ★CONVERT は 本番で 動いている★のを 私は 別の 用で 見ています。
 *    ⇒ 台帳は ★決まった 引数の 形で 押した 結果★なので
 *      ★形が 合わないだけで「動かない」に 見える★事が 有る。
 *    ⇒★★だから 先に「本物の 引数」で 押し直す★★
 *
 *  ★配列は Excel の 書き方で 書く★（{1;2;3} ＝縦, {1,2,3} ＝横）
 *    ⇒ 実Excel も エンジンも 同じ 字を 読む
 *
 *  使い方: node docs/measured/kansuu46/shiki-wo-tsukuru2.mjs
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ここ = path.dirname(fileURLToPath(import.meta.url));

/* ★縦の 並び★ … y は x の 2倍＋1（ぴったり 直線＝答えが 読みやすい） */
const Y = '{3;5;7;9;11;13}';
const X = '{1;2;3;4;5;6}';
/* ★増える 並び★（GROWTH・LOGEST 用）＝2の 累乗 */
const YE = '{2;4;8;16;32;64}';
/* ★ばらついた 並び★（ぴったりでは 分からない 所を 見る） */
const YB = '{3;5;6;10;10;14}';
/* ★時系列★（FORECAST.ETS 用）＝月ごと・季節 4 */
const TL = '{45292;45323;45352;45383;45413;45444;45474;45505;45536;45566;45597;45627}';
const VS = '{10;12;16;14;11;13;17;15;12;14;18;16}';

const 表 = [
  /* ── 直線・増える ── */
  ['TREND', 'TREND(' + Y + ')'],
  ['TREND', 'INDEX(TREND(' + Y + '),1,1)'],
  ['TREND', 'INDEX(TREND(' + Y + '),6,1)'],
  ['TREND', 'INDEX(TREND(' + Y + ',' + X + ',{7;8}),1,1)'],
  ['TREND', 'INDEX(TREND(' + Y + ',' + X + ',{7;8}),2,1)'],
  ['TREND', 'INDEX(TREND(' + YB + ',' + X + ',{7}),1,1)'],
  ['TREND', 'INDEX(TREND(' + YB + ',' + X + ',{7},FALSE),1,1)'],
  ['TREND', 'ROWS(TREND(' + Y + '))'],
  ['GROWTH', 'INDEX(GROWTH(' + YE + '),1,1)'],
  ['GROWTH', 'INDEX(GROWTH(' + YE + '),6,1)'],
  ['GROWTH', 'INDEX(GROWTH(' + YE + ',' + X + ',{7;8}),1,1)'],
  ['GROWTH', 'INDEX(GROWTH(' + YE + ',' + X + ',{7;8}),2,1)'],
  ['GROWTH', 'INDEX(GROWTH(' + YB + ',' + X + ',{7}),1,1)'],
  ['GROWTH', 'INDEX(GROWTH(' + YB + ',' + X + ',{7},FALSE),1,1)'],
  ['LOGEST', 'INDEX(LOGEST(' + YE + ',' + X + '),1,1)'],
  ['LOGEST', 'INDEX(LOGEST(' + YE + ',' + X + '),1,2)'],
  ['LOGEST', 'INDEX(LOGEST(' + YB + ',' + X + '),1,1)'],
  ['LOGEST', 'INDEX(LOGEST(' + YB + ',' + X + '),1,2)'],
  ['LOGEST', 'INDEX(LOGEST(' + YB + ',' + X + ',TRUE,TRUE),3,1)'],
  ['LOGEST', 'INDEX(LOGEST(' + YB + ',' + X + ',TRUE,TRUE),4,1)'],
  ['LOGEST', 'ROWS(LOGEST(' + YB + ',' + X + ',TRUE,TRUE))'],
  ['LOGEST', 'COLUMNS(LOGEST(' + YB + ',' + X + ',TRUE,TRUE))'],
  /* ── 誤差関数 ── */
  ['ERF.PRECISE', 'ERF.PRECISE(0)'],
  ['ERF.PRECISE', 'ERF.PRECISE(0.5)'],
  ['ERF.PRECISE', 'ERF.PRECISE(1)'],
  ['ERF.PRECISE', 'ERF.PRECISE(2)'],
  ['ERF.PRECISE', 'ERF.PRECISE(-1)'],
  ['ERF.PRECISE', 'ERF.PRECISE(3.5)'],
  ['ERFC.PRECISE', 'ERFC.PRECISE(0)'],
  ['ERFC.PRECISE', 'ERFC.PRECISE(0.5)'],
  ['ERFC.PRECISE', 'ERFC.PRECISE(1)'],
  ['ERFC.PRECISE', 'ERFC.PRECISE(2)'],
  ['ERFC.PRECISE', 'ERFC.PRECISE(-1)'],
  ['ERFC.PRECISE', 'ERFC.PRECISE(3.5)'],
  /* ── 予測（ETS） ── */
  ['FORECAST.ETS', 'FORECAST.ETS(45658,' + VS + ',' + TL + ')'],
  ['FORECAST.ETS', 'FORECAST.ETS(45658,' + VS + ',' + TL + ',4)'],
  ['FORECAST.ETS', 'FORECAST.ETS(45658,' + VS + ',' + TL + ',1)'],
  ['FORECAST.ETS.SEASONALITY', 'FORECAST.ETS.SEASONALITY(' + VS + ',' + TL + ')'],
  ['FORECAST.ETS.CONFINT', 'FORECAST.ETS.CONFINT(45658,' + VS + ',' + TL + ')'],
  ['FORECAST.ETS.STAT', 'FORECAST.ETS.STAT(' + VS + ',' + TL + ',1)'],
  ['FORECAST.ETS.STAT', 'FORECAST.ETS.STAT(' + VS + ',' + TL + ',2)'],
  ['FORECAST.ETS.STAT', 'FORECAST.ETS.STAT(' + VS + ',' + TL + ',3)'],
  ['FORECAST.ETS.STAT', 'FORECAST.ETS.STAT(' + VS + ',' + TL + ',4)'],
  ['FORECAST.ETS.STAT', 'FORECAST.ETS.STAT(' + VS + ',' + TL + ',5)'],
  ['FORECAST.ETS.STAT', 'FORECAST.ETS.STAT(' + VS + ',' + TL + ',6)'],
  ['FORECAST.ETS.STAT', 'FORECAST.ETS.STAT(' + VS + ',' + TL + ',7)'],
  ['FORECAST.ETS.STAT', 'FORECAST.ETS.STAT(' + VS + ',' + TL + ',8)'],
  /* ── 表・情報 ── */
  ['AREAS', 'AREAS(A1:B3)'],
  ['AREAS', 'AREAS((A1:B3,D1:D2))'],
  ['AREAS', 'AREAS((A1,B2,C3,D4))'],
  ['CELL', 'CELL("row",B7)'],
  ['CELL', 'CELL("col",C2)'],
  ['CELL', 'CELL("address",B7)'],
  ['CELL', 'CELL("width",B7)'],
  ['CELL', 'CELL("type",B7)'],
  ['CELL', 'CELL("prefix",B7)'],
  ['CELL', 'CELL("protect",B7)'],
  ['CELL', 'CELL("format",B7)'],
  ['CONVERT', 'CONVERT(1,"m","cm")'],
  ['CONVERT', 'CONVERT(1,"lbm","kg")'],
  ['CONVERT', 'CONVERT(100,"C","F")'],
  ['CONVERT', 'CONVERT(1,"day","hr")'],
  ['CONVERT', 'CONVERT(1,"kibyte","byte")'],
  ['CONVERT', 'CONVERT(1,"あ","cm")'],
  ['FILTERXML', 'FILTERXML("<a><b>1</b><b>2</b></a>","//b[1]")'],
  ['FILTERXML', 'FILTERXML("<a><b>1</b><b>2</b></a>","//b[2]")'],
  ['FILTERXML', 'COUNTA(FILTERXML("<a><b>1</b><b>2</b></a>","//b"))'],
  ['INFO', 'INFO("numfile")'],
  ['INFO', 'INFO("recalc")'],
  ['INFO', 'INFO("release")'],
  ['INFO', 'INFO("system")'],
  ['INFO', 'INFO("osversion")'],
  ['RANDARRAY', 'ROWS(RANDARRAY(3,2))'],
  ['RANDARRAY', 'COLUMNS(RANDARRAY(3,2))'],
  ['RANDARRAY', 'ROWS(RANDARRAY())'],
  ['RANDARRAY', 'AND(RANDARRAY(1,1)>=0,RANDARRAY(1,1)<1)'],
  ['RANDARRAY', 'AND(RANDARRAY(1,1,5,10,TRUE)>=5,RANDARRAY(1,1,5,10,TRUE)<=10)'],
  /* ★★でたらめ同士を 比べる 式は 入れない★★
     ＝実Excel も こちらも ★毎回 答えが 変わる★＝★物差しに ならない★
     （1度 `INT(RANDARRAY(…))=RANDARRAY(…)` を 入れて しまい、
       実Excel が FALSE・こちらが TRUE で「合わない」に 見えた。
       ★中身は どちらも 正しく、★式が 悪かった★） */
  ['RANDARRAY', 'AND(RANDARRAY(1,1,7,7,TRUE)=7)'],
  ['RANDARRAY', 'AND(RANDARRAY(3,1,0,1)>=0)'],
  /* ★「整数で 返すか」は ここでは 見ない★
     ＝`SUMPRODUCT(--(MOD(RANDARRAY(…),1)=0))` を 入れたら
       実Excel 5／こちら 2 に なった。中身は ★SUMPRODUCT と でたらめの 表の 組み合わせ★の 話で
       ★RANDARRAY が 整数を 返すか どうかでは ない★。
     ⇒★整数かどうかは 試験の 中で 計算だけを 直に 押して 見る★（tests/formula-yosoku.test.mjs） */
  ['TRIMRANGE', 'ROWS(TRIMRANGE(A1:A10))'],
  ['TRIMRANGE', 'COLUMNS(TRIMRANGE(A1:D1))'],
  ['PERCENTOF', 'PERCENTOF(2,{1;2;3;4})'],
  ['PERCENTOF', 'PERCENTOF({1;2},{1;2;3;4})'],
];

const 行 = 表.map(([名, 式]) => 名 + '\t=' + 式);
fs.writeFileSync(path.join(ここ, 'cases-yosoku.txt'), 行.join('\n') + '\n', { encoding: 'utf-8' });
console.log('★式 ' + 行.length + '本／関数 ' + new Set(表.map((x) => x[0])).size + '個★ … cases-yosoku.txt に 書いた');
