/* shiki-wo-tsukuru.mjs — ★実Excel に 打たせる 式を 作る★（2026-09-07）
 *
 *  ★★なぜ 先に 式を 作るのか★★
 *    お金の 関数は ★日数の 数え方（basis）が 5通り★あり、
 *    ★1つ 試して 合っても 残り 4つが 外れている★事が 普通に 起きます。
 *    ⇒★★同じ 式を basis 0〜4 で 全部 打たせる★★
 *    ⇒ 作る 前に ★答えを 実Excel から 取る★（私が 計算しない）
 *
 *  ★★式は 1か所に だけ 置く★★
 *    ここで 作った `cases-kane.txt` を
 *      ①実Excel に 打たせる（golden を 取る）
 *      ②出来た 物を 押す（試験）
 *    の ★両方が 読む★＝★食い違いようが ない★
 *
 *  使い方: node docs/measured/kansuu46/shiki-wo-tsukuru.mjs
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ここ = path.dirname(fileURLToPath(import.meta.url));

/* ★日付は DATE(...) で 書く★＝地域の 書き方に 左右されない */
const 表 = [
  /* [関数, 引数（basis を 除く）, basis を 付けるか] */
  ['ACCRINT', 'DATE(2008,3,1),DATE(2008,8,31),DATE(2008,5,1),0.1,1000,2', true],
  ['ACCRINT', 'DATE(2008,3,5),DATE(2008,9,5),DATE(2008,7,15),0.075,1000,4', true],
  ['ACCRINTM', 'DATE(2008,4,1),DATE(2008,6,15),0.1,1000', true],
  ['ACCRINTM', 'DATE(2007,10,3),DATE(2008,2,28),0.06,5000', true],
  ['AMORDEGRC', '2400,DATE(2008,8,19),DATE(2008,12,31),300,1,0.15', true],
  ['AMORDEGRC', '2400,DATE(2008,8,19),DATE(2008,12,31),300,3,0.15', true],
  ['AMORLINC', '2400,DATE(2008,8,19),DATE(2008,12,31),300,1,0.15', true],
  ['AMORLINC', '2400,DATE(2008,8,19),DATE(2008,12,31),300,2,0.15', true],
  ['COUPDAYBS', 'DATE(2011,1,25),DATE(2011,11,15),2', true],
  ['COUPDAYBS', 'DATE(2008,2,29),DATE(2020,8,31),4', true],
  ['COUPDAYS', 'DATE(2011,1,25),DATE(2011,11,15),2', true],
  ['COUPDAYS', 'DATE(2008,2,29),DATE(2020,8,31),4', true],
  ['COUPDAYSNC', 'DATE(2011,1,25),DATE(2011,11,15),2', true],
  ['COUPDAYSNC', 'DATE(2008,2,29),DATE(2020,8,31),4', true],
  ['COUPNCD', 'DATE(2011,1,25),DATE(2011,11,15),2', true],
  ['COUPNCD', 'DATE(2008,2,29),DATE(2020,8,31),4', true],
  ['COUPNUM', 'DATE(2011,1,25),DATE(2011,11,15),2', true],
  ['COUPNUM', 'DATE(2008,2,29),DATE(2020,8,31),4', true],
  ['COUPPCD', 'DATE(2011,1,25),DATE(2011,11,15),2', true],
  ['COUPPCD', 'DATE(2008,2,29),DATE(2020,8,31),4', true],
  ['DISC', 'DATE(2018,7,1),DATE(2048,1,1),97.975,100', true],
  ['DISC', 'DATE(2008,1,25),DATE(2008,6,15),97.975,100', true],
  ['DURATION', 'DATE(2018,7,1),DATE(2048,1,1),0.08,0.09,2', true],
  ['DURATION', 'DATE(2008,1,1),DATE(2016,1,1),0.08,0.09,4', true],
  ['INTRATE', 'DATE(2008,2,15),DATE(2008,5,15),1000000,1014420', true],
  ['INTRATE', 'DATE(2007,4,1),DATE(2008,3,31),500000,520000', true],
  ['MDURATION', 'DATE(2008,1,1),DATE(2016,1,1),0.08,0.09,2', true],
  ['MDURATION', 'DATE(2018,7,1),DATE(2048,1,1),0.08,0.09,4', true],
  ['ODDFPRICE', 'DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2', true],
  ['ODDFYIELD', 'DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0575,84.5,100,2', true],
  ['ODDLPRICE', 'DATE(2008,2,7),DATE(2008,6,15),DATE(2007,10,15),0.0375,0.0405,100,2', true],
  ['ODDLYIELD', 'DATE(2008,4,20),DATE(2008,6,15),DATE(2007,12,24),0.0375,99.875,100,2', true],
  ['PRICE', 'DATE(2008,2,15),DATE(2017,11,15),0.0575,0.065,100,2', true],
  ['PRICE', 'DATE(2008,2,15),DATE(2017,11,15),0.0575,0.065,100,4', true],
  ['PRICEDISC', 'DATE(2008,2,16),DATE(2008,3,1),99.795,100', true],
  ['PRICEDISC', 'DATE(2008,1,25),DATE(2008,6,15),0.0525,100', true],
  ['PRICEMAT', 'DATE(2008,2,15),DATE(2008,4,13),DATE(2007,11,11),0.061,0.061', true],
  ['PRICEMAT', 'DATE(2008,2,15),DATE(2008,4,13),DATE(2007,11,11),0.061,0.08', true],
  ['RECEIVED', 'DATE(2008,2,15),DATE(2008,5,15),1000000,0.0575', true],
  ['RECEIVED', 'DATE(2007,4,1),DATE(2008,3,31),500000,0.04', true],
  ['YIELD', 'DATE(2008,2,15),DATE(2016,11,15),0.0575,95.04287,100,2', true],
  ['YIELD', 'DATE(2008,2,15),DATE(2016,11,15),0.0575,95.04287,100,1', true],
  ['YIELDDISC', 'DATE(2008,2,16),DATE(2008,3,1),99.795,100', true],
  ['YIELDDISC', 'DATE(2008,1,25),DATE(2008,6,15),97.975,100', true],
  ['YIELDMAT', 'DATE(2008,3,15),DATE(2008,11,3),DATE(2007,11,8),0.0625,100.0123', true],
  ['YIELDMAT', 'DATE(2008,3,15),DATE(2008,11,3),DATE(2007,11,8),0.0625,95', true],
  /* ★★ここから 下は「どちらの 数え方か」を 決める 為に 足した 式★★
     ⇒ 1組だけだと ★2つの 考えが どちらも 合ってしまう★ ＝ 決められない
     ⇒★答えが 割れる 日付を わざと 選んだ★ */
  /* PRICEMAT／YIELDMAT の B が「発行→満期」か「発行→決済」かで 割れる 組 */
  ['PRICEMAT', 'DATE(2007,6,10),DATE(2008,3,5),DATE(2007,3,10),0.05,0.05', true],
  ['PRICEMAT', 'DATE(2011,9,1),DATE(2012,4,30),DATE(2011,5,1),0.05,0.06', true],
  ['PRICEMAT', 'DATE(2008,6,15),DATE(2010,1,15),DATE(2007,1,15),0.05,0.06', true],
  ['YIELDMAT', 'DATE(2007,6,10),DATE(2008,3,5),DATE(2007,3,10),0.05,99', true],
  ['YIELDMAT', 'DATE(2011,9,1),DATE(2012,4,30),DATE(2011,5,1),0.05,99', true],
  ['YIELDMAT', 'DATE(2008,6,15),DATE(2010,1,15),DATE(2007,1,15),0.05,99', true],
  /* ODDFPRICE／ODDFYIELD … 端数の 期間の 長さが basis で 割れる 組 */
  ['ODDFPRICE', 'DATE(2008,1,15),DATE(2012,6,30),DATE(2007,11,20),DATE(2008,6,30),0.06,0.055,100,2', true],
  ['ODDFYIELD', 'DATE(2008,1,15),DATE(2012,6,30),DATE(2007,11,20),DATE(2008,6,30),0.06,98.5,100,2', true],
  ['ODDFPRICE', 'DATE(2009,3,10),DATE(2015,2,28),DATE(2008,12,5),DATE(2009,8,31),0.045,0.05,100,4', true],
  /* ODDLPRICE／ODDLYIELD … 別の 日付でも 同じ 数え方か */
  ['ODDLPRICE', 'DATE(2009,3,10),DATE(2009,8,31),DATE(2008,11,30),0.045,0.05,100,4', true],
  ['ODDLYIELD', 'DATE(2009,3,10),DATE(2009,8,31),DATE(2008,11,30),0.045,99.5,100,4', true],
  /* 利払日が 月末に 寄る 組（ここで 1日 ずれると 利息が 毎回 違う） */
  ['COUPDAYS', 'DATE(2008,3,31),DATE(2015,8,31),2', true],
  ['COUPDAYSNC', 'DATE(2008,3,31),DATE(2015,8,31),2', true],
  ['COUPDAYBS', 'DATE(2008,3,31),DATE(2015,8,31),2', true],
  ['COUPNCD', 'DATE(2008,3,31),DATE(2015,8,31),2', true],
  ['COUPPCD', 'DATE(2008,3,31),DATE(2015,8,31),2', true],
  ['COUPNUM', 'DATE(2008,3,31),DATE(2015,8,31),2', true],
  ['PRICE', 'DATE(2008,3,31),DATE(2015,8,31),0.05,0.06,100,2', true],
  ['YIELD', 'DATE(2008,3,31),DATE(2015,8,31),0.05,95,100,2', true],
  ['DURATION', 'DATE(2008,3,31),DATE(2015,8,31),0.05,0.06,2', true],
  /* ACCRINT … 期間を 3つ またぐ 組（丸ごとの 期間が 2つ） */
  ['ACCRINT', 'DATE(2008,1,15),DATE(2009,1,15),DATE(2008,11,20),0.06,1000,4', true],
  ['ACCRINT', 'DATE(2007,2,28),DATE(2008,2,29),DATE(2007,12,31),0.06,1000,2', true],
  /* ★VDB は basis を 取らない★（代わりに 期間と no_switch） */
  ['VDB', '2400,300,3650,0,1', false],
  ['VDB', '2400,300,120,0,1', false],
  ['VDB', '2400,300,10,0,1', false],
  ['VDB', '2400,300,10,2,5', false],
  ['VDB', '2400,300,10,0,0.875', false],
  ['VDB', '2400,300,10,6,10,1.5', false],
  ['VDB', '2400,300,10,0,2,2,TRUE', false],
  ['VDB', '2400,300,10,0,2,2,FALSE', false],
  ['VDB', '10000,1000,5,0,1', false],
  ['VDB', '10000,1000,5,4,5', false],
  ['VDB', '10000,1000,5,0,5', false],
  ['VDB', '10000,1000,5,1.5,3.5', false],
  ['VDB', '10000,1000,5,0,3,1', false],
  ['VDB', '10000,1000,5,0,3,3', false],
  ['VDB', '10000,0,5,3,5,2,TRUE', false],
  ['VDB', '10000,0,5,3,5,2,FALSE', false],
  /* AMOR … 期を 進めた 時に 1円ずつ ずれないか */
  ['AMORDEGRC', '2400,DATE(2008,8,19),DATE(2008,12,31),300,5,0.15', true],
  ['AMORDEGRC', '2400,DATE(2008,8,19),DATE(2008,12,31),300,6,0.15', true],
  ['AMORDEGRC', '10000,DATE(2010,1,20),DATE(2010,12,31),500,2,0.25', true],
  ['AMORLINC', '2400,DATE(2008,8,19),DATE(2008,12,31),300,6,0.15', true],
  ['AMORLINC', '10000,DATE(2010,1,20),DATE(2010,12,31),500,3,0.25', true],
  /* ★端数の 初回が「1期より 長い」時の 割り引き方を 決める 為の 組★
     ⇒ 発行日を ★ちょうど 準利払日に 合わせて★ 端数を 無くし、
       ★決済が どの 期に 居るか だけ★を 変える */
  ['ODDFPRICE', 'DATE(2009,3,1),DATE(2013,1,1),DATE(2009,1,1),DATE(2010,1,1),0.06,0.05,100,2', true],
  ['ODDFPRICE', 'DATE(2009,9,1),DATE(2013,1,1),DATE(2009,1,1),DATE(2010,1,1),0.06,0.05,100,2', true],
  ['ODDFPRICE', 'DATE(2009,3,1),DATE(2013,1,1),DATE(2008,7,1),DATE(2010,1,1),0.06,0.05,100,2', true],
  ['ODDFPRICE', 'DATE(2009,7,10),DATE(2013,1,1),DATE(2008,7,1),DATE(2010,1,1),0.06,0.05,100,2', true],
  ['ODDFPRICE', 'DATE(2009,3,1),DATE(2013,1,1),DATE(2009,4,1),DATE(2010,1,1),0.06,0.05,100,2', true],
  ['ODDFYIELD', 'DATE(2009,3,1),DATE(2013,1,1),DATE(2009,1,1),DATE(2010,1,1),0.06,99,100,2', true],
  ['ODDFYIELD', 'DATE(2009,9,1),DATE(2013,1,1),DATE(2009,1,1),DATE(2010,1,1),0.06,99,100,2', true],
  /* ★AMORDEGRC の 終わり方を 決める 為に 期を 0 から 並べて 打たせる★
     （1つ2つでは「半分に する 回」と「0に なる 回」の 決まりが 分からない） */
  ['AMORDEGRC', '10000,DATE(2010,1,20),DATE(2010,12,31),500,0,0.25', true],
  ['AMORDEGRC', '10000,DATE(2010,1,20),DATE(2010,12,31),500,1,0.25', true],
  ['AMORDEGRC', '10000,DATE(2010,1,20),DATE(2010,12,31),500,2,0.25', true],
  ['AMORDEGRC', '10000,DATE(2010,1,20),DATE(2010,12,31),500,3,0.25', true],
  ['AMORDEGRC', '10000,DATE(2010,1,20),DATE(2010,12,31),500,4,0.25', true],
  ['AMORDEGRC', '10000,DATE(2010,1,20),DATE(2010,12,31),500,5,0.25', true],
  ['AMORDEGRC', '10000,DATE(2010,1,20),DATE(2010,12,31),500,6,0.25', true],
  ['AMORDEGRC', '10000,DATE(2010,1,20),DATE(2010,12,31),500,0,0.2', true],
  ['AMORDEGRC', '10000,DATE(2010,1,20),DATE(2010,12,31),500,1,0.2', true],
  ['AMORDEGRC', '10000,DATE(2010,1,20),DATE(2010,12,31),500,2,0.2', true],
  ['AMORDEGRC', '10000,DATE(2010,1,20),DATE(2010,12,31),500,3,0.2', true],
  ['AMORDEGRC', '10000,DATE(2010,1,20),DATE(2010,12,31),500,4,0.2', true],
  ['AMORDEGRC', '10000,DATE(2010,1,20),DATE(2010,12,31),500,5,0.2', true],
  ['AMORDEGRC', '10000,DATE(2010,1,20),DATE(2010,12,31),500,6,0.2', true],
  ['AMORDEGRC', '10000,DATE(2010,1,20),DATE(2010,12,31),500,7,0.2', true],
  ['AMORDEGRC', '10000,DATE(2010,1,20),DATE(2010,12,31),0,0,0.1', true],
  ['AMORDEGRC', '10000,DATE(2010,1,20),DATE(2010,12,31),0,1,0.1', true],
  ['AMORDEGRC', '10000,DATE(2010,1,20),DATE(2010,12,31),0,2,0.1', true],
  ['AMORDEGRC', '10000,DATE(2010,1,20),DATE(2010,12,31),0,3,0.1', true],
  ['AMORDEGRC', '10000,DATE(2010,1,20),DATE(2010,12,31),0,4,0.1', true],
  ['AMORDEGRC', '10000,DATE(2010,1,20),DATE(2010,12,31),0,5,0.1', true],
  ['AMORDEGRC', '2400,DATE(2008,8,19),DATE(2008,12,31),300,0,0.15', true],
  ['AMORDEGRC', '2400,DATE(2008,8,19),DATE(2008,12,31),300,1,0.15', true],
  ['AMORDEGRC', '2400,DATE(2008,8,19),DATE(2008,12,31),300,2,0.15', true],
  ['AMORDEGRC', '2400,DATE(2008,8,19),DATE(2008,12,31),300,3,0.15', true],
  ['AMORDEGRC', '2400,DATE(2008,8,19),DATE(2008,12,31),300,4,0.15', true],
  ['AMORDEGRC', '2400,DATE(2008,8,19),DATE(2008,12,31),300,5,0.15', true],
  ['AMORDEGRC', '2400,DATE(2008,8,19),DATE(2008,12,31),300,6,0.15', true],
];

const 行 = [];
for (const [名, 引数, basisか] of 表) {
  if (basisか) {
    for (let b = 0; b <= 4; b++) 行.push(名 + '\t=' + 名 + '(' + 引数 + ',' + b + ')');
    /* ★basis を 省いた 形★も 見る（省くと 0 と 同じか） */
    行.push(名 + '\t=' + 名 + '(' + 引数 + ')');
  } else {
    行.push(名 + '\t=' + 名 + '(' + 引数 + ')');
  }
}

/* ★日付を 返す 3つは 数のままだと 読み違える★
   ⇒ TEXT で 年月日に して ★2通り★ 取る（数と 字の 両方） */
const 日付を返す = ['COUPNCD', 'COUPPCD'];
const 足す = [];
for (const l of 行) {
  const [名, 式] = l.split('\t');
  if (日付を返す.includes(名)) 足す.push(名 + '\t=TEXT(' + 式.slice(1) + ',"yyyy-mm-dd")');
}
行.push(...足す);

fs.writeFileSync(path.join(ここ, 'cases-kane.txt'), 行.join('\n') + '\n', { encoding: 'utf-8' });
const 数 = new Set(行.map((l) => l.split('\t')[0])).size;
console.log('★式 ' + 行.length + '本／関数 ' + 数 + '個★ … cases-kane.txt に 書いた');
