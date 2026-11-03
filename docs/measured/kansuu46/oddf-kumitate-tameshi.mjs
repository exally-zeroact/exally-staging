/* oddf-kumitate-tameshi.mjs — ★ODDFPRICE の 組み立てを 何通りか 試す★（2026-09-18）
 *
 *  ★★先に 分かった 事★★
 *    ⑴★部品は 210/210 全部 合って います★（`oddf-buhin-awaseru.mjs`）
 *      ＝★日数の 数え方は 元では ありません★
 *    ⑵★落ちる 組の 形が ★1つの 条件★に 揃いました★（`oddf-doko-ga-chigau.mjs`）
 *      ★発行が 準利払日の 上に 無い★ かつ ★NC >= 2★
 *        組1 NC=1 発行×  -> ○全部     組4〜7 発行○      -> ○全部
 *        ★組2 NC=2 発行× -> 違5／合1★  ★組3 NC=3 発行× -> 違4／合2★
 *    ⑶★今の 台には その 形に だけ 効く ★+1 の 手当て★が 在ります★
 *      `if (!発が準日 && 並.length - 1 >= 2) DSC += 1;`
 *      ＝★70/96 まで 上げた 手当て★／★但し 合わせ切れて いません★
 *
 *  ★★だから ここで 試すのは 1つ★★
 *    ★端数の 利札を ★まるごと 1つ★と して 割り引くのでは なく
 *      ★準期間ごとの かけら に 分けて それぞれ 割り引く★★
 *    ＝Microsoft の 覚書きの `Σ(i=1..NC) DCi/NLi / (1+y/f)^(i-1+DSC/E)` の 形
 *
 *  ★★試した 形と 結果（2026-09-18・★全部 外れ★）★★
 *    今の台 ............ ★33★（これが 一番 良い）
 *    手当てなし ........ 30   ＝`+1` は 3本 効いて いる
 *    手当てをAに ....... 30
 *    DSCはDFC-A ........ 29
 *    DFCは準日から ..... 24
 *    Aは準日から ....... 24
 *    DFCもAも準日から .. 24
 *    かけらごと ........ ★6★（一番 悪い）
 *    YEARFRACで出す .... 18   ＝`YF_発初 x f` などを そのまま 使う
 *    端は全体から引く .. 22   ＝30/360 は 足し算に ならないので 全体から 引く
 *  ⇒★★9つ 潰しました／どれも 今より 悪い★★
 *  ⇒★★当てずっぽうの 案を 木に 残しません★★＝★口で 呼ぶ 形だけ 残します★
 *
 *  ★★残差を 測りました（★これが 一番 効く 手掛かり★）★★
 *    ★落ちる 9本の（うち − 実Excel）★
 *      組B basis 0 ... -0.00084   （券 3 の ★0.03%★）
 *      組B basis 1 ... +0.01440   （券 3 の ★0.48%★ ＝★181日の うち 約0.9日★）
 *      組B basis 2 ... +0.01582   組B basis 3 ... +0.01561
 *      組C basis 1 ... -0.00081   （券 1.125 の ★0.07%★）
 *      組C basis 0 ... -0.01343   組C basis 4 ... -0.02687（★90日の うち 約2.2日★）
 *    ⇒★★1回の 利払い ぶんでは ありません／★1〜2日ぶん★です★★
 *    ⇒★★構えの 間違いでは なく ★日数の 端★の 話★★
 *    ⇒★だから「割り引き方を 変える」案が 全部 外れた のは 筋が 通ります★
 *    ⇒★★次に 見るのは ★かけらの 日数★です★★
 *       （★組ごとに 1つの basis だけ ほぼ 合って いる★＝★basis で 端が 動く★）
 *
 *  ★★これは 当てでは ありません★★
 *    ・★形を 1つ 決めて 48本 全部で 数を 出します★
 *    ・★今より 悪く なったら 捨てます★（★紙は 1本も 触りません★）
 *    ・★試す 前に「今 何本 合うか」を 出します★
 *
 *  使い方: node docs/measured/kansuu46/oddf-kumitate-tameshi.mjs
 */
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '../../..');
const require_ = createRequire(path.join(ROOT, 'package.json'));
const K = require_(path.join(ROOT, 'lib/formula-kane.js'));

const 紙道 = path.join(ROOT, 'docs/measured/kansuu46/golden-kane-2026-09-07.tsv');
const 行 = fs.readFileSync(紙道, 'utf-8').split(/\r?\n/).filter((l) => l && !l.startsWith('#'));

function 引数を割る(s) {
  const 出 = []; let 深 = 0, 今 = '';
  for (const ch of s) {
    if (ch === '(') 深++;
    if (ch === ')') 深--;
    if (ch === ',' && 深 === 0) { 出.push(今); 今 = ''; continue; }
    今 += ch;
  }
  出.push(今);
  return 出;
}
function 数に(s) {
  const m = /^DATE\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)$/.exec(s.trim());
  if (!m) return Number(s.trim());
  return K.日から数(Number(m[1]), Number(m[2]), Number(m[3]));
}
const 日 = (n) => K.数から日(n);
const 数 = (d) => K.日から数(d.y, d.m, d.d);

/* ★準利払日の 並び★（初回から f ぶんずつ さかのぼり 発行を 越えるまで）
     ★`formula-kane.js` の `準の並び前へ` と 同じ 形を 作ります★
     ＝★向こうは 外に 出して いません★ので ここで 作り直します */
function 準日たち(発行, 初回, 頻度) {
  const 月 = 12 / 頻度;
  const 初 = 日(初回);
  const 末日 = (y, m) => new Date(Date.UTC(y, m, 0)).getUTCDate();
  const 末 = (初.d === 末日(初.y, 初.m));
  const 出 = [初回];
  for (let k = 1; k < 400; k++) {
    const m = 初.m - k * 月;
    const yy = 初.y + Math.floor((m - 1) / 12);
    const mm = ((m - 1) % 12 + 12) % 12 + 1;
    const dd = 末 ? 末日(yy, mm) : Math.min(初.d, 末日(yy, mm));
    const n = K.日から数(yy, mm, dd);
    出.unshift(n);
    if (n <= 発行) break;
  }
  return 出;
}
/* ★1期の 長さ★（`比()` の `Bの長さ` と 同じ） */
function 期の長さ(a, c, f, b) {
  if (b === 1) return Math.round(c - a);
  if (b === 3) return 365 / f;
  return 360 / f;
}
/* ★basis ごとの 日数★（`formula-kane.js` の `日数`） */
const 日数 = (a, c, b) => K.日数(日(a), 日(c), b);

/* ══ ★試す 組み立て★ ══ */
function 組み立てる(v, 形) {
  const [決済, 満期, 発行, 初回, 利率v, 利回りv, 償還, 頻度] = v;
  const b = v.length > 8 ? v[8] : 0, f = 頻度;
  const 準 = 準日たち(発行, 初回, f);
  const NC = 準.length - 1;
  /* ★かけら★ ... i 番目の 準期間 [準[i-1], 準[i]] と [発行, 初回] の 重なり */
  const かけら = [];
  for (let i = 1; i <= NC; i++) {
    const a = 準[i - 1], c = 準[i];
    const s = Math.max(a, 発行), e = Math.min(c, 初回);
    かけら.push({ DC: e > s ? 日数(s, e, b) : 0, NL: 期の長さ(a, c, f, b), 終: c });
  }
  /* ★DSC★ ... 決済 -> 初回 を ★準期間の 数★で（★今の 台と 同じ 数え方★） */
  let DSC = 0;
  for (let i = 1; i <= NC; i++) {
    const a = 準[i - 1], c = 準[i];
    const s = Math.max(a, 決済), e = c;
    if (e <= s) continue;
    if (s === a) { DSC += 1; continue; }              /* まるごとの 期は 1 */
    DSC += 日数(s, e, b) / 期の長さ(a, c, f, b);
  }
  /* ★A★ ... 発行 -> 決済（まとめて 割る） */
  let A = 0;
  for (let i = 1; i <= NC; i++) {
    const a = 準[i - 1], c = 準[i];
    const s = Math.max(a, 発行), e = Math.min(c, 決済);
    if (e <= s) continue;
    A += 日数(s, e, b) / 期の長さ(a, c, f, b);
  }
  const N = 1 + K.利払回数(日(初回), 日(満期), f);
  const 割 = 1 + 利回りv / f, 券 = 100 * 利率v / f;

  let 出 = 償還 / Math.pow(割, (N - 1) + DSC);
  if (形 === 'まとめて') {
    let DFC = 0;
    for (const x of かけら) DFC += (x.DC === 0 ? 0 : (x.DC === x.NL ? 1 : x.DC / x.NL));
    出 += 券 * DFC / Math.pow(割, DSC);
  } else if (形 === 'かけらごと') {
    /* ★★かけら i は ★その 準日に 払われる★ と して 割り引く★★
         ＝一番 後ろ（初回）の かけらは DSC ／ 1つ 前は DSC-1 ... */
    for (let i = 0; i < かけら.length; i++) {
      const x = かけら[i];
      const 比 = (x.DC === 0) ? 0 : (x.DC === x.NL ? 1 : x.DC / x.NL);
      const 指 = DSC - (かけら.length - 1 - i);
      出 += 券 * 比 / Math.pow(割, 指);
    }
  }
  for (let k = 2; k <= N; k++) 出 += 券 / Math.pow(割, (k - 1) + DSC);
  return 出 - 券 * A;
}

/* ══ ★48本で 数える★ ══ */
const 組 = [];
for (const l of 行) {
  const c = l.split('\t');
  if (c[0] !== 'ODDFPRICE') continue;
  const v = 引数を割る(c[1].slice(c[1].indexOf('(') + 1, c[1].lastIndexOf(')'))).map(数に);
  組.push({ v: v, 正: Number(c[2]), 式: c[1], 答字: c[2] });
}
const 近い = (a, b) => isFinite(a) && isFinite(b) && Math.abs(a - b) <= Math.max(1e-9, Math.abs(b) * 1e-9);

console.log('');
console.log('★★ODDFPRICE の 組み立てを 試す★★（★紙 ' + 組.length + '本★）');
console.log('');
const 再現 = [];
for (const 形 of ['今の台', 'かけらごと', '手当てなし', '手当てをAに', 'DSCはDFC-A',
  'DFCは準日から', 'Aは準日から', 'DFCもAも準日から', 'YEARFRACで出す', '端は全体から引く', '月末に揃える（前の形）', '端は準日ごとに数える（前の形）', '端は並びから', '1は外す（前の形）']) {
  let 合 = 0, 違 = 0, 数えない = 0;
  const 外 = [];
  for (const g of 組) {
    if (!isFinite(g.正)) { 数えない++; continue; }    /* #NUM! の 行 */
    let 出;
    try {
      /* ★★写しを やめて ★同じ 道★の 形の 口を 使います★★（2026-09-18）
           ＝写しは 今の 台を 再現 出来ませんでした（33本 対 24本）
           ＝★`lib/formula-kane.js` に 形の 口を 1つ 足しました★
           ＝★既定（省略）は 今の 形★＝★お客さんの 道は 変わりません★ */
      const b9 = g.v.length > 8 ? g.v[8] : 0;
      出 = (形 === '写し')
        ? 組み立てる(g.v, 'まとめて')
        : K.初回端数の価格を試す(g.v[0], g.v[1], g.v[2], g.v[3], g.v[4], g.v[5], g.v[6], g.v[7], b9,
          形 === '今の台' ? '今の形' : 形);
    } catch (e) { 出 = NaN; }
    if (近い(出, g.正)) 合++;
    else { 違++; if (外.length < 4) 外.push('      ' + g.式.slice(0, 78) + '\n        正 ' + g.正 + ' ／ 出 ' + 出); }
  }
  再現.push(合);
  console.log('  ★' + 形.padEnd(10) + '★  合 ' + String(合).padStart(3) + ' ／ 違 ' + String(違).padStart(3)
    + ' ／ 数えない（#NUM!）' + 数えない);
  for (const s of 外) console.log(s);
}
/* ══ ★★一番 大事な 門★★ ══（2026-09-18）
     ★私の 記憶に 書いて ある 事★
       「★見込みを 出す 前に ★道具が 今の 数を 再現できるか★を 見る★」
     ⇒★この 道具の 「まとめて」は ★今の 台と 同じ 形の つもり★です★
     ⇒★同じ 数が 出なければ ★別の 物を 測って います★★
     ⇒★★その まま 「かけらごと が 悪い」と 読んでは いけません★★ */
console.log('');
/* ★★40 -> 41★★（2026-09-18）
     ＝`basis 1 にも (E - BS) / E を 当てる` を 木に 入れました
     ＝★この 数は 「今の 台が 何本 合うか」★＝★直したら ここも 直す★ */
if (再現[0] !== 41) {
  console.log('★★★この 道具は 今の 台を 再現 出来て いません★★★');
  console.log('    今の台 ' + 再現[0] + '本');
  console.log('    ⇒★だから 「かけらごと」の 数から 何も 言えません★');
  console.log('    ⇒★先に 「まとめて」を 今の 台と 同じ 数に してから★');
  console.log('      ★違うのは DSC の 数え方（+1 の 手当て）と `比()` の 決まりです★');
} else {
  console.log('★この 道具は 今の 台を 再現 出来て います★（' + 再現[0] + '本）');
  console.log('  ⇒★だから 「かけらごと」の 数を 読んで よい★');
}
console.log('');
console.log('★言えない 事★');
console.log('  ・★`#NUM!` の 行は 数えて いません★（★数と 比べられない★）');
console.log('  ・★ODDFYIELD は 押して いません★（★価格が 決まれば 一度に 直る★＝別の 道具で 確かめ済）');
