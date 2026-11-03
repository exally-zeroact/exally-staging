/* osu-dai-dake.mjs -- ★同じ 紙を ★自前の 台だけ★で 押す★（2026-09-18）
 *
 *  ★訳★
 *    `osu-kami-webkit.mjs` は ★お客さんの 道★（借り物）を 押して ★99/170（58.2%）★。
 *    ⇒★★「繋いだら どこまで 行くか」は ★別に 測らないと 分かりません★★
 *    ⇒★同じ 170本を ★`lib/shiki-*` だけ★で 押します★（★借り物を 建てません★）
 *
 *  ★★これは「繋いだ後の 数」では ありません★★
 *    ・繋ぐと ★`convertFormula`（JIS→DBCS・YEN→DOLLAR 等）★を 通ります
 *    ・繋ぐと ★JS層（`_jsSet`）★が 先に 横取りする 物が 在ります
 *    ⇒★★ここで 出る 数は「台 単体の 力」です★★
 *    ⇒★★繋いだ 後は 必ず `osu-kami-webkit.mjs` で 測り直します★★
 *
 *
 *  ★★★一番 大きい 弱み（★先に 書きます★）★★★
 *    ★材料は 1組だけです★（8枠目・9枠目の 紙の 頭の 物）。
 *    ★でも 紙は 6枚 在り、★紙ごとに 材料が 違います★★。
 *    ⇒★★材料が 違う 紙の 行は「合わない」と 出ます＝★欠陥では ありません★★★
 *    ⇒★今 分かって いる 分（4本）★
 *        =FORECAST(6,B1:B5,A1:A5) / =FORECAST.LINEAR(...) / =LOOKUP(3,A1:A5,B1:B5)
 *          ... その 紙の B は ★2,4,6,8,10★（ここでは B3 が =1/0）
 *        =BYROW(A1:B2,LAMBDA(r,SUM(r)))
 *          ... その 紙の A1:B2 は ★1,2／2,4★
 *    ⇒★★直し方 ... 紙ごとに 材料を 持たせる★★（★まだ して いません★）
 *
 *  ★見て いない 事★
 *    ・★溢れ（スピル）の 2つ目 以降★は 見て いません（左上だけ）
 *    ・★材料は 1組だけ★（紙の 頭の 物）
 *
 *  使い方: node docs/measured/osu-dai-dake.mjs
 */
import path from 'node:path';
import fs from 'node:fs';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';

const ここ = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(ここ, '..', '..');
const require_ = createRequire(path.join(ROOT, 'package.json'));
const Tsu = require_(path.join(ROOT, 'lib/shiki-tsunagi.js'));
require_(path.join(ROOT, 'lib/shiki-kansuu.js'));

/* ★★XML を 読む 道具を 渡します★★（2026-09-18・★私が 自分で 踏みました★）
   `FILTERXML` は 台が ★道具が 無ければ 計算しない★ 作りです
     （`lib/shiki-tsunagi.js` ... 「半分 合う 答えを 出さない」）
   ⇒node には `DOMParser` / `XPathResult` が 在りません
   ⇒★渡さないと 37本 全部 `#VALUE!` に なります★
   ⇒★私は それを ★「台の 欠陥 22本」★と 数えて 出して しまいました★
   ⇒★★渡したら 32/37 でした★★（★欠陥は 22本では なく 5本★）
   ⇒★★node の 台で 出した 数は 画面の 数では ありません★★ */
try {
  const { JSDOM } = require_('jsdom');
  const 窓 = new JSDOM('<!doctype html>').window;
  Tsu.XML道具を渡す({ DOMParser: 窓.DOMParser, XPathResult: 窓.XPathResult });
} catch (e) {
  console.log('  NG   ★jsdom が 入って いません★＝★FILTERXML は 測れません★');
  console.log('       ★飛ばしません★ ... `npm install` して ください');
  process.exit(3);
}

const H = require_(path.join(ROOT, 'lib/shiki-hyou.js'));

/* ★8枠目・9枠目の 紙の 頭に 書いて ある 材料★（★既定★） */
const 既定の材料 = [
  ['A1', 1], ['A2', 2], ['A3', 3], ['A4', 4], ['A5', 5],
  ['B1', 1], ['B2', 2], ['B3', '=1/0'], ['B4', 4], ['B5', 5],
  ['C1', 1], ['C2', 3], ['C3', 5], ['C4', 7], ['C5', 9],
  ['D1', 9], ['D2', 7], ['D3', 5], ['D4', 3], ['D5', 1],
  ['F1', 1], ['G1', 2], ['F2', 10], ['G2', 20],
];

/* ★★式の 列・答えの 列は ★紙ごとに 名指しします★★
   ＝★機械に 探させようと しましたが ★9枠目の 紙で 外れました★★
     （並びが「式 ／ ★見込み★ ／ .Value2 ／ 答え」で、★見込みを 答えと 読んで いました★）
   ＝★飾り（★）の 割合で 見分ける 案も 13% で 通って しまいました★
   ⇒★★探す 決まりを 作らず 1枚ずつ 目で 見て 名指しします★★
   ⇒★門★ ... 各紙に ★対照★を 1本 持ち、合わなければ ★その場で 止めます★ */
const 紙たち = [
  { 名: 'golden-kansuu-8kaime-2026-09-18.tsv', 式列: 1, 答列: 2,
    対照: { 式: '=PERMUT(0,0)', 答: '1' } },
  { 名: 'golden-kansuu-9kaime-2026-09-18.tsv', 式列: 2, 答列: 4,
    対照: { 式: '=PERMUT(5,2)', 答: '20' } },
  { 名: 'golden-kansuu-7kaime-2026-09-18.tsv', 式列: 1, 答列: 2,
    対照: { 式: '=PERMUT(5,2)', 答: '20' } },
  { 名: 'golden-oddl-6kaime-2026-09-18.tsv', 式列: 1, 答列: 2,
    対照: { 式: '=ODDLPRICE(DATE(2009,3,10),DATE(2009,8,31),DATE(2008,11,30),0,0.05,100,4,0)',
           答: '97.696364140993609' } },
  { 名: 'golden-oddf-buhin-2026-09-16.tsv', 式列: 3, 答列: 4,
    対照: { 式: '=COUPDAYBS(DATE(2008,11,11),DATE(2021,3,1),2,0)', 答: '70' } },
  { 名: 'golden-oddf-to-46ko-2026-09-16.tsv', 式列: 2, 答列: 3, 対照: null,
    /* ★★この 紙だけ 材料が 違います★★（出どころ ... kansuu46-no-dodai.mjs の 材料()）
       ＝A1:A5 = 1,2,3,4,5 ／ ★B1:B5 = 2,4,6,8,10★
       ＝裏取り ... =FORECAST(6,B1:B5,A1:A5) が 12（★B = 2x でしか 12に ならない★） */
    材料: { B1: 2, B2: 4, B3: 6, B4: 8, B5: 10 } },
  { 名: 'golden-oddf-2kaime-2026-09-16.tsv', 式列: 1, 答列: 2,
    対照: { 式: '=ODDFPRICE(DATE(2009,7,2),DATE(2013,1,1),DATE(2009,1,1),DATE(2010,1,1),0.06,0.05,100,2,1)',
           答: '103.09945989078082' } },
  { 名: 'golden-oddf-3kaime-2026-09-16.tsv', 式列: 1, 答列: 2,
    対照: { 式: '=ODDFPRICE(DATE(2009,1,2),DATE(2013,1,1),DATE(2009,1,1),DATE(2009,7,1),0.06,0.05,100,2,1)',
           答: '-2146826252' } },
  { 名: 'golden-oddf-4kaime-2026-09-17.tsv', 式列: 1, 答列: 2,
    対照: { 式: '=ODDFPRICE(DATE(2009,3,1),DATE(2013,1,1),DATE(2009,1,1),DATE(2009,7,1),0,0.05,100,2,0)',
           答: '-2146826252' } },
  { 名: 'golden-oddf-5kaime-2026-09-17.tsv', 式列: 1, 答列: 2,
    対照: { 式: '=ODDFPRICE(DATE(2009,3,1),DATE(2013,1,1),DATE(2009,1,1),DATE(2009,7,1),0.06,0.03,100,2,0)',
           答: '-2146826252' } },
  /* ★★ここから 下は ★マスを 1つも 指しません★★（材料が 要りません） */
  { 名: 'kansuu46/golden-kane-2026-09-07.tsv', 式列: 1, 答列: 2,
    対照: { 式: '=ACCRINT(DATE(2008,3,1),DATE(2008,8,31),DATE(2008,5,1),0.1,1000,2,0)',
           答: '16.666666666666664' } },
  { 名: 'kansuu46/golden-convert-2026-09-07.tsv', 式列: 1, 答列: 2,
    対照: { 式: '=CONVERT(1,"g","g")', 答: '1' } },
  { 名: 'kansuu46/golden-convert2-2026-09-07.tsv', 式列: 1, 答列: 2,
    対照: { 式: '=CONVERT(1,"kft","ft")', 答: '#N/A' } },
  { 名: 'kansuu46/golden-convert3-2026-09-07.tsv', 式列: 1, 答列: 2,
    対照: { 式: '=CONVERT(1,"kBTU","BTU")', 答: '#N/A' } },
  { 名: 'kansuu46/golden-filterxml-2026-09-07.tsv', 式列: 1, 答列: 2,
    対照: { 式: '=FILTERXML("<a><b>1</b>","//b")', 答: '#VALUE!' } },
  { 名: 'kansuu46/golden-isomitted-2026-09-08.tsv', 式列: 1, 答列: 2,
    対照: { 式: '=LAMBDA(x,y,ISOMITTED(y))(1,2)', 答: 'False' } },
  /* ★★入れなかった 紙と その 訳（★書いて 残します★）★★
     `kansuu46/golden-cell6` `golden-cell7` ... =CELL("format",A1)
        ＝★マスに 付いた 表示形式で 答えが 変わります★＝★板では 作れません★
     `golden-hoyuu-27` ... =A1+A2-0.3 ＝★材料が 要り、1列目が 紙の 名前★
     `golden-86-karimono`（1,998本）`golden-346`（3,593本）
        ＝★マス参照が 多く 材料が 要ります★＝★次に 足します★ */
];



/* ══ ★★測り道具の 限りで 落ちる 物（★名指しで 許す★）★★ ══
     ★jsdom は XPath の `name()` を 持って いません★（2026-09-18 実測）
       jsdom .......... `//*[name()='b']` -> `#VALUE!`
       ★本物（WebKit）★ -> ★["1","2","3"]★（★読めます★）
     ＝`lib/formula-filterxml.js` は ★自分で XPath を 作って いません★
       ＝`document.evaluate`（★ブラウザの 本物★）に そのまま 渡す
     ⇒★★台の 欠陥では ありません★★＝★この 道具の 限りです★
     ★裏取りの 道具★ `docs/measured/xpath-name-honmono-de-osu.mjs`
     ★★合ったら 赤に します★★（＝jsdom が 持つ ように なったら 許しを 外す） */
const 道具の限り = new Map([
  ["=FILTERXML(\"<r><b>1</b><b>2</b><b>3</b><c id='x'>A</c><c id='y'>B</c><d><e>ne</e></d></r>\",\"//*[name()='b']\")",
    '★jsdom が name() を 持たない★（本物では 読めます）'],
  ["=ROWS(FILTERXML(\"<r><b>1</b><b>2</b><b>3</b><c id='x'>A</c><c id='y'>B</c><d><e>ne</e></d></r>\",\"//*[name()='b']\"))",
    '★jsdom が name() を 持たない★（本物では 読めます）'],
]);
let 道具の限りで落ちた = 0;
const 許しを外せ = [];

/* ══ ★★関数ごとの 数（★どこが 一番 大きいか を 出す★）★★ ══ */
function 関数名(式) {
  const m = /^=([A-Z0-9_.]+)\(/.exec(式);
  return m ? m[1] : '(不明)';
}

/* ★★真偽の 字だけ 大小を 揃えてから 比べます★★（2026-09-18・Exally1 の 決め）
   ★訳★ ... 紙の 列は `.Value2`＝`False` ／ ★画面に 出るのは `FALSE`★
           ＝★お客さんが 見るのは 画面です★（`.Value2` は 測る 道具の 都合）
   ⇒★台は `FALSE` の まま★／★突き合わせる 側で 揃えます★
   ★★真偽だけです★★＝★他の 字の 大小は 揃えません★（★本物の 違いを 隠さない 為★） */
const 真偽をそろえる = (x) => {
  const s = String(x).trim();
  if (/^(TRUE|FALSE)$/i.test(s)) return s.toUpperCase();
  return s;
};

const 裸 = (s) => String(s === undefined ? '' : s).replace(/★/g, '').replace(/`/g, '').trim();

/* ★実Excel の .Value2 は 誤りを 負の 数で 返します★（記憶の 決まり） */
const 誤りの数 = {
  '-2146826281': '#DIV/0!', '-2146826252': '#NUM!', '-2146826246': '#N/A',
  '-2146826273': '#VALUE!', '-2146826265': '#REF!', '-2146826259': '#NAME?',
  '-2146826288': '#NULL!', '-2146826238': '#CALC!', '-2146826245': '#SPILL!',
};

function 答になるか(a) {
  if (a === '') return false;
  /* ★Excel 自身が 式を 受け付けなかった 行は 答えでは ありません★ */
  if (/打てません|受け付けません|HRESULT/.test(a)) return false;
  return true;
}

const 出し先 = (process.argv.find((a) => a.startsWith('--出し=')) || '').split('=').slice(1).join('=');
const 行ごとの出し = [];
const 問い = [];
const 紙ごと = [];
for (const p of 紙たち) {
  const 生 = fs.readFileSync(path.join(ここ, p.名), 'utf-8').replace(/^\uFEFF/, '');
  const 行たち = 生.split(/\r?\n/)
    .filter((l) => l && !l.startsWith('#') && l.includes('\t'))
    .map((l) => l.split('\t'));
  /* ★★門 ... 対照が 合わなければ 列が ずれて います★★ */
  if (p.対照) {
    const r = 行たち.find((c) => 裸(c[p.式列]) === p.対照.式);
    if (!r) { console.log('  NG   ★' + p.名 + ' に 対照の 式が ありません★'); process.exit(9); }
    if (裸(r[p.答列]) !== p.対照.答) {
      console.log('  NG   ★' + p.名 + ' の 答えの 列が ずれて います★'
        + ' ... 対照 ' + p.対照.式 + ' の 答えは ' + p.対照.答
        + ' の はずが [' + 裸(r[p.答列]) + ']');
      process.exit(9);
    }
  }
  let n = 0;
  for (const c of 行たち) {
    const 式 = 裸(c[p.式列]);
    let 答 = 裸(c[p.答列]);
    if (!式.startsWith('=')) continue;
    if (!答になるか(答)) continue;
    if (誤りの数[答]) 答 = 誤りの数[答];
    問い.push({ 紙: p.名, 材料: p.材料 || null, 式, 答 });
    n += 1;
  }
  紙ごと.push(p.名 + ' ... ' + n + '本（式 ' + p.式列 + '列目 ／ 答え ' + p.答列 + '列目'
    + (p.対照 ? ' ／ ★対照 ok★' : ' ／ ★対照 なし★') + '）');
}


console.log('');
console.log('[osu-dai-dake] ★同じ 紙を 自前の 台だけで 押す★');
console.log('  ★★拾った 式 ... ' + 問い.length + '本★★（★お客さんの 道と 同じ 分母★）');
console.log('  ★借り物は 建てて いません★');

/* ★★紙ごとに 板を 立て直します★★
   ＝★紙ごとに 材料が 違うから★（★前は 1組で 通して 4本 嘘の 赤を 出しました★） */
let 合 = 0, 違 = 0, 知らない = 0, 転 = 0;
const 外れ = [];
const 知らない名 = new Map();
const 字違い = [];
const 関数ごと = new Map();
const 印 = (式, どう) => {
  const n = 関数名(式);
  if (!関数ごと.has(n)) 関数ごと.set(n, { 合: 0, 違: 0, 無: 0 });
  関数ごと.get(n)[どう] += 1;
};
let 今の紙 = null;
let 板 = null;
for (const q of 問い) {
  if (q.紙 !== 今の紙) {
    今の紙 = q.紙;
    板 = new H.表();
    for (const kv of 既定の材料) 板.打つ(kv[0], kv[1]);
    if (q.材料) for (const k of Object.keys(q.材料)) 板.打つ(k, q.材料[k]);
  }
  let 出;
  try {
    板.打つ('J21', q.式);
    出 = 板.字('J21');
  } catch (e) { 転 += 1; 外れ.push(q.式 + ' => ★転んだ★ ' + String(e.message).slice(0, 70)); continue; }
  const 字 = (出 === null || 出 === undefined) ? '' : String(出);
  if (字 === '#NAME?') {
    行ごとの出し.push(q.式 + '\t#NAME?\t' + q.答);
    知らない += 1;
    const m = /^=([A-Z0-9_.]+)\(/.exec(q.式);
    const n = m ? m[1] : '(不明)';
    知らない名.set(n, (知らない名.get(n) || 0) + 1);
    印(q.式, '無');
    continue;
  }
  const 数どうし = /^[-+]?[0-9]*[.]?[0-9]+([eE][-+]?[0-9]+)?$/.test(q.答) && /^[-+]?[0-9]*[.]?[0-9]+([eE][-+]?[0-9]+)?$/.test(字);
  const 同 = (真偽をそろえる(字) === 真偽をそろえる(q.答))
    || (数どうし && Math.abs(Number(字) - Number(q.答)) <= Math.max(1e-9, Math.abs(Number(q.答)) * 1e-9));
  行ごとの出し.push(q.式 + '\t' + 字 + '\t' + q.答
    + (道具の限り.has(q.式) ? '\t★測り道具の限り★' : ''));
  if (同) {
    合 += 1;
    印(q.式, '合');
    /* ★★許した 物が 合ったら 赤に します★★＝★許しを 外して ください★ */
    if (道具の限り.has(q.式)) 許しを外せ.push(q.式);
    /* ★★「数は 同じ・書き方が 違う」だけを 数えます★★
       ★★桁の 数は 数えません★★
         ＝紙の 列は `.Value2`（17桁）／うちは ★画面に 出る 字★
         ＝★元から 別の 物なので 比べても 意味が ありません★
       ⇒★指数の 書き方（1E-10 と 0.0000000001）だけ★を 出します
         ＝★これは 画面に そのまま 出る 違いです★ */
    const 指数か = (x) => /[eE][-+]?[0-9]+$/.test(String(x));
    if (数どうし && 指数か(字) !== 指数か(q.答)) {
      字違い.push(q.式 + ' => 出た [' + 字 + '] ／実Excel [' + q.答 + ']');
    }
  } else if (道具の限り.has(q.式)) {
    道具の限りで落ちた += 1;
    印(q.式, '合');            /* ★台の 欠陥では ないので 合った 側に 入れます★ */
    合 += 1;
  } else { 違 += 1; 印(q.式, '違'); 外れ.push(q.式 + ' => 出た [' + 字 + '] ／実Excel [' + q.答 + ']'); }
}

console.log('');
console.log('  ★★合った ' + 合 + ' / ' + 問い.length + '★★'
  + ' ／ 違った ' + 違 + ' ／ ★台が 知らない（#NAME?） ' + 知らない + '★ ／ 転んだ ' + 転);
console.log('  ★合った 割合 ... ' + (問い.length ? (100 * 合 / 問い.length).toFixed(1) : '0') + '%★');
console.log('');
console.log('  ★★台が 知らない 関数★★');
for (const [n, c] of [...知らない名.entries()].sort((a, b) => b[1] - a[1])) {
  console.log('    ' + n + ' ... ' + c + '本');
}
console.log('');
console.log('  ★★数は 同じ・★指数の 書き方★が 違う ... ' + 字違い.length + '本★★'
  + '（★合った に 入れて います／★画面に 出る 違い★です★）');
for (const s2 of 字違い.slice(0, 6)) console.log('    ・' + s2);
if (字違い.length > 6) console.log('    ...（残り ' + (字違い.length - 6) + '本）');
console.log('');
console.log('  ★★合わない ' + 外れ.length + '本★★');
for (const s of 外れ) console.log('    ・' + s);

/* ══ ★関数ごと（★合わない 数が 多い 順★・上 20）★ ══ */
console.log('');
console.log('  ★★関数ごと ... 合わない 数が 多い 順（上 20）★★');
console.log('    関数              合った  間違い  知らない');
const 並び = [...関数ごと.entries()]
  .map(([n, v]) => ({ n, ...v, 悪: v.違 + v.無 }))
  .filter((x) => x.悪 > 0)
  .sort((a, b) => b.悪 - a.悪)
  .slice(0, 20);
for (const x of 並び) {
  console.log('    ' + x.n.padEnd(18)
    + String(x.合).padStart(5) + String(x.違).padStart(8) + String(x.無).padStart(10));
}

if (出し先) {
  fs.writeFileSync(出し先, 行ごとの出し.join('\n') + '\n', 'utf-8');
  console.log('');
  console.log('  ★1行ずつ 出しました ... ' + 出し先 + '（' + 行ごとの出し.length + '行）★');
}

console.log('');
console.log('  ★★測り道具の 限りで 落ちた（★名指しで 許した★） ... ' + 道具の限りで落ちた + '本★★');
for (const [f, w] of 道具の限り.entries()) {
  console.log('    ' + w);
  console.log('        ' + f.slice(0, 100));
}
if (許しを外せ.length) {
  console.log('');
  console.log('  ★★★合ったのに 許しの ままです ... ' + 許しを外せ.length + '本★★★');
  console.log('    ★許しを 外して ください★');
  for (const f of 許しを外せ) console.log('        ' + f.slice(0, 100));
  process.exitCode = 5;
}
