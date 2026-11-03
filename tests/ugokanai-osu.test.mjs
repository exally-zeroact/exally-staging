/* ugokanai-osu.test.mjs — ★「動かない」と 数えた 物を ★1つ残らず 本物の 式で 押す★★（2026-09-07）
 *
 *  ★★なぜ 要るか（指示役 2026-09-07）★★
 *    `CONVERT` は 台帳の「動かない 43個」に 載っていた。
 *    ⇒★でも 本当は 動いていて、しかも ★答えが 違った★（0.4536／実Excel 0.45359237）
 *    ⇒★★`#NAME?` なら お客さんは 気づく／★それらしい 数★は 気づかない★★
 *    ⇒ 見つかったのは ★偶然★だった
 *      （棚の 見張り kansuu-tana が 押すのは ★棚の 13個★で、
 *        台帳の 43個とは ★別の 集まり★だった）
 *    ⇒★★だから 台帳に 載っている 物を 1つ残らず 押す★★
 *
 *  ★★何を 見るか★★
 *    ①`#NAME?` が 返る … ★正しい★（名前が 無い＝お客さんに はっきり 見える）
 *    ②★何か 値が 返る★ … ★★赤★★（CONVERT と 同じ 家＝それらしい 間違い）
 *    ③押した 数を 出す（★何本 中 何本★）
 *
 *  ★★外へ 出る 物も ここでは 安全★★
 *    WEBSERVICE などは ★作っていない★ので 押しても ★名前が 無いだけ★。
 *    ⇒ 実Excel には ★打たせません★（本当に 外へ 出て お金が かかる）
 *
 *  使い方: node tests/ugokanai-osu.test.mjs
 *          node tests/ugokanai-osu.test.mjs --self-test
 */
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath, pathToFileURL } from 'node:url';

const ここ = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(ここ, '..');
const require_ = createRequire(import.meta.url);
const 直に走った = process.argv[1] && pathToFileURL(process.argv[1]).href === import.meta.url;
const 自己試験 = 直に走った && process.argv.includes('--self-test');

let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ✓ ' + n); } catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + (e && e.message)); } };

/* ── 本番と 同じ 物を 積む ── */
const HFns = require_(path.join(ROOT, 'hyperformula.full.min.js'));
const EF = require_(path.join(ROOT, 'exally-formula.js'));
const 積1 = EF.registerExallyFunctions(HFns) === true;
const HF0 = HFns.HyperFormula;
const H = Object.assign(Object.create(HF0), HFns,
  { registerFunctionPlugin: HF0.registerFunctionPlugin.bind(HF0) });
const 積 = [
  require_(path.join(ROOT, 'lib/formula-extra-plug.js')).つなぐ(H, require_(path.join(ROOT, 'lib/formula-extra.js'))),
  require_(path.join(ROOT, 'lib/formula-nokori-plug.js')).つなぐ(H, require_(path.join(ROOT, 'lib/formula-nokori.js'))),
  require_(path.join(ROOT, 'lib/formula-kane-plug.js')).つなぐ(H, require_(path.join(ROOT, 'lib/formula-kane.js'))),
  require_(path.join(ROOT, 'lib/formula-yosoku-plug.js')).つなぐ(H, require_(path.join(ROOT, 'lib/formula-yosoku.js'))),
];
const hf = HF0.buildEmpty({ licenseKey: 'gpl-v3', useArrayArithmetic: true, smartRounding: false });
const SID = hf.getSheetId(hf.addSheet('S'));
EF.initExallyFormula(hf);
const JS層 = EF._jsComputeFormula || null;

const 下敷き = [];
for (let r = 0; r < 8; r++) 下敷き.push([r + 1, (r + 1) * 2, (r + 1) * 3, (r + 1) * 4]);

function 押す(式) {
  if (typeof JS層 === 'function') {
    try {
      const v = JS層(0, 式);
      if (v !== null && v !== undefined) return { 出: v, 道: 'JS層' };
    } catch (e) { return { 出: '★JS層で 例外★', 道: 'JS層' }; }
  }
  let 後;
  try { 後 = EF.convertFormula(式); } catch (e) { return { 出: '#NAME?', 道: '書き換えで 例外' }; }
  try {
    const 表 = 下敷き.map((r) => r.slice());
    表.push([後]);
    for (let i = 0; i < 40; i++) 表.push([null]);
    hf.setSheetContent(SID, 表);
    const v = hf.getCellValue({ sheet: SID, row: 下敷き.length, col: 0 });
    if (v && v.type) return { 出: '#' + v.type, 道: 'engine' };
    return { 出: v, 道: 'engine' };
  } catch (e) { return { 出: '#NAME?', 道: 'engine で 例外' }; }
}

/** ★台帳の 名前 → 本物の 形の 式★
 *  ★手で 決め打った 引数では「形が 合わないだけ」で 名前が 通らない★ので
 *  ★その 関数が 本当に 使われる 形★で 1本ずつ 書く。 */
const 式たち = {
  'BYCOL': '=BYCOL(A1:B2,LAMBDA(c,SUM(c)))',
  'BYROW': '=BYROW(A1:B2,LAMBDA(r,SUM(r)))',
  'CALL': '=CALL("kernel32","GetTickCount","J")',
  'CUBEKPIMEMBER': '=CUBEKPIMEMBER("x","y",1)',
  'CUBEMEMBER': '=CUBEMEMBER("x","[a].[b]")',
  'CUBEMEMBERPROPERTY': '=CUBEMEMBERPROPERTY("x","[a].[b]","p")',
  'CUBERANKEDMEMBER': '=CUBERANKEDMEMBER("x","[a]",1)',
  'CUBESET': '=CUBESET("x","[a].[b].children")',
  'CUBESETCOUNT': '=CUBESETCOUNT(A1)',
  'CUBEVALUE': '=CUBEVALUE("x","[a].[b]")',
  'FORECAST.ETS': '=FORECAST.ETS(45658,{10;12;16;14},{45292;45323;45352;45383})',
  'FORECAST.ETS.CONFINT': '=FORECAST.ETS.CONFINT(45658,{10;12;16;14},{45292;45323;45352;45383})',
  'FORECAST.ETS.SEASONALITY': '=FORECAST.ETS.SEASONALITY({10;12;16;14},{45292;45323;45352;45383})',
  'FORECAST.ETS.STAT': '=FORECAST.ETS.STAT({10;12;16;14},{45292;45323;45352;45383},1)',
  'GETPIVOTDATA': '=GETPIVOTDATA("売上",A1)',
  'GROUPBY': '=GROUPBY(A1:A4,B1:B4,SUM)',
  'IMAGE': '=IMAGE("https://example.invalid/a.png")',
  'LAMBDA': '=LAMBDA(x,x+1)',
  'ODDFPRICE': '=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,0)',
  'ODDFYIELD': '=ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0575,84.5,100,2,0)',
  'ODDLPRICE': '=ODDLPRICE(DATE(2008,2,7),DATE(2008,6,15),DATE(2007,10,15),0.0375,0.0405,100,2,0)',
  'ODDLYIELD': '=ODDLYIELD(DATE(2008,4,20),DATE(2008,6,15),DATE(2007,12,24),0.0375,99.875,100,2,0)',
  'PHONETIC': '=PHONETIC(A1)',
  'PIVOTBY': '=PIVOTBY(A1:A4,B1:B4,C1:C4,SUM)',
  'REGISTER.ID': '=REGISTER.ID("kernel32","GetTickCount","J")',
  'RTD': '=RTD("a",,"b")',
  /* ★式を 置く 所（A9）を 含む 範囲に しない★
     ＝A1:A10 に すると ★自分を 指す★ので #CYCLE に なり、
       「名前が 通らない」かどうかが 見えなく なる（1度 踏んだ） */
  'TRIMRANGE': '=TRIMRANGE(C1:C4)',
};

const 台帳 = fs.readFileSync(path.join(ROOT, 'docs/measured/exally-missing-2026-09-07.txt'), 'utf-8')
  .split('\n').map((s) => s.trim()).filter(Boolean);

console.log('\n[ugokanai-osu] 「動かない」と 数えた 物を 本物の 式で 押す');
console.log('  台帳 … ' + 台帳.length + '個');

T('★台帳の 全部に 押す 式が 用意して 在る★（1つでも 抜けたら そこが 穴）', () => {
  const 無 = 台帳.filter((n) => !式たち[n]);
  if (無.length) throw new Error('★押す 式が 無い★: ' + 無.join(' / '));
});

T('★★押したら 全部 名前が 通らない（それらしい 数が 出ない）★★', () => {
  const 出た = [];
  let 押した = 0;
  for (const n of 台帳) {
    押した++;
    const r = 押す(式たち[n]);
    const s = String(r.出);
    /* ★#NAME? か、書き換え／engine が 名前を 知らずに 投げた 物★ … 正しい */
    if (s === '#NAME?' || s === '#NAME' || s === '#ERROR!' || s === '#ERROR') continue;
    出た.push(n + ' … ' + 式たち[n] + '\n        → ' + s + '（' + r.道 + '）');
  }
  console.log('      押した 式 … ' + 押した + '／' + 台帳.length + '本');
  if (押した !== 台帳.length) throw new Error('押し残しが 有る');
  if (出た.length) {
    throw new Error('★★動かないと 書いてあるのに 何かが 返る★★（CONVERT と 同じ 家）\n      '
      + 出た.join('\n      '));
  }
});

T('★★「動かない」と 数えた 物は 全部 AI の 紙にも 載っている★★（指示役 2026-09-07）', () => {
  /* ★載っていない 関数は AI にとって「知らない 関数」★
     ⇒★知らないので 止めようが ない＝AI が 勧めて しまう★
     ⇒ お客さんが 打つ ⇒ #NAME? ⇒「出来ると 言ったのに 出来ない」
     ⇒★動くのに「動かない」と 言うのと ★同じ 家の 裏返し★★ */
  const EX = require_(path.join(ROOT, 'lib/formula-extra.js'));
  const 棚 = (EX.数える && EX.数える().足さない) || {};
  const 無 = 台帳.filter((n) => !棚[n]);
  if (無.length) throw new Error('★AI の 紙に 載っていない★: ' + 無.join(' / '));
  const 短 = 台帳.filter((n) => String(棚[n]).length < 8);
  if (短.length) throw new Error('★訳が 短すぎる★: ' + 短.join(' / '));
  console.log('      台帳 ' + 台帳.length + '個 … 全部 AI の 紙に 在る');
});

T('★棚に だけ 在る 物は「実Excel にも 無い」物だけ★（数が ずれない）', () => {
  /* ★棚の 方が 多いのは ★実Excel の 一覧に そもそも 無い 名前★の 時だけ★
     （実測 … FIELDVALUE は Excel 16.0 build 20326 の 一覧に 無い） */
  const EX = require_(path.join(ROOT, 'lib/formula-extra.js'));
  const 棚 = Object.keys((EX.数える && EX.数える().足さない) || {});
  const 実 = new Set(fs.readFileSync(path.join(ROOT, 'docs/measured/excel-functions-2026-09-06.txt'), 'utf-8')
    .split(/\r?\n/).map((x) => x.trim()).filter(Boolean));
  const 台 = new Set(台帳);
  const 余 = 棚.filter((n) => !台.has(n) && 実.has(n));
  if (余.length) {
    throw new Error('★実Excel に 在って 動くのに 棚に 載っている★（AIが 客に 嘘を 言う）: ' + 余.join(' / '));
  }
  console.log('      棚 ' + 棚.length + '個 ／ 台帳 ' + 台帳.length + '個'
    + ' ／ 差は 実Excel の 一覧に 無い 名前だけ … '
    + 棚.filter((n) => !台.has(n)).join(' ') );
});

T('★積み忘れが 無い★（積み忘れると 全部 名前が 通らず ★偽の 緑★に なる）', () => {
  if (!積1) throw new Error('exally-formula を 積めていない');
  for (const n of 積) if (!(n > 0)) throw new Error('プラグインを 積めていない（' + 積.join('/') + '）');
  if (typeof JS層 !== 'function') throw new Error('JS層に 手が 届かない');
  /* ★物差しが 生きている 事の 確かめ★＝作った 物は ちゃんと 答えを 返す */
  const v = 押す('=CONVERT(1,"lbm","kg")');
  if (Math.abs(Number(v.出) - 0.45359237) > 1e-12) {
    throw new Error('★作った 物まで 名前が 通らない＝押せていない★（' + v.出 + '）');
  }
});

if (自己試験) {
  console.log('\n[self-test] ★物差しが 効いているか★');
  T('★動く 物を 台帳に 混ぜたら 赤に なる★', () => {
    const r = 押す('=CONVERT(1,"m","cm")');
    const s = String(r.出);
    if (s === '#NAME?' || s === '#ERROR!') throw new Error('動く 物なのに 名前が 通らない');
  });
  T('★前の CONVERT（4桁で 丸めた 物）なら 赤に なる★', () => {
    if (Math.abs(0.4536 - 0.45359237) <= 1e-12) throw new Error('丸めた 数を 見分けられない');
  });
}

console.log('\n' + pass + ' passed, ' + fail + ' failed');
if (直に走った) process.exit(fail ? 1 : 0);
