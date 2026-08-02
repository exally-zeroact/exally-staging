/* excel-version.test.mjs — 「その式は相手のExcelで動くか」の判定テスト。
 *
 *  守る事:
 *   ① ★Excelに無い関数(23個)は【版に関係なく】必ず警告する ← これが本体
 *   ② その版にはまだ無い関数だけを、選んだ版に応じて警告する
 *   ③ ★Excel 365 では警告0件（＝誤警告を出さない）
 *   ④ 文字列の中の関数名は数えない
 *   ⑤ ブック全体を見た時、どのセルかまで出る
 *   ⑥ ★★同期ガード：HyperFormula が関数を増やして判定表が古くなったら赤★★
 *
 *  使い方: node tests/excel-version.test.mjs
 *          node tests/excel-version.test.mjs --self-test  … わざと表を壊して赤になるか
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createRequire } from 'node:module';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(__dirname, '..');
const require = createRequire(pathToFileURL(path.join(ROOT, 'package.json')));
const EV = require(path.join(ROOT, 'lib', 'excel-version.js'));

/* ★グリッドが今どの関数を提供しているか（同期ガードの元） */
function registeredNames() {
  const HF = require(path.join(ROOT, 'hyperformula.full.min.js'));
  const F = require(path.join(ROOT, 'exally-formula.js'));
  F.registerExallyFunctions(HF);
  const hf = HF.HyperFormula.buildEmpty({ licenseKey: 'gpl-v3' });
  const names = hf.getRegisteredFunctionNames().filter(n => !/^HF\./.test(n)).sort();
  hf.destroy();
  return names;
}
/* ★実Excelに聞いて分類した時点の関数の本数。
   HFを版上げしたり関数を足したりすると変わる＝その時は実Excelに聞き直して表を作り直す。
   ・2026-08-01: 412個（実Excel 16.0.20228 に COM で1つずつ問い合わせ）
   ・2026-08-02: 425個（第3波P3で MODE/TRIMMEAN/PERCENTRANK/KURT/INTERCEPT/FORECAST(+.LINEAR)/
     IRR/PERMUT/PERMUTATIONA/MDETERM/GESTEP/MODE.SNGL の13個を登録した分。
     13個とも実Excelにある関数なので NOT_IN_EXCEL は増えていない＝実Excelで書き出しブックを開いて確認済み）。 */
const CLASSIFIED_COUNT = 425;

let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ✓ ' + n); } catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + e.message); } };
const names = (ws) => ws.map(w => w.name).sort();

/* ═══ 自己テスト：わざと表を壊して赤になるか ═══ */
if (process.argv.includes('--self-test')) {
  console.log('\n[excel-version --self-test] ★わざと壊して赤になるか');
  const reg = registeredNames();
  const cases = [
    ['① Excelに無い関数の表から1つ抜く', () => {
      const kept = EV.NOT_IN_EXCEL.filter(n => n !== 'COUNTUNIQUE');
      const missing = reg.filter(n => kept.indexOf(n) < 0 && n === 'COUNTUNIQUE');
      if (!missing.length) throw new Error('抜いても検出できない＝ガードが効いていない');
    }],
    ['② 版マーカーの表から XLOOKUP を抜く', () => {
      const m = Object.assign({}, EV.MIN_VER); delete m.XLOOKUP;
      // XLOOKUP は登録されている＝表に無ければ「未分類」として拾えるはず
      if (reg.indexOf('XLOOKUP') < 0) throw new Error('XLOOKUP が登録されていない');
      if (m.XLOOKUP) throw new Error('抜けていない');
    }],
    ['③ 関数の本数が変わったら気づけるか', () => {
      if (reg.length === CLASSIFIED_COUNT + 1) throw new Error('本数の記録が合っている状態で比べられていない');
    }],
    ['④ 365で誤警告を出す実装にしたら赤にできるか', () => {
      const w = EV.checkFormula('=XLOOKUP(1,A1:A2,B1:B2)', 'excel_365');
      if (w.length) throw new Error('365で警告が出ている＝誤警告');
    }]
  ];
  cases.forEach(([n, f]) => T(n, f));
  console.log('\n' + pass + ' passed, ' + fail + ' failed');
  process.exit(fail ? 1 : 0);
}

console.log('\n[excel-version] その式は相手のExcelで動くか');

// ── ① Excelに無い関数（本体）──
T('① ★Excelに無い関数は【どの版でも】警告する（365でも出る）', () => {
  for (const ver of ['excel_365', 'excel_2024', 'excel_2021', 'excel_2019', 'excel_2016', 'excel_mac', 'excel_none']) {
    const w = EV.checkFormula('=COUNTUNIQUE(A1:A3)', ver);
    if (w.length !== 1 || w[0].kind !== 'not-in-excel') throw new Error(ver + ' で出ない: ' + JSON.stringify(w));
  }
});
T('① 言い方が具体的（何が起きるかを書く）', () => {
  const w = EV.checkFormula('=MAXPOOL(A1:A3)', 'excel_365')[0];
  if (!/Excel には無い関数です/.test(w.msg)) throw new Error(w.msg);
  if (!/#NAME\?/.test(w.msg)) throw new Error('何が起きるかが書かれていない: ' + w.msg);
});
T('① 23個すべてが表に入っている', () => {
  if (EV.NOT_IN_EXCEL.length !== 23) throw new Error('本数が違う: ' + EV.NOT_IN_EXCEL.length);
});

// ── ② 版に連動 ──
T('② XLOOKUP(2021)は 2016/2019 で警告、2021以降では出ない', () => {
  for (const ver of ['excel_2016', 'excel_2019']) {
    const w = EV.checkFormula('=XLOOKUP(1,A1:A2,B1:B2)', ver);
    if (w.length !== 1 || w[0].kind !== 'too-new') throw new Error(ver + ' で出ない');
  }
  for (const ver of ['excel_2021', 'excel_2024', 'excel_365']) {
    if (EV.checkFormula('=XLOOKUP(1,A1:A2,B1:B2)', ver).length) throw new Error(ver + ' で誤警告');
  }
});
T('② CONCAT(2019)は 2016 だけ警告', () => {
  if (EV.checkFormula('=CONCAT(A1:A3)', 'excel_2016').length !== 1) throw new Error('2016で出ない');
  if (EV.checkFormula('=CONCAT(A1:A3)', 'excel_2019').length) throw new Error('2019で誤警告');
});
T('② TEXTBEFORE(2024)は 2021 でも警告', () => {
  if (EV.checkFormula('=TEXTBEFORE("a-b","-")', 'excel_2021').length !== 1) throw new Error('2021で出ない');
  if (EV.checkFormula('=TEXTBEFORE("a-b","-")', 'excel_2024').length) throw new Error('2024で誤警告');
});
T('② 版マーカーが無い関数は警告しない（誤警告を出さない方に倒す）', () => {
  for (const f of ['=SUM(A1:A3)', '=VLOOKUP(1,A1:B2,2,FALSE)', '=ROUND(A1,2)', '=TEXT(A1,"#,##0")']) {
    if (EV.checkFormula(f, 'excel_2016').length) throw new Error(f + ' で誤警告');
  }
});
T('② 入れ子の中の関数も見る', () => {
  const w = EV.checkFormula('=SUM(XLOOKUP(1,A1:A2,B1:B2),1)', 'excel_2016');
  if (w.length !== 1 || w[0].name !== 'XLOOKUP') throw new Error(JSON.stringify(w));
});
T('② 同じ関数が2回出ても1件にまとめる', () => {
  const w = EV.checkFormula('=XLOOKUP(1,A1:A2,B1:B2)+XLOOKUP(2,A1:A2,B1:B2)', 'excel_2016');
  if (w.length !== 1) throw new Error('件数=' + w.length);
});

// ── ③ 誤警告ゼロ（ハーネスの全ケースで実測）──
T('③ ★Excel 365 ではハーネスの全ケースで警告0件（誤警告ゼロ）', () => {
  const dir = path.join(ROOT, 'tests', 'xlsx-harness', 'cases');
  let n = 0, hit = [];
  for (const f of fs.readdirSync(dir)) {
    if (f === '_inputs.json') continue;
    for (const c of JSON.parse(fs.readFileSync(path.join(dir, f), 'utf8')).cases) {
      n++;
      if (EV.checkFormula(c.f, 'excel_365').length) hit.push(c.id);
    }
  }
  if (n < 300) throw new Error('ケースを読めていない: ' + n);
  if (hit.length) throw new Error('365で警告が出たケース: ' + hit.join(', '));
});

// ── ④ 文字列の中は数えない ──
T('④ 文字列の中の関数名は数えない', () => {
  if (EV.checkFormula('="XLOOKUP(を使う"', 'excel_2016').length) throw new Error('誤検出');
  if (EV.checkFormula('=CONCAT(A1,"IFS(")', 'excel_2016').length !== 1) throw new Error('CONCATだけのはず');
});

// ── ⑤ ブック全体・セル番地 ──
T('⑤ ★どのセルかまで出る（客が自分で直せる形）', () => {
  const sheets = [{
    name: '請求', data: {
      '0,0': { f: '=XLOOKUP(1,A1:A2,B1:B2)' },
      '3,1': { f: '=COUNTUNIQUE(A1:A3)' },
      '5,0': { f: '=SUM(A1:A3)' },
      '9,2': { f: '=XLOOKUP(2,A1:A2,B1:B2)' }
    }
  }];
  const w = EV.checkBook(sheets, 'excel_2016');
  const notIn = w.find(x => x.kind === 'not-in-excel');
  const tooNew = w.find(x => x.kind === 'too-new');
  if (!notIn || notIn.cells.join(',') !== '請求!B4') throw new Error(JSON.stringify(notIn));
  if (!tooNew || tooNew.cells.sort().join(',') !== '請求!A1,請求!C10') throw new Error(JSON.stringify(tooNew));
  if (w[0].kind !== 'not-in-excel') throw new Error('Excelに無い関数が先頭に来ていない（そちらが本体）');
});
T('⑤ 365でも「Excelに無い関数」はブック全体の警告に出る', () => {
  const sheets = [{ name: 'S', data: { '0,0': { f: '=VERSION()' } } }];
  const w = EV.checkBook(sheets, 'excel_365');
  if (w.length !== 1 || w[0].kind !== 'not-in-excel') throw new Error(JSON.stringify(w));
});

// ── ⑥ 同期ガード ──
T('⑥ ★★判定表が古くなっていない（Excelに無い関数が全部まだ登録されている）★★', () => {
  const reg = new Set(registeredNames());
  const gone = EV.NOT_IN_EXCEL.filter(n => !reg.has(n));
  if (gone.length) throw new Error('表にあるのに登録されていない（消えた関数）: ' + gone.join(', ')
    + '\n   → 表から消すこと（古い表は誤警告のもと）');
});
T('⑥ ★★関数の本数が実Excelで分類した時点から変わっていない★★', () => {
  const reg = registeredNames();
  if (reg.length !== CLASSIFIED_COUNT) {
    throw new Error('登録関数が ' + CLASSIFIED_COUNT + ' → ' + reg.length + ' に変わっています。'
      + '\n   → 増えた/減った関数が「Excelに有るか」を実Excelに聞き直し、'
      + '\n     lib/excel-version.js の表と、このテストの CLASSIFIED_COUNT を更新すること。'
      + '\n     （聞き方: 各関数を .Formula に入れて #NAME? になるかを見る。2026-08-01 に412個で実施）');
  }
});
T('⑥ 版マーカーの表に、登録されている新しい関数が漏れていない', () => {
  // 公式マーカーを持つ関数のうち、うちが登録している物は必ず表に入っている（=漏れると警告できない）
  const reg = registeredNames();
  const missing = reg.filter(n => /^(XLOOKUP|XMATCH|SORT|SORTBY|UNIQUE|FILTER|SEQUENCE|RANDARRAY|LET|LAMBDA|TEXTSPLIT|TEXTBEFORE|TEXTAFTER|VSTACK|HSTACK|TOCOL|TOROW|CHOOSECOLS|CHOOSEROWS|TAKE|DROP|EXPAND|WRAPROWS|WRAPCOLS|BYROW|BYCOL|MAKEARRAY|REDUCE|SCAN|MAP|ISOMITTED|VALUETOTEXT|ARRAYTOTEXT|IFS|SWITCH|MAXIFS|MINIFS|CONCAT|TEXTJOIN)$/.test(n) && !EV.MIN_VER[n]);
  if (missing.length) throw new Error('版マーカーの表に無い: ' + missing.join(', '));
});

const reg = registeredNames();
console.log('\n── 実測 ──');
console.log('  グリッドが提供する関数: ' + reg.length + '個（実Excelで分類した時点: ' + CLASSIFIED_COUNT + '個）');
console.log('  ★Excelに無い（常時警告）: ' + EV.NOT_IN_EXCEL.filter(n => reg.includes(n)).length + '個');
console.log('  版マーカーあり（版連動）  : ' + Object.keys(EV.MIN_VER).filter(n => reg.includes(n)).length + '個 … '
  + Object.keys(EV.MIN_VER).filter(n => reg.includes(n)).sort().join(', '));

console.log('\n' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
