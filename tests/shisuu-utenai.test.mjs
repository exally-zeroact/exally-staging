/* shisuu-utenai.test.mjs — ★指数の 字を 打つと #ERROR に なって いた★（2026-09-10）
 *
 *  ★★何が 起きて いたか★★
 *    お客さんが ★指数の 字を 打つと #ERROR★ に なって いました。
 *      =1.64E-14 ／ =1E+20 ／ =1.5E+20 ／ =2E5 ／ =1E-5 ／ =1E-4 … ★全部 #ERROR★
 *      ★=1.64e-14（小文字）は 通る★
 *    ⇒★エンジンの 式読みが ★大文字 E を 数と 読めない★★
 *    ⇒★★「見えない」より 悪い＝★打てない★★
 *
 *  ★★「E14 という セル参照と ぶつかって いる」は ★外れ★でした★★
 *    押して 確かめました
 *      `=E14` → 999 ／ `=SUM(E1:E14)` → 1002 ★どちらも 正しく 動く★
 *    ⇒ ぶつかって いません。★大文字 E を 数として 読まないだけ★
 *    ⇒★見当を 直す 前に 押した★（★見当で 直さない★）
 *
 *  ★★一番 危ないのは ★セル参照を 壊す 事★★★
 *    ⇒★だから この 試験は ★壊れる 側から★ 書いて あります★
 *    ⇒★触っては いけない 物★ … `=E14` `=SUM(E1:E14)` `=$E$14` `=Sheet1!E14`
 *                              `="1.64E-14"`（字の 中）／`=TEXT(A1,"0.00E+00")`（書式の 中）
 *                              `=A1E5`（前に 英字）／`=EXP(2)`（関数名）
 *
 *  ★★見て いない 範囲（★書かない 見張りは「全部 守った」と 読まれる★）★★
 *    ・★出る 字は 直して いません★
 *        うち `1.64e-14`（小文字 e）／実Excel `1.64E-14`（大文字 E・仮数6桁・指数2桁）
 *        ⇒★別の 直し（見た目）★／★この PR では 直しません★
 *    ・★一重引用符（'）の 中は 見て いません★（シート名の 囲み）
 *    ・★実Excel が どこから 指数に するか（境目）は 別に 測ります★
 *      （0.000001 は そのまま／0.0000001 も そのまま＝★幅で 変わる 見込み・未測定★）
 *
 *  使い方: node tests/shisuu-utenai.test.mjs [--self-test]
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createRequire } from 'node:module';

const ここ = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(ここ, '..');
const 直に走った = process.argv[1] && pathToFileURL(process.argv[1]).href === import.meta.url;
const 自己試験 = 直に走った && process.argv.includes('--self-test');

let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ✓ ' + n); } catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + (e && e.message)); } };

const require_ = createRequire(path.join(ROOT, 'package.json'));
const HFns = require_(path.join(ROOT, 'hyperformula.full.min.js'));
const EF = require_(path.join(ROOT, 'exally-formula.js'));
EF.registerExallyFunctions(HFns);
const hf = HFns.HyperFormula.buildEmpty({
  licenseKey: 'gpl-v3', useArrayArithmetic: true, smartRounding: false,
});
const SID = hf.getSheetId(hf.addSheet('S'));
EF.initExallyFormula(hf);

/* ★材料★ … E14=999 ／ E1=1 ／ E2=2 ／ A1=5 */
const 板 = [];
for (let r = 0; r < 22; r++) 板.push([null, null, null, null, null, null]);
板[13][4] = 999; 板[0][4] = 1; 板[1][4] = 2; 板[0][0] = 5;
const 式の行 = 20;

/* ★本番と 同じ 道★ ①JS層 → ②convertFormula → エンジン */
function 押す(f) {
  hf.setSheetContent(SID, 板.map((r) => r.slice()));
  const js = EF._jsComputeFormula(0, f);
  if (js !== null && js !== undefined) return String(js);
  const 盤 = 板.map((r) => r.slice());
  盤[式の行][0] = EF.convertFormula(f);
  hf.setSheetContent(SID, 盤);
  const v = hf.getCellValue({ sheet: SID, row: 式の行, col: 0 });
  if (v && v.type) return '#' + v.type;
  return v === null || v === undefined ? '(空)' : String(v);
}

console.log('\n[shisuu-utenai] ★指数の 字を 打つと #ERROR に なって いた★');

T('★★指数の 字が 打てる（★前は 全部 #ERROR★）★★', () => {
  const 組 = [
    ['=1.64E-14', '1.64e-14'],
    ['=1E+20', '100000000000000000000'],
    ['=1.5E+20', '150000000000000000000'],
    ['=2E5', '200000'],
    ['=1.5E2', '150'],
    ['=1E-5', '0.00001'],
    ['=1E-4', '0.0001'],
    ['=1.640928159047308E-14', '1.640928159047308e-14'],
  ];
  for (const [f, 正] of 組) {
    const r = 押す(f);
    if (r === '#ERROR') throw new Error('★' + f + ' が #ERROR★＝★お客さんが 打てない★');
    if (String(r) !== 正) throw new Error('★' + f + ' … うち ' + r + ' ／ 期待 ' + 正 + '★');
  }
  console.log('      … ' + 組.length + '通り とも 打てる');
});

T('★★★セル参照を 1つも 壊して いない（★一番 危ない所★）★★★', () => {
  const 組 = [
    ['=E14', '999', '★E14＝ふつうの セル★'],
    ['=SUM(E1:E14)', '1002', '★範囲★'],
    ['=$E$14', '999', '★絶対★'],
    ['=E1+E14', '1000', '★足し算★'],
    ['=E14*1E5', '99900000', '★セルと 指数が 混ざる★'],
  ];
  for (const [f, 正, 札] of 組) {
    const r = 押す(f);
    if (String(r) !== 正) {
      throw new Error('★' + 札 + '（' + f + '）… うち ' + r + ' ／ 期待 ' + 正 + '★'
        + ' ⇒★セル参照を 壊すのが 一番 危ない★');
    }
  }
  console.log('      … ' + 組.length + '通り とも 無事');
});

T('★★字の 中・書式の 中を 触って いない★★', () => {
  const 組 = [
    ['="1.64E-14"', '1.64E-14', '★字の 中＝大文字の まま★'],
    ['=TEXT(1234,"0.00E+00")', null, '★書式の 中★'],
  ];
  const r1 = 押す(組[0][0]);
  if (r1 !== '1.64E-14') {
    throw new Error('★字の 中を 触った（' + r1 + '）★＝★お客さんが 打った 字が 変わる★');
  }
  /* ★書式は 直した 字が エンジンに 渡らない 事だけ 見る★（答えは 別の 話） */
  const 変換 = EF.convertFormula('=TEXT(1234,"0.00E+00")');
  if (変換.indexOf('0.00E+00') < 0) {
    throw new Error('★書式の 中の E を 小文字に した（' + 変換 + '）★＝★書式が 壊れる★');
  }
  console.log('      … 字の 中も 書式の 中も 大文字の まま');
});

T('★前に 英字が 在る 物は 触らない（★セルの 名前かも★）★', () => {
  for (const f of ['=A1E5', '=AE5', '=EXP(0)']) {
    const 変換 = EF.convertFormula(f);
    if (変換 !== f) throw new Error('★' + f + ' を ' + 変換 + ' に 変えた★');
  }
  console.log('      … 3通り とも そのまま');
});

T('★★ぶつかりの 見当は 外れて いた（★押して 確かめた★）★★', () => {
  /* ★「E14 は セル参照だから ぶつかる」は 外れ★
     ⇒ 素の エンジンでも `=E14` は 動く＝★大文字 E を 数と 読まないだけ★ */
  const 素 = HFns.HyperFormula.buildEmpty({ licenseKey: 'gpl-v3', smartRounding: false });
  const s = 素.getSheetId(素.addSheet('T'));
  const b = [];
  for (let r = 0; r < 20; r++) b.push([null, null, null, null, null]);
  b[13][4] = 999;
  b[18][0] = '=E14';
  素.setSheetContent(s, b);
  const v = 素.getCellValue({ sheet: s, row: 18, col: 0 });
  if (Number(v) !== 999) {
    throw new Error('★素の エンジンで =E14 が ' + JSON.stringify(v) + '★＝★見当が 当たって いた事に なる★');
  }
  console.log('      … 素の エンジンでも =E14 → 999（★ぶつかって いない★）');
});

/* ══ ★自己試験（★壊すのは 写し★＝ファイルは 1バイトも 触らない）★ ══ */
if (自己試験) {
  console.log('\n★自己試験（★壊すのは 写し★＝ファイルは 1バイトも 触らない）★');

  T('★★書き換えが コードに 在る（空振りして いない）★★', () => {
    const s = fs.readFileSync(path.join(ROOT, 'exally-formula.js'), 'utf-8');
    if (!/function _指数のEを小文字に/.test(s)) throw new Error('★書き換えが 無い★');
    if (!/f = _指数のEを小文字に\(f\);/.test(s)) throw new Error('★convertFormula から 呼んで いない★');
    console.log('      … 書き換えと 呼び出しの 両方が 在る');
  });

  T('★★書き換えを 外すと #ERROR に 戻る（＝効いて いる 証拠）★★', () => {
    /* ★素の エンジンに 直に 渡す＝書き換えを 通さない★ */
    const 素 = HFns.HyperFormula.buildEmpty({ licenseKey: 'gpl-v3', smartRounding: false });
    const s = 素.getSheetId(素.addSheet('T'));
    const b = [[null], [null], [null]];
    b[0][0] = '=1.64E-14';
    素.setSheetContent(s, b);
    const v = 素.getCellValue({ sheet: s, row: 0, col: 0 });
    const 字 = (v && v.type) ? ('#' + v.type) : String(v);
    if (字.indexOf('ERROR') < 0) {
      throw new Error('★書き換え無しでも 通る（' + 字 + '）★＝★この 直しは 要らなかった 事に なる★');
    }
    console.log('      … 書き換え無しでは ' + 字 + '（★だから 要る★）');
  });

  T('★出る 字は まだ 直して いない と 書いて 在る（★未完を 緑に しない★）★', () => {
    const s = fs.readFileSync(path.join(ここ, 'shisuu-utenai.test.mjs'), 'utf-8');
    if (!/出る 字は 直して いません/.test(s)) {
      throw new Error('★「出る 字は 直して いない」の 断りが 消えた★');
    }
    /* ★実際 小文字の まま 出る事も 押して 確かめる★ */
    const r = 押す('=1.64E-14');
    if (String(r) !== '1.64e-14') throw new Error('★' + r + '★（1.64e-14 のはず）');
    console.log('      … 断りも 在り、実際 小文字の まま（実Excel は 1.64E-14）');
  });
}

console.log('\n' + pass + ' passed, ' + fail + ' failed');
if (fail) process.exit(1);
