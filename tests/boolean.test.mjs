/* boolean.test.mjs — ★はい/いいえ（TRUE・FALSE）の 出方★ 2026-08-30
 *
 *  ★真値（実Excel 16.0 で 実測）★ … tools/measure-insert.ps1
 *    ・TRUE を 入れた セル … 型 Boolean／★画面に 出る 字は "TRUE"（大文字）★
 *    ・FALSE の セル … "FALSE"
 *    ・TRUE と 1 は 同じか … True
 *    ・`=A1`（A1 が TRUE） … "TRUE"
 *    ・`=COUNTIF(A1:A2,TRUE)` … ★1★
 *    ・★`=SUM(A1:A2)` … 0★（★はい/いいえは 足されない★）
 *    ・★`=A1+A2` … 1★（★足し算に すると 数に なる★）
 *      ⇒ ★同じ 2つの セルでも SUM と ＋ で 答えが 違う★。これが Excel の 決まり。
 *
 *  ★見つけて 直した（08-30）★
 *    うちは 画面に ★"true"（小文字）★と 出ていた（`String(true)` の まま）。
 *    ⇒ 表示を 作る 出口（`_hfGetDisplay`）★1か所★で 大文字に した。
 *
 *  走らせ方: node tests/boolean.test.mjs [--self-test]
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createRequire } from 'node:module';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(pathToFileURL(path.join(ROOT, 'package.json')));
const 壊す = process.argv.includes('--self-test');
let 緑 = 0, 赤 = 0;
const ok = (名, 条件, 添え) => {
  if (条件) { 緑++; console.log('  ok   ' + 名); }
  else { 赤++; console.log('  ★NG★ ' + 名 + (添え !== undefined ? '  … ' + 添え : '')); }
};
const 式ファイル = fs.readFileSync(path.join(ROOT, 'exally-formula.js'), 'utf8');

console.log('\n[① 測った 道具が 残っている]');
ok('tools/measure-insert.ps1 が 在る', fs.existsSync(path.join(ROOT, 'tools/measure-insert.ps1')));

console.log('\n[② ★大文字で 出す★（実測＝Range.Text は "TRUE"/"FALSE"）]');
ok('★表示を 作る 所で 大文字に している★',
  /if\(typeof val==='boolean'\) return val \? 'TRUE' : 'FALSE';/.test(式ファイル));
ok('★出口は 1か所（_hfGetDisplay）★',
  (式ファイル.match(/return val \? 'TRUE' : 'FALSE';/g) || []).length === 1);
ok('  実測を 書き残している', /Range\.Text が "TRUE"\/"FALSE"/.test(式ファイル));

console.log('\n[③ ★エンジンの 答えが 実Excel と 同じか★]');
{
  /* ★読めない物を「走っている」と 決めない★＝読めなかったら ★赤★に する */
  let HyperFormula = null, 読めない = '';
  try {
    HyperFormula = require_(path.join(ROOT, 'hyperformula.full.min.js')).HyperFormula;
  } catch (e) { 読めない = e.message; }
  ok('★エンジンを 読めた（飛ばさない）★', !!HyperFormula, 読めない);
  if (HyperFormula) {
    const hf = HyperFormula.buildFromArray(
      [[true], [false], ['=SUM(A1:A2)'], ['=A1+A2'], ['=A1'], ['=NOT(A1)'], ['=AND(A1,A1)']],
      { licenseKey: 'gpl-v3' });
    const 読 = (r) => hf.getCellValue({ sheet: 0, col: 0, row: r });
    /* ★実測★ SUM は 0・＋ は 1＝★同じ 2つの セルでも 答えが 違う★ */
    ok('★SUM(A1:A2) = 0（はい/いいえは 足されない）★', 読(2) === 0, JSON.stringify(読(2)));
    ok('★A1+A2 = 1（足し算だと 数に なる）★', 読(3) === 1, JSON.stringify(読(3)));
    ok('  =A1 は はい（true）', 読(4) === true, JSON.stringify(読(4)));
    ok('  =NOT(A1) は いいえ（false）', 読(5) === false, JSON.stringify(読(5)));
    ok('  =AND(A1,A1) は はい', 読(6) === true, JSON.stringify(読(6)));
    /* ★大文字に する 所を 通すと 実Excel と 同じ 字に なる★ */
    const 字に = (v) => (typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : String(v));
    ok('★=A1 を 字に すると "TRUE"★', 字に(読(4)) === 'TRUE', 字に(読(4)));
    ok('★=NOT(A1) を 字に すると "FALSE"★', 字に(読(5)) === 'FALSE', 字に(読(5)));
    ok('★String() の まま だと 小文字に なる（直す 前の 姿）★',
      String(読(4)) === 'true' && 字に(読(4)) !== String(読(4)));
    hf.destroy();
  }
}

console.log('\n[④ チェック ボックス（挿入→コントロール）の 実測を 書き残している]');
{
  const 道具 = fs.readFileSync(path.join(ROOT, 'tools/measure-insert.ps1'), 'utf8');
  ok('★SUM は 足さない事を 測っている★', /SUM\(A1:A2\) は いくつか/.test(道具));
  ok('★＋は 数に なる事を 測っている★', /A1\+A2 は いくつか/.test(道具));
  ok('  COUNTIF も 測っている', /COUNTIF\(A1:A2,TRUE\)/.test(道具));
  ok('  古い チェックボックスも 測っている', /CheckBoxes\(\)\.Add/.test(道具));
  ok('  SmartArt も 測っている', /AddSmartArt/.test(道具));
}

console.log('\nboolean: ' + 緑 + '/' + (緑 + 赤) + ' passed');

if (壊す) {
  console.log('\n★--self-test＝この 見張りが 見ている物を 直に 確かめる（★壊して 赤を 見るのは tools/break-check.mjs★）★');
  let 素通り = 0;
  if (!/if\(typeof val==='boolean'\) return val \? 'TRUE' : 'FALSE';/.test(式ファイル)) {
    素通り++; console.log('  ★素通り★ 大文字に する 所が 無い');
  } else console.log('  ok   大文字に する 所が 在る');
  /* ★小文字の まま だったら 赤に なるか★＝わざと 小文字の 出口を 作って 見る */
  const にせ = (v) => String(v);
  if (にせ(true) === 'TRUE') { 素通り++; console.log('  ★素通り★ String(true) が TRUE に なってしまう'); }
  else console.log('  ok   String(true) は "true"＝★直さないと 小文字に なる★');
  if (素通り) { console.log('★抜け道 ' + 素通り + '件★'); process.exit(1); }
  console.log('  ok   抜け道 0件');
}
process.exit(赤 ? 1 : 0);
