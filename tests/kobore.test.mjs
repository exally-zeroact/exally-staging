/* kobore.test.mjs — ★足した 関数が 2つ以上 返せるか（こぼれるか）★ 2026-09-04
 *
 *  ★これは 新しい 機能では ありません★＝★本番に 前から 在った 穴★です。
 *
 *  ★見つけ方（2026-09-04）★
 *    FILTERXML を 作っている 途中、★2つ以上 返ると #VALUE! に なる★ので 調べたら
 *    ★2026-08-31 に 足した 10本が 前から 同じ★だった。
 *    ★本番（a7a078e）で 打った 実測★
 *      =TAKE(A1:B3,2) → ★#VALUE!★ ／ =TAKE(A1:A3,1) → 1（★1つなら 通る★）
 *      =UNIQUE(A1:A3)（エンジンが 元から 持つ 物）→ 1/3/5 ★こぼれる★
 *      ⇒★10本中 9本が #VALUE!★（MODE.MULT は 別の 訳で 通らない）
 *
 *  ★なぜ 気づかれなかったか★
 *    ★試験が lib の 純関数を 直に 呼んでいた★＝形は 保てている
 *    ★繋ぐ 側は「形を 保って 返している」を ★字合わせ★で 見ているだけ★
 *    ⇒★エンジンに 通して 打っていなかった★（★lib 緑で 完成と するな★）
 *
 *  ★直し方★
 *    エンジンの 決まり＝配列を 返す 関数は ★sizeOfResultArrayMethod★（大きさを 先に 言う）が 要る。
 *    エンジン自身の UNIQUE / SORT / FILTER も そう 宣言している（2026-09-04 中を 読んで 確かめた）。
 *
 *  ★この 試験は エンジンに 通して 打つ★（★字合わせでは ない★）
 *
 *  走らせ方: node tests/kobore.test.mjs [--self-test]
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createRequire } from 'node:module';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(pathToFileURL(path.join(ROOT, 'package.json')));

let pass = 0, fail = 0;
const T = (n, よい, 添え) => {
  if (よい) { pass++; console.log('  ok   ' + n); }
  else { fail++; console.log('  NG   ' + n + (添え ? '\n       ' + 添え : '')); }
};

const M = require_(path.join(ROOT, 'hyperformula.full.min.js'));
const HF = M.HyperFormula;
/* ★画面では window.HyperFormula が 全部 持っている★＝node では 同じ 形に 混ぜる */
const H = Object.assign(Object.create(HF), M,
  { registerFunctionPlugin: HF.registerFunctionPlugin.bind(HF) });
const P = require_(path.join(ROOT, 'lib/formula-extra-plug.js'));
const 数 = P.つなぐ(H, require_(path.join(ROOT, 'lib/formula-extra.js')));

console.log('');
console.log('[kobore] ★足した 関数が 2つ以上 返せるか（エンジンに 通して 打つ）★');
T('★繋がった（0本では ない）★', 数 > 0, '繋いだ ' + 数 + '本');

/** ★式を 打って 返った 大きさと 中身を 数える★ */
function 打つ(式) {
  const 表 = [[1, 2, null, null, null], [3, 4, null, null, null], [5, 6, null, null, null],
    [null, null, null, null, null], [null, null, null, null, null], [null, null, null, null, null]];
  表[0][2] = 式;
  const hf = HF.buildFromArray(表, { licenseKey: 'gpl-v3' });
  const 読 = (r, c) => { const v = hf.getCellValue({ sheet: 0, col: c, row: r });
    return (v && v.value !== undefined) ? v.value : v; };
  let 行 = 0, 列 = 0;
  for (let r = 0; r < 6; r++) { if (読(r, 2) === null || 読(r, 2) === undefined) break; 行++; }
  for (let c = 2; c < 8; c++) { if (読(0, c) === null || 読(0, c) === undefined) break; 列++; }
  return { 行: 行, 列: 列, 先頭: 読(0, 2) };
}

/* ★実Excel と 同じ 形に なるか★（形は 実Excel の 決まりどおり）
   ★元の 表★ A1:B3 = 1,2 / 3,4 / 5,6 */
const 組 = [
  ['TAKE', '=TAKE(A1:B3,2)', 2, 2, 1],
  ['DROP', '=DROP(A1:B3,1)', 2, 2, 3],
  ['CHOOSECOLS', '=CHOOSECOLS(A1:B3,1)', 3, 1, 1],
  ['CHOOSEROWS', '=CHOOSEROWS(A1:B3,1,2)', 2, 2, 1],
  ['TOCOL', '=TOCOL(A1:B3)', 6, 1, 1],
  ['TOROW', '=TOROW(A1:B3)', 1, 6, 1],
  ['WRAPROWS', '=WRAPROWS(A1:A3,2,0)', 2, 2, 1],
  ['WRAPCOLS', '=WRAPCOLS(A1:A3,2,0)', 2, 2, 1],
  ['EXPAND', '=EXPAND(A1:B3,4,3,0)', 4, 3, 1],
];
let 合 = 0;
for (const [名, 式, 行, 列, 先] of 組) {
  const r = 打つ(式);
  const よい = (r.行 === 行 && r.列 === 列 && r.先頭 === 先);
  if (よい) 合++;
  T('★' + 名 + ' … ' + 行 + '行' + 列 + '列 で こぼれる★',
    よい, 式 + ' → ' + r.行 + '行' + r.列 + '列 先頭 ' + JSON.stringify(r.先頭));
}
console.log('       … 打った 総数 ' + 組.length + '本 ／ ★こぼれた ' + 合 + '本★'
  + '（★本番 a7a078e では 0本★）');

/* ★1つだけ 返る 物は 前から 通っていた＝壊していないか★ */
{
  const a = 打つ('=TAKE(A1:B3,1,1)');
  T('★1つだけ 返る 形も そのまま★', a.行 === 1 && a.列 === 1 && a.先頭 === 1,
    JSON.stringify(a));
  const b = 打つ('=AVERAGEIFS(A1:A3,B1:B3,">1")');
  T('★1つ返しの AVERAGEIFS は そのまま★', b.行 === 1 && b.列 === 1 && b.先頭 === 3,
    JSON.stringify(b));
}

/* ★★まだ 直っていない 物を 隠さない★★ */
{
  const hf = HF.buildFromArray([[1, 2, '=MODE.MULT(A1:B3)'], [1, 2, null], [3, 9, null]],
    { licenseKey: 'gpl-v3' });
  const v = hf.getCellValue({ sheet: 0, col: 2, row: 0 });
  const 字 = (v && v.value !== undefined) ? String(v.value) : String(v);
  console.log('       … ★MODE.MULT（重なり 在り）は まだ ' + 字 + '★'
    + (v && v.message ? ('「' + v.message + '」') : '')
    + '（★本番でも 同じ＝直っていない／別の 一件★）');
  T('★MODE.MULT が まだ 直っていない事を 隠していない★', 字 === '#VALUE!',
    '★直ったなら この 試験を 書き替える（緑に する）★');
  const hf2 = HF.buildFromArray([[1, 2, '=MODE.MULT(A1:B3)'], [3, 4, null], [5, 6, null]],
    { licenseKey: 'gpl-v3' });
  const w = hf2.getCellValue({ sheet: 0, col: 2, row: 0 });
  T('★重なりが 無い時の #N/A は 実Excel と 同じ★',
    ((w && w.value !== undefined) ? String(w.value) : String(w)) === '#N/A');
}

/* ── わざと 壊して 赤に なるか ── */
if (process.argv.includes('--self-test')) {
  console.log('');
  console.log('★大きさの 宣言を 外して 赤に なるか★');
  const { execFileSync } = await import('node:child_process');
  const 道 = path.join(ROOT, 'lib/formula-extra-plug.js');
  const 元 = fs.readFileSync(道, 'utf8');
  const 壊す = [
    ['★大きさの 宣言を 全部 外す（前の 姿に 戻す）★',
      (t) => t.replace(/sizeOfResultArrayMethod: '大きさ[A-Z.]+',\s*/g, '')],
    ['★大きさを いつも 1×1 と 言う★',
      (t) => t.replace('return new ArraySize(幅, 高);', 'return ArraySize.scalar();')],
    ['★TAKE の 宣言だけ 外す★',
      (t) => t.replace("sizeOfResultArrayMethod: '大きさTAKE',        ", '')],
  ];
  for (const [名, f] of 壊す) {
    const 壊れ = f(元);
    if (壊れ === 元) { console.log('  ★素通り★  ' + 名 + '（印が 古い＝直せ）'); fail++; continue; }
    fs.writeFileSync(道, 壊れ);
    let 赤 = false;
    try { execFileSync(process.execPath, [path.join(ROOT, 'tests', 'kobore.test.mjs')], { stdio: 'pipe' }); }
    catch (e) { 赤 = true; }
    fs.writeFileSync(道, 元);                 /* ★必ず 戻す★ */
    console.log((赤 ? '  赤くなった  ' : '  ★素通り★  ') + 名);
    if (!赤) fail++;
  }
  T('★本物は 壊していない（戻した）★', fs.readFileSync(道, 'utf8') === 元);
}

console.log('');
console.log('kobore: ' + pass + ' 緑 / ' + fail + ' 赤');
process.exit(fail ? 1 : 0);
