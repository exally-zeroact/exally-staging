/* awaseru2.mjs — ★予測・統計／表・情報 を「本番の 道」で 押して 実Excel と 比べる★（2026-09-07）
 *
 *  ★★まず 疑う★★
 *    台帳（exally-missing）は ★決まった 引数の 形で 押した 結果★です。
 *    ⇒★形が 合わないだけで「動かない」に 見える★事が 有る。
 *    ⇒ だから ★本物の 式★（実Excel に 打たせたのと 同じ 字）で 押し直す。
 *
 *  ★本番の 道★＝ JS層(_jsComputeFormula) → convertFormula → HyperFormula
 *    ★プラグインは 本番と 同じ 4つ 全部 積む★（1つ 忘れると 嘘の 数に なる）
 *
 *  使い方: node docs/measured/kansuu46/awaseru2.mjs [関数名]
 */
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';

const ここ = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(ここ, '..', '..', '..');
const require_ = createRequire(import.meta.url);

const HFns = require_(path.join(ROOT, 'hyperformula.full.min.js'));
const EF = require_(path.join(ROOT, 'exally-formula.js'));
EF.registerExallyFunctions(HFns);
const HF0 = HFns.HyperFormula;
const H = Object.assign(Object.create(HF0), HFns,
  { registerFunctionPlugin: HF0.registerFunctionPlugin.bind(HF0) });
require_(path.join(ROOT, 'lib/formula-extra-plug.js')).つなぐ(H, require_(path.join(ROOT, 'lib/formula-extra.js')));
require_(path.join(ROOT, 'lib/formula-nokori-plug.js')).つなぐ(H, require_(path.join(ROOT, 'lib/formula-nokori.js')));
require_(path.join(ROOT, 'lib/formula-kane-plug.js')).つなぐ(H, require_(path.join(ROOT, 'lib/formula-kane.js')));
require_(path.join(ROOT, 'lib/formula-yosoku-plug.js')).つなぐ(H, require_(path.join(ROOT, 'lib/formula-yosoku.js')));

const hf = HF0.buildEmpty({ licenseKey: 'gpl-v3' });
const SID = hf.getSheetId(hf.addSheet('S'));
EF.initExallyFormula(hf);
const JS層 = EF._jsComputeFormula || null;

/* ★エンジンの 赤の 名前を 実Excel の 書き方に 直す★
   （エンジンは NA・DIV_BY_ZERO、実Excel は #N/A・#DIV/0!） */
const 赤の名 = (t) => ({
  NA: '#N/A', DIV_BY_ZERO: '#DIV/0!', VALUE: '#VALUE!', NUM: '#NUM!',
  NAME: '#NAME?', REF: '#REF!', CYCLE: '#CYCLE!', ERROR: '#ERROR!', SPILL: '#SPILL!',
}[t] || ('#' + t));

/* ★下敷きの 表★＝A1:D8 に 数を 置く（CELL・AREAS が 見る 所） */
const 下敷き = [];
for (let r = 0; r < 8; r++) 下敷き.push([r + 1, (r + 1) * 2, (r + 1) * 3, (r + 1) * 4]);

function 押す(式) {
  if (typeof JS層 === 'function') {
    try {
      const v = JS層(0, 式);
      if (v !== null && v !== undefined) return { 出: v, 道: 'JS層' };
    } catch (e) { return { 出: '★例外★ ' + e.message, 道: 'JS層' }; }
  }
  let 後;
  try { 後 = EF.convertFormula(式); } catch (e) { return { 出: '★書き換えで 例外★', 道: '—' }; }
  try {
    /* ★こぼれる 場所を 空けておく★＝縦に 伸びる 答え（TREND 等）は
       置き場が 無いと エンジンが 赤を 返す＝★式の 所為に 見える★ */
    const 表 = 下敷き.map((r) => r.slice());
    表.push([後]);
    /* ★空文字は「中身が 在る」扱い＝こぼれる 場所を 塞ぐ★（#SPILL! に なる）
       ⇒ null を 置く */
    for (let i = 0; i < 40; i++) 表.push([null]);
    hf.setSheetContent(SID, 表);
    const v = hf.getCellValue({ sheet: SID, row: 下敷き.length, col: 0 });
    if (v && v.type) return { 出: 赤の名(v.type), 道: 'engine' };
    return { 出: v, 道: 'engine' };
  } catch (e) { return { 出: '★例外★ ' + e.message, 道: 'engine' }; }
}

const 絞り = (process.argv[2] || '').toUpperCase();
const 金 = fs.readFileSync(path.join(ここ, 'golden-yosoku-2026-09-07.tsv'), 'utf-8')
  .split('\n').filter((l) => l && !l.startsWith('#'))
  .map((l) => { const p = l.split('\t'); return { 名: p[0], 式: p[1], 答: p[2], 型: p[3] }; });

const 束 = new Map();
for (const g of 金) {
  if (絞り && g.名 !== 絞り) continue;
  const r = 押す(g.式);
  let 合 = false;
  if (g.型 === 'Double' && typeof r.出 === 'number') {
    const 正 = Number(g.答);
    合 = Math.abs(r.出 - 正) <= Math.max(1e-9, Math.abs(正) * 1e-9);
  } else if (String(r.出) === String(g.答)) 合 = true;
  else if (g.型 === 'Boolean' && String(r.出).toUpperCase() === String(g.答).toUpperCase()) 合 = true;
  const s = 束.get(g.名) || { 合: 0, 違: 0, 例: [] };
  if (合) s.合++; else { s.違++; if (s.例.length < 3) s.例.push('      ' + g.式 + '\n        正 ' + g.答 + ' ／ 出 ' + r.出 + '（' + r.道 + '）'); }
  束.set(g.名, s);
}

console.log('\n[awaseru2] 本番の 道で 押して 実Excel と 比べる');
console.log('  ★答えは 実Excel（16.0 build 20326 / UI 1041）★\n');
const 並 = [...束.entries()].sort((a, b) => b[1].違 - a[1].違);
let 全合 = 0, 全違 = 0;
for (const [名, s] of 並) {
  全合 += s.合; 全違 += s.違;
  console.log('  ' + (s.違 ? '✗' : '✓') + ' ' + 名.padEnd(24) + ' 合 ' + String(s.合).padStart(2) + ' ／ 違 ' + String(s.違).padStart(2));
  for (const e of s.例) console.log(e);
}
console.log('\n  ★合った ' + 全合 + ' ／ 違った ' + 全違 + '★');
