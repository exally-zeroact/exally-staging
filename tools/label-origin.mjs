/* label-origin.mjs — ★札は「切った物」か「言い換えた物」か を 数える★ 2026-08-30
 *
 *  ★なぜ★（監査役の 宿題・2026-08-30）
 *    「短く 切る」は 監査役の 決めなので 問題ない。
 *    ★でも「中身一覧」は 切ったのでは なく ★言い換え★★＝
 *    ★お客さんが「Excel で 見た あの機能」を 探せなくなる恐れが ある★。
 *
 *  ★分け方（★字が 残っているかで 決める★・見た目で 決めつけない）★
 *    そのまま … 実Excelの 名前と 札が 同じ
 *    切った   … ★連続 2字以上★が そのまま 残っている（例 ウィンドウ枠の固定 → 枠の固定）
 *    1字だけ  … ★連続で 残っているのが 1字だけ★（きわどい＝別に 出す）
 *    言い換え … ★1字も 続けて 残っていない★（例 ナビゲーション → 中身一覧）
 *
 *  ★注意★
 *    ・末尾の「...」「…」は 落としてから 比べる（うちは 窓を 開くので 元から 落としている）
 *    ・★間の 空白は 詰めて 比べる★（「フォント サイズ」と「フォントサイズ」を 別扱いに しない）
 *
 *  走らせ方: node tools/label-origin.mjs
 */
import fs from 'node:fs';
import path from 'node:path';

const ROOT = path.join(path.dirname(new URL(import.meta.url).pathname.replace(/^\/([A-Za-z]:)/, '$1')), '..');
const ラベル = (await import('file://' + path.join(ROOT, 'lib/ribbon-label.js').replace(/\\/g, '/'))).default;
const spec = fs.readFileSync(path.join(ROOT, 'lib/ribbon-spec.js'), 'utf8');

const 部品 = [];
for (const 行 of spec.split(String.fromCharCode(10))) {
  const m = /\{ t: '([^']*)', g: '([^']*)', p: '([^']*)', a: (.*)\},?\s*$/.exec(行.trim());
  if (m && m[4].trim() !== 'null') 部品.push({ t: m[1], g: m[2], p: m[3] });
}
/* ★空振りを その場で 止める★ */
if (部品.length < 200) {
  console.log('★読めていません★ 結んだ部品 = ' + 部品.length + '個');
  process.exit(1);
}

const 詰める = (s) => String(s).replace(/[\s　]/g, '');

/** ★2つの 字の 中で いちばん 長い「続けて 同じ 所」の 長さ★ */
function 一番長く続く(a, b) {
  const x = 詰める(a), y = 詰める(b);
  let 最大 = 0;
  const 表 = new Array(y.length + 1).fill(0);
  for (let i = 1; i <= x.length; i++) {
    let 前 = 0;
    for (let j = 1; j <= y.length; j++) {
      const いま = 表[j];
      表[j] = (x[i - 1] === y[j - 1]) ? 前 + 1 : 0;
      if (表[j] > 最大) 最大 = 表[j];
      前 = いま;
    }
  }
  return 最大;
}

const 群 = { そのまま: [], 切った: [], 一字だけ: [], 言い換え: [] };
for (const v of 部品) {
  const 元 = ラベル._尻を落とす(v.p);
  const 札 = ラベル.札(v.p, v.t, v.g, 1);
  const 場所 = v.t + '|' + v.g + '|';
  if (詰める(元) === 詰める(札)) { 群.そのまま.push(場所 + 元); continue; }
  const n = 一番長く続く(元, 札);
  const 行 = 場所 + '「' + 元 + '」→「' + 札 + '」（続けて ' + n + '字）';
  if (n >= 2) 群.切った.push(行);
  else if (n === 1) 群.一字だけ.push(行);
  else 群.言い換え.push(行);
}

const 全 = 部品.length;
console.log('★結んだ 部品 = ' + 全 + '個★\n');
console.log('  そのまま  … ' + String(群.そのまま.length).padStart(4) + '個（実Excelの 名前の まま）');
console.log('  ★切った★ … ' + String(群.切った.length).padStart(4) + '個（連続 2字以上 残っている）');
console.log('  ★1字だけ★… ' + String(群.一字だけ.length).padStart(4) + '個（きわどい）');
console.log('  ★言い換え★… ' + String(群.言い換え.length).padStart(4) + '個（1字も 続けて 残っていない）');

console.log('\n★★言い換えた物 = ' + 群.言い換え.length + '個★★');
群.言い換え.forEach((v) => console.log('  ★' + v));

console.log('\n★1字だけ 残っている物 = ' + 群.一字だけ.length + '個★（切ったか 言い換えか きわどい）');
群.一字だけ.forEach((v) => console.log('  ・' + v));

console.log('\n★切った物 = ' + 群.切った.length + '個★（先頭 12個だけ 出す）');
群.切った.slice(0, 12).forEach((v) => console.log('  ・' + v));

/* ★検算★＝4つの 合計が 全部の数と 合うか（黙って 落としていないか） */
const 合計 = 群.そのまま.length + 群.切った.length + 群.一字だけ.length + 群.言い換え.length;
console.log('\n★検算★ ' + 群.そのまま.length + ' + ' + 群.切った.length + ' + ' + 群.一字だけ.length
  + ' + ' + 群.言い換え.length + ' = ' + 合計 + ' / ' + 全 + ' … ' + (合計 === 全 ? '★合う★' : '★合わない★'));
process.exit(合計 === 全 ? 0 : 1);
