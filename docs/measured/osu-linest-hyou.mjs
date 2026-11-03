/* ★LINEST の 紙を ★うちの 本番の 道★で 押して 突き合わせる★（2026-09-09・2件目）
 *
 *  ★正★ … `golden-linest-hyou-2026-09-09.tsv`（実Excel の 実測）
 *  ★押す 道★ … 本番と 同じ ①JS層 `_jsComputeFormula` → ②`convertFormula` → エンジン
 *
 *  ★材料は 紙から 読む★（★手で 写さない★＝2026-09-08/09 に 3回 写し間違えた）
 *  ★数だけで 済ませない★＝★出た 字を 並べて 出す★
 *
 *  使い方: node docs/measured/osu-linest-hyou.mjs
 */

/* ★★★この 台の 数は 画面の 数では ありません★★★（2026-09-10 に 書いた）
   ★ここは `setSheetContent` で ★板ごと★ 入れて います★（本番は 1マスずつ）。
   ⇒ 2026-09-10 … 板ごと 入れた せいで 裸の `=LINEST(…)` が ★#VALUE!★ に なり、
     ★「本番が 壊れて いる」と 報告する 一歩 手前★まで 行きました。
     ⇒ 実配信を ★ブラウザで 押したら 4本とも 動いて いました★＝★台の 産物★
   ⇒★画面の 事を 言いたい なら ★ブラウザで 押して ください★★
     （`setCellFormula` は エンジンに 入れるだけ／`setCell(r,c,v)` が 画面側・★シート番号は 取らない★）
   ⇒ 見張り `tests/hakaridai-mon.test.mjs` が この 断りの 有無を 見ます */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createRequire } from 'node:module';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..', '..');
const 紙 = path.join(ROOT, 'docs/measured/golden-linest-hyou-2026-09-09.tsv');
const 出す先 = path.join(ROOT, 'docs/measured/golden-linest-hyou-awase-2026-09-09.tsv');

const require_ = createRequire(path.join(ROOT, 'package.json'));
/* ★★本番の 道は 1本★★（2026-09-15）＝`docs/measured/honban-no-michi.mjs`
   ★前は この 建て方が ★道具 6本に 写されて いました★（つなぐ行 のべ 35）
   ★写しが 1本でも 本番と ずれたら 嘘の 数字が 出ます★
     （2026-09-09〜10 に ★それで 1日に 5回 転んで います★）
   ★これから 借り物を 外します★＝★自前の 関数 86個の 付け先を 変える★
     ⇒★建て方が 6本 在ると 必ず 1本 忘れます★ */
const { 建てる } = await import(pathToFileURL(path.join(ROOT, 'docs/measured/honban-no-michi.mjs')).href);
const 道 = await 建てる();
const HFns = 道.HFns;
const EF = 道.EF;
const HF0 = 道.HF0;
const H = 道.H;
const hf = 道.hf;
const SID = 道.SID;
console.log('★本番と 同じ ' + 道.積んだ + '本の プラグインを つないだ★');

/* ══ ★紙を 読む（材料も 答えも）★ ══════════════════════════ */
const 材料 = {};
const 答え = [];
for (const l of fs.readFileSync(紙, 'utf-8').split('\n')) {
  if (l.startsWith('#材料')) {
    const c = l.split('\t');
    if (c.length >= 3) 材料[c[1].trim()] = Number(c[2]);
    continue;
  }
  if (l.startsWith('#') || !l.trim()) continue;
  const c = l.split('\t');
  if (c.length < 6) continue;
  答え.push({ 組: c[0].trim(), 式: c[1].trim(), 正: c[2].trim(), 型: c[3].trim(), 場所: c[5].trim() });
}
if (!Object.keys(材料).length || !答え.length) {
  console.error('★紙から 読めなかった（材料 ' + Object.keys(材料).length + ' ／ 行 ' + 答え.length + '）★');
  process.exit(2);
}

/* ★マスの 名前を 行・列に★ */
function マスを解く(s) {
  const m = /^([A-Z]+)(\d+)$/.exec(s);
  if (!m) return null;
  let c = 0;
  for (const ch of m[1]) c = c * 26 + (ch.charCodeAt(0) - 64);
  return { 行: Number(m[2]) - 1, 列: c - 1 };
}
let 最大行 = 0, 最大列 = 0;
const 置く = [];
for (const k of Object.keys(材料)) {
  const rc = マスを解く(k);
  if (!rc) continue;
  置く.push([rc, 材料[k]]);
  if (rc.行 > 最大行) 最大行 = rc.行;
  if (rc.列 > 最大列) 最大列 = rc.列;
}
const 板 = [];
for (let r = 0; r <= 最大行 + 3; r++) 板.push(new Array(最大列 + 4).fill(null));
for (const [rc, v] of 置く) 板[rc.行][rc.列] = v;

/* ══ ★本番と 同じ 道で 押す★ ══════════════════════════════ */
const 式の行 = 最大行 + 2;   /* ★材料の 下★に 置く（材料を 潰さない） */
function 押す(式) {
  /* ①JS層 … 本番は これを 先に 見る */
  hf.setSheetContent(SID, 板.map((r) => r.slice()));
  try {
    const r = EF._jsComputeFormula(0, 式);
    if (r !== null && r !== undefined) return { 道: 'JS層', 値: String(r) };
  } catch (e) { return { 道: 'JS層', 値: '★投げた（' + (e && e.message) + '）★' }; }
  /* ②エンジン */
  let 後;
  try { 後 = EF.convertFormula(式); } catch (e) { return { 道: 'エンジン', 値: '★変換で 投げた★' }; }
  try {
    const 盤 = 板.map((r) => r.slice());
    while (盤.length <= 式の行) 盤.push(new Array(最大列 + 4).fill(null));
    盤[式の行][0] = 後;
    hf.setSheetContent(SID, 盤);
    const v = hf.getCellValue({ sheet: SID, row: 式の行, col: 0 });
    if (v && v.type) return { 道: 'エンジン', 値: '#' + v.type };
    if (v === null || v === undefined) return { 道: 'エンジン', 値: '(空)' };
    return { 道: 'エンジン', 値: String(v) };
  } catch (e) { return { 道: 'エンジン', 値: '★投げた（' + (e && e.message) + '）★' }; }
}

/* ★合うか（★幅は 場所ごとに 変える＝1つの 数で 決めない★）★ */
const 幅 = 1e-9;
/* ★★誤りの 名前の 書き方を 揃える（2026-09-09 に 足した）★★
   実Excel は `#N/A` / `#DIV/0!` / `#NUM!` と 書き、エンジンは `#NA` / `#DIV_BY_ZERO` / `#NUM` と 返す。
   ⇒★同じ 誤りなのに「違う」と 数えて いました★（★字が 違うだけ★）
   ⇒★名前を 揃えてから 比べる★
   ★注意★ これは ★同じ 誤りを 同じと 見る★だけです。
          ★違う 誤り（#REF! と #NUM! など）は そのまま 違うと 数えます★ */
const 誤りの名 = (s) => {
  const t = String(s).trim().toUpperCase().replace(/[#!]/g, '').replace(/[\/_]/g, '');
  const 表 = { NA: 'NA', DIVBY0: 'DIV0', DIVBYZERO: 'DIV0', DIV0: 'DIV0', NUM: 'NUM',
    VALUE: 'VALUE', REF: 'REF', NAME: 'NAME', NULL: 'NULL', SPILL: 'SPILL',
    CALC: 'CALC', CYCLE: 'CYCLE', ERROR: 'ERROR' };
  return 表[t] || null;
};
function 合うか(う, 正) {
  if (う === 正) return true;
  const 誤り1 = 誤りの名(う), 誤り2 = 誤りの名(正);
  if (誤り1 && 誤り2) return 誤り1 === 誤り2;   /* ★どちらも 誤り＝名前で 比べる★ */
  if (誤り1 || 誤り2) return false;             /* ★片方だけ 誤り＝違う★ */
  const a = Number(う), b = Number(正);
  if (!isFinite(a) || !isFinite(b)) return false;
  if (b === 0) return a === 0;
  return Math.abs((a - b) / b) <= 幅;
}

const 行 = [];
const 言う = (s) => { 行.push(s); console.log(s); };
言う('# ★LINEST の 紙を うちの 本番の 道で 押して 突き合わせた★（2026-09-09・2件目）');
言う('#');
言う('# ★正★ golden-linest-hyou-2026-09-09.tsv（実Excel の 実測）');
言う('# ★押す 道★ 本番と 同じ ①JS層 → ②convertFormula → エンジン');
言う('# ★材料は 紙から 読んだ★（' + Object.keys(材料).length + 'マス）★手で 写して いません★');
言う('# ★合う 幅★ 相対 ' + 幅 + '（★字が そのまま 同じ 物は そのまま 合う★）');
言う('#');
言う('# 組\t式\t実Excel\tうち\t道\t判じ\t何の 場所');

const 数 = { 合った: 0, 違う: 0, NAME: 0, 誤りが違う: 0 };
const 例 = [];
for (const x of 答え) {
  const r = 押す(x.式);
  let 判;
  if (r.値 === '#NAME') { 判 = '★★#NAME?＝在るように 見えて 実は 無い★★'; 数.NAME++; }
  else if (合うか(r.値, x.正)) { 判 = '合った'; 数.合った++; }
  else if (/^#/.test(x.正) || /^#/.test(r.値)) { 判 = '★誤りの 出方が 違う★'; 数.誤りが違う++; }
  else { 判 = '★★違う★★'; 数.違う++; }
  if (判 !== '合った' && 例.length < 14) 例.push({ 式: x.式, 正: x.正, う: r.値, 場所: x.場所 });
  行.push(x.組 + '\t' + x.式 + '\t' + x.正 + '\t' + r.値 + '\t' + r.道 + '\t' + 判 + '\t' + x.場所);
}

言う('#');
言う('# ★★締め★★ 全 ' + 答え.length + '本');
言う('#   合った ……………………………………… ' + 数.合った);
言う('#   ★#NAME?（在るように 見えて 実は 無い）★ … ' + 数.NAME);
言う('#   ★値が 違う★ ………………………………… ' + 数.違う);
言う('#   ★誤りの 出方が 違う★ ……………………… ' + 数.誤りが違う);
言う('#   ―― 足すと ' + (数.合った + 数.NAME + 数.違う + 数.誤りが違う) + ' ／ 全 ' + 答え.length + '本');

fs.writeFileSync(出す先, 行.join('\n') + '\n');
console.log('\n★合わない 物の 実物（★数だけで 済ませない＝出た 字を 並べる★）★');
for (const e of 例) {
  console.log('  ' + e.式.padEnd(46) + ' 実Excel=' + String(e.正).padEnd(22) + ' うち=' + String(e.う).padEnd(22) + ' ' + e.場所);
}
console.log('\n★書いた … ' + 出す先 + '★');
if (数.合った + 数.NAME + 数.違う + 数.誤りが違う !== 答え.length) {
  console.error('★数が 合わない＝数え落として いる★'); process.exit(3);
}
