/* kansuu-519-atai-souname.mjs — ★関数 519個の うち 「値が 実Excel と 同じか」を 何個 測れて いるか★（2026-11-07・Exally-値 席）
 *
 *  ★★この 数で 何が 決まるか★★
 *    「関数の 名前 496/519」は ★名前が 在るか★ だけ。★値が 同じか★ の 分母が 無かった。
 *    ⇒ ① 519個の うち ★実Excel の 値を 控えた 行が 1本でも 在る★ 関数 ＝ 測れて いる 分母
 *       ② その 中で ★全部の 行が 合った★ 関数 ／ 1本でも 外れた 関数
 *       ③ ★控えが 1本も 無い★ 関数 ＝ 次に 実Excel で 取る 名簿（ここが 次の 仕事の 大きさ）
 *
 *  ★★材料★★（★読み取り ライブラリが 作った 値は 使わない＝全部 実Excel が 出した 値★）
 *    A. `osu-dai-dake.mjs --出し=` の 1行ずつ（紙 16枚・2116本・紙ごとの 盤面は あちらが 持つ）
 *    B. `golden-sounate-uchikata-2026-09-20.tsv`（319個・★盤面が 違う★＝下の 盤面で 押す）
 *       盤面の 出どころ ... `tsukuru-jitsu-excel-sounate-319.ps1:109-118`
 *       ★打った マスに 打つ★（ROW() 等が マスで 変わる 為）
 *
 *  ★★数え方★★
 *    ・関数の 名前は ★一番 外の 関数★ だけで 数える（中に 入った 関数は 数えない＝甘く しない）
 *    ・合う ... osu-dai-dake.mjs と 同じ 決まり（真偽の 大小だけ 揃える・数は 相対 1e-9）
 *    ・実Excel が #NAME? を 返した 行は ★比べない★（相手が 知らない＝値の 比べに ならない）
 *
 *  ★★見て いない 事★★
 *    ・★溢れ（スピル）の 2つ目 以降★（左上だけ）
 *    ・★形の 数★（1関数 1形 だけの 物が 多い＝「全部 合った」は 「控えた 形では 合った」）
 *    ・★空の マスの ずれ★は この 道具の 外（盤面に 空を 含む 形が 少ない）
 *    ・★構造化参照★（テーブル[列]）の 形は ★1本も 無い★＝古い A1 の 道だけを 測って いる
 *    ・★台 単体★の 数（画面の 道 convertFormula・_jsSet を 通して いない）
 *
 *  使い方: node docs/measured/kansuu-519-atai-souname.mjs [--名簿=出す先.tsv]
 */
import path from 'node:path';
import fs from 'node:fs';
import os from 'node:os';
import { execFileSync } from 'node:child_process';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';

const ここ = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(ここ, '..', '..');
const require_ = createRequire(path.join(ROOT, 'package.json'));

/* ── A. 2116本 ── */
const 一時 = path.join(os.tmpdir(), 'kansuu-519-dai-' + process.pid + '.tsv');
execFileSync(process.execPath, [path.join(ここ, 'osu-dai-dake.mjs'), '--出し=' + 一時],
  { stdio: ['ignore', 'ignore', 'inherit'] });
const 行A = fs.readFileSync(一時, 'utf-8').split(/\r?\n/).filter(Boolean).map((l) => l.split('\t'));
fs.unlinkSync(一時);

/* ── B. 319個 ── */
const Tsu = require_(path.join(ROOT, 'lib/shiki-tsunagi.js'));
require_(path.join(ROOT, 'lib/shiki-kansuu.js'));
const H = require_(path.join(ROOT, 'lib/shiki-hyou.js'));
try {
  const { JSDOM } = require_('jsdom');
  const 窓 = new JSDOM('<!doctype html>').window;
  Tsu.XML道具を渡す({ DOMParser: 窓.DOMParser, XPathResult: 窓.XPathResult });
} catch (e) { console.log('NG ★jsdom が 無い★'); process.exit(3); }

const 誤りの数 = {
  '-2146826281': '#DIV/0!', '-2146826252': '#NUM!', '-2146826246': '#N/A',
  '-2146826273': '#VALUE!', '-2146826265': '#REF!', '-2146826259': '#NAME?',
  '-2146826288': '#NULL!', '-2146826238': '#CALC!', '-2146826245': '#SPILL!',
};
const 盤面B = [['A1', 1], ['A2', 2], ['A3', 2], ['A4', 3], ['A5', 3], ['A6', 4],
  ['B1', 10], ['B2', 20], ['B3', 30], ['D1', 'abc'], ['E1', 1], ['F1', 2], ['E2', 3], ['F2', 4]];
const 生B = fs.readFileSync(path.join(ここ, 'golden-sounate-uchikata-2026-09-20.tsv'), 'utf-8').replace(/^﻿/, '');
const 行B = [];
let B対照 = false;
for (const l of 生B.split(/\r?\n/)) {
  if (!l || l.startsWith('#')) continue;
  const c = l.split('\t');
  const [マス, , 式, 値] = c;
  if (!式 || !式.startsWith('=') || 値 === undefined || 値 === '' || /打てません|HRESULT/.test(値)) continue;
  const 答 = 誤りの数[値] || 値;
  if (答 === '#NAME?') continue;
  const 板 = new H.表();
  for (const kv of 盤面B) 板.打つ(kv[0], kv[1]);
  let 字;
  try { 板.打つ(マス, 式); const o = 板.字(マス); 字 = (o === null || o === undefined) ? '' : String(o); }
  catch (e) { 字 = '★転んだ★'; }
  if (式 === '=ABS(A1:A3)' && 答 === '1') B対照 = true;
  行B.push([式, 字, 答]);
}
if (!B対照) { console.log('NG ★319 の 紙の 対照（=ABS(A1:A3) → 1）が 無い＝列が ずれた★'); process.exit(9); }

/* ── C／D. kansuu46 の 346・まだ（昔の 道 .Formula）と その 打ち直し（今の 道 Formula2） ──
   盤面（3枚 共通・紙の 頭）... A1:A5=1..5 ／ B1:B5=2,4,6,8,10 ／ D1=2024/1/1(45292) ／ D2=2026/1/1(46023)
   ★昔の 道の 紙は 範囲を 渡すと 暗黙の 交わり（1マス）★＝その 式は ★打ち直しの 紙（D）の 値を 使う★
   （mada の 紙の 注：範囲を 渡す 812本を 打ち直した＝D）。D に 無い 式だけ C から 採る。 */
const 盤面C = [['A1', 1], ['A2', 2], ['A3', 3], ['A4', 4], ['A5', 5],
  ['B1', 2], ['B2', 4], ['B3', 6], ['B4', 8], ['B5', 10], ['D1', 45292], ['D2', 46023]];
const 盤面C式 = 盤面C.map((kv) => (kv[0] === 'D1' ? ['D1', '=DATE(2024,1,1)'] : kv[0] === 'D2' ? ['D2', '=DATE(2026,1,1)'] : kv));
const 押す = (盤面, マス, 式) => {
  const 板 = new H.表();
  for (const kv of 盤面) 板.打つ(kv[0], kv[1]);
  try { 板.打つ(マス, 式); const o = 板.字(マス); return (o === null || o === undefined) ? '' : String(o); }
  catch (e) { return '★転んだ★'; }
};
const 読む = (名) => fs.readFileSync(path.join(ここ, 名), 'utf-8').replace(/^﻿/, '')
  .split(/\r?\n/).filter((l) => l && !l.startsWith('#') && l.includes('\t')).map((l) => l.split('\t'));
const 行D = [];
const D式 = new Set();
let Di = 0, D対照 = false;
for (const c of 読む('golden-furui-kami-torinaoshi-2026-09-21.tsv')) {
  const 式 = c[1], 値 = c[3];
  if (!式 || !式.startsWith('=') || 値 === undefined || 値 === '') continue;
  const マス = 'BZ1'; Di++;   /* ★既存の 試験 shiki-kansuu-kami.test.mjs:1509 と 同じ★（台は 暗黙の 交わりで 解く＝下の 行だと 外れる）*/
  const 答 = 誤りの数[値] || 値;
  D式.add(式);
  if (答 === '#NAME?') continue;
  if (式 === '=ABS(B1:B5)' && 答 === '2') D対照 = true;
  行D.push([式, 押す(盤面C, マス, 式), 答]);
}
if (!D対照) { console.log('NG ★打ち直しの 紙の 対照（=ABS(B1:B5) → 2）が 無い★'); process.exit(9); }
const 行C = [];
const C式 = new Set();
let C対照 = false;
for (const 名 of ['kansuu46/golden-346-2026-09-08.tsv', 'kansuu46/golden-mada-2026-09-14.tsv']) {
  for (const c of 読む(名)) {
    const 式 = c[1], 値 = c[2];
    if (!式 || !式.startsWith('=') || 値 === undefined || 値 === '' || /打てません|HRESULT/.test(値)) continue;
    if (D式.has(式) || C式.has(式)) continue;
    C式.add(式);
    const 答 = 誤りの数[値] || 値;
    if (答 === '#NAME?') continue;
    if (式 === '=ABS(A2)' && 答 === '2') C対照 = true;
    /* ★実Excel は H1 に .Formula で 打った（toru-346.ps1:185-209）・D1/D2 は 式 =DATE(…)★ */
    行C.push([式, 押す(盤面C式, 'H1', 式), 答]);
  }
}
if (!C対照) { console.log('NG ★346 の 紙の 対照（=ABS(A2) → 2）が 無い★'); process.exit(9); }

/* ── 比べる（osu-dai-dake.mjs と 同じ 決まり） ── */
const 真偽 = (x) => { const s = String(x).trim(); return /^(TRUE|FALSE)$/i.test(s) ? s.toUpperCase() : s; };
const 数 = /^[-+]?[0-9]*[.]?[0-9]+([eE][-+]?[0-9]+)?$/;
const 合うか = (字, 答) => 真偽(字) === 真偽(答)
  || (数.test(字) && 数.test(答) && Math.abs(Number(字) - Number(答)) <= Math.max(1e-9, Math.abs(Number(答)) * 1e-9));
const 外の名 = (式) => { const m = /^=\s*(?:_xlfn\.|_xlws\.)?([A-Z0-9_.]+)\(/i.exec(式); return m ? m[1].toUpperCase() : null; };

/* osu-dai-dake の 「測り道具の 限り」（jsdom の name()）は あちらで 合った 扱い＝ここでも 合った に 入れる */
/* ★毎回 変わる 関数は 比べない★（値が 同じに なる 筈が 無い） */
const 毎回 = new Set(['RAND', 'RANDBETWEEN', 'RANDARRAY', 'NOW', 'TODAY']);
/* ★精度の 外れ★＝数どうしで 相対 1e-6 以内（1e-9 は 超える）＝別に 数える（直すかは 別の 決め） */
const 精度か = (字, 答) => 数.test(字) && 数.test(答) && Math.abs(Number(字) - Number(答)) <= Math.max(1e-12, Math.abs(Number(答)) * 1e-6);
const 関数 = new Map();
const 足す = (名, よい, 行) => {
  if (!名) return;
  if (!関数.has(名)) 関数.set(名, { 合: 0, 外: 0, 精: 0, 例: null });
  const v = 関数.get(名);
  if (よい === true) v.合 += 1; else if (よい === '精度') v.精 += 1; else { v.外 += 1; if (!v.例) v.例 = 行; }
};
let 行数 = 0, 行合 = 0, 毎回の行 = 0;
for (const [式, 字, 答, 印] of 行A) {
  if (毎回.has(外の名(式))) { 毎回の行 += 1; continue; }
    const よい = (合うか(字, 答) || /測り道具の限り/.test(印 || '')) ? true : 精度か(字, 答) ? '精度' : false;
  行数 += 1; if (よい === true) 行合 += 1;
  足す(外の名(式), よい, 式 + ' → 台 [' + 字 + '] ／ 実Excel [' + 答 + ']');
}
for (const [式, 字, 答] of 行B) {
  if (毎回.has(外の名(式))) { 毎回の行 += 1; continue; }
    const よい = (合うか(字, 答)) ? true : 精度か(字, 答) ? '精度' : false;
  行数 += 1; if (よい === true) 行合 += 1;
  足す(外の名(式), よい, 式 + ' → 台 [' + 字 + '] ／ 実Excel [' + 答 + ']（319の盤面）');
}
for (const [札, 行たち] of [['346/まだ・昔の道', 行C], ['打ち直し・今の道', 行D]]) {
  for (const [式, 字, 答] of 行たち) {
    if (毎回.has(外の名(式))) { 毎回の行 += 1; continue; }
    const よい = (合うか(字, 答)) ? true : 精度か(字, 答) ? '精度' : false;
    行数 += 1; if (よい === true) 行合 += 1;
    足す(外の名(式), よい, 式 + ' → 台 [' + 字 + '] ／ 実Excel [' + 答 + ']（' + 札 + '）');
  }
}

const 名簿 = fs.readFileSync(path.join(ここ, 'excel-functions-2026-09-06.txt'), 'utf-8')
  .split(/\r?\n/).map((s) => s.trim().toUpperCase()).filter(Boolean);
const 測れた = [], 全部合う = [], 外れた = [], 控え無し = [], 精度だけ = [], 毎回の名 = [];
for (const n of 名簿) {
  const v = 関数.get(n);
  if (!v) { (毎回.has(n) ? 毎回の名 : 控え無し).push(n); continue; }
  if (毎回.has(n)) { 毎回の名.push(n); continue; }
  測れた.push(n);
  (v.外 ? 外れた : v.精 ? 精度だけ : 全部合う).push(n);
}
const 名簿の外 = [...関数.keys()].filter((n) => !名簿.includes(n));

console.log('[kansuu-519-atai-souname] ★関数 519個の 値の 総なめ（台 単体・実Excel の 控えと 比べる）★');
console.log('  行 ... A ' + 行A.length + '本 ＋ B（319の紙）' + 行B.length + '本 ＋ C（346/まだ・昔の道）' + 行C.length + '本 ＋ D（打ち直し）' + 行D.length + '本 ＝ ' + 行数 + '本 ／ 合った ' + 行合);
console.log('  ★名簿 ' + 名簿.length + '個★');
console.log('    ① 控えが 在る（値を 比べられた）... ' + 測れた.length);
console.log('    ② うち 全部の 行が 合った ........ ' + 全部合う.length);
console.log('       うち 精度の 差だけ（相対1e-6内）.. ' + 精度だけ.length + ' ... ' + 精度だけ.join(' '));
console.log('       うち 1本でも 外れた ............ ' + 外れた.length);
console.log('    （毎回 変わる＝比べない ' + 毎回の名.length + '個・' + 毎回の行 + '本 ... ' + 毎回の名.join(' ') + '）');
console.log('    ③ ★控えが 1本も 無い★ ........... ' + 控え無し.length + '（＝次に 実Excel で 取る 名簿）');
console.log('  （名簿の 外の 名前 ' + 名簿の外.length + '個 ... ' + 名簿の外.slice(0, 12).join(' ') + '）');
console.log('');
console.log('  ★★外れた 関数（1本目の 例）★★');
for (const n of 外れた) console.log('    ' + n.padEnd(16) + ' 合' + 関数.get(n).合 + ' 外' + 関数.get(n).外 + ' 精' + 関数.get(n).精 + '  ' + 関数.get(n).例.slice(0, 200));
console.log('');
console.log('  ★★控えが 無い ' + 控え無し.length + '個★★');
for (let i = 0; i < 控え無し.length; i += 10) console.log('    ' + 控え無し.slice(i, i + 10).join(' '));

const 出し = (process.argv.find((a) => a.startsWith('--名簿=')) || '').split('=').slice(1).join('=');
if (出し) {
  fs.writeFileSync(出し, 名簿.map((n) => {
    const v = 関数.get(n);
    return n + '\t' + (毎回.has(n) ? '毎回変わる' : v ? (v.外 ? '外れ' : v.精 ? '精度だけ' : '合う') : '控え無し') + '\t' + (v ? v.合 : 0) + '\t' + (v ? v.外 : 0);
  }).join('\n') + '\n', 'utf-8');
  console.log('\n  ★1関数 1行 ... ' + 出し + '★');
}
console.log('\n  名簿 ' + 名簿.length + '／控え在り ' + 測れた.length + '／全部合う ' + 全部合う.length);
