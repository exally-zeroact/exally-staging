/* typed-value.test.mjs — ★E3「1,234 が文字列＝合計に入らない」の見張り★
 *
 * 何を守るか（★期限 2026-09-30・金が落ちる所★）
 *   `1,234` と打つと文字列のままで、=SUM() が拾わなかった。
 *   ★#ERROR は出ない。合計が黙って小さくなる★＝いちばん気づけない壊れ方。
 *
 * 突き合わせる相手は ★実Excel(COM)で打ち込んで出た答え★
 *   tests/fixtures/typed-value-golden.json（30通り・うちで作った期待値ではない）
 *
 * 見る所は3つ
 *   ① 打った文字の解釈が Excel と同じか（lib/typed-value.js）
 *   ② ★計算に渡す値（book.html の toHFVal）と 書き出す値（lib/grid-xlsx.js の asValue）が同じか★
 *      ズレると「画面の合計」と「落としたファイルの合計」が食い違う＝いちばんやってはいけない壊れ方
 *   ③ ★=SUM() が実際に拾うか★（本番と同じ経路で計算させる）
 *
 * 使い方: node tests/typed-value.test.mjs [--self-test]
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createRequire } from 'node:module';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(__dirname, '..');
const require = createRequire(pathToFileURL(path.join(ROOT, 'package.json')));

const TV = require(path.join(ROOT, 'lib', 'typed-value.js'));
const GX = require(path.join(ROOT, 'lib', 'grid-xlsx.js'));
const GOLD = JSON.parse(fs.readFileSync(path.join(__dirname, 'fixtures', 'typed-value-golden.json'), 'utf8'));

let ng = 0, n = 0;
function ok(cond, label, extra) {
  n++;
  if (cond) { console.log('  ✓ ' + label); return true; }
  ng++;
  console.log('  ✗ ' + label + (extra ? '\n      ' + extra : ''));
  return false;
}

/* ★今回 直さないと決めた物（Excel の真値は fixture に残してある）★
   ここに載せた物は「Excel と違ってよい」ではなく ★「まだ測って直していない」と明示する場所★。
   足す時は ★実Excelで測ってから★。 */
const NOT_YET = {
  '(1,234)': '会計のマイナス。カンマ無しの (1234) を測っていないので触らない',
  '1,234%': '値が100分の1になる＝桁区切りとは別の話',
  '$1,234': '通貨記号。★¥1,234 は Excel も文字のまま（実測）★＝記号ごとに答えが違う',
  '1,234E2': '指数。カンマと混ぜない',
};

/* ── ① 打った文字の解釈が Excel と同じか ────────────────────────── */
function testAgainstExcel() {
  console.log('\n[① 打った文字の解釈が 実Excel と同じか]');
  let same = 0, notYet = 0, bad = [];
  for (const p of GOLD.probes) {
    /* うちの答え＝書き出す側の関数（asValue）を通した物＝★本番が使う経路★ */
    const got = GX.asValue(p.raw);
    const gotType = typeof got === 'number' ? 'n' : (typeof got === 'boolean' ? 'b' : 's');
    const agree = (gotType === p.type) && (p.type !== 'n' || Number(got) === Number(p.value));
    if (agree) { same++; continue; }
    if (NOT_YET[p.raw]) { notYet++; continue; }
    bad.push(`「${p.raw}」 Excel=${p.type} ${p.value} ／ うち=${gotType} ${got}`);
  }
  console.log('      ── 実測 ── ' + GOLD.probes.length + '通りを数えて 一致 ' + same
    + ' / ★まだ直していない（一覧に理由あり）★ ' + notYet + ' / 食い違い ' + bad.length);
  ok(bad.length === 0, '★実Excel と食い違う物が無い（まだ直していない物を除く）', bad.join('\n      '));
  ok(same >= 20, '★実際に比べた数が20通り以上（検査が空振りしていない）＝' + same);
  ok(notYet === Object.keys(NOT_YET).length,
    '★「まだ直していない」の数が一覧と合う（黙って増えていない）＝' + notYet + '/' + Object.keys(NOT_YET).length);
}

/* ── ② 計算に渡す値と 書き出す値が同じか ───────────────────────── */
function testBothSidesAgree() {
  console.log('\n[② 計算に渡す値（toHFVal）と 書き出す値（asValue）が同じか]');
  /* book.html の toHFVal を そのまま切り出して動かす（実物の行を読む＝写しを作らない） */
  const html = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf8');
  const m = /function toHFVal\(v\)\{[\s\S]*?\n\}/.exec(html);
  if (!ok(!!m, 'book.html から toHFVal を取り出せた')) return;
  const fn = new Function('TypedValue', 'parseDateStr', 'dateSerial',
    m[0] + '; return toHFVal;')(TV, GX.dateFmtFor ? parseDateStrOf(GX) : null, GX.dateSerial);
  let bad = [];
  for (const p of GOLD.probes) {
    const a = fn(p.raw);
    const b = GX.asValue(p.raw);
    const aNum = typeof a === 'number', bNum = typeof b === 'number';
    if (aNum !== bNum || (aNum && a !== b)) bad.push(`「${p.raw}」 計算=${typeof a} ${a} ／ 書き出し=${typeof b} ${b}`);
  }
  ok(bad.length === 0, '★30通りとも 計算に渡す値と書き出す値が同じ（画面と落としたファイルの合計が食い違わない）',
    bad.join('\n      '));
  console.log('      ── 実測 ── ' + GOLD.probes.length + '通りを突き合わせて 食い違い ' + bad.length);
}
/* grid-xlsx は parseDateStr を外に出していないので、同じ規則の物を作って渡す（日付は本題ではない） */
function parseDateStrOf() {
  return function (s) {
    if (typeof s !== 'string') return null;
    const mm = s.match(/^(\d{4})([/-])(\d{1,2})\2(\d{1,2})$/);
    if (!mm) return null;
    const y = +mm[1], mo = +mm[3], d = +mm[4];
    if (mo < 1 || mo > 12 || d < 1 || d > 31) return null;
    const dt = new Date(y, mo - 1, d);
    if (dt.getFullYear() !== y || dt.getMonth() !== mo - 1 || dt.getDate() !== d) return null;
    return { y, m: mo, d };
  };
}

/* ── ③ 本番と同じ経路で =SUM() が拾うか ─────────────────────────── */
async function testSumPicksItUp() {
  console.log('\n[③ 本番と同じ経路（book.html）で =SUM() が拾うか]');
  let JSDOM;
  try { ({ JSDOM } = require('jsdom')); }
  catch { ng++; console.log('  ✗ jsdom が無い＝本番経路を通せない。★緑ではない★'); return; }

  const html = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf8');
  const dom = new JSDOM(html.replace(/<script[\s\S]*?<\/script>/g, ''), {
    runScripts: 'dangerously', pretendToBeVisual: true, url: 'http://localhost/',
  });
  const win = dom.window, doc = win.document;
  win.fetch = () => Promise.reject(new Error('no net'));
  win.scrollTo = () => { };
  win.alert = () => { };
  const stubCtx = new Proxy({}, {
    get: (t, k) => (k === 'measureText' ? (() => ({ width: 10 }))
      : k === 'canvas' ? { width: 800, height: 600 }
        : k === 'getImageData' ? (() => ({ data: [] })) : (() => { })),
  });
  win.HTMLCanvasElement.prototype.getContext = () => stubCtx;
  const inject = (code) => { const el = doc.createElement('script'); el.textContent = code; doc.body.appendChild(el); };
  for (const mm of html.matchAll(/<script src="([^"]+)"><\/script>/g)) {
    const src = mm[1].split('?')[0];
    if (/^https?:/.test(src)) continue;
    inject(fs.readFileSync(path.join(ROOT, src), 'utf8'));
  }
  for (const mm of html.matchAll(/<script(?![^>]*\ssrc=)[^>]*>([\s\S]*?)<\/script>/g)) inject(mm[1]);
  await new Promise(res => {
    if (doc.readyState === 'complete') return res();
    win.addEventListener('load', res);
    setTimeout(res, 3000);
  });
  if (!ok(typeof win.TypedValue !== 'undefined', '★book.html が lib/typed-value.js を読んでいる')) return;
  if (!ok(typeof win.setCellFormula === 'function', 'setCellFormula が使える')) return;
  win.initFormulaEngine(['Sheet1']);

  /* A1=1,234 / A2=２，３４５（全角） / A3=1,23（Excelでも文字） を打って合計する */
  win.setCellFormula('Sheet1', 0, 0, '1,234');
  win.setCellFormula('Sheet1', 1, 0, '２，３４５');
  win.setCellFormula('Sheet1', 2, 0, '1,23');
  const sum = win.setCellFormula('Sheet1', 4, 0, '=SUM(A1:A3)');
  ok(Number(sum) === 3579, '★=SUM が 1,234 と ２，３４５ を拾う（1234+2345=3579）／実際は ' + sum);
  const cnt = win.setCellFormula('Sheet1', 5, 0, '=COUNT(A1:A3)');
  ok(Number(cnt) === 2, '★数として数えるのは2つ（1,23 は Excel と同じく文字のまま）／実際は ' + cnt);
  try { win.close(); } catch (e) { /* 閉じられなくても検査は済んでいる */ }
}

/* ── ④ 配線 ────────────────────────────────────────────────── */
function testWiring() {
  console.log('\n[④ 配線]');
  const html = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf8');
  const iTv = html.indexOf('lib/typed-value.js');
  const iGx = html.indexOf('lib/grid-xlsx.js');
  ok(iTv > 0, 'book.html が lib/typed-value.js を読んでいる');
  ok(iTv > 0 && iGx > 0 && iTv < iGx, '★typed-value.js は grid-xlsx.js より先に読む（後だと掴めない）');
  const gx = fs.readFileSync(path.join(ROOT, 'lib', 'grid-xlsx.js'), 'utf8');
  ok(/TypedValue\.parseTyped\(/.test(gx), '★書き出す側が TypedValue を呼んでいる（規則を写していない）');
  ok(/TypedValue\.parseTyped\(/.test(html), '★計算する側も TypedValue を呼んでいる');
  ok(!/isNaN\(v\.replace\(|replace\(\/,\/g/.test(html), '★book.html の中に カンマを外す規則を写していない（規則は1本だけ）');
}

/* ── わざと壊して赤になるか ─────────────────────────────────── */
function selfTest() {
  console.log('\n★--self-test★ わざと壊して、赤になる通り数を数える');
  const ways = [];
  const P = TV.parseTyped;
  ways.push(['① 3桁未満の組を数にしない（1,23）', P('1,23') === null]);
  ways.push(['② 3桁未満の組を数にしない（12,34）', P('12,34') === null]);
  ways.push(['③ 空の組を数にしない（1,,234）', P('1,,234') === null]);
  ways.push(['④ 先頭にカンマは数にしない（,234）', P(',234') === null]);
  ways.push(['⑤ 末尾がカンマは数にしない（1,）', P('1,') === null]);
  ways.push(['⑥ 後ろのマイナスは数にしない（1,234-）', P('1,234-') === null]);
  ways.push(['⑦ 空白入りは数にしない（1 234）', P('1 234') === null]);
  ways.push(['⑧ 単位付きは数にしない（1,234円）', P('1,234円') === null]);
  ways.push(['⑨ 括弧は今回 対象外（(1,234) は文字のまま）', P('(1,234)') === null]);
  ways.push(['⑩ %は今回 対象外（1,234% は文字のまま）', P('1,234%') === null]);
  ways.push(['⑪ 桁区切りは数になる（1,234→1234）', (P('1,234') || {}).num === 1234]);
  ways.push(['⑫ 4桁の組も数になる（1,2345→12345・Excel実測）', (P('1,2345') || {}).num === 12345]);
  ways.push(['⑬ 全角も数になる（１，２３４→1234）', (P('１，２３４') || {}).num === 1234]);
  ways.push(['⑭ 全角の数字だけも数になる（１２３４→1234）', (P('１２３４') || {}).num === 1234]);
  ways.push(['⑮ 前後の空白は落とす（" 1,234 "→1234）', (P(' 1,234 ') || {}).num === 1234]);
  ways.push(['⑯ マイナスが効く（-1,234→-1234）', (P('-1,234') || {}).num === -1234]);
  ways.push(['⑰ 小数が効く（1,234.5→1234.5）', (P('1,234.5') || {}).num === 1234.5]);
  ways.push(['⑱ 書式は #,##0（小数なし）', (P('1,234') || {}).numFmt === '#,##0']);
  ways.push(['⑲ 書式は #,##0.00（小数あり）', (P('1,234.5') || {}).numFmt === '#,##0.00']);
  ways.push(['⑳ 文字でない物は触らない（数値そのもの）', P(1234) === null]);

  let red = 0;
  for (const [label, caught] of ways) {
    if (caught) { red++; console.log('  ✓ ' + label); }
    else { ng++; console.log('  ✗ ' + label + ' → ★素通りした（見張りの穴）'); }
    n++;
  }
  console.log('\n  ── 実測 ── 確かめた ' + ways.length + ' 通り / 通った ' + red + ' 通り');
}

/* ── 実行 ─────────────────────────────────────────────────── */
console.log('[typed-value] 打ち込んだ文字の解釈（E3: 1,234 が合計に入らない）');
if (process.argv.includes('--self-test')) {
  selfTest();
} else {
  testAgainstExcel();
  testBothSidesAgree();
  await testSumPicksItUp();
  testWiring();
}
console.log('\n' + (n - ng) + ' passed, ' + ng + ' failed');
process.exit(ng ? 1 : 0);
