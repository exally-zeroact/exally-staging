/* table-refs.test.mjs — ★表の名前での参照（Table[列名]）が A1 範囲に直るか★
 *
 * 何を守る検査か（2026-08-18 に実物で分かった事）:
 *   司さんの実物（代行計算表2026.xlsb・式15,126本）のうち ★11,669本★ が Table[列名] の形で、
 *   ★1本残らず #ERROR★ になっていた。原因は計算エンジンではなく ★読み込みライブラリ★。
 *     実Excelの真値 =INDEX(R8.1[白石正人], MATCH(B4, R8.1[日付], 0))
 *     SheetJS の式  =INDEX(Table1[#Data],  MATCH(B4, Table1[#Data], 0))   ←列名が消えている
 *   .xlsx では消えない。★消えるのは .xlsb だけ★（下の ⑦ で両方 測る）。
 *
 * 見本 tests/fixtures/table-refs-sample.{xlsb,xlsx} は ★実Excel(COM)で作った物★。
 *   真値 tests/fixtures/table-refs-golden.json も実Excelの .Formula / .Text / .Value2。
 *   ★うちで作った偽物では、偽物自体が嘘をつく★ので、必ず実Excelに作らせている。
 *
 * 使い方:
 *   node tests/table-refs.test.mjs
 *   node tests/table-refs.test.mjs --self-test   ★わざと壊して、何通りで赤になるかを数える★
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createRequire } from 'node:module';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(__dirname, '..');
const require = createRequire(pathToFileURL(path.join(ROOT, 'package.json')));
const FIX = path.join(__dirname, 'fixtures');

const TR = require(path.join(ROOT, 'lib', 'table-refs.js'));
const Zip = require(path.join(ROOT, 'lib', 'zip-surgeon.js'));
const XLSX = require(path.join(ROOT, 'lib', 'xlsx.full.min.js'));

let ng = 0, n = 0;
function ok(cond, label, extra) {
  n++;
  if (cond) { console.log('  ✓ ' + label); return true; }
  ng++;
  console.log('  ✗ ' + label + (extra ? '\n      ' + extra : ''));
  return false;
}
function eq(got, want, label) {
  return ok(String(got) === String(want), label, 'got[' + got + '] want[' + want + ']');
}

const gold = JSON.parse(fs.readFileSync(path.join(FIX, 'table-refs-golden.json'), 'utf8'));

/* ★この見本で「こう直る」と決めた形。実Excelの範囲（golden.table）から人が確かめた物。
   data=A3:F7 / header=A2:F2 / totals=A8:F8 / 列= 日付,名前,金額,時間,"単価 ",倍 */
const WANT = {
  '元表|H1': '=SUM(C3:C7)',
  '元表|H2': '=INDEX(C3:C7, MATCH("う", B3:B7, 0))',
  '元表|H3': '=SUM(C3:D7)',
  '元表|H4': '=COUNTA(A2:F2)',
  '元表|H5': '=SUM(A8:F8)',
  '元表|H6': '=COUNTA(A2:F8)',
  '元表|H7': '=SUM(C3:C7)',
  '元表|H8': '=SUM(E3:E7)',
  '元表|H9': '=COUNTA(A2:F7)',
  '元表|H10': '=SUM(C3:C7)-SUM(D3:D7)',
  '別シート|A1': '=SUM(元表!C3:C7)',
  '別シート|A2': '=INDEX(元表!E3:E7, MATCH("え", 元表!B3:B7, 0))',
  '別シート|A3': '=COUNTA(元表!A2:F8)',
  '元表|F3': '=C3*2',
  '元表|F4': '=C4*2',
  '元表|F5': '=C5*2',
  '元表|F6': '=C6*2',
  '元表|F7': '=C7*2',
};

function decodeAddr(a) {
  const m = /^([A-Z]+)(\d+)$/.exec(a);
  let c = 0;
  for (let i = 0; i < m[1].length; i++) c = c * 26 + (m[1].charCodeAt(i) - 64);
  return { r: parseInt(m[2], 10) - 1, c: c - 1 };
}
function readFixture(name) {
  const bytes = new Uint8Array(fs.readFileSync(path.join(FIX, name)));
  const wb = XLSX.read(bytes, { type: 'array', cellFormula: true, cellNF: true });
  return { bytes, wb };
}

/* ── ① 表の定義（.xlsb）が実Excelと同じか ───────────────────────── */
async function testTableDefsXlsb() {
  console.log('\n[① 表の定義(.xlsb)が実Excelと同じか]');
  const { bytes } = readFixture('table-refs-sample.xlsb');
  const zip = Zip.read(bytes);
  const defs = await TR._loadTables(zip, true);
  ok(defs.length === 1, '表を1つ読めた（' + defs.length + '）');
  if (!defs.length) return;
  const d = defs[0];
  const colN = TR._colName;
  eq(colN(d.cl1) + (d.rw1 + 1) + ':' + colN(d.cl2) + (d.rw2 + 1), gold.table.range, '範囲が実Excelと同じ');
  eq(d.sheet, gold.table.sheet, 'シート名が実Excelと同じ');
  eq(d.header, 1, '見出し行は1行');
  eq(d.totals, 1, '合計行は1行');
  eq(d.cols.join('|'), gold.table.cols.join('|'), '★列名が実Excelと同じ（空白入りの「単価 」を含む）');
  /* ★中身の行が 実Excel の DataBodyRange と同じか（見出し/合計の数え方が逆でも気付ける） */
  const dataFirst = d.rw1 + d.header, dataLast = d.rw2 - d.totals;
  eq(colN(d.cl1) + (dataFirst + 1) + ':' + colN(d.cl2) + (dataLast + 1), gold.table.data, '★中身の行が実Excelの DataBodyRange と同じ');
}

/* ── ② 表の定義（.xlsx）も同じか ─────────────────────────────── */
async function testTableDefsXlsx() {
  console.log('\n[② 表の定義(.xlsx)も同じか]');
  const { bytes } = readFixture('table-refs-sample.xlsx');
  const zip = Zip.read(bytes);
  const defs = await TR._loadTables(zip, false);
  ok(defs.length === 1, '表を1つ読めた（' + defs.length + '）');
  if (!defs.length) return;
  const d = defs[0];
  eq(d.name, gold.table.name, '表の名前');
  eq(d.sheet, gold.table.sheet, 'シート名');
  eq(d.cols.join('|'), gold.table.cols.join('|'), '列名');
  eq(d.header + '/' + d.totals, '1/1', '見出し行/合計行');
}

/* ── ③④ 直した式が「こう直る」と決めた形と同じか（.xlsb / .xlsx）───── */
async function testRewrite(file, kind, tag) {
  console.log('\n[' + tag + ' 直した式（' + kind + '）]');
  const { bytes, wb } = readFixture(file);
  const res = await TR.resolve(bytes, kind, wb, Zip);
  ok(res.ok, '読めた（why=' + (res.why || '-') + '）');
  eq(res.stats.refused, 0, '★断ったセルが0本');
  let bad = 0;
  for (const key of Object.keys(WANT)) {
    const [sheet, addr] = key.split('|');
    const rc = decodeAddr(addr);
    const got = res.fixes[sheet + '|' + rc.r + ',' + rc.c];
    if (String(got) !== WANT[key]) { bad++; console.log('      ' + key + ' got[' + got + '] want[' + WANT[key] + ']'); }
  }
  ok(bad === 0, '★' + Object.keys(WANT).length + '本すべてが決めた形になった（違い ' + bad + '）');
  /* ★数えた物の数を出す（0件でしたが「見ていないだけ」にならないように）★ */
  console.log('      ── 実測 ── 表 ' + res.stats.tables + ' / 表の参照を含むセル ' + res.stats.cells
    + ' / 直した ' + res.stats.fixed + ' / 断った ' + res.stats.refused);
  ok(res.stats.cells >= 18, '見ているセルが18本以上ある（検査が空振りしていない）');
}

/* ── ⑤ 範囲の作り方（境界）───────────────────────────────────── */
function testToRange() {
  console.log('\n[⑤ 範囲の作り方（境界を実物の形で）]');
  const def = { rw1: 1, rw2: 7, cl1: 0, cl2: 5, header: 1, totals: 1, cols: ['日付', '名前', '金額', '時間', '単価 ', '倍'], sheet: '元表' };
  eq(TR._toRange(def, { band: 'data', colFirst: 2, colLast: 2, fromSheet: '元表' }), 'C3:C7', '中身・1列');
  eq(TR._toRange(def, { band: 'all', fromSheet: '元表' }), 'A2:F8', 'ぜんぶ');
  eq(TR._toRange(def, { band: 'headers', fromSheet: '元表' }), 'A2:F2', '見出し行');
  eq(TR._toRange(def, { band: 'totals', fromSheet: '元表' }), 'A8:F8', '合計行');
  eq(TR._toRange(def, { band: 'thisRow', curRow: 4, colFirst: 2, colLast: 2, fromSheet: '元表' }), 'C5', '自分の行（1セル）');
  eq(TR._toRange(def, { band: 'data', colFirst: 2, colLast: 2, fromSheet: '別シート' }), '元表!C3:C7', '★別シートからはシート名が付く');
  ok(TR._toRange(def, { band: 'thisRow', curRow: 1, colFirst: 2 }) === null, '★自分の行が見出し行なら 直さない（null）');
  ok(TR._toRange(def, { band: 'thisRow', curRow: 7, colFirst: 2 }) === null, '★自分の行が合計行なら 直さない（null）');
  ok(TR._toRange(def, { band: 'data', colFirst: 6 }) === null, '★列が表からはみ出したら 直さない（列は0〜5）');
  ok(TR._toRange(def, { band: 'data', colFirst: -1 }) === null, '★列が負なら 直さない');
  ok(TR._toRange(def, { band: 'data', colFirst: 2, colLast: 1 }) === null, '★列の順が逆なら 直さない');
  const noHead = { rw1: 1, rw2: 7, cl1: 0, cl2: 5, header: 0, totals: 0, cols: [], sheet: 'S' };
  ok(TR._toRange(noHead, { band: 'headers' }) === null, '★見出し行が無い表で [#Headers] は 直さない');
  ok(TR._toRange(noHead, { band: 'totals' }) === null, '★合計行が無い表で [#Totals] は 直さない');
  eq(TR._toRange(noHead, { band: 'data' }), 'A2:F8', '見出しも合計も無い表の中身は表ぜんぶ');
  ok(TR._toRange(def, { band: 'しらない band' }) === null, '★知らない行の種類は 直さない');
}

/* ── ⑥ 辻褄が合わなければ直さない ───────────────────────────── */
function testRefuse() {
  console.log('\n[⑥ 辻褄が合わなければ直さない（壊すより断る）]');
  const byId = { 1: { rw1: 1, rw2: 7, cl1: 0, cl2: 5, header: 1, totals: 1, cols: [], sheet: '元表', id: 1 } };
  const ctx = { row: 3, sheet: '元表' };
  const L = (o) => Object.assign({ coltype: 1, code: 0, id: 1, colFirst: 2, colLast: 2 }, o);
  eq(TR._rewriteXlsbFormula('=SUM(Table1[#Data])', [L({})], byId, ctx), '=SUM(C3:C7)', 'まっとうな1本は直る');
  ok(TR._rewriteXlsbFormula('=SUM(Table1[#Data])', [L({}), L({})], byId, ctx) === null, '★数が合わない（式1・拾2）→ 直さない');
  ok(TR._rewriteXlsbFormula('=SUM(Table1[#Data])+SUM(Table1[#Data])', [L({})], byId, ctx) === null, '★数が合わない（式2・拾1）→ 直さない');
  ok(TR._rewriteXlsbFormula('=SUM(Table9[#Data])', [L({})], byId, ctx) === null, '★表idが合わない → 直さない');
  ok(TR._rewriteXlsbFormula('=SUM(Table1[#Totals])', [L({})], byId, ctx) === null, '★行の種類が合わない → 直さない');
  ok(TR._rewriteXlsbFormula('=SUM(Table2[#Data])', [L({ id: 2 })], byId, ctx) === null, '★知らない表 → 直さない');
  ok(TR._rewriteXlsbFormula('=SUM(Table1[#Data])', [L({ colFirst: 99 })], byId, ctx) === null, '★列が表の外 → 直さない');
  ok(TR._rewriteXlsbFormula('=1+1', [], byId, ctx) === null, '表の参照が無い式は触らない');
  /* ★消された表（id=0xFFFFFFFF）は Excel 自身が #REF! と書いている（実物の .Formula で確認）★
     ここで断ると IFERROR まで道連れで #ERROR になり、Excel は空を出すのに うちだけ #ERROR になる。 */
  eq(TR._rewriteXlsbFormula('=IFERROR(Table1[#Data]-Table4294967295[#?Current],"")',
    [L({}), L({ id: 0xFFFFFFFF, code: 16 })], byId, ctx),
    '=IFERROR(C3:C7-#REF!,"")', '★消された表は #REF! に置く（Excelと同じ・IFERROR が拾える）');
}

/* ── ⑦ .xlsb だけが壊れている事を、両方 読んで確かめる ───────────── */
function testOnlyXlsbIsBroken() {
  console.log('\n[⑦ 読み込みライブラリは .xlsb でだけ列名を捨てる]');
  const b = readFixture('table-refs-sample.xlsb').wb;
  const x = readFixture('table-refs-sample.xlsx').wb;
  const fb = String(b.Sheets['元表']['H2'].f), fx = String(x.Sheets['元表']['H2'].f);
  ok(/Table\d+\[#Data\]/.test(fb), '.xlsb は Table<数>[#Data] に化けている（' + fb.slice(0, 46) + '…）');
  ok(fx.indexOf('売上T[金額]') >= 0 && fx.indexOf('売上T[名前]') >= 0, '.xlsx は列名が残っている（' + fx.slice(0, 46) + '…）');
  ok(fb.indexOf('金額') < 0, '★.xlsb からは列名が完全に消えている＝エンジン側では絶対に直せない');
}

/** 受け取り口（js/book-open.js）を読んでいる画面を全部 数えて返す */
function pagesWithOpenPath() {
  const pages = fs.readdirSync(ROOT).filter(f => /\.html$/i.test(f));
  const opens = pages.filter(p => fs.readFileSync(path.join(ROOT, p), 'utf8').indexOf('js/book-open.js') >= 0);
  return { pages, opens };
}

/* ── ⑧ 本番と同じ経路で計算して、実Excelの答えと合うか ──────────── */
async function testProductionPath(page) {
  console.log('\n[⑧ ' + page + ' を丸ごと載せて計算し、実Excelの答えと合うか]');
  let JSDOM;
  try { ({ JSDOM } = require('jsdom')); }
  catch { ng++; console.log('  ✗ jsdom が無い＝本番経路を通せない。★緑ではない★'); return; }

  const html = fs.readFileSync(path.join(ROOT, page), 'utf8');
  const dom = new JSDOM(html.replace(/<script[\s\S]*?<\/script>/g, ''), {
    runScripts: 'dangerously', pretendToBeVisual: true, url: 'http://localhost/',
  });
  const win = dom.window, doc = win.document;
  win.fetch = () => Promise.reject(new Error('no net'));
  win.scrollTo = () => { };
  win.alert = () => { };
  /* jsdom に無いが実ブラウザには在る物（zip の取り出しが使う） */
  win.DecompressionStream = globalThis.DecompressionStream;
  win.CompressionStream = globalThis.CompressionStream;
  win.Response = globalThis.Response;
  win.Blob = globalThis.Blob;
  const stubCtx = new Proxy({}, {
    get: (t, k) => (k === 'measureText' ? (() => ({ width: 10 }))
      : k === 'canvas' ? { width: 800, height: 600 }
        : k === 'getImageData' ? (() => ({ data: [] })) : (() => { })),
  });
  win.HTMLCanvasElement.prototype.getContext = () => stubCtx;
  const inject = (code) => { const el = doc.createElement('script'); el.textContent = code; doc.body.appendChild(el); };
  for (const m of html.matchAll(/<script src="([^"]+)"><\/script>/g)) {
    const src = m[1].split('?')[0];
    if (/^https?:/.test(src)) continue;
    inject(fs.readFileSync(path.join(ROOT, src), 'utf8'));
  }
  for (const m of html.matchAll(/<script(?![^>]*\ssrc=)[^>]*>([\s\S]*?)<\/script>/g)) inject(m[1]);
  /* ★load を必ず待つ。待たずに開くと initFormulaEngine が engine を作り直して式が全部 空になる★ */
  await new Promise(res => {
    if (doc.readyState === 'complete') return res();
    win.addEventListener('load', res);
    setTimeout(res, 3000);
  });
  /* 本番の _ensureXlsx と同じ順で読む（★table-refs.js が book-open.js より前★） */
  for (const f of ['lib/xlsx.full.min.js', 'lib/xlsx-io.js', 'lib/zip-surgeon.js',
    'lib/xlsx-edit.js', 'lib/xlsb-edit.js', 'lib/table-refs.js', 'js/book-open.js']) {
    inject(fs.readFileSync(path.join(ROOT, f), 'utf8'));
  }
  if (!ok(typeof win.BookOpen !== 'undefined', 'BookOpen が載った')) return;
  if (!ok(typeof win.TableRefs !== 'undefined', '★TableRefs が載った')) return;
  if (!ok(typeof win.initFormulaEngine === 'function' && typeof win.loadSheetIntoEngine === 'function',
    '★' + page + ' に計算する側へ流す口がある（initFormulaEngine / loadSheetIntoEngine）')) return;

  const bytes = fs.readFileSync(path.join(FIX, 'table-refs-sample.xlsb'));
  const file = {
    name: 'table-refs-sample.xlsb',
    arrayBuffer: () => Promise.resolve(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength)),
  };
  const res = await win.BookOpen.openFile(file);
  ok(res.sheets.length === 2, 'シートが2枚（' + res.sheets.length + '）');
  ok(res.opened.tableRefs && res.opened.tableRefs.fixed > 0,
    '★直した本数が残っている（' + JSON.stringify(res.opened.tableRefs) + '）');

  win.sheets = res.sheets;
  win.activeSheet = 0;
  win._engineLoaded = {};
  win.initFormulaEngine(res.sheets.map(s => s.name));
  /* ★全シートを流してから数える。1枚だけ流して数えると 参照先が空のまま計算される★ */
  for (let i = 0; i < res.sheets.length; i++) win.loadSheetIntoEngine(i);

  const idx = {};
  res.sheets.forEach((s, i) => { idx[s.name] = i; });
  let bad = 0, seen = 0;
  for (const c of gold.cases) {
    const rc = decodeAddr(c.addr);
    const sh = res.sheets[idx[c.sheet]];
    const cell = sh && sh.data[rc.r + ',' + rc.c];
    const got = cell ? cell.d : undefined;
    seen++;
    const want = c.value;
    const nGot = Number(String(got).replace(/,/g, ''));
    const same = (typeof want === 'number')
      ? (isFinite(nGot) && Math.abs(nGot - want) <= 1e-9 * Math.max(1, Math.abs(want)))
      : String(got).trim() === String(want).trim();
    if (!same) { bad++; console.log('      ' + c.sheet + '!' + c.addr + ' [' + c.label + '] うち[' + got + '] Excel[' + want + ']'); }
  }
  ok(seen === gold.cases.length, '実Excelの真値を ' + seen + ' 本 見た（検査が空振りしていない）');
  ok(bad === 0, '★' + seen + '本すべて 実Excelと同じ答えになった（違い ' + bad + '）');
  try { win.close(); } catch (e) { /* 閉じられなくても検査は済んでいる */ }
}

/* ── ⑨ 配線（本番の読み込みに入っているか）────────────────────── */
function testWiring(pages, opens) {
  console.log('\n[⑨ 配線]');
  /* ★受け取り口(book-open.js)を読んでいる画面を全部 数えてから見る★
     テスト線の book.html には受け取り口が無い（別の系統）。
     「無いから飛ばす」ではなく ★数えた物の数を出す★＝見ていないだけの緑を作らない。 */
  console.log('      ── 実測 ── 画面 ' + pages.length + '枚を見て、受け取り口を読んでいるのは ' + opens.length + '枚: ' + (opens.join(',') || '（無し）'));
  for (const p of opens) {
    const t = fs.readFileSync(path.join(ROOT, p), 'utf8');
    const iTable = t.indexOf('lib/table-refs.js');
    const iOpen = t.indexOf('js/book-open.js');
    ok(iTable > 0, p + ' が lib/table-refs.js を読んでいる');
    ok(iTable > 0 && iTable < iOpen, '★' + p + ' は table-refs.js を book-open.js より先に読む（後だと呼べない）');
  }
  const open = fs.readFileSync(path.join(ROOT, 'js', 'book-open.js'), 'utf8');
  ok(/TableRefs\.resolve\(/.test(open), '★book-open.js が TableRefs.resolve を呼んでいる');
  ok(/sheetToGrid\(wb\.Sheets\[nm\], nm, trFixes\)/.test(open), '★直した式が画面の形に渡っている');
  ok(fs.existsSync(path.join(ROOT, 'lib', 'table-refs.js')), 'lib/table-refs.js が実在する');
}

/* ── わざと壊して赤になるか ─────────────────────────────────── */
async function selfTest() {
  console.log('\n★--self-test★ わざと壊して、赤になる通り数を数える');
  const ways = [];
  const byId = { 1: { rw1: 1, rw2: 7, cl1: 0, cl2: 5, header: 1, totals: 1, cols: [], sheet: '元表', id: 1 } };
  const ctx = { row: 3, sheet: '元表' };
  const L = (o) => Object.assign({ coltype: 1, code: 0, id: 1, colFirst: 2, colLast: 2 }, o);
  const def = byId[1];

  ways.push(['① 拾った数が式より多い', TR._rewriteXlsbFormula('=SUM(Table1[#Data])', [L({}), L({})], byId, ctx) === null]);
  ways.push(['② 拾った数が式より少ない', TR._rewriteXlsbFormula('=SUM(Table1[#Data])+SUM(Table1[#Data])', [L({})], byId, ctx) === null]);
  ways.push(['③ 表idが食い違う', TR._rewriteXlsbFormula('=SUM(Table9[#Data])', [L({})], byId, ctx) === null]);
  ways.push(['④ 行の種類が食い違う', TR._rewriteXlsbFormula('=SUM(Table1[#All])', [L({})], byId, ctx) === null]);
  ways.push(['⑤ 知らない表を指している', TR._rewriteXlsbFormula('=SUM(Table7[#Data])', [L({ id: 7 })], byId, ctx) === null]);
  ways.push(['⑥ 列が表からはみ出す', TR._rewriteXlsbFormula('=SUM(Table1[#Data])', [L({ colFirst: 99 })], byId, ctx) === null]);
  ways.push(['⑦ 自分の行が中身の外', TR._toRange(def, { band: 'thisRow', curRow: 1, colFirst: 2 }) === null]);
  ways.push(['⑧ 見出し行が無いのに [#Headers]', TR._toRange({ rw1: 0, rw2: 5, cl1: 0, cl2: 2, header: 0, totals: 0, cols: [] }, { band: 'headers' }) === null]);
  ways.push(['⑨ 合計行が無いのに [#Totals]', TR._toRange({ rw1: 0, rw2: 5, cl1: 0, cl2: 2, header: 0, totals: 0, cols: [] }, { band: 'totals' }) === null]);
  ways.push(['⑩ 中身が0行の表', TR._toRange({ rw1: 0, rw2: 0, cl1: 0, cl2: 2, header: 1, totals: 0, cols: [] }, { band: 'data' }) === null]);
  ways.push(['⑪ 記録の終端が合わない', TR._rgceRange(9, new Uint8Array(4), 0, 4) === null]);
  ways.push(['⑫ 歩き終わりが合わない部品', TR._walkRecords(new Uint8Array([1, 200, 0, 0])) === null]);
  ways.push(['⑬ 表の部品が壊れている', TR._readTableBin(new Uint8Array([87, 4, 1, 2, 3, 4])) === null]);
  ways.push(['⑭ .xlsx の知らない列名', (() => {
    const d = { rw1: 1, rw2: 7, cl1: 0, cl2: 5, header: 1, totals: 1, cols: ['日付', '金額'], sheet: 'S', name: 'T' };
    return TR._rewriteTextFormula('=SUM(T[存在しない列])', [d], { row: 3, sheet: 'S' }, null) === null;
  })()]);
  ways.push(['⑮ 別シートなのにシート名が付かない形', (() => {
    const r = TR._toRange(def, { band: 'data', colFirst: 2, colLast: 2, fromSheet: '別シート' });
    return r === '元表!C3:C7';
  })()]);

  /* ★見張りそのものが空振りしていないか＝正しい物は緑のままである事も確かめる★ */
  ways.push(['⑯ まっとうな式は直る（緑のままである事）',
    TR._rewriteXlsbFormula('=SUM(Table1[#Data])', [L({})], byId, ctx) === '=SUM(C3:C7)']);

  let red = 0;
  for (const [label, caught] of ways) {
    if (caught) { red++; console.log('  ✓ ' + label + ' → 捕まえた'); }
    else { ng++; console.log('  ✗ ' + label + ' → ★素通りした（見張りの穴）'); }
    n++;
  }
  console.log('\n  ── 実測 ── わざと壊した ' + ways.length + ' 通り / 赤になった ' + red + ' 通り');
}

/* ── 実行 ─────────────────────────────────────────────────── */
console.log('[table-refs] 表の名前での参照（Table[列名]）→ A1範囲');
if (process.argv.includes('--self-test')) {
  await selfTest();
} else {
  await testTableDefsXlsb();
  await testTableDefsXlsx();
  await testRewrite('table-refs-sample.xlsb', 'xlsb', '③');
  await testRewrite('table-refs-sample.xlsx', 'xlsx', '④');
  testToRange();
  testRefuse();
  testOnlyXlsbIsBroken();
  const { pages, opens } = pagesWithOpenPath();
  if (opens.length) {
    for (const p of opens) await testProductionPath(p);
  } else {
    /* ★このリポジトリには受け取り口を持つ画面が1枚も無い（テスト線＝別の系統）。
       ★「対象が0枚」と数で言う。飛ばした事を緑に見せない★ */
    console.log('\n[⑧ 画面を丸ごと載せる検査]');
    console.log('      ── 実測 ── 画面 ' + pages.length + '枚を見て、受け取り口(js/book-open.js)を読んでいるのは ★0枚★');
    console.log('      ＝この repo の画面は受け取り口を持たない系統。画面ごしの検査は ★対象が無い★（③④⑤⑥⑦で部品は測っている）');
  }
  testWiring(pages, opens);
}
console.log('\n' + (n - ng) + ' passed, ' + ng + ' failed');
process.exit(ng ? 1 : 0);
