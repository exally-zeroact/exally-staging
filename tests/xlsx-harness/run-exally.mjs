// run-exally.mjs — ★本番と同じ経路で計算する。
//
//   なぜ「本番経路」なのか:
//     book.html の setCellFormula() は
//       ① _jsComputeFormula(独自層。約60関数を横取りしてJSで計算)
//       ② convertFormula(文字列書換。TRUE→TRUE() など)→ HyperFormula
//       ③ HFがエラーを返したら evalFormula で再挑戦
//     の三段になっている。生の HyperFormula で測ると実測で5件も答えが違い、検証にならない。
//     だから book.html を丸ごと jsdom に載せて、本物の setCellFormula を呼ぶ。
//
//   ついでに3つ固定する(将来 book.html が変わって生HFに落ちたら気付けるように):
//     (a) 経路スパイ  … 独自層/書換/HF のどれが答えを出したかを1件ずつ記録
//     (b) 生HFとの差分 … 同じケースを生HFでも回して「独自層が救済している件数と中身」を出す
//     (c) 入口の単一性 … 独自層の入口が _jsComputeFormula の対象表ただ1つであること
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createRequire } from 'node:module';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
export const ROOT = path.join(__dirname, '..', '..');
const require = createRequire(pathToFileURL(path.join(ROOT, 'package.json')));

/* ── ケース読み込み ───────────────────────────────────────── */
export function loadCases() {
  const dir = path.join(__dirname, 'cases');
  const inputs = JSON.parse(fs.readFileSync(path.join(dir, '_inputs.json'), 'utf8'));
  const files = fs.readdirSync(dir).filter(f => f.endsWith('.json') && f !== '_inputs.json').sort();
  const cases = [];
  for (const f of files) {
    const j = JSON.parse(fs.readFileSync(path.join(dir, f), 'utf8'));
    for (const c of j.cases) cases.push({ ...c, group: j.group, file: f });
  }
  const seen = new Set();
  for (const c of cases) {
    if (seen.has(c.id)) throw new Error('ケースIDが重複: ' + c.id);
    seen.add(c.id);
  }
  return { inputs, cases, files };
}

/* ── 入力セルの書き込み ────────────────────────────────────
 *  ★文字列は ="..." の形で入れる。
 *    理由(実測): HyperFormula は setCellContents('0007') を数値7に変換してしまい、
 *    文字列セルを文字列のまま持てない。Excel側は NumberFormat='@' で文字列のまま。
 *    そのままだと「関数の検証」ではなく「入力の化け」を測る事になり、比べる物がズレる。
 *    ="0007" なら両者とも「文字列の0007を持つセル」になり、関数の比較として正しい。
 *  ★入力そのものが化ける件(グリッドの実挙動)は別枠で測る=下の inputFidelity。
 */
function addrToRC(addr) {
  const m = addr.match(/^([A-Z]+)(\d+)$/);
  let col = 0;
  for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64);
  return { r: parseInt(m[2], 10) - 1, c: col - 1 };
}
function inputValueFor(spec) {
  if (spec.t === 'n') return spec.v;
  if (spec.t === 'b') return spec.v;
  if (spec.t === 's') return '="' + String(spec.v).replace(/"/g, '""') + '"';
  throw new Error('unknown input type: ' + spec.t);
}

/* ── 生の HyperFormula(比較用) ─────────────────────────────── */
export function runRawHF(inputs, cases) {
  const { HyperFormula } = require(path.join(ROOT, 'hyperformula.full.min.js'));
  const ERR = { DIV_BY_ZERO: '#DIV/0!', NUM: '#NUM!', NA: '#N/A', VALUE: '#VALUE!', REF: '#REF!', NAME: '#NAME?', CYCLE: '#CYCLE!', NULL: '#NULL!', SPILL: '#SPILL!' };
  const hf = HyperFormula.buildEmpty({ licenseKey: 'gpl-v3' });
  hf.addSheet('Sheet1');
  const sid = hf.getSheetId('Sheet1');
  for (const [addr, spec] of Object.entries(inputs.cells)) {
    const { r, c } = addrToRC(addr);
    hf.setCellContents({ sheet: sid, row: r, col: c }, inputValueFor(spec));
  }
  const out = {};
  let row = 0;
  for (const cs of cases) {
    const cell = { sheet: sid, row: row++, col: 20 };
    let v;
    try {
      hf.setCellContents(cell, cs.f);
      v = hf.getCellValue(cell);
    } catch (e) { v = 'THROW:' + (e && e.message); }
    if (v && typeof v === 'object' && v.type) v = ERR[v.type] || ('#' + v.type);
    else if (Array.isArray(v)) v = 'ARRAY';
    out[cs.id] = { v, jsType: v === null ? 'null' : typeof v };
  }
  return out;
}

/* ── 本番経路(book.html) ──────────────────────────────────── */
export async function runProductionPath(inputs, cases) {
  let JSDOM;
  try { ({ JSDOM } = require('jsdom')); }
  catch { return { skipped: 'jsdom未導入(npm i)=本番経路の検証をスキップ。★緑ではない★' }; }

  const html = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf8');
  // 本物のDOMを作る(script は後から順に流す)。UIの初期化が動くので実物に近い。
  const dom = new JSDOM(html.replace(/<script[\s\S]*?<\/script>/g, ''), {
    runScripts: 'dangerously', pretendToBeVisual: true, url: 'http://localhost/'
  });
  const win = dom.window, doc = win.document;
  win.fetch = () => Promise.reject(new Error('no net'));
  win.scrollTo = () => {};
  win.alert = () => {};
  // jsdom に canvas は無い。グリッド描画は検証対象外なので最小限のスタブを与える
  //  (描画で落ちても計算エンジンの検証は続けられるようにする)。
  const stubCtx = new Proxy({}, {
    get: (t, k) => (k === 'measureText' ? (() => ({ width: 10 }))
      : k === 'canvas' ? { width: 800, height: 600 }
        : k === 'getImageData' ? (() => ({ data: [] })) : (() => {}))
  });
  win.HTMLCanvasElement.prototype.getContext = () => stubCtx;

  const winErrors = [];
  win.addEventListener('error', e => winErrors.push('window.error: ' + (e.message || e)));
  win.addEventListener('unhandledrejection', e => winErrors.push('unhandledrejection: ' + (e.reason?.message || e.reason)));

  const injectErrors = [];
  const inject = (code, tag) => {
    const el = doc.createElement('script');
    el.textContent = code;
    try { doc.body.appendChild(el); } catch (e) { injectErrors.push(tag + ': ' + e.message); }
  };
  for (const m of html.matchAll(/<script src="([^"]+)"><\/script>/g)) {
    const src = m[1].split('?')[0];               // ★?v= を落として実ファイルを読む
    if (/^https?:/.test(src)) continue;           // 外部CDNは読まない(ネットに出ない)
    inject(fs.readFileSync(path.join(ROOT, src), 'utf8'), src);
  }
  let n = 0;
  for (const m of html.matchAll(/<script(?![^>]*\ssrc=)[^>]*>([\s\S]*?)<\/script>/g)) inject(m[1], 'inline#' + (++n));

  if (typeof win.initFormulaEngine !== 'function' || typeof win.setCellFormula !== 'function') {
    throw new Error('book.html から initFormulaEngine / setCellFormula が見つからない(本番経路を通せない)');
  }
  win.initFormulaEngine(['Sheet1']);
  if (!win.hf) throw new Error('HyperFormula の初期化に失敗(本番経路を通せない)');

  /* ── (a) 経路スパイ ── */
  const spy = { js: 0, convert: 0, hfSet: 0, current: null, perCase: {} };
  const origJs = win._jsComputeFormula, origConv = win.convertFormula;
  const origHfSet = win.hf.setCellContents.bind(win.hf);
  win._jsComputeFormula = function (sheet, v) {
    const r = origJs.apply(this, arguments);
    if (spy.current) { spy.perCase[spy.current].jsCalled = true; if (r !== null) spy.perCase[spy.current].jsAnswered = true; }
    spy.js++;
    return r;
  };
  win.convertFormula = function (f) {
    const r = origConv.apply(this, arguments);
    if (spy.current) { spy.perCase[spy.current].convertCalled = true; if (r !== f) spy.perCase[spy.current].rewritten = r; }
    spy.convert++;
    return r;
  };
  win.hf.setCellContents = function (...a) { if (spy.current) spy.perCase[spy.current].hfSet = true; spy.hfSet++; return origHfSet(...a); };

  /* ── 入力セル ── */
  for (const [addr, spec] of Object.entries(inputs.cells)) {
    const { r, c } = addrToRC(addr);
    win.hf.setCellContents({ sheet: win.hf.getSheetId('Sheet1'), row: r, col: c }, inputValueFor(spec));
  }

  /* ── 全ケースを本番経路で ── */
  const out = {};
  let row = 0;
  for (const cs of cases) {
    spy.current = cs.id;
    spy.perCase[cs.id] = {};
    let v;
    try { v = win.setCellFormula('Sheet1', row, 20, cs.f); }
    catch (e) { v = 'THROW:' + (e && e.message); }
    row++;
    out[cs.id] = { v: v === null || v === undefined ? null : String(v), route: spy.perCase[cs.id] };
  }
  spy.current = null;

  /* ── 入力の型が保たれるか(グリッドの実挙動。★別枠で正直に測る) ── */
  //  ★期待値はここでは決めない。Excelが同じ文字をどう解釈したか(golden)と後で突き合わせる。
  const inputFidelity = {};
  let pr = 500;
  for (const p of (inputs.inputProbes?.probes || [])) {
    win.setCellFormula('Sheet1', pr, 25, p.raw);                     // 本物の入力経路で入れて
    const got = win.hf.getCellValue({ sheet: win.hf.getSheetId('Sheet1'), row: pr, col: 25 });
    inputFidelity[p.id] = {
      label: p.label, raw: p.raw,
      got: got === null ? null : String(got),
      jsType: got === null ? 'null' : typeof got
    };
    pr++;
  }

  /* ── (c) 独自層の入口が1つだけか ── */
  const formulaSrc = fs.readFileSync(path.join(ROOT, 'exally-formula.js'), 'utf8');
  const jsSetCount = (formulaSrc.match(/var\s+_jsSet\s*=/g) || []).length;
  const entryPoints = (formulaSrc.match(/function\s+_jsComputeFormula\s*\(/g) || []).length;
  const pluginRegistered = win._pluginRegistered === true;
  const pluginFuncs = Array.isArray(win._PLUGIN_FUNCS) ? win._PLUGIN_FUNCS.slice().sort() : [];
  const jsSetNames = (() => {
    const m = formulaSrc.match(/var _jsSet = \{([\s\S]*?)\};/);
    return m ? [...m[1].matchAll(/([A-Z][A-Z0-9.]*)\s*:\s*1/g)].map(x => x[1]).sort() : [];
  })();

  const closeErrs = [];
  try { dom.window.close(); } catch (e) { closeErrs.push(e.message); }

  return {
    out, inputFidelity,
    spyTotals: { js: spy.js, convert: spy.convert, hfSet: spy.hfSet },
    entry: { jsSetCount, entryPoints, pluginRegistered, pluginCount: pluginFuncs.length },
    split: { plugin: pluginFuncs, jsSet: jsSetNames },
    domErrors: winErrors.concat(injectErrors)
  };
}

/* ── 単体実行 ── */
if (process.argv[1] && pathToFileURL(process.argv[1]).href === import.meta.url) {
  const { inputs, cases } = loadCases();
  const res = await runProductionPath(inputs, cases);
  console.log('本番経路:', Object.keys(res.out || {}).length, '件 / DOM例外', (res.domErrors || []).length, '件');
  console.log('経路スパイ合計:', JSON.stringify(res.spyTotals));
  console.log('入力の型:', JSON.stringify(res.inputFidelity, null, 1));
}
