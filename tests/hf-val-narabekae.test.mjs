/* hf-val-narabekae.test.mjs — ★並べ替えなどで 板を 流し直しても 打った 値は 打った時と 同じ 形で 計算側（HF）へ★（2026-11-13 件e）
 *  ★なぜ★ _pushGrid が cell.v を 字のまま HF へ 流していた＝新しい本で 打った '1,234'・'2,000'・'2026/8/1' が
 *   並べ替えの 後 HF の 中で 字になり ★=SORT(B2:B4) の 先頭が 5 でなく「1,234」★（並べ替え 無しなら 5）。
 *   =A2+30・SUM は 台が 答えるので 画面は 合う＝★HF に 落ちる 式（溢れ）だけ 黙って 違う★。
 *  ★ファイルの 字の マス（d字）は 字のまま★＝通すと '1,234' が 数に 化ける。
 *  使い方: node tests/hf-val-narabekae.test.mjs ／ --self-test＝直しを 1つずつ 壊して（3通り）赤になるか */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(path.join(ROOT, 'package.json'));
let JSDOM;
try { ({ JSDOM } = require_('jsdom')); }
catch { console.log('★jsdomが入っていません（SKIPを緑と呼ばない）。npm ci してください。'); process.exit(1); }

const 元のhtml = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf8');
const 元のエンジン = fs.readFileSync(path.join(ROOT, 'exally-formula.js'), 'utf8');

/* ★壊し方★＝直しを 1つずつ 元の形へ 戻す。★当たらなければ 壊せていない＝その場で 落とす★ */
const 直し = "(cell.d字 ? cell.v : toHFVal(cell.v))";
const 壊し方 = {
  通さない: (h) => h.replace(直し, 'cell.v'),
  字も通す: (h) => h.replace(直し, 'toHFVal(cell.v)'),
  持ち越し: (h) => h.replace("  _打った答えは字か = false;\n  // 時刻は文字列として扱う", "  // 時刻は文字列として扱う"),
};

const CANVAS_STUB = `(function(){var noop=function(){};var ctx=new Proxy({},{get:function(t,k){if(k==='measureText')return function(){return{width:40};};if(k==='canvas')return{width:800,height:600};if(k==='getImageData')return function(){return{data:[]};};if(k==='createLinearGradient'||k==='createPattern')return function(){return{addColorStop:noop};};return noop;}});HTMLCanvasElement.prototype.getContext=function(){return ctx;};})();`;
function 開く(html) {
  const dom = new JSDOM(html.replace(/<script[\s\S]*?<\/script>/g, ''), {
    runScripts: 'dangerously', url: 'http://localhost/', pretendToBeVisual: true,
    beforeParse(w) {
      w.fetch = () => Promise.reject(new Error('no net')); w.scrollTo = () => {}; w.alert = () => {}; w.confirm = () => true;
      w.matchMedia = w.matchMedia || (() => ({ matches: false, addListener() {}, removeListener() {}, addEventListener() {}, removeEventListener() {} }));
      w.requestAnimationFrame = (cb) => setTimeout(() => cb(Date.now()), 0); w.cancelAnimationFrame = () => {};
      w.console = { log() {}, info() {}, warn() {}, error() {}, debug() {} };
      w.eval(CANVAS_STUB);
    },
  });
  const win = dom.window, doc = win.document;
  const inject = (code) => { const el = doc.createElement('script'); el.textContent = code; doc.body.appendChild(el); };
  for (const src of [...html.matchAll(/<script src="([^"]+)"><\/script>/g)].map((m) => m[1].split('?')[0]).filter((s) => !/^https?:/.test(s))) {
    const p = path.join(ROOT, src);
    if (src === 'exally-formula.js') { inject(元のエンジン); continue; }
    if (fs.existsSync(p)) { try { inject(fs.readFileSync(p, 'utf8')); } catch (e) { /* 続ける */ } }
  }
  for (const m of html.matchAll(/<script(?![^>]*\ssrc=)[^>]*>([\s\S]*?)<\/script>/g)) { try { inject(m[1]); } catch (e) { /* 続ける */ } }
  try { doc.dispatchEvent(new win.Event('DOMContentLoaded', { bubbles: true })); } catch (e) { /* 続ける */ }
  try { win.dispatchEvent(new win.Event('load')); } catch (e) { /* 続ける */ }
  return win;
}
const 打つ = (w, i, r, c, v) => w.eval(`switchSheet(${i}); setCell(${r},${c},${JSON.stringify(v)}); loadSheetIntoEngine(${i});`);
const 画 = (w, i, r, c) => w.eval(`(sheets[${i}].data['${r},${c}']||{}).d`);
const 名たち = (w) => JSON.stringify({ HF: w.eval('hf.getSheetNames()'), 画面: w.eval('sheets.map(function(s){return s.name})') });
const 揃う = (w) => { const x = JSON.parse(名たち(w)); if (JSON.stringify(x.HF) !== JSON.stringify(x.画面)) throw new Error('板の名が 揃わない ' + 名たち(w)); };
const eq = (a, b, m) => { if (String(a) !== String(b)) throw new Error((m ? m + ': ' : '') + '期待=' + JSON.stringify(b) + ' 実際=' + JSON.stringify(a)); };
const hfv = (w, i, r, c) => w.eval(`JSON.stringify(hf.getCellValue({sheet:${i},row:${r},col:${c}}))`);

function 回す(html, 黙る) {
  let pass = 0, fail = 0;
  const T = (n, fn) => {
    try { fn(); pass++; if (!黙る) console.log('  ok   ' + n); }
    catch (e) { fail++; if (!黙る) console.log('  NG   ' + n + '\n       ' + (e && e.message)); }
  };
  const 表を打つ = (w) => {
    打つ(w, 0, 0, 0, '日'); 打つ(w, 0, 0, 1, '額');
    打つ(w, 0, 1, 0, '2026/8/31'); 打つ(w, 0, 1, 1, '1,234');
    打つ(w, 0, 2, 0, '2026/8/1');  打つ(w, 0, 2, 1, '5');
    打つ(w, 0, 3, 0, '2026/8/15'); 打つ(w, 0, 3, 1, '2,000');
  };
  const 並べる = (w) => w.eval("selR1=0;selC1=0;selR2=3;selC2=1; sortRange('asc');");

  T('並べ替えの 後も =SORT(B2:B4) の 先頭は 5（打った 1,234・2,000 は 数のまま）', () => {
    const w = 開く(html); 表を打つ(w); 並べる(w);
    eq(w.eval("sheets[0].data['1,1'].v"), '5', '並べ替えが 効いた');
    打つ(w, 0, 0, 7, '=SORT(B2:B4)');
    eq(画(w, 0, 0, 7), '5');
    eq(hfv(w, 0, 2, 1), '2000', 'HF の B3'); eq(hfv(w, 0, 3, 1), '1234', 'HF の B4');
  });
  T('並べ替えの 後も 打った 日付は 計算側で 日付の 数（2026/8/1 ＝ 46235）', () => {
    const w = 開く(html); 表を打つ(w); 並べる(w);
    eq(hfv(w, 0, 1, 0), '46235', 'HF の A2');
    打つ(w, 0, 0, 6, '=SORT(A2:A4)+0'); eq(画(w, 0, 0, 6), '46235');
  });
  T('字の 答えの 式（=TEXT）の 直後に 打った 日付も 並べ替えの 後 日付の 数（d字 を 持ち越さない）', () => {
    const w = 開く(html);
    打つ(w, 0, 0, 0, '日'); 打つ(w, 0, 0, 1, '額');
    打つ(w, 0, 0, 5, '=TEXT(1,"0")'); 打つ(w, 0, 1, 0, '2026/8/31');
    打つ(w, 0, 0, 5, '=TEXT(1,"0")'); 打つ(w, 0, 2, 0, '2026/8/1');
    打つ(w, 0, 0, 5, '=TEXT(1,"0")'); 打つ(w, 0, 3, 0, '2026/8/15');
    打つ(w, 0, 1, 1, '1'); 打つ(w, 0, 2, 1, '2'); 打つ(w, 0, 3, 1, '3');
    eq(w.eval("!!sheets[0].data['2,0'].d字"), false, '打った 日付の d字');
    並べる(w);
    eq(hfv(w, 0, 1, 0), '46235', 'HF の A2');
    打つ(w, 0, 0, 6, '=SORT(A2:A4)+0'); eq(画(w, 0, 0, 6), '46235');
  });
  T('ファイルの 字の マス（d字）は 流し直しても 字のまま（数に 化けない）', () => {
    const w = 開く(html);
    w.eval(`sheets[0].data['0,0'] = { v:'1,234', f:'1,234', d:'1,234', d字:true };
            sheets[0].data['1,0'] = { v:5, f:'5', d:'5' };
            _pushGrid(0, sheets[0]);`);
    eq(hfv(w, 0, 0, 0), '"1,234"', 'HF の A1（字）'); eq(hfv(w, 0, 1, 0), '5', 'HF の A2（数）');
  });
  return { pass, fail };
}

if (process.argv.includes('--self-test')) {
  console.log('\n[hf-val-narabekae --self-test] ★直しを 1つずつ 壊すと 赤になるか★');
  let 空振り = 0;
  for (const [名, 壊す] of Object.entries(壊し方)) {
    const h = 壊す(元のhtml);
    if (h === 元のhtml) { console.log('  NG   壊し方「' + 名 + '」が 当たらない（book.html の形が変わった）'); 空振り++; continue; }
    const r = 回す(h, true);
    if (r.fail > 0) console.log('  ok   「' + 名 + '」を壊すと 赤 ' + r.fail);
    else { console.log('  NG   「' + 名 + '」を壊しても 緑＝空振り'); 空振り++; }
  }
  console.log(空振り ? '\n★空振り ' + 空振り + '★' : '\n全部 赤になった');
  process.exit(空振り ? 1 : 0);
}
console.log('\n[hf-val-narabekae] 板を 流し直しても 打った 値は 打った時と 同じ 形で 計算側へ');
const r = 回す(元のhtml, false);
console.log('\n' + r.pass + ' ok / ' + r.fail + ' NG');
process.exit(r.fail ? 1 : 0);
