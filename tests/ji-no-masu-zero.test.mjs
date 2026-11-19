/* ji-no-masu-zero.test.mjs — ★字の マス '007 の 頭の 0 を 落とさない★（2026-11-19）
 *  ★なぜ★ ファイルから 読んだ 字の マス（d字）'007' を =A1&"x" が 使うと 台が 7 に 読み "7x"（実Excel 007x）・
 *   HF も '0123' を 123 に 読んだ（開いた時・並べ替え・値の取り消し・ゴールシークの戻し）。
 *   手で '007 と打つと 台が ' ごと 字にし "'007x"・LEN 4（実Excel 007x・3）。
 *  ★直し★ 台の 打つ＝頭の ' は 字の 印（外す）／打った物＝d字 の 字の マスは "'"+v（式かは f の 頭）／HF の 3口＝"'"+v
 *  使い方: node tests/ji-no-masu-zero.test.mjs ／ --self-test＝直しを 1つずつ 壊して（6通り）赤になるか */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(path.join(ROOT, 'package.json'));
let JSDOM;
try { ({ JSDOM } = require_('jsdom')); }
catch { console.log('★jsdomが入っていません（SKIPを緑と呼ばない）。npm ci してください。'); process.exit(1); }

const 元のエンジン = fs.readFileSync(path.join(ROOT, 'exally-formula.js'), 'utf8');
const 読む = (f) => fs.readFileSync(path.join(ROOT, f), 'utf8');
const 元 = { 'book.html': 読む('book.html'), 'lib/shiki-hyou.js': 読む('lib/shiki-hyou.js'), 'lib/shiki-ita-awase.js': 読む('lib/shiki-ita-awase.js') };

/* ★壊し方★＝直しを 1つずつ 元の形へ 戻す。★当たらなければ 壊せていない＝その場で 落とす★ */
const 壊し方 = {
  "台：頭の ' を外さない": ['lib/shiki-hyou.js', `} else if (s.charAt(0) === "'") {`, `} else if (false) {`],
  "打った物：d字 に ' を付けない": ['lib/shiki-ita-awase.js', `if (セル.d字 && 字 !== '') return "'" + 字;`, `if (false) return "'" + 字;`],
  '打った物：式を v の頭で判じる': ['lib/shiki-ita-awase.js', `if (f.charAt(0) === '=') return f;`, `if (f.charAt(0) === '=' || v.charAt(0) === '=') return f || v;`],
  'HF：板の流し直し': ['book.html', `(cell.d字 ? "'" + cell.v : toHFVal(cell.v))`, `(cell.d字 ? cell.v : toHFVal(cell.v))`],
  'HF：値の取り消し': ['book.html', `row:r,col:c}, "'" + String(op.old)); }catch(e){}`, `row:r,col:c}, String(op.old)); }catch(e){}`],
  'HF：ゴールシークの戻し': ['book.html', `(控え.d字 ? "'" + 控え.v : toHFVal(控え.v))`, `(控え.d字 ? 控え.v : toHFVal(控え.v))`],
};

const CANVAS_STUB = `(function(){var noop=function(){};var ctx=new Proxy({},{get:function(t,k){if(k==='measureText')return function(){return{width:40};};if(k==='canvas')return{width:800,height:600};if(k==='getImageData')return function(){return{data:[]};};if(k==='createLinearGradient'||k==='createPattern')return function(){return{addColorStop:noop};};return noop;}});HTMLCanvasElement.prototype.getContext=function(){return ctx;};})();`;
function 開く(物) {
  const html = 物['book.html'];
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
    if (src === 'exally-formula.js') { inject(元のエンジン); continue; }
    if (物[src] !== undefined) { inject(物[src]); continue; }
    const p = path.join(ROOT, src);
    if (fs.existsSync(p)) { try { inject(fs.readFileSync(p, 'utf8')); } catch (e) { /* 続ける */ } }
  }
  for (const m of html.matchAll(/<script(?![^>]*\ssrc=)[^>]*>([\s\S]*?)<\/script>/g)) { try { inject(m[1]); } catch (e) { /* 続ける */ } }
  try { doc.dispatchEvent(new win.Event('DOMContentLoaded', { bubbles: true })); } catch (e) { /* 続ける */ }
  try { win.dispatchEvent(new win.Event('load')); } catch (e) { /* 続ける */ }
  win.eval('switchSheet(0)');
  return win;
}
/* ★打った後の 計算し直しは 150ms 後（_scheduleRecalc）★＝試験では 同じ recalcSheet を その場で 済ませる */
const 打つ = (w, r, c, v) => w.eval(`switchSheet(0); setCell(${r},${c},${JSON.stringify(v)}); recalcSheet(0, sheets[0].data);`);
/* ★ファイルから 読んだ 字の マス★＝js/book-open.js の 形（v・d・d字）。★開いた 時と 同じく 板ごと 流す★ */
const 字のマス = (w, r, c, v) => w.eval(`sheets[0].data['${r},${c}']={v:${JSON.stringify(v)},f:'',d:${JSON.stringify(v)},d字:true};`);
const 流す = (w) => w.eval('_pushGrid(0, sheets[0]); recalcSheet(0, sheets[0].data);');
const 画 = (w, r, c) => w.eval(`String((sheets[0].data['${r},${c}']||{}).d)`);
const hfv = (w, r, c) => w.eval(`JSON.stringify(hf.getCellValue({sheet:_hfSid(0),row:${r},col:${c}}))`);
const eq = (a, b, m) => { if (String(a) !== String(b)) throw new Error((m ? m + ': ' : '') + '期待=' + JSON.stringify(b) + ' 実際=' + JSON.stringify(a)); };

function 回す(物, 黙る) {
  let pass = 0, fail = 0;
  const T = (n, fn) => {
    try { fn(); pass++; if (!黙る) console.log('  ok   ' + n); }
    catch (e) { fail++; if (!黙る) console.log('  NG   ' + n + '\n       ' + (e && e.message)); }
  };
  T('ファイルの字 007 を =A1&"x" が使うと 007x・LEN 3（台）', () => {
    const w = 開く(物); 字のマス(w, 0, 0, '007'); 流す(w);
    打つ(w, 0, 1, '=A1&"x"'); 打つ(w, 0, 2, '=LEN(A1)'); 打つ(w, 0, 3, '=ISTEXT(A1)');
    eq(画(w, 0, 1), '007x', 'B1'); eq(画(w, 0, 2), '3', 'LEN'); eq(画(w, 0, 3), 'TRUE', 'ISTEXT');
  });
  T('開いた直後の 計算側（HF）も 字の "007"', () => {
    const w = 開く(物); 字のマス(w, 0, 0, '007'); 流す(w);
    eq(hfv(w, 0, 0), '"007"');
  });
  T('並べ替えの後も 字のまま（=SORT の 答えは 007）', () => {
    const w = 開く(物); 字のマス(w, 0, 0, '0123'); 字のマス(w, 1, 0, '007'); 流す(w);
    w.eval("selR1=0;selC1=0;selR2=1;selC2=0; sortRange('asc');");
    eq(hfv(w, 0, 0), '"007"', 'HF A1'); eq(hfv(w, 1, 0), '"0123"', 'HF A2');
    打つ(w, 0, 4, '=SORT(A1:A2)'); eq(画(w, 0, 4), '007', 'SORT');
  });
  T('9 を打って Ctrl+Z で戻すと 台も HF も 007', () => {
    const w = 開く(物); 字のマス(w, 0, 0, '007'); 流す(w);
    打つ(w, 0, 1, '=A1&"x"'); 打つ(w, 0, 0, '9'); eq(画(w, 0, 1), '9x', '打った後');
    w.eval('doUndo(); recalcSheet(0, sheets[0].data);');
    eq(hfv(w, 0, 0), '"007"', 'HF A1'); eq(画(w, 0, 1), '007x', '台の B1');
  });
  T('ゴールシークの戻しで 字の マスを 戻すと HF も 007', () => {
    const w = 開く(物); 字のマス(w, 0, 0, '007'); 流す(w);
    打つ(w, 0, 0, '5');
    w.eval("_ゴールの元へ戻す(0, 0, 0, {v:'007',f:'',d:'007',d字:true})");
    eq(hfv(w, 0, 0), '"007"');
  });
  T("手で '007 と打つと =A1&\"x\" は 007x・LEN 3（' は答えに入らない）", () => {
    const w = 開く(物); 打つ(w, 0, 0, "'007"); 打つ(w, 0, 1, '=A1&"x"'); 打つ(w, 0, 2, '=LEN(A1)');
    eq(画(w, 0, 1), '007x', 'B1'); eq(画(w, 0, 2), '3', 'LEN');
  });
  T('ファイルの字 "=abc" は 式に ならない（=abcx・ISFORMULA FALSE）', () => {
    const w = 開く(物); 字のマス(w, 0, 0, '=abc'); 流す(w);
    打つ(w, 0, 1, '=A1&"x"'); 打つ(w, 0, 2, '=ISFORMULA(A1)');
    eq(画(w, 0, 1), '=abcx', 'B1'); eq(画(w, 0, 2), 'FALSE', 'ISFORMULA');
  });
  T("打った数 007（' なし）は 前と同じく 数の 7", () => {
    const w = 開く(物); 打つ(w, 0, 0, '007'); 打つ(w, 0, 1, '=A1+1'); eq(画(w, 0, 1), '8');
  });
  T('式が返す字（=TEXT(7,"000")）は 前と同じ 007x', () => {
    const w = 開く(物); 打つ(w, 0, 0, '=TEXT(7,"000")'); 打つ(w, 0, 1, '=A1&"x"'); eq(画(w, 0, 1), '007x');
  });
  return { pass, fail };
}

if (process.argv.includes('--self-test')) {
  let 赤 = 0, 全 = 0;
  for (const [名, [f, a, b]] of Object.entries(壊し方)) {
    全++;
    if (!元[f].includes(a)) { console.log('  NG   壊せない（字が 当たらない）: ' + 名); process.exit(1); }
    const 物 = { ...元, [f]: 元[f].replace(a, () => b) };
    const r = 回す(物, true);
    if (r.fail > 0) { 赤++; console.log('  ok   守り「' + 名 + '」を 抜く→赤 ' + r.fail); }
    else console.log('  NG   守り「' + 名 + '」を 抜いても 緑のまま');
  }
  console.log(`\n${赤}/${全} 通りで赤くなった`);
  process.exit(赤 === 全 ? 0 : 1);
}
const r = 回す(元, false);
console.log(`\n字の マスの 頭の 0: 緑 ${r.pass} / 赤 ${r.fail}`);
process.exit(r.fail ? 1 : 0);
