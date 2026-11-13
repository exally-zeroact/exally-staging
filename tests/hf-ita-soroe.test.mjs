/* hf-ita-soroe.test.mjs — ★計算側（HyperFormula）の板と 画面の sheets[] を 番号も名も 揃える★（件B'・2026-11-13）
 *  ★なぜ★ 本番 818bd778 で 実測（jsdom・客の画面の data.d）：
 *    ㋑ 給料明細を受け取る（book.html receivePasteData が sheets[0].name だけ替え、HF は Sheet1 のまま）
 *       → シートを足して =SORT('給料明細'!A1:A2) ＝ ★"#REF!"★（正 50）
 *    ㋒ そのあと タブを作り直す（rebuildSheetTabs が 名で 板を足す＝HF に「給料明細」が 3枚目に増える）
 *       → 足した Sheet3 に 打った 7 が ★HF の「給料明細」の板★に 入る（番号は そのまま渡すため）
 *         ＝ =SORT('給料明細'!A1:A2) が ★"7"★（正 50）＝★黙って 違う答え★
 *    ・切り出しの 新しいシートの 中身は HF に 流れていなかった
 *    ・つなげて出す は 名が「つないだ表」固定＝2回で 同じ名が 2枚
 *    ・名の比べが 大小を区別（HF は区別しない＝'DATA' が在ると 'Data' の板は 黙って捨てられる）
 *    ・initFormulaEngine の板の輪が 投げると initExallyFormula が走らず hf と _hf が 別物に分かれる
 *  ★普通の参照（='給料明細'!A1）は 台が 先に答えるので 画面は 正しい＝★赤が出るのは HF に落ちる式（溢れる式 SORT 等・book.html _台に聞く）★
 *  ★INDIRECT の 板またぎは 台が 自分で #REF! を返す（lib/shiki-kansuu.js の INDIRECT・未測定）＝物差しに 使わない★
 *  使い方: node tests/hf-ita-soroe.test.mjs ／ --self-test＝直しを 1つずつ 壊して（5通り）
 *  ★引用符なしの 日本語の 板名（=SORT(切り!A1:A2)）は 打った式が HF で #ERROR＝前からの 別の穴（この試験は 引用符つきで 書く）★それぞれ 赤になるか
 */
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
const 壊し方 = {
  揃える: (h) => h.replace('function 板を揃える(){', 'function 板を揃える(){ return;'),
  中身: (h) => h.replace('if(_中身が在る(sheets[i])) _pushGrid(i, sheets[i]);', ''),
  大小: (h) => h.replace("var 小 = function(x){ return String(x).toLowerCase(); };", "var 小 = function(x){ return String(x); };"),
  つなぎの名: (h) => h.replace("var 名 = 使えるシート名('つないだ表');", "var 名 = 'つないだ表';"),
  板の輪: (h) => h.replace("    try { hf.addSheet(sheetNames[i]); }\n    catch(e){ hf.addSheet(); 板がずれた('「' + sheetNames[i] + '」と 同じ名の シートが 在ります'); }", '    hf.addSheet(sheetNames[i]);'),
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

function 回す(html, 黙る) {
  let pass = 0, fail = 0;
  const T = (n, fn) => {
    try { fn(); pass++; if (!黙る) console.log('  ok   ' + n); }
    catch (e) { fail++; if (!黙る) console.log('  NG   ' + n + '\n       ' + (e && e.message)); }
  };
  const 給料明細 = (w) => w.eval(`receivePasteData({payType:'monthly',sheets:[{cells:[{r:0,c:0,v:100},{r:1,c:0,v:50}]}]})`);
  const 名の列 = (w) => JSON.parse(w.eval('JSON.stringify(sheets.map(function(s){return s.name}))'));

  T('㋐ 名を替えない時は 前と同じ（Sheet2 から =SORT(Sheet1!A1:A2) ＝ 50）', () => {
    const w = 開く(html);
    打つ(w, 0, 0, 0, 100); 打つ(w, 0, 1, 0, 50); w.eval('addSheet()'); 打つ(w, 1, 0, 0, '=SORT(Sheet1!A1:A2)');
    eq(画(w, 1, 0, 0), '50'); 揃う(w);
  });
  T("㋑ 給料明細を受け取って 足したシートから =SORT('給料明細'!A1:A2) ＝ 50", () => {
    const w = 開く(html); 給料明細(w);
    w.eval('addSheet()'); 打つ(w, 1, 0, 0, "=SORT('給料明細'!A1:A2)");
    eq(画(w, 1, 0, 0), '50'); 揃う(w);
  });
  T('㋒ タブを作り直して 足した Sheet3 の 8・7 は Sheet3 に入る（給料明細は 50 のまま）', () => {
    const w = 開く(html); 給料明細(w);
    w.eval('addSheet()'); w.eval('rebuildSheetTabs()'); w.eval('addSheet()'); 打つ(w, 2, 0, 0, 8); 打つ(w, 2, 1, 0, 7);
    打つ(w, 1, 0, 0, "=SORT('給料明細'!A1:A2)"); 打つ(w, 1, 0, 2, '=SORT(Sheet3!A1:A2)');
    eq(画(w, 1, 0, 0), '50', '給料明細'); eq(画(w, 1, 0, 2), '7', 'Sheet3'); 揃う(w);
  });
  T("切り出しの 新しいシートの 中身が 計算側に入る（=SORT('切り'!A1:A2) ＝ 3）", () => {
    const w = 開く(html);
    w.eval(`_手順の待ち = { シート:0, 変える:{}, 出すシート:[{ name:'切り', data:{'0,0':{v:5},'1,0':{v:3}}, 何行:2 }] }; 手順を書く();`);
    打つ(w, 0, 0, 0, "=SORT('切り'!A1:A2)");
    eq(画(w, 0, 0, 0), '3'); 揃う(w);
  });
  T("名の比べは 大小を区別しない（'DATA' が在る所へ 'Data' は 別の名になる）", () => {
    const w = 開く(html);
    w.eval(`_手順の待ち = { シート:0, 変える:{}, 出すシート:[{ name:'DATA', data:{'0,0':{v:9},'1,0':{v:8}}, 何行:2 }, { name:'Data', data:{'0,0':{v:2},'1,0':{v:1}}, 何行:2 }] }; 手順を書く();`);
    const n = 名の列(w);
    eq(n.length, 3); if (n[2].toLowerCase() === 'data') throw new Error('同じ名（大小違い）が 2枚 ' + JSON.stringify(n));
    打つ(w, 0, 0, 0, "=SORT('" + n[2] + "'!A1:A2)"); eq(画(w, 0, 0, 0), '1'); 揃う(w);
  });
  T('つなげて出す を 2回 押しても 名が重ならず 計算側と 揃う', () => {
    const w = 開く(html);
    w.eval(`_この画面の表たち = function(){ return [{名:'L'},{名:'R'}]; };
      DataModel.つなげる = function(){ return [['k','v'],[1,2]]; };
      つながりたち.length = 0; つながりたち.push({左表:'L',右表:'R',左列:'k',右列:'k'});
      つなげて出す(0); つなげて出す(0);`);
    const n = 名の列(w);
    eq(n.length, 3); if (n[1] === n[2]) throw new Error('同じ名が 2枚 ' + JSON.stringify(n));
    打つ(w, 0, 0, 0, "=SORT('" + n[2] + "'!B1:B2)"); eq(画(w, 0, 0, 0), '2'); 揃う(w);
  });
  T('シートを足す名は 在る名と 重ならない（切り出しの「Sheet3」の後に 足すと Sheet4）', () => {
    const w = 開く(html);
    w.eval(`_手順の待ち = { シート:0, 変える:{}, 出すシート:[{ name:'Sheet3', data:{'0,0':{v:1}}, 何行:1 }] }; 手順を書く(); addSheet();`);
    eq(w.eval('sheets[2].name'), 'Sheet4'); 打つ(w, 2, 0, 0, 9); 打つ(w, 2, 1, 0, 4); 打つ(w, 0, 0, 0, '=SORT(Sheet4!A1:A2)');
    eq(画(w, 0, 0, 0), '4'); 揃う(w);
  });
  T('initFormulaEngine に 大小だけ違う名が来ても hf と exally-formula の _hf は 同じ物・番号は 詰まらない', () => {
    const w = 開く(html);
    w.eval(`initFormulaEngine(['A','a','B']);`);
    eq(w.eval('hf === _hf'), true, 'hf===_hf'); eq(w.eval('hf.countSheets()'), 3, '板の数'); eq(w.eval("hf.getSheetId('B')"), 2, 'B の番号');
  });
  return { pass, fail };
}

if (process.argv.includes('--self-test')) {
  console.log('\n[hf-ita-soroe --self-test] ★直しを 1つずつ 壊すと 赤になるか★');
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
console.log('\n[hf-ita-soroe] ★計算側の板と 画面のシートを 揃える★');
const r = 回す(元のhtml, false);
console.log('\n' + r.pass + ' ok / ' + r.fail + ' NG');
process.exit(r.fail ? 1 : 0);
