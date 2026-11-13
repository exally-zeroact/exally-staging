/* ana9-oi.test.mjs — ★窓（覆い）を 開いている間・鍵の 間は 板へ 何も 入らない★（2026-11-14・穴9）
 *  ★なぜ★ 窓（position:fixed;inset:0 の 覆い）も 鍵（body.exally-locked）も 指・マウスを 止めるだけで、
 *    focus は 隠れた 入力欄・body・リボンの ボタンに 残る＝Ctrl+V・Ctrl+D・Ctrl+Z・打った字・貼り付け・落とし物・
 *    隠れた ボタンの Enter が 後ろの 板を 書き換えた。実Excel の 窓は 閉じるまで 板へ 戻れない。
 *  ★確定の 動き（入力欄の blur）は 今のまま★＝この試験も それを 見る。
 *  使い方: node tests/ana9-oi.test.mjs ／ --self-test＝守りを 1本ずつ 抜いて 回し、どれも 赤に なるか
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(__dirname, '..');

/* ★守り＝[ファイル, 元の字, 代わり]。--self-test=<名> で 1本だけ 抜く★ */
const 守り = {
  '鍵：後の受け手を止める': ['js/auth.js', '    e.stopImmediatePropagation();\n', ''],
  '鍵：変換は止めない': ['js/auth.js', "  var 鍵で止める型 = ['keydown',", "  var 鍵で止める型 = ['compositionend', 'keydown',"],
  '鍵：押しも止める': ['js/auth.js', "'click', 'dblclick'", "'dblclick'"],
  '鍵：ログインの窓の栓': ['js/auth.js', '    if (!止める) { ログインの窓に栓(lo); return; }', '    if (!止める) { return; }'],
  '鍵：リボンのキー': ['book.html', "            if (document.body.classList.contains('exally-locked')) return;   /*", '            /*'],
  'L4：自前の印': ['book.html', '      if(今 === _鍵の印) return;\n', ''],
  'L4：控えと比べる': ['book.html', '(_帯の控え === null || fi_.value !== _帯の控え) && ', ''],
  'L4：セルの欄を捨てる': ['book.html', "      if(ci_ && !editingCell && ci_.value) ci_.value = '';\n", ''],
  'W：onKD': ['book.html', '  if(_覆いが開いている()){\n', '  if(false){\n'],
  'W：doUndo': ['book.html', '  if(_覆いが開いている()) return;   /* ★窓（覆い）が 開いている間は 板へ 書かない★（2026-11-14・穴9 層W） */\n  if(!undoStack.length) return;', '  if(!undoStack.length) return;'],
  'W：doRedo': ['book.html', '  if(_覆いが開いている()) return;   /* ★窓（覆い）が 開いている間は 板へ 書かない★（2026-11-14・穴9 層W） */\n  if(!redoStack.length) return;', '  if(!redoStack.length) return;'],
  'W：ci input': ["book.html", "  ci.addEventListener('input',function(){\n    if(_覆いが開いている()) return;", "  ci.addEventListener('input',function(){\n"],
  'W：ci paste': ["book.html", "  ci.addEventListener('paste', function(e){\n    if(_覆いが開いている()) return;", "  ci.addEventListener('paste', function(e){\n"],
  'W：ci keydown': ["book.html", "  ci.addEventListener('keydown',function(e){\n    if(_覆いが開いている()) return;", "  ci.addEventListener('keydown',function(e){\n"],
  'W：fi keydown': ["book.html", "  fi.addEventListener('keydown',function(e){\n    if(_覆いが開いている()) return;", "  fi.addEventListener('keydown',function(e){\n"],
  'W：fi input': ["book.html", "  fi.addEventListener('input',function(){\n    if(_覆いが開いている()) return;", "  fi.addEventListener('input',function(){\n"],
  'W：リボンのキー': ['book.html', "            if (typeof _リボンキーを窓が止める === 'function' && _リボンキーを窓が止める()) return;", ''],
  'W：開いた時の focus': ['book.html', '      _窓へ戻す(上);\n    });\n', '    });\n'],
  'W：抜けたら戻す': ['book.html', '      _窓へ戻す(上);\n    }, true);', '    }, true);'],
  'W：drop': ['book.html', '    if (f && _覆いが開いている()) {', '    if (false) {'],
};
if (process.argv.includes('--self-test')) {
  const { spawnSync } = await import('node:child_process');
  let 抜けない = 0;
  for (const 名 of Object.keys(守り)) {
    const r = spawnSync(process.execPath, [fileURLToPath(import.meta.url), '--self-test=' + 名], { encoding: 'utf8' });
    const 赤 = (r.stdout.match(/^\s+NG /mg) || []).length;
    console.log((r.status === 1 && 赤 > 0 ? '  ok   ' : '  NG   ') + '守り「' + 名 + '」を 抜く→赤 ' + 赤 + (r.status === 2 ? '（抜く字が 見つからない）' : ''));
    if (!(r.status === 1 && 赤 > 0)) 抜けない++;
  }
  console.log(抜けない ? '\n★守りを 抜いても 赤に ならない 物が ' + 抜けない + '本＝試験が 空振り★' : '\n守り ' + Object.keys(守り).length + '本 全部 赤');
  process.exit(抜けない ? 1 : 0);
}
const 抜く名 = (process.argv.find((a) => a.startsWith('--self-test=')) || '').slice('--self-test='.length);
const 字を変える = (file, code) => {
  code = code.replace(/\r\n/g, '\n');
  if (!抜く名) return code;
  const [先, 元, 代わり] = 守り[抜く名] || [];
  if (!先) { console.log('★守りの 名が 無い: ' + 抜く名); process.exit(2); }
  if (先 !== file) return code;
  if (code.split(元).length !== 2) { console.log('★抜く字が 1か所に 当たらない: ' + 抜く名); process.exit(2); }
  return code.replace(元, () => 代わり);
};

let pass = 0, fail = 0;
const T = async (n, fn) => { try { await fn(); pass++; console.log('  ok   ' + n); } catch (e) { fail++; console.log('  NG   ' + n + '\n       ' + (e && e.message)); } };
const ok = (c, m) => { if (!c) throw new Error(m || 'expected truthy'); };
const eq = (a, b, m) => { if (a !== b) throw new Error((m ? m + ': ' : '') + '期待=' + JSON.stringify(b) + ' 実際=' + JSON.stringify(a)); };

let JSDOM;
try { ({ JSDOM } = await import('jsdom')); }
catch { console.log('★jsdomが入っていません（SKIPを緑と呼ばない）。npm ci してください。'); process.exit(1); }

const html = 字を変える('book.html', fs.readFileSync(path.join(ROOT, 'book.html'), 'utf8'));
const CANVAS_STUB = `
(function(){
  var noop=function(){};
  var ctx=new Proxy({}, { get:function(t,k){
    if(k==='measureText') return function(){ return {width:40}; };
    if(k==='canvas') return {width:800,height:600};
    if(k==='getImageData') return function(){ return {data:[]}; };
    if(k==='createLinearGradient'||k==='createPattern') return function(){ return {addColorStop:noop}; };
    return noop;
  }});
  HTMLCanvasElement.prototype.getContext=function(){ return ctx; };
})();`;
/* ★受け手の 型を 集める★（試験 d）＝読み込む 前に addEventListener を 包む */
const 受け手 = [];
const dom = new JSDOM(html.replace(/<script[\s\S]*?<\/script>/g, ''), {
  runScripts: 'dangerously', url: 'http://localhost/', pretendToBeVisual: true,
  beforeParse(w) {
    w.fetch = () => Promise.reject(new Error('no net'));
    w.scrollTo = () => {}; w.alert = () => {};
    w.matchMedia = w.matchMedia || (() => ({ matches: false, addListener() {}, removeListener() {}, addEventListener() {}, removeEventListener() {} }));
    w.requestAnimationFrame = (cb) => setTimeout(() => cb(Date.now()), 0);
    w.cancelAnimationFrame = () => {};
    w.eval(CANVAS_STUB);
    const 元 = w.EventTarget.prototype.addEventListener;
    w.EventTarget.prototype.addEventListener = function (t, f, o) { 受け手.push({ 的: this, 型: t }); return 元.call(this, t, f, o); };
  },
});
const win = dom.window, doc = win.document;
// jsdom の document.hasFocus() は blur の中で偽を返す。本物は 同じ ページの 中で focus が 動く時 真＝そう揃える（裏の 試験だけ 偽に する）
doc.hasFocus = () => true;
const inject = (code) => { const el = doc.createElement('script'); el.textContent = code; doc.body.appendChild(el); };
for (const src of [...html.matchAll(/<script src="([^"]+)"><\/script>/g)].map((m) => m[1].split('?')[0]).filter((s) => !/^https?:/.test(s))) {
  const p = path.join(ROOT, src);
  if (fs.existsSync(p)) { try { inject(字を変える(src, fs.readFileSync(p, 'utf8'))); } catch (e) { /* 続ける */ } }
}
for (const m of html.matchAll(/<script(?![^>]*\ssrc=)[^>]*>([\s\S]*?)<\/script>/g)) { try { inject(m[1]); } catch (e) { /* 続ける */ } }
try { doc.dispatchEvent(new win.Event('DOMContentLoaded', { bubbles: true })); } catch (e) { /* 続ける */ }
try { win.dispatchEvent(new win.Event('load')); } catch (e) { /* 続ける */ }

const 待つ = (ms = 0) => new Promise((r) => setTimeout(r, ms));
const $ = (id) => doc.getElementById(id);
const ci = $('cell-input'), fi = $('formula-input');
const 値 = (k) => { const x = win.sheets[win.activeSheet].data[k]; return x ? (x.f || x.v) : undefined; };
const 新しい本 = () => { win.eval("if(editingCell) cancelEdit(); sheets.length=1; sheets[0]={name:'Sheet1',data:{},colW:{},rowH:{}}; activeSheet=0; undoStack.length=0; redoStack.length=0; copyBuffer=null;"); };
const 選ぶ = (r1, c1, r2 = r1, c2 = c1) => win.eval(`selR1=${r1};selC1=${c1};selR2=${r2};selC2=${c2};`);
const 打って確定 = (r, c, 字) => { 選ぶ(r, c); win.startEdit(r, c); ci.value = 字; win.commitEdit(); };
const キー = (的, key, o = {}) => { const e = new win.KeyboardEvent('keydown', { key, bubbles: true, cancelable: true, ...o }); 的.dispatchEvent(e); return e; };
const 事 = (的, 型, 足す = {}) => { const e = new win.Event(型, { bubbles: true, cancelable: true }); for (const [k, v] of Object.entries(足す)) Object.defineProperty(e, k, { value: v }); 的.dispatchEvent(e); return e; };
const 貼る = (的, 字) => 事(的, 'paste', { clipboardData: { getData: (t) => (/text/.test(t) ? 字 : '') } });
const 落とす = (的) => 事(的, 'drop', { dataTransfer: { files: [{ name: 'b.xlsx', size: 10 }] } });
const 開く = async (id) => { const el = $(id); if (id === 'fmtModal') el.classList.add('show'); else el.style.display = 'flex'; await 待つ(); return el; };
const 閉じる = async (id) => { const el = $(id); if (id === 'fmtModal') el.classList.remove('show'); else el.style.display = 'none'; await 待つ(); };
const 鍵 = async (掛ける) => { doc.body.classList[掛ける ? 'add' : 'remove']('exally-locked'); await 待つ(); };
const 中か = (id) => $(id).contains(doc.activeElement);
let 開いた本 = 0;
win.openBookFile = () => { 開いた本++; };

console.log('\n[ana9-oi] 窓を 開いている間・鍵の 間は 板へ 何も 入らない' + (抜く名 ? '（--self-test=' + 抜く名 + '＝赤に なるはず）' : ''));
if (doc.body.classList.contains('exally-locked')) await 鍵(false);
/* ログインの 窓（js/exally-login.js が 倉庫の 有る 時だけ 足す）＝試験では 自分で 足す */
if (!$('loginOv')) { const lo = doc.createElement('div'); lo.id = 'loginOv'; lo.innerHTML = '<input id="loginPass"><button id="btnLogin">ログイン</button>'; doc.body.appendChild(lo); }

/* ══ 層W（窓） ══ */
新しい本(); 打って確定(0, 0, 'a'); 選ぶ(0, 0);
ci.focus();
await 開く('findOverlay');
await T('★窓を 開くと focus が 窓の 最初の 欄へ 入る（後ろの セルの 欄に 残らない）★', () => { ok(中か('findOverlay'), 'activeElement=' + (doc.activeElement && (doc.activeElement.id || doc.activeElement.tagName))); });
await T('★窓の 間に セルの 欄へ focus を 移しても 窓へ 戻る★', () => { ci.focus(); ok(中か('findOverlay'), 'activeElement=' + (doc.activeElement && doc.activeElement.id)); });
const 窓の欄 = () => (中か('findOverlay') ? doc.activeElement : $('findOverlay').querySelector('input'));
await T('★窓の 欄で Ctrl+V＝後ろの 板へ 貼らない★', () => { 選ぶ(0, 0); win.ctxCopy(); 選ぶ(0, 1); キー(窓の欄(), 'v', { ctrlKey: true }); eq(値('0,1'), undefined, 'B1'); });
await T('★窓の 欄で Ctrl+D＝下へ 写さない★', () => { 選ぶ(0, 0, 1, 0); キー(窓の欄(), 'd', { ctrlKey: true }); eq(値('1,0'), undefined, 'A2'); });
await T('★窓の 欄で Ctrl+Z＝後ろの 板を 戻さない★', () => { const n = win.undoStack.length; ok(n > 0, '戻す物が 在る'); キー(窓の欄(), 'z', { ctrlKey: true }); eq(win.undoStack.length, n, 'undo の 数'); eq(値('0,0'), 'a', 'A1'); });
await T('★↩ を 直に 呼んでも 窓の 間は 戻さない★', () => { const n = win.undoStack.length; win.doUndo(); eq(win.undoStack.length, n); eq(値('0,0'), 'a', 'A1'); });
await T('★↪ を 直に 呼んでも 窓の 間は やり直さない★', async () => {
  await 閉じる('findOverlay'); win.doUndo(); const m = win.redoStack.length; ok(m > 0, 'やり直す物が 在る'); const 前 = 値('0,0');
  await 開く('findOverlay'); win.doRedo(); eq(win.redoStack.length, m, 'redo の 数'); eq(値('0,0'), 前, 'A1');
  await 閉じる('findOverlay'); win.doRedo(); await 開く('findOverlay');
});
await T('窓の 欄の Ctrl+C は ブラウザの まま（既定を 止めない）／Ctrl+S は 止める（ブラウザの 保存へ 落とさない）', () => {
  eq(キー(窓の欄(), 'c', { ctrlKey: true }).defaultPrevented, false, 'Ctrl+C');
  eq(キー(窓の欄(), 's', { ctrlKey: true }).defaultPrevented, true, 'Ctrl+S');
});
await T('★窓の 間に 隠れた セルの 欄へ 字・Enter・貼り付けが 来ても 板を 書かない★', () => {
  選ぶ(2, 0);
  ci.value = 'x'; 事(ci, 'input'); eq(win.editingCell, null, 'editingCell');
  ci.value = ''; キー(ci, 'Enter'); eq(win.editingCell, null, 'Enter で 打ち始めない');
  貼る(ci, '貼った'); eq(値('2,0'), undefined, 'A3');
});
await T('★窓の 間に 帯へ F4・字が 来ても 動かない★', () => {
  fi.value = '=A1'; キー(fi, 'F4'); eq(fi.value, '=A1', 'F4');
  fi.value = '=打ちかけ'; キー(fi, 'Escape'); eq(fi.value, '=打ちかけ', 'Escape で 帯を 戻さない');
  win.eval('_currentRefs=null;'); fi.value = '=B2'; 事(fi, 'input'); eq(win.eval('_currentRefs'), null, '帯の input');
});
await T('★窓の 間に リボンの キーチップを 通さない★', () => {
  if (!win.Ribbon) throw new Error('リボンが 読めていない（試験が 空振り）');
  let 呼んだ = 0; const 元 = win.Ribbon.キーを受ける; win.Ribbon.キーを受ける = () => { 呼んだ++; return false; };
  try { キー(窓の欄(), 'Alt', { altKey: true }); } finally { win.Ribbon.キーを受ける = 元; }
  eq(呼んだ, 0, 'キーを受ける');
});
await T('★窓の 間に 本を 落としても 差し替えない★', () => { 開いた本 = 0; 落とす(doc.body); eq(開いた本, 0); });
await 閉じる('findOverlay');
await T('書式の 窓（#fmtModal）も 同じ＝落としても 差し替えない・Ctrl+V しない', async () => {
  await 開く('fmtModal'); try {
    開いた本 = 0; 落とす(doc.body); eq(開いた本, 0, '落とす');
    選ぶ(0, 0); win.ctxCopy(); 選ぶ(3, 3); キー(doc.body, 'v', { ctrlKey: true }); eq(値('3,3'), undefined, 'D4');
  } finally { await 閉じる('fmtModal'); }
});
await T('入れ子＝fit の 上に 関数の 窓を 開いて 閉じると、fit が 開いたままなら 抜けた focus は fit へ 戻る', async () => {
  await 開く('fitOverlay'); await 開く('funcOverlay');
  const f = $('funcOverlay').querySelector('input,button'); if (f) f.focus();
  await 閉じる('funcOverlay');
  ci.focus(); ok(中か('fitOverlay'), 'activeElement=' + (doc.activeElement && doc.activeElement.id));
  await 閉じる('fitOverlay');
});
await T('名簿の 外の 窓（ウォッチ）は 止めない＝開いたまま セルへ 打てる', async () => {
  $('watchOverlay').style.display = 'block'; await 待つ();
  try { ci.focus(); eq(doc.activeElement, ci, 'focus'); 選ぶ(4, 0); ci.value = 'w'; 事(ci, 'input'); ok(win.editingCell, '打ち始める'); win.commitEdit(); eq(値('4,0'), 'w'); }
  finally { $('watchOverlay').style.display = 'none'; }
});
await T('窓を 閉じれば 前の 通り 打てる・貼れる・戻せる', () => {
  選ぶ(5, 0); ci.focus(); ci.value = 'y'; 事(ci, 'input'); ok(win.editingCell, '打ち始める'); win.commitEdit(); eq(値('5,0'), 'y');
  選ぶ(6, 0); 貼る(ci, 'p'); eq(値('6,0'), 'p', '貼る');
  const n = win.undoStack.length; win.doUndo(); eq(win.undoStack.length, n - 1, '戻す');
});

/* ══ 層L（鍵）＝js/auth.js の 入口1か所 ══ */
await 待つ(150);   /* 前の 試験の blur が 残した 100ms の 確定（ci の blur の setTimeout）を 先に 流す */
新しい本(); 打って確定(0, 0, 'a'); 選ぶ(0, 0);
await 待つ(150);
await T('★鍵の 間＝打っていた セルの 欄で Escape・Enter が 来ても 打ちかけを 捨てない・確定しない★', async () => {
  選ぶ(1, 0); win.startEdit(1, 0); ci.value = 'ab';
  await 鍵(true);
  try { キー(ci, 'Escape'); ok(win.editingCell, 'Escape で 捨てない'); eq(ci.value, 'ab'); キー(ci, 'Enter'); eq(値('1,0'), undefined, 'Enter で 確定しない'); }
  finally { await 鍵(false); win.cancelEdit(); }
});
await T('★鍵の 間＝帯の F4・Escape が 効かない★', async () => {
  fi.value = '=A1'; await 鍵(true);
  try { fi.focus(); キー(fi, 'F4'); eq(fi.value, '=A1', 'F4'); fi.value = '=打ちかけ'; キー(fi, 'Escape'); eq(fi.value, '=打ちかけ', 'Escape'); } finally { await 鍵(false); fi.blur(); await 待つ(150); win.updateBar(); }
});
await T('★鍵の 間＝隠れた セルの 欄へ 貼っても 板を 書かない・字で 打ち始めない★', async () => {
  選ぶ(2, 0); await 鍵(true);
  try { 貼る(ci, '前の人の字'); eq(値('2,0'), undefined, 'A3'); ci.value = 'q'; 事(ci, 'input'); eq(win.editingCell, null); }
  finally { await 鍵(false); }
});
await T('★鍵の 間＝隠れた ↩ ボタンの 押し（Enter・Space の 既定の click）で 戻さない★', async () => {
  const n = win.undoStack.length; ok(n > 0, '戻す物が 在る');
  await 鍵(true);
  try { doc.querySelector('.fmt-undo-redo').click(); eq(win.undoStack.length, n); eq(値('0,0'), 'a', 'A1'); } finally { await 鍵(false); }
});
await T('★鍵の 間＝body・document の Ctrl+V・Ctrl+Z も 板へ 届かない★', async () => {
  選ぶ(0, 0); win.ctxCopy(); 選ぶ(3, 0); const n = win.undoStack.length;
  await 鍵(true);
  try { キー(doc.body, 'v', { ctrlKey: true }); キー(doc.body, 'z', { ctrlKey: true }); eq(値('3,0'), undefined, 'A4'); eq(win.undoStack.length, n); }
  finally { await 鍵(false); }
});
await T('★鍵の 間＝落としても 本を 差し替えない★', async () => { 開いた本 = 0; await 鍵(true); try { 落とす(doc.body); eq(開いた本, 0); } finally { await 鍵(false); } });
await T('鍵の 間＝字・貼り付けの 既定は 止める／キーの 既定は 止めない（F5・Tab・Ctrl+R は 効く）', async () => {
  await 鍵(true);
  try {
    eq(事(ci, 'beforeinput').defaultPrevented, true, 'beforeinput');
    eq(貼る(fi, 'x').defaultPrevented, true, 'paste');
    eq(キー(doc.body, 'F5').defaultPrevented, false, 'F5');
    eq(キー(ci, 'Tab').defaultPrevented, false, 'Tab');
  } finally { await 鍵(false); }
});
await T('★鍵の 間でも ログインの 窓の 中は そのまま（受け手が 走る・既定を 止めない）★', async () => {
  let 走った = 0; const pw = $('loginPass'); const f = () => { 走った++; }; pw.addEventListener('keydown', f); pw.addEventListener('paste', f);
  await 鍵(true);
  try {
    eq(キー(pw, 'Enter').defaultPrevented, false, 'Enter'); eq(貼る(pw, 'pw').defaultPrevented, false, 'paste');
    eq(走った, 2, '受け手'); let 押した = 0; $('btnLogin').addEventListener('click', () => 押した++); $('btnLogin').click(); eq(押した, 1, 'ログインの ボタン');
  } finally { await 鍵(false); pw.removeEventListener('keydown', f); pw.removeEventListener('paste', f); }
});
await T('★鍵の 間に ログインの 欄で 打つ Ctrl+V・Ctrl+A・Enter・Alt は 板へ 上がらない（onKD・リボンの キーチップ）★', async () => {
  新しい本(); 打って確定(0, 0, 'a'); 選ぶ(0, 0); win.ctxCopy(); 選ぶ(2, 2); const n = win.undoStack.length; const pw = $('loginPass');
  let 呼んだ = 0; const 元 = win.Ribbon && win.Ribbon.キーを受ける; if (win.Ribbon) win.Ribbon.キーを受ける = () => { 呼んだ++; return false; };
  await 鍵(true);
  try {
    キー(pw, 'v', { ctrlKey: true }); キー(pw, 'z', { ctrlKey: true }); キー(pw, 'a', { ctrlKey: true }); キー(pw, 'Enter'); キー(pw, 'Alt', { altKey: true });
    eq(値('2,2'), undefined, 'C3 に 貼らない'); eq(win.undoStack.length, n, 'undo'); eq(win.eval('selR1+","+selC1+","+selR2+","+selC2'), '2,2,2,2', '選び'); eq(呼んだ, 0, 'キーチップ');
  } finally { await 鍵(false); if (win.Ribbon) win.Ribbon.キーを受ける = 元; }
});
await T('★鍵の 間に 変換が 終わっても 変換の 旗が 残らない＝解いた後 抜ければ 今まで通り 確定する★', async () => {
  選ぶ(4, 0); win.startEdit(4, 0); ci.value = 'あ'; 事(ci, 'compositionstart');
  await 鍵(true); 事(ci, 'compositionend'); await 鍵(false);
  ci.dispatchEvent(new win.FocusEvent('blur')); await 待つ(150);
  eq(値('4,0'), 'あ', 'A5');
});

/* ══ L4（解けた時の 後始末） ══ */
新しい本(); 打って確定(0, 0, 'a'); 選ぶ(0, 0); win.updateBar();
const クラスを付け直す = async () => { doc.body.setAttribute('class', doc.body.getAttribute('class') || ''); await 待つ(); };
await T('★鍵の 間に 帯へ 入った 字（IME の 残り）は 解くと 捨てる（付け直しで 鳴っても 控えを 取り直さない）★', async () => {
  await 鍵(true); fi.value = 'x'; await クラスを付け直す(); await 鍵(false);
  eq(fi.value, 'a', '帯');
});
await T('★鍵の 無い 時に class を 外し直しても 帯の 打ちかけを 捨てない★', async () => { fi.value = 'abc'; await クラスを付け直す(); eq(fi.value, 'abc'); win.updateBar(); });
await T('★鍵の 前の 帯の 打ちかけは 解いても 残る（タブが 裏の 時・自前キーボードの 時）★', async () => {
  doc.hasFocus = () => false; fi.value = 'abc'; await 鍵(true); doc.hasFocus = () => true; await 鍵(false); eq(fi.value, 'abc', '裏');
  win.updateBar(); $('formula-kb').style.display = 'block'; fi.value = 'kbd'; await 鍵(true); await 鍵(false); eq(fi.value, 'kbd', '自前キーボード');
  $('formula-kb').style.display = 'none'; win.updateBar(); eq(値('0,0'), 'a', 'A1 は 元のまま');
});
await T('★打っていない セルの 欄に 鍵の 間の 字が 残れば 解くと 捨てる・打っていた 欄は 触らない★', async () => {
  await 鍵(true); ci.value = 'か'; await 鍵(false); eq(ci.value, '', '打っていない'); eq(win.editingCell, null);
  選ぶ(1, 0); win.startEdit(1, 0); ci.value = '続き'; await 鍵(true); await 鍵(false); eq(ci.value, '続き', '打っていた'); win.cancelEdit();
});

/* ══ d：受け手の 型＝鍵で 止める 型か、通す 型か の どちらかに 在る ══ */
const 通す型 = new Set(['focus', 'blur', 'focusin', 'focusout', 'compositionstart', 'compositionupdate', 'compositionend', 'scroll', 'load',
  'DOMContentLoaded', 'resize', 'mousemove', 'mouseup', 'mouseover', 'mouseout', 'mouseenter', 'mouseleave', 'pointermove', 'pointerup',
  'pointercancel', 'pointerleave', 'pointerenter', 'pointerover', 'pointerout', 'touchstart', 'touchmove', 'touchend', 'touchcancel', 'wheel', 'dragleave', 'dragend', 'dragstart',
  'beforeunload', 'unload', 'pagehide', 'pageshow', 'error', 'unhandledrejection', 'message', 'visibilitychange', 'selectionchange',
  'orientationchange', 'online', 'offline', 'storage', 'popstate', 'hashchange', 'select', 'gesturestart', 'gesturechange', 'gestureend',
  'copy', 'afterprint', 'beforeprint', 'animationend', 'transitionend', 'invalid', 'reset', 'toggle', 'contextlost', 'contextrestored']);
await T('★板へ 届く 受け手の 型は 全部「鍵で 止める」か「通す」の 表に 在る★（表に 無い 型が 増えたら 人が 決める）', () => {
  const 止める = new Set(win.__鍵で止める型 || []);
  ok(止める.size > 0, '鍵の 表が 読めない');
  const 見る = new Set([ci, fi, doc, win, doc.body, ...doc.querySelectorAll('#ribbon button, .fmt-undo-redo')]);
  const 型 = new Set(受け手.filter((x) => 見る.has(x.的)).map((x) => x.型));
  for (const el of 見る) { if (el && el.attributes) for (const a of el.attributes) if (/^on/.test(a.name)) 型.add(a.name.slice(2)); }
  for (const el of [ci, fi]) for (const k of Object.keys(el)) if (/^on/.test(k) && typeof el[k] === 'function') 型.add(k.slice(2));
  const 無い = [...型].filter((t) => !止める.has(t) && !通す型.has(t));
  console.log('       （集めた 型 ' + 型.size + '：' + [...型].sort().join(' ') + '）');
  eq(無い.join(' '), '', '表に 無い 型');
});

console.log('\n' + (fail ? '★NG ' + fail + '本★ ' : '') + 'ok ' + pass + '本');
process.exit(fail ? 1 : 0);
