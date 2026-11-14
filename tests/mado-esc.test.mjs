/* mado-esc.test.mjs — ★Esc で 一番上の 窓を 閉じる／窓を 全部 閉じたら 的を セルの 欄へ 戻す★（2026-11-15）
 *  ★なぜ★ 実Excel の 窓は Esc で 閉じ、閉じると セルへ そのまま 打てる。今までは Esc で 閉じず、
 *    閉じるボタンで 閉じた 後は 的が body に 落ちて 打った 字が 消えていた。
 *  使い方: node tests/mado-esc.test.mjs ／ --self-test＝守りを 1本ずつ 抜いて 回し、どれも 赤に なるか
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(__dirname, '..');

/* ★守り＝[ファイル, 元の字, 代わり]。--self-test=<名> で 1本だけ 抜く★ */
const 守り = {
  'Esc：変換中 isComposing': ['book.html', "if (e.key !== 'Escape' || e.isComposing || e.keyCode === 229", "if (e.key !== 'Escape' || e.keyCode === 229"],
  'Esc：変換中 229': ['book.html', "if (e.key !== 'Escape' || e.isComposing || e.keyCode === 229", "if (e.key !== 'Escape' || e.isComposing"],
  'Esc：SELECT は1回見送る': ['book.html', "if (t && t.tagName === 'SELECT' && _一覧が開いたかも === t) { _一覧が開いたかも = null; return; }", 'if (false) {}'],
  'Esc：SELECT の印を倒す': ['book.html', "window.addEventListener(型, function(e){ if (e.target === _一覧が開いたかも) _一覧が開いたかも = null; }, true);", 'void 型;'],
  'Esc：SELECT のキーで印': ['book.html', "_一覧が開いたかも = t;\n    }, true);", 'void 0;\n    }, true);'],
  'Esc：SELECT の押しで印': ['book.html', "      _一覧が開いたかも = s;\n", ''],
  'Esc：窓の外の欄': ['book.html', "      if (!(t === document.body || t === document.documentElement || 上.contains(t))) return;\n", ''],
  'Esc：ログインと一覧': ['book.html', "      if (_見えているか('loginOv') || _見えているか('fnPalette')) return;\n", ''],
  'Esc：既定を止める': ['book.html', "{ e.preventDefault(); e.stopPropagation(); }\n    }, true);", "{ e.stopPropagation(); }\n    }, true);"],
  'Esc：表の1枚': ['book.html', "  sortOverlay:'並べ替えの窓を閉じる', ", '  '],
  'Esc：フォームは外す': ['book.html', "  diffOverlay:'#diffCancel',", "  formOverlay:'フォームを閉じる', diffOverlay:'#diffCancel',"],
  '戻す：鍵の間は印を動かさない': ['book.html', "      if (document.body.classList.contains('exally-locked')) return;   /* 鍵の 間は 印を 動かさない */\n", ''],
  '戻す：鍵の間は奪わない': ['book.html', "  if (document.body.classList.contains('exally-locked') || _覆いが開いている()) return;\n  var ci", "  if (_覆いが開いている()) return;\n  var ci"],
  '戻す：まだ開いていれば奪わない': ['book.html', "  if (document.body.classList.contains('exally-locked') || _覆いが開いている()) return;\n  var ci", "  if (document.body.classList.contains('exally-locked')) return;\n  var ci"],
  '戻す：他の欄は奪わない': ['book.html', "  if (奪ってよい) ci.focus();", "  ci.focus();"],
  '戻す：呼ぶ': ['book.html', '      setTimeout(_閉じた後の的を戻す, 0);\n', ''],
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

console.log('\n[mado-esc] Esc で 一番上の 窓を 閉じる・閉じたら 的を セルの 欄へ' + (抜く名 ? '（--self-test=' + 抜く名 + '＝赤に なるはず）' : ''));
if (doc.body.classList.contains('exally-locked')) await 鍵(false);
if (!$('loginOv')) { const lo = doc.createElement('div'); lo.id = 'loginOv'; lo.style.display = 'none'; lo.innerHTML = '<input id="loginPass">'; doc.body.appendChild(lo); }
/* 試験では 倉庫が 無い＝ログインの 窓が 開いたまま（本物は 鍵と 一緒に だけ 開く）＝鍵を 外したので 閉じる */
$('loginOv').classList.remove('open'); $('loginOv').style.display = 'none';
const 見える = (id) => win.getComputedStyle($(id)).display !== 'none';
const 的 = () => doc.activeElement && (doc.activeElement.id || doc.activeElement.tagName);
const 名簿 = win.eval('_覆いの名簿.slice()'), 表 = win.eval('_窓を閉じる表');

await T('★表の 窓＝名簿 − formOverlay（足りない・余る 0）★', () => {
  const 要る = 名簿.filter((id) => id !== 'formOverlay').sort(), 在る = Object.keys(表).sort();
  eq(JSON.stringify(在る), JSON.stringify(要る));
  ok(要る.length === 44, '44枚 実際=' + 要る.length);
});
await T('★表の 閉じる 関数・ボタンが 全部 在る★', () => {
  const 無い = Object.entries(表).filter(([, n]) => (n[0] === '#' ? !$(n.slice(1)) : typeof win[n] !== 'function')).map(([id]) => id);
  eq(無い.join(','), '');
});
/* 差分の 窓の やめるは 開く 関数が 付ける＝試験でも 付ける */
let やめた = 0;
const 差分を開く = async () => { $('diffCancel').onclick = () => { やめた++; $('diffOverlay').style.display = 'none'; }; await 開く('diffOverlay'); };
const 閉じない = [];
for (const id of Object.keys(表)) {
  新しい本(); ci.focus();
  if (id === 'diffOverlay') await 差分を開く(); else await 開く(id);
  const 上 = win.eval('_覆いが開いている()');
  const t = 上 && 上.id === id && $(id).contains(doc.activeElement) ? doc.activeElement : $(id);
  let e;
  try { e = キー(t, 'Escape'); } catch (er) { 閉じない.push(id + '（落ちた ' + er.message + '）'); await 閉じる(id); continue; }
  await 待つ(); await 待つ();
  if (見える(id) || !e.defaultPrevented) { 閉じない.push(id + (上 && 上.id !== id ? '（上=' + 上.id + '）' : '')); await 閉じる(id); continue; }
  if (doc.activeElement !== ci) 閉じない.push(id + '（的=' + 的() + '）');
}
await T('★44枚 全部＝窓の 中の Esc で 閉じて 的が セルの 欄へ 戻る★', () => { eq(閉じない.join(' / '), ''); });
await T('差分の 窓の Esc＝やめる を 押す（つながない）', () => { eq(やめた, 1); });
await T('★フォームの 窓（閉じて 入れる）は Esc で 閉じない＝黙って 書かない★', async () => {
  await 開く('formOverlay'); const e = キー($('formOverlay').querySelector('input,button') || $('formOverlay'), 'Escape'); await 待つ();
  ok(見える('formOverlay'), '閉じた'); eq(e.defaultPrevented, false); await 閉じる('formOverlay');
});
await T('★字を 変換中（isComposing・keyCode 229）は 閉じない★', async () => {
  await 開く('findOverlay'); const 欄 = $('findOverlay').querySelector('input');
  キー(欄, 'Escape', { isComposing: true }); await 待つ(); ok(見える('findOverlay'), 'isComposing で 閉じた');
  const e = new win.KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true }); Object.defineProperty(e, 'keyCode', { value: 229 });
  欄.dispatchEvent(e); await 待つ(); ok(見える('findOverlay'), '229 で 閉じた'); await 閉じる('findOverlay');
});
await T('★SELECT＝一覧を 開いたかも（押した・Alt+↓・F4・空白）の 後は 1回目を 見送り 2回目で 閉じる★', async () => {
  const id = 'sortOverlay'; const 押し方 = [
    ['押した', (s) => 事(s, 'mousedown')], ['Alt+↓', (s) => キー(s, 'ArrowDown', { altKey: true })],
    ['F4', (s) => キー(s, 'F4')], ['空白', (s) => キー(s, ' ')]];
  for (const [名, 開ける] of 押し方) {
    await 開く(id); const s = $(id).querySelector('select'); ok(s, 'select'); s.focus();
    開ける(s); キー(s, 'Escape'); await 待つ(); ok(見える(id), 名 + '→1回目で 閉じた');
    キー(s, 'Escape'); await 待つ(); ok(!見える(id), 名 + '→2回目で 閉じない');
  }
});
await T('★SELECT＝開いていない（focus だけ）・選んだ 後・離れて 戻った 後 は 1回で 閉じる★', async () => {
  const id = 'sortOverlay';
  await 開く(id); let s = $(id).querySelector('select'); s.focus(); キー(s, 'Escape'); await 待つ(); ok(!見える(id), 'focus だけ');
  await 開く(id); s.focus(); 事(s, 'mousedown'); 事(s, 'change'); キー(s, 'Escape'); await 待つ(); ok(!見える(id), '選んだ 後');
  await 開く(id); s.focus(); 事(s, 'mousedown'); 事(s, 'focusout'); s.focus(); キー(s, 'Escape'); await 待つ(); ok(!見える(id), '離れて 戻った 後');
});
await T('窓の 外の 隠れた 欄へ 来た Esc では 閉じない（穴9 の 守り の 内）', async () => {
  await 開く('findOverlay'); キー(fi, 'Escape'); await 待つ(); ok(見える('findOverlay')); await 閉じる('findOverlay');
});
await T('body へ 来た Esc（的が 落ちた 時）は 閉じる', async () => {
  await 開く('findOverlay'); キー(doc.body, 'Escape'); await 待つ(); ok(!見える('findOverlay'));
});
await T('ログインの 窓が 見えている・関数の 一覧が 開いている 時は 閉じない', async () => {
  await 開く('findOverlay'); const 欄 = $('findOverlay').querySelector('input');
  $('loginOv').style.display = 'block'; キー(欄, 'Escape'); await 待つ(); $('loginOv').style.display = 'none'; ok(見える('findOverlay'), 'loginOv');
  $('fnPalette').style.display = 'flex'; キー(欄, 'Escape'); await 待つ(); $('fnPalette').style.display = 'none'; ok(見える('findOverlay'), 'fnPalette');
  await 閉じる('findOverlay');
});
await T('★鍵の 間は Esc で 閉じない★', async () => {
  await 開く('findOverlay'); await 鍵(true); キー(doc.body, 'Escape'); await 待つ(); await 鍵(false);
  ok(見える('findOverlay')); await 閉じる('findOverlay');
});
await T('★入れ子＝fit の 上の 関数の 窓だけ 閉じる・的は fit へ（セルの 欄へ 抜けない）★', async () => {
  await 開く('fitOverlay'); await 開く('funcOverlay');
  const f = $('funcOverlay').querySelector('input,button'); f.focus();
  キー(f, 'Escape'); await 待つ(); await 待つ();
  ok(!見える('funcOverlay'), 'func が 閉じない'); ok(見える('fitOverlay'), 'fit も 閉じた'); ok(中か('fitOverlay'), '的=' + 的());
  キー(doc.activeElement, 'Escape'); await 待つ(); await 待つ(); ok(!見える('fitOverlay'), 'fit'); eq(doc.activeElement, ci, '最後は セルの 欄');
});
await T('★閉じるボタンで 閉じても 的が セルの 欄へ（body に 落ちない）★', async () => {
  await 開く('findOverlay'); win.closeFind(); doc.activeElement.blur(); await 待つ(); await 待つ(); eq(doc.activeElement, ci, '的=' + 的());
});
await T('閉じた 直後に 他の 欄へ 的を 移した 時は 奪わない', async () => {
  await 開く('findOverlay'); win.closeFind(); fi.focus(); await 待つ(); await 待つ(); eq(doc.activeElement, fi, '的=' + 的());
});
await T('★鍵の 間に 閉じても 奪わない／鍵の 間に 開いていた 窓を 解けた 後に 閉じたら 戻る★', async () => {
  await 開く('findOverlay'); await 鍵(true); await 閉じる('findOverlay'); await 待つ(); ok(doc.activeElement !== ci, '鍵の 間に 奪った');
  await 鍵(false); await 開く('findOverlay'); await 鍵(true); await 鍵(false);
  win.closeFind(); doc.activeElement.blur(); await 待つ(); await 待つ(); eq(doc.activeElement, ci, '的=' + 的());
});
/* ci.focus を 数える（戻す 関数が 呼んだかを 見る＝窓の 閉じ込めが 後で 戻しても 分かる） */
const 数える = async (手) => { let n = 0; const 元 = ci.focus; ci.focus = function () { n++; return 元.apply(this, arguments); }; try { await 手(); } finally { ci.focus = 元; } return n; };
await T('★鍵の 間に 窓の 形が 変わっても 印を 倒さない＝解けた 後に 閉じたら 戻る★', async () => {
  await 開く('findOverlay'); await 鍵(true);
  $('findOverlay').style.zIndex = '1001'; await 待つ(); $('findOverlay').style.zIndex = ''; await 待つ();
  await 鍵(false); win.closeFind(); doc.activeElement.blur(); await 待つ(); await 待つ(); eq(doc.activeElement, ci, '的=' + 的());
});
await T('★閉じた 直後（戻す 前）に 鍵が 掛かったら 奪わない★', async () => {
  await 開く('findOverlay'); fi.blur();
  const n = await 数える(async () => { win.closeFind(); await Promise.resolve(); doc.body.classList.add('exally-locked'); await 待つ(); await 待つ(); });
  await 鍵(false); eq(n, 0, 'ci.focus の 数');
});
await T('★閉じた 直後（戻す 前）に 別の 窓が 開いたら 奪わない★', async () => {
  await 開く('findOverlay');
  const n = await 数える(async () => { win.closeFind(); await Promise.resolve(); $('tnOverlay').style.display = 'flex'; await 待つ(); await 待つ(); });
  eq(n, 0, 'ci.focus の 数'); ok(中か('tnOverlay'), '的=' + 的()); await 閉じる('tnOverlay');
});

console.log('\n' + (fail ? '★NG ' + fail + '本★ ' : '') + 'ok ' + pass + '本');
process.exit(fail ? 1 : 0);
