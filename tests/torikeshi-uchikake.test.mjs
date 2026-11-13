/* torikeshi-uchikake.test.mjs — ★打っている途中の ↩／↪ は 歴史を 動かさない★（2026-11-13・穴⑤）
 *  ★なぜ★ ↩ ボタンを 押すと 入力欄の blur は 100ms 後に 確定し、その前に doUndo が 走る＝打ちかけのまま 前の操作を 戻し、
 *    後から 打った字が 確定して やり直しが 消えた。別の板の 控えなら switchSheet の commitEdit が 打ちかけを 先に 確定した。
 *    実Excel は セルを 編集している 間は 元に戻す・やり直しが 押せない（Ctrl+Z も 今は onKD で 弾く）＝ボタンも 同じに する。
 *  使い方: node tests/torikeshi-uchikake.test.mjs ／ --self-test＝直す前の形（打ちかけでも 歴史を 動かす）で 赤になるか
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(__dirname, '..');
const require_ = createRequire(path.join(ROOT, 'package.json'));
let pass = 0, fail = 0;
const T = async (n, fn) => { try { await fn(); pass++; console.log('  ok   ' + n); } catch (e) { fail++; console.log('  NG   ' + n + '\n       ' + (e && e.message)); } };
const ok = (c, m) => { if (!c) throw new Error(m || 'expected truthy'); };
const eq = (a, b, m) => { if (a !== b) throw new Error((m ? m + ': ' : '') + '期待=' + JSON.stringify(b) + ' 実際=' + JSON.stringify(a)); };

let JSDOM;
try { ({ JSDOM } = await import('jsdom')); }
catch { console.log('★jsdomが入っていません（SKIPを緑と呼ばない）。npm ci してください。'); process.exit(1); }

const html = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf8');
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
const dom = new JSDOM(html.replace(/<script[\s\S]*?<\/script>/g, ''), {
  runScripts: 'dangerously', url: 'http://localhost/', pretendToBeVisual: true,
  beforeParse(w) {
    w.fetch = () => Promise.reject(new Error('no net'));
    w.scrollTo = () => {}; w.alert = () => {};
    w.matchMedia = w.matchMedia || (() => ({ matches: false, addListener() {}, removeListener() {}, addEventListener() {}, removeEventListener() {} }));
    w.requestAnimationFrame = (cb) => setTimeout(() => cb(Date.now()), 0);
    w.cancelAnimationFrame = () => {};
    w.eval(CANVAS_STUB);
  },
});
const win = dom.window, doc = win.document;
const inject = (code) => { const el = doc.createElement('script'); el.textContent = code; doc.body.appendChild(el); };
for (const src of [...html.matchAll(/<script src="([^"]+)"><\/script>/g)].map((m) => m[1].split('?')[0]).filter((s) => !/^https?:/.test(s))) {
  const p = path.join(ROOT, src);
  if (fs.existsSync(p)) { try { inject(fs.readFileSync(p, 'utf8')); } catch (e) { /* 続ける */ } }
}
for (const m of html.matchAll(/<script(?![^>]*\ssrc=)[^>]*>([\s\S]*?)<\/script>/g)) { try { inject(m[1]); } catch (e) { /* 続ける */ } }
try { doc.dispatchEvent(new win.Event('DOMContentLoaded', { bubbles: true })); } catch (e) { /* 続ける */ }
try { win.dispatchEvent(new win.Event('load')); } catch (e) { /* 続ける */ }
doc.body.classList.remove('exally-locked');   /* ★ログインが 済んだ 形で 試す★（試験は 倉庫を 読まない＝auth.js が 鍵を 掛ける。鍵の 間は 門が キー・押しを 止める＝穴9） */
/* ★保存の 道具は 本番でも 初めての 保存で 後から 読む（_ensureXlsx と 同じ 順）★ */

const SELF = process.argv.includes('--self-test');
if (SELF) {
  /* ★直す前の 動きを 真似る★＝打っている途中でも ↩／↪ が 歴史を 動かす */
  win.eval("(function(){ var u=doUndo, r=doRedo;"
    + " function 外して(f){ return function(){ var k=editingCell, m=_editMode; editingCell=null; _editMode=false; try{ return f(); } finally{ if(editingCell===null) editingCell=k; if(_editMode===false) _editMode=m; } }; }"
    + " window.doUndo=外して(u); window.doRedo=外して(r); })()");
  /* ボタンを 押しても 入力欄から 抜けない 印（onmousedown）も 外す */
  for (const b of doc.querySelectorAll('.fmt-undo-redo')) b.removeAttribute('onmousedown');
}
const 板 = (i) => win.sheets[i];
const 値 = (i, k) => { const x = 板(i).data[k]; return x ? (x.f || x.v) : undefined; };
const 新しい本 = () => { win.eval("sheets.length=1; sheets[0]={name:'Sheet1',data:{},colW:{},rowH:{}}; activeSheet=0; undoStack.length=0; redoStack.length=0;"); if (typeof win.rebuildSheetTabs === 'function') win.rebuildSheetTabs(); };
/* ★本物の 押し方の 順★＝↩ を 押すと 先に 入力欄から 抜け（blur は 100ms 後に 確定）、その前に onclick（doUndo）が 走る */
const 打ちかけ = (r, c, 字) => { win.startEdit(r, c); doc.getElementById('cell-input').value = 字; };
const 抜けて確定 = () => { if (win.editingCell) win.commitEdit(); };
console.log('\n[torikeshi-uchikake] 打っている途中の ↩／↪ は 歴史を 動かさない（実Excel は 編集中は 元に戻すが 押せない）' + (SELF ? '（--self-test＝赤に なるはず）' : ''));

新しい本();
win.setCell(0, 0, 'a');
打ちかけ(0, 1, '打ちかけ');
await T('★同じ板で 打ちかけのまま ↩＝前の操作を 戻さない★', () => { win.doUndo(); 抜けて確定(); eq(値(0, '0,0'), 'a', 'A1'); eq(値(0, '0,1'), '打ちかけ', 'B1'); });
await T('確定した後の ↩ は 打った字から 戻す', () => { win.doUndo(); eq(値(0, '0,1'), undefined, 'B1'); eq(値(0, '0,0'), 'a', 'A1'); });

新しい本();
win.addSheet(); const 板2 = win.activeSheet;
win.switchSheet(0); win.setCell(0, 0, '板1');
win.switchSheet(板2);
打ちかけ(0, 0, '板2の打ちかけ');
await T('★別の板で 打ちかけのまま ↩＝板を 移らない・板1 を 戻さない★', () => { win.doUndo(); eq(win.activeSheet, 板2, '開いている板'); 抜けて確定(); eq(値(0, '0,0'), '板1', '板1 A1'); eq(値(板2, '0,0'), '板2の打ちかけ', '板2 A1'); });
await T('打ちかけを 確定した後の ↩ は 板2 の 字を 戻す', () => { win.doUndo(); eq(値(板2, '0,0'), undefined, '板2 A1'); eq(値(0, '0,0'), '板1', '板1 A1'); });

新しい本();
win.setCell(0, 0, 'a'); win.setCell(1, 0, 'b'); win.doUndo();
打ちかけ(5, 5, 'x');
await T('★打ちかけのまま ↪＝やり直さない★', () => { win.doRedo(); 抜けて確定(); eq(値(0, '1,0'), undefined, 'A2'); eq(値(0, '5,5'), 'x', 'F6'); });
await T('打った後は やり直しが 消える（実Excel と 同じ＝新しい 操作）', () => { eq(win.redoStack.length, 0, 'redoStack'); });

新しい本();
win.setCell(0, 0, 'a');
await T('打っていない時の ↩／↪ は 今のまま', () => { win.doUndo(); eq(値(0, '0,0'), undefined, 'Z'); win.doRedo(); eq(値(0, '0,0'), 'a', 'Y'); });

/* ★数式の帯★＝実Excel は 数式バーで 打っている間も 元に戻すが 押せない */
const fi = doc.getElementById('formula-input'), kb = doc.getElementById('formula-kb');
const ボタン = (i) => doc.querySelectorAll('.fmt-undo-redo')[i];
/* 本物の 押し方＝mousedown の 既定（入力欄から 抜ける＝blur）が 先・その後 click（doUndo） */
const 押す = (b) => {
  const md = new win.MouseEvent('mousedown', { bubbles: true, cancelable: true });
  b.dispatchEvent(md);
  if (!md.defaultPrevented && doc.activeElement && doc.activeElement !== doc.body) doc.activeElement.blur();
  b.dispatchEvent(new win.MouseEvent('click', { bubbles: true, cancelable: true }));
};

新しい本();
win.setCell(0, 0, 'a'); win.sel(1, 0, 1, 0); win.updateBar();
kb.style.display = 'block'; win.eval('_editMode=true'); fi.value = '自前キーボードの打ちかけ';
await T('★自前キーボードの帯で 打ちかけのまま ↩＝前の操作を 戻さない★', () => { 押す(ボタン(0)); eq(値(0, '0,0'), 'a', 'A1'); eq(fi.value, '自前キーボードの打ちかけ', '帯の字'); });
await T('帯を 確定した後の ↩ は 帯で 打った字から 戻す', () => { win.fkbCommit(); eq(値(0, '1,0'), '自前キーボードの打ちかけ', 'A2 確定'); win.doUndo(); eq(値(0, '1,0'), undefined, 'A2'); eq(値(0, '0,0'), 'a', 'A1'); });

新しい本();
win.setCell(0, 0, 'a'); win.sel(1, 0, 1, 0); win.updateBar(); kb.style.display = 'none';
fi.focus(); fi.value = 'PCの帯で打った'; fi.dispatchEvent(new win.Event('input'));
await T('★PC の帯で 打ちかけのまま ↩ ボタン＝帯から 抜けず 前の操作も 戻さない★', () => { 押す(ボタン(0)); eq(doc.activeElement, fi, '帯に 居る'); eq(値(0, '0,0'), 'a', 'A1'); eq(win.undoStack.length, 1, '控えの数'); });
await T('PC の帯から 抜けたら 確定し、その後の ↩ で 戻る', () => { fi.blur(); eq(値(0, '1,0'), 'PCの帯で打った', 'A2 確定'); win.doUndo(); eq(値(0, '1,0'), undefined, 'A2'); eq(値(0, '0,0'), 'a', 'A1'); });

await T('↩／↪ ボタンの mousedown は 既定を 止める（押しても 入力欄から 抜けない）', () => {
  for (const i of [0, 1]) { const md = new win.MouseEvent('mousedown', { bubbles: true, cancelable: true }); ボタン(i).dispatchEvent(md); ok(md.defaultPrevented, 'ボタン' + i); }
});

console.log('\n' + pass + ' passed, ' + fail + ' failed');
if (SELF) {
  const 効く = fail >= 7;
  console.log(効く ? '★--self-test：直す前の形で ' + fail + ' 本 赤＝見張りが効いている★' : '★--self-test：赤が ' + fail + ' 本しか出ない＝空振り★');
  process.exit(効く ? 0 : 1);
}
process.exit(fail ? 1 : 0);
