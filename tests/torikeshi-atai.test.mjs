/* torikeshi-atai.test.mjs — ★値の取り消し／やり直しは 打った板へ・シートの保護は 控えを 捨てる★（2026-11-13）
 *  ★なぜ★ 本番 818bd778 の doRedo は 値の時 setCell（開いている板へ 書く）＝板1で打つ→板2へ移って Ctrl+Z→Ctrl+Y で
 *    ★板2 の 同じマスが 板1 の字で 上書きされた（客のデータが 壊れる）★。値の Ctrl+Z は 打った板へ 戻すが 移らなかった。
 *    行の挿入→シートを保護→Ctrl+Z で ★保護まで 巻き戻った★（行の控えは 板を 丸ごと 写すので protect も 入る）。
 *    ★条件付き書式・ロック・範囲の許可など 控えを積まない直しも 同じ形で 巻き戻る＝穴①（積まない操作）の側で 塞ぐ★
 *  直し＝値の道でも 控えの板へ 移る（redoStack へ 積む前＝打ちかけの commitEdit が redoStack を空にするため）／保護を変えたら 控えの中の protect を 今の物に合わせる
 *  使い方: node tests/torikeshi-atai.test.mjs ／ --self-test＝直す前の形（値で 板を移らない・控えの protect を 合わせない）で 赤になるか
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
/* ★保存の 道具は 本番でも 初めての 保存で 後から 読む（_ensureXlsx と 同じ 順）★ */

const SELF = process.argv.includes('--self-test');
if (SELF) {
  /* ★直す前の 動きを 真似る★＝値の取り消し・やり直しでは 板を移らない／保護しても 控えを 捨てない */
  win.eval("(function(){ var sw=switchSheet; var 中=false;"
    + " var u=doUndo, r=doRedo;"
    + " function 値か(st){ var o=st[st.length-1]; return o && !o.type; }"
    + " window.doUndo=function(){ if(値か(undoStack)){ var k=switchSheet; window.switchSheet=function(){}; try{ return u(); } finally{ window.switchSheet=k; } } return u(); };"
    + " window.doRedo=function(){ if(値か(redoStack)){ var k=switchSheet; window.switchSheet=function(){}; try{ return r(); } finally{ window.switchSheet=k; } } return r(); };"
    + " window._控えの保護を合わせる=function(){};"
    + " })()");
}
const 板 = (i) => win.sheets[i];
const 値 = (i, k) => { const x = 板(i).data[k]; return x ? (x.f || x.v) : undefined; };
const 新しい本 = () => { win.eval("sheets.length=1; sheets[0]={name:'Sheet1',data:{},colW:{},rowH:{}}; activeSheet=0; undoStack.length=0; redoStack.length=0;"); if (typeof win.rebuildSheetTabs === 'function') win.rebuildSheetTabs(); };
console.log('\n[torikeshi-atai] 値の取り消し・やり直しは 打った板へ／保護は 巻き戻さない' + (SELF ? '（--self-test＝赤に なるはず）' : ''));

新しい本();
win.addSheet(); const 板2 = win.activeSheet; win.setCell(0, 0, '板2の字', true);
win.switchSheet(0); win.setCell(0, 0, '打った');
await T('値を打つ→別の板で Ctrl+Z＝打った板が 戻る・別の板は そのまま', () => { win.switchSheet(板2); win.doUndo(); eq(値(0, '0,0'), undefined, '板1 A1'); eq(値(板2, '0,0'), '板2の字', '板2 A1'); });
await T('値の Ctrl+Z で 打った板へ 移る（実Excel と同じ）', () => { eq(win.activeSheet, 0, '開いている板'); });
await T('★別の板で Ctrl+Y＝打った板に やり直す・別の板を 上書きしない★', () => { win.switchSheet(板2); win.doRedo(); eq(値(0, '0,0'), '打った', '板1 A1'); eq(値(板2, '0,0'), '板2の字', '板2 A1'); });
await T('値の Ctrl+Y で 打った板へ 移る', () => { eq(win.activeSheet, 0, '開いている板'); });
await T('やり直した後の Ctrl+Z も 打った板だけ 戻す', () => { win.switchSheet(板2); win.doUndo(); eq(値(0, '0,0'), undefined, '板1 A1'); eq(値(板2, '0,0'), '板2の字', '板2 A1'); });

新しい本();
win.setCell(0, 0, 'x', true);
await T('同じ板の 値の取り消し・やり直しは 今のまま', () => { win.setCell(0, 0, 'y'); win.doUndo(); eq(値(0, '0,0'), 'x', 'Z の後'); win.doRedo(); eq(値(0, '0,0'), 'y', 'Y の後'); eq(win.activeSheet, 0, '板'); });

新しい本();
win.addSheet(); const 板2e = win.activeSheet; win.setCell(0, 0, '板2の字', true);
win.switchSheet(0); win.setCell(0, 0, '打った');
/* ★打ちかけのまま ↩ は 何もしない（穴⑤・実Excel は 編集中 押せない＝tests/torikeshi-uchikake）＝確定してから 戻す形で 見る★ */
await T('★別の板で 打って確定→↩→↩→↪↪＝別の板を 上書きしない・やり直しが 消えない★', () => {
  win.switchSheet(板2e); win.startEdit(1, 1); doc.getElementById('cell-input').value = '打った2'; win.commitEdit();
  win.doUndo(); eq(値(板2e, '1,1'), undefined, '1回目の↩ 板2 B2');
  win.doUndo(); eq(値(0, '0,0'), undefined, '2回目の↩ 板1 A1'); eq(win.activeSheet, 0, '↩の後の板');
  eq(win.redoStack.length, 2, 'redoStack');
  win.doRedo(); eq(値(0, '0,0'), '打った', '↪の後 板1 A1'); eq(値(板2e, '0,0'), '板2の字', '板2 A1');
  win.doRedo(); eq(値(板2e, '1,1'), '打った2', '↪↪の後 板2 B2'); eq(値(板2e, '0,0'), '板2の字', '板2 A1');
});

新しい本();
win.setCell(0, 0, '1', true); win.setCell(1, 0, '2', true);
win.sel(1, 0, 1, 0); win.ctxInsertRow(); win.シートを保護する();
await T('★行を入れる→保護→Ctrl+Z＝行は 戻る・保護は 外れない★', () => { win.doUndo(); eq(!!板(0).protect, true, '保護'); eq(値(0, '1,0'), '2', 'A2'); eq(値(0, '2,0'), undefined, 'A3'); });
await T('保護しても ほかの控えは 捨てない（Ctrl+Y で 行の挿入を やり直せる・保護は そのまま）', () => { win.doRedo(); eq(値(0, '2,0'), '2', 'A3'); eq(!!板(0).protect, true, '保護'); });
win.シートの保護をやめる();
await T('保護を やめた後の Ctrl+Z で 保護が 戻らない', () => { win.doUndo(); eq(!!板(0).protect, false, '保護'); });
/* ★やり直しの後の 行の Ctrl+Z が 正しく 戻るかは 穴③（表の席の持ち分・doRedo が op を そのまま積む）＝ここでは 見ない★ */

console.log('\n' + pass + ' passed, ' + fail + ' failed');
if (SELF) {
  const 効く = fail >= 5;
  console.log(効く ? '★--self-test：直す前の形で ' + fail + ' 本 赤＝見張りが効いている★' : '★--self-test：赤が ' + fail + ' 本しか出ない＝空振り★');
  process.exit(効く ? 0 : 1);
}
process.exit(fail ? 1 : 0);
