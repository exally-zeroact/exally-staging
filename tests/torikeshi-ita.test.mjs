/* torikeshi-ita.test.mjs — ★取り消し／やり直しは 控えを取った板へ 戻す★（押し0a・2026-11-07）
 *  ★なぜ★ 前は 開いている板へ 書き戻した＝実測（ef6cf5a・jsdom）：
 *    板1で行を挿入→板2へ移って Ctrl+Z ＝★板2 が板1の控えで上書き（名が Sheet1 になり中身が消える）・板1 は戻らない★
 *    板1で太字→板2で Ctrl+Z ＝★板1 は太字のまま・板2 の A1 に空のマスが出来る★
 *    本を開き直しても 前の本の控えが 残っていた（openBookFile は undoStack を空にしなかった）
 *  使い方: node tests/torikeshi-ita.test.mjs ／ --self-test＝控えから板の番号を消す（直す前の形）と赤になるか
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
for (const f of ['lib/xlsx.full.min.js', 'lib/xlsx-io.js', 'lib/zip-surgeon.js', 'lib/xlsx-edit.js', 'lib/xlsb-edit.js']) inject(fs.readFileSync(path.join(ROOT, f), 'utf8'));

const SELF = process.argv.includes('--self-test');
if (SELF) {
  /* ★直す前の 動きを 真似る★＝控えから 板の番号を 消す（開いている板へ 書き戻す 前の形） */
  win.eval("(function(){ var p=Array.prototype.push; undoStack.push=function(o){ if(o&&(o.type==='rowcol'||o.type==='format')) delete o.sheet; return p.apply(this, arguments); }; })()");
}
const 板 = (i) => win.sheets[i];
const 値 = (i, k) => { const x = 板(i).data[k]; return x ? (x.f || x.v) : undefined; };
const 新しい本 = () => { win.eval("sheets.length=1; sheets[0]={name:'Sheet1',data:{},colW:{},rowH:{}}; activeSheet=0; undoStack.length=0; redoStack.length=0;"); if (typeof win.rebuildSheetTabs === 'function') win.rebuildSheetTabs(); };
console.log('\n[torikeshi-ita] 取り消しは 控えを取った板へ' + (SELF ? '（--self-test＝赤に なるはず）' : ''));

新しい本();
win.setCell(0, 0, '1', true); win.setCell(1, 0, '2', true);
win.addSheet(); const 板2 = win.activeSheet; win.setCell(0, 0, '板2の字', true);
win.switchSheet(0); win.sel(1, 0, 1, 0); win.ctxInsertRow();
await T('行の挿入→別の板で Ctrl+Z＝別の板は そのまま', () => { win.switchSheet(板2); win.doUndo(); eq(板(板2).name, 'Sheet2', '板2の名'); eq(値(板2, '0,0'), '板2の字', '板2 A1'); });
await T('行の挿入→別の板で Ctrl+Z＝挿入した板が 戻る', () => { eq(値(0, '1,0'), '2', '板1 A2'); eq(値(0, '2,0'), undefined, '板1 A3'); });
await T('戻した板へ 移る（実Excel と同じ）', () => { eq(win.activeSheet, 0, '開いている板'); });
await T('別の板で Ctrl+Y＝挿入した板に やり直す・別の板は そのまま', () => { win.switchSheet(板2); win.doRedo(); eq(値(0, '2,0'), '2', '板1 A3'); eq(値(板2, '0,0'), '板2の字', '板2 A1'); });

新しい本();
win.setCell(0, 0, 'a', true); win.addSheet(); const 板2b = win.activeSheet;
win.switchSheet(0); win.sel(0, 0, 0, 0); win.applyFormat('bold', true);
await T('太字→別の板で Ctrl+Z＝太字の板が 戻る', () => { ok(板(0).data['0,0'].bold, '前提：太字になっていない'); win.switchSheet(板2b); win.doUndo(); eq(!!板(0).data['0,0'].bold, false, '板1 A1 太字'); });
await T('太字→別の板で Ctrl+Z＝別の板に 空のマスを 作らない', () => { eq(板(板2b).data['0,0'], undefined, '板2 A1'); });
await T('別の板で Ctrl+Y＝太字の板に やり直す', () => { win.switchSheet(板2b); win.doRedo(); eq(!!板(0).data['0,0'].bold, true, '板1 A1 太字'); eq(板(板2b).data['0,0'], undefined, '板2 A1'); });

新しい本();
win.setCell(0, 0, 'りんご', true); win.addSheet(); const 板2c = win.activeSheet; win.setCell(0, 0, '板2', true);
win.switchSheet(0); doc.getElementById('findWhat').value = 'りんご'; doc.getElementById('findTo').value = 'みかん'; win.replaceAll();
await T('全部置換→別の板で Ctrl+Z＝置換した板が 戻る・別の板は そのまま', () => { eq(値(0, '0,0'), 'みかん', '置換した後'); win.switchSheet(板2c); win.doUndo(); eq(値(0, '0,0'), 'りんご', '板1 A1'); eq(値(板2c, '0,0'), '板2', '板2 A1'); });

新しい本();
win.setCell(0, 0, 'x', true);
await T('同じ板の 値の取り消しは 今のまま', () => { win.setCell(0, 0, 'y'); win.doUndo(); eq(値(0, '0,0'), 'x', 'A1'); });

新しい本();
win.setCell(0, 0, 's', true); win.addSheet(); const 板2d = win.activeSheet;
win.switchSheet(0); win.sel(0, 0, 0, 0);
await T('セルのスタイル→別の板で Ctrl+Z＝スタイルの板が 戻る・別の板に 空のマスを 作らない', () => {
  const 前 = JSON.stringify(板(0).data['0,0']);
  ok(win.CellStyles && win.CellStyles.一覧 && win.CellStyles.一覧.length, '前提：スタイルの一覧が無い');
  let i = 0; for (; i < win.CellStyles.一覧.length; i++) { win.セルのスタイルを当てる(i); if (JSON.stringify(板(0).data['0,0']) !== 前) break; }
  ok(JSON.stringify(板(0).data['0,0']) !== 前, '前提：スタイルが 当たっていない');
  win.switchSheet(板2d); win.doUndo();
  eq(JSON.stringify(板(0).data['0,0']), 前, '板1 A1'); eq(板(板2d).data['0,0'], undefined, '板2 A1');
});

await T('書式・行列の控えを 積む所は 全部 板の番号つき（book.html の字を 数える）', () => {
  const 所 = [...html.matchAll(/undoStack\.push\(\s*\{([^}]*type\s*:\s*'(?:format|rowcol)'[^}]*)\}/g)];
  ok(所.length >= 9, '積む所が ' + 所.length + ' か所しか 見つからない（数え方が 壊れた）');
  const 抜け = 所.filter((m) => !/\bsheet\s*:/.test(m[1])).map((m) => html.slice(0, m.index).split('\n').length);
  eq(抜け.join(','), '', '板の番号が 無い行');
});

await T('本を開いたら 前の本の控えを 捨てる（openBookFile の 中）', () => {
  const f = win.openBookFile.toString();
  const i = f.indexOf('sheets = res.sheets;');
  ok(i >= 0, 'openBookFile に sheets = res.sheets が無い');
  const 後 = f.slice(i, i + 400);
  ok(/undoStack\s*=\s*\[\]/.test(後) && /redoStack\s*=\s*\[\]/.test(後), '開いた直後に undoStack と redoStack を空にしていない');
});

console.log('\n' + pass + ' passed, ' + fail + ' failed');
if (SELF) {
  const 効く = fail >= 5;
  console.log(効く ? '★--self-test：直す前の形で ' + fail + ' 本 赤＝見張りが効いている★' : '★--self-test：赤が ' + fail + ' 本しか出ない＝空振り★');
  process.exit(効く ? 0 : 1);
}
process.exit(fail ? 1 : 0);
