/* torikeshi-dji.test.mjs — ★値の取り消しで ファイルの字の マス（d字）が 数に 化けない★（2026-11-14 件f）
 *  ★なぜ★ staging 5446df1e の doUndo（値）は 控えの v を toHFVal に 通して HF へ 入れ、d字 を 戻さなかった＝
 *    ファイルから 読んだ 字 '1,234' に 7 を 打って Ctrl+Z すると HF は 数 1234・d字 は 消え ISTEXT が 黙って 逆に なった。
 *    同じ 行で ★0 の マスが Ctrl+Z で 消えた★（!op.old）・★数の マスを 戻すと f が 数で 数式バーが 落ちた★・
 *    溢れ先（v 無し）の 控えは undefined・時刻 '9:00' の 見た目が "0.375" に なった。
 *  直し＝setCell の 控えに oldD／oldD字・redo にも 載せる・字の マスは 字のまま HF へ・f は 字・空へは 本当に 空の 時だけ。
 *  使い方: node tests/torikeshi-dji.test.mjs ／ --self-test＝直しを 1つずつ 外して 全部 赤に なるか（子を 起こす・7通り）
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { spawnSync } from 'node:child_process';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(__dirname, '..');

/* ★直しを 1つずつ 外す 壊し★（字が 当たらなければ「置換できず」＝赤と 数えない） */
const 壊し = [
  ['字の マスを 字のまま HF へ 入れない', "  var _字のまま = !!op.oldD字 && !_式 && !_空へ;", "  var _字のまま = false;"],
  ['やり直しに 印を 載せない', ", oldD:op.oldD, oldD字:op.oldD字});", "});"],
  ['0 を 空と 見る（前の !op.old）', "  var _空へ = (op.old === '' || op.old === null || op.old === undefined);", "  var _空へ = !op.old;"],
  ['f に 数を 入れる', "{v:op.old, f:String(op.old), d:String(", "{v:op.old, f:op.old, d:String("],
  ['前の 見た目に 戻さない', "    if(!_式 && op.oldD !== undefined && op.oldD !== null) d = op.oldD;", ""],
  ['溢れ先の 空（undefined）を 空と 見ない', "(op.old === '' || op.old === null || op.old === undefined);", "(op.old === '' || op.old === null);"],
  ['帯の 下書きを 前の 見た目と 控える', "var _前のd = _下書き中 ? _打ちかけ前.d :", "var _前のd = false ? _打ちかけ前.d :"],
  ['式の「答えは字」を 戻さない', "if(_字のまま || (_式 && typeof _hf答えの型", "if(_字のまま || (false && typeof _hf答えの型"],
];

const orig = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf8');
if (process.argv.includes('--self-test')) {
  console.log('\n[torikeshi-dji --self-test] 直しを 1つずつ 外して 赤に なるか');
  const 素 = spawnSync(process.execPath, [fileURLToPath(import.meta.url)], { encoding: 'utf8' });
  if (素.status !== 0) { console.log('  ★壊していない 形が 赤★＝数えても 意味が ありません\n' + 素.stdout); process.exit(1); }
  console.log('  壊していない 形は 緑');
  let 赤 = 0;
  for (let i = 0; i < 壊し.length; i++) {
    if (orig.split(壊し[i][1]).length !== 2) { console.log('  ★置換できず★  ' + 壊し[i][0]); continue; }
    const r = spawnSync(process.execPath, [fileURLToPath(import.meta.url), '--break=' + i], { encoding: 'utf8' });
    if (r.status !== 0) { 赤++; console.log('  赤くなった  ' + 壊し[i][0]); }
    else console.log('  ★緑のまま★  ' + 壊し[i][0]);
  }
  console.log('\n  ' + 赤 + '/' + 壊し.length + ' 通りで赤くなった');
  process.exit(赤 === 壊し.length ? 0 : 1);
}

let JSDOM;
try { ({ JSDOM } = await import('jsdom')); }
catch { console.log('★jsdomが入っていません（SKIPを緑と呼ばない）。npm ci してください。'); process.exit(1); }

let html = orig;
const 壊す番 = process.argv.find((a) => a.startsWith('--break='));
if (壊す番) { const k = 壊し[Number(壊す番.slice(8))]; html = html.replace(k[1], k[2]); }

let pass = 0, fail = 0;
const T = async (n, fn) => { try { await fn(); pass++; console.log('  ok   ' + n); } catch (e) { fail++; console.log('  NG   ' + n + '\n       ' + (e && e.message)); } };
const ok = (c, m) => { if (!c) throw new Error(m || 'expected truthy'); };
const eq = (a, b, m) => { if (a !== b) throw new Error((m ? m + ': ' : '') + '期待=' + JSON.stringify(b) + ' 実際=' + JSON.stringify(a)); };
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
const E = (s) => win.eval(s);
const 新しい本 = () => { E("sheets.length=1; sheets[0]={name:'Sheet1',data:{},colW:{},rowH:{}}; activeSheet=0; undoStack.length=0; redoStack.length=0;"); if (typeof win.rebuildSheetTabs === 'function') win.rebuildSheetTabs(); };
const HF = (r, c, i = 0) => E(`JSON.stringify(hf.getCellValue({sheet:_hfSid(${i}),row:${r},col:${c}}))`);
const マス = (k, i = 0) => win.sheets[i].data[k];
const 置く = (k, cell) => { E(`sheets[0].data['${k}']=${JSON.stringify(cell)}; _pushGrid(0,sheets[0]);`); };
console.log('\n[torikeshi-dji] 値の取り消しで ファイルの字の マス（d字）が 数に 化けない' + (壊す番 ? '（' + 壊す番 + '）' : ''));

新しい本();
置く('0,0', { v: '1,234', f: '1,234', d: '1,234', d字: true });
await T('★ファイルの字 1,234 に 7→Ctrl+Z＝HF は 字 "1,234" のまま★', () => { eq(HF(0, 0), '"1,234"', '打つ前'); win.setCell(0, 0, '7'); eq(HF(0, 0), '7', '打った後'); win.doUndo(); eq(HF(0, 0), '"1,234"', 'Ctrl+Z の後'); });
await T('★マスも 戻る（d は 1,234・d字 は true）★', () => { eq(マス('0,0').d, '1,234', 'd'); eq(マス('0,0').d字, true, 'd字'); eq(マス('0,0').v, '1,234', 'v'); });
await T('★↪→↩ でも 字に 戻る（印が やり直しを 通る）★', () => { win.doRedo(); eq(HF(0, 0), '7', '↪の後'); win.doUndo(); eq(HF(0, 0), '"1,234"', '↩の後'); eq(マス('0,0').d字, true, 'd字'); });

新しい本();
置く('0,0', { v: '1,234', f: '1,234', d: '1,234', d字: true });
win.setCell(0, 0, '7'); win.addSheet();
await T('★別の板で Ctrl+Z（打った板へ 移る 道）でも 字のまま★', () => { win.doUndo(); eq(win.activeSheet, 0, '板'); eq(HF(0, 0), '"1,234"', 'HF'); eq(マス('0,0').d字, true, 'd字'); });

新しい本();
win.setCell(0, 0, '=TEXT(1,"0")');
await T('★=TEXT の 式に 5→Ctrl+Z＝「答えは字」が 戻る★', () => { eq(マス('0,0').d字, true, '打った時'); win.setCell(0, 0, '5'); win.doUndo(); eq(マス('0,0').f, '=TEXT(1,"0")', 'f'); eq(マス('0,0').d字, true, 'd字'); });

新しい本();
win.setCell(0, 0, '1,234');
await T('打った 1,234（字で ない）に 7→Ctrl+Z＝HF は 数 1234 のまま（今の 動きを 壊さない）', () => { win.setCell(0, 0, '7'); win.doUndo(); eq(HF(0, 0), '1234', 'HF'); ok(!マス('0,0').d字, 'd字 が 付いた'); });

新しい本();
置く('2,0', { d: 'x', d字: true, _溢れ元: '5,3' });
await T('★溢れ先（v 無し）に 打って Ctrl+Z＝HF に "undefined" を 入れない★', () => { win.setCell(2, 0, '7'); win.doUndo(); ok(HF(2, 0) !== '"undefined"', 'HF=' + HF(2, 0)); ok(!(マス('2,0') && マス('2,0').f === 'undefined'), 'f が "undefined"'); });

新しい本();
置く('0,0', { v: 0, f: '0', d: '0' });
await T('★0 の マスに 7→Ctrl+Z＝マスが 消えない（0 に 戻る）★', () => { win.setCell(0, 0, '7'); win.doUndo(); ok(マス('0,0'), 'マスが 消えた'); eq(マス('0,0').v, 0, 'v'); eq(HF(0, 0), '0', 'HF'); });

新しい本();
置く('0,0', { v: 1234, f: '1234', d: '1,234', numFmt: '#,##0' });
await T('★数の マスに 7→Ctrl+Z＝f は 字（数式バーが 落ちない）★', () => {
  win.setCell(0, 0, '7'); win.doUndo(); eq(typeof マス('0,0').f, 'string', 'f の 型'); E('sel(0,0,0,0)'); win.updateBar(); eq(マス('0,0').d, '1,234', 'd');
});

新しい本();
win.setCell(0, 0, '9:00');
await T('★時刻 9:00 に 7→Ctrl+Z＝見た目は 9:00（0.375 に しない）★', () => { const 前 = マス('0,0').d; win.setCell(0, 0, '7'); win.doUndo(); eq(マス('0,0').d, 前, 'd'); });

/* ★数式の帯で 確定→Ctrl+Z＝見える字（d）も 前に 戻る★（2026-11-14・件f の続き・Exally [424b60] の taiketsu が 測った）
   帯は 打つ度に d を 下書きに 書き換える（件⑥）＝setCell が その下書きを「前の 見た目」と 控え、Ctrl+Z で v は 100 なのに d は 200 が 残った */
const fi = doc.getElementById('formula-input'), kb = doc.getElementById('formula-kb');
const 帯で打つ = (字) => { fi.value = 字; fi.dispatchEvent(new win.Event('input')); };
const 選ぶ = (r, c) => { win.sel(r, c, r, c); win.updateBar(); };
const 帯の確定 = {
  'Enter': () => { kb.style.display = 'none'; fi.focus(); 帯で打つ('200'); fi.dispatchEvent(new win.KeyboardEvent('keydown', { key: 'Enter', bubbles: true })); },
  'blur（抜ける）': () => { kb.style.display = 'none'; fi.focus(); 帯で打つ('200'); fi.blur(); },
  '自前キーボードの 確定': () => { kb.style.display = 'block'; E('_editMode=true'); 帯で打つ('200'); win.fkbCommit(); E('_editMode=false'); kb.style.display = 'none'; },
};
for (const [名, 確定] of Object.entries(帯の確定)) {
  新しい本(); E('_打ちかけ前=null'); win.setCell(0, 0, '100'); 選ぶ(0, 0); 確定();
  await T('★帯で 100→200 を ' + 名 + '→Ctrl+Z＝v も d も 100★', () => { eq(String(マス('0,0').v), '200', '確定した v'); win.doUndo(); eq(String(マス('0,0').v), '100', 'v'); eq(String(マス('0,0').d), '100', 'd'); eq(HF(0, 0), '100', 'HF'); });
}
新しい本(); E('_打ちかけ前=null'); 選ぶ(1, 0); kb.style.display = 'none'; fi.focus(); 帯で打つ('5'); fi.dispatchEvent(new win.KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
await T('★空のマスに 帯で 5 を 確定→Ctrl+Z＝空に 戻る（下書きの 5 を 残さない）★', () => { win.doUndo(); ok(!マス('1,0') || マス('1,0').d === '' || マス('1,0').d === undefined, 'd=' + JSON.stringify(マス('1,0') && マス('1,0').d)); });
新しい本(); E('_打ちかけ前=null'); win.setCell(0, 0, '100'); 選ぶ(0, 0); kb.style.display = 'block'; 帯で打つ('ab'); E('_fiCursorPos=1'); win.insertLineBreak(); kb.style.display = 'none';
await T('★帯の 中で 改行（setCell が 走る）→Ctrl+Z＝見える字は 100★', () => { win.doUndo(); eq(String(マス('0,0').v), '100', 'v'); eq(String(マス('0,0').d), '100', 'd'); });
新しい本(); E('_打ちかけ前=null');
置く('0,0', { v: '1,234', f: '1,234', d: '1,234', d字: true }); 選ぶ(0, 0); kb.style.display = 'none'; fi.focus(); 帯で打つ('7'); fi.blur();
await T('★ファイルの字 1,234 に 帯で 7→Ctrl+Z＝見える字 1,234・HF は 字★', () => { win.doUndo(); eq(マス('0,0').d, '1,234', 'd'); eq(HF(0, 0), '"1,234"', 'HF'); });

console.log('\n' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
