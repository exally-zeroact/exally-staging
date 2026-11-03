/* hyou-ireru-ui.test.mjs — ★AIが出した表を「入れる」／★1押しで 元に戻る★★
 *
 *  ★司さん 2026-09-05★
 *    「コピーとかやなくて ★反映するとか★ なんか他の言い回しで
 *      ★押したら 自動で★ できるようにせな やり方が 古いやろ」
 *    「入れてええが 入れて（見積書や 在庫管理表など 表）
 *      ★やっぱ 思っとるのと 違うと 思ったら ボタン1つで 元に戻せるよな？★」
 *
 *  ★★字だけ 見ても 意味が ない所★★
 *    「元に戻せる」は ★押して 戻るまで 見ないと 分からない★。
 *    ★前は 1セルずつ undo を 積んでいた★＝10行5列なら ★50回 押さないと 戻らない★
 *    ＝客から 見れば ★戻せない★のと 同じ。
 *    ⇒ ここでは ★本物の book.html を jsdom で 動かし、
 *      入れて → ★1回だけ★ 元に戻して → ★入れる前と 同じか★を 見る。
 *
 *  土台は tests/grid-edit-ui.mjs と 同じ（★作る前に 探した★）。
 *
 *  使い方: node tests/hyou-ireru-ui.test.mjs
 *          node tests/hyou-ireru-ui.test.mjs --self-test
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(__dirname, '..');
let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ok   ' + n); } catch (e) { fail++; console.log('  NG   ' + n + '\n       ' + (e && e.message)); } };
const ok = (c, m) => { if (!c) throw new Error(m || 'expected truthy'); };
const eq = (a, b, m) => { if (a !== b) throw new Error((m ? m + ': ' : '') + '期待=' + JSON.stringify(b) + ' 実際=' + JSON.stringify(a)); };

/** ★純関数★＝「1押しで 戻るか」の 判定その物（self-test で 作り物を 通せる）
 *  @param 積んだ数  入れる時に undo に 積んだ 数
 *  @param 入れたセル数 */
export function 一押しで戻るか(積んだ数, 入れたセル数) {
  if (積んだ数 <= 0) return { 戻る: false, なぜ: '控えを 1つも 積んでいない＝戻せない' };
  if (積んだ数 === 1) return { 戻る: true, なぜ: '' };
  return { 戻る: false, なぜ: '控えを ' + 積んだ数 + '個 積んでいる＝' + 積んだ数 + '回 押さないと 戻らない'
    + '（入れたセル ' + 入れたセル数 + '個）' };
}

if (process.argv.includes('--self-test')) {
  console.log('\n[hyou-ireru-ui --self-test] わざと壊して赤になるか');
  T('★★1セルずつ 積む 作りは 赤（2026-09-06 まで これだった）★★', () => {
    const r = 一押しで戻るか(50, 50);
    ok(!r.戻る, '★50回 押さないと 戻らないのに 通した★');
    ok(/50回 押さないと/.test(r.なぜ), '理由が 違う: ' + r.なぜ);
  });
  T('★1つだけ 積むなら 通る', () => { ok(一押しで戻るか(1, 50).戻る); });
  T('★1つも 積まないのは 赤（戻せない）', () => {
    const r = 一押しで戻るか(0, 50);
    ok(!r.戻る); ok(/1つも 積んでいない/.test(r.なぜ));
  });
  console.log('\n' + pass + ' passed, ' + fail + ' failed');
  process.exit(fail ? 1 : 0);
}

/* ══ 本番＝本物の book.html を 動かす ═══════════════════════ */
let JSDOM;
try { ({ JSDOM } = await import('jsdom')); }
catch { console.log('★jsdomが入っていません。この検証は飛ばせません（SKIPを緑と呼ばない）。npm install してください。'); process.exit(1); }

console.log('\n[hyou-ireru-ui] 本物の book.html で「入れて・1押しで 戻る」か');

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
    w.scrollTo = () => {};
    w.alert = () => {};
    w.matchMedia = w.matchMedia || (() => ({ matches: false, addListener() {}, removeListener() {}, addEventListener() {}, removeEventListener() {} }));
    w.requestAnimationFrame = (cb) => setTimeout(() => cb(Date.now()), 0);
    w.cancelAnimationFrame = () => {};
    w.eval(CANVAS_STUB);
  },
});
const win = dom.window, doc = win.document;
const inject = (code) => { const el = doc.createElement('script'); el.textContent = code; doc.body.appendChild(el); };

const srcs = [...html.matchAll(/<script src="([^"]+)"><\/script>/g)].map((m) => m[1].split('?')[0])
  .filter((s) => !/^https?:/.test(s));
const loaded = [], skipped = [];
for (const src of srcs) {
  const p = path.join(ROOT, src);
  if (!fs.existsSync(p)) { skipped.push(src + '(無い)'); continue; }
  try { inject(fs.readFileSync(p, 'utf8')); loaded.push(src); }
  catch (e) { skipped.push(src + '(' + String(e.message).slice(0, 40) + ')'); }
}
let inlineNG = 0;
for (const m of html.matchAll(/<script(?![^>]*\ssrc=)[^>]*>([\s\S]*?)<\/script>/g)) {
  try { inject(m[1]); } catch (e) { inlineNG++; }
}
try { doc.dispatchEvent(new win.Event('DOMContentLoaded', { bubbles: true })); } catch (e) { /* 続ける */ }
try { win.dispatchEvent(new win.Event('load')); } catch (e) { /* 続ける */ }

T('★本物の 画面が 立ち上がっている（立たないと 何を 測っても 嘘）', () => {
  ok(typeof win.setCell === 'function', 'setCell が 無い＝インラインが 途中で 止まっている');
  ok(typeof win.doUndo === 'function', 'doUndo が 無い');
  ok(typeof win.AIの表をシートに入れる === 'function', '★入れる 口が 無い★');
  ok(typeof win._pushRowColUndo === 'function', '控えを 積む 口が 無い');
});

/* ★入れる 前の 姿を 控える★（[[feedback_baseline_must_be_copied_before_measuring]]） */
const 前の姿 = JSON.stringify(win.sheets[win.activeSheet].data);
const 前のundo = win.undoStack.length;

const 表 = [
  ['品名', '数量', '単価', '金額'],
  ['ねじ', '10', '50', '=B2*C2'],
  ['ボルト', '4', '120', '=B3*C3'],
  ['合計', '', '', '=SUM(D2:D3)'],
];
win._AIの表たち.push(表);
const 番 = win._AIの表たち.length - 1;
win.selR1 = 0; win.selC1 = 0;
const 入れた = win.AIの表をシートに入れる(番);
const 積んだ = win.undoStack.length - 前のundo;

T('★表が 本当に 入った（セルに 値が 在る）', () => {
  eq(入れた, 4, '入れた行数');
  const d = win.sheets[win.activeSheet].data;
  eq(d['0,0'] && d['0,0'].v, '品名', 'A1');
  eq(d['1,0'] && d['1,0'].v, 'ねじ', 'A2');
  eq(d['3,3'] && d['3,3'].v, '=SUM(D2:D3)', 'D4（数式も そのまま）');
});

T('★★控えは 1つだけ（1押しで 戻る作り）★★', () => {
  const r = 一押しで戻るか(積んだ, 16);
  ok(r.戻る, '★' + r.なぜ + '★');
});

T('★★1回 押したら 入れる前に 戻る（絵ではなく 中身で 見る）★★', () => {
  win.doUndo();
  const 後の姿 = JSON.stringify(win.sheets[win.activeSheet].data);
  if (後の姿 !== 前の姿) {
    throw new Error('★1回では 戻っていない★'
      + '\n   入れる前 … ' + 前の姿.slice(0, 120)
      + '\n   1回 押した後 … ' + 後の姿.slice(0, 120));
  }
});

/* ★★逆向きも 見る★★＝「いつも 戻る」に なっていないか
   ＝控えを 積まなければ 戻らない事を 実際に 確かめる（見張りが 効いている 証拠） */
T('★★控えを 積まずに 入れたら 戻らない（この 検査が 空振りしていない）★★', () => {
  const 前 = JSON.stringify(win.sheets[win.activeSheet].data);
  win.setCell(0, 0, 'わざと', true);   /* ★積まずに★ 書く */
  win.doUndo();
  const 後 = JSON.stringify(win.sheets[win.activeSheet].data);
  if (後 === 前) throw new Error('★積まなくても 戻った＝この 検査は 何も 見ていない★');
  win.setCell(0, 0, '', true);
});

console.log('\n── 実測 ──');
console.log('  読み込んだ部品 … ' + loaded.length + '本' + (skipped.length ? '（読めなかった ' + skipped.join(' / ') + '）' : ''));
console.log('  流せなかった インライン … ' + inlineNG + '本');
console.log('  入れた … ' + 表.length + '行 × ' + 表[0].length + '列（' + (表.length * 表[0].length) + 'セル）');
console.log('  ★undo に 積んだ 数 … ' + 積んだ + '個★（1なら 1押しで 戻る）');
console.log('\n' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
