// grid-edit-ui.mjs — ★本物の book.html を丸ごと読み込み、本物の insertRefAddr を動かす★
//   純関数のテスト(grid-refedit.test.mjs)は「判定」を固定する。
//   ここでは ★画面の中の本物の関数★ を呼んで、数式バーの値が実際に直るかを見る。
//   依存: jsdom。未導入なら赤（SKIPを緑と呼ばない）。
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(__dirname, '..');
let JSDOM;
try { ({ JSDOM } = await import('jsdom')); }
catch { console.log('★jsdomが入っていません。この検証は飛ばせません（SKIPを緑と呼ばない）。npm install してください。'); process.exit(1); }

let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ok   ' + n); } catch (e) { fail++; console.log('  NG   ' + n + '\n       ' + (e && e.message)); } };
const ok = (c, m) => { if (!c) throw new Error(m || 'expected truthy'); };
const eq = (a, b, m) => { if (a !== b) throw new Error((m ? m + ': ' : '') + 'expected ' + JSON.stringify(b) + ' got ' + JSON.stringify(a)); };

console.log('\n[grid-edit-ui] 本物の book.html で「書き間違えた式を直せるか」');

const html = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf8');

/* canvas は jsdom に無いので、最低限の偽物を用意する（描画は見ない・落とさないため） */
const CANVAS_STUB = `
(function(){
  var noop=function(){};
  var ctx={ save:noop,restore:noop,beginPath:noop,moveTo:noop,lineTo:noop,stroke:noop,fill:noop,
    fillRect:noop,strokeRect:noop,clearRect:noop,rect:noop,clip:noop,arc:noop,closePath:noop,
    fillText:noop,strokeText:noop,setLineDash:noop,translate:noop,scale:noop,setTransform:noop,
    measureText:function(){return {width:40};}, createLinearGradient:function(){return {addColorStop:noop};},
    drawImage:noop, putImageData:noop, getImageData:function(){return {data:[]};} };
  HTMLCanvasElement.prototype.getContext=function(){ return ctx; };
})();`;

const dom = new JSDOM(html, {
  runScripts: 'dangerously', url: 'http://localhost/', pretendToBeVisual: true,
  resources: undefined,
  beforeParse(w) {
    w.fetch = () => Promise.reject(new Error('no net'));
    w.scrollTo = () => {};
    w.matchMedia = w.matchMedia || (() => ({ matches: false, addListener() {}, removeListener() {}, addEventListener() {}, removeEventListener() {} }));
    w.requestAnimationFrame = (cb) => setTimeout(() => cb(Date.now()), 0);
    w.cancelAnimationFrame = () => {};
    w.eval(CANVAS_STUB);
  },
});
const win = dom.window, doc = win.document;

/* 外部scriptは jsdom が取りに行かないので、手で流し込む（?v= を落として実ファイルを読む） */
const srcs = [...html.matchAll(/<script src="([^"]+)"><\/script>/g)].map((m) => m[1].split('?')[0])
  .filter((s) => !/^https?:/.test(s));
const loaded = [], skipped = [];
for (const src of srcs) {
  const p = path.join(ROOT, src);
  if (!fs.existsSync(p)) { skipped.push(src + '(無い)'); continue; }
  try { win.eval(fs.readFileSync(p, 'utf8')); loaded.push(src); }
  catch (e) { skipped.push(src + '(' + String(e.message).slice(0, 40) + ')'); }
}
console.log('  読み込んだ部品: ' + loaded.join(' / '));
if (skipped.length) console.log('  読めなかった部品: ' + skipped.join(' / '));

T('lib/grid-refedit.js が入っている（window.GridRefEdit）', () => {
  ok(win.GridRefEdit && typeof win.GridRefEdit.refEditAt === 'function', 'GridRefEdit が無い');
});

/* ── ★本物の insertRefAddr を動かす★ ── */
const hasReal = typeof win.insertRefAddr === 'function' && typeof win.getRefInput === 'function';
console.log('  本物の insertRefAddr: ' + (hasReal ? '★あり（これを動かす）★' : 'なし（純関数で確かめる）'));

function drive(formula, cursor, pickR, pickC) {
  /* 数式バーに式を入れて、カーソルを置いて、セルを選んだ時と同じ呼び出しをする */
  const fi = doc.getElementById('formula-input');
  ok(fi, 'formula-input が無い');
  fi.value = formula;
  try { fi.setSelectionRange(cursor, cursor); } catch (e) { /* jsdom差 */ }
  win._fiCursorPos = cursor;
  win.refMode = true;
  win._editMode = true;
  win.insertRefAddr(pickR, pickC, pickR, pickC);
  return fi.value;
}

if (hasReal) {
  T('★A2 の「=B1+30」で B1 にカーソルを置いて A1 を選ぶ → =A1+30★（本物の関数）', () => {
    eq(drive('=B1+30', 2, 0, 0), '=A1+30');
  });
  T('★カーソルが末尾（数字の直後）で A1 を選んでも式が壊れない★（前は =B1+30A1）', () => {
    eq(drive('=B1+30', 6, 0, 0), '=B1+30');
  });
  T('今までの組み立て方が壊れていない（=A1+ → B2 を差し込む）', () => {
    eq(drive('=A1+', 4, 1, 1), '=A1+B2');
  });
} else {
  const R = win.GridRefEdit;
  T('★=B1+30 の B1 を A1 に直せる★', () => { eq(R.refEditAt('=B1+30', 2, 'A1').v, '=A1+30'); });
  T('★末尾で触っても壊れない★', () => { eq(R.refEditAt('=B1+30', 6, 'A1').v, '=B1+30'); });
  T('組み立ては今までどおり', () => { eq(R.refEditAt('=A1+', 4, 'B2').v, '=A1+B2'); });
}

console.log(`\n  ${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
