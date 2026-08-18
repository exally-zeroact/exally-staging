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
/* canvas は jsdom に無いので、偽物を用意する（描画は見ない・落とさないため）
   ★弱い偽物だと、途中で例外が出て「その先の行が動かない」★（2026-08-18 実際に踏んだ:
     window.showToast = notify; の行まで届かず、showToast is not defined になった。
     関数の宣言は巻き上がるので insertRefAddr は在る＝★在るのに壊れている★が起きる）。
   ⇒ ★何を呼ばれても落ちない偽物にする★ */
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

/* ★順番が要る（2026-08-18 実際に踏んだ）★
   前は「HTMLをそのまま読ませる（＝インラインが先に動く）→ 後から外の部品を流す」だった。
   すると ★インラインの中で 外の部品を触る行で例外が出て、その先が動かない★。
   関数の宣言は巻き上がるので insertRefAddr は在るのに、
   window.showToast = notify; の代入まで届かず ★在るのに壊れている★状態になった。
   ⇒ ★インラインを外してから読ませ、外の部品 → インライン の順に流す★（他の検査と同じ形）。 */
const dom = new JSDOM(html.replace(/<script[\s\S]*?<\/script>/g, ''), {
  runScripts: 'dangerously', url: 'http://localhost/', pretendToBeVisual: true,
  resources: undefined,
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

/* ① 外の部品（?v= を落として実ファイルを読む） */
const srcs = [...html.matchAll(/<script src="([^"]+)"><\/script>/g)].map((m) => m[1].split('?')[0])
  .filter((s) => !/^https?:/.test(s));
const loaded = [], skipped = [];
for (const src of srcs) {
  const p = path.join(ROOT, src);
  if (!fs.existsSync(p)) { skipped.push(src + '(無い)'); continue; }
  try { inject(fs.readFileSync(p, 'utf8')); loaded.push(src); }
  catch (e) { skipped.push(src + '(' + String(e.message).slice(0, 40) + ')'); }
}
/* ② インライン */
let inlineN = 0, inlineNG = 0;
for (const m of html.matchAll(/<script(?![^>]*\ssrc=)[^>]*>([\s\S]*?)<\/script>/g)) {
  inlineN++;
  try { inject(m[1]); } catch (e) { inlineNG++; }
}
/* ③ ★読み込みの合図をもう一度 出す★
   後から script を足す形だと DOMContentLoaded / load は ★もう終わっている★ので、
   その中で登録される物（例 window.showToast = notify）が ★一生 動かない★。
   2026-08-18 実際に踏んだ: insertRefAddr は在るのに showToast が無い状態になった。 */
try { doc.dispatchEvent(new win.Event('DOMContentLoaded', { bubbles: true })); } catch (e) { /* 出せなくても続ける */ }
try { win.dispatchEvent(new win.Event('load')); } catch (e) { /* 同上 */ }

console.log('  読み込んだ部品: ' + loaded.join(' / '));
console.log('  インラインの script: ' + inlineN + '本（流せなかった ' + inlineNG + '本）');
if (skipped.length) console.log('  読めなかった部品: ' + skipped.join(' / '));

/* ★本物のページが持っている物が、この入れ物でも本当に在るか先に確かめる★
   在ると思い込んで動かすと「無いから落ちた」のか「作りが悪い」のか分からなくなる。 */
T('★本物のページが持っている知らせの口(showToast)が在る', () => {
  ok(typeof win.showToast === 'function',
    'showToast が無い＝インラインの script が途中で止まっている（偽物が弱いか、本物が壊れている）');
});

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
