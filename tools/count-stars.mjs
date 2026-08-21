/* count-stars.mjs — ★客に出る字に「★」が いくつ在るかを 描き終わった画面から数える★
 *
 *  なぜ（2026-08-22 指示役の裁定）
 *    ★★は うちの覚え書きの印であって 客の字ではない★。大事さは 太字・色・大きさで出す。
 *    ★ただし 一斉置換はするな。先に数えろ★ ＝ この道具は「数えるだけ」。1バイトも直さない。
 *
 *  数え方（3つに分ける。混ぜると嘘になる）
 *    Ａ ★今 見えている字★     … 画面を組み立てた後の 見える所の ★
 *    Ｂ ★在るが 今は出ていない字★ … display:none / hidden の中の ★（押すと出る物）
 *    Ｃ ★JSが組み立てて 客に出す字★ … showToast / textContent= / innerHTML= / placeholder に
 *        直接 書いてある文字列の ★（★注記(コメント)は数えない★）
 *    ★Ｃは「文字列の中」しか見ていない＝下限★。組み立てで足される物は 数えられない。
 *
 *  使い方: node tools/count-stars.mjs [--json]
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(__dirname, '..');
const JSON_OUT = process.argv.includes('--json');

let JSDOM;
try { ({ JSDOM } = await import('jsdom')); }
catch { console.log('★jsdomが入っていません。npm install してください（SKIPを緑と呼ばない）'); process.exit(1); }

const 画面 = ['book.html', 'hub.html', 'chat.html', 'index.html'].filter((f) => fs.existsSync(path.join(ROOT, f)));

const CANVAS_STUB = [
  '(function(){var noop=function(){};var ctx=new Proxy({},{get:function(t,k){',
  ' if(k==="measureText")return function(){return{width:40};};',
  ' if(k==="canvas")return{width:900,height:600};',
  ' if(k==="getImageData")return function(){return{data:[]};};',
  ' if(k==="createLinearGradient"||k==="createPattern")return function(){return{addColorStop:noop};};',
  ' return noop;}});HTMLCanvasElement.prototype.getContext=function(){return ctx;};})();',
].join('\n');

const 星の数 = (s) => (s.match(/★/g) || []).length;

/** 行の後ろに付いた注記( // …)を落とす。★http:// を消さない★／文字列の中の // は狙わない */
function 後ろの注記を落とす(line) {
  const i = line.indexOf('//');
  if (i < 0) return line;
  if (i > 0 && line[i - 1] === ':') return line;      // http:// など
  const 前 = line.slice(0, i);
  const q = (前.match(/'/g) || []).length + (前.match(/"/g) || []).length;
  if (q % 2 === 1) return line;                        // 文字列の中
  return 前;
}

/** 画面を本当に組み立ててから、見える所と 隠れている所を 分けて数える */
function 描いて数える(file) {
  const html = fs.readFileSync(path.join(ROOT, file), 'utf8');
  const dom = new JSDOM(html.replace(/<script[\s\S]*?<\/script>/g, ''), {
    runScripts: 'dangerously', url: 'http://localhost/', pretendToBeVisual: true,
    beforeParse(w) {
      w.fetch = () => Promise.reject(new Error('no net'));
      w.scrollTo = () => {}; w.alert = () => {};
      w.matchMedia = w.matchMedia || (() => ({ matches: false, addListener() {}, removeListener() {}, addEventListener() {}, removeEventListener() {} }));
      w.requestAnimationFrame = (cb) => setTimeout(() => cb(1), 0); w.cancelAnimationFrame = () => {};
      w.eval(CANVAS_STUB);
    },
  });
  const win = dom.window, doc = win.document;
  const inject = (c) => { const el = doc.createElement('script'); el.textContent = c; doc.body.appendChild(el); };
  const srcs = [...html.matchAll(/<script src="([^"]+)"><\/script>/g)].map((m) => m[1].split('?')[0]).filter((s) => !/^https?:/.test(s));
  for (const s of srcs) { const p = path.join(ROOT, s); if (fs.existsSync(p)) inject(fs.readFileSync(p, 'utf8')); }
  for (const m of html.matchAll(/<script(?![^>]*\ssrc=)[^>]*>([\s\S]*?)<\/script>/g)) inject(m[1]);
  try { doc.dispatchEvent(new win.Event('DOMContentLoaded', { bubbles: true })); } catch (e) { /* 続ける */ }
  try { win.dispatchEvent(new win.Event('load')); } catch (e) { /* 続ける */ }

  /* ★「見えているか」は 先祖をたどって display:none / hidden を見る★
     （offsetParent では 位置を固定した物が 漏れる＝別アプリで踏んだ穴） */
  const 隠れているか = (el) => {
    for (let n = el; n && n.nodeType === 1; n = n.parentElement) {
      if (n.hasAttribute && n.hasAttribute('hidden')) return true;
      const st = win.getComputedStyle(n);
      if (st && (st.display === 'none' || st.visibility === 'hidden')) return true;
    }
    return false;
  };

  const 見える = [], 隠れ = [];
  const walker = doc.createTreeWalker(doc.body, win.NodeFilter.SHOW_TEXT);
  for (let n = walker.nextNode(); n; n = walker.nextNode()) {
    const t = n.nodeValue || '';
    if (!t.includes('★')) continue;
    const 親 = n.parentElement;
    if (!親) continue;
    if (親.closest('script,style,template')) continue;
    (隠れているか(親) ? 隠れ : 見える).push({ 字: t.replace(/\s+/g, ' ').trim().slice(0, 70), 個: 星の数(t) });
  }
  /* ★window.close() は 走っている途中の描き直しを落とす（ここで落ちた）★ので閉じない */
  return { 見える, 隠れ };
}

/** JSが組み立てて 客に出す字（★注記は数えない★） */
function 出す字を数える() {
  const 見るファイル = [
    ...画面,
    ...fs.readdirSync(path.join(ROOT, 'lib')).filter((f) => f.endsWith('.js') && f !== 'xlsx.full.min.js').map((f) => 'lib/' + f),
    ...fs.readdirSync(path.join(ROOT, 'js')).filter((f) => f.endsWith('.js')).map((f) => 'js/' + f),
  ];
  const 出口 = /(showToast|_addPopupMsg|addAIChatMsg|\.textContent\s*=|\.innerHTML\s*=|placeholder\s*=|\.title\s*=)/;
  const 当たり = [];
  for (const f of 見るファイル) {
    const src = fs.readFileSync(path.join(ROOT, f), 'utf8')
      .replace(/\/\*[\s\S]*?\*\//g, '')          // ブロックの注記
      .replace(/^\s*\/\/.*$/gm, '')              // 行の注記
      .split(/\r?\n/).map(後ろの注記を落とす).join(String.fromCharCode(10))
      .replace(/<!--[\s\S]*?-->/g, '');          // HTMLの注記
    for (const line of src.split(/\r?\n/)) {
      if (!line.includes('★')) continue;
      if (!出口.test(line)) continue;
      当たり.push({ ファイル: f, 個: 星の数(line), 字: line.trim().slice(0, 90) });
    }
  }
  return 当たり;
}

const 表 = { _数え方: 'Ａ=今見えている / Ｂ=在るが今は出ていない / Ｃ=JSが組み立てて出す字（★下限★・文字列の中だけ）', 画面ごと: {} };
let A = 0, B = 0;
for (const f of 画面) {
  const r = 描いて数える(f);
  const a = r.見える.reduce((s, x) => s + x.個, 0);
  const b = r.隠れ.reduce((s, x) => s + x.個, 0);
  A += a; B += b;
  表.画面ごと[f] = { 'Ａ今見えている': a, 'Ｂ在るが出ていない': b, 見える中身: r.見える.slice(0, 12), 隠れ中身: r.隠れ.slice(0, 12) };
}
const c = 出す字を数える();
const C = c.reduce((s, x) => s + x.個, 0);
表.合計 = { 'Ａ今見えている': A, 'Ｂ在るが出ていない': B, 'Ｃ JSが出す字(下限)': C };
表['Ｃの内訳'] = c;

if (JSON_OUT) { console.log(JSON.stringify(表, null, 2)); process.exit(0); }

console.log('');
console.log('[count-stars] ★客に出る字の「★」を数える（直しません・数えるだけ）★');
for (const f of 画面) {
  const r = 表.画面ごと[f];
  console.log('  ' + f.padEnd(12) + ' Ａ今見えている ' + String(r['Ａ今見えている']).padStart(3)
    + ' ／ Ｂ在るが出ていない ' + String(r['Ｂ在るが出ていない']).padStart(3));
  for (const x of r.見える中身) console.log('      Ａ ' + x.字);
  for (const x of r.隠れ中身.slice(0, 6)) console.log('      Ｂ ' + x.字);
}
console.log('');
console.log('  ★Ｃ JSが組み立てて出す字（下限・注記は数えていない）= ' + C + '個 / ' + c.length + '行★');
for (const x of c) console.log('      ' + x.ファイル + ' … ' + x.字);
console.log('');
console.log('  ★合計 Ａ' + A + ' ／ Ｂ' + B + ' ／ Ｃ' + C + '（Ｃは下限）★');
process.exit(0);
