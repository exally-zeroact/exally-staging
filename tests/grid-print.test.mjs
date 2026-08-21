/* grid-print.test.mjs — ★印刷が 実Excelの既定どおりか／白紙を刷らないか★
 *
 *  真値は tests/fixtures/excel-print-golden.json（実Excel 16.0.20228 を COM で読んだ）。
 *  ★ソースを読むだけにしない★＝本物の book.html を載せて ★本物の printSheet を呼ぶ★。
 *
 *  ★一番 大事な所★
 *    ・★中身が0なら 印刷ダイアログを出さない★（白紙の窓が開いて「固まった」に見える事故）
 *    ・★紙だけの新しい窓で刷る★（画面のCSSを紙用に切り替えない＝全アプリの決まり）
 *    ・★式は 答えで刷る★（紙に「=A1+1」と出さない）
 *
 *  --self-test … わざと壊して「何通りで赤くなるか」を数える（★repo は読むだけ★）
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(__dirname, '..');
const SELF = process.argv.includes('--self-test');
const OVERRIDE = process.env.EXALLY_PRINT_OVERRIDE ? JSON.parse(process.env.EXALLY_PRINT_OVERRIDE) : {};
const srcPath = (rel) => OVERRIDE[rel] || path.join(ROOT, rel);

let JSDOM;
try { ({ JSDOM } = await import('jsdom')); }
catch { console.log('★jsdomが入っていません。この検証は飛ばせません（SKIPを緑と呼ばない）。npm install してください。'); process.exit(1); }

const GOLD = JSON.parse(fs.readFileSync(path.join(ROOT, 'tests/fixtures/excel-print-golden.json'), 'utf8'));

let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ok   ' + n); } catch (e) { fail++; console.log('  NG   ' + n + '\n       ' + (e && e.message)); } };
const ok = (c, m) => { if (!c) throw new Error(m || 'expected truthy'); };
const eq = (a, b, m) => { if (a !== b) throw new Error((m ? m + ': ' : '') + 'expected ' + JSON.stringify(b) + ' got ' + JSON.stringify(a)); };

const CANVAS_STUB = [
  '(function(){var noop=function(){};var ctx=new Proxy({},{get:function(t,k){',
  ' if(k==="measureText")return function(){return{width:40};};',
  ' if(k==="canvas")return{width:900,height:600};',
  ' if(k==="getImageData")return function(){return{data:[]};};',
  ' if(k==="createLinearGradient"||k==="createPattern")return function(){return{addColorStop:noop};};',
  ' return noop;}});HTMLCanvasElement.prototype.getContext=function(){return ctx;};})();',
].join('\n');

/* ★開いた窓と 刷った回数を数える偽物★（本物のブラウザの窓は出さない） */
const WINDOW_SPY = [
  'window.__開いた窓 = [];',
  'window.open = function(){',
  '  var 書いた = "";',
  '  var w = { __刷った: 0, focus: function(){}, print: function(){ this.__刷った++; },',
  '            document: { open:function(){}, close:function(){}, write:function(s){ 書いた += s; } },',
  '            get 中身(){ return 書いた; } };',
  '  window.__開いた窓.push(w); return w;',
  '};',
].join('\n');

const html = fs.readFileSync(srcPath('book.html'), 'utf8');
const dom = new JSDOM(html.replace(/<script[\s\S]*?<\/script>/g, ''), {
  runScripts: 'dangerously', url: 'http://localhost/', pretendToBeVisual: true,
  beforeParse(w) {
    w.fetch = () => Promise.reject(new Error('no net')); w.scrollTo = () => {}; w.alert = () => {};
    w.matchMedia = w.matchMedia || (() => ({ matches: false, addListener() {}, removeListener() {}, addEventListener() {}, removeEventListener() {} }));
    w.requestAnimationFrame = (cb) => setTimeout(() => cb(1), 0); w.cancelAnimationFrame = () => {};
    w.eval(CANVAS_STUB);
  },
});
const win = dom.window, doc = win.document;
const inject = (c) => { const el = doc.createElement('script'); el.textContent = c; doc.body.appendChild(el); };
const srcs = [...html.matchAll(/<script src="([^"]+)"><\/script>/g)].map((m) => m[1].split('?')[0]).filter((s) => !/^https?:/.test(s));
for (const s of srcs) { const p = srcPath(s); if (fs.existsSync(p)) inject(fs.readFileSync(p, 'utf8')); }
for (const m of html.matchAll(/<script(?![^>]*\ssrc=)[^>]*>([\s\S]*?)<\/script>/g)) inject(m[1]);
try { doc.dispatchEvent(new win.Event('DOMContentLoaded', { bubbles: true })); } catch (e) { /* 続ける */ }
try { win.dispatchEvent(new win.Event('load')); } catch (e) { /* 続ける */ }
inject(WINDOW_SPY);

console.log('\n[grid-print] ★本物の book.html で 実際に印刷を押す★');
console.log('  真値 = ' + GOLD._measured_with);

const sh = () => win.sheets[win.activeSheet];
function reset() { sh().data = {}; win.sel(0, 0, 0, 0); win.__開いた窓.length = 0; }
function toast(fn) {
  const 出た = []; const 元 = win.showToast;
  win.showToast = function (m) { 出た.push(String(m)); };
  try { fn(); } finally { win.showToast = 元; }
  return 出た;
}
const 最後の窓 = () => win.__開いた窓[win.__開いた窓.length - 1];

/* ── 土台 ── */
T('★画面が立ち上がっていて printSheet と GridPrint が在る', () => {
  ok(typeof win.printSheet === 'function', 'printSheet が無い');
  ok(win.GridPrint && typeof win.GridPrint.buildHtml === 'function', 'GridPrint が無い');
});
T('★Ctrl+P で 印刷が動く（実際にキーを送る）★', () => {
  reset();
  win.setCell(0, 0, 'あ');
  const ev = new win.KeyboardEvent('keydown', { key: 'p', ctrlKey: true, bubbles: true, cancelable: true });
  doc.dispatchEvent(ev);
  ok(ev.defaultPrevented, 'ブラウザの印刷を止めていない');
  eq(win.__開いた窓.length, 1, 'Ctrl+P で 紙の窓が開かない');
});
T('★右クリックにも 印刷が在る★', () => {
  const items = [...doc.querySelectorAll('.ctx-item')].map((e) => e.getAttribute('onclick') || '');
  ok(items.some((x) => x.startsWith('printSheet()')), '印刷の口が無い');
});

/* ── ① ★白紙を刷らない★（一番 大事） ── */
T('★中身が0なら 窓も開かず 印刷ダイアログも出さない（理由を出す）★', () => {
  reset();
  const 出た = toast(() => win.printSheet());
  eq(win.__開いた窓.length, 0, '★中身が無いのに 紙の窓を開いた＝白紙の印刷ダイアログ★');
  ok(出た.join('／').indexOf('中身') >= 0, '理由を言っていない: ' + 出た.join('／'));
});
T('★空白だけでも 刷らない★', () => {
  reset();
  sh().data['0,0'] = { v: '', f: '', d: '' };
  win.printSheet();
  eq(win.__開いた窓.length, 0, '空白だけなのに 窓を開いた');
});

/* ── ② ★紙だけの新しい窓で刷る★ ── */
T('★紙だけの新しい窓を開いて そこに書く（画面のCSSを紙用に切り替えない）★', () => {
  reset();
  win.setCell(0, 0, 'あ');
  win.printSheet();
  eq(win.__開いた窓.length, 1, '窓が開いていない');
  const 中身 = 最後の窓().中身;
  ok(中身.indexOf('<!doctype html>') === 0, '紙の窓に 丸ごとの HTML を書いていない');
  ok(中身.indexOf('<table>') >= 0, '表を書いていない');
  /* ★画面側に @media print を足していない事★（紙用のCSSに切り替える作りにしない） */
  ok(html.indexOf('@media print') < 0, '★画面のCSSに @media print が在る＝紙用に切り替える作りになっている★');
});

/* ── ③ 実Excel の既定どおりか ── */
T('★A4・縦・余白が 実Excel の既定と同じ★', () => {
  reset();
  win.setCell(0, 0, 'あ');
  win.printSheet();
  const 中身 = 最後の窓().中身;
  ok(/size: A4 portrait/.test(中身), 'A4縦になっていない');
  const cm = GOLD.既定.余白cm;
  ok(中身.indexOf('margin: ' + cm.上下 + 'cm ' + cm.左右 + 'cm') >= 0,
    '余白が実測（上下' + cm.上下 + 'cm・左右' + cm.左右 + 'cm）と違う');
});
T('★枠線を刷らない・行と列の番号も刷らない（実Excel の既定）★', () => {
  reset();
  win.setCell(0, 0, 'あ');
  win.setCell(0, 1, 'い');
  win.printSheet();
  const 中身 = 最後の窓().中身;
  eq(GOLD.既定.枠線を刷る, false, 'golden が違う');
  eq(GOLD.既定['行と列の番号を刷る'], false, 'golden が違う');
  ok(!/td,th \{[^}]*border:1px solid #BBBBBB/.test(中身), '★枠線を刷っている★');
  ok(中身.indexOf('>A</th>') < 0, '★列の番号(A)を刷っている★');
});

/* ── ④ 何を刷るか ── */
T('★式は 答えで刷る（紙に =A1+1 と出さない）★', () => {
  reset();
  win.setCell(0, 0, '2');
  win.setCell(1, 0, '=A1+1');
  win.printSheet();
  const 中身 = 最後の窓().中身;
  ok(中身.indexOf('>3<') >= 0, '答えが出ていない');
  ok(中身.indexOf('=A1+1') < 0, '★紙に 式が出ている★');
});
T('★まだ計算していない式のセルは 紙に式を出さない★', () => {
  /* ★2026-08-21 実際に踏んだ★:
     打ち込んだ式のセルは 答え(d)が入っているので、「式をそのまま出す」壊し方が
     ★結果が同じになり 素通りした★。
     ★答えがまだ入っていない形（v が式・d が空）★ で 別に確かめる。 */
  eq(win.GridPrint.shown({ v: '=A1+1', f: '=A1+1', d: '' }), '',
    '★答えが無い式を 紙に そのまま出している★');
  eq(win.GridPrint.shown({ v: '=A1+1', f: '=A1+1', d: '3' }), '3', '答えが在る時は 答えを出す');
});
T('★選んだ範囲が 空っぽなら 刷らない（白紙を刷らない）★', () => {
  /* ★2026-08-21 実際に踏んだ★:
     中身が0の時は「中身が在る所」を探す所で止まるので、
     ★「字が1つも無いなら刷らない」を消しても 素通りした★。
     ★範囲を選んでいて その中が空っぽ★ の時だけ 通る道なので そこを確かめる。 */
  reset();
  win.setCell(0, 0, 'あ');          /* 別の所には中身が在る */
  win.sel(3, 3, 5, 5);              /* ★空っぽの所を選ぶ★ */
  const 出た = toast(() => win.printSheet());
  eq(win.__開いた窓.length, 0, '★空っぽの範囲なのに 紙の窓を開いた＝白紙★');
  ok(出た.join('／').indexOf('中身') >= 0, '理由を言っていない: ' + 出た.join('／'));
});
T('★2セル以上 選んでいれば その範囲だけ刷る★', () => {
  reset();
  win.setCell(0, 0, 'これは出る');
  win.setCell(5, 0, 'これは出ない');
  win.sel(0, 0, 1, 0);
  win.printSheet();
  const 中身 = 最後の窓().中身;
  ok(中身.indexOf('これは出る') >= 0, '選んだ所が出ていない');
  ok(中身.indexOf('これは出ない') < 0, '★選んでいない所まで刷った★');
});
T('★1セルだけの時は 中身が在る所を ぜんぶ刷る★', () => {
  reset();
  win.setCell(0, 0, '上');
  win.setCell(5, 2, '下');
  win.sel(0, 0, 0, 0);
  win.printSheet();
  const 中身 = 最後の窓().中身;
  ok(中身.indexOf('上') >= 0 && 中身.indexOf('下') >= 0, '全部 出ていない');
});
T('★書式（太字・色・寄せ）が 紙にも出る★', () => {
  reset();
  win.setCell(0, 0, 'あ');
  Object.assign(sh().data['0,0'], { bold: true, color: '#B04A3A', align: 'center' });
  win.printSheet();
  const 中身 = 最後の窓().中身;
  ok(中身.indexOf('font-weight:700') >= 0, '太字が出ていない');
  ok(中身.indexOf('#B04A3A') >= 0, '色が出ていない');
  ok(中身.indexOf('text-align:center') >= 0, '寄せが出ていない');
});
T('★字が欠けない（折り返して全部 出す）★＝Excel との違い', () => {
  reset();
  win.setCell(0, 0, 'とても長い字をここに入れて紙で欠けないかを見る');
  win.printSheet();
  const 中身 = 最後の窓().中身;
  ok(中身.indexOf('とても長い字をここに入れて紙で欠けないかを見る') >= 0, '長い字が切られている');
  ok(中身.indexOf('word-break:break-all') >= 0, '折り返す作りになっていない');
  ok(GOLD['★うちの方が良い所（違いとして残す）★'].字が欠けない, '違いが golden に書かれていない');
});

/* ── ⑤ ★真似られない事／作っていない物★ ── */
T('★「1ページの行数は 機械で変わる＝真似ない」が golden に書いてある★', () => {
  const x = GOLD['★真似られない事★'];
  ok(x && x['1ページの行数'], '真似られない事が書かれていない');
  eq(x['1ページの行数'].実測, 39);
  eq(x['1ページの行数'].紙から計算すると, 41);
  ok(String(x['1ページの行数'].なぜ違うか).length > 10, '理由が書かれていない');
});
T('★作っていない物が golden に書いてある（全部 作った事にしない）★', () => {
  ok(Array.isArray(GOLD['作っていない物']) && GOLD['作っていない物'].length >= 4, '作っていない物が書かれていない');
  ok(String(GOLD._how).length > 20, 'どう測ったかが書かれていない');
});

console.log('\n  ' + pass + ' passed, ' + fail + ' failed');

if (SELF) {
  const { spawnSync } = await import('node:child_process');
  const os = await import('node:os');
  const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'exally-print-'));
  console.log('\n[self-test] わざと壊して 赤くなるかを数える（★repo は読むだけ★）');
  const BREAKS = [
    ['book.html', 'Ctrl+P の割り当てを消す', (s) => s.replace("    if(ek==='p'){ e.preventDefault(); printSheet(); return; }\n", '')],
    ['book.html', '右クリックの 印刷を消す', (s) => s.replace('<div class="ctx-item" onclick="printSheet()">🖨️ 印刷（Ctrl+P）</div>', '')],
    ['book.html', '★中身が0でも 窓を開く（白紙の印刷ダイアログ）★', (s) => s.replace("  if(!html){ showToast('刷る中身が在りません（先に何か入れてね）'); return; }", '  if(!html){ html = "<html></html>"; }')],
    ['book.html', '選んだ範囲を見ない（いつも全部 刷る）', (s) => s.replace("  var 範囲 = (selR1!==selR2 || selC1!==selC2)\n    ? { r1:selR1, c1:selC1, r2:selR2, c2:selC2 } : null;", '  var 範囲 = null;')],
    ['lib/grid-print.js', '★A4横にする（実測は縦）★', (s) => s.replace("    var 縦 = (o.向き 　!== 'landscape');", '    var 縦 = false;')],
    ['lib/grid-print.js', '★余白を変える★', (s) => s.replace("'@page { size: A4 ' + (縦 ? 'portrait' : 'landscape') + '; margin: 1.9cm 1.78cm; }',", "'@page { size: A4 ' + (縦 ? 'portrait' : 'landscape') + '; margin: 1cm; }',")],
    ['lib/grid-print.js', '★枠線を刷ってしまう★', (s) => s.replace('    var 枠線 = !!o.枠線;                          /* ★既定は刷らない（実測）★ */', '    var 枠線 = true;')],
    ['lib/grid-print.js', '★行と列の番号を刷ってしまう★', (s) => s.replace('    var 行列番号 = !!o.行列番号;                  /* ★既定は刷らない（実測）★ */', '    var 行列番号 = true;')],
    ['lib/grid-print.js', '★紙に 式をそのまま出す★', (s) => s.replace("    if (s.charAt(0) === '=') return '';", '')],
    ['lib/grid-print.js', '★中身が無くても HTML を返す（白紙を刷る）★', (s) => s.replace('    if (!何か出た) return null;                  /* ★字が1つも無いなら 刷らない★ */', '')],
    ['lib/grid-print.js', '★長い字を折り返さない（紙で欠ける）★', (s) => s.replace('word-break:break-all;', '')],
  ];
  let red = 0;
  for (const [rel, name, brk] of BREAKS) {
    const orig = fs.readFileSync(path.join(ROOT, rel), 'utf8');
    const bad = brk(orig);
    if (bad === orig) { console.log('  ★置換できず★  ' + name); continue; }
    const tmpFile = path.join(TMP, rel.replace(/[\\/]/g, '_'));
    fs.writeFileSync(tmpFile, bad, 'utf8');
    const env = Object.assign({}, process.env, { EXALLY_PRINT_OVERRIDE: JSON.stringify({ [rel]: tmpFile }) });
    const isRed = spawnSync(process.execPath, [path.join(__dirname, 'grid-print.test.mjs')], { encoding: 'utf8', env }).status !== 0;
    if (isRed) { red++; console.log('  赤くなった  ' + name); }
    else console.log('  ★素通り★  ' + name);
  }
  for (const rel of ['book.html', 'lib/grid-print.js']) {
    const now = fs.readFileSync(path.join(ROOT, rel), 'utf8');
    if (now.includes('var 縦 = false;') || now.includes('var 枠線 = true;') || now.includes('html = "<html></html>"')) {
      console.log('  ★NG★ ' + rel + ' に わざと壊した物が残っている');
      process.exit(1);
    }
  }
  try { fs.rmSync(TMP, { recursive: true, force: true }); } catch (e) { /* 消せなくても検査は済んでいる */ }
  console.log('\n  ' + red + '/' + BREAKS.length + ' 通りで赤くなった');
  process.exit(red === BREAKS.length ? 0 : 1);
}

process.exit(fail ? 1 : 0);
