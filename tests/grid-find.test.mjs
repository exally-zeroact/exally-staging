/* grid-find.test.mjs — ★検索と置換が 実Excelと同じに動くか★
 *
 *  真値は tests/fixtures/excel-find-golden.json（実Excel 16.0.20228 を COM で動かして取った）。
 *  ★ソースを読むだけにしない★＝本物の book.html を載せて ★本物の replaceAll / findNextCell を呼ぶ★。
 *
 *  ★一番 大事な所★
 *    ・★置換は「式」を見る★＝答えが変わる（実測①⑦）
 *    ・だから ★式を何本 書き換えたかを 必ず知らせる★（黙って値を変えない）
 *    ・★取り消し1回で丸ごと戻る★
 *
 *  --self-test … わざと壊して「何通りで赤くなるか」を数える（★repo は読むだけ★）
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(__dirname, '..');
const SELF = process.argv.includes('--self-test');
const OVERRIDE = process.env.EXALLY_FIND_OVERRIDE ? JSON.parse(process.env.EXALLY_FIND_OVERRIDE) : {};
const srcPath = (rel) => OVERRIDE[rel] || path.join(ROOT, rel);

let JSDOM;
try { ({ JSDOM } = await import('jsdom')); }
catch { console.log('★jsdomが入っていません。この検証は飛ばせません（SKIPを緑と呼ばない）。npm install してください。'); process.exit(1); }

const GOLD = JSON.parse(fs.readFileSync(path.join(ROOT, 'tests/fixtures/excel-find-golden.json'), 'utf8'));

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

console.log('\n[grid-find] ★本物の book.html で 実際に探して置き換える★');
console.log('  真値 = ' + GOLD._measured_with);

const sh = () => win.sheets[win.activeSheet];
const A = (r, c) => sh().data[r + ',' + c] || null;
const 出る = (r, c) => { const x = A(r, c); if (!x) return ''; const v = (x.d !== undefined && x.d !== null && x.d !== '') ? x.d : x.v; return (v === undefined || v === null) ? '' : String(v); };
function reset() { sh().data = {}; sh().freezeRow = 0; sh().freezeCol = 0; win.sel(0, 0, 0, 0); }
function 窓(what, to, opts) {
  /* ★知らせを毎回 空にしてから始める★（2026-08-21 実際に踏んだ）
     前の検査が残した「見つかりません」が そのまま残っていて、
     ★「黙る」壊し方が 素通りした★。★前の字を見て 緑にしない★ */
  doc.getElementById('findMsg').textContent = '';
  doc.getElementById('findWhat').value = what;
  doc.getElementById('findTo').value = to === undefined ? '' : to;
  doc.getElementById('findCase').checked = !!(opts && opts.matchCase);
  doc.getElementById('findWhole').checked = !!(opts && opts.whole);
}
const 知らせ = () => doc.getElementById('findMsg').textContent;

/* ── 土台 ── */
T('★画面が立ち上がっていて openFind / replaceAll / GridFind が在る', () => {
  ok(typeof win.openFind === 'function', 'openFind が無い');
  ok(typeof win.findNextCell === 'function', 'findNextCell が無い');
  ok(typeof win.replaceOne === 'function', 'replaceOne が無い');
  ok(typeof win.replaceAll === 'function', 'replaceAll が無い');
  ok(win.GridFind && typeof win.GridFind.replacedText === 'function', 'GridFind が無い');
});
T('★Ctrl+F / Ctrl+H で 窓が開く（実際にキーを送る）★', () => {
  const key = (k) => { const ev = new win.KeyboardEvent('keydown', { key: k, ctrlKey: true, bubbles: true, cancelable: true }); doc.dispatchEvent(ev); return ev; };
  const ov = doc.getElementById('findOverlay');
  ov.style.display = 'none';
  const e1 = key('f');
  ok(e1.defaultPrevented, 'Ctrl+F が 既定を止めていない');
  eq(ov.style.display, 'flex', 'Ctrl+F で窓が開かない');
  eq(doc.getElementById('findReplaceRow').style.display, 'none', '検索なのに 置換の欄が出ている');
  win.closeFind();
  const e2 = key('h');
  ok(e2.defaultPrevented, 'Ctrl+H が 既定を止めていない');
  eq(ov.style.display, 'flex', 'Ctrl+H で窓が開かない');
  eq(doc.getElementById('findReplaceRow').style.display, 'flex', '置換なのに 置換の欄が無い');
  win.closeFind();
});

/* ── ① ★置換は式を見る★ ── */
T('★置換は「式」を見る（=A1*2 が =A1*9 になり 答えが 81）★＝実Excelで実測', () => {
  reset();
  win.setCell(0, 0, '2');
  win.setCell(0, 1, '=A1*2');
  win.setCell(0, 2, 'りんご2個');
  const g = GOLD['①式を見るのか答えを見るのか'].結果;
  窓('2', '9');
  win.replaceAll();
  eq(A(0, 0).v, g.A1式, 'A1');
  eq(A(0, 1).f, g.B1式, 'B1の式');
  eq(出る(0, 1), g.B1答え, 'B1の答え');
  eq(A(0, 2).v, g.C1式, 'C1');
});

T('★読み込んだファイルの形（値が空・式だけ在る）でも 式を見る★', () => {
  /* ★2026-08-21 実際に踏んだ★:
     打ち込んだセルは v にも式が入るので、「式ではなく値を見る」壊し方が
     ★結果が同じになり 素通りした★。
     ファイルから読んだセルは ★v が空で f に式★（js/book-open.js の sheetToGrid）
     なので、その形でも見られるかを 別に確かめる。 */
  reset();
  sh().data['0,0'] = { v: 2, f: '', d: 2 };
  sh().data['0,1'] = { v: '', f: '=A1*2', d: 4 };      /* ★ファイルから読んだ形★ */
  ok(win.GridFind.match(A(0, 1), '2', {}), '★式の中の 2 を見つけられない＝答えを見ている★');
  eq(win.GridFind.replacedText(A(0, 1), '2', '9', {}), '=A1*9');
});

/* ── ⑦ ★一番 危ない所★ ── */
T('★=A1+100 に 10→20 で =A1+200（答え 110→220）★＝実Excelで実測', () => {
  reset();
  win.setCell(0, 0, '10');
  win.setCell(1, 0, '=A1+100');
  eq(出る(1, 0), '110', '材料が違う');
  const g = GOLD['⑦★一番 危ない所★'].結果;
  窓('10', '20');
  win.replaceAll();
  eq(A(0, 0).v, g.A1, 'A1');
  eq(A(1, 0).f, g.A2式, 'A2の式');
  eq(出る(1, 0), g.A2答え, 'A2の答え');
});
T('★式を何本 書き換えたかを 必ず知らせる★（黙って値を変えない）', () => {
  reset();
  win.setCell(0, 0, '10');
  win.setCell(1, 0, '=A1+100');
  窓('10', '20');
  win.replaceAll();
  const m = 知らせ();
  ok(m.indexOf('2か所') >= 0, '何か所 置き換えたか 出していない: ' + m);
  ok(m.indexOf('式') >= 0 && m.indexOf('1本') >= 0, '★式が何本かを 出していない★: ' + m);
  ok(m.indexOf('Ctrl+Z') >= 0, '戻せる事を言っていない: ' + m);
});
T('★式が無い時は 式の話をしない（余計な脅かしをしない）★', () => {
  reset();
  win.setCell(0, 0, 'あい');
  win.setCell(1, 0, 'あいうえお');
  窓('あい', 'X');
  win.replaceAll();
  const m = 知らせ();
  ok(m.indexOf('2か所') >= 0, m);
  ok(m.indexOf('式が') < 0, '式が無いのに 式の話をした: ' + m);
});

/* ── ② 大文字と小文字 ── */
T('大文字と小文字 … ★既定は 区別しない★＝実Excelで実測', () => {
  reset();
  GOLD['②大文字と小文字'].既定.入れた物.forEach((v, i) => win.setCell(i, 0, v));
  窓('abc', 'X');
  win.replaceAll();
  eq([0, 1, 2].map((i) => A(i, 0).v).join('|'), GOLD['②大文字と小文字'].既定.結果.join('|'));
});
T('大文字と小文字 … ★区別する にすると abc だけ★＝実Excelで実測', () => {
  reset();
  GOLD['②大文字と小文字'].区別する.入れた物.forEach((v, i) => win.setCell(i, 0, v));
  窓('abc', 'X', { matchCase: true });
  win.replaceAll();
  eq([0, 1, 2].map((i) => A(i, 0).v).join('|'), GOLD['②大文字と小文字'].区別する.結果.join('|'));
});

/* ── ③ 一部か全体か ── */
T('一部か全体か … ★既定は 一部でも置き換える★＝実Excelで実測', () => {
  reset();
  GOLD['③一部か全体か'].既定.入れた物.forEach((v, i) => win.setCell(i, 0, v));
  窓('あい', 'X');
  win.replaceAll();
  eq([0, 1].map((i) => A(i, 0).v).join('|'), GOLD['③一部か全体か'].既定.結果.join('|'));
});
T('一部か全体か … ★セル全体だけ にすると 短い方だけ★＝実Excelで実測', () => {
  reset();
  GOLD['③一部か全体か'].セル全体だけ.入れた物.forEach((v, i) => win.setCell(i, 0, v));
  窓('あい', 'X', { whole: true });
  win.replaceAll();
  eq([0, 1].map((i) => A(i, 0).v).join('|'), GOLD['③一部か全体か'].セル全体だけ.結果.join('|'));
});

/* ── ④ 全角と半角 ── */
T('★全角と半角は 別物★＝実Excelで実測', () => {
  reset();
  /* ★setCell は打った字を半角に直すので ここは直接 置く★（読み込んだファイルと同じ形） */
  sh().data['0,0'] = { v: 'ABC', f: 'ABC', d: 'ABC' };
  sh().data['1,0'] = { v: 'ＡＢＣ', f: 'ＡＢＣ', d: 'ＡＢＣ' };
  窓('ABC', 'X');
  win.replaceAll();
  eq([0, 1].map((i) => A(i, 0).v).join('|'), GOLD['④全角と半角'].結果.join('|'));
});

/* ── ⑤ ワイルドカード ── */
for (const [pat, cs] of Object.entries(GOLD['⑤ワイルドカード'])) {
  T('ワイルドカード … 「' + pat + '」＝実Excelで実測（' + cs.意味 + '）', () => {
    reset();
    cs.入れた物.forEach((v, i) => win.setCell(i, 0, v));
    窓(pat, 'X');
    win.replaceAll();
    eq(cs.入れた物.map((_, i) => A(i, 0).v).join('|'), cs.結果.join('|'));
  });
}

/* ── ⑥ 次を探す順 ── */
T('★次を探すのは「選んだセルの次」から★＝実Excelで実測（A1を選んで探すと B1）', () => {
  reset();
  win.setCell(0, 0, 'x'); win.setCell(0, 1, 'x'); win.setCell(1, 0, 'x');
  win.sel(0, 0, 0, 0);
  窓('x', '');
  const 見つけた = win.findNextCell();
  ok(見つけた, '見つからなかった');
  eq(win.colLetter(見つけた.c) + (見つけた.r + 1), GOLD['⑥次を探す順'].A1を選んで探す);
  eq(win.selR1, 見つけた.r, '選び直していない（行）');
  eq(win.selC1, 見つけた.c, '選び直していない（列）');
});
T('★見つからない時は そう言う（黙って何もしない、にしない）★', () => {
  reset();
  win.setCell(0, 0, 'あ');
  win.sel(0, 0, 0, 0);
  窓('ない言葉', '');
  const r = win.findNextCell();
  eq(r, null);
  ok(知らせ().indexOf('見つかりません') >= 0, '理由を言っていない: ' + 知らせ());
});
T('★すべて置き換えで 見つからない時も そう言う（黙らない）★', () => {
  reset();
  win.setCell(0, 0, 'あ');
  窓('ない言葉', 'X');
  win.replaceAll();
  eq(A(0, 0).v, 'あ', '無いのに置き換えた');
  ok(知らせ().indexOf('見つかりません') >= 0, '理由を言っていない: ' + 知らせ());
});
T('★探す言葉が空なら 何もせず そう言う★', () => {
  reset();
  win.setCell(0, 0, 'あ');
  窓('', 'X');
  win.replaceAll();
  eq(A(0, 0).v, 'あ', '空なのに置き換えた');
  ok(知らせ().indexOf('探す言葉') >= 0, '理由を言っていない: ' + 知らせ());
});

/* ── 取り消し ── */
T('★すべて置き換えても 取り消し1回で丸ごと戻る★', () => {
  reset();
  for (let r = 0; r < 5; r++) win.setCell(r, 0, 'あ' + r);
  窓('あ', 'X');
  win.replaceAll();
  eq([0, 1, 2, 3, 4].map((i) => A(i, 0).v).join('|'), 'X0|X1|X2|X3|X4');
  win.doUndo();
  eq([0, 1, 2, 3, 4].map((i) => A(i, 0).v).join('|'), 'あ0|あ1|あ2|あ3|あ4', '1回で戻っていない');
});

/* ── 作っていない物の台帳 ── */
T('★作っていない物が golden に書いてある（全部 作った事にしない）★', () => {
  ok(Array.isArray(GOLD['作っていない物']) && GOLD['作っていない物'].length >= 3, '作っていない物が書かれていない');
  ok(String(GOLD._how).length > 20, 'どう測ったかが書かれていない');
});

console.log('\n  ' + pass + ' passed, ' + fail + ' failed');

if (SELF) {
  const { spawnSync } = await import('node:child_process');
  const os = await import('node:os');
  const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'exally-find-'));
  console.log('\n[self-test] わざと壊して 赤くなるかを数える（★repo は読むだけ★）');
  const BREAKS = [
    ['book.html', 'Ctrl+F の割り当てを消す', (s) => s.replace("    if(ek==='f'){ e.preventDefault(); openFind(false); return; }\n", '')],
    ['book.html', 'Ctrl+H の割り当てを消す', (s) => s.replace("    if(ek==='h'){ e.preventDefault(); openFind(true); return; }\n", '')],
    ['book.html', '★式を何本 書き換えたかを 出さない★', (s) => s.replace("    + (式の数 ? '／うち 式が ' + 式の数 + '本＝答えが変わります' : '')", '')],
    ['book.html', '★取り消しの控えを取らない★', (s) => s.replace('  _pushRowColUndo();                       /* ★取り消し1回で丸ごと戻る★ */', '')],
    ['book.html', '見つからない時に 黙る', (s) => s.replace("  if(!直す.length){ _findMsg('見つかりません'); return; }", '  if(!直す.length){ return; }')],
    ['book.html', '探す言葉が空でも やってしまう', (s) => s.replace("  if(!what){ _findMsg('探す言葉を入れてね'); return; }\n  var opts = _findOpts();", '  var opts = _findOpts();')],
    ['lib/grid-find.js', '★式ではなく 答えを見る★（Excelと違う）', (s) => s.replace(
      "    if (typeof cell.f === 'string' && cell.f.charAt(0) === '=') return cell.f;",
      "    if (typeof cell.f === 'string' && cell.f.charAt(0) === '=') return String(cell.d === undefined ? '' : cell.d);")],
    ['lib/grid-find.js', '★大文字と小文字を いつも区別する★', (s) => s.replace("return new RegExp(out, opts.matchCase ? 'g' : 'gi');", "return new RegExp(out, 'g');")],
    ['lib/grid-find.js', '★セル全体だけ が効かない★', (s) => s.replace("    if (opts.whole) out = '^' + out + '$';", '')],
    ['lib/grid-find.js', '★~* が効かない（* そのものを探せない）★', (s) => s.replace("      if (ch === '~') {", '      if (false) {')],
    ['lib/grid-find.js', '★? が1文字にならない★', (s) => s.replace("      if (ch === '?') { out += '[\\\\s\\\\S]'; continue; }    /* 1文字 */", '')],
    ['lib/grid-find.js', '★次からではなく 今いる所から探す★', (s) => s.replace('    var start = (今.r * cols + 今.c + 1) % 全部;      /* ★次から★ */', '    var start = (今.r * cols + 今.c) % 全部;')],
  ];
  let red = 0;
  for (const [rel, name, brk] of BREAKS) {
    const orig = fs.readFileSync(path.join(ROOT, rel), 'utf8');
    const bad = brk(orig);
    if (bad === orig) { console.log('  ★置換できず★  ' + name); continue; }
    const tmpFile = path.join(TMP, rel.replace(/[\\/]/g, '_'));
    fs.writeFileSync(tmpFile, bad, 'utf8');
    const env = Object.assign({}, process.env, { EXALLY_FIND_OVERRIDE: JSON.stringify({ [rel]: tmpFile }) });
    const isRed = spawnSync(process.execPath, [path.join(__dirname, 'grid-find.test.mjs')], { encoding: 'utf8', env }).status !== 0;
    if (isRed) { red++; console.log('  赤くなった  ' + name); }
    else console.log('  ★素通り★  ' + name);
  }
  for (const rel of ['book.html', 'lib/grid-find.js']) {
    const now = fs.readFileSync(path.join(ROOT, rel), 'utf8');
    if (now.includes("new RegExp(out, 'g');") || now.includes('if (false) {')) {
      console.log('  ★NG★ ' + rel + ' に わざと壊した物が残っている');
      process.exit(1);
    }
  }
  try { fs.rmSync(TMP, { recursive: true, force: true }); } catch (e) { /* 消せなくても検査は済んでいる */ }
  console.log('\n  ' + red + '/' + BREAKS.length + ' 通りで赤くなった');
  process.exit(red === BREAKS.length ? 0 : 1);
}

process.exit(fail ? 1 : 0);
