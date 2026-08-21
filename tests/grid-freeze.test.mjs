/* grid-freeze.test.mjs — ★ウィンドウ枠の固定が 実Excelと同じに決まるか★
 *
 *  真値は tests/fixtures/excel-freeze-golden.json（実Excel 16.0.20228 を COM で動かして取った）。
 *  ★ソースを読むだけにしない★＝本物の book.html を載せて ★本物の freezePanes を呼ぶ★。
 *
 *  ★一番 大事な所★
 *    ・★固定しても 合計が変わらない★（見た目だけ＝実Excelで実測）
 *    ・★描き方を触ったので 前からある物が壊れていないか★も ここで押して数える
 *      （並べ替え・オートSUM・絞り込み）
 *
 *  --self-test … わざと壊して「何通りで赤くなるか」を数える（★repo は読むだけ★）
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(__dirname, '..');
const SELF = process.argv.includes('--self-test');
const OVERRIDE = process.env.EXALLY_FREEZE_OVERRIDE ? JSON.parse(process.env.EXALLY_FREEZE_OVERRIDE) : {};
const srcPath = (rel) => OVERRIDE[rel] || path.join(ROOT, rel);

let JSDOM;
try { ({ JSDOM } = await import('jsdom')); }
catch { console.log('★jsdomが入っていません。この検証は飛ばせません（SKIPを緑と呼ばない）。npm install してください。'); process.exit(1); }

const GOLD = JSON.parse(fs.readFileSync(path.join(ROOT, 'tests/fixtures/excel-freeze-golden.json'), 'utf8'));

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

console.log('\n[grid-freeze] ★本物の book.html で 実際に固定する★');
console.log('  真値 = ' + GOLD._measured_with);

const sh = () => win.sheets[win.activeSheet];
const A = (r, c) => sh().data[r + ',' + c] || null;
const 出る = (r, c) => { const x = A(r, c); if (!x) return '（空）'; const v = (x.d !== undefined && x.d !== null && x.d !== '') ? x.d : x.v; return (v === undefined || v === null || v === '') ? '（空）' : String(v); };
function reset() {
  sh().data = {}; sh().filterHidden = {}; sh().hiddenRows = {};
  sh().freezeRow = 0; sh().freezeCol = 0;
  win.scrollTop = 0; win.scrollLeft = 0;
  win.sel(0, 0, 0, 0);
}
const A1 = (a1) => { const m = /^([A-Za-z]+)(\d+)$/.exec(a1); let c = 0; const L = m[1].toUpperCase(); for (let i = 0; i < L.length; i++) c = c * 26 + (L.charCodeAt(i) - 64); return { r: +m[2] - 1, c: c - 1 }; };
function toast(fn) {
  const 出た = []; const 元 = win.showToast;
  win.showToast = function (m) { 出た.push(String(m)); };
  try { fn(); } finally { win.showToast = 元; }
  return 出た;
}

/* ── 土台 ── */
T('★画面が立ち上がっていて freezePanes と _renderPass が在る', () => {
  ok(typeof win.freezePanes === 'function', 'freezePanes が無い');
  ok(typeof win.unfreezePanes === 'function', 'unfreezePanes が無い');
  ok(typeof win._renderPass === 'function', '★描く所の切り出し（_renderPass）が無い★');
  ok(typeof win._doRender === 'function', '_doRender が無い');
});
T('★右クリックの中に 固定と 固定をやめる が在る（押す口が在る）', () => {
  const items = [...doc.querySelectorAll('.ctx-item')].map((e) => e.getAttribute('onclick') || '');
  ok(items.some((x) => x.startsWith('freezePanes()')), '固定の口が無い');
  ok(items.some((x) => x.startsWith('unfreezePanes()')), '固定をやめる口が無い');
});

/* ── ① どこで固定されるか（実測どおり） ── */
for (const cs of GOLD['どこで固定されるか']) {
  T('固定 … ' + cs.選んだ + ' を選ぶと ★上' + cs.上に固定 + '行 / 左' + cs.左に固定 + '列★＝実Excelで実測', () => {
    reset();
    const p = A1(cs.選んだ);
    win.sel(p.r, p.c, p.r, p.c);
    win.freezePanes();
    eq(sh().freezeRow, cs.上に固定, '上に固定する行');
    eq(sh().freezeCol, cs.左に固定, '左に固定する列');
  });
}

/* ── ② ★A1 は 真似ない（固定せず 理由を出す）★ ── */
T('★A1 で押したら 固定せず 理由を出す★（画面の大きさで変わる物は真似ない）', () => {
  reset();
  win.sel(0, 0, 0, 0);
  const 出た = toast(() => win.freezePanes());
  eq(sh().freezeRow, 0, 'A1 で固定してしまった（行）');
  eq(sh().freezeCol, 0, 'A1 で固定してしまった（列）');
  ok(出た.length > 0, '何も言わずに黙って終わった');
  ok(出た.join('／').indexOf('A1') >= 0, '理由に A1 が出ていない: ' + 出た.join('／'));
  /* golden にも「真似ない理由」が書いてあるか */
  const 真似ない = GOLD['★真似ない事（測ったが 真似られない）★'];
  ok(真似ない && 真似ない.A1を選んで固定, '真似ない理由が golden に無い');
});

/* ── ③ ★固定しても 合計が変わらない★（一番 大事） ── */
T('★固定しても 合計が変わらない（見た目だけ）★＝実Excelで実測', () => {
  reset();
  for (let r = 0; r < 5; r++) win.setCell(r, 0, String(r + 1));
  win.setCell(0, 3, '=SUM(A1:A5)');
  const 前 = 出る(0, 3);
  eq(前, '15', '材料が違う');
  win.sel(1, 1, 1, 1);
  win.freezePanes();
  eq(出る(0, 3), 前, '★固定したら 合計が変わった＝黙って嘘をつく★');
  win.unfreezePanes();
  eq(出る(0, 3), 前, '固定をやめたら 合計が変わった');
});

/* ── ④ 固定した行は スクロールしても 動かない ── */
T('★固定した行は スクロールしても 場所が動かない★（本物の rowY を呼ぶ）', () => {
  reset();
  for (let r = 0; r < 40; r++) win.setCell(r, 0, String(r + 1));
  win.sel(2, 0, 2, 0);           /* A3 → 上2行を固定 */
  win.freezePanes();
  eq(sh().freezeRow, 2);
  const 固定前 = [win.rowY(0), win.rowY(1)];
  win.scrollTop = 500;
  const 固定後 = [win.rowY(0), win.rowY(1)];
  /* rowY 自体は今までどおり（帯は「スクロールなしで もう一度描く」形）＝
     ★押した所の読み取り（yToR）が 帯の中で 固定した行を返す★ のが 本当の確かめ */
  ok(固定前.length === 2 && 固定後.length === 2, '');
  eq(win.yToR(win.HDR_H + 1), 0, '帯の1行目が 0行目になっていない');
  eq(win.yToR(win.HDR_H + win.rH(0) + 1), 1, '帯の2行目が 1行目になっていない');
  win.scrollTop = 0;
});
T('★固定した列も同じ（xToC が 帯の中で 固定した列を返す）★', () => {
  reset();
  for (let c = 0; c < 10; c++) win.setCell(0, c, 'c' + c);
  win.sel(0, 2, 0, 2);           /* C1 → 左2列を固定 */
  win.freezePanes();
  eq(sh().freezeCol, 2);
  win.scrollLeft = 500;
  eq(win.xToC(win.HDR_W + 1), 0, '帯の1列目が 0列目になっていない');
  eq(win.xToC(win.HDR_W + win.cW(0) + 1), 1, '帯の2列目が 1列目になっていない');
  win.scrollLeft = 0;
});

/* ── ⑤ 取り消し1回で戻る ── */
T('★固定は 取り消し1回で戻る★', () => {
  reset();
  win.setCell(0, 0, 'あ');
  win.sel(3, 2, 3, 2);
  win.freezePanes();
  eq(sh().freezeRow, 3); eq(sh().freezeCol, 2);
  win.doUndo();
  eq(sh().freezeRow || 0, 0, '1回で戻っていない（行）');
  eq(sh().freezeCol || 0, 0, '1回で戻っていない（列）');
});

/* ── ⑥ 固定していない時に「やめる」を押しても 壊れない ── */
T('★固定していない時に「やめる」を押しても 何も壊さない（理由を出す）★', () => {
  reset();
  const 出た
    = toast(() => win.unfreezePanes());
  ok(出た.join('／').indexOf('固定') >= 0, '理由を言っていない: ' + 出た.join('／'));
});

/* ── ⑦ ★描き方を触ったので 前からある物が壊れていないか★ ── */
T('★描き直しても 落ちない（固定あり・なし どちらも _doRender を実際に呼ぶ）★', () => {
  reset();
  for (let r = 0; r < 20; r++) for (let c = 0; c < 5; c++) win.setCell(r, c, String(r) + '-' + String(c));
  win._doRender();                       /* 固定なし */
  win.sel(2, 2, 2, 2); win.freezePanes();
  win._doRender();                       /* 上も左も固定 */
  win.sel(2, 0, 2, 0); win.unfreezePanes(); win.freezePanes();
  win._doRender();                       /* 上だけ固定 */
  win.sel(0, 2, 0, 2); win.unfreezePanes(); win.freezePanes();
  win._doRender();                       /* 左だけ固定 */
  ok(true, '');
});
T('★固定している時は 帯のぶん もう一度 描いている★（描かなければ 固定は効かない）', () => {
  /* ★2026-08-21 実際に踏んだ★:
     jsdom には本物の canvas が無いので「何が塗られたか」は見られない。
     状態（freezeRow など）だけ見ていたら ★帯を描かない壊し方が 素通りした★。
     ⇒ ★描く回数を数える★＝固定していれば 帯のぶん もう一度 呼ばれるはず。 */
  const 数える = () => {
    let n = 0;
    const 元 = win._renderPass;
    win._renderPass = function () { n++; return 元.apply(this, arguments); };
    try { win._doRender(); } finally { win._renderPass = 元; }
    return n;
  };
  reset();
  for (let r = 0; r < 20; r++) for (let c = 0; c < 5; c++) win.setCell(r, c, String(r) + '-' + String(c));
  eq(数える(), 1, '固定なしなのに 何回も描いている');
  win.sel(2, 0, 2, 0); win.freezePanes();
  eq(数える(), 2, '★上だけ固定＝帯のぶん もう一度 描いていない★');
  win.unfreezePanes(); win.sel(0, 2, 0, 2); win.freezePanes();
  eq(数える(), 2, '★左だけ固定＝帯のぶん もう一度 描いていない★');
  win.unfreezePanes(); win.sel(2, 2, 2, 2); win.freezePanes();
  eq(数える(), 4, '★上も左も固定＝4つの区画を描いていない★');
});
T('★前からある物① 並べ替えが まだ動く★', () => {
  reset();
  ['名前', 'い', 'ろ', 'は'].forEach((v, i) => win.setCell(i, 0, v));
  ['数', '3', '1', '2'].forEach((v, i) => win.setCell(i, 1, v));
  win.sel(0, 1, 0, 1);
  win.sortRange('asc');
  eq([0, 1, 2, 3].map((i) => 出る(i, 1)).join('|'), '数|1|2|3');
});
T('★前からある物② オートSUM が まだ動く★', () => {
  reset();
  for (let r = 0; r < 10; r++) win.setCell(r, 0, String(r + 1));
  win.sel(10, 0, 10, 0);
  win.autoSum();
  eq(A(10, 0).f, '=SUM(A1:A10)');
});
T('★前からある物③ 絞り込みが まだ動く★', () => {
  reset();
  ['くだもの', 'りんご', 'みかん', 'りんご'].forEach((v, i) => win.setCell(i, 0, v));
  ['数', '3', '1', '5'].forEach((v, i) => win.setCell(i, 1, v));
  win.sel(1, 0, 1, 0);
  win.filterByValue();
  eq(!!(sh().filterHidden || {})[2], true, 'みかんが隠れていない');
  eq(!!(sh().filterHidden || {})[1], false, 'りんごまで隠れた');
});
T('★前からある物④ 固定した上で 絞り込んでも 合計が変わらない★', () => {
  reset();
  ['くだもの', 'りんご', 'みかん', 'りんご'].forEach((v, i) => win.setCell(i, 0, v));
  ['数', '3', '1', '5'].forEach((v, i) => win.setCell(i, 1, v));
  win.setCell(0, 3, '=SUM(B2:B4)');
  const 前 = 出る(0, 3);
  eq(前, '9', '材料が違う');
  win.sel(1, 0, 1, 0); win.freezePanes();
  win.sel(1, 0, 1, 0); win.filterByValue();
  eq(出る(0, 3), 前, '★固定＋絞り込みで 合計が変わった★');
});

/* ── ⑧ ★Alt+= の「上と左が両方 すぐ隣」は 未測定＝どちらを取ったか出す★ ── */
T('★Alt+= ⓓ（上と左が両方 すぐ隣）は どちらを取ったか 画面に出す★（未測定を黙って決めない）', () => {
  reset();
  [0, 1, 2].forEach((r) => win.setCell(r, 2, '1'));   /* C1:C3（上） */
  [0, 1].forEach((c) => win.setCell(3, c, '2'));      /* A4:B4（左） */
  win.sel(3, 2, 3, 2);                                /* C4 */
  const 出た = toast(() => win.autoSum());
  ok(A(3, 2) && A(3, 2).f.indexOf('=SUM(') === 0, '式が入っていない');
  ok(出た.join('／').indexOf('上') >= 0 || 出た.join('／').indexOf('左') >= 0,
    '★どちらを取ったか 出していない★: ' + 出た.join('／'));
  ok(出た.join('／').indexOf('戻して') >= 0, '直せる事を言っていない: ' + 出た.join('／'));
});
T('★迷わない時は 余計な知らせを出さない★（ⓐ 上だけ）', () => {
  reset();
  for (let r = 0; r < 5; r++) win.setCell(r, 0, String(r + 1));
  win.sel(5, 0, 5, 0);
  const 出た = toast(() => win.autoSum());
  eq(出た.length, 0, '余計な知らせが出た: ' + 出た.join('／'));
});

/* ── ⑨ ★作っていない物を 台帳に書いてある★ ── */
T('★作っていない物が golden に書いてある（全部 作った事にしない）★', () => {
  ok(Array.isArray(GOLD['作っていない物']) && GOLD['作っていない物'].length >= 2, '作っていない物が書かれていない');
  ok(String(GOLD._how).length > 20, 'どう測ったかが書かれていない');
});

console.log('\n  ' + pass + ' passed, ' + fail + ' failed');

if (SELF) {
  const { spawnSync } = await import('node:child_process');
  const os = await import('node:os');
  const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'exally-freeze-'));
  console.log('\n[self-test] わざと壊して 赤くなるかを数える（★repo は読むだけ★）');
  const BREAKS = [
    ['book.html', '右クリックの 固定を消す', (s) => s.replace('<div class="ctx-item" onclick="freezePanes()">🧊 ここで固定（選んだセルの 上と左）</div>', '')],
    ['book.html', '★A1 でも固定してしまう★（真似ても合わない物を真似る）', (s) => s.replace("  if(selR1 === 0 && selC1 === 0){\n    showToast('A1 では固定できません（固定したい行の1つ下・列の1つ右を選んでね）');\n    return;\n  }", '')],
    ['book.html', '固定する行を 1つ ずらす', (s) => s.replace('  sh.freezeRow = selR1;                    /* 選んだセルの ★上の行★ を固定 */', '  sh.freezeRow = selR1 + 1;')],
    ['book.html', '固定する列を 1つ ずらす', (s) => s.replace('  sh.freezeCol = selC1;                    /* 選んだセルの ★左の列★ を固定 */', '  sh.freezeCol = selC1 + 1;')],
    ['book.html', '取り消しの控えを取らない', (s) => s.replace('  _pushRowColUndo();                       /* ★取り消し1回で戻る★ */', '')],
    ['book.html', '★押した所の読み取りが 帯を見ない★（固定した行を押せない）', (s) => s.replace('  var fz = sheets[activeSheet].freezeRow||0;\n  if(fz > 0){', '  var fz = 0;\n  if(fz > 0){')],
    ['book.html', '★左の帯の読み取りが 効かない★', (s) => s.replace('  var fzc = sheets[activeSheet].freezeCol||0;\n  if(fzc > 0){', '  var fzc = 0;\n  if(fzc > 0){')],
    ['book.html', '★帯をもう一度 描かない★（固定しても画面が変わらない）', (s) => s.replace('  _renderPass();                                       /* ① 流れる所（今までどおり） */', '  _renderPass(); return;')],
    ['book.html', '★Alt+= が どちらを取ったか 出さない★', (s) => s.replace('  if(got.ambiguous){', '  if(false){')],
    ['lib/autosum.js', '★上と左の両方が在る事に 気づかない★', (s) => s.replace('    var まよう = 上あり && 左あり;', '    var まよう = false;')],
  ];
  let red = 0;
  for (const [rel, name, brk] of BREAKS) {
    const orig = fs.readFileSync(path.join(ROOT, rel), 'utf8');
    const bad = brk(orig);
    if (bad === orig) { console.log('  ★置換できず★  ' + name); continue; }
    const tmpFile = path.join(TMP, rel.replace(/[\\/]/g, '_'));
    fs.writeFileSync(tmpFile, bad, 'utf8');
    const env = Object.assign({}, process.env, { EXALLY_FREEZE_OVERRIDE: JSON.stringify({ [rel]: tmpFile }) });
    const isRed = spawnSync(process.execPath, [path.join(__dirname, 'grid-freeze.test.mjs')], { encoding: 'utf8', env }).status !== 0;
    if (isRed) { red++; console.log('  赤くなった  ' + name); }
    else console.log('  ★素通り★  ' + name);
  }
  for (const rel of ['book.html', 'lib/autosum.js']) {
    const now = fs.readFileSync(path.join(ROOT, rel), 'utf8');
    if (now.includes('_renderPass(); return;') || now.includes('var まよう = false;') || now.includes('sh.freezeRow = selR1 + 1;')) {
      console.log('  ★NG★ ' + rel + ' に わざと壊した物が残っている');
      process.exit(1);
    }
  }
  try { fs.rmSync(TMP, { recursive: true, force: true }); } catch (e) { /* 消せなくても検査は済んでいる */ }
  console.log('\n  ' + red + '/' + BREAKS.length + ' 通りで赤くなった');
  process.exit(red === BREAKS.length ? 0 : 1);
}

process.exit(fail ? 1 : 0);
