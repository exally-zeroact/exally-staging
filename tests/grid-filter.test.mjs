/* grid-filter.test.mjs — ★絞り込み（フィルター）が 実Excelと同じに動くか★
 *
 *  真値は tests/fixtures/excel-filter-golden.json（実Excel 16.0.20228 を COM で動かして取った）。
 *  ★ソースを読むだけにしない★＝本物の book.html を載せて ★本物の filterByValue を呼ぶ★。
 *
 *  ★一番 大事な所★ … ★絞っても 合計が変わらない★（隠れた行も足す＝実Excelで実測）
 *
 *  --self-test … わざと壊して「何通りで赤くなるか」を数える（★repo は読むだけ★）
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(__dirname, '..');
const SELF = process.argv.includes('--self-test');
const OVERRIDE = process.env.EXALLY_FILTER_OVERRIDE ? JSON.parse(process.env.EXALLY_FILTER_OVERRIDE) : {};
const srcPath = (rel) => OVERRIDE[rel] || path.join(ROOT, rel);

let JSDOM;
try { ({ JSDOM } = await import('jsdom')); }
catch { console.log('★jsdomが入っていません。この検証は飛ばせません（SKIPを緑と呼ばない）。npm install してください。'); process.exit(1); }

const GOLD = JSON.parse(fs.readFileSync(path.join(ROOT, 'tests/fixtures/excel-filter-golden.json'), 'utf8'));

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

console.log('\n[grid-filter] ★本物の book.html で 実際に絞り込む★');
console.log('  真値 = ' + GOLD._measured_with);

const sh = () => win.sheets[win.activeSheet];
const A = (r, c) => sh().data[r + ',' + c] || null;
const 出る = (r, c) => { const x = A(r, c); if (!x) return '（空）'; const v = (x.d !== undefined && x.d !== null && x.d !== '') ? x.d : x.v; return (v === undefined || v === null || v === '') ? '（空）' : String(v); };
function reset() { sh().data = {}; sh().filterHidden = {}; sh().hiddenRows = {}; win.sel(0, 0, 0, 0); }
/** golden の置いた物を そのまま並べる（見出し1行＋中身5行＋合計3本） */
function 置く() {
  reset();
  GOLD['置いた物'].見出し.forEach((v, c) => win.setCell(0, c, v));
  GOLD['置いた物'].中身.forEach((row, i) => { win.setCell(i + 1, 0, row[0]); win.setCell(i + 1, 1, String(row[1])); });
  win.setCell(0, 3, '=SUM(B2:B6)');
}

/* ── 土台 ── */
T('★画面が立ち上がっていて filterByValue と GridFilter が在る', () => {
  ok(typeof win.filterByValue === 'function', 'filterByValue が無い');
  ok(typeof win.clearFilter === 'function', 'clearFilter が無い');
  ok(win.GridFilter && typeof win.GridFilter.byValue === 'function', 'GridFilter が無い');
});
T('★右クリックの中に 絞り込みと 解除が在る（押す口が在る）', () => {
  const items = [...doc.querySelectorAll('.ctx-item')].map((e) => e.getAttribute('onclick') || '');
  ok(items.some((x) => x.startsWith('filterByValue()')), '絞り込みの口が無い');
  ok(items.some((x) => x.startsWith('clearFilter()')), '解除の口が無い');
});

/* ── ① 隠れ方が実測どおりか ── */
T('★「りんご」で絞ると 実Excelと同じ行が隠れる（見出しは隠れない）★', () => {
  置く();
  win.sel(1, 0, 1, 0);                      /* A2＝「りんご」 */
  win.filterByValue();
  const f = sh().filterHidden || {};
  const 出た = {};
  for (let r = 0; r <= 5; r++) 出た[String(r + 1)] = !!f[r];   /* golden は1起算 */
  eq(JSON.stringify(出た), JSON.stringify(GOLD['りんごで絞った'].行の隠れ方));
});
T('★中身は消えていない（隠れているだけ）★', () => {
  const 中 = GOLD['置いた物'].中身;
  for (let i = 0; i < 中.length; i++) eq(出る(i + 1, 0), 中[i][0], (i + 2) + '行目');
});
T('★見えている行の数が 実測と同じ★', () => {
  const f = sh().filterHidden || {};
  let 見えている = 0;
  for (let r = 0; r <= 5; r++) if (!f[r]) 見えている++;
  eq(見えている, GOLD['りんごで絞った'].見えている行 + 1, '見出しを含めた数');
});

/* ── ② ★合計が変わらない★（ここが一番 大事） ── */
T('★絞っても 合計が変わらない（隠れた行も足す）★＝実Excelで実測', () => {
  置く();
  const 前 = 出る(0, 3);
  eq(前, GOLD['合計はどうなるか'].絞る前.SUM, '絞る前の合計');
  win.sel(1, 0, 1, 0);
  win.filterByValue();
  eq(出る(0, 3), GOLD['合計はどうなるか'].絞った後.SUM, '★絞った後に 合計が変わった＝黙って嘘をつく★');
});
T('★手で隠しても 合計が変わらない★＝実Excelで実測', () => {
  置く();
  sh().hiddenRows = { 2: true };
  win.render();
  eq(出る(0, 3), GOLD['合計はどうなるか'].手で行を隠した.SUM);
});

/* ── ③ フィルターと 手で隠した行は 別物 ── */
T('★解除しても 人が手で隠した行は 出てこない（別々に持っている）★', () => {
  置く();
  sh().hiddenRows = { 3: true };            /* 人が手で隠した行 */
  win.sel(1, 0, 1, 0);
  win.filterByValue();
  win.clearFilter();
  eq(!!(sh().hiddenRows || {})[3], true, '★手で隠した行まで出てきた★');
  eq(Object.keys(sh().filterHidden || {}).length, 0, '絞り込みが残っている');
});

/* ── ④ 解除で全部 戻る ── */
T('★解除すると 全部の行が戻る★＝実Excelで実測', () => {
  置く();
  win.sel(1, 0, 1, 0);
  win.filterByValue();
  win.clearFilter();
  const f = sh().filterHidden || {};
  const 出た = {};
  for (let r = 0; r <= 5; r++) 出た[String(r + 1)] = !!f[r];
  eq(JSON.stringify(出た), JSON.stringify(GOLD['解除したら'].行の隠れ方));
});

/* ── ⑤ 行の高さが 0 になる（画面から消える） ── */
T('★隠れた行は 画面の高さが 0 になる（本当に見えなくなる）★', () => {
  置く();
  win.sel(1, 0, 1, 0);
  win.filterByValue();
  eq(win.rH(2), 0, '3行目（みかん）が隠れていない');
  ok(win.rH(1) > 0, '2行目（りんご）まで隠れている');
  eq(win.rH(0) > 0, true, '見出しが隠れている');
});

/* ── ⑥ 見出しの行では絞らない ── */
T('★見出しの行を選んで押したら「見出しでは絞れない」と言う★（黙って何もしない、にしない）', () => {
  置く();
  win.sel(0, 0, 0, 0);
  /* ★出た知らせを読む★＝「絞らなかった」だけでなく ★理由を言ったか★ を見る。
     （2026-08-21 実際に踏んだ: 理由を出す行を消しても、別の行に引っかかって
       「絞らない」結果は同じになり ★見張りが素通りした★） */
  const 出た = [];
  const 元 = win.showToast;
  win.showToast = function (m) { 出た.push(String(m)); };
  try { win.filterByValue(); } finally { win.showToast = 元; }
  eq(Object.keys(sh().filterHidden || {}).length, 0, '見出しで絞ってしまった');
  ok(出た.length > 0, '何も言わずに黙って終わった');
  ok(出た.join('／').indexOf('見出し') >= 0,
    '理由が「見出しだから」になっていない: ' + 出た.join('／'));
});

/* ── ⑦ ★Excelの置き場所と同じか（実物から取った一覧に在る）★ ── */
T('★Excelの右クリックにも「選択したセルの値でフィルター」が在る（実物から取った）★', () => {
  const l = GOLD['Excelの右クリックにも在る（実物から取った）'];
  ok(l && Array.isArray(l.フィルター), '実物の一覧が無い');
  ok(l.フィルター.some((x) => x.indexOf('選択したセルの値') >= 0), '実物の一覧に無い＝勝手に作った事になる');
  ok(l.並べ替え.indexOf('昇順') >= 0 && l.並べ替え.indexOf('降順') >= 0, '並べ替えの一覧が違う');
});

/* ── ⑧ ★作っていない物を 台帳に書いてある★ ── */
T('★作っていない物が golden に書いてある（全部 作った事にしない）★', () => {
  ok(Array.isArray(GOLD['作っていない物']) && GOLD['作っていない物'].length >= 3, '作っていない物が書かれていない');
  ok(String(GOLD._how).length > 20, 'どう測ったかが書かれていない');
});

console.log('\n  ' + pass + ' passed, ' + fail + ' failed');

if (SELF) {
  const { spawnSync } = await import('node:child_process');
  const os = await import('node:os');
  const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'exally-filter-'));
  console.log('\n[self-test] わざと壊して 赤くなるかを数える（★repo は読むだけ★）');
  const BREAKS = [
    ['book.html', '右クリックの 絞り込みを消す', (s) => s.replace('<div class="ctx-item" onclick="filterByValue()">🔽 選んだセルの値で絞り込む</div>', '')],
    ['book.html', '★隠した行を 手で隠した所と混ぜる★（解除で人の行まで出る）', (s) => s.replace('  sh.filterHidden = {};\n  for(var i=0;i<res.hide.length;i++) sh.filterHidden[res.hide[i]] = true;', '  sh.hiddenRows = sh.hiddenRows || {};\n  for(var i=0;i<res.hide.length;i++) sh.hiddenRows[res.hide[i]] = true;')],
    ['book.html', '見出しの行でも絞れるようにする', (s) => s.replace("  if(見出し && selR1===rng.r1){ showToast('見出しの行では絞り込めません（下の行を選んでね）'); return; }", '')],
    ['book.html', '★隠すのではなく 中身を消す★', (s) => s.replace('  for(var i=0;i<res.hide.length;i++) sh.filterHidden[res.hide[i]] = true;', '  for(var i=0;i<res.hide.length;i++){ sh.filterHidden[res.hide[i]] = true; for(var c=rng.c1;c<=rng.c2;c++) delete data[res.hide[i]+","+c]; }')],
    ['book.html', '隠れた行の高さを 0 にしない（見た目が変わらない）', (s) => s.replace('  if(sh.filterHidden && sh.filterHidden[r]) return 0;             /* ★絞り込みで隠れた行（別物として持つ）★ */', '')],
    ['book.html', '解除で戻さない', (s) => s.replace('  sh.filterHidden = {};\n  var ctl = document.getElementById(\'ctx-clear-filter\'); if(ctl) ctl.style.display = \'none\';', '  var ctl = document.getElementById(\'ctx-clear-filter\'); if(ctl) ctl.style.display = \'none\';')],
    ['lib/grid-filter.js', '★同じ値の行を 逆に隠す★', (s) => s.replace("      if (text(get(r, keyCol)) === 見たい) keep.push(r); else hide.push(r);", "      if (text(get(r, keyCol)) === 見たい) hide.push(r); else keep.push(r);")],
    ['lib/grid-filter.js', '見出し行も 隠す対象にする', (s) => s.replace('    var 先頭 = rng.r1 + (見出しあり ? 1 : 0);', '    var 先頭 = rng.r1;')],
  ];
  let red = 0;
  for (const [rel, name, brk] of BREAKS) {
    const orig = fs.readFileSync(path.join(ROOT, rel), 'utf8');
    const bad = brk(orig);
    if (bad === orig) { console.log('  ★置換できず★  ' + name); continue; }
    const tmpFile = path.join(TMP, rel.replace(/[\\/]/g, '_'));
    fs.writeFileSync(tmpFile, bad, 'utf8');
    const env = Object.assign({}, process.env, { EXALLY_FILTER_OVERRIDE: JSON.stringify({ [rel]: tmpFile }) });
    const isRed = spawnSync(process.execPath, [path.join(__dirname, 'grid-filter.test.mjs')], { encoding: 'utf8', env }).status !== 0;
    if (isRed) { red++; console.log('  赤くなった  ' + name); }
    else console.log('  ★素通り★  ' + name);
  }
  for (const rel of ['book.html', 'lib/grid-filter.js']) {
    const now = fs.readFileSync(path.join(ROOT, rel), 'utf8');
    if (now.includes('hide.push(r); else keep.push(r)') || now.includes('sh.hiddenRows[res.hide[i]] = true')) {
      console.log('  ★NG★ ' + rel + ' に わざと壊した物が残っている');
      process.exit(1);
    }
  }
  try { fs.rmSync(TMP, { recursive: true, force: true }); } catch (e) { /* 消せなくても検査は済んでいる */ }
  console.log('\n  ' + red + '/' + BREAKS.length + ' 通りで赤くなった');
  process.exit(red === BREAKS.length ? 0 : 1);
}

process.exit(fail ? 1 : 0);
