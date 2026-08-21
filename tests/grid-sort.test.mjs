/* grid-sort.test.mjs — ★並べ替えが 実Excelと同じに並ぶか★
 *
 *  真値は tests/fixtures/excel-sort-golden.json（実Excel 16.0.20228 を COM で動かして取った）。
 *  ★ソースを読むだけにしない★＝本物の book.html を載せて ★本物の sortRange を呼ぶ★。
 *
 *  --self-test … わざと壊して「何通りで赤くなるか」を数える
 *                ★repo は読むだけ★（壊した中身は temp に置いて env で子へ渡す）
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(__dirname, '..');
const SELF = process.argv.includes('--self-test');
const OVERRIDE = process.env.EXALLY_SORT_OVERRIDE ? JSON.parse(process.env.EXALLY_SORT_OVERRIDE) : {};
const srcPath = (rel) => OVERRIDE[rel] || path.join(ROOT, rel);

let JSDOM;
try { ({ JSDOM } = await import('jsdom')); }
catch { console.log('★jsdomが入っていません。この検証は飛ばせません（SKIPを緑と呼ばない）。npm install してください。'); process.exit(1); }

const GOLD = JSON.parse(fs.readFileSync(path.join(ROOT, 'tests/fixtures/excel-sort-golden.json'), 'utf8'));

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

console.log('\n[grid-sort] ★本物の book.html で 実際に並べ替える★');
console.log('  真値 = ' + GOLD._measured_with);

/* ── 土台（無ければ この先の緑は全部 嘘） ── */
T('★画面が立ち上がっていて sortRange と GridSort が在る', () => {
  ok(Array.isArray(win.sheets) && win.sheets.length > 0, 'sheets が無い');
  ok(typeof win.sortRange === 'function', 'sortRange が無い');
  ok(win.GridSort && typeof win.GridSort.order === 'function', 'GridSort が無い');
});
T('★右クリックの中に 昇順・降順が在る（押す口が在る）', () => {
  const items = [...doc.querySelectorAll('.ctx-item')].map((e) => (e.getAttribute('onclick') || '') + '|' + e.textContent.trim());
  ok(items.some((x) => x.startsWith("sortRange('asc')")), '昇順の口が無い');
  ok(items.some((x) => x.startsWith("sortRange('desc')")), '降順の口が無い');
});

const A = (r, c) => win.sheets[win.activeSheet].data[r + ',' + c] || null;
const 出る = (r, c) => { const x = A(r, c); if (!x) return '（空）'; const v = (x.d !== undefined && x.d !== null && x.d !== '') ? x.d : x.v; return (v === undefined || v === null || v === '') ? '（空）' : String(v); };
function reset() { win.sheets[win.activeSheet].data = {}; win.sel(0, 0, 0, 0); }
function 置く(列, 値たち) { 値たち.forEach((v, i) => { if (v !== null && v !== '（空）') win.setCell(i, 列, v); }); }

/* ── ① 並ぶ順（実測どおりか） ── */
for (const dir of ['昇順', '降順']) {
  T('並べ替え … ' + dir + ' が 実Excelと同じ順（数→文字→論理→エラー／空白は最後）', () => {
    reset();
    /* 「（空）」は置かない＝空のまま */
    置く(0, GOLD['並ぶ順'].入れた物);
    win.sel(0, 0, GOLD['並ぶ順'].入れた物.length - 1, 0);   /* 範囲を選ぶ＝その範囲だけ並べ替える */
    win.sortRange(dir === '昇順' ? 'asc' : 'desc');
    /* ★論理値の見せ方だけ うちは小文字（true）＝golden の「うちとの違い」に書いてある★
       並び順は同じなので、ここでは 大文字小文字を揃えて 並びだけを見る。 */
    const そろえる = (x) => String(x).toUpperCase();
    const 出た = GOLD['並ぶ順'].入れた物.map((_, i) => そろえる(出る(i, 0)));
    eq(出た.join('|'), GOLD['並ぶ順'][dir].map(そろえる).join('|'));
  });
}

/* ── ② 見出しの自動判定（実測どおりか） ── */
for (const [名, cs] of Object.entries(GOLD['見出しの自動判定'])) {
  if (名 === '意味') continue;
  T('見出しの判定 … ' + 名 + '（' + (cs.見出しか ? '見出しとみなす' : '見出しではない') + '）', () => {
    reset();
    置く(0, cs.入れた物);
    win.sel(0, 0, 0, 0);                                   /* ★1セルだけ選ぶ＝表を自動で広げる★ */
    win.sortRange('asc');
    const 出た = cs.入れた物.map((_, i) => 出る(i, 0));
    eq(出た.join('|'), cs.結果.join('|'));
  });
}

/* ── ③ 行はセットで動く ── */
T('★行はセットで動く★（名前と数がバラバラにならない）＝実Excelで実測', () => {
  reset();
  置く(0, ['名前', 'い', 'ろ', 'は']);
  置く(1, ['数', '3', '1', '2']);
  win.sel(0, 1, 0, 1);                                     /* B列の見出しを選ぶ＝B列がキー */
  win.sortRange('asc');
  eq([0, 1, 2, 3].map((i) => 出る(i, 0)).join('|'), GOLD['行はセットで動く'].結果.A.join('|'), 'A列');
  eq([0, 1, 2, 3].map((i) => 出る(i, 1)).join('|'), GOLD['行はセットで動く'].結果.B.join('|'), 'B列');
});

/* ── ④ ★式はどうなるか★（ここが一番 大事） ── */
T('★相対参照の式は 各行で正しいまま／$付きは 指す先が変わる★＝実Excelで実測', () => {
  reset();
  置く(0, ['3', '1', '2']);
  win.setCell(0, 1, '=A1*10'); win.setCell(1, 1, '=A2*10'); win.setCell(2, 1, '=A3*10');
  win.setCell(0, 2, '=$A$1'); win.setCell(1, 2, '=$A$2'); win.setCell(2, 2, '=$A$3');
  win.sel(0, 0, 2, 2);
  win.sortRange('asc');
  const g = GOLD['式はどうなるか'].並べ替え後;
  eq([0, 1, 2].map((i) => 出る(i, 0)).join('|'), g.A.join('|'), 'A列');
  eq([0, 1, 2].map((i) => A(i, 1).f).join('|'), g.B_式.join('|'), 'B列の式');
  eq([0, 1, 2].map((i) => 出る(i, 1)).join('|'), g.B_答え.join('|'), 'B列の答え');
  eq([0, 1, 2].map((i) => A(i, 2).f).join('|'), g.C_式.join('|'), 'C列の式（$は動かない）');
  eq([0, 1, 2].map((i) => 出る(i, 2)).join('|'), g.C_答え.join('|'), 'C列の答え（★指す先が変わる★）');
});

/* ── ⑤ 表の広がり（1セルだけ選んだ時） ── */
T('★1セルだけ選ぶと 空の行と列で囲まれた かたまりだけ動く（離れた物は動かない）★', () => {
  reset();
  置く(0, ['名前', 'い', 'ろ', 'は']);
  置く(1, ['数', '3', '1', '2']);
  win.setCell(0, 4, 'はなれた');                            /* E1 相当（間に空列） */
  win.setCell(6, 0, 'あきをはさんだ');                       /* 空行をはさんだ物 */
  win.sel(2, 1, 2, 1);
  win.sortRange('asc');
  eq([0, 1, 2, 3].map((i) => 出る(i, 1)).join('|'), '数|1|2|3', '表の中');
  eq(出る(0, 4), 'はなれた', '離れた物が動いてしまった');
  eq(出る(6, 0), 'あきをはさんだ', '空行の向こうが動いてしまった');
});

/* ── ⑥ 取り消しが1回で効く ── */
T('★並べ替えは 取り消し1回で丸ごと戻る★', () => {
  reset();
  置く(0, ['名前', 'い', 'ろ', 'は']);
  置く(1, ['数', '3', '1', '2']);
  win.sel(0, 1, 0, 1);
  win.sortRange('asc');
  eq([0, 1, 2, 3].map((i) => 出る(i, 1)).join('|'), '数|1|2|3');
  win.doUndo();
  eq([0, 1, 2, 3].map((i) => 出る(i, 1)).join('|'), '数|3|1|2', '1回で戻っていない');
});

/* ── ⑦ 既に並んでいる時は 何もしない ── */
T('★もう並んでいる時は 書き換えない（黙って動かさない）★', () => {
  reset();
  置く(0, ['1', '2', '3']);
  win.sel(0, 0, 2, 0);
  win.sortRange('asc');
  eq([0, 1, 2].map((i) => 出る(i, 0)).join('|'), '1|2|3');
});

/* ── ⑧ ★測っていない事を 台帳に書いてある★ ── */
T('★見つけた違い（TRUE を true と出す）が golden に書いてある★', () => {
  const 違い = GOLD['うちとの違い（並べ替えで見つけた物）'];
  ok(違い && 違い['論理値の見せ方'], '見つけた違いが書かれていない');
  /* ★本当に まだ小文字で出るか を実際に見る（直したのに台帳が古い、を防ぐ）★ */
  reset();
  win.setCell(0, 0, 'TRUE');
  const 出 = 出る(0, 0);
  ok(出 === 'true' || 出 === 'TRUE', '思っていない出方: ' + 出);
  if (出 === 'TRUE') throw new Error('★もう大文字で出ている＝golden の「まだ直していません」が古い★');
});
T('★測っていない事が golden に書いてある（全部 測った事にしない）★', () => {
  ok(Array.isArray(GOLD['測っていない事']) && GOLD['測っていない事'].length >= 2, '測っていない事が書かれていない');
  ok(String(GOLD._how).length > 20, 'どう測ったかが書かれていない');
});

console.log('\n  ' + pass + ' passed, ' + fail + ' failed');

if (SELF) {
  const { spawnSync } = await import('node:child_process');
  const os = await import('node:os');
  const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'exally-sort-'));
  console.log('\n[self-test] わざと壊して 赤くなるかを数える（★repo は読むだけ★）');
  const BREAKS = [
    ['book.html', '右クリックの 昇順を消す', (s) => s.replace('<div class="ctx-item" onclick="sortRange(\'asc\')">⬆️ 昇順で並べ替え</div>', '')],
    ['book.html', '★式を移った先の行に読み替えない★（相対参照が壊れる）', (s) => s.replace('next.f = shiftFormula(src.f, ずれ, 0);', 'next.f = src.f;')],
    ['book.html', '見出しを いつも動かす', (s) => s.replace('var 先頭 = rng.r1 + (見出し ? 1 : 0);', 'var 先頭 = rng.r1;')],
    ['book.html', '表を広げず 1セルだけ並べ替える', (s) => s.replace('rng = GridSort.region(中身がある, selR1, selC1, ROWS-1, COLS-1);', 'rng = { r1:selR1, c1:selC1, r2:selR1, c2:selC1 };')],
    ['book.html', '取り消しの控えを取らない', (s) => s.replace('  _pushRowColUndo();                       /* ★1回の取り消しで丸ごと戻せる★ */', '')],
    ['book.html', 'キーの列を いつもA列にする', (s) => s.replace('var キー列 = Math.max(rng.c1, Math.min(selC1, rng.c2));', 'var キー列 = rng.c1;')],
    ['lib/grid-sort.js', '★空白を最後にしない★', (s) => s.replace('return 中身.map(function (x) { return x.i; }).concat(空の物.map(function (x) { return x.i; }));', 'return 空の物.map(function (x) { return x.i; }).concat(中身.map(function (x) { return x.i; }));')],
    ['lib/grid-sort.js', '型の順を変える（文字を数より先に）', (s) => s.replace('var 数 = 0, 文字 = 1, 論理 = 2, エラー = 3, 空 = 4;', 'var 数 = 1, 文字 = 0, 論理 = 2, エラー = 3, 空 = 4;')],
    ['lib/grid-sort.js', '見出しの判定を いつも true にする', (s) => s.replace('    if (!先頭に文字) return false;', '    return true;')],
    ['lib/grid-sort.js', '表の広がりを 縦だけにする', (s) => s.replace('      while (c1 > 0 && 列に何かある(c1 - 1, r1, r2)) { c1--; 広がった = true; }', '')],
  ];
  let red = 0;
  for (const [rel, name, brk] of BREAKS) {
    const orig = fs.readFileSync(path.join(ROOT, rel), 'utf8');
    const bad = brk(orig);
    if (bad === orig) { console.log('  ★置換できず★  ' + name); continue; }
    const tmpFile = path.join(TMP, rel.replace(/[\\/]/g, '_'));
    fs.writeFileSync(tmpFile, bad, 'utf8');
    const env = Object.assign({}, process.env, { EXALLY_SORT_OVERRIDE: JSON.stringify({ [rel]: tmpFile }) });
    const isRed = spawnSync(process.execPath, [path.join(__dirname, 'grid-sort.test.mjs')], { encoding: 'utf8', env }).status !== 0;
    if (isRed) { red++; console.log('  赤くなった  ' + name); }
    else console.log('  ★素通り★  ' + name);
  }
  /* ★repo を1バイトも触っていない事を その場で確かめる★ */
  for (const rel of ['book.html', 'lib/grid-sort.js']) {
    const now = fs.readFileSync(path.join(ROOT, rel), 'utf8');
    if (now.includes('next.f = src.f;') || now.includes('var 数 = 1, 文字 = 0')) {
      console.log('  ★NG★ ' + rel + ' に わざと壊した物が残っている');
      process.exit(1);
    }
  }
  try { fs.rmSync(TMP, { recursive: true, force: true }); } catch (e) { /* 消せなくても検査は済んでいる */ }
  console.log('\n  ' + red + '/' + BREAKS.length + ' 通りで赤くなった');
  process.exit(red === BREAKS.length ? 0 : 1);
}

process.exit(fail ? 1 : 0);
