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
  /* ★2026-08-31：右クリックの 中身が ★開いた時に 作る★ 形に なった★
     （実 Excel の 1,384命令に 合わせて 25→37個。組（★）に した）
     ⇒ ★開いてから 探す★。印は onclick ではなく data-ctx。 */
  win.sel(0, 0, 0, 0);
  win.showCtxMenu(10, 10);
  const items = [...doc.querySelectorAll('#ctx-menu .ctx-item')]
    .map((e) => e.getAttribute('data-ctx') || '');
  ok(items.indexOf('freezePanes') >= 0, '固定の口が無い：' + items.join(' / '));
  ok(items.indexOf('unfreezePanes') >= 0, '固定をやめる口が無い');
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

/* ── ⑩ ★固定したまま 下・右へ動かすと 帯の下・右に 動かした先が描かれる★（2026-11-07 実WebKit で 空白を見た） ──
   描き始めを 帯の0行目から取っていた＝画面1枚分（約30行）で止まり、101行目より下が描かれず 帯の下が空白だった。
   ★描く時に getCell へ来た 行・列を 数える★（ctx は作り物＝絵ではなく「どの行を描きに行ったか」で見る） */
function 描いた行と列(fzR, fzC, sT, sL) {
  const sh0 = sh(); const 元 = win.getCell, 元H = win.wrapH, 元W = win.wrapW;
  sh0.freezeRow = fzR; sh0.freezeCol = fzC; win.scrollTop = sT; win.scrollLeft = sL;
  win.wrapH = 600; win.wrapW = 900;                  /* jsdom は 0＝置かないと 何も描かれず 空振りする */
  const 行 = new Set(), 列 = new Set();
  win.getCell = function (r, c) { 行.add(r); 列.add(c); return 元.apply(this, arguments); };
  try { win._doRender(); } finally { win.getCell = 元; win.wrapH = 元H; win.wrapW = 元W; }
  return { 行, 列 };
}
T('★固定3行で 2400px 下へ動かすと 流れる所の行（101〜120）を描きに行く★', () => {
  reset();
  const { 行 } = 描いた行と列(3, 0, 2400, 0);
  const 見える = Math.floor(2400 / (win.ROW_H * win.scale));
  for (let r = 見える + 4; r < 見える + 20; r++) ok(行.has(r), '★' + (r + 1) + '行目を 描いていない＝帯の下が空白★');
  ok(行.has(0) && 行.has(2), '固定した帯（1〜3行目）を描いていない');
});
T('★固定3列で 右へ動かすと 流れる所の列を描きに行く★', () => {
  reset();
  const { 列 } = 描いた行と列(0, 3, 0, 3000);
  const 見える = Math.floor(3000 / (win.COL_W * win.scale));
  for (let c = 見える + 4; c < 見える + 8; c++) ok(列.has(c), '★' + (c + 1) + '列目を 描いていない＝帯の右が空白★');
});
T('★PageDown は 帯の下に見えている行の数だけ 動く★（帯の0行目を一番上と取ると 100行以上 飛んだ）', () => {
  reset();
  const 元H = win.wrapH; win.wrapH = 600;
  try {
    const 動き = (fz) => { sh().freezeRow = fz; win.sel(110, 0, 110, 0); win.scrollTop = 2400; const 前 = win.selR1; win.movePage(1); return win.selR1 - 前; };
    const 無し = 動き(0), 有り = 動き(3);
    /* ★帯の下に見えている行＝固定なしより ちょうど3行少ない★（多いと飛びすぎ・同じだと帯に隠れた3行を毎回飛ばす） */
    eq(無し - 有り, 3, '★固定3行で ' + 有り + '行 動いた（固定なしは ' + 無し + '行）★');
  } finally { win.wrapH = 元H; sh().freezeRow = 0; }
});

/* ── ⑪ ★固定・分割が 画面より大きくても 描くのは 画面に入る分だけ★（2026-11-07 測った：固定1048575行で 1回 64秒・getCell 4194万回） ──
   ★getCell の数に 上限を置き 超えたら throw★＝直しを外しても 固まらずに すぐ赤になる。
   ★比べるのは 同じ形同士★（行だけ同士・行＋列同士・分割同士）。wrapH/wrapW は jsdom で0＝置かないと 空振り */
const 上限 = 50000;
class 描きすぎ extends Error {}
function 描いた数(設定) {
  const sh0 = sh(); const 元 = win.getCell, 元H = win.wrapH, 元W = win.wrapW, 元T = win.scrollTop, 元L = win.scrollLeft;
  const 元分割 = Object.assign({}, win.分割);
  sh0.freezeRow = 設定.fzR || 0; sh0.freezeCol = 設定.fzC || 0;
  win.分割.行 = 設定.分R || 0; win.分割.列 = 設定.分C || 0; win.分割.上のずれ = 0; win.分割.左のずれ = 0;
  win.scrollTop = 0; win.scrollLeft = 0; win.wrapH = 600; win.wrapW = 900;
  let n = 0;
  win.getCell = function () { if (++n > 上限) throw new 描きすぎ('getCell ' + 上限 + ' 回を越えた'); return 元.apply(this, arguments); };
  try { win._doRender(); }
  catch (e) { if (!(e instanceof 描きすぎ)) throw e; }
  finally {
    win.getCell = 元; win.wrapH = 元H; win.wrapW = 元W; win.scrollTop = 元T; win.scrollLeft = 元L;
    Object.assign(win.分割, 元分割); sh0.freezeRow = 0; sh0.freezeCol = 0;
  }
  return n;
}
const 同じ程度 = (大, 小, 名) => {
  ok(大 <= 上限, '★' + 名 + '：getCell ' + 上限 + ' 回を越えた＝画面の外まで描いている★');
  ok(大 <= 小 * 2 + 200, '★' + 名 + '：大きい時 ' + 大 + ' 回・小さい時 ' + 小 + ' 回＝画面の外まで描いている★');
  ok(小 > 0, '★' + 名 + '：小さい時 1回も描いていない＝試験が空振り★');
};
T('★固定 1048575行でも 描く数は 固定30行と同じ程度★', () => {
  reset();
  同じ程度(描いた数({ fzR: win.ROWS - 1 }), 描いた数({ fzR: 30 }), '行の固定');
});
T('★固定 1048575行×16383列でも 描く数は 固定30×30と同じ程度★', () => {
  reset();
  同じ程度(描いた数({ fzR: win.ROWS - 1, fzC: win.COLS - 1 }), 描いた数({ fzR: 30, fzC: 30 }), '行と列の固定');
});
T('★分割 1048575行×16383列でも 描く数は 分割30×30と同じ程度★', () => {
  reset();
  同じ程度(描いた数({ 分R: win.ROWS - 1, 分C: win.COLS - 1 }), 描いた数({ 分R: 30, 分C: 30 }), '分割');
});

/* ── ⑫ ★帯の囲い（clip）は「本当の高さ」と「画面」の小さい方・境目の線は 本当の位置★（2026-11-07 taiketsu 任意1・3）
   ⑪ は描く数の上の端しか見ない＝打ち切りすぎ（帯の下に流れる所が出る）・画面より小さい固定が変わる・線が画面の端に出る に気づかない。
   ★ctx を 書き留める作り物に 差し替えて rect と moveTo を 数で見る★ */
function 描いた囲い(設定) {
  const sh0 = sh(); const 元ctx = win.ctx, 元H = win.wrapH, 元W = win.wrapW, 元T = win.scrollTop, 元L = win.scrollLeft;
  const 元分割 = Object.assign({}, win.分割);
  sh0.freezeRow = 設定.fzR || 0; sh0.freezeCol = 設定.fzC || 0;
  win.分割.行 = 設定.分R || 0; win.分割.列 = 設定.分C || 0; win.分割.上のずれ = 0; win.分割.左のずれ = 0;
  win.scrollTop = 0; win.scrollLeft = 0; win.wrapH = 600; win.wrapW = 900;
  const 囲い = [], 線 = []; let 前 = null, 色 = '';
  win.ctx = new Proxy(元ctx, {
    get(t, k) {
      if (k === 'rect') return (x, y, w, h) => { 前 = [x, y, w, h]; };
      if (k === 'clip') return () => { if (前) 囲い.push(前); };
      if (k === 'moveTo') return (x, y) => { if (色 === '#3D9E72') 線.push([x, y]); };
      return t[k];
    },
    set(t, k, v) { if (k === 'strokeStyle') 色 = v; return true; },
  });
  let 高さ, 幅, n = 0; const 元getCell = win.getCell;
  win.getCell = function () { if (++n > 上限) throw new 描きすぎ('getCell ' + 上限 + ' 回を越えた'); return 元getCell.apply(this, arguments); };   /* 壊した時も 固まらずに赤 */
  try { 高さ = 設定.分R || 設定.分C ? win.分割の高さ() : win._freezeH(); 幅 = 設定.分R || 設定.分C ? win.分割の幅() : win._freezeW(); win._doRender(); }
  finally {
    win.getCell = 元getCell; win.ctx = 元ctx; win.wrapH = 元H; win.wrapW = 元W; win.scrollTop = 元T; win.scrollLeft = 元L;
    Object.assign(win.分割, 元分割); sh0.freezeRow = 0; sh0.freezeCol = 0;
  }
  return { 囲い, 線, 高さ, 幅 };
}
const 囲いが在る = (囲い, w, h, 名) => ok(囲い.some((r) => r[0] === 0 && r[1] === 0 && r[2] === w && r[3] === h), '★' + 名 + '：囲い 0,0,' + w + ',' + h + ' が無い★ 在ったのは ' + JSON.stringify(囲い));
T('★画面より小さい固定（5行×2列）は 帯も線も 今まで通り★', () => {
  reset();
  const { 囲い, 線, 高さ, 幅 } = 描いた囲い({ fzR: 5, fzC: 2 });
  ok(高さ > 0 && 高さ < 600 - win.HDR_H && 幅 > 0 && 幅 < 900 - win.HDR_W, '試験の前提（画面より小さい）が崩れた ' + 高さ + '×' + 幅);
  囲いが在る(囲い, 900, win.HDR_H + 高さ, '上の帯');
  囲いが在る(囲い, win.HDR_W + 幅, 600, '左の帯');
  囲いが在る(囲い, win.HDR_W + 幅, win.HDR_H + 高さ, '左上の角');
  ok(線.some((p) => p[0] === 0 && p[1] === Math.round(win.HDR_H + 高さ) + 0.5), '★横の境目の線が 本当の位置に無い★ ' + JSON.stringify(線));
  ok(線.some((p) => p[1] === 0 && p[0] === Math.round(win.HDR_W + 幅) + 0.5), '★縦の境目の線が 本当の位置に無い★ ' + JSON.stringify(線));
});
T('★画面より大きい固定は 帯を ちょうど画面まで描き・線は画面の外（本当の位置）★', () => {
  reset();
  const { 囲い, 線, 高さ, 幅 } = 描いた囲い({ fzR: win.ROWS - 1, fzC: win.COLS - 1 });
  囲いが在る(囲い, 900, 600, '上の帯＝画面まで（短いと 帯の下に流れる所が出る）');
  ok(線.length === 2, '線が2本でない ' + JSON.stringify(線));
  ok(線.every((p) => p[0] > 900 + 2 || p[1] > 600 + 2), '★境目の線が 画面の中か端に出た★（打ち切った高さで引くと 600.5／900.5＝端に かかる） ' + JSON.stringify(線));
  ok(線.some((p) => p[1] === Math.round(win.HDR_H + 高さ) + 0.5) && 線.some((p) => p[0] === Math.round(win.HDR_W + 幅) + 0.5), '★線が 本当の位置に無い★');
});
/* ★行だけ・列だけで見る★（両方だと 左の帯の囲いが ちょうど 900×600 になり 上の帯の短さを隠す＝self-test で素通りした） */
T('★画面より大きい固定（行だけ・列だけ）も 帯を ちょうど画面まで描く★', () => {
  reset();
  囲いが在る(描いた囲い({ fzR: win.ROWS - 1 }).囲い, 900, 600, '行だけの固定＝上の帯');
  囲いが在る(描いた囲い({ fzC: win.COLS - 1 }).囲い, 900, 600, '列だけの固定＝左の帯');
});
T('★画面より大きい分割（行だけ・列だけ）も 側を ちょうど画面まで描く★', () => {
  reset();
  囲いが在る(描いた囲い({ 分R: win.ROWS - 1 }).囲い, 900, 600, '行だけの分割＝上の側');
  囲いが在る(描いた囲い({ 分C: win.COLS - 1 }).囲い, 900, 600, '列だけの分割＝左の側');
});

/* ── ⑬ ★選んだ見出しの強調は 見えている行・列だけ塗る★（2026-11-07 測った：A列丸ごと選ぶと 描くたび fillRect 1,048,610回・599ms）
   ★見出しの色（C_SEL_HDR）の fillRect に 上限を置き 超えたら throw★＝打ち切りを外すと すぐ赤。
   ★打ち切りすぎも見る★＝行の見出しの塗りが 画面の上（HDR_H）から下（600）まで 隙間なく 届いているか */
const 見出し上限 = 3000;
function 見出しの塗り(設定) {
  const sh0 = sh(); const 元ctx = win.ctx, 元H = win.wrapH, 元W = win.wrapW, 元T = win.scrollTop, 元L = win.scrollLeft;
  const 元分割 = Object.assign({}, win.分割);
  win.sel(設定.r1, 設定.c1, 設定.r2, 設定.c2);
  if (設定.足す) win.eval('selRanges=' + JSON.stringify(設定.足す) + ';');
  sh0.freezeRow = 設定.fzR || 0; sh0.freezeCol = 設定.fzC || 0;
  win.分割.行 = 設定.分R || 0; win.分割.列 = 設定.分C || 0; win.分割.上のずれ = 0; win.分割.左のずれ = 0;
  win.wrapH = 600; win.wrapW = 900; win.scrollTop = 設定.top || 0; win.scrollLeft = 0;
  const 行 = [], 列 = []; let 色 = '', n = 0;
  win.ctx = new Proxy(元ctx, {
    get(t, k) {
      if (k === 'fillRect') return (x, y, w, h) => {
        if (色 !== win.C_SEL_HDR) return;
        if (++n > 見出し上限) throw new 描きすぎ('見出しの fillRect ' + 見出し上限 + ' 回を越えた');
        if (x === 0 && w === win.HDR_W) 行.push([y, y + h]); else if (y === 0) 列.push([x, x + w]);
      };
      return t[k];
    },
    set(t, k, v) { if (k === 'fillStyle') 色 = v; return true; },
  });
  let 越えた = false;
  try { win._doRender(); }
  catch (e) { if (!(e instanceof 描きすぎ)) throw e; 越えた = true; }
  finally {
    win.ctx = 元ctx; win.wrapH = 元H; win.wrapW = 元W; win.scrollTop = 元T; win.scrollLeft = 元L;
    Object.assign(win.分割, 元分割); sh0.freezeRow = 0; sh0.freezeCol = 0;
    win.eval('selRanges=[];');
  }
  return { 行, 列, n, 越えた };
}
/* 画面の a〜b が 塗りの並びで 隙間なく 埋まっているか（行の見出しは 縦・列の見出しは 横） */
const 埋まる = (並び, a, b, 名) => {
  const s = 並び.filter((p) => p[1] > a && p[0] < b).sort((p, q) => p[0] - q[0]);
  let 端 = a;
  for (const [p0, p1] of s) { if (p0 > 端) break; 端 = Math.max(端, p1); }
  ok(端 >= b, '★' + 名 + '：' + a + '〜' + b + ' のうち ' + 端 + ' までしか塗っていない＝打ち切りすぎ★ 塗り ' + s.length + ' 個');
};
const 見出し見る = (r, 名, 列も) => {
  ok(!r.越えた, '★' + 名 + '：見出しの fillRect ' + 見出し上限 + ' 回を越えた＝画面の外まで塗っている★');
  ok(r.n > 0, '★' + 名 + '：1回も塗っていない＝試験が空振り★');
  埋まる(r.行, win.HDR_H, 600, 名 + '・行の見出し');
  if (列も) 埋まる(r.列, win.HDR_W, 900, 名 + '・列の見出し');
};
const 全行 = () => win.ROWS - 1, 全列 = () => win.COLS - 1;
T('★A列を丸ごと選んでも 見出しの塗りは 画面の分だけ・上から下まで届く★', () => {
  reset();
  見出し見る(見出しの塗り({ r1: 0, c1: 0, r2: 全行(), c2: 0 }), 'A列丸ごと');
  見出し見る(見出しの塗り({ r1: 0, c1: 0, r2: 全行(), c2: 0, top: 5000 * win.ROW_H * win.scale + 7 }), 'A列丸ごと・下へ動かした（半端な位置）');
});
T('★全部選んでも 行・列の見出しが 画面の分だけ・端まで届く★', () => {
  reset();
  見出し見る(見出しの塗り({ r1: 0, c1: 0, r2: 全行(), c2: 全列() }), '全部', true);
});
T('★固定3行×2列・分割で 下へ動かしても 見出しの塗りが 帯の下まで届く★', () => {
  reset();
  const top = 5000 * win.ROW_H * win.scale;
  見出し見る(見出しの塗り({ r1: 0, c1: 0, r2: 全行(), c2: 全列(), fzR: 3, fzC: 2, top }), '固定3×2', true);
  見出し見る(見出しの塗り({ r1: 0, c1: 0, r2: 全行(), c2: 全列(), fzR: 全行(), top }), '固定 1048575行', true);
  見出し見る(見出しの塗り({ r1: 0, c1: 0, r2: 全行(), c2: 全列(), 分R: 3, 分C: 2, top }), '分割3×2', true);
});
T('★Ctrl で足した範囲（selRanges）も 見出しの塗りは 画面の分だけ★', () => {
  reset();
  const 足す = [{ r1: 0, c1: 3, r2: 全行(), c2: 3, sheet: sh().name }, { r1: 2, c1: 0, r2: 2, c2: 全列(), sheet: sh().name }];
  見出し見る(見出しの塗り({ r1: 0, c1: 0, r2: 0, c2: 0, 足す }), '足した範囲', true);
});
T('★隠れた行が 多くても 見出しの塗りが 下まで届く★', () => {
  reset();
  for (let r = 10; r < 2010; r++) sh().hiddenRows[r] = true;
  try { 見出し見る(見出しの塗り({ r1: 0, c1: 0, r2: 全行(), c2: 0 }), '2000行 隠した'); }
  finally { sh().hiddenRows = {}; }
});

/* ── ⑭ ★低い行・細い列（Excel 方眼紙）でも 画面の下・右の端まで描く★（2026-11-07 実WebKit で見た：列幅24で T列から右・行高8で 23行目から下が 白く残った）
   ★見出しの字（列の字・行の番号）の 一番 右・一番 下が 画面の端まで 届いているか★を 数で見る */
function 見出しの字(設定) {
  const sh0 = sh(); const 元ctx = win.ctx, 元H = win.wrapH, 元W = win.wrapW, 元T = win.scrollTop, 元L = win.scrollLeft;
  const 元幅 = sh0.colW, 元高 = sh0.rowH;
  sh0.colW = {}; sh0.rowH = {};
  for (let c = 0; c < 400; c++) if (設定.幅) sh0.colW[c] = 設定.幅;
  for (let r = 0; r < 400; r++) if (設定.高さ) sh0.rowH[r] = 設定.高さ;
  win.wrapH = 600; win.wrapW = 900; win.scrollTop = 0; win.scrollLeft = 0;
  let 右 = 0, 下 = 0, n = 0;
  win.ctx = new Proxy(元ctx, {
    get(t, k) {
      if (k === 'fillText') return (s, x, y) => { if (y === win.HDR_H / 2) 右 = Math.max(右, x); else if (x === win.HDR_W - 4) 下 = Math.max(下, y); };
      if (k === 'fillRect') return () => { if (++n > 上限) throw new 描きすぎ('fillRect ' + 上限 + ' 回を越えた'); };
      return t[k];
    },
    set() { return true; },
  });
  try { win._doRender(); }
  finally { win.ctx = 元ctx; win.wrapH = 元H; win.wrapW = 元W; win.scrollTop = 元T; win.scrollLeft = 元L; sh0.colW = 元幅; sh0.rowH = 元高; }
  return { 右, 下 };
}
T('★細い列（幅24）でも 列の字が 画面の右の端まで 描かれる★', () => {
  reset();
  const { 右 } = 見出しの字({ 幅: 24 });
  ok(右 >= 900 - 24, '★列の字の 一番右が ' + 右 + '＝画面の右（900）まで 描いていない★');
});
T('★低い行（高さ8）でも 行の番号が 画面の下の端まで 描かれる★', () => {
  reset();
  const { 下 } = 見出しの字({ 高さ: 8 });
  ok(下 >= 600 - 8, '★行の番号の 一番下が ' + 下 + '＝画面の下（600）まで 描いていない★');
});
T('★既定の大きさでも 今まで通り 端まで 描かれる★', () => {
  reset();
  const { 右, 下 } = 見出しの字({});
  ok(右 >= 900 - win.COL_W && 下 >= 600 - win.ROW_H, '既定で 端まで 届いていない 右' + 右 + ' 下' + 下);
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
    /* ★2026-08-31：右クリックの 中身の 正本が 画面から lib/ctx-menu.js へ 移った★
       ★壊す 先も 一緒に 移す★（印が 古いままだと ★素通り★＝見張りが 眠る） */
    ['lib/ctx-menu.js', '右クリックの 固定を消す',
      (s) => s.replace("画面: 'freezePanes'", "画面: 'ctxCopy'")],
    ['book.html', '★A1 でも固定してしまう★（真似ても合わない物を真似る）', (s) => s.replace("  if(selR1 === 0 && selC1 === 0){\n    showToast('A1 では固定できません（固定したい行の1つ下・列の1つ右を選んでね）');\n    return;\n  }", '')],
    ['book.html', '固定する行を 1つ ずらす', (s) => s.replace('  sh.freezeRow = selR1;                    /* 選んだセルの ★上の行★ を固定 */', '  sh.freezeRow = selR1 + 1;')],
    ['book.html', '固定する列を 1つ ずらす', (s) => s.replace('  sh.freezeCol = selC1;                    /* 選んだセルの ★左の列★ を固定 */', '  sh.freezeCol = selC1 + 1;')],
    ['book.html', '取り消しの控えを取らない', (s) => s.replace('  _pushRowColUndo();                       /* ★取り消し1回で戻る★ */', '')],
    ['book.html', '★押した所の読み取りが 帯を見ない★（固定した行を押せない）', (s) => s.replace('  var fz = 帯を見ない ? 0 : (sheets[activeSheet].freezeRow||0);\n  if(fz > 0){', '  var fz = 0;\n  if(fz > 0){')],
    ['book.html', '★左の帯の読み取りが 効かない★', (s) => s.replace('  var fzc = 帯を見ない ? 0 : (sheets[activeSheet].freezeCol||0);\n  if(fzc > 0){', '  var fzc = 0;\n  if(fzc > 0){')],
    ['book.html', '★帯をもう一度 描かない★（固定しても画面が変わらない）', (s) => s.replace('  _renderPass();                                       /* ① 流れる所（今までどおり） */', '  _renderPass(); return;')],
    ['book.html', '★描き始めを 帯の行から取る★（固定して下へ動かすと 帯の下が空白）', (s) => s.replace('var sC=xToC(HDR_W, true), sR=yToR(HDR_H, true)', 'var sC=xToC(HDR_W, true), sR=yToR(HDR_H)')],
    ['book.html', '★描き始めを 帯の列から取る★（固定して右へ動かすと 帯の右が空白）', (s) => s.replace('var sC=xToC(HDR_W, true), sR=yToR(HDR_H, true)', 'var sC=xToC(HDR_W), sR=yToR(HDR_H, true)')],
    ['book.html', '★PageDown が 帯の0行目を一番上と取る★', (s) => s.replace('var top = yToR(HDR_H+_freezeH()+1, true)', 'var top = yToR(HDR_H+1)')],
    ['book.html', '★PageDown が 帯の高さを引かない★（帯に隠れた行を毎回飛ばす）', (s) => s.replace('var top = yToR(HDR_H+_freezeH()+1, true)', 'var top = yToR(HDR_H+1, true)')],
    ['book.html', '★帯を見ない が 効かない★（yToR）', (s) => s.replace("var fz = 帯を見ない ? 0 : (sheets[activeSheet].freezeRow||0);", "var fz = sheets[activeSheet].freezeRow||0;")],
    ['book.html', '★固定の帯の高さを 画面で打ち切らない★', (s) => s.replace('var _fH=Math.min(_fHall, Math.max(0,_wH-HDR_H)), _fW=Math.min(_fWall, Math.max(0,_wW-HDR_W));', 'var _fH=_fHall, _fW=Math.min(_fWall, Math.max(0,_wW-HDR_W));')],
    ['book.html', '★固定の帯の幅を 画面で打ち切らない★', (s) => s.replace('var _fH=Math.min(_fHall, Math.max(0,_wH-HDR_H)), _fW=Math.min(_fWall, Math.max(0,_wW-HDR_W));', 'var _fH=Math.min(_fHall, Math.max(0,_wH-HDR_H)), _fW=_fWall;')],
    ['book.html', '★帯を 打ち切りすぎる★（帯の下に 流れる所が出る）', (s) => s.replace('var _fH=Math.min(_fHall, Math.max(0,_wH-HDR_H)),', 'var _fH=Math.min(_fHall, Math.max(0,_wH-HDR_H-200)),')],
    ['book.html', '★画面より小さい固定まで 削る★', (s) => s.replace('var _fH=Math.min(_fHall, Math.max(0,_wH-HDR_H)), _fW=Math.min(_fWall, Math.max(0,_wW-HDR_W));', 'var _fH=Math.min(_fHall-1, Math.max(0,_wH-HDR_H)), _fW=Math.min(_fWall-1, Math.max(0,_wW-HDR_W));')],
    ['book.html', '★横の境目の線を 打ち切った高さで引く★（画面の下の端に線が出る）', (s) => s.replace('var _ly=Math.round(HDR_H+_fHall)+0.5', 'var _ly=Math.round(HDR_H+_fH)+0.5')],
    ['book.html', '★縦の境目の線を 打ち切った幅で引く★', (s) => s.replace('var _lx=Math.round(HDR_W+_fWall)+0.5', 'var _lx=Math.round(HDR_W+_fW)+0.5')],
    ['book.html', '★分割の側を 打ち切りすぎる★', (s) => s.replace('var _h = Math.min(分割の高さ(), Math.max(0, _wH - HDR_H)),', 'var _h = Math.min(分割の高さ(), Math.max(0, _wH - HDR_H - 200)),')],
    ['book.html', '★分割の側を 画面で打ち切らない★', (s) => s.replace('var _h = Math.min(分割の高さ(), Math.max(0, _wH - HDR_H)), _w = Math.min(分割の幅(), Math.max(0, _wW - HDR_W));', 'var _h = 分割の高さ(), _w = 分割の幅();')],
    ['book.html', '★見出しの強調（行）を 打ち切らない★', (s) => s.replace('for(var r=Math.max(sR,selR1);r<=Math.min(eR+1,selR2);r++) ctx.fillRect(0,rowY(r),HDR_W,rH(r));', 'for(var r=selR1;r<=selR2;r++) ctx.fillRect(0,rowY(r),HDR_W,rH(r));')],
    ['book.html', '★見出しの強調（列）を 打ち切らない★', (s) => s.replace('for(var c=Math.max(sC,selC1);c<=Math.min(eC+1,selC2);c++) ctx.fillRect(colX(c),0,cW(c),HDR_H);', 'for(var c=selC1;c<=selC2;c++) ctx.fillRect(colX(c),0,cW(c),HDR_H);')],
    ['book.html', '★足した範囲の見出し（行）を 打ち切らない★', (s) => s.replace('for(var _hrr=Math.max(sR,_hr.r1);_hrr<=Math.min(eR+1,_hr.r2);_hrr++)', 'for(var _hrr=_hr.r1;_hrr<=_hr.r2;_hrr++)')],
    ['book.html', '★足した範囲の見出し（列）を 打ち切らない★', (s) => s.replace('for(var _hc=Math.max(sC,_hr.c1);_hc<=Math.min(eC+1,_hr.c2);_hc++)', 'for(var _hc=_hr.c1;_hc<=_hr.c2;_hc++)')],
    ['book.html', '★見出しの強調を 打ち切りすぎる（下）★', (s) => s.replace('r<=Math.min(eR+1,selR2);r++) ctx.fillRect(0,rowY(r)', 'r<=Math.min(eR-3,selR2);r++) ctx.fillRect(0,rowY(r)')],
    ['book.html', '★見出しの強調を 打ち切りすぎる（上）★', (s) => s.replace('for(var r=Math.max(sR,selR1);r<=Math.min(eR+1,selR2);r++) ctx.fillRect(0,rowY(r)', 'for(var r=Math.max(sR+1,selR1);r<=Math.min(eR+1,selR2);r++) ctx.fillRect(0,rowY(r)')],
    ['book.html', '★見出しの強調（列）を 打ち切りすぎる（右）★', (s) => s.replace('c<=Math.min(eC+1,selC2);c++) ctx.fillRect(colX(c),0', 'c<=Math.min(eC-3,selC2);c++) ctx.fillRect(colX(c),0')],
    ['book.html', '★歩く数を 既定の高さで 数える★（低い行で 下が白い）', (s) => s.replace('Math.ceil(wrapH/Math.max(1,_低い*scale))', 'Math.ceil(wrapH/Math.max(1,ROW_H*scale))')],
    ['book.html', '★歩く数を 既定の幅で 数える★（細い列で 右が白い）', (s) => s.replace('Math.ceil(wrapW/Math.max(1,_細い*scale))', 'Math.ceil(wrapW/Math.max(1,COL_W*scale))')],
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
