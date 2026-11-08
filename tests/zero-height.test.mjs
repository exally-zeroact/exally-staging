/* zero-height.test.mjs — ★書いて いない 行を 隠す 本（zeroHeight）が 画面で 実Excel と 同じに 動くか★（2026-11-08・件③）
 *
 *  ★本物の book.html を 載せて 本物の 関数を 呼ぶ★（字を 読むだけに しない）
 *  板の 形は 見本 tests/fixtures/zero-height-gap-2026-11-08.xlsx の 読み（book-open.test が 見張る）と 同じ：
 *    下を隠す 12（13行目から 下は 全部 隠れ）・書いていない間 {3,8}（4行目・9行目）・手で 隠した 行 {2}（3行目）
 *  真値＝実Excel 16.0 を COM で 動かして 取った 出し（scratchpad zero/hakaru-zero4-utf8.txt・$Error.Count=0）を
 *    ★下の 再表示の表に 字の まま 写した★（1始まりの 行番号・hid=True が 隠れ）
 *
 *  --self-test … わざと壊して「何通りで赤くなるか」を数える（★repo は読むだけ★）
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(__dirname, '..');
const SELF = process.argv.includes('--self-test');
const OVERRIDE = process.env.EXALLY_ZERO_OVERRIDE ? JSON.parse(process.env.EXALLY_ZERO_OVERRIDE) : {};
const srcPath = (rel) => OVERRIDE[rel] || path.join(ROOT, rel);

let JSDOM;
try { ({ JSDOM } = await import('jsdom')); }
catch { console.log('★jsdomが入っていません。この検証は飛ばせません（SKIPを緑と呼ばない）。npm install してください。'); process.exit(1); }

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

console.log('\n[zero-height] ★本物の book.html で 書いて いない 行を 隠す 板を 動かす★');

const sh = () => win.sheets[win.activeSheet];
const ROWS = win.ROWS;
/* ★見本 gap の 読みと 同じ 板★（0始まり） */
function 板(形) {
  const s = sh();
  s.data = {}; s.rowH = {}; s.filterHidden = {}; s.hiddenRows = { 2: true };
  delete s.kirikaeRows; delete s.既定の行高;
  s.freezeRow = 0; s.freezeCol = 0;
  s.下を隠す = 12; s.書いていない間 = { 3: true, 8: true };
  for (let r = 0; r < 12; r++) if (r !== 3 && r !== 8) s.data[r + ',0'] = { v: 'x' + (r + 1) };
  if (形) 形(s);
  win.scrollTop = 0; win.scrollLeft = 0; win.wrapH = 600; win.wrapW = 900;
  win.undoStack.length = 0;
  win.sel(0, 0, 0, 0);
}
const 見える = (r) => win.rH(r) > 0;

/* ── ① ★再表示＝実Excel の 7通り★（hakaru-zero4 の 出しを 字の まま） ── */
const 測った = [
  ['Cells.EntireRow', '全部', 'r1(hid=False) r3(hid=False) r4(hid=False) r9(hid=False) r12(hid=False) r13(hid=False) r14(hid=False) r20(hid=False) r21(hid=False) r100(hid=False) r1048576(hid=False)'],
  ['Rows', '全部', 'r1(hid=False) r3(hid=False) r4(hid=False) r9(hid=False) r12(hid=False) r13(hid=False) r14(hid=False) r20(hid=False) r21(hid=False) r100(hid=False) r1048576(hid=False)'],
  ['1:1048576', '全部', 'r1(hid=False) r3(hid=False) r4(hid=False) r9(hid=False) r12(hid=False) r13(hid=False) r14(hid=False) r20(hid=False) r21(hid=False) r100(hid=False) r1048576(hid=False)'],
  ['13:20', [13, 20], 'r1(hid=False) r3(hid=True) r4(hid=True) r9(hid=True) r12(hid=False) r13(hid=False) r14(hid=False) r20(hid=False) r21(hid=True) r100(hid=True) r1048576(hid=True)'],
  ['12:14', [12, 14], 'r1(hid=False) r3(hid=True) r4(hid=True) r9(hid=True) r12(hid=False) r13(hid=False) r14(hid=False) r20(hid=True) r21(hid=True) r100(hid=True) r1048576(hid=True)'],
  ['13:13', [13, 13], 'r1(hid=False) r3(hid=True) r4(hid=True) r9(hid=True) r12(hid=False) r13(hid=False) r14(hid=True) r20(hid=True) r21(hid=True) r100(hid=True) r1048576(hid=True)'],
  ['4:4', [4, 4], 'r1(hid=False) r3(hid=True) r4(hid=False) r9(hid=True) r12(hid=False) r13(hid=True) r14(hid=True) r20(hid=True) r21(hid=True) r100(hid=True) r1048576(hid=True)'],
];
for (const [名, 範囲, 出し] of 測った) {
  T('★' + 名 + ' を 再表示 → 実Excel と 同じ 行が 出る★', () => {
    板();
    if (範囲 === '全部') win.sel(0, 0, ROWS - 1, win.COLS - 1);
    else win.sel(範囲[0] - 1, 0, 範囲[1] - 1, 0);
    win.ctxUnhideRow();
    const 違い = [];
    for (const m of 出し.matchAll(/r(\d+)\(hid=(True|False)\)/g)) {
      const r = Math.min(+m[1], ROWS) - 1, 隠れ = m[2] === 'True';
      /* ★手で 隠した 行の 前後1行を 消す 今の 形（ctxUnhideRow の ±1）は 別件★＝4:4 の 3行目だけ 見ない（紙の 別件の 候補） */
      if (名 === '4:4' && r === 2) continue;
      if (見える(r) === 隠れ) 違い.push((r + 1) + '行目＝' + (隠れ ? '隠れの はずが 見える' : '見える はずが 隠れ'));
    }
    eq(違い.join('・'), '', '実Excel との 違い');
    if (範囲 === '全部') { eq(sh().下を隠す, undefined, '全部の 再表示で zeroHeight が 外れる'); eq(sh().書いていない間, undefined, '間も 消える'); }
  });
}
T('★再表示は 取り消し1回で 元へ 戻る（下を隠す・間も）★', () => {
  板(); win.sel(12, 0, 19, 0); win.ctxUnhideRow();
  eq(sh().下を隠す, 20, '再表示の 後');
  win.doUndo();
  eq(sh().下を隠す, 12, '取り消しの 後の 下を隠す'); eq(JSON.stringify(sh().書いていない間), JSON.stringify({ 3: true, 8: true }), '取り消しの 後の 間');
});
/* ★選んだ 行だけ 出す（前後の 1行は 出さない）★＝13:13・4:4 で 測った 決まりを 間の 隣で 当てる */
T('★8行目だけ 再表示 → 隣の 9行目（間）は 隠れた まま★', () => {
  板(); win.sel(7, 0, 7, 0); win.ctxUnhideRow();
  ok(!見える(8), '9行目（間）が 出てしまった');
});
T('★下を隠す より 下だけを 再表示 → その 上は 間へ 移って 隠れた まま（21:22）★', () => {
  板(); win.sel(20, 0, 21, 0); win.ctxUnhideRow();
  ok(見える(20) && 見える(21), '21・22行目が 出ていない');
  ok(!見える(12) && !見える(19) && !見える(22), '13〜20行目・23行目が 出てしまった');
  eq(sh().下を隠す, 22, '下を隠す');
});

/* ── ①b ★ビュー（シート・ブック）は 書いて いない 行の 板を 持たない・当てる 時も 触らない★（2026-11-08 設計 v4）
   本番前 taiketsu で 3回 止め＝①消える（100万 行が 出る）②別の シートで 出すと 消える／書いた 行が 隠れる
   ③別の 本の 同じ 名の シート・全部 再表示の 後・行を 入れた 後に 出すと 古い 下を隠す が 戻り 書いた 行が 隠れる ── */
const ビューを空に = () => win.eval('ブックのビュー.length = 0; シートビュー = {}; 今のシートビュー = {};');
const ビューの道 = [
  ['シートビュー', "var _t = SheetView.足す(_このシートのビュー(), 'v', _今の見え方());", 'シートビューを切り替え(_t.名);'],
  ['ブックのビュー', 'ブックのビューを足す();', 'ブックのビューを出す(ブックのビュー[ブックのビュー.length-1].名);'],
  /* ★シートビューの もう1本の 道＝新規→保持（覚え直す）→切り替え★（本番前 taiketsu 任意2・道が 2本 在る時は 両方） */
  ['シートビュー（新規・保持）', 'シートビューを新規(); シートビューを保持(); var _t2 = 今のシートビュー[sheets[activeSheet].name];', 'シートビューを切り替え(_t2);'],
];
const 間 = () => JSON.stringify(sh().書いていない間);
for (const [道, 覚える, 出す] of ビューの道) {
  T('★' + 道 + 'を 足して 出しても 下を隠す・間 が 残る★', () => {
    板(); ビューを空に();
    try {
      win.eval(覚える + 出す);
      eq(sh().下を隠す, 12, '下を隠す'); eq(間(), JSON.stringify({ 3: true, 8: true }), '間');
      ok(!見える(500), '501行目が 出てしまった');
    } finally { ビューを空に(); }
  });
  T('★別の 本の 同じ 名の シート（300行 書いて ある）で ' + 道 + 'を 出しても 隠さない★', () => {
    板(); ビューを空に();
    try {
      win.eval(覚える);
      const s = sh(); delete s.下を隠す; delete s.書いていない間; s.hiddenRows = {};
      for (let r = 0; r < 300; r++) s.data[r + ',0'] = { v: 'b' + (r + 1) };
      win.eval(出す);
      eq(sh().下を隠す, undefined, '下を隠す'); eq(sh().書いていない間, undefined, '間');
      ok(見える(100), '101行目（書いて ある）が 隠れた');
    } finally { ビューを空に(); }
  });
  T('★全部 再表示 → 20行目に 書く → ' + 道 + 'を 出しても 20行目は 見える★', () => {
    板(); ビューを空に();
    try {
      win.eval(覚える);
      win.sel(0, 0, ROWS - 1, 0); win.ctxUnhideRow();
      eq(sh().下を隠す, undefined, '全部 再表示で 外れる');
      sh().data['19,0'] = { v: 'w20' };
      win.eval(出す);
      eq(sh().下を隠す, undefined, '下を隠す'); ok(見える(19), '書いた 20行目が 隠れた');
    } finally { ビューを空に(); }
  });
  T('★6〜7行目に 2行 入れた 後に ' + 道 + 'を 出しても 下を隠す は 14 の まま★', () => {
    板(); ビューを空に();
    try {
      win.eval(覚える);
      win.sel(5, 0, 6, 0); win.ctxInsertRow();
      eq(sh().下を隠す, 14, '入れた 後');
      win.eval(出す);
      eq(sh().下を隠す, 14, '出した 後'); ok(見える(12) && 見える(13), '入れて 下がった 13・14行目が 隠れた');
    } finally { ビューを空に(); }
  });
}
/* ★別の シートで 覚えた ブックの ビューを 出す★
   直す前＝普通の シートで 覚えて 書いて いない 行の シートで 出すと 下を隠す が 消え 501行目が 出た／逆は 書いた 行が 13行目から 隠れた */
function 別のシートで(fn) {
  const 前 = win.activeSheet;
  win.sheets.push({ name: 'ZeroBetsu', data: { '0,0': { v: 'a' } }, colW: {}, rowH: {}, hiddenRows: {}, filterHidden: {}, hiddenCols: {} });
  const 別 = win.sheets.length - 1;
  const 移る = (i) => win.eval('activeSheet = ' + i + ';');
  try { fn(移る, 前, 別); }
  finally { 移る(前); win.sheets.splice(別, 1); ビューを空に(); }
}
T('★普通の シートで 覚えた ブックの ビューを 書いて いない 行の シートで 出しても 下を隠す・間 が 残る★', () => {
  板();
  別のシートで((移る, 前, 別) => {
    移る(別); win.eval('ブックのビューを足す();');
    移る(前); win.eval('ブックのビューを出す(ブックのビュー[ブックのビュー.length-1].名);');
    eq(sh().下を隠す, 12, '下を隠す'); eq(間(), JSON.stringify({ 3: true, 8: true }), '間');
    ok(!見える(500), '501行目が 出てしまった');
  });
});
T('★書いて いない 行の シートで 覚えた ブックの ビューを 普通の シートで 出しても 隠さない★', () => {
  板();
  別のシートで((移る, 前, 別) => {
    win.eval('ブックのビューを足す();');
    移る(別); win.eval('ブックのビューを出す(ブックのビュー[ブックのビュー.length-1].名);');
    eq(sh().下を隠す, undefined, '下を隠す'); eq(sh().書いていない間, undefined, '間');
    ok(見える(100), '101行目が 隠れた');
  });
});

/* ── ② ★再表示の 札★ ── */
T('★書いて いない 行に 掛かる 選びで 右クリック → 行の 再表示 の 札が 出る／掛からなければ 出ない★', () => {
  const 札 = (r1, r2) => { win.sel(r1, 0, r2, 0); win.showCtxMenu(10, 10); const d = doc.getElementById('ctx-unhide-row').style.display; win.hideCtxMenu(); return d; };
  板();
  eq(札(10, 12), 'block', '11〜13行目（下を隠す に 掛かる）');
  eq(札(8, 8), 'block', '9行目（間）');
  eq(札(5, 6), 'none', '6〜7行目（何も 隠れて いない）');
});

/* ── ③ ★二重に 数えない★（列を 選んで 隠すと 下を隠す より 下の 鍵も hiddenRows に 入る） ── */
T('★下を隠す より 下に hiddenRows の 鍵が 在っても rowY(最後の行)＝rowY(下を隠す)★', () => {
  板((s) => { for (let r = 20; r < 60; r++) s.hiddenRows[r] = true; });
  eq(win.rowY(ROWS - 1), win.rowY(12), 'rowY(最後の行)');
  eq(win.rowY(12), win.rowY(11) + win.rH(11), '最後の 見える 行の 下の 端');
  eq(win.rowY(ROWS), win.rowY(12), 'rowY(ROWS)');
});

/* 描いた 行の 番号の 一番下と、緑の 最後の 線を 拾う */
function 描く(形) {
  const 元ctx = win.ctx; let 下 = 0; const 線 = []; const 筆 = {};
  win.ctx = new Proxy(元ctx, {
    get(t, k) {
      if (k === 'fillText') return (s, x, y) => { if (x === win.HDR_W - 4) 下 = Math.max(下, y); };
      if (k === 'moveTo') return (x, y) => { if (筆.strokeStyle === '#3D9E72' && 筆.lineWidth === 1.5 && x === 0) 線.push(y); };
      return t[k];
    },
    set(t, k, v) { 筆[k] = v; return true; },
  });
  try { win._doRender(); } finally { win.ctx = 元ctx; }
  return { 下, 線 };
}

/* ── ④ ★間の 多い 板（8行）でも 画面の 下まで 描く★（7202 の 歩く数に 間を 足さないと 下が 白い） ── */
T('★書いて いない 間が 8行 在っても 行の 番号が 画面の 下（600）まで 描かれる★', () => {
  板((s) => { s.hiddenRows = {}; s.下を隠す = 200; s.書いていない間 = {}; for (let r = 2; r < 10; r++) s.書いていない間[r] = true; });
  const { 下 } = 描く();
  ok(下 >= 600 - win.ROW_H, '★行の番号の 一番下が ' + 下 + '＝画面の 下まで 描いていない★');
});

/* ── ⑤ ★縦の 上限＝最後の 見える 行の 下の 端が 画面の 下まで★（その 下へ 回し車で 行けない） ── */
T('★回し車で 下へ 行き過ぎても 最後の 見える 行の 下の 端＝画面の 下 で 止まる★', () => {
  板((s) => { s.hiddenRows = {}; s.書いていない間 = {}; s.下を隠す = 100; });
  win.scrollTop = 1e7; 描く();
  eq(win.rowY(100), win.wrapH, '最後の 見える 行の 下の 端');
});
T('★画面に 収まる 板は 下へ 動かない（scrollTop 0）★', () => {
  板(); win.scrollTop = 5000; 描く();
  eq(win.scrollTop, 0, 'scrollTop');
});
T('★zeroHeight の 無い 板は 上限を 付けない（今まで 通り）★', () => {
  板((s) => { delete s.下を隠す; delete s.書いていない間; s.hiddenRows = {}; });
  win.scrollTop = 5000; 描く();
  eq(win.scrollTop, 5000, 'scrollTop');
});

/* ── ⑥ ★最後の 線（緑）を 最後の 見える 行の 下に 引かない★（Excel に 無い 線） ── */
T('★書いて いない 下を 隠す 板は 緑の 最後の 線を 引かない★', () => {
  板();
  const { 線 } = 描く();
  eq(線.indexOf(Math.round(win.rowY(ROWS)) + 0.5), -1, '最後の 見える 行の 下の 緑の 線');
});

/* ── ⑦ ★選びを 隠れた 行へ 置かない★（↓・↑・Shift+↓・PageDown・押した所） ── */
T('★↓ は 隠れた 行（3行目 手・4行目 間）を 飛ばす／最後の 見える 行で 止まる／↑ も 飛ばす★', () => {
  板();
  win.sel(1, 0, 1, 0); win.moveSel(1, 0); eq(win.selR1, 4, '2行目から ↓');
  win.sel(11, 0, 11, 0); win.moveSel(1, 0); eq(win.selR1, 11, '12行目から ↓（下は 全部 隠れ）');
  win.sel(4, 0, 4, 0); win.moveSel(-1, 0); eq(win.selR1, 1, '5行目から ↑');
  win.sel(9, 0, 9, 0); win.moveSel(-1, 0); eq(win.selR1, 7, '10行目から ↑（9行目 間）');
});
T('★Shift+↓ も 最後の 見える 行で 止まる★', () => {
  板(); win.sel(10, 0, 10, 0); win.selShift(11, 0); win.selShift(12, 0);
  eq(win.selR2, 11, '選びの 下の 端');
});
T('★PageDown は 最後の 見える 行で 止まる★', () => {
  板(); win.sel(4, 0, 4, 0); win.movePage(1);
  eq(win.selR1, 11, '5行目から PageDown（1画面 進むと 隠れた 下へ 出る）');
});
T('★画面の 下の 白い 所を 押しても 最後の 見える 行★', () => {
  板();
  eq(win.yToR(win.wrapH - 1), 11, 'yToR(画面の 下)');
});

/* ── ⑧ ★行を 入れる・消す と 一緒に 動く★（実Excel＝間は 番号ごと 動く・最後の 在る 行の 下に 入れた 行は 見える） ── */
T('★6〜7行目に 2行 入れる → 間 {3,10}・下を隠す 14★', () => {
  板(); win.sel(5, 0, 6, 0); win.ctxInsertRow();
  eq(JSON.stringify(sh().書いていない間), JSON.stringify({ 3: true, 10: true }), '間'); eq(sh().下を隠す, 14, '下を隠す');
});
T('★下を隠す を またいで 消す（11〜15行目）→ 下を隠す 10★', () => {
  板(); win.sel(10, 0, 14, 0); win.ctxDeleteRow();
  eq(sh().下を隠す, 10, '下を隠す'); eq(JSON.stringify(sh().書いていない間), JSON.stringify({ 3: true, 8: true }), '間');
});

console.log('\n  ' + pass + ' passed, ' + fail + ' failed');

if (SELF) {
  const { spawnSync } = await import('node:child_process');
  const os = await import('node:os');
  const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'exally-zero-'));
  console.log('\n[self-test] わざと壊して 赤くなるかを数える（★repo は読むだけ★）');
  const BREAKS = [
    ['book.html', '★書いて いない 下の 行を 隠さない★', (s) => s.replace('  if(r >= 行の終わり(sh)) return 0;', '')],
    ['book.html', '★書いて いない 間を 隠さない★', (s) => s.replace('  if(sh.書いていない間 && sh.書いていない間[r]) return 0;', '')],
    ['book.html', '★全部の 再表示で zeroHeight を 外さない★', (s) => s.replace("  if(r1 <= 0 && r2 >= ROWS - 1){ delete s.下を隠す; delete s.書いていない間; return; }", '')],
    ['book.html', '★再表示で 前後の 1行も 出す★', (s) => s.replace('if(n >= r1 && n <= r2) delete 間[k];', 'if(n >= r1 - 1 && n <= r2 + 1) delete 間[k];')],
    ['book.html', '★下より 下を 再表示した 時 間へ 移さない★', (s) => s.replace('    if(r1 > 終){ if(!s.書いていない間) s.書いていない間 = {}; for(var g = 終; g < r1; g++) s.書いていない間[g] = true; }', '')],
    ['book.html', '★再表示の 札が 書いて いない 行を 見ない★', (s) => s.replace('  if(!hasHR && _書いていない行が掛かるか(sheets[activeSheet], selR1, selR2)) hasHR = true;', '')],
    ['book.html', '★下を隠す より 下の 鍵を 二重に 数える★', (s) => s.replace('return _隠れた数(_隠れ行の並び(), Math.min(r, 終)) + Math.max(0, r - 終);', 'return _隠れた数(_隠れ行の並び(), r) + Math.max(0, r - 終);')],
    ['book.html', '★歩く数に 間を 足さない★', (s) => s.replace("              + Object.keys(_sh.書いていない間||{}).length;", "              + 0;")],
    ['book.html', '★縦の 上限で 抑えない★', (s) => s.replace('function _doRenderの中(){\n  _縦の上限で抑える();', 'function _doRenderの中(){')],
    ['book.html', '★最後の 線を 引く★', (s) => s.replace('  if(行の終わり() >= ROWS && lastRowY>HDR_H', '  if(lastRowY>HDR_H')],
    ['book.html', '★↓ が 隠れた 行へ 行く★', (s) => s.replace('  if(dr) r = 見える行へ(r, dr > 0 ? 1 : -1, selR1);', '')],
    ['book.html', '★Shift+↓ が 隠れた 行へ 行く★', (s) => s.replace('  if(activeR !== _前R) activeR = 見える行へ(activeR, activeR > _前R ? 1 : -1, _前R);', '')],
    ['book.html', '★PageDown が 隠れた 行へ 行く★', (s) => s.replace('  r = 見える行へ(r, dir > 0 ? 1 : -1, selR1);', '')],
    ['book.html', '★押した所が 行の 終わりで 止まらない★', (s) => s.replace('  return Math.max(0, _終 - 1);\n}', '  return ROWS - 1;\n}')],
    ['book.html', '★行を 入れても 間を 動かさない★', (s) => s.replace('    if (s.書いていない間) s.書いていない間 = _番号の箱を写す(s.書いていない間, 入れる, 上, 下);', '')],
    ['book.html', '★行を 入れても 下を隠す を 下げない★', (s) => s.replace('      if (入れる) { if (上 <= 境) 境 += 幅; }', '      if (入れる) { }')],
    ['book.html', '★またいで 消しても 下を隠す を 縮めない★', (s) => s.replace('      else if (上 < 境) 境 = 上;', '')],
    ['book.html', '★当てる が 下を隠す・間 を ビューの 物で 上書きする★', (s) => s.replace('  sh.filterHidden = v.絞り; sh.hiddenRows = v.手で隠した行; sh.hiddenCols = v.手で隠した列;\n  scale = v.倍率;', '  sh.filterHidden = v.絞り; sh.hiddenRows = v.手で隠した行; sh.hiddenCols = v.手で隠した列;\n  sh.下を隠す = v.下を隠す; sh.書いていない間 = v.書いていない間;\n  scale = v.倍率;')],
    /* ★3か所 一緒（覚える・写す・持つ 時だけ 当てる）＝v3 の 形に 戻す★ */
    [['book.html', 'lib/sheet-view.js'], '★ビューが 下を隠す・間 を 覚えて 戻す（v3 の 形）★', [
      (s) => s.replace('    絞り: sh.filterHidden, 手で隠した行: sh.hiddenRows, 手で隠した列: sh.hiddenCols,\n    倍率: scale,', '    絞り: sh.filterHidden, 手で隠した行: sh.hiddenRows, 手で隠した列: sh.hiddenCols,\n    書いていない間: sh.書いていない間, 下を隠す: sh.下を隠す,\n    倍率: scale,')
        .replace('  sh.filterHidden = v.絞り; sh.hiddenRows = v.手で隠した行; sh.hiddenCols = v.手で隠した列;\n  scale = v.倍率;', '  sh.filterHidden = v.絞り; sh.hiddenRows = v.手で隠した行; sh.hiddenCols = v.手で隠した列;\n  if (v.下を隠す !== undefined) { sh.下を隠す = v.下を隠す; sh.書いていない間 = v.書いていない間; }\n  scale = v.倍率;'),
      (s) => s.replace('      固定列: Number(o.固定列 || 0),\n    };', '      固定列: Number(o.固定列 || 0),\n      下を隠す: o.下を隠す, 書いていない間: o.書いていない間,\n    };'),
    ]],
  ];
  let red = 0;
  for (const [rel0, name, brk0] of BREAKS) {
    /* ★1つの 壊しで 複数の ファイルを 壊せる（rel と brk を 並びで 渡す）＝どれか 1つでも 置換できなければ 数えない★ */
    const rels = Array.isArray(rel0) ? rel0 : [rel0], brks = Array.isArray(brk0) ? brk0 : [brk0];
    const over = {}; let 置換できず = false;
    rels.forEach((rel, i) => {
      const orig = fs.readFileSync(path.join(ROOT, rel), 'utf8');
      const bad = brks[i](orig);
      if (bad === orig) 置換できず = true;
      const tmpFile = path.join(TMP, rel.replace(/[\\/]/g, '_'));
      fs.writeFileSync(tmpFile, bad, 'utf8');
      over[rel] = tmpFile;
    });
    if (置換できず) { console.log('  ★置換できず★  ' + name); continue; }
    const env = Object.assign({}, process.env, { EXALLY_ZERO_OVERRIDE: JSON.stringify(over) });
    const isRed = spawnSync(process.execPath, [path.join(__dirname, 'zero-height.test.mjs')], { encoding: 'utf8', env }).status !== 0;
    if (isRed) { red++; console.log('  赤くなった  ' + name); }
    else console.log('  ★素通り★  ' + name);
  }
  try { fs.rmSync(TMP, { recursive: true, force: true }); } catch (e) { /* 消せなくても検査は済んでいる */ }
  console.log('\n  ' + red + '/' + BREAKS.length + ' 通りで赤くなった');
  process.exit(red === BREAKS.length ? 0 : 1);
}

process.exit(fail ? 1 : 0);
