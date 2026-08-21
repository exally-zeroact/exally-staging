// excel-shortcuts.test.mjs — ★本物の book.html に 実際にキーを送って ショートカットを確かめる★
//   真値は tests/fixtures/excel-shortcuts-golden.json（実Excel 16.0.20228 から機械で取った物）。
//   ★ソースを読むだけにしない。実際に keydown を送って 画面の中の値が変わったかを見る★
//   --self-test … わざと壊して「何通りで赤くなるか」を数える（見張りが本当に見張っているか）
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(__dirname, '..');
const SELF = process.argv.includes('--self-test');

let JSDOM;
try { ({ JSDOM } = await import('jsdom')); }
catch { console.log('★jsdomが入っていません。この検証は飛ばせません（SKIPを緑と呼ばない）。npm install してください。'); process.exit(1); }

const GOLD = JSON.parse(fs.readFileSync(path.join(ROOT, 'tests/fixtures/excel-shortcuts-golden.json'), 'utf8'));

let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ok   ' + n); } catch (e) { fail++; console.log('  NG   ' + n + '\n       ' + (e && e.message)); } };
const ok = (c, m) => { if (!c) throw new Error(m || 'expected truthy'); };
const eq = (a, b, m) => { if (a !== b) throw new Error((m ? m + ': ' : '') + 'expected ' + JSON.stringify(b) + ' got ' + JSON.stringify(a)); };

const CANVAS_STUB = [
  '(function(){',
  '  var noop=function(){};',
  '  var ctx=new Proxy({}, { get:function(t,k){',
  '    if(k==="measureText") return function(){ return {width:40}; };',
  '    if(k==="canvas") return {width:900,height:600};',
  '    if(k==="getImageData") return function(){ return {data:[]}; };',
  '    if(k==="createLinearGradient"||k==="createPattern") return function(){ return {addColorStop:noop}; };',
  '    return noop;',
  '  }});',
  '  HTMLCanvasElement.prototype.getContext=function(){ return ctx; };',
  '})();',
].join('\n');

/* ★わざと壊す時も repo のファイルは1バイトも書き換えない★（2026-08-18 実際に踏んだ）
   前は book.html を直接 壊して→戻す形にしていた。すると
     ① 検査を2本 同時に走らせると、片方の「壊した状態」をもう片方が読んで ★嘘の赤★ が出る
     ② 途中で止めると ★壊した行が repo に残る★（実際に `Alt+= を測らずに作る` の行が残った）
   ⇒ 壊した中身は ★temp に置いて★、子の検査に「こっちを読め」と env で渡す。
      repo は読むだけ＝同時に何本走っても安全・途中で止めても残らない。 */
const OVERRIDE = process.env.EXALLY_SHORTCUT_OVERRIDE ? JSON.parse(process.env.EXALLY_SHORTCUT_OVERRIDE) : {};
const srcPath = (rel) => OVERRIDE[rel] || path.join(ROOT, rel);

function bootPage() {
  const html = fs.readFileSync(srcPath('book.html'), 'utf8');
  const dom = new JSDOM(html.replace(/<script[\s\S]*?<\/script>/g, ''), {
    runScripts: 'dangerously', url: 'http://localhost/', pretendToBeVisual: true,
    beforeParse(w) {
      w.fetch = () => Promise.reject(new Error('no net'));
      w.scrollTo = () => {}; w.alert = () => {};
      w.matchMedia = w.matchMedia || (() => ({ matches: false, addListener() {}, removeListener() {}, addEventListener() {}, removeEventListener() {} }));
      w.requestAnimationFrame = (cb) => setTimeout(() => cb(1), 0);
      w.cancelAnimationFrame = () => {};
      w.eval(CANVAS_STUB);
    },
  });
  const win = dom.window, doc = win.document;
  const inject = (code) => { const el = doc.createElement('script'); el.textContent = code; doc.body.appendChild(el); };
  const srcs = [...html.matchAll(/<script src="([^"]+)"><\/script>/g)].map((m) => m[1].split('?')[0]).filter((s) => !/^https?:/.test(s));
  const loaded = [], missing = [];
  for (const src of srcs) {
    const p = srcPath(src);
    if (!fs.existsSync(p)) { missing.push(src); continue; }
    inject(fs.readFileSync(p, 'utf8'));
    loaded.push(src);
  }
  let inlineNG = 0;
  for (const m of html.matchAll(/<script(?![^>]*\ssrc=)[^>]*>([\s\S]*?)<\/script>/g)) {
    try { inject(m[1]); } catch (e) { inlineNG++; }
  }
  try { doc.dispatchEvent(new win.Event('DOMContentLoaded', { bubbles: true })); } catch (e) { /* 続ける */ }
  try { win.dispatchEvent(new win.Event('load')); } catch (e) { /* 続ける */ }
  return { win, doc, loaded, missing, inlineNG };
}

const { win, doc, loaded, missing, inlineNG } = bootPage();
console.log('\n[excel-shortcuts] ★本物の book.html に 実際にキーを送る★');
console.log('  外の部品 ' + loaded.length + '本（読めなかった ' + missing.length + '本）／インラインで流せなかった ' + inlineNG + '本');
console.log('  真値 = ' + GOLD._measured_with);

function key(k, mod) {
  mod = mod || {};
  const ev = new win.KeyboardEvent('keydown', {
    key: k, bubbles: true, cancelable: true,
    ctrlKey: !!mod.ctrl, shiftKey: !!mod.shift, altKey: !!mod.alt, metaKey: !!mod.meta,
  });
  (mod.target || doc).dispatchEvent(ev);
  return ev;
}
const A = (r, c) => win.sheets[win.activeSheet].data[r + ',' + c] || null;
function reset() {
  win.sheets[win.activeSheet].data = {};
  win.sel(0, 0, 0, 0);
  const fi = doc.getElementById('formula-input');
  if (fi) fi.blur();
  doc.body.focus();
}

/* ── 土台が本当に立っているか（立っていなければ この先の緑は全部 嘘） ── */
T('★画面が立ち上がっている（sheets と onKD が在る）', () => {
  ok(Array.isArray(win.sheets) && win.sheets.length > 0, 'sheets が無い＝インラインが途中で止まっている');
  ok(typeof win.onKD === 'function', 'onKD が無い');
});
T('★送ったキーが本当に届いている（受け口が document に付いている）', () => {
  let reached = 0;
  const spy = win.onKD;
  win.onKD = function () { reached++; return spy.apply(this, arguments); };
  doc.addEventListener('keydown', win.onKD);
  key('ArrowDown');
  doc.removeEventListener('keydown', win.onKD);
  win.onKD = spy;
  ok(reached > 0, 'キーが届いていない');
});
T('★4つの助けの関数が本当に在る★（無いと押した瞬間に落ちる）', () => {
  for (const f of ['fillFromEdge', 'putNowValue', 'movePage', 'toggleAbsRef']) {
    ok(typeof win[f] === 'function', f + ' が無い');
  }
});

/* ── ① Ctrl+D（上のセルを写す）── 真値: 式も書式も写す ── */
T('Ctrl+D … 上のセルの★式★を写す（1セル選択）', () => {
  reset();
  win.setCell(0, 0, '=1+2');
  win.sel(1, 0, 1, 0);
  const ev = key('d', { ctrl: true });
  ok(ev.defaultPrevented, 'ブラウザの既定（ブックマーク）を止めていない');
  eq(A(1, 0) && A(1, 0).f, '=1+2');
});
T('Ctrl+D … ★書式も写す★（太字・背景色・表示形式）＝実Excelで実測', () => {
  reset();
  win.setCell(0, 0, '=1+2');
  Object.assign(win.sheets[win.activeSheet].data['0,0'], { bold: true, bgColor: '#FFFF00', numFmt: '#,##0.00' });
  win.sel(1, 0, 1, 0);
  key('d', { ctrl: true });
  const d = A(1, 0);
  eq(!!(d && d.bold), GOLD.fill_down.bold_copied, '太字');
  eq(d && d.numFmt, GOLD.fill_down.numfmt_copied, '表示形式');
  eq(d && d.bgColor, '#FFFF00', '背景色');
});
T('Ctrl+D … ★写した先の元の書式は捨てる★（斜体が残らない）＝実Excelで実測', () => {
  reset();
  win.setCell(0, 0, '9');
  win.setCell(1, 0, '1');
  win.sheets[win.activeSheet].data['1,0'].italic = true;
  win.sel(1, 0, 1, 0);
  key('d', { ctrl: true });
  eq(!!(A(1, 0) && A(1, 0).italic), GOLD.fill_down.italic_of_destination_kept, '先の斜体');
});
T('Ctrl+D … 範囲を選んだ時は ★先頭の行★ を下へ写す', () => {
  reset();
  win.setCell(0, 0, '7');
  win.sel(0, 0, 3, 0);
  key('d', { ctrl: true });
  for (let r = 1; r <= 3; r++) eq(A(r, 0) && String(A(r, 0).v), '7', 'r=' + r);
});
T('Ctrl+D … 1行目（上が無い）では何もしない・落ちない', () => {
  reset();
  win.sel(0, 0, 0, 0);
  key('d', { ctrl: true });
  eq(A(0, 0), null);
});

/* ── ② Ctrl+R（左のセルを写す）── 真値: 参照がずれる ── */
T('Ctrl+R … ★式の中の参照がずれる★（=A2*2 → =B2*2）＝実Excelで実測', () => {
  reset();
  win.setCell(0, 0, GOLD.fill_right.source);
  win.sel(0, 1, 0, 1);
  const ev = key('r', { ctrl: true });
  ok(ev.defaultPrevented, '既定（再読み込み）を止めていない');
  eq(A(0, 1) && A(0, 1).f, GOLD.fill_right.result_at_B1);
});
T('Ctrl+D … 下へ写した式も参照がずれる（=A1+1 → =A2+1）', () => {
  reset();
  win.setCell(0, 1, '=A1+1');
  win.sel(1, 1, 1, 1);
  key('d', { ctrl: true });
  eq(A(1, 1) && A(1, 1).f, '=A2+1');
});

/* ── ③ Ctrl+; / Ctrl+: ── 真値: 打った日付は yyyy/m/d・時刻は h:mm ── */
const goldDateFmt = GOLD.typed_datetime.find((t) => t.typed === '2026/8/18').numfmt;
const goldTimeFmt = GOLD.typed_datetime.find((t) => t.typed === '14:30').numfmt;
T('Ctrl+; … 今日の日付が「値」で入る（式ではない）＋書式 ' + goldDateFmt, () => {
  reset();
  win.sel(2, 2, 2, 2);
  const ev = key(';', { ctrl: true });
  ok(ev.defaultPrevented, '既定を止めていない');
  const c = A(2, 2); ok(c, 'セルが空');
  ok(/^\d{4}\/\d{1,2}\/\d{1,2}$/.test(String(c.v)), '入った物が日付でない: ' + c.v);
  ok(String(c.v).charAt(0) !== '=', '式になっている（Excelは値を入れる）');
  eq(c.numFmt, goldDateFmt, '書式');
});
T('Ctrl+: … 今の時刻が入る＋書式 ' + goldTimeFmt, () => {
  reset();
  win.sel(3, 3, 3, 3);
  key(':', { ctrl: true });
  const c = A(3, 3); ok(c, 'セルが空');
  ok(/^\d{1,2}:\d{2}$/.test(String(c.v)), '入った物が時刻でない: ' + c.v);
  eq(c.numFmt, goldTimeFmt, '書式');
});
T('Ctrl+; … 範囲を選んでいたら 選んだ全部に入る', () => {
  reset();
  win.sel(5, 0, 6, 1);
  key(';', { ctrl: true });
  for (const k of ['5,0', '5,1', '6,0', '6,1']) ok(win.sheets[win.activeSheet].data[k], k + ' が空');
});

/* ── ④ Ctrl+B / I / U / 1（ボタンと同じ物をキーからも） ── */
const FMT_KEYS = [['b', 'bold', '太字'], ['i', 'italic', '斜体'], ['u', 'underline', '下線']];
for (const [k, prop, name] of FMT_KEYS) {
  T('Ctrl+' + k.toUpperCase() + ' … ' + name + 'が付く／もう一度で外れる', () => {
    reset();
    win.setCell(1, 1, 'あ');
    win.sel(1, 1, 1, 1);
    key(k, { ctrl: true });
    eq(!!(A(1, 1) && A(1, 1)[prop]), true, '1回目');
    key(k, { ctrl: true });
    eq(!!(A(1, 1) && A(1, 1)[prop]), false, '2回目');
  });
}
T('Ctrl+1 … セルの書式設定の窓が開く', () => {
  reset();
  win.sel(0, 0, 0, 0);
  const ev = key('1', { ctrl: true });
  ok(ev.defaultPrevented, '既定を止めていない');
  const m = doc.getElementById('fmtModal');
  ok(m, '書式の窓（#fmtModal）が見つからない');
  ok(m.classList.contains('show'), '窓が開いていない（show が付いていない）');
});

/* ── ⑤ 行・列を選ぶ ── */
T('Ctrl+Space … ★列★ を丸ごと選ぶ', () => {
  reset();
  win.sel(3, 2, 3, 2);
  const ev = key(' ', { ctrl: true });
  ok(ev.defaultPrevented, '既定を止めていない');
  eq(win.selR1, 0); eq(win.selR2, win.ROWS - 1);
  eq(win.selC1, 2); eq(win.selC2, 2);
});
T('Shift+Space … ★行★ を丸ごと選ぶ', () => {
  reset();
  win.sel(3, 2, 3, 2);
  const ev = key(' ', { shift: true });
  ok(ev.defaultPrevented, '既定を止めていない');
  eq(win.selC1, 0); eq(win.selC2, win.COLS - 1);
  eq(win.selR1, 3); eq(win.selR2, 3);
});

/* ── ⑥ シート移動・1画面 上下 ── */
T('Ctrl+PageDown / PageUp … シートを移る', () => {
  if (win.sheets.length < 2 && typeof win.addSheet === 'function') win.addSheet();
  ok(win.sheets.length >= 2, 'シートを2枚にできなかった（この確認は飛ばせない）');
  win.switchSheet(0);
  key('PageDown', { ctrl: true });
  eq(win.activeSheet, 1, '次のシートへ');
  key('PageUp', { ctrl: true });
  eq(win.activeSheet, 0, '前のシートへ');
  key('PageUp', { ctrl: true });
  eq(win.activeSheet, 0, '一番左で止まる（0より小さくならない）');
});
T('PageDown / PageUp … 1画面ぶん 下・上へ動く', () => {
  /* ★jsdom には「大きさ」が無いので wrapH（画面の高さ）は 0 のまま★
     ＝そのままだと 1行しか動かず、本当に1画面ぶん動くのかを見られない。
     本物の画面と同じ高さを入れてから押す（本物の resize() が入れる物と同じ変数）。 */
  win.wrapH = 600;
  win.switchSheet(0);
  win.sel(0, 0, 0, 0);
  key('PageDown');
  const down = win.selR1;
  const mieru = Math.floor((600 - win.HDR_H) / win.ROW_H);
  ok(down > 1, '1画面ぶん動いていない（' + down + '行目）');
  ok(Math.abs(down - mieru) <= 2, '動いた行数が 見えている行数と合わない（動いた ' + down + ' ／ 見えている ' + mieru + '）');
  key('PageUp');
  eq(win.selR1, 0, '上へ戻る');
});

/* ── ⑦ F4（$の切替）── 4つの形は実測・順番は未測定 ── */
T('F4 … 数式バーの参照の $ が 4回で一周する（$A$1 → A$1 → $A1 → A1）', () => {
  const fi = doc.getElementById('formula-input');
  ok(fi, 'formula-input が無い');
  fi.value = '=A1+B2';
  fi.focus();
  try { fi.setSelectionRange(2, 2); } catch (e) { /* jsdom差 */ }
  win._fiCursorPos = 2;
  const seen = [];
  for (let i = 0; i < 4; i++) { key('F4', { target: fi }); seen.push(fi.value); }
  eq(seen.join(' → '), '=$A$1+B2 → =A$1+B2 → =$A1+B2 → =A1+B2');
  eq(seen[3], GOLD.f4_forms[3].replace('+$B$2', '+B2').replace('+B$2', '+B2').replace('+$B2', '+B2'), '一周して元に戻る');
});
T('F4 … 参照の上にいない時は 何も変えない（壊さない）', () => {
  const fi = doc.getElementById('formula-input');
  fi.value = '=SUM(1,2)';
  fi.focus();
  try { fi.setSelectionRange(6, 6); } catch (e) { /* jsdom差 */ }
  win._fiCursorPos = 6;
  key('F4', { target: fi });
  eq(fi.value, '=SUM(1,2)');
});

/* ── ⑧ Ctrl+S（書き出す）── */
T('Ctrl+S … 書き出しが呼ばれる（ブラウザの保存は止める）', () => {
  reset();
  const real = win.saveXlsx;
  let called = 0;
  win.saveXlsx = function () { called++; };
  const ev = key('s', { ctrl: true });
  win.saveXlsx = real;
  ok(ev.defaultPrevented, 'ブラウザの「ページを保存」を止めていない');
  eq(called, 1);
});

/* ── ⑨ ★測っていない物は作っていない★（作った事にしない） ── */
/* ── ⑨ Alt+=（オートSUM）★実Excelを押して測った3通りと一致するか★ ── */
T('★Alt+= の真値が golden に入っている（人が実物を押して測った）★', () => {
  eq(GOLD.autosum.measured, true, 'まだ測っていない事になっている');
  eq(GOLD.autosum.cases.length, 3, '測った通りの数が違う');
  ok(String(GOLD.autosum.how).length > 20, 'どう測ったかが書かれていない');
  ok(Array.isArray(GOLD.autosum.not_measured) && GOLD.autosum.not_measured.length > 0,
    '★測っていない所★ が書かれていない（全部 測った事にしない）');
});
for (const cs of GOLD.autosum.cases) {
  T('Alt+= … ' + cs.label + '（' + cs.at + ' で押す → ' + cs.formula + '）＝実Excelで実測', () => {
    reset();
    /* golden の setup を そのまま置く（数の中身は式に出ないので 1 でよい） */
    const put = (a1) => {
      const m = /^([A-Za-z]+)(\d+)$/.exec(a1);
      let c = 0; const L = m[1].toUpperCase();
      for (let i = 0; i < L.length; i++) c = c * 26 + (L.charCodeAt(i) - 64);
      return { r: +m[2] - 1, c: c - 1 };
    };
    /* ★置く場所は golden の fill を そのまま使う★
       日本語の setup から読み取ろうとしたら ★「C4 は空」まで置いてしまった★（2026-08-21 実際に踏んだ） */
    ok(Array.isArray(cs.fill) && cs.fill.length, 'golden に fill が無い＝どこに数を置くか決まっていない');
    for (const 範囲 of cs.fill) {
      const [a, b] = 範囲.split(':');
      const p1 = put(a), p2 = b ? put(b) : p1;
      for (let r = p1.r; r <= p2.r; r++) for (let c = p1.c; c <= p2.c; c++) win.setCell(r, c, '5');
    }
    const at = put(cs.at);
    win.sel(at.r, at.c, at.r, at.c);
    const ev = key('=', { alt: true });
    ok(ev.defaultPrevented, '既定を止めていない');
    const cell = A(at.r, at.c);
    ok(cell, cs.at + ' に何も入っていない');
    eq(cell.f, cs.formula, cs.at);
  });
}
T('★足す数が無い所で押しても 何も入れない（勝手に広げない＝分かりません）★', () => {
  reset();
  win.sel(4, 4, 4, 4);
  key('=', { alt: true });
  eq(A(4, 4), null, '足す数が無いのに何か入れた');
});
T('★F4 の「回る順番」は機械で測っていない と golden に書いてある★', () => {
  eq(GOLD.f4_order_measured, false);
  ok(String(GOLD.f4_order_note).length > 20, '理由が書かれていない');
});

console.log('\n  ' + pass + ' passed, ' + fail + ' failed');

/* ── 見張りが本当に見張っているか（わざと壊して 何通りで赤くなるか） ── */
if (SELF) {
  const { spawnSync } = await import('node:child_process');

  console.log('\n[self-test] わざと壊して 赤くなるかを数える');
  const BREAKS = [
    ['Ctrl+D の割り当てを消す', (s) => s.replace("if(ek==='d'){ e.preventDefault(); fillFromEdge('down'); return; }", '')],
    ['Ctrl+R の割り当てを消す', (s) => s.replace("if(ek==='r'){ e.preventDefault(); fillFromEdge('right'); return; }", '')],
    ['Ctrl+; の割り当てを消す', (s) => s.replace("if(e.key===';'){ e.preventDefault(); putNowValue('date'); return; }", '')],
    ['Ctrl+: の割り当てを消す', (s) => s.replace("if(e.key===':'){ e.preventDefault(); putNowValue('time'); return; }", '')],
    ['Ctrl+S の割り当てを消す', (s) => s.replace("if(ek==='s'){ e.preventDefault(); saveXlsx(); return; }", '')],
    ['Ctrl+B の割り当てを消す', (s) => s.replace("if(ek==='b'){ e.preventDefault(); toggleFormat('bold'); return; }", '')],
    ['Ctrl+1 の割り当てを消す', (s) => s.replace("if(e.key==='1'){ e.preventDefault(); openFmtModal(); return; }", '')],
    ['Ctrl+Space の割り当てを消す', (s) => s.replace("if(e.key===' '){ e.preventDefault(); sel(0,selC1,ROWS-1,selC2); render(); return; }", '')],
    ['Shift+Space の割り当てを消す', (s) => s.replace("if(e.shiftKey && e.key===' '){ e.preventDefault(); sel(selR1,0,selR2,COLS-1); render(); moved=true; }", '')],
    ['シート移動の割り当てを消す', (s) => s.replace("if(e.key==='PageDown'){ e.preventDefault(); switchSheet(Math.min(sheets.length-1, activeSheet+1)); return; }", '')],
    ['1画面 上下 を1行だけにする', (s) => s.replace('var n = Math.max(1, bot - top);', 'var n = 1;')],
    ['Ctrl+D で書式を写さない', (s) => s.replace('for(var i=0;i<FMT.length;i++){ if(src[FMT[i]]!==undefined) keep[FMT[i]] = src[FMT[i]]; }', '')],
    ['Ctrl+D で先の書式を残す（実Excelと違う）', (s) => s.replace('var keep = {};', 'var keep = Object.assign({}, prev);')],
    ['Ctrl+R で参照をずらさない', (s) => s.replace('if(typeof src.f===\'string\' && src.f.charAt(0)===\'=\'){ raw = shiftFormula(src.f, r-sr, c-sc); }', 'if(typeof src.f===\'string\' && src.f.charAt(0)===\'=\'){ raw = src.f; }')],
    ['日付の書式を当てない', (s) => s.replace("fmt='yyyy/m/d'; }", "fmt=null; }")],
    ['時刻の書式を日付にする', (s) => s.replace("fmt='h:mm';", "fmt='yyyy/m/d';")],
    ['F4 の $ 切替を止める', (s) => s.replace('var res = GridRefEdit.toggleAbsAt(el.value, pos);', 'var res = {ok:false};')],
    ['Alt+= の割り当てを消す', (s) => s.replace("  if(e.altKey && e.key==='='){ e.preventDefault(); autoSum(); moved=true; }   /* オートSUM */\n", '')],
    ['Alt+= で 足す数が無くても勝手に入れる', (s) => s.replace(
      "if(!got){ showToast('足す数が見つかりません（すぐ上か すぐ左に数を並べてね）'); return; }",
      "if(!got){ setCell(selR1,selC1,'=SUM(A1:A2)'); return; }")],
    ['既定の動きを止めない（Ctrl+D の preventDefault を消す）', (s) => s.replace("if(ek==='d'){ e.preventDefault(); fillFromEdge('down'); return; }", "if(ek==='d'){ fillFromEdge('down'); return; }")],
  ];

  const LIB_BREAKS = [
    ['F4 の回る順番を変える（lib）', (s) => s.replace('var next = { 0: 3, 3: 1, 1: 2, 2: 0 }[state];', 'var next = { 0: 1, 1: 3, 3: 2, 2: 0 }[state];')],
  ];
  const SUM_BREAKS = [
    ['Alt+= が ★空きで止まらない★（離れた数まで拾う）', (s) => s.replace(
      'while (top > 0 && isNum(get(top - 1, c))) top--;',
      'while (top > 0) top--;')],
    ['Alt+= が 上と左の順番を逆にする', (s) => s.replace(
      'if (上あり) {', 'if (false) {')],
    ['Alt+= が 見つからない時に 勝手に範囲を作る（分かりません を出さない）', (s) => s.replace(
      '    /* ③ ★分かりません★（勝手に広げない） */\n    return null;',
      "    return { formula: '=SUM(A1:A2)', dir: 'up', from: 'A1', to: 'A2' };")],
    ['Alt+= が 1個だけの時に A1:A1 のような書き方をする', (s) => s.replace(
      "'=SUM(' + (a === b ? a : a + ':' + b) + ')', dir: 'up'",
      "'=SUM(' + a + ':' + b + ')', dir: 'up'")],
  ];
  let red = 0, total = BREAKS.length + LIB_BREAKS.length + SUM_BREAKS.length;
  const os = await import('node:os');
  const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'exally-shortcuts-'));
  /* ★repo は読むだけ。壊した中身は temp に置いて 子に env で渡す★ */
  const tryBreak = (rel, name, brk) => {
    const orig = fs.readFileSync(path.join(ROOT, rel), 'utf8');
    const bad = brk(orig);
    if (bad === orig) { console.log('  ★置換できず★  ' + name + '（壊し方が当たっていない）'); return; }
    const tmpFile = path.join(TMP, rel.replace(/[\\/]/g, '_'));
    fs.writeFileSync(tmpFile, bad, 'utf8');
    const env = Object.assign({}, process.env, { EXALLY_SHORTCUT_OVERRIDE: JSON.stringify({ [rel]: tmpFile }) });
    const isRed = spawnSync(process.execPath, [path.join(__dirname, 'excel-shortcuts.test.mjs')], { encoding: 'utf8', env }).status !== 0;
    if (isRed) { red++; console.log('  赤くなった  ' + name); }
    else console.log('  ★素通り★  ' + name);
  };
  for (const [name, brk] of BREAKS) tryBreak('book.html', name, brk);
  for (const [name, brk] of LIB_BREAKS) tryBreak('lib/grid-refedit.js', name, brk);
  for (const [name, brk] of SUM_BREAKS) tryBreak('lib/autosum.js', name, brk);
  /* ★repo を1バイトも触っていない事を その場で確かめる（前は壊した行が残った）★ */
  for (const rel of ['book.html', 'lib/grid-refedit.js', 'lib/autosum.js']) {
    const now = fs.readFileSync(path.join(ROOT, rel), 'utf8');
    if (now.includes('var res = {ok:false};') || now.includes("setCell(selR1,selC1,'=SUM(A1:A2)')")
      || now.includes('while (top > 0) top--;') || now.includes('if (false) {')) {
      console.log('  ★NG★ ' + rel + ' に わざと壊した物が残っている');
      process.exit(1);
    }
  }
  try { fs.rmSync(TMP, { recursive: true, force: true }); } catch (e) { /* 消せなくても検査は済んでいる */ }
  console.log('\n  ' + red + '/' + total + ' 通りで赤くなった');
  process.exit(red === total ? 0 : 1);
}

process.exit(fail ? 1 : 0);
