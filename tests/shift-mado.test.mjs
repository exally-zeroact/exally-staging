/* shift-mado.test.mjs — ★挿入／削除で「どちらへ 詰めるか」を 聞く 窓★ 2026-09-03
 *
 *  ★司さんの 決め（2026-09-03）＝「Excelと 同じ」★
 *    ★実Excel は 右クリックの「挿入(I)...」「削除(D)...」に ★...★ が 付いている＝★聞く★★
 *    （2026-09-03 に 実Excel を ★本物の マウス★で 開いて 目で 確かめた
 *      ⇒ docs/EXCEL_CELL_CTXMENU_2026-09-03.md ／ shot/excel_cell_ctxmenu_jitsubutsu.png）
 *
 *  ★見る 物★
 *    ①★4つ 出る★（右/下 or 左/上・行全体・列全体）
 *    ②★はじめから 選ばれているのは 前と 同じ★（挿入＝下へ／削除＝上へ）
 *       ＝★OK を 1回 押せば 前と 同じ★（★1押しで 出来ていた物を 2段に 落としすぎない★）
 *    ③★4つとも 前から 在る 手を 呼ぶ★＝★新しい 働きを 作っていない★
 *    ④★本当に 動く★＝選んだ 向きに 中身が 詰まる（jsdom で 実際に 押す）
 *    ⑤★alert / prompt / confirm を 使っていない★
 *
 *  ★窓の 中の 字は 実Excel を 撮って 合わせる★＝★まだ 撮れていない★（下の「未測定」）
 *
 *  走らせ方: node tests/shift-mado.test.mjs [--self-test]
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(__dirname, '..');
const OVERRIDE = process.env.EXALLY_SHIFT_OVERRIDE ? JSON.parse(process.env.EXALLY_SHIFT_OVERRIDE) : {};
const srcPath = (rel) => OVERRIDE[rel] || path.join(ROOT, rel);

let JSDOM;
try { ({ JSDOM } = await import('jsdom')); }
catch { console.log('★jsdomが入っていません。この検証は飛ばせません（SKIPを緑と呼ばない）'); process.exit(1); }

let pass = 0, fail = 0;
const T = (n, fn) => {
  try { fn(); pass++; console.log('  ok   ' + n); }
  catch (e) { fail++; console.log('  NG   ' + n + '\n       ' + (e && e.message)); }
};
const 言う = (よい, 文) => { if (!よい) throw new Error(文); };

const CANVAS_STUB = [
  '(function(){var noop=function(){return noop;};',
  'var ctx=new Proxy({measureText:function(t){return {width:(t||"").length*7};}},',
  '{get:function(t,k){ if(k in t) return t[k];',
  ' return noop;},set:function(t,k,v){ t[k]=v; return true; }});',
  'HTMLCanvasElement.prototype.getContext=function(){return ctx;};})();',
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

console.log('');
console.log('[shift-mado] ★挿入／削除で どちらへ 詰めるかを 聞く★');

const sh = () => win.sheets[win.activeSheet];
const 中身 = (r, c) => {
  const v = sh().data[r + ',' + c];
  return v ? String(v.f !== undefined && v.f !== '' ? v.f : (v.v || '')) : '';
};
const 下地 = () => {
  sh().data = {};
  /* A1..A4 に あ い う え ／ B1..B4 に 1 2 3 4 */
  const a = ['あ', 'い', 'う', 'え'];
  for (let r = 0; r < 4; r++) { win.setCell(r, 0, a[r]); win.setCell(r, 1, String(r + 1)); }
};
const 選ぶ = (値) => {
  const el = doc.querySelector('input[name="shiftPick"][value="' + 値 + '"]');
  言う(!!el, '「' + 値 + '」の 選び が 無い');
  doc.querySelectorAll('input[name="shiftPick"]').forEach((e) => { e.checked = false; });
  el.checked = true;
};

/* ── ①4つ 出る ── */
for (const どっち of ['挿入', '削除']) {
  T('★' + どっち + ' … 4つ 出る★', () => {
    下地(); win.sel(1, 0, 1, 0);
    const n = win.詰める向きの窓を開く(どっち);
    言う(n === 4, '選びが ' + n + '個（4では ない）');
    言う(doc.getElementById('shiftOverlay').style.display === 'flex', '窓が 出ていない');
    言う(doc.getElementById('shiftTitle').textContent === どっち, '題が 違う');
    win.詰める向きの窓を閉じる();
  });
}

/* ── ②はじめから 選ばれている 物＝★実Excel と 同じ★（2026-09-04 に 実物を 撮って 合わせた） ──
   ★実Excel★ … 挿入＝★右方向にシフト(I)★／削除＝★左方向にシフト(L)★
     絵 … shot/excel_dialog_sonyu.png ／ shot/excel_dialog_sakujo.png
   ★前は 挿入＝下／削除＝上★（＝直す前の うちの 動きに 合わせていた）
   ⇒★司さんの 決めは「Excelと 同じ」★＝実物に 合わせた */
T('★はじめは 挿入＝右方向・削除＝左方向（実Excel と 同じ）★', () => {
  下地(); win.sel(1, 0, 1, 0);
  win.挿入の窓を開く();
  言う(doc.querySelector('input[name="shiftPick"]:checked').value === 'right',
    '挿入の はじめが 右方向に なっていない');
  win.詰める向きの窓を閉じる();
  win.削除の窓を開く();
  言う(doc.querySelector('input[name="shiftPick"]:checked').value === 'left',
    '削除の はじめが 左方向に なっていない');
  win.詰める向きの窓を閉じる();
});

/* ── ★窓の 形も 実物と 同じか★（2026-09-04 に 撮って 合わせた） ── */
T('★OK が 左・キャンセルが 右／字は「キャンセル」（実Excel と 同じ）★', () => {
  下地(); win.sel(1, 0, 1, 0);
  win.削除の窓を開く();
  const 箱 = doc.getElementById('shiftOk').parentNode;
  const ボタン = [...箱.querySelectorAll('button')].map((b) => b.textContent.trim());
  言う(ボタン[0] === 'OK', '1つ目が OK では ない（今 ' + ボタン.join('・') + '）');
  言う(ボタン[1] === 'キャンセル', '2つ目が キャンセルでは ない（今 ' + ボタン.join('・') + '）');
  win.詰める向きの窓を閉じる();
});
T('★題の 下に もう一度 「挿入」「削除」と 出る（実Excel と 同じ）★', () => {
  下地(); win.sel(1, 0, 1, 0);
  win.挿入の窓を開く();
  言う(doc.getElementById('shiftLead').textContent === '挿入', '小見出しが 挿入に なっていない');
  win.詰める向きの窓を閉じる();
  win.削除の窓を開く();
  言う(doc.getElementById('shiftLead').textContent === '削除', '小見出しが 削除に なっていない');
  win.詰める向きの窓を閉じる();
});

/* ── ③4つとも 前から 在る 手を 呼ぶ（新しい 働きを 作っていない） ── */
T('★呼ぶ先は 全部 前から 在る 手★', () => {
  const 素 = fs.readFileSync(srcPath('book.html'), 'utf8');
  const 在る = {};
  for (const m of 素.match(/function\s+([^\s(){}]+)\s*\(/g) || []) {
    在る[m.replace(/^function\s+/, '').replace(/\s*\($/, '')] = 1;
  }
  for (const n of ['セルを挿入', 'セルを削除', 'ctxInsertRow', 'ctxInsertCol',
    'ctxDeleteRow', 'ctxDeleteCol']) 言う(!!在る[n], n + ' が 無い');
});

/* ── ④本当に 動く（4通り × 2） ── */
T('★削除・上方向 … A列だけ 上へ 詰まる（B列は 動かない）★', () => {
  下地(); win.sel(0, 0, 0, 0);
  win.削除の窓を開く(); 選ぶ('up'); win.詰める向きを決める();
  言う(中身(0, 0) === 'い' && 中身(1, 0) === 'う', 'A列が 上へ 詰まっていない（' + 中身(0, 0) + '）');
  言う(中身(0, 1) === '1' && 中身(1, 1) === '2', '★B列が 動いた★（' + 中身(0, 1) + '）');
});
T('★削除・左方向 … その行だけ 左へ 詰まる★', () => {
  下地(); win.sel(0, 0, 0, 0);
  win.削除の窓を開く(); 選ぶ('left'); win.詰める向きを決める();
  言う(中身(0, 0) === '1', '左へ 詰まっていない（' + 中身(0, 0) + '）');
  言う(中身(1, 0) === 'い', '★別の 行が 動いた★（' + 中身(1, 0) + '）');
});
T('★削除・行全体 … 1行 まるごと 消える★', () => {
  下地(); win.sel(0, 0, 0, 0);
  win.削除の窓を開く(); 選ぶ('row'); win.詰める向きを決める();
  言う(中身(0, 0) === 'い' && 中身(0, 1) === '2', '行が 消えていない（' + 中身(0, 0) + '/' + 中身(0, 1) + '）');
});
T('★削除・列全体 … 1列 まるごと 消える★', () => {
  下地(); win.sel(0, 0, 0, 0);
  win.削除の窓を開く(); 選ぶ('col'); win.詰める向きを決める();
  言う(中身(0, 0) === '1' && 中身(3, 0) === '4', '列が 消えていない（' + 中身(0, 0) + '）');
});
T('★挿入・下方向 … A列だけ 下へ ずれる（B列は 動かない）★', () => {
  下地(); win.sel(0, 0, 0, 0);
  win.挿入の窓を開く(); 選ぶ('down'); win.詰める向きを決める();
  言う(中身(0, 0) === '' && 中身(1, 0) === 'あ', 'A列が 下へ ずれていない');
  言う(中身(0, 1) === '1', '★B列が 動いた★（' + 中身(0, 1) + '）');
});
T('★挿入・行全体 … 1行 まるごと 増える★', () => {
  下地(); win.sel(0, 0, 0, 0);
  win.挿入の窓を開く(); 選ぶ('row'); win.詰める向きを決める();
  言う(中身(0, 0) === '' && 中身(1, 0) === 'あ' && 中身(1, 1) === '1', '行が 増えていない');
});

/* ── ⑤止まる 物を 使っていない ── */
T('★alert / prompt / confirm を 使っていない★', () => {
  const 素 = fs.readFileSync(srcPath('book.html'), 'utf8');
  const i = 素.indexOf('function 詰める向きの窓を開く');
  const j = 素.indexOf('function drawText', i);
  const 中 = 素.slice(i, j > i ? j : i + 4000);
  言う(!/\balert\(|\bprompt\(|\bconfirm\(/.test(中), '画面が 止まる 物を 使っている');
});

/* ── ★実Excel の 窓と 突き合わせた（2026-09-04）★ ──
   ★撮り方★ tools/excel-shift-dialog-shot.ps1（★本物の マウスで 右クリック → アクセスキー★）
     ★踏んだ 罠★＝★アクセスキーが 1つだけの 行は 押した 瞬間に 実行される★
       削除(D) … メニューに D は 1つ ⇒★押した 瞬間に 窓が 出る（Enter を 押しては いけない）★
       挿入(I) … リンク(I) と かぶる ⇒★選ばれるだけ＝Enter が 要る★
       （★Enter を 足したら 窓が 出ず 2回 撮り直した★）
   ★実物（絵 … shot/excel_dialog_sonyu.png ／ excel_dialog_sakujo.png）★
     挿入 … ★右方向にシフト(I)★【はじめ】／下方向にシフト(D)／行全体(R)／列全体(C)
     削除 … ★左方向にシフト(L)★【はじめ】／上方向にシフト(U)／行全体(R)／列全体(C)
     どちらも ★題の 下に もう一度 「挿入」「削除」★／★OK が 左・キャンセルが 右★
   ★合っていた★ … 4つの 言い方（右/下・左/上・行全体・列全体）
   ★違っていた★ … ★はじめの 選び★（うちは 下/上 → ★右/左 に 直した★）
                  ★ボタンの 並びと 字★（うちは やめる/OK → ★OK/キャンセル に 直した★）
   ⇒★「窓の 中の 字は 未測定」は 消えました★ */

/* ── わざと 壊して 赤に なるか ── */
if (process.argv.includes('--self-test')) {
  console.log('');
  console.log('★わざと 壊して 赤に なるか★');
  const 元 = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf8');
  const 壊す = [
    ['★選びを 3つに 減らす★',
      (s) => s.replace("    { 値: 'col',   字: '列全体',          呼ぶ: function(){ return ctxDeleteCol(); } }", '')],
    ['★はじめの 選びを 変える（挿入＝行全体に する）★',
      (s) => s.replace("var SHIFT_HAJIME = { 挿入: 'right', 削除: 'left' };",
        "var SHIFT_HAJIME = { 挿入: 'row', 削除: 'left' };")],
    /* ★並びを 本当に 入れ替える★（id を 入れ替えるだけでは 並びは 変わらない
       ＝2026-09-04 実測で 素通りした） */
    ['★OK と キャンセルの 並びを 入れ替える（実Excel と 違う 形に する）★',
      (s) => {
        const 行 = s.split(String.fromCharCode(10));
        const a = 行.findIndex((l) => l.indexOf('id="shiftOk"') >= 0);
        const b = 行.findIndex((l) => l.indexOf('id="shiftCancel"') >= 0);
        if (a < 0 || b < 0) return s;
        const t = 行[a]; 行[a] = 行[b]; 行[b] = t;
        return 行.join(String.fromCharCode(10));
      }],
    ['★選んだ 向きを 無視して いつも 上へ 詰める★',
      (s) => s.replace("    var r = 一覧[i].呼ぶ();", "    var r = セルを削除('up');")],
    ['★prompt で 聞く（画面が 止まる）★',
      (s) => s.replace('  var ov = document.getElementById(\'shiftOverlay\');\n  if (!ov) return 0;',
        '  var ov = document.getElementById(\'shiftOverlay\');\n  if (!ov) return 0;\n  prompt(\'どちら？\');')],
  ];
  const tmp = path.join(ROOT, 'tests', '_shift_broken.html');
  const { execFileSync } = await import('node:child_process');
  for (const [名, f] of 壊す) {
    const 壊れ = f(元);
    if (壊れ === 元) { console.log('  ★素通り★  ' + 名 + '（印が 古い＝直せ）'); fail++; continue; }
    fs.writeFileSync(tmp, 壊れ);
    let 赤 = false;
    try {
      execFileSync(process.execPath, [path.join(ROOT, 'tests', 'shift-mado.test.mjs')], {
        env: { ...process.env, EXALLY_SHIFT_OVERRIDE: JSON.stringify({ 'book.html': tmp }) },
        stdio: 'pipe',
      });
    } catch (e) { 赤 = true; }
    console.log((赤 ? '  赤くなった  ' : '  ★素通り★  ') + 名);
    if (!赤) fail++;
  }
  try { fs.unlinkSync(tmp); } catch (e) { /* 無くてよい */ }
}

console.log('');
console.log('shift-mado: ' + pass + ' 緑 / ' + fail + ' 赤');
process.exit(fail ? 1 : 0);
