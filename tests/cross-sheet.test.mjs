/* cross-sheet.test.mjs — ★他のシートを参照している合計が、黙って小さくなっていないか★
 *
 * なぜ必要か（2026-08-11・指示役が本番の実機で発見）:
 *   見本ブック（4月/5月/6月/まとめ）を開いて 4月 C4 を 1→99 にすると、
 *   まとめの合計は 527,000 にならなければいけない。実際は ★186,000★ になった。
 *     まとめ B4 ='4月'!E14  → 186,000  ○
 *     まとめ B5 ='5月'!E14  → ★空★    ×
 *     まとめ B6 ='6月'!E14  → ★空★    ×
 *   ★#ERROR は出ない。空になって合計が黙って 34万円 小さくなる。誰も気づけない。★
 *
 *   原因: 見ているシートだけを計算する側へ流していたので、まとめを計算する時点で
 *   5月・6月 が計算する側に入っていない（空のシートを参照して空が返る）。
 *   しかも一度流したシートは二度と計算し直さないので、
 *   ★5月・6月 を開いてから まとめへ戻っても 186,000 のまま★になる。
 *
 * ★この検査は「合計」だけを見ない★
 *   合計だけ見ると、別の理由（たまたま同じ数）で通ってしまう。B4/B5/B6 の1本ずつを見る。
 *
 * 材料: tests/fixtures/cross-sheet-sample.xlsb（★実Excelで作り、真値もExcelに出させた★）
 *   開いた時      88,000 + 143,000 + 198,000 = 429,000
 *   4月C4を99に → 186,000 + 143,000 + 198,000 = ★527,000★
 *
 * 使い方: node tests/cross-sheet.test.mjs
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createRequire } from 'node:module';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require = createRequire(pathToFileURL(path.join(ROOT, 'package.json')));
const FIX = path.join(ROOT, 'tests/fixtures/cross-sheet-sample.xlsb');

let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ✓ ' + n); } catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + (e && e.message)); } };
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

let JSDOM;
try { ({ JSDOM } = require('jsdom')); }
catch {
  console.log('\n[cross-sheet] ★jsdom が入っていません（npm install）。★緑ではありません★');
  process.exit(1);
}

/* ── book.html を丸ごと載せる（＝本番と同じ関数を呼ぶ） ──
   ★load が終わるのを必ず待つ★。book.html は load の中で canvas を掴み、
   ★initFormulaEngine をもう一度呼ぶ★。待たずにブックを開くと、
   その後で load が走って ★計算する側が空に作り直され、全部の式が空になる★
   （2026-08-11 実測。本番では load は人がファイルを開く前に終わっている）。 */
async function boot() {
  const html = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf8');
  const dom = new JSDOM(html.replace(/<script[\s\S]*?<\/script>/g, ''), {
    runScripts: 'dangerously', pretendToBeVisual: true, url: 'http://localhost/'
  });
  const win = dom.window, doc = win.document;
  win.fetch = () => Promise.reject(new Error('no net'));
  win.scrollTo = () => {}; win.alert = () => {};
  const ctx = new Proxy({}, {
    get: (t, k) => (k === 'measureText' ? (() => ({ width: 10 }))
      : k === 'canvas' ? { width: 800, height: 600 }
        : k === 'getImageData' ? (() => ({ data: [] })) : (() => {}))
  });
  win.HTMLCanvasElement.prototype.getContext = () => ctx;
  const errs = [];
  win.addEventListener('error', (e) => errs.push(String(e.message || e)));
  const inject = (code) => { const el = doc.createElement('script'); el.textContent = code; doc.body.appendChild(el); };
  for (const m of html.matchAll(/<script src="([^"]+)"><\/script>/g)) {
    const src = m[1].split('?')[0];
    if (/^https?:/.test(src)) continue;
    inject(fs.readFileSync(path.join(ROOT, src), 'utf8'));
  }
  for (const m of html.matchAll(/<script(?![^>]*\ssrc=)[^>]*>([\s\S]*?)<\/script>/g)) inject(m[1]);
  /* 本番が「開く／保存」の時に遅れて読む物（_ensureXlsx と同じ並び・同じファイル） */
  for (const p of ['lib/xlsx.full.min.js', 'lib/zip-surgeon.js', 'lib/xlsx-edit.js', 'lib/xlsb-edit.js', 'js/book-open.js']) {
    inject(fs.readFileSync(path.join(ROOT, p), 'utf8'));
  }
  if (doc.readyState !== 'complete') await new Promise((r) => win.addEventListener('load', r, { once: true }));
  else win.dispatchEvent(new win.Event('load'));
  return { win, errs };
}

/* ★本番の openBookFile と同じ順番でブックを開く★ */
async function openBook(win, bytes) {
  const res = await win.BookOpen.openFile({ name: 'cross-sheet-sample.xlsb', arrayBuffer: () => Promise.resolve(bytes.buffer.slice(0)) });
  win.sheets = res.sheets;
  win.activeSheet = 0;
  win.initFormulaEngine(res.sheets.map((s) => s.name));
  win.loadSheetIntoEngine(0);
  return res;
}
const cellD = (win, sheetIdx, r, c) => {
  const cell = win.sheets[sheetIdx].data[r + ',' + c];
  return cell === undefined ? '(セルが無い)' : String(cell.d);
};
const num = (s) => Number(String(s).replace(/[,\s]/g, ''));

console.log('\n[cross-sheet] 他のシートを参照している合計が 黙って小さくなっていないか');

const bytes = new Uint8Array(fs.readFileSync(FIX));
const SUM_ROW = 6, B = 1;               // まとめ B7（0起点で r=6, c=1）
const MATOME = 3, APRIL = 0;

/* ═══ ① 材料が本当に「他シート参照」を持っているか（★空振りしていない事の確認★） ═══ */
{
  const { win } = await boot();
  const res = await openBook(win, bytes);
  T('材料: シートが 4月/5月/6月/まとめ の4枚ある', () => {
    const got = res.sheets.map((s) => s.name).join(',');
    if (got !== '4月,5月,6月,まとめ') throw new Error(got);
  });
  T("材料: まとめ B5 が ='5月'!E14（他シート参照）を持っている", () => {
    const f = win.sheets[MATOME].data['4,1'].f;
    if (!/5月/.test(String(f))) throw new Error('式が違う: ' + f);
  });
  T('材料: Excelが書き残した合計は 429,000（開いた時）', () => {
    const wanted = 429000;
    const XLSX = require(path.join(ROOT, 'lib/xlsx.full.min.js'));
    const wb = XLSX.read(bytes, { type: 'array', cellFormula: true });
    const v = wb.Sheets['まとめ']['B7'].v;
    if (v !== wanted) throw new Error('Excelの控え=' + v);
  });
  win.close();
}

/* ═══ ② ★本番の手順そのまま★ 開く → 4月C4を99 → まとめを見る ═══ */
{
  const { win, errs } = await boot();
  await openBook(win, bytes);
  win.setCell(3, 2, '99');               // 4月 C4 = 99（commitEdit が呼ぶ本物の入口）
  await sleep(300);                      // _scheduleRecalc(150ms) が終わるのを待つ
  const e4 = cellD(win, APRIL, 3, 4), e14 = cellD(win, APRIL, 13, 4);
  win.switchSheet(MATOME);               // ★まとめのタブを押す★
  await sleep(50);
  const b4 = cellD(win, MATOME, 3, B), b5 = cellD(win, MATOME, 4, B), b6 = cellD(win, MATOME, 5, B);
  const sum = cellD(win, MATOME, SUM_ROW, B);

  console.log(`\n   4月 E4=${e4} / E14=${e14}`);
  console.log(`   まとめ B4=${JSON.stringify(b4)} B5=${JSON.stringify(b5)} B6=${JSON.stringify(b6)} 合計=${JSON.stringify(sum)}`);

  T('4月 E4 = 99,000', () => { if (num(e4) !== 99000) throw new Error(e4); });
  T('4月 E14 = 186,000', () => { if (num(e14) !== 186000) throw new Error(e14); });
  T("★まとめ B5 が空でない（='5月'!E14 = 143,000）", () => {
    if (b5 === '' || b5 === 'undefined') throw new Error('★空になっている★＝合計が黙って小さくなる');
    if (num(b5) !== 143000) throw new Error(b5);
  });
  T("★まとめ B6 が空でない（='6月'!E14 = 198,000）", () => {
    if (b6 === '' || b6 === 'undefined') throw new Error('★空になっている★＝合計が黙って小さくなる');
    if (num(b6) !== 198000) throw new Error(b6);
  });
  T('まとめ B4 = 186,000', () => { if (num(b4) !== 186000) throw new Error(b4); });
  T('★★まとめ 合計 = 527,000（合格の値）★★', () => {
    if (num(sum) !== 527000) throw new Error(sum + ' ← 34万円 少ない。誰も気づけない形の間違い');
  });
  T('画面のJS例外が0', () => { if (errs.length) throw new Error(errs.slice(0, 3).join(' / ')); });
  win.close();
}

/* ═══ ③ ★先にまとめを見てから 4月を直しても、まとめが古いままにならない★ ═══ */
{
  const { win } = await boot();
  await openBook(win, bytes);
  win.switchSheet(MATOME);               // 先に まとめ を見る（429,000 のはず）
  await sleep(50);
  const before = cellD(win, MATOME, SUM_ROW, B);
  win.switchSheet(APRIL);
  await sleep(50);
  win.setCell(3, 2, '99');               // 4月 C4 = 99
  await sleep(300);
  win.switchSheet(MATOME);               // まとめへ戻る
  await sleep(50);
  const after = cellD(win, MATOME, SUM_ROW, B);
  console.log(`\n   （先にまとめを見た場合）直す前=${JSON.stringify(before)} → 直した後=${JSON.stringify(after)}`);
  T('先にまとめを見た時の合計 = 429,000', () => { if (num(before) !== 429000) throw new Error(before); });
  T('★4月を直して戻ったら 527,000（古いまま残らない）', () => {
    if (num(after) !== 527000) throw new Error(after + ' ← 開いた時の数字が残っている');
  });
  win.close();
}

/* ═══ ④ ★何も打たずにタブを見て回っただけなら「変わっていない」★ ═══
   シートを開くたびに計算し直すようにしたので、うっかり
   ★触っていないのに「変わった」ことになって、保存でファイルが書き換わる★のを止める。 */
{
  const { win } = await boot();
  await openBook(win, bytes);
  for (const i of [1, 2, 3, 0, 3]) { win.switchSheet(i); await sleep(30); }
  T('★全シートを見て回っても 1セルも「変わった」ことにならない', () => {
    const changed = win.BookOpen.anyChanged(win.sheets);
    if (changed) {
      const detail = win.sheets.map((sh) => sh.name + ':' + Object.keys(win.BookOpen.changedCells(sh)).length).join(' ');
      throw new Error('変わった扱いのセルがある → ' + detail);
    }
  });
  const saved = await win.BookOpen.saveOpened(win.sheets);
  T('★保存したバイト数が元と同じ（' + bytes.length + 'バイト）', () => {
    if (saved.bytes.length !== bytes.length) throw new Error(saved.bytes.length + ' ≠ ' + bytes.length);
    if (!saved.log || !saved.log.noChange) throw new Error('「変わっていない」道を通っていない');
  });
  win.close();
}

console.log('\n' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);

