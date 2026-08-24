/* excel-parity.test.mjs — ★docs/EXCEL_PARITY.md の数字を 機械で数え直す★
 *
 * なぜ必要か（2026-08-18 方針の上書き）
 *   司さん「Excelの最上級として作る／つけられる機能は全て作る」。
 *   ★作る前に端から数える★と決めたが、★数えた表は必ず古くなる★。
 *   1機能 作るたびに手で書き換える約束は、必ずどこかで守られなくなる。
 *   ⇒ ★数字は毎回 機械で数え直し、表とズレたら赤★にする。
 *
 * 数える物（★全部 実際に押す／DOMから拾う。推測しない★）
 *   ① ショートカット … 本物の book.html に keydown を1つずつ送って「効いたか」を見る
 *   ② 画面の操作     … button / select / input / onclick を全部 拾う
 *
 * 使い方: node tests/excel-parity.test.mjs [--self-test]
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(__dirname, '..');
const require = createRequire(import.meta.url);

let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ✓ ' + n); } catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + (e && e.message)); } };

/* ★Excel 365（日本語UI）の代表的なショートカット★
   ここを増やしたら 表(docs/EXCEL_PARITY.md)の「数えた○個」も一緒に直す。 */
export const SHORTCUTS = [
  { key: 'c', ctrl: true, name: 'Ctrl+C コピー' },
  { key: 'v', ctrl: true, name: 'Ctrl+V 貼り付け' },
  { key: 'x', ctrl: true, name: 'Ctrl+X 切り取り' },
  { key: 'z', ctrl: true, name: 'Ctrl+Z 元に戻す' },
  { key: 'y', ctrl: true, name: 'Ctrl+Y やり直す' },
  { key: 'a', ctrl: true, name: 'Ctrl+A 全選択' },
  { key: 's', ctrl: true, name: 'Ctrl+S 書き出す' },
  { key: 'f', ctrl: true, name: 'Ctrl+F 検索' },
  { key: 'h', ctrl: true, name: 'Ctrl+H 置換' },
  { key: 'b', ctrl: true, name: 'Ctrl+B 太字' },
  { key: 'i', ctrl: true, name: 'Ctrl+I 斜体' },
  { key: 'u', ctrl: true, name: 'Ctrl+U 下線' },
  { key: 'p', ctrl: true, name: 'Ctrl+P 印刷' },
  { key: 'd', ctrl: true, name: 'Ctrl+D 上のセルをコピー' },
  { key: 'r', ctrl: true, name: 'Ctrl+R 左のセルをコピー' },
  { key: '1', ctrl: true, name: 'Ctrl+1 セルの書式設定' },
  { key: ';', ctrl: true, name: 'Ctrl+; 今日の日付' },
  { key: ':', ctrl: true, shift: true, name: 'Ctrl+: 今の時刻' },
  { key: '=', alt: true, name: 'Alt+= オートSUM' },
  { key: 'ArrowDown', ctrl: true, name: 'Ctrl+↓ 端まで移動' },
  { key: 'ArrowUp', ctrl: true, name: 'Ctrl+↑ 端まで移動' },
  { key: 'ArrowRight', ctrl: true, name: 'Ctrl+→ 端まで移動' },
  { key: 'ArrowLeft', ctrl: true, name: 'Ctrl+← 端まで移動' },
  { key: 'ArrowDown', ctrl: true, shift: true, name: 'Ctrl+Shift+↓ 端まで選択' },
  { key: 'ArrowRight', ctrl: true, shift: true, name: 'Ctrl+Shift+→ 端まで選択' },
  { key: ' ', shift: true, name: 'Shift+Space 行を選ぶ' },
  { key: ' ', ctrl: true, name: 'Ctrl+Space 列を選ぶ' },
  { key: 'Home', ctrl: true, name: 'Ctrl+Home A1へ' },
  { key: 'End', ctrl: true, name: 'Ctrl+End 最後のセルへ' },
  { key: 'PageDown', ctrl: true, name: 'Ctrl+PageDown 次のシート' },
  { key: 'PageUp', ctrl: true, name: 'Ctrl+PageUp 前のシート' },
  { key: 'F4', name: 'F4 直前の操作を繰り返す／$の切替' },
  { key: 'F2', name: 'F2 セルを編集' },
  { key: 'Home', name: 'Home 行の先頭へ' },
  { key: 'End', name: 'End 行の末尾へ' },
  { key: 'PageDown', name: 'PageDown 1画面下' },
  { key: 'PageUp', name: 'PageUp 1画面上' },
  { key: 'Delete', name: 'Delete 中身を消す' },
  { key: 'Escape', name: 'Esc やめる' },
  { key: 'Enter', name: 'Enter 下へ' },
  { key: 'Tab', name: 'Tab 右へ' },
  { key: 'ArrowDown', name: '↓ 下へ' },
];

/** 表(docs/EXCEL_PARITY.md)に書いてある数字を読む */
export function readDoc(text) {
  const m1 = /★(\d+)個を数えて 在る (\d+) ／ 無い (\d+)★/.exec(text);
  const m2 = /★(\d+)個（同じ物を除くと (\d+)個）★/.exec(text);
  return {
    counted: m1 ? +m1[1] : null, have: m1 ? +m1[2] : null, missing: m1 ? +m1[3] : null,
    controls: m2 ? +m2[1] : null, uniqueControls: m2 ? +m2[2] : null,
  };
}

/* ══ self-test ══════════════════════════════════════════════════════ */
if (process.argv.includes('--self-test')) {
  console.log('\n[excel-parity --self-test] 判定そのものが空振りしていないか');
  T('★表の数字を読み取れる', () => {
    const d = readDoc('★42個を数えて 在る 21 ／ 無い 21★ と ★184個（同じ物を除くと 178個）★');
    if (d.counted !== 42 || d.have !== 21 || d.missing !== 21) throw new Error('ショートカットの数を読めない');
    if (d.controls !== 184 || d.uniqueControls !== 178) throw new Error('操作の数を読めない');
  });
  T('★数字が書いていなければ null（黙って0にしない）', () => {
    const d = readDoc('数字のない文');
    if (d.counted !== null || d.controls !== null) throw new Error('0 に化けた');
  });
  T('★在る＋無い＝数えた数 の形になっている（表の中で辻褄が合う）', () => {
    const d = readDoc('★42個を数えて 在る 21 ／ 無い 21★');
    if (d.have + d.missing !== d.counted) throw new Error('足しても合わない');
  });
  T('★一覧が空になっていない', () => { if (SHORTCUTS.length < 20) throw new Error('一覧が痩せている'); });
  console.log('\n  ── 実測 ── 確かめた ' + (pass + fail) + ' 通り / 通った ' + pass + ' 通り');
  console.log('\n' + pass + ' passed, ' + fail + ' failed');
  process.exit(fail ? 1 : 0);
}

/* ══ 本番（本物の book.html を載せて 実際に押す） ═══════════════════ */
console.log('\n[excel-parity] Excel との差を 機械で数え直す');

let JSDOM;
try { ({ JSDOM } = require('jsdom')); }
catch { console.log('  ✗ jsdom が無い＝押せない。★緑ではない★'); process.exit(1); }

const html = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf8');
const dom = new JSDOM(html.replace(/<script[\s\S]*?<\/script>/g, ''), {
  runScripts: 'dangerously', pretendToBeVisual: true, url: 'http://localhost/',
});
const win = dom.window, doc = win.document;
win.fetch = () => Promise.reject(new Error('no net'));
win.scrollTo = () => { }; win.alert = () => { };
const stub = new Proxy({}, { get: (t, k) => (k === 'measureText' ? (() => ({ width: 10 })) : k === 'canvas' ? { width: 800, height: 600 } : k === 'getImageData' ? (() => ({ data: [] })) : (() => { })) });
win.HTMLCanvasElement.prototype.getContext = () => stub;
const inject = (c) => { const el = doc.createElement('script'); el.textContent = c; doc.body.appendChild(el); };
for (const m of html.matchAll(/<script src="([^"]+)"><\/script>/g)) {
  const s = m[1].split('?')[0];
  if (/^https?:/.test(s)) continue;
  inject(fs.readFileSync(path.join(ROOT, s), 'utf8'));
}
for (const m of html.matchAll(/<script(?![^>]*\ssrc=)[^>]*>([\s\S]*?)<\/script>/g)) inject(m[1]);
await new Promise(r => { if (doc.readyState === 'complete') return r(); win.addEventListener('load', r); setTimeout(r, 3000); });

/* 材料を置く（空だと「効いたか」が分からない＝空振りする） */
win.setCell(0, 0, '10'); win.setCell(1, 0, '20'); win.setCell(2, 0, '30');
win.setCell(0, 1, 'あ'); win.setCell(1, 1, 'い');

function snapshot() {
  return JSON.stringify({
    sel: [win.selR1, win.selC1, win.selR2, win.selC2],
    sheet: win.activeSheet,
    data: JSON.stringify(win.sheets[win.activeSheet].data).length,
    undo: (win.undoStack || []).length, redo: (win.redoStack || []).length,
    editing: !!win.editingCell, scrollTop: win.scrollTop, scrollLeft: win.scrollLeft,
    dialogs: [...doc.querySelectorAll('div')].filter(d => d.style && d.style.display === 'flex').length,
    copy: win.copyBuffer ? JSON.stringify(win.copyBuffer).length : 0,
  });
}

const results = [];
for (const s of SHORTCUTS) {
  win.sel(1, 1, 1, 1);
  win.editingCell = null;
  const before = snapshot();
  const ev = new win.KeyboardEvent('keydown', {
    key: s.key, ctrlKey: !!s.ctrl, shiftKey: !!s.shift, altKey: !!s.alt, bubbles: true, cancelable: true,
  });
  try { doc.dispatchEvent(ev); } catch (e) { /* 落ちても続ける */ }
  const works = ev.defaultPrevented || before !== snapshot();
  results.push({ name: s.name, works });
}
const have = results.filter(r => r.works).length;
const missing = results.length - have;

/* ★ログインの窓は 数から外す（2026-08-23）★
   この表は ★Excel との差★を数える物。ログイン（メール・パスワード・ログイン・新規登録・忘れた）は
   ★Excel には無い物＝表計算の操作ではない★ので、混ぜると「Excelに近づいた」に見えて 嘘になる。 */
const controls = [...doc.querySelectorAll('button, select, input, [onclick]')]
  .filter((el) => !el.closest('#loginOv'));
const uniq = new Set(controls.map(el =>
  ((el.textContent || '').trim().replace(/\s+/g, ' ').slice(0, 30) || el.getAttribute('title') || el.id)
  + '|' + (el.getAttribute('onclick') || '')));

console.log('      ── 実測 ── ショートカット ' + results.length + '個を押して 在る ' + have + ' / 無い ' + missing);
console.log('      ── 実測 ── 画面の操作 ' + controls.length + '個（同じ物を除くと ' + uniq.size + '個）');

const docPath = path.join(ROOT, 'docs', 'EXCEL_PARITY.md');
T('★数えた表 docs/EXCEL_PARITY.md が在る', () => {
  if (!fs.existsSync(docPath)) throw new Error('表が無い＝数えていない');
});
const d = readDoc(fs.readFileSync(docPath, 'utf8'));
T('★表のショートカットの数が 今の実測と同じ（古くなっていない）', () => {
  if (d.counted !== results.length || d.have !== have || d.missing !== missing) {
    throw new Error('表[' + d.counted + '個/在る' + d.have + '/無い' + d.missing + '] ≠ 実測['
      + results.length + '個/在る' + have + '/無い' + missing + ']　→ docs/EXCEL_PARITY.md を直す');
  }
});
T('★表の「画面の操作」の数が 今の実測と同じ', () => {
  if (d.controls !== controls.length || d.uniqueControls !== uniq.size) {
    throw new Error('表[' + d.controls + '/' + d.uniqueControls + '] ≠ 実測['
      + controls.length + '/' + uniq.size + ']　→ docs/EXCEL_PARITY.md を直す');
  }
});
T('★押して確かめた数が0でない（検査が空振りしていない）', () => {
  if (have < 5) throw new Error('在ると数えた物が ' + have + ' 個しかない');
});

try { win.close(); } catch (e) { /* 閉じられなくても検査は済んでいる */ }
console.log('\n' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
