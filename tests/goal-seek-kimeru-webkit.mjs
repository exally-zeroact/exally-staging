/* goal-seek-kimeru-webkit.mjs ･･･ ★ゴールシークの 見つかった後＝決める・やめる・Esc・取り消し★ 2026-11-18
 *
 *  ★見る 物★ 変える マスが 空・字・数・書式つきの 4通りで ①やめる ②Esc ③決める→Ctrl+Z ④Ctrl+Z→Ctrl+Y
 *  ★走らせ方★ node tests/goal-seek-kimeru-webkit.mjs
 */
import path from 'node:path'; import http from 'node:http'; import fs from 'node:fs';
import os from 'node:os';
import { createRequire } from 'node:module';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { borrow, launch } from '../scripts/_borrow-playwright.mjs';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(pathToFileURL(path.join(ROOT, 'package.json')));
const XLSX = require_(path.join(ROOT, 'lib/xlsx.full.min.js'));

let pass = 0, fail = 0;
const T = (n, ok, m) => {
  if (ok) { pass++; console.log('  ok   ' + n); }
  else { fail++; console.log('  NG   ' + n + (m ? '\n       ' + m : '')); }
};

function 立てる(root) {
  const 型 = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8',
    '.css': 'text/css; charset=utf-8', '.json': 'application/json; charset=utf-8',
    '.png': 'image/png', '.svg': 'image/svg+xml', '.ico': 'image/x-icon' };
  const s = http.createServer((q, r) => {
    const f = path.join(root, decodeURIComponent(String(q.url).split('?')[0]).replace(/^\/+/, ''));
    if (!f.startsWith(root) || !fs.existsSync(f) || fs.statSync(f).isDirectory()) { r.statusCode = 404; return r.end('no'); }
    r.setHeader('content-type', 型[path.extname(f).toLowerCase()] || 'application/octet-stream');
    fs.createReadStream(f).pipe(r);
  });
  return new Promise((x) => s.listen(0, '127.0.0.1', () => x({
    url: 'http://127.0.0.1:' + s.address().port, 閉じる: () => s.close(),
  })));
}

/* B1=C1*2（目標 10 → C1=5 が 見つかる）・D1=C1&"x" */
const ws = XLSX.utils.aoa_to_sheet([[5, null, null]]);
ws['B1'] = { t: 'n', f: 'C1*2', v: 0 };
ws['D1'] = { t: 's', f: 'C1&"x"', v: 'x' };
ws['!ref'] = 'A1:D1';
const wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, ws, 'あ');
const 本 = path.join(os.tmpdir(), 'exally-goal-seek-kimeru.xlsx');
fs.writeFileSync(本, XLSX.write(wb, { bookType: 'xlsx', type: 'buffer' }));

console.log('[ゴールシーク 見つかった後] ★やめる・Esc・決める・取り消し★');
const wk = await borrow('goal-seek-kimeru', 'webkit');
const browser = await launch('goal-seek-kimeru', wk, {}, 'webkit');
const 配信 = await 立てる(ROOT);
const page = await browser.newPage({ viewport: { width: 1280, height: 900 } });
page.on('pageerror', (e) => console.log('      [落ちた] ' + String(e.message).slice(0, 200)));
try {
  await page.goto(配信.url + '/book.html', { waitUntil: 'load', timeout: 120000 });
  await page.evaluate(() => {
    document.body.classList.remove('exally-locked');
    const ov = document.getElementById('loginOv');
    if (ov) { ov.classList.remove('open'); ov.style.display = 'none'; }
  });
  await page.setInputFiles('#bookFileInput', 本);
  await page.waitForFunction(() => {
    const ss = window.sheets || [];
    return ss.some((s) => s && s.data && Object.keys(s.data).length > 0)
      && typeof window.setCell === 'function' && typeof window['ゴールシークを決める'] === 'function';
  }, null, { timeout: 120000 });
  await page.waitForTimeout(900);

  const 回す = (場面, 流れ) => page.evaluate(async (a) => {
    const { 場面, 流れ } = a;
    const d = window.sheets[window.activeSheet].data;
    const sid = window._hfSid(window.activeSheet);
    delete d['0,2'];
    if (場面.cell) d['0,2'] = JSON.parse(JSON.stringify(場面.cell));
    window.hf.setCellContents({ sheet: sid, row: 0, col: 2 }, 場面.hf === undefined ? null : 場面.hf);
    window.recalcSheet(window.activeSheet, d);
    window.undoStack.length = 0; window.redoStack.length = 0;
    const 形 = () => JSON.stringify(d['0,2'] === undefined ? null : d['0,2']);
    const 状態 = () => ({ 形: 形(), hf: String(window.hf.getCellValue({ sheet: sid, row: 0, col: 2 })),
      D: String(window.hf.getCellValue({ sheet: sid, row: 0, col: 3 })), B: String(window.hf.getCellValue({ sheet: sid, row: 0, col: 1 })) });
    const 前 = 状態();
    window['ゴールシークを開く']();
    document.getElementById('gsTarget').value = 'B1';
    document.getElementById('gsValue').value = '10';
    document.getElementById('gsChange').value = 'C1';
    const 返 = window['ゴールシークをやる']();
    const 見つけ = 状態();
    const 画面 = { 求める: document.getElementById('gsGo').style.display, 決める: document.getElementById('gsOk').style.display };
    const 再 = window['ゴールシークをやる']();
    const 窓 = () => document.getElementById('goalOverlay').style.display;
    let 後 = null, 積み = window.undoStack.length;
    if (流れ === 'やめる') { window['ゴールシークを閉じる'](); 後 = 状態(); }
    else if (流れ === 'esc') { if (document.activeElement && document.activeElement.blur) document.activeElement.blur();
      document.body.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', keyCode: 27, bubbles: true, cancelable: true })); 後 = 状態(); }
    else {
      window['ゴールシークを決める'](); 積み = window.undoStack.length; const 決め後 = 状態();
      window.doUndo(); const 戻し = 状態();
      if (流れ === '決めてUndo') 後 = 戻し;
      else { window.doRedo(); 後 = 状態(); 後.戻し = 戻し; }
      後.決め後 = 決め後;
    }
    return { 前, 見つけ, 後, 返: String(返), 再: String(再), 画面, 積み, 窓: 窓(), 保留: typeof ゴールの保留 !== 'undefined' && !!ゴールの保留 };
  }, { 場面, 流れ });

  const 例 = [
    ['空の マス', { cell: null }],
    ['字の マス（abc）', { cell: { v: 'abc', f: 'abc', d: '' }, hf: 'abc' }],
    ['数の マス（3）', { cell: { v: '3', f: '3', d: '' }, hf: 3 }],
    ['書式つきの 空の マス', { cell: { v: '', f: '', d: '', bold: true, bgColor: '#FFEE99' } }],
  ];
  for (const [名, 場面] of 例) {
    for (const [流れ, 見出し] of [['やめる', 'やめる'], ['esc', 'Esc']]) {
      const r = await 回す(場面, 流れ);
      T('★' + 名 + '＝見つかった（C1=5）★', r.返 === '5' && r.見つけ.B === '10', JSON.stringify(r));
      T('★' + 名 + '＝見つかった後は 求めるを 隠し 決めるを 出す・2回目の 求めるは 断る★', r.画面.求める === 'none' && r.画面.決める === '' && r.再 === '0', JSON.stringify(r));
      T('★' + 名 + '＝' + 見出し + '＝data・計算側・つられる マスが 元と 同じ・窓が 閉じる★',
        r.後.形 === r.前.形 && r.後.hf === r.前.hf && r.後.D === r.前.D && r.後.B === r.前.B && r.窓 === 'none' && !r.保留 && r.積み === 0, JSON.stringify(r));
    }
    const u = await 回す(場面, '決めてUndo');
    T('★' + 名 + '＝決める＝取り消し1件・窓が 閉じ 値は 残る★', u.積み === 1 && u.後.決め後.hf === '5' && u.後.決め後.B === '10' && u.窓 === 'none', JSON.stringify(u));
    T('★' + 名 + '＝決める→Ctrl+Z＝元の 形★', u.後.形 === u.前.形 && u.後.hf === u.前.hf && u.後.D === u.前.D, JSON.stringify(u));
    const ur = await 回す(場面, '決めてUndoRedo');
    T('★' + 名 + '＝Ctrl+Z→Ctrl+Y＝見つけた 値★', ur.後.hf === '5' && ur.後.B === '10' && ur.後.D === '5x', JSON.stringify(ur));
  }
} finally {
  await page.close().catch(() => {});
  await browser.close().catch(() => {});
  配信.閉じる();
}
console.log('\ngoal-seek-kimeru: ' + pass + ' 緑 / ' + fail + ' 赤');
process.exit(fail ? 1 : 0);
