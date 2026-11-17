/* goal-seek-modosu-webkit.mjs ･･･ ★ゴールシークが 見つからず 戻す時 変える マスを 元の 形に 戻す★ 2026-11-17
 *
 *  ★★なぜ 要るか★★
 *    変える マスの 元の 値を 数に して 0 に 化かし、失敗の 時は 置く(0) で 戻して いた。
 *    ⇒ 式の マスは 0 の 定数に なり 式が 消え、字の マスも 0 に なって 「元に 戻しました」と 嘘を 出した。
 *  ★見る 物★ 空・字・書式つき空の マスが 失敗の 後 元と 同じ（data も 計算側 HF も）／式の マスは 断って 触らない
 *  ★走らせ方★ node tests/goal-seek-modosu-webkit.mjs
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

/* A1=5・B1=C1*0（目標 5 には 届かない）・C1 は 場面ごとに 作る */
const ws = XLSX.utils.aoa_to_sheet([[5, null, null]]);
ws['B1'] = { t: 'n', f: 'C1*0', v: 0 };
ws['!ref'] = 'A1:C1';
const wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, ws, 'あ');
const 本 = path.join(os.tmpdir(), 'exally-goal-seek-modosu.xlsx');
fs.writeFileSync(本, XLSX.write(wb, { bookType: 'xlsx', type: 'buffer' }));

console.log('[ゴールシーク 失敗の 戻し] ★変える マスの 元の 形が 残るか★');
const wk = await borrow('goal-seek-modosu', 'webkit');
const browser = await launch('goal-seek-modosu', wk, {}, 'webkit');
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
      && typeof window.setCell === 'function' && typeof window['ゴールシークをやる'] === 'function';
  }, null, { timeout: 120000 });
  await page.waitForTimeout(900);

  const 回す = (場面) => page.evaluate((場面) => {
    const d = window.sheets[window.activeSheet].data;
    const sid = window._hfSid(window.activeSheet);
    delete d['0,2'];
    if (場面.cell) d['0,2'] = JSON.parse(JSON.stringify(場面.cell));
    window.hf.setCellContents({ sheet: sid, row: 0, col: 2 }, 場面.hf === undefined ? null : 場面.hf);
    window.recalcSheet(window.activeSheet, d);
    const 形 = () => JSON.stringify(d['0,2'] === undefined ? null : d['0,2']);
    const 前 = 形(); const hf前 = window.hf.getCellValue({ sheet: sid, row: 0, col: 2 });
    document.getElementById('gsTarget').value = 'B1';
    document.getElementById('gsValue').value = '5';
    document.getElementById('gsChange').value = 'C1';
    const 返 = window['ゴールシークをやる']();
    const hf後 = window.hf.getCellValue({ sheet: sid, row: 0, col: 2 });
    return { 前, 後: 形(), hf前: String(hf前), hf後: String(hf後), 返: String(返),
      文: document.getElementById('gsWhy').textContent };
  }, 場面);

  const 例 = [
    ['空の マス', { cell: null }],
    ['字の マス（abc）', { cell: { v: 'abc', f: 'abc', d: '' }, hf: 'abc' }],
    ['書式つきの 空の マス', { cell: { v: '', f: '', d: '', bold: true, bgColor: '#FFEE99' } }],
    ['字の マス（007・ファイル由来 d字）', { cell: { v: '007', f: '007', d: '', d字: true }, hf: '007' }],
  ];
  for (const [名, 場面] of 例) {
    const r = await 回す(場面);
    T('★' + 名 + '＝失敗の 後 data が 元と 同じ★', r.後 === r.前, JSON.stringify(r));
    T('★' + 名 + '＝計算側（HF）も 元と 同じ★', r.hf後 === r.hf前, JSON.stringify(r));
    T('★' + 名 + '＝「戻しました」と 出る★', /元に 戻しました/.test(r.文), JSON.stringify(r));
  }
  const 式 = await 回す({ cell: { v: '=A1+1', f: '=A1+1', d: '6' }, hf: '=A1+1' });
  T('★式の マス＝断って 触らない★', 式.後 === 式.前 && /式では なく/.test(式.文) && 式.返 === '0', JSON.stringify(式));
} finally {
  await page.close().catch(() => {});
  await browser.close().catch(() => {});
  配信.閉じる();
}
console.log('\ngoal-seek-modosu: ' + pass + ' 緑 / ' + fail + ' 赤');
process.exit(fail ? 1 : 0);
