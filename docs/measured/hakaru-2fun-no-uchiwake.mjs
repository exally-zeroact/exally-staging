/* hakaru-2fun-no-uchiwake.mjs
 *   -- ★開くのに 2分 かかる ── その 内訳★（113）（2026-09-25）
 *
 *  ★★なぜ★★
 *    司さん「★遅すぎる ことないか★」
 *    ⇒実物は ★109〜115秒★で 開きます（4回 とも ほぼ 同じ）
 *    ⇒★どこで 使って いるかは 誰も 割って いません★
 *    ⇒★当てずっぽうで 直すと 見当違いを 直します★
 *
 *  ★★やり方（★repo の 字を 1文字も 変えません★）★★
 *    ブラウザの 中で ★上から 包んで 時間を 取ります★（`addInitScript`）
 *      ・`XLSX.read` ............. 包みを ほどいて 表に する（借り物）
 *      ・`BookOpen.openFile` ..... 開く 全体
 *      ・`BookOpen.sheetToGrid` .. 板 1枚を 画面の 形に する
 *      ・`initFormulaEngine` ..... 計算の 台を 作る
 *      ・`loadSheetIntoEngine` ... 板を 計算に 流す
 *      ・`render` ................ 描く
 *    ⇒★呼ばれた 回数と 合計の 時間★を 出します
 *    ★`window.XLSX` は 後から 出来ます★ので ★出来た 瞬間に 包みます★
 *      （`Object.defineProperty` で 待ち構える）
 *
 *  ★★出さない 物★★ ... ★マスの 値／シートの 名／式の 字★
 *  ★★1バイトも 書きません★★
 *
 *  使い方: node docs/measured/hakaru-2fun-no-uchiwake.mjs "<xlsb>" [--秒 300]
 */
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import http from 'node:http';
import crypto from 'node:crypto';
import { fileURLToPath } from 'node:url';
import { borrow, launch } from '../../scripts/_borrow-playwright.mjs';

const ここ = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(ここ, '..', '..');
const 材料 = process.argv[2];
const 秒 = process.argv.includes('--秒') ? Number(process.argv[process.argv.indexOf('--秒') + 1]) : 300;
if (!材料 || !fs.existsSync(材料)) { console.log('★材料が 在りません★'); process.exit(3); }
const 中 = fs.readFileSync(材料);
console.log('★材料★ ' + path.basename(材料) + '（' + 中.length.toLocaleString() + ' バイト）');
console.log('  sha256 ' + crypto.createHash('sha256').update(中).digest('hex'));

const 型 = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8', '.json': 'application/json', '.png': 'image/png',
  '.svg': 'image/svg+xml', '.ico': 'image/x-icon' };
const server = http.createServer((req, res) => {
  const u = decodeURIComponent(req.url.split('?')[0]);
  const p = path.join(ROOT, u === '/' ? 'book.html' : u.replace(/^\//, ''));
  if (!p.startsWith(ROOT) || !fs.existsSync(p) || fs.statSync(p).isDirectory()) { res.writeHead(404); res.end(); return; }
  res.writeHead(200, { 'Content-Type': 型[path.extname(p)] || 'application/octet-stream' });
  res.end(fs.readFileSync(p));
});
await new Promise((ok) => server.listen(0, '127.0.0.1', ok));
const もと = 'http://127.0.0.1:' + server.address().port;

const wk = await borrow('uchiwake', 'webkit');
const browser = await launch('uchiwake', wk, {}, 'webkit');
const page = await browser.newPage({ viewport: { width: 1280, height: 900 } });

/* ══ ★★2026-09-25 ── ★包む やり方は やめました★★ ══
     ＝`Object.defineProperty(window,'XLSX',...)` で 待ち構えたら
       ★画面が 落ちました★（`Target page, context or browser has been closed`）
     ＝★測る 為に 相手を 触ると 相手が 壊れます★（今日 2回目）
   ⇒★触らずに 「節目に いつ 着いたか」だけ 観ます★
     ・台が 窓に 出た 時刻（11本 それぞれ）
     ・`sheets` の 板が 増えた 時刻／マスが 入った 時刻
     ・描かれた 時刻（canvas に 色が 付いた）
   ⇒★差を 取れば どの 段に 何秒 かかったかが 出ます★ */

await page.goto(もと + '/book.html', { waitUntil: 'load', timeout: 120000 });
await page.evaluate(() => {
  document.body.classList.remove('exally-locked');
  const ov = document.getElementById('loginOv');
  if (ov) { ov.classList.remove('open'); ov.style.display = 'none'; }
  window.FileOut = { deliver: () => Promise.reject(new Error('★書き出しては いけません★')) };
  /* ★包む のは やめました★（上の 覚書き）＝★触らずに 観るだけ★ */
  window.__t0 = performance.now();
});
console.log('★渡します★／★開き終わるまで 邪魔しません★');
await page.setInputFiles('#bookFileInput', 材料);

let 開いた = 0, 断られ = '';
const 節目 = {};
const 名たち = ['XLSX', 'XlsxIO', 'ZipSurgeon', 'XlsxEdit', 'XlsbEdit', 'XlsbJitai',
  'TableRefs', 'XlsxKazari', 'XlsxZukei', 'DiffPreview', 'BookOpen'];
for (let t = 0; t < 秒; t += 1) {
  await new Promise((r) => setTimeout(r, 1000));
  const い = await page.evaluate((名たち) => {
    const ss = window.sheets || [];
    let マス = 0;
    ss.forEach((sh) => { マス += Object.keys((sh && sh.data) || {}).length; });
    const 在る = {};
    名たち.forEach((n) => { 在る[n] = !!window[n]; });
    let 描 = 0;
    try {
      const cv = document.getElementById('grid-canvas');
      if (cv && cv.width) {
        const g = cv.getContext('2d');
        const d = g.getImageData(0, 0, Math.min(cv.width, 300), Math.min(cv.height, 60)).data;
        for (let q = 0; q < d.length; q += 40) { if (d[q] < 250 || d[q + 1] < 250 || d[q + 2] < 250) 描 += 1; }
      }
    } catch (e) { 描 = -1; }
    const 知 = (document.querySelector('.toast, #toast, .toast-h') || {}).textContent || '';
    return { マス, 板: ss.length, 在る, 描, 秒: performance.now() - window.__t0, 知: String(知).slice(0, 80) };
  }, 名たち).catch(() => null);
  if (!い) break;
  名たち.forEach((n) => { if (い.在る[n] && !節目['台:' + n]) 節目['台:' + n] = い.秒; });
  if (い.板 > 1 && !節目['板が 増えた']) 節目['板が 増えた'] = い.秒;
  if (い.マス > 0 && !節目['マスが 入った']) 節目['マスが 入った'] = い.秒;
  if (い.描 > 0 && !節目['画面に 色が 付いた']) 節目['画面に 色が 付いた'] = い.秒;
  if (い.マス > 0 && い.描 > 0) { 開いた = い.秒; break; }
  if (い.知.indexOf('開けませんでした') >= 0) { 断られ = い.知; break; }
}
const 計 = {};

console.log('');
if (断られ) { console.log('★★断られました★★ ' + 断られ); }
else if (開いた) { console.log('★★開きました★★ ' + Math.round(開いた) + ' ms'); }
else { console.log('★★' + 秒 + '秒 経っても 何も 起きません★★'); }
console.log('');
console.log('★★節目に 着いた 時刻（★触らずに 観ただけ★）★★');
const 並 = Object.entries(節目).sort((a, b) => a[1] - b[1]);
let 前 = 0;
並.forEach(([な, ms]) => {
  console.log('  ' + String(Math.round(ms)).padStart(7) + ' ms  （+' + String(Math.round(ms - 前)).padStart(6) + '）  ' + な);
  前 = ms;
});
if (!並.length) console.log('  ★1つも 着いて いません★');

await browser.close();
server.close();
