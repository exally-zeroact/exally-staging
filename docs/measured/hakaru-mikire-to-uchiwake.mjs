/* hakaru-mikire-to-uchiwake.mjs
 *   -- ★見切れ★と ★2分の 内訳★を 実物で 数える★（106）（2026-09-24）
 *
 *  ★★なぜ★★
 *    司さん 2026-09-24
 *      「★読み込んだ ファイル 見切れとん とかも 自動で 調整しろ★」
 *      「★5分 待つって なんど／そんな 待てるか★」
 *    ⇒★直す 前に どこが どれだけか を 数えます★（★見立てで 直さない★）
 *
 *  ★★①見切れ★★
 *    ・★描かれた 字が マスより 広い★物を 数える（`scrollWidth > clientWidth`）
 *    ・★`#` で 埋まって いる★物を 数える（実Excel の「幅が 足りない」印）
 *    ・★どの 列に 何個★かまで 出す（★直す 所を 決められる ように★）
 *  ★★②内訳★★
 *    ・`performance` の 印（`mark`）では なく ★画面の 出来上がりで 刻みます★
 *      ⇒`sheets` に 入るまで ／ 描かれるまで ／ 計算が 終わるまで
 *  ★★出さない 物★★ ... ★マスの 値／シートの 名／式の 字★（数と 列の 字だけ）
 *  ★★1バイトも 書きません★★
 *
 *  使い方: node docs/measured/hakaru-mikire-to-uchiwake.mjs "<xlsb>" [--秒 300]
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
if (!材料 || !fs.existsSync(材料)) { console.log('★材料が 在りません★ ... ' + 材料); process.exit(3); }
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

const wk = await borrow('mikire-uchiwake', 'webkit');
const browser = await launch('mikire-uchiwake', wk, {}, 'webkit');
const page = await browser.newPage({ viewport: { width: 1280, height: 900 } });
const 落ち = [];
page.on('pageerror', (e) => 落ち.push(String((e && e.message) || e).slice(0, 300)));

await page.goto(もと + '/book.html', { waitUntil: 'load', timeout: 120000 });
await page.evaluate(() => {
  document.body.classList.remove('exally-locked');
  const ov = document.getElementById('loginOv');
  if (ov) { ov.classList.remove('open'); ov.style.display = 'none'; }
  window.FileOut = { deliver: () => Promise.reject(new Error('★書き出しては いけません★')) };
  window.__t0 = performance.now();
});

console.log('');
console.log('★★②内訳を 刻みます★★（★画面の 出来上がりで 見ます★）');
await page.setInputFiles('#bookFileInput', 材料);

const 刻み = [];
let 入った = 0, 描けた = 0;
for (let t = 0; t < 秒; t += 2) {
  await new Promise((r) => setTimeout(r, 2000));
  const い = await page.evaluate(() => {
    const ss = window.sheets || [];
    let マス = 0;
    ss.forEach((sh) => { マス += Object.keys((sh && sh.data) || {}).length; });
    const 描 = document.querySelectorAll('#grid .cell, #grid td, .grid-cell').length;
    return { 板: ss.length, マス: マス, 描: 描, 秒: performance.now() - window.__t0 };
  }).catch(() => null);
  if (!い) break;
  刻み.push(い);
  if (!入った && い.マス > 0) { 入った = い.秒; console.log('  ★台に 入った★ ' + Math.round(い.秒) + ' ms（板 ' + い.板 + '／マス ' + い.マス.toLocaleString() + '）'); }
  if (入った && !描けた && い.描 > 0) { 描けた = い.秒; console.log('  ★画面に 描けた★ ' + Math.round(い.秒) + ' ms（出た マス ' + い.描 + '個）'); }
  if (入った && 描けた) break;
}
if (!入った) { console.log('  ★' + 秒 + '秒 経っても 台に 入りません★'); }

/* ── ★①見切れを 数えます★ ── */
console.log('');
console.log('★★①見切れ★★（★描かれた 字が マスより 広い／`#` で 埋まって いる★）');
const 見切れ = await page.evaluate(() => {
  const 箱 = document.querySelectorAll('#grid .cell, #grid td, .grid-cell');
  const 列ごと = {};
  let はみ出し = 0, シャープ = 0, 見た = 0;
  箱.forEach((e) => {
    const 字 = (e.textContent || '').trim();
    if (!字) return;
    見た += 1;
    const 列 = e.dataset && e.dataset.c !== undefined ? e.dataset.c
      : (e.getAttribute && e.getAttribute('data-c')) || '?';
    const はみ = e.scrollWidth > e.clientWidth + 1;
    const しゃ = /^#+$/.test(字) || /^#{3,}/.test(字);
    if (はみ || しゃ) {
      列ごと[列] = (列ごと[列] || 0) + 1;
      if (はみ) はみ出し += 1;
      if (しゃ) シャープ += 1;
    }
  });
  return { 見た, はみ出し, シャープ, 列ごと };
});
console.log('  ★字の 在る マス（画面に 出て いる 分）★ ' + 見切れ.見た + '個');
console.log('  ★はみ出して いる★ ' + 見切れ.はみ出し + '個 ／ ★`#` で 埋まって いる★ ' + 見切れ.シャープ + '個');
const 並 = Object.entries(見切れ.列ごと).sort((a, b) => b[1] - a[1]).slice(0, 10);
並.forEach(([c, n]) => console.log('     列 ' + c + ' ... ' + n + '個'));

const 絵 = path.join(os.tmpdir(), 'exally-mikire.png');
await page.screenshot({ path: 絵, fullPage: false });
console.log('');
console.log('★絵★ ' + 絵 + '（' + fs.statSync(絵).size + 'B）');
console.log('★落ち★ ' + 落ち.length + '件');
落ち.slice(0, 3).forEach((x) => console.log('   ' + x));

await browser.close();
server.close();
