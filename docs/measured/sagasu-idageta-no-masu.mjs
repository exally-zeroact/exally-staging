/* sagasu-idageta-no-masu.mjs
 *   -- ★`####` に なって いる マスの 場所を 出す★（111）（2026-09-24）
 *
 *  ★★なぜ★★
 *    実物で ★450マス★が `####` に なって いる（Exally1 の 実測）。
 *    ⇒★幅も 字の 大きさも 実Excel と 合って います★（私が 108/109 で 測った）
 *    ⇒★では なぜ `####` か★
 *    ⇒★★一番 先に 確かめる 事＝「実Excel でも そこは `####` なのか」★★
 *       ＝実Excel の `.Text` は ★幅が 足りなければ `#####` を 返します★
 *       ⇒★両方 `####` なら 直す 物は 在りません★
 *       ⇒★Excel だけ 数字なら そこが 欠陥★
 *
 *  ★★出さない 物★★
 *    ★マスの 値（数字そのもの）は 出しません★
 *    ＝出すのは ★板・行・列・桁数・測った 幅・列の 点★ だけ
 *
 *  ★★1バイトも 書きません★★
 *
 *  使い方: node docs/measured/sagasu-idageta-no-masu.mjs "<xlsb>" [--何個 20] [--秒 300]
 */
import fs from 'node:fs';
import path from 'node:path';
import http from 'node:http';
import crypto from 'node:crypto';
import { fileURLToPath } from 'node:url';
import { borrow, launch } from '../../scripts/_borrow-playwright.mjs';

const ここ = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(ここ, '..', '..');
const 材料 = process.argv[2];
const 何個 = process.argv.includes('--何個') ? Number(process.argv[process.argv.indexOf('--何個') + 1]) : 20;
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

const wk = await borrow('idageta', 'webkit');
const browser = await launch('idageta', wk, {}, 'webkit');
const page = await browser.newPage({ viewport: { width: 1280, height: 900 } });
await page.goto(もと + '/book.html', { waitUntil: 'load', timeout: 120000 });
await page.evaluate(() => {
  document.body.classList.remove('exally-locked');
  const ov = document.getElementById('loginOv');
  if (ov) { ov.classList.remove('open'); ov.style.display = 'none'; }
  window.FileOut = { deliver: () => Promise.reject(new Error('★書き出しては いけません★')) };
});
console.log('★渡します★／★開き終わるまで 邪魔しません★');
await page.setInputFiles('#bookFileInput', 材料);
await page.waitForFunction(() => {
  const ss = window.sheets || [];
  let m = 0; ss.forEach((sh) => { m += Object.keys((sh && sh.data) || {}).length; });
  return m > 0;
}, null, { timeout: 秒 * 1000 });

const 出 = await page.evaluate((何個) => {
  const ss = window.sheets || [];
  const 列の字 = (c) => { let s = ''; c += 1; while (c > 0) { s = String.fromCharCode(64 + ((c - 1) % 26 + 1)) + s; c = Math.floor((c - 1) / 26); } return s; };
  const 出 = []; let 数 = 0;
  for (let i = 0; i < ss.length && 出.length < 何個 * 3; i += 1) {
    const sh = ss[i]; const d = (sh && sh.data) || {};
    for (const k of Object.keys(d)) {
      const c = d[k];
      if (!c) continue;
      const [r, col] = k.split(',').map(Number);
      let 出字 = '';
      try { 出字 = String(typeof 表示の字 === 'function' ? 表示の字(c, r, col) : (c.d !== undefined ? c.d : c.v)); } catch (e) { 出字 = String(c.d !== undefined ? c.d : c.v); }
      if (!出字) continue;
      const 井桁 = /^#+$/.test(出字);
      if (!井桁) continue;
      数 += 1;
      if (出.length < 何個) {
        const w = (typeof cW === 'function') ? cW(col) : ((sh.colW && sh.colW[col]) || 72);
        const 生 = c.v;
        const 桁 = String(生 === undefined || 生 === null ? '' : 生).length;
        出.push({ 板: sh.name, 行: r + 1, 列: 列の字(col), 列番: col, 幅: w, 桁: 桁, 井桁の長さ: 出字.length });
      }
    }
  }
  return { 数, 出 };
}, 何個);

console.log('');
console.log('★★`####` に なって いる マス★★ ' + 出.数 + '個（★先頭 ' + 出.出.length + '個を 出します★）');
console.log('  板\t所\t列の点\t元の桁数\t井桁の数');
出.出.forEach((x) => {
  console.log('  ' + x.板 + '\t' + x.列 + x.行 + '\t' + x.幅 + '\t' + x.桁 + '\t' + x.井桁の長さ);
});
console.log('');
console.log('★★次に 実Excel で 同じ 所を 見ます★★（`.Text` が `#####` を 返すか）');
await browser.close();
server.close();
