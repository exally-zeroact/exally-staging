/* kuraberu-hashira-haba.mjs
 *   -- ★実Excel と Exally で 列の 幅を 1列ずつ 並べる★（108）（2026-09-24）
 *
 *  ★★なぜ★★
 *    司さん「★Excel内で 見切れてない とこが 見切れとる★」
 *    ⇒★合わせる 相手（実Excel の 点）は 取って あります★
 *       `docs/measured/toru-jitsu-excel-no-hashira-haba.ps1`（107）
 *    ⇒★Exally 側の 点を 取って 並べます★
 *
 *  ★★前に 2回 外した★★（★同じ 穴を 踏まない ように 書いて おきます★）
 *    ⑴★HTML の マスを 数えようと した★
 *       ⇒Exally の 表は ★キャンバス（1枚の 絵）★＝★マスは 1つも 在りません★
 *       ⇒★どんなに 見切れて いても 必ず 0個★＝★偽の 緑★
 *    ⑵★2秒ごとに 画面へ 問い合わせた★
 *       ⇒★測る 行為が 測る 相手を 遅くしました★（2分で 開く 本が 300秒でも 開かない）
 *    ⇒★だから この 道具は ★開き終わってから 1回だけ★ 訊きます★
 *
 *  ★★取る 物（★画面の 内側の 数★）★★
 *    ・`sheets[i].colW[列]` ... ★Exally が 決めた 列の 点★
 *    ・`sheets[i].rowH[行]` ... 行の 点
 *    ・★一字の幅★（`measureText('0')` の 床）と ★本の 既定の 字★
 *    ・★マスごとの 字の 大きさ★の 出方（既定と 違う 物が 何個 在るか）
 *  ★★出さない 物★★ ... ★マスの 値／式の 字★（★板の 名と 数だけ★）
 *  ★★1バイトも 書きません★★
 *
 *  使い方: node docs/measured/kuraberu-hashira-haba.mjs "<xlsb>" [--秒 300]
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

const wk = await borrow('hashira-haba', 'webkit');
const browser = await launch('hashira-haba', wk, {}, 'webkit');
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
console.log('★渡します★（`#bookFileInput`）／★開き終わるまで 邪魔しません★');
await page.setInputFiles('#bookFileInput', 材料);
await page.waitForFunction(() => {
  const ss = window.sheets || [];
  let マス = 0;
  ss.forEach((sh) => { マス += Object.keys((sh && sh.data) || {}).length; });
  return マス > 0;
}, null, { timeout: 秒 * 1000 });

const 出 = await page.evaluate(() => {
  const ss = window.sheets || [];
  const 字幅 = (() => {
    try {
      const c = document.createElement('canvas').getContext('2d');
      const f = (window.sheets && window.sheets[0] && window.sheets[0].__font) || null;
      const 名 = (window.BookOpen && window.BookOpen.current && (window.BookOpen.current() || {}).既定の字) || '';
      const 出 = {};
      ['11px 游ゴシック', '11px sans-serif', '18px 游ゴシック'].forEach((s) => {
        c.font = s; 出[s] = c.measureText('0').width;
      });
      return { 名: 名, 実測: 出 };
    } catch (e) { return { だめ: String(e.message) }; }
  })();
  const 板 = ss.slice(0, 3).map((sh) => {
    const d = (sh && sh.data) || {};
    const 大きさ = {};
    Object.keys(d).forEach((k) => {
      const c = d[k];
      const s = (c && c.s && (c.s.fontSize || c.s.sz)) || (c && c.fontSize) || null;
      const な = s === null ? '(無し)' : String(s);
      大きさ[な] = (大きさ[な] || 0) + 1;
    });
    return {
      名: sh.name,
      マス: Object.keys(d).length,
      colW: sh.colW ? Object.keys(sh.colW).slice(0, 14).map((k) => [k, sh.colW[k]]) : [],
      rowH: sh.rowH ? Object.keys(sh.rowH).slice(0, 4).map((k) => [k, sh.rowH[k]]) : [],
      字の大きさ: Object.entries(大きさ).sort((a, b) => b[1] - a[1]).slice(0, 6),
    };
  });
  return { 秒: performance.now() - window.__t0, 板の数: ss.length, 字幅, 板 };
});

console.log('');
console.log('★開くまで★ ' + Math.round(出.秒) + ' ms ／ 板 ' + 出.板の数 + '枚');
console.log('★一字の幅（webkit で 実測）★ ' + JSON.stringify(出.字幅.実測));
出.板.forEach((b) => {
  console.log('');
  console.log('★板「' + b.名 + '」★ マス ' + b.マス.toLocaleString());
  console.log('  ★列の 点★ ' + b.colW.map(([c, w]) => c + ':' + w).join(' / '));
  console.log('  ★行の 点★ ' + b.rowH.map(([r, h]) => r + ':' + h).join(' / '));
  console.log('  ★字の 大きさの 出方★ ' + b.字の大きさ.map(([s, n]) => s + '=' + n + '個').join(' / '));
});
console.log('');
console.log('★落ち★ ' + 落ち.length + '件');
落ち.slice(0, 3).forEach((x) => console.log('   ' + x));
const 絵 = path.join(os.tmpdir(), 'exally-hashira-haba.png');
await page.screenshot({ path: 絵, fullPage: false });
console.log('★絵★ ' + 絵 + '（' + fs.statSync(絵).size + 'B）');
await browser.close();
server.close();
