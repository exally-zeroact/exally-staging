/* hakaru-egaku-hayasa.mjs — ★1回 描くのに 何ミリ秒 かかるか★ 2026-09-24
 *
 *  ★★なぜ 測るか★★
 *    2026-09-24 に 見切れを 直す 中で `ctx.beginPath()` の 抜けを 見つけた。
 *    ★`rect()` は 今の 道に 足す★ので、`beginPath()` が 無いと
 *    ★描いた マスの 数だけ 道が 伸び続けます★。
 *    ⇒★「切り取りが だんだん 重くなる」のでは ないか★ と 疑った。
 *    ⇒★疑いのままに しない★＝★同じ 材料で 直す前／直した後を 測る★。
 *
 *  ★★自分の 疑いに 都合の よい 材料を 選ばない★★
 *    ・材料は ★引数で 渡します★（同じ 物を 2回 使う）
 *    ・★描くのは 20回★＝1回の 揺れを 均す（★一番 速い回・真ん中・一番 遅い回★を 出す）
 *    ・★描く 所（`_renderPass`）だけ★を 測る＝読み込みは 別（前の 紙に 在る）
 *
 *  ★1バイトも 書きません★（保存の 窓は 出しません）
 *
 *  走らせ方: node docs/measured/hakaru-egaku-hayasa.mjs <材料の道> [--札=なまえ]
 */
import path from 'node:path'; import http from 'node:http'; import fs from 'node:fs';
import { fileURLToPath } from 'node:url';
import { borrow, launch } from '../../scripts/_borrow-playwright.mjs';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '../..');
const 材料 = process.argv[2] || path.join(ROOT, 'tests/fixtures/cross-sheet-sample.xlsb');
const 札 = (process.argv.find((x) => x.indexOf('--札=') === 0) || '--札=（札なし）').slice('--札='.length);

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

if (!fs.existsSync(材料)) { console.log('★材料が 有りません★ ' + 材料); process.exit(1); }
console.log('★1回 描くのに 何ミリ秒 かかるか★  札 ' + 札);
console.log('  材料 ... ' + path.basename(材料) + '（' + fs.statSync(材料).size.toLocaleString() + ' バイト）');
console.log('  ★本の 字の 印★ sha256(book.html) = '
  + (await import('node:crypto')).createHash('sha256')
    .update(fs.readFileSync(path.join(ROOT, 'book.html'))).digest('hex').slice(0, 16));

const wk = await borrow('egaku-hayasa', 'webkit');
const browser = await launch('egaku-hayasa', wk, {}, 'webkit');
const page = await browser.newPage({ viewport: { width: 1400, height: 900 } });
const 配信 = await 立てる(ROOT);

try {
  await page.goto(配信.url + '/book.html', { waitUntil: 'load', timeout: 120000 });
  await page.evaluate(() => {
    document.body.classList.remove('exally-locked');
    const ov = document.getElementById('loginOv');
    if (ov) { ov.classList.remove('open'); ov.style.display = 'none'; }
    window.FileOut = { deliver: () => Promise.reject(new Error('★書き出しては いけません★')) };
  });
  await page.setInputFiles('#bookFileInput', 材料);
  await page.waitForFunction(() => {
    const ss = window.sheets || [];
    for (let i = 0; i < ss.length; i++) {
      if (ss[i] && ss[i].data && Object.keys(ss[i].data).length > 0) return true;
    }
    return false;
  }, null, { timeout: 900000 });
  await page.waitForTimeout(800);

  const 出 = await page.evaluate(() => {
    if (typeof window._renderPass !== 'function') return { だめ: '_renderPass が 有りません' };
    const みな = [];
    for (let i = 0; i < 20; i++) {
      const a = performance.now();
      window._renderPass();
      みな.push(performance.now() - a);
    }
    みな.sort((x, y) => x - y);
    const sh = window.sheets[window.activeSheet] || {};
    return {
      速い: みな[0], 真ん中: みな[10], 遅い: みな[19],
      板: sh.name, マス: Object.keys(sh.data || {}).length,
      板の数: (window.sheets || []).length,
    };
  });
  if (出.だめ) { console.log('  ★' + 出.だめ + '★'); process.exit(1); }
  console.log('');
  console.log('  板 ' + 出.板 + '（' + 出.マス.toLocaleString() + 'マス／全 ' + 出.板の数 + '枚）');
  console.log('  ★1回 描く★ ... 速い ' + Math.round(出.速い) + ' ms'
    + ' ／ ★真ん中 ' + Math.round(出.真ん中) + ' ms★'
    + ' ／ 遅い ' + Math.round(出.遅い) + ' ms（20回）');
} finally {
  await page.close().catch(() => {});
  await browser.close().catch(() => {});
  配信.閉じる();
}
