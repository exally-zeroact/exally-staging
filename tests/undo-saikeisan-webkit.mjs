/* undo-saikeisan-webkit.mjs ･･･ ★取り消し（Ctrl+Z）の後も つられた マスを 計算し直す★ 2026-11-17
 *
 *  ★★なぜ 要るか★★
 *    A1 を 7 に 打つと B1（=A1*3+5）は 26 に なる。取り消すと A1 は 5 に 戻るが、
 *    ★B1 は 26 の まま★（実測・500ms 待っても）＝合計が 古い まま 見える。
 *    `doUndo` の 値の 道は `render(); updateBar();` だけで 波及を 起こして いなかった
 *    （打つ 側の setCell は `_scheduleRecalc` を 呼ぶ）。
 *
 *  ★見る 物★ ⑴打つと B1=26／⑵取り消すと A1=5 と ★B1=20 と C1=21（2段）★／⑶やり直すと B1=26
 *  ★走らせ方★ node tests/undo-saikeisan-webkit.mjs
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

/* A1=5・B1=A1*3+5（20）・C1=B1+1（21）＝2段の つながり */
const ws = XLSX.utils.aoa_to_sheet([[5, null, null]]);
ws['B1'] = { t: 'n', f: 'A1*3+5', v: 20 };
ws['C1'] = { t: 'n', f: 'B1+1', v: 21 };
ws['!ref'] = 'A1:C1';
const wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, ws, 'あ');
const 本 = path.join(os.tmpdir(), 'exally-undo-saikeisan.xlsx');
fs.writeFileSync(本, XLSX.write(wb, { bookType: 'xlsx', type: 'buffer' }));

console.log('[取り消しの 後の 再計算] ★つられた マスが 古い まま 残らないか★');
const wk = await borrow('undo-saikeisan', 'webkit');
const browser = await launch('undo-saikeisan', wk, {}, 'webkit');
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
      && typeof window.setCell === 'function' && typeof window.doUndo === 'function' && typeof window.doRedo === 'function';
  }, null, { timeout: 120000 });
  await page.waitForTimeout(900);

  const 見る = () => page.evaluate(() => {
    const d = window.sheets[window.activeSheet].data;
    const g = (k) => (d[k] ? String(d[k].d !== undefined && d[k].d !== '' ? d[k].d : d[k].v) : '（無し）');
    return { A1: g('0,0'), B1: g('0,1'), C1: g('0,2') };
  });

  const 前 = await 見る();
  T('★始め＝A1=5・B1=20・C1=21★', 前.A1 === '5' && 前.B1 === '20' && 前.C1 === '21', JSON.stringify(前));

  await page.evaluate(() => { window.setCell(0, 0, '7'); });
  await page.waitForTimeout(700);
  const 打 = await 見る();
  T('★打つと B1=26・C1=27（打つ 側は もとから 波及する）★', 打.A1 === '7' && 打.B1 === '26' && 打.C1 === '27', JSON.stringify(打));

  await page.evaluate(() => { window.doUndo(); });
  await page.waitForTimeout(700);
  const 戻 = await 見る();
  T('★取り消すと A1=5★', 戻.A1 === '5', JSON.stringify(戻));
  T('★★取り消すと つられた B1 も 20 に 戻る★★（★古い 26 の まま に しない★）', 戻.B1 === '20', JSON.stringify(戻));
  T('★取り消すと 2段目の C1 も 21 に 戻る★', 戻.C1 === '21', JSON.stringify(戻));

  await page.evaluate(() => { window.doRedo(); });
  await page.waitForTimeout(700);
  const 再 = await 見る();
  T('★やり直すと A1=7・B1=26・C1=27★', 再.A1 === '7' && 再.B1 === '26' && 再.C1 === '27', JSON.stringify(再));
} finally {
  await page.close().catch(() => {});
  await browser.close().catch(() => {});
  配信.閉じる();
}
console.log('\nundo-saikeisan: ' + pass + ' 緑 / ' + fail + ' 赤');
process.exit(fail ? 1 : 0);
