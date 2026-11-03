/* bahttext-webkit.mjs — ★本物の ブラウザで 実Excel の ブックを 開いて 計算し直す★ 2026-09-04
 *
 *  ★なぜ★
 *    部品（lib）と エンジンだけ 緑でも ★画面で 動く とは 限らない★。
 *    ★実Excel が 作った ブック★（B列＝=BAHTTEXT(A列) の 式／C列＝実Excel が 出した 答え）を
 *    ★お客さんの 道（ファイルを 開く）★で 読み込み、★うちが 計算し直した B列★と
 *    ★実Excel の 答え C列★を 1つずつ 見比べる。
 *
 *  ★見本★ tests/fixtures/bahttext-sample.xlsx（tools/bahttext-sample.ps1）
 *    ★司さんの 実物は 1バイトも 触っていない★（新規ブック）
 *
 *  走らせ方: node tests/bahttext-webkit.mjs
 */
import path from 'node:path'; import http from 'node:http'; import fs from 'node:fs';
import { fileURLToPath } from 'node:url';
import { borrow, launch } from '../scripts/_borrow-playwright.mjs';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const 見本 = path.join(ROOT, 'tests/fixtures/bahttext-sample.xlsx');
let pass = 0, fail = 0;
const T = (n, よい, 添え) => {
  if (よい) { pass++; console.log('  ok   ' + n); }
  else { fail++; console.log('  NG   ' + n + (添え ? '\n       ' + 添え : '')); }
};
function 立てる(root) {
  const 型 = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8',
    '.css': 'text/css; charset=utf-8', '.json': 'application/json; charset=utf-8',
    '.png': 'image/png', '.svg': 'image/svg+xml' };
  const s = http.createServer((q, r) => {
    const f = path.join(root, decodeURIComponent(String(q.url).split('?')[0]).replace(/^\/+/, ''));
    if (!f.startsWith(root) || !fs.existsSync(f) || fs.statSync(f).isDirectory()) { r.statusCode = 404; return r.end('no'); }
    r.setHeader('content-type', 型[path.extname(f).toLowerCase()] || 'application/octet-stream');
    fs.createReadStream(f).pipe(r);
  });
  return new Promise((x) => s.listen(0, '127.0.0.1', () => x({ url: 'http://127.0.0.1:' + s.address().port, 閉じる: () => s.close() })));
}

console.log('');
console.log('[bahttext-webkit] ★実Excel の ブックを 本物の ブラウザで 計算し直す★');
if (!fs.existsSync(見本)) { console.log('  ★未測定★ 見本が 在りません: ' + 見本); process.exit(1); }

const wk = await borrow('bahttext', 'webkit');
const browser = await launch('bahttext', wk, {}, 'webkit');
const page = await browser.newPage({ viewport: { width: 1280, height: 800 } });
const 配信 = await 立てる(ROOT);
try {
  await page.goto(配信.url + '/book.html', { waitUntil: 'load' });
  /* ★何で 出したか★＝手元・★鍵を 外した（ログインは 通していない）★・ファイルは 本物の 口 */
  await page.evaluate(() => {
    document.body.classList.remove('exally-locked');
    const ov = document.getElementById('loginOv');
    if (ov) { ov.classList.remove('open'); ov.style.display = 'none'; }
  });
  T('★画面が lib/bahttext.js を 読み込んでいる★',
    await page.evaluate(() => typeof window.BahtText === 'object' && !!window.BahtText.字にする));
  await page.setInputFiles('#bookFileInput', 見本);
  await page.waitForFunction(() => {
    const s = window.sheets && window.sheets[window.activeSheet];
    return s && s.data && Object.keys(s.data).length > 3;
  }, null, { timeout: 60000 });
  await page.waitForTimeout(1200);

  const 表 = await page.evaluate(() => {
    const d = (window.sheets[window.activeSheet].data) || {};
    const 出 = [];
    for (let r = 0; r < 12; r++) {
      const b = d[r + ',1'], c = d[r + ',2'];
      if (!b && !c) continue;
      出.push({ 行: r + 1,
        式: b && b.f, うち: b ? (b.d !== undefined ? b.d : b.v) : undefined,
        Excel: c ? (c.d !== undefined ? c.d : c.v) : undefined });
    }
    return 出;
  });
  T('★見本が 読めた（行が 在る）★', 表.length >= 7, '行 ' + 表.length);
  let 合 = 0, 見 = 0;
  for (const x of 表) {
    if (x.Excel === undefined || x.Excel === '') continue;
    見++;
    if (String(x.うち) === String(x.Excel)) 合++;
    else T('★' + x.行 + '行目★', false, '式 ' + x.式 + '／Excel「' + x.Excel + '」／うち「' + x.うち + '」');
  }
  T('★★実Excel の 答えと 1つ残らず 同じ（' + 合 + ' / ' + 見 + '）★★', 見 > 0 && 合 === 見,
    '見た ' + 見 + '行 ／ 合った ' + 合 + '行');
  console.log('       … 見た 総数 ' + 表.length + '行 ／ 見比べた ' + 見 + '行 ／ ★合った ' + 合 + '行★');
  console.log('       … 例 ' + 表.slice(0, 3).map((x) => x.式 + ' → ' + x.うち).join(' ／ '));
  if (process.env.EXALLY_SHOT) {
    await page.screenshot({ path: process.env.EXALLY_SHOT });
    console.log('       … 絵 ' + process.env.EXALLY_SHOT);
  }
} finally {
  配信.閉じる();
  await browser.close();
}
console.log('');
console.log('bahttext-webkit: ' + pass + ' 緑 / ' + fail + ' 赤');
process.exit(fail ? 1 : 0);
