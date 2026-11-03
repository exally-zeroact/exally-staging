/* hakaru-fuda-wo-osu-to-kotae-ga-kawaru-ka.mjs ･･･ ★板の 札を 押すと 答えが 変わるか★ 2026-09-25
 *
 *  ★★なぜ 測るか★★
 *    経営者1 の 実測（2026-09-25）
 *      ★開いた だけ ..... 654個中 654個 ＝ 100.00% 実Excel と 同じ★
 *      ★札を 15枚 押した 後 ... 15,799個中 117個 違う★
 *    ⇒経営者1 の 見立て「★札を 押すと 計算し直されて 実Excel の 答えが 上書きされる★」
 *    ⇒但し `book.html` の `switchSheet` は ★もう 塞いで あります★（17473行）
 *       `if (!_開いた直後は計算しない && ... ) recalcSheet(...)`
 *    ⇒★どちらかが 古い物を 見て います★
 *    ⇒★だから 実Excel を 使わずに 「前と 後」だけで 数えます★
 *
 *  ★★何を 数えるか★★
 *    ⑴開いた 直後に ★全部の 板の 答え（`cell.d`）を 覚える★
 *    ⑵★画面の `switchSheet` で 板を 1枚ずつ 押して 回る★（★真似ません★）
 *    ⑶もう 一度 覚えて ★変わった マスを 数える★
 *    ⇒★0個なら 「押しても 上書きされない」★
 *    ⇒★1個でも 在れば その 板と 番地を 出します★（★中身は 出しません＝数と 番地だけ★）
 *
 *  ★これは 「印が 付いたか」では なく 「答えが 変わったか」です★
 *
 *  ★走らせ方★
 *    node docs/measured/hakaru-fuda-wo-osu-to-kotae-ga-kawaru-ka.mjs "<本の 道>" [--台 webkit]
 */
import path from 'node:path'; import http from 'node:http'; import fs from 'node:fs';
import { fileURLToPath } from 'node:url';
import { borrow, launch } from '../../scripts/_borrow-playwright.mjs';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..', '..');
const 引数 = process.argv.slice(2);
const 本 = 引数.filter((a) => a.slice(0, 2) !== '--')[0];
const 取る = (名, 既定) => { const i = 引数.indexOf('--' + 名); return i >= 0 ? 引数[i + 1] : 既定; };
const 台 = 取る('台', 'webkit');
/* ★★`--保存` ･･･ 札の 代わりに 「保存」を 押して 数えます★★
     `saveOpenedBook()` は 窓を 出す 前に ★板 15枚 全部を 計算し直して いました★。
     ⇒1マスも 打って いない お客さんの 答えが ★うちの 答えで 上書きされます★
     ⇒★書き出しは しません★（`FileOut.deliver` は 断る 物に 差し替えて います） */
const 保存で = 引数.indexOf('--保存') >= 0;
if (!本 || !fs.existsSync(本)) { console.log('★本が 在りません★ ' + 本); process.exit(2); }

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

console.log('[札を 押すと 答えが 変わるか] ★実Excel を 使わず 前と 後だけで 数えます★');
console.log('  ★本★ ' + path.basename(本) + '（★読むだけ／1バイトも 書きません★）／★台★ ' + 台);

const wk = await borrow('fuda-osu-kotae', 台);
const browser = await launch('fuda-osu-kotae', wk, {}, 台);
const page = await browser.newPage({ viewport: { width: 1280, height: 900 } });
const 配信 = await 立てる(ROOT);

try {
  let 声 = '';
  page.on('console', (m) => { const t = String(m.text()); if (t.indexOf('開いた 直後の 計算') >= 0) 声 = t; });
  page.on('pageerror', (e) => console.log('      [落ちた] ' + String(e.message).slice(0, 200)));
  await page.goto(配信.url + '/book.html', { waitUntil: 'load', timeout: 180000 });
  await page.evaluate(() => {
    document.body.classList.remove('exally-locked');
    const ov = document.getElementById('loginOv');
    if (ov) { ov.classList.remove('open'); ov.style.display = 'none'; }
    window.FileOut = { deliver: () => Promise.reject(new Error('★書き出しては いけません★')) };
  });
  await page.setInputFiles('#bookFileInput', 本);
  await page.waitForFunction(() => {
    const ss = window.sheets || [];
    return ss.some((s) => s && s.data && Object.keys(s.data).length > 0);
  }, null, { timeout: 180000 });
  await page.waitForTimeout(800);
  if (声) console.log('  [画面の 声] ' + 声);

  /* ★★要る 関数が 無ければ 数を 出さずに 止まります★★（2026-09-25 の 決め） */
  const 無い = await page.evaluate(() => ['switchSheet', 'sheets'].filter((n) => window[n] === undefined));
  if (無い.length) { console.log('  ★★測れません★★ 在りません ... ' + 無い.join(' / ')); process.exit(8); }

  const 覚える = () => page.evaluate(() => {
    /* ★式の マスの 答え（`d`）だけを 覚えます★＝定数は 変わりようが ありません */
    const 出 = {};
    (window.sheets || []).forEach((s, i) => {
      const d = (s && s.data) || {};
      for (const k in d) {
        const c = d[k];
        if (c && typeof c.f === 'string' && c.f.charAt(0) === '=') 出[i + '|' + k] = String(c.d);
      }
    });
    return 出;
  });

  const 前 = await 覚える();
  const 板の名 = await page.evaluate(() => (window.sheets || []).map((s) => s.name));
  console.log('  ★覚えた 式の マス★ ' + Object.keys(前).length.toLocaleString() + '個 ／ 板 ' + 板の名.length + '枚');

  /* ══ ★画面の 関数で 触ります★ ══（★真似ません★） */
  const 始め = Date.now();
  if (保存で) {
    const 無2 = await page.evaluate(() => (typeof window.saveOpenedBook === 'function' ? [] : ['saveOpenedBook']));
    if (無2.length) { console.log('  ★★測れません★★ 在りません ... saveOpenedBook'); process.exit(8); }
    /* ★窓が 出るので 待ちません＝計算し直しは 窓より 先に 走ります★ */
    await page.evaluate(() => { try { window.saveOpenedBook(); } catch (e) { window.__保存の落ち = String(e && e.message); } });
    await page.waitForTimeout(3000);
    const 落 = await page.evaluate(() => window.__保存の落ち || null);
    if (落) console.log('  ★保存で 落ちました★ ' + 落);
    console.log('  ★「保存」を 1回 押した★ ／ ' + (Date.now() - 始め).toLocaleString() + ' ms');
  } else {
    for (let i = 0; i < 板の名.length; i++) {
      await page.evaluate((n) => { window.switchSheet(n); }, i);
      await page.waitForTimeout(120);
    }
    console.log('  ★押して 回った★ ' + 板の名.length + '枚 ／ ' + (Date.now() - 始め).toLocaleString() + ' ms');
  }

  const 後 = await 覚える();

  const 変わった = [];
  for (const k in 前) if (前[k] !== (k in 後 ? 後[k] : 前[k])) 変わった.push(k);
  const 板ごと = {};
  変わった.forEach((k) => { const i = k.split('|')[0]; 板ごと[i] = (板ごと[i] || 0) + 1; });

  console.log('');
  console.log('  ══ ★答えが 変わった マス★ ══');
  console.log('    ★' + 変わった.length.toLocaleString() + '個★'
    + '（見た ' + Object.keys(前).length.toLocaleString() + '個 の '
    + (Object.keys(前).length ? (Math.round((変わった.length / Object.keys(前).length) * 10000) / 100) : 0) + '%）');
  if (変わった.length === 0) {
    console.log('    ⇒★★' + (保存で ? '「保存」を 押しても' : '札を 押しても') + ' 実Excel の 答えは 上書きされません★★');
    console.log('    ⇒★塞ぎが 効いて います★（`switchSheet` 17473行 ／ `saveOpenedBook` 20368行）');
  } else {
    Object.keys(板ごと).sort((a, b) => 板ごと[b] - 板ごと[a]).slice(0, 15).forEach((i) => {
      console.log('      ' + (板の名[i] || i) + ' ... ' + 板ごと[i].toLocaleString() + '個');
    });
    console.log('    ★番地の 例（★中身は 出しません★）★ ' + 変わった.slice(0, 8).map((k) => {
      const [i, rc] = k.split('|'); return (板の名[i] || i) + '!' + rc;
    }).join(' / '));
  }
  console.log('  ★これは 「答えが 変わったか」で 数えて います★（印では ありません）');
} finally {
  await page.close().catch(() => {});
  await browser.close().catch(() => {});
  配信.閉じる();
}
