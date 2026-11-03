/* hakaru-jitsubutsu-ga-hiraku-ka.mjs
 *   -- ★司さんの 実物が お客さんの 道で 開くか★（105）（2026-09-24）
 *
 *  ★★なぜ★★
 *    `hakaru-hozon-no-mado-no-hayasa.mjs` で 実物を 渡したら
 *    ★5分 待っても `window.sheets` に 中身が 入りませんでした★。
 *    ⇒★「遅い」のか 「落ちて いる」のか 分かりません★
 *    ⇒★分けずに 「遅い」と 書いたら 嘘に なります★
 *
 *  ★★この 道具が する 事★★
 *    ・言づて（console）と 落ち（pageerror）と 取れなかった 物を ★全部 拾う★
 *    ・★10秒ごとに 「今 どこまで 来たか」★を 出す（板の 数・マスの 数）
 *    ・★落ちたら その 字を そのまま 出す★
 *
 *  ★★出さない 物★★
 *    ★マスの 値／シートの 名／式の 字は 1つも 出しません★
 *    ＝出すのは ★数と 時間と 落ちの 字★だけ
 *
 *  ★★1バイトも 書きません★★
 *    `FileOut.deliver` を ★断る 物★に 替えてから 渡します（窓は 出しません）
 *
 *  使い方: node docs/measured/hakaru-jitsubutsu-ga-hiraku-ka.mjs "<xlsb への 道>" [--秒 300]
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
console.log('★待つ★ ' + 秒 + '秒 ／ ★1バイトも 書きません★');

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

const wk = await borrow('jitsubutsu-hiraku', 'webkit');
const browser = await launch('jitsubutsu-hiraku', wk, {}, 'webkit');
const page = await browser.newPage({ viewport: { width: 1280, height: 900 } });

const 言づて = [], 落ち = [], 取れず = [];
page.on('console', (m) => { if (m.type() === 'error' || m.type() === 'warning') 言づて.push(m.type() + ': ' + m.text().slice(0, 300)); });
page.on('pageerror', (e) => 落ち.push(String((e && e.message) || e).slice(0, 400)));
page.on('requestfailed', (r) => 取れず.push(r.url().replace(もと, '') + ' ... ' + (r.failure() || {}).errorText));
/* ★★2026-09-25 ── `Failed to load resource` の ★どれが★ を 捕まえます★★
     ＝`requestfailed` は ★繋がらなかった 時★だけ／★404 は 「取れた」扱い★です
     ⇒★200 以外を 全部 拾います★（★「取れなかった 0件」で 見落として いました★） */
page.on('response', (res) => {
  if (res.status() >= 400) 取れず.push('★' + res.status() + '★ ' + res.url().replace(もと, ''));
});

/* ══ ★★2026-09-25 ── ★黙って 止まる★ を 捕まえます★★ ══
     ＝★4回 中 3回 「1枚 / 0マス」の まま 止まりました★（落ち 0件・取れず 0件）
     ⇒★捕まえて いない 所が 在る★
     ⇒①★受け止められて いない 約束（unhandledrejection）★
       ②★読み込んだ 台 1本ずつの 成否★（`_loadScript` は 11本を 順に 読む）
       ③★200 以外の 返り★
     ★①が 一番 怪しい★＝★約束が 断られても 誰も 受け止めなければ 画面は 黙って 止まります★ */
await page.addInitScript(() => {
  window.__ｼ = { 断られ: [], 台: [] };
  window.addEventListener('unhandledrejection', (e) => {
    window.__ｼ.断られ.push(String((e.reason && e.reason.message) || e.reason).slice(0, 300));
  });
  window.addEventListener('error', (e) => {
    const t = e.target;
    if (t && (t.tagName === 'SCRIPT' || t.tagName === 'LINK')) {
      window.__ｼ.台.push('★取れません★ ' + (t.src || t.href || '?'));
    }
  }, true);
});
await page.goto(もと + '/book.html', { waitUntil: 'load', timeout: 120000 });
await page.evaluate(() => {
  document.body.classList.remove('exally-locked');
  const ov = document.getElementById('loginOv');
  if (ov) { ov.classList.remove('open'); ov.style.display = 'none'; }
  window.FileOut = { deliver: () => Promise.reject(new Error('★書き出しては いけません★')) };
  window.__t0 = performance.now();
});
console.log('');
console.log('★渡します★（`#bookFileInput`＝画面の「読み込む」と 同じ 口）');
await page.setInputFiles('#bookFileInput', 材料);

let 開いた = false;
let 断られ = '';
const 刻み = [];
for (let t = 0; t < 秒; t += 10) {
  await new Promise((r) => setTimeout(r, 10000));
  const い = await page.evaluate(() => {
    const ss = window.sheets || [];
    let マス = 0;
    ss.forEach((sh) => { マス += Object.keys((sh && sh.data) || {}).length; });
    /* ★★2026-09-25 ── ★止まって いる 段を 名指しします★★
         ＝「受け止められて いない 約束」も 「取れなかった 台」も ★0件★でした
         ＝★台は 揃って いるのに 進まない★のか、★台が まだ 来て いない★のか を 分けます
         ⇒★11本の 台が 窓に 居るか★を 1本ずつ 見ます（`_ensureXlsx` が 読む 順）
         ⇒★居ない 台が 在れば そこで 止まって います★
         ⇒★全部 居るのに 進まなければ 止まって いるのは `BookOpen.openFile` の 中★ */
    const 名たち = ['XLSX', 'XlsxIO', 'ZipSurgeon', 'XlsxEdit', 'XlsbEdit', 'XlsbJitai',
      'TableRefs', 'XlsxKazari', 'XlsxZukei', 'DiffPreview', 'BookOpen'];
    const 無い = 名たち.filter((n) => !window[n]);
    let 開いた = null;
    try { 開いた = !!(window.BookOpen && window.BookOpen.isOpened && window.BookOpen.isOpened()); } catch (e) { 開いた = '(訊けません)'; }
    const 知らせ = (document.querySelector('.toast, #toast, .toast-h') || {}).textContent || '';
    return { 板: ss.length, マス: マス, 秒: performance.now() - window.__t0,
      無い台: 無い, 開いた: 開いた, 知らせ: String(知らせ).slice(0, 120) };
  }).catch((e) => ({ だめ: String(e.message).slice(0, 200) }));
  if (い.だめ) { console.log('  ' + (t + 10) + '秒 ... ★窓に 訊けません★ ' + い.だめ); break; }
  刻み.push(い);
  console.log('  ' + (t + 10) + '秒 ... 板 ' + い.板 + '枚 ／ マス ' + い.マス.toLocaleString()
    + ' ／ ★無い台 ' + (い.無い台 || []).length + '本★' + ((い.無い台 || []).length ? '（' + い.無い台.join(',') + '）' : '')
    + ' ／ 開いた ' + い.開いた
    + (い.知らせ ? ' ／ 知らせ「' + い.知らせ + '」' : '')
    + (落ち.length ? ' ／ ★落ち ' + 落ち.length + '件★' : ''));
  if (い.マス > 0) { 開いた = true; console.log('  ⇒★開きました（' + Math.round(い.秒) + ' ms）★'); break; }
  /* ★★2026-09-25 ── ★断られたら すぐ 止めます★★
       ＝私は ★「開けませんでした」の 知らせを 見ずに★ 190秒 待って いました
       ＝★『何も 起きない』と 『断られた のに 見て いない』は 別★
       ⇒★知らせが 出たら その場で 終わり★（★待ち続けて 「止まった」と 言わない★） */
  if (い.知らせ && い.知らせ.indexOf('開けませんでした') >= 0) {
    console.log('  ⇒★★断られました（' + Math.round(い.秒) + ' ms）★★ ' + い.知らせ);
    断られ = い.知らせ;
    break;
  }
}

const しるし = await page.evaluate(() => window.__ｼ || { 断られ: [], 台: [] }).catch(() => ({ 断られ: [], 台: [] }));
console.log('');
console.log('★★答え★★ ... ' + (開いた ? '★開きました★' : (断られ ? '★断られました★ ' + 断られ : '★' + 秒 + '秒 経っても 何も 起きません★')));
console.log('★★受け止められて いない 約束★★ ' + しるし.断られ.length + '件');
しるし.断られ.slice(0, 5).forEach((x) => console.log('   ' + x));
console.log('★★台の 読み込みで 取れなかった 物★★ ' + しるし.台.length + '件');
しるし.台.slice(0, 5).forEach((x) => console.log('   ' + x));
console.log('★落ち（pageerror）★ ' + 落ち.length + '件');
落ち.slice(0, 5).forEach((x) => console.log('   ' + x));
console.log('★言づての 赤・黄★ ' + 言づて.length + '件');
言づて.slice(0, 8).forEach((x) => console.log('   ' + x));
console.log('★取れなかった 物★ ' + 取れず.length + '件');
取れず.slice(0, 5).forEach((x) => console.log('   ' + x));

/* ★★板を 切り替えてから 撮ります★★（2026-09-25）
     ＝`####` が 出て いると 言われた のは ★給料表★（Exally1 の 数え直し）
     ＝★1枚目だけ 撮っても そこは 見えません★
     ⇒`--板 給料表` で 切り替えます（★お客さんの 道＝タブを 押す★） */
const 見る板 = process.argv.includes('--板') ? process.argv[process.argv.indexOf('--板') + 1] : '';
if (見る板) {
  const 切れた = await page.evaluate((な) => {
    const ss = window.sheets || [];
    const i = ss.findIndex((s) => s && s.name === な);
    if (i < 0) return { だめ: '板が 在りません' };
    if (typeof switchSheet === 'function') { switchSheet(i); return { 板: i }; }
    if (typeof シートを切り替える === 'function') { シートを切り替える(i); return { 板: i }; }
    window.activeSheet = i;
    if (typeof render === 'function') render();
    return { 板: i, 手で: true };
  }, 見る板);
  console.log('★板を 切り替えました★ ' + 見る板 + ' ⇒ ' + JSON.stringify(切れた));
  await new Promise((r) => setTimeout(r, 4000));
}

const 絵 = path.join(os.tmpdir(), 'exally-jitsubutsu-hiraku.png');
await page.screenshot({ path: 絵, fullPage: false });
console.log('★絵★ ' + 絵 + '（' + fs.statSync(絵).size + 'B）');

await browser.close();
server.close();
process.exit(開いた ? 0 : 1);
