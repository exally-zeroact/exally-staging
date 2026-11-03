/* kazoeru-1uchime-no-iriguchi.mjs ･･･ ★1打ち目の 間に 台の 入口が 何回 呼ばれたか★ 2026-10-05
 *
 *  ★★なぜ★★
 *    本番 14955ae（10-05）は 開くのが 速く なった 代わりに ★最初の 1マスで 11.5秒（chromium）／17.2秒（WebKit）固まる★。
 *    CPU の 記録（道具114 --1打ち目）では 1打ち目の 79% が 式の 台（lib/shiki-hyou.js）＝
 *      読む揃った名で 18.7% ／ 引数にする 10.5% ／ 番地から名 10.4% ／ 直す 10.1% ／ 名から番地 6.1% ／ 四角を読む 5.5%
 *    ★割合だけでは 直し方は 決まらない★（記憶）＝★回数と 別々の 種類★まで 数える
 *  ★★数える 物（台の 口 `数えを出す()`）★★
 *    四角を読む（回・中の マス・索引に 在った マス）／読む／揃った名で（回・中身に 無かった）／計算する／
 *    ★四角の 種類★（別々の 四角の 数・2回以上 読んだ 四角・読み直しの 回数）／空と答えた数
 *  ★打つ 所は 画面の 関数★（switchSheet(0) → setCell(5000,0,'1')＝道具131 と 同じ 所）
 *  ★秒は 出すが 目安に しない★（数えの 加算と 種類の 表の 重みが 乗る＝秒は 道具131 で 別に 測る）
 *  ★走らせ方★ node docs/measured/kazoeru-1uchime-no-iriguchi.mjs "<本の 道>" [--台 chromium|webkit] [--種類 なし]
 *  ★出すのは 数だけ★（★本の 中身は 1字も 出しません・1バイトも 書きません★）
 */
import path from 'node:path'; import http from 'node:http'; import fs from 'node:fs';
import { fileURLToPath } from 'node:url';
import { execSync } from 'node:child_process';
import { borrow, launch } from '../../scripts/_borrow-playwright.mjs';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..', '..');
const 引数 = process.argv.slice(2);
const 本 = 引数.filter((a) => a.slice(0, 2) !== '--' && !['chromium', 'webkit', 'なし'].includes(a))[0];
const 取る = (名, 既定) => { const i = 引数.indexOf('--' + 名); return i >= 0 ? 引数[i + 1] : 既定; };
const 台 = 取る('台', 'chromium');
const 種類も = 取る('種類', 'あり') !== 'なし';
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

let 印 = '?';
try { 印 = execSync('git rev-parse --short HEAD', { cwd: ROOT }).toString().trim() + (execSync('git status --porcelain', { cwd: ROOT }).toString().trim() ? '（手元に 未commit あり）' : '（手元は 綺麗）'); } catch (e) { /* 印が 取れない 時は ? */ }
console.log('[1打ち目の 入口] 本 ' + path.basename(本) + ' ／ 台 ' + 台 + ' ／ 木 ' + 印 + ' ／ 四角の 種類 ' + (種類も ? '取る' : '取らない'));

const wk = await borrow('1uchime-no-iriguchi', 台);
const browser = await launch('1uchime-no-iriguchi', wk, {}, 台);
const page = await browser.newPage({ viewport: { width: 1280, height: 900 } });
const 配信 = await 立てる(ROOT);
let 終わり値 = 0;
try {
  page.on('pageerror', (e) => console.log('      [落ちた] ' + String(e.message).slice(0, 200)));
  await page.goto(配信.url + '/book.html', { waitUntil: 'load', timeout: 180000 });
  await page.evaluate(() => {
    document.body.classList.remove('exally-locked');
    const ov = document.getElementById('loginOv');
    if (ov) { ov.classList.remove('open'); ov.style.display = 'none'; }
    window.FileOut = { deliver: () => Promise.reject(new Error('★書き出しては いけません★')) };
  });
  await page.setInputFiles('#bookFileInput', 本);
  await page.waitForFunction(() => (window.sheets || []).some((s) => s && s.data && Object.keys(s.data).length > 0), null, { timeout: 300000 });
  await page.waitForTimeout(800);
  const 出 = await page.evaluate((種類も) => {
    let 表 = null;
    try { if (typeof window._台の板を得る === 'function') 表 = window._台の板を得る(); } catch (e) { 表 = null; }
    if (!表 || typeof 表.数えを始める !== 'function' || typeof 表.数えを出す !== 'function') return { 誤り: '台の 数えの 口が 無い' };
    const 前の空 = 表.空と答えた数();
    表.数えを始める(種類も);
    const t = performance.now();
    window.switchSheet(0); window.setCell(5000, 0, '1');
    const 秒 = Math.round(performance.now() - t);
    /* ★打った 後に 台が 作り直されて いないか★（作り直されたら 数えは 新しい 台に 無い） */
    let 表2 = null;
    try { 表2 = window._台の板を得る(); } catch (e) { 表2 = null; }
    /* ★答えの 指紋★＝画面の 全部の マスの 出る 字（板・番地・字）を 並べて 1本の 字に（★中身は 出さない★） */
    const 並び = [];
    (window.sheets || []).forEach((sh, i) => { const d = (sh && sh.data) || {}; Object.keys(d).sort().forEach((k) => { const c = d[k]; if (!c || typeof c !== 'object') return; 並び.push(i + '|' + k + '|' + String(c.d === undefined ? '' : c.d) + '|' + String(c.v === undefined ? '' : c.v)); }); });
    /* ★台の 答えの 指紋★＝画面の 字が 借り物から 来て いても 台の 直しを 見られる 様に、台の 全部の マスの 答えを 直に 並べる
         （`数えを出す` を 先に 取る＝ここで 読むと 計算が 走って 数えが 増える 為） */
    const 数 = 表.数えを出す();
    const 台の並び = Object.keys(表.中身 || {}).sort().map((n) => { let t = ''; try { t = String(表.字(n)); } catch (e) { t = '★投げた★'; } return n + '|' + t; });
    return { 秒, 同じ台: 表2 === 表, 数, 台の字: 台の並び.join('\n'), 台のマス: 台の並び.length, 空と答えた: 表.空と答えた数() - 前の空, 索引の数: 表.索引の数(), 答えの字: 並び.join('\n'), 答えのマス: 並び.length };
  }, 種類も);
  if (出.答えの字 !== undefined) { const h = (await import('node:crypto')).createHash('sha256').update(出.答えの字).digest('hex'); 出.答えの指紋 = h.slice(0, 16); delete 出.答えの字; }
  if (出.台の字 !== undefined) { 出.台の指紋 = (await import('node:crypto')).createHash('sha256').update(出.台の字).digest('hex').slice(0, 16); delete 出.台の字; }
  if (出.誤り) { console.log('★' + 出.誤り + '★'); 終わり値 = 8; }
  else {
    console.log('  1打ち目 ' + 出.秒.toLocaleString() + ' ms（★数えの 重みが 乗る＝目安に しない★）／ 打つ 前後で 同じ 台 ' + 出.同じ台);
    const f = (n) => (typeof n === 'number' ? n.toLocaleString() : String(n));
    for (const k of Object.keys(出.数)) console.log('  ' + k + ' ... ' + f(出.数[k]));
    console.log('  空と答えた（四角の 中で 索引に 無かった）... ' + f(出.空と答えた));
    console.log('  索引の数 ... ' + f(出.索引の数));
    console.log('  ★答えの 指紋（1打ち目の 後の 画面の 全部の マス d/v）★ ' + 出.答えの指紋 + ' ／ マス ' + f(出.答えのマス));
    console.log('  ★台の 答えの 指紋（台の 全部の マスの 字）★ ' + 出.台の指紋 + ' ／ マス ' + f(出.台のマス));
    if (!出.同じ台) { console.log('★打った 後に 台が 替わった＝数えは 前の 台の 物★'); 終わり値 = 1; }
    if (出.数.計算する === 0) { console.log('★計算する が 0回＝1打ち目で 何も 計算して いない（測る 所を 外した）★'); 終わり値 = 1; }
  }
} finally { await page.close(); await browser.close(); 配信.閉じる(); }
process.exit(終わり値);
