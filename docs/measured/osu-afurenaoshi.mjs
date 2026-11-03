/* osu-afurenaoshi.mjs — ★じゃまを 消したら 溢れ直すか★を 本物の ブラウザで 押して 絵に する（2026-09-11）
 *
 *  ★★なぜ 作ったか★★
 *    ★実Excel★（COM で 打って 測った 2026-09-11）
 *      C3 に じゃま ⇒ C1「#スピル!」／★じゃまを 消す ⇒ C1..C5 = 1,2,3,4,5★（すぐ 溢れ直す）
 *    ★うち（直す 前）★
 *      じゃまを 消しても ★#SPILL! の まま★＝式を 打ち直すまで 直らない
 *    ⇒★借り物の 中は 読まず★、外から ★同じ 式を 入れ直す★事に した（book.html `_溢れ直しを試す`）
 *
 *  ★★入れた 直後に 画面を 固めました（2026-09-11）★★
 *    打ち直す → 計算し直す → まだ #SPILL → また 打ち直す …で ★ぐるぐる★
 *    ⇒★止め金（_溢れ直し中）を 付けた★＝1回の 計算で 1度だけ
 *    ⇒★この 道具は その 止め金が 効いて いる 事も 見ます★（固まったら 時間切れで 落ちる）
 *
 *  ★この 道具の 数は 画面の 数です★＝本物の ブラウザで ★マスを 押して 字を 打って★ います。
 *    押す 場所は 画面自身の `colX()/rowY()` から 取ります（当て推量の 座標を 使わない）。
 *  ★司さんの 実物には 触りません★＝手元に 立てた 配信の book.html だけ。
 *
 *  使い方: node docs/measured/osu-afurenaoshi.mjs
 */
import fs from 'node:fs';
import http from 'node:http';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { borrow, launch, unmeasured } from '../../scripts/_borrow-playwright.mjs';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..', '..');
const TAG = 'afurenaoshi';
const 出す先 = path.join(ROOT, 'docs/measured/e-afurenaoshi-2026-09-11.png');

/* ★立て方は tests/hyou-no-soto-webkit.mjs と 同じ★（読むだけ・127.0.0.1 だけ） */
function 立てる(root) {
  const 型 = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8',
    '.mjs': 'text/javascript; charset=utf-8', '.css': 'text/css; charset=utf-8',
    '.json': 'application/json; charset=utf-8', '.png': 'image/png', '.svg': 'image/svg+xml' };
  const s = http.createServer((req, res) => {
    const 道 = decodeURIComponent(String(req.url).split('?')[0]);
    const f = path.join(root, 道.replace(/^\/+/, ''));
    if (!f.startsWith(root) || !fs.existsSync(f) || fs.statSync(f).isDirectory()) { res.statusCode = 404; return res.end('no'); }
    res.setHeader('content-type', 型[path.extname(f).toLowerCase()] || 'application/octet-stream');
    fs.createReadStream(f).pipe(res);
  });
  return new Promise((r) => s.listen(0, '127.0.0.1', () => {
    r({ url: 'http://127.0.0.1:' + s.address().port, 閉じる: () => s.close() });
  }));
}

const chromium = await borrow(TAG, 'chromium');
if (!chromium) { unmeasured(TAG, 'chromium'); process.exit(0); }
const browser = await launch(TAG, chromium, {}, 'chromium');
if (!browser) { unmeasured(TAG, 'chromium'); process.exit(0); }
const page = await browser.newPage({ viewport: { width: 1000, height: 620 } });
const 配信 = await 立てる(ROOT);

let 悪い = 0;
const T = (n, よい, 添え) => {
  if (よい) console.log('  ok   ' + n);
  else { 悪い++; console.log('  NG   ' + n + (添え ? '\n       ' + 添え : '')); }
};

try {
  await page.goto(配信.url + '/book.html', { waitUntil: 'load' });
  /* ★鍵を 外す★＝ログインは 通って いません（手元の 配信だけ） */
  await page.evaluate(() => {
    document.body.classList.remove('exally-locked');
    const ov = document.getElementById('loginOv');
    if (ov) ov.style.display = 'none';
  });
  await page.waitForTimeout(500);

  /* ★押す 場所は 画面自身に 聞く★ */
  const 居所 = (r, c) => page.evaluate(([r, c]) => {
    const cv = document.getElementById('grid-canvas');
    const b = cv.getBoundingClientRect();
    return { x: b.left + colX(c) + 20, y: b.top + rowY(r) + 8 };
  }, [r, c]);
  /* ★切り取る 所も 画面に 聞く★＝当て推量の 座標を 焼き込まない
     ★A..D の 幅までに 止める★＝右に 出る 手引きの 吹き出しを 入れない */
  const 切り取り = async () => {
    const p = await 居所(0, 0);
    return { x: 0, y: Math.max(0, Math.round(p.y) - 26), width: 420, height: 190 };
  };

  const 打つ = async (r, c, 字) => {
    const p = await 居所(r, c);
    await page.mouse.click(p.x, p.y);
    await page.keyboard.type(字);
    await page.keyboard.press('Enter');
    await page.keyboard.press('Escape');
    await page.waitForTimeout(150);
  };
  const 消す = async (r, c) => {
    const p = await 居所(r, c);
    await page.mouse.click(p.x, p.y);
    await page.keyboard.press('Delete');
    await page.waitForTimeout(200);
  };
  /* ★画面の 板を そのまま 読む★（出て いる 字） */
  const 読む = () => page.evaluate(() => {
    const 板 = sheets[activeSheet];
    const 出 = [];
    for (let r = 0; r < 5; r++) {
      const c = 板.data[r + ',2'];
      出.push(c ? String(c.d != null && c.d !== '' ? c.d : (c.v != null ? c.v : '')) : '');
    }
    return 出;
  });

  console.log('');
  console.log('[' + TAG + '] ★本物の ブラウザで 押す★');

  /* ══ ①種 ══ */
  const 種 = [3, 1, 5, 2, 4];
  for (let i = 0; i < 種.length; i++) await 打つ(i, 0, String(種[i]));
  /* ══ ②じゃま ══ */
  await 打つ(2, 2, 'JAMA');
  /* ══ ③溢れる 式 ══ */
  await 打つ(0, 2, '=SORT(A1:A5)');

  const 前 = await 読む();
  T('★じゃまが 在る 時は #SPILL!★', String(前[0]).indexOf('SPILL') >= 0, 'C1=' + 前[0]);
  T('★じゃまの 字は 残る（実Excel と 同じ）★', 前[2] === 'JAMA', 'C3=' + 前[2]);
  await page.screenshot({ path: path.join(ROOT, 'docs/measured/_afure-mae.png'), clip: await 切り取り() });

  /* ══ ④じゃまを 消す ⇒★溢れ直すか★ ══ */
  await 消す(2, 2);
  await page.waitForTimeout(400);
  const 後 = await 読む();
  T('★じゃまを 消したら 溢れ直す★', String(後[0]).indexOf('SPILL') < 0, 'C1=' + 後[0]);
  T('★下まで 溢れる★', 後.filter((v) => v !== '').length === 5, JSON.stringify(後));
  T('★止め金が 効いて いる（固まって いない）★', true);
  await page.screenshot({ path: path.join(ROOT, 'docs/measured/_afure-ato.png'), clip: await 切り取り() });

  /* ══ ⑤2枚を 1枚に ══ */
  const 前絵 = fs.readFileSync(path.join(ROOT, 'docs/measured/_afure-mae.png')).toString('base64');
  const 後絵 = fs.readFileSync(path.join(ROOT, 'docs/measured/_afure-ato.png')).toString('base64');
  const 合 = await page.evaluate(async ([a, b]) => {
    const load = (s) => new Promise((r) => { const i = new Image(); i.onload = () => r(i); i.src = 'data:image/png;base64,' + s; });
    const [x1, x2] = await Promise.all([load(a), load(b)]);
    const P = 18, L = 38, W = x1.width, H = x1.height;
    const c = document.createElement('canvas');
    c.width = W * 2 + P * 3; c.height = H + L + P * 2;
    const g = c.getContext('2d');
    g.fillStyle = '#fff'; g.fillRect(0, 0, c.width, c.height);
    g.fillStyle = '#333'; g.font = 'bold 17px sans-serif';
    g.fillText('じゃまが 在る … #SPILL!', P, P + 22);
    g.fillText('じゃまを 消した … 溢れ直す', P * 2 + W, P + 22);
    g.drawImage(x1, P, P + L); g.drawImage(x2, P * 2 + W, P + L);
    g.strokeStyle = '#ccc'; g.strokeRect(P + 0.5, P + L + 0.5, W, H); g.strokeRect(P * 2 + W + 0.5, P + L + 0.5, W, H);
    return c.toDataURL('image/png').split(',')[1];
  }, [前絵, 後絵]);
  fs.writeFileSync(出す先, Buffer.from(合, 'base64'));
  fs.unlinkSync(path.join(ROOT, 'docs/measured/_afure-mae.png'));
  fs.unlinkSync(path.join(ROOT, 'docs/measured/_afure-ato.png'));
  console.log('');
  console.log('★絵を 書いた … ' + 出す先 + '（' + fs.statSync(出す先).size + ' バイト）★');
} catch (e) {
  悪い++;
  console.error('★落ちた … ' + e.message + '★');
} finally {
  await browser.close().catch(() => {});
  配信.閉じる();
}
console.log('');
console.log(悪い ? '★★違う ' + 悪い + '本★★' : '★★全部 合った★★');
process.exit(悪い ? 1 : 0);
