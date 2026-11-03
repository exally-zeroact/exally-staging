/* hakaru-1uchi-me-no-machi.mjs ･･･ ★1打ち目で 画面が どれだけ 固まるか★ 2026-09-25
 *
 *  ★★なぜ 測るか（経営者1 の 問い）★★
 *    「開いた 直後は 計算しない」に した ので
 *    ★重い 仕事は 消えて いません＝「最初の 1打ち」に 移りました★。
 *    ⇒経営者1「★1打ち目の 間 画面は どう 見えますか★」
 *      「★固まって 見えるなら 『待って います』の 印が 要ります★」
 *      「★前に 『開くまで 111秒 画面が 塞がって いた』と 測った 形です★」
 *      「★お客さんは 『壊れた』と 思って 閉じます★」
 *    ⇒★今 画面に 「待って います」の 印は 1つも 在りません★（`book.html` を 数えて 0件）
 *
 *  ★★何を 数えるか（★印では なく 絵が 動いたか★）★★
 *    `requestAnimationFrame` の 時刻を 全部 拾い、★隣との 空き★を 見ます。
 *      ・普通に 動いて いる ･･･ 16ms 前後
 *      ・★台が 塞がって いる ･･･ その間 1回も 来ません★
 *    ⇒★一番 長い 空き ＝ 画面が 固まって いた 時間★
 *    ⇒★これは 「印が 付いたか」では なく 「絵が 動いたか」です★
 *
 *  ★★打つ 所は 画面の 関数を 呼びます★★（2026-09-25 の 決め）
 *    `setCell` ＝ お客さんが 1マス 打った 時に 画面が 呼ぶ 物。★真似ません★
 *
 *  ★走らせ方★
 *    node docs/measured/hakaru-1uchi-me-no-machi.mjs "<本の 道>" [--台 webkit] [--待つ 90]
 *  ★出すのは 数だけ★（★本の 中身は 1字も 出しません★）
 */
import path from 'node:path'; import http from 'node:http'; import fs from 'node:fs';
import { fileURLToPath } from 'node:url';
import { borrow, launch } from '../../scripts/_borrow-playwright.mjs';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..', '..');
const 引数 = process.argv.slice(2);
const 本 = 引数.filter((a) => a.slice(0, 2) !== '--')[0];
const 取る = (名, 既定) => { const i = 引数.indexOf('--' + 名); return i >= 0 ? 引数[i + 1] : 既定; };
const 台 = 取る('台', 'webkit');
const 待つ秒 = Number(取る('待つ', '90'));
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

console.log('[1打ち目の 待ち] ★1打ち目で 画面が どれだけ 固まるか★');
console.log('  ★本★ ' + path.basename(本) + '（★読むだけ／1バイトも 書きません★）');
console.log('  ★台★ ' + 台 + ' ／ ★待つ★ ' + 待つ秒 + '秒');

const wk = await borrow('1uchi-me-no-machi', 台);
const browser = await launch('1uchi-me-no-machi', wk, {}, 台);
const page = await browser.newPage({ viewport: { width: 1280, height: 900 } });
const 配信 = await 立てる(ROOT);

try {
  let 声 = '';
  page.on('console', (m) => {
    const t = String(m.text());
    if (t.indexOf('開いた 直後の 計算') >= 0 || t.indexOf('計算を 始めます') >= 0) {
      声 += (声 ? ' ／ ' : '') + t.slice(0, 150);
    }
  });
  page.on('pageerror', (e) => console.log('      [落ちた] ' + String(e.message).slice(0, 200)));
  await page.goto(配信.url + '/book.html', { waitUntil: 'load', timeout: 180000 });
  await page.evaluate(() => {
    document.body.classList.remove('exally-locked');
    const ov = document.getElementById('loginOv');
    if (ov) { ov.classList.remove('open'); ov.style.display = 'none'; }
    window.FileOut = { deliver: () => Promise.reject(new Error('★書き出しては いけません★')) };
  });

  const 開く始め = Date.now();
  await page.setInputFiles('#bookFileInput', 本);
  await page.waitForFunction(() => {
    const ss = window.sheets || [];
    return ss.some((s) => s && s.data && Object.keys(s.data).length > 0);
  }, null, { timeout: 180000 });
  await page.waitForTimeout(800);
  console.log('  ★開くまで★ ' + (Date.now() - 開く始め).toLocaleString() + ' ms');
  if (声) console.log('  [画面の 声] ' + 声);

  /* ★絵が 動いたかを 拾い始めます★（`requestAnimationFrame` の 時刻） */
  await page.evaluate(() => {
    window.__刻 = [];
    const f = (t) => { window.__刻.push(t); requestAnimationFrame(f); };
    requestAnimationFrame(f);
  });
  await page.waitForTimeout(1500);
  const 前の空き = await page.evaluate(() => {
    const k = window.__刻; let m = 0;
    for (let i = 1; i < k.length; i++) m = Math.max(m, k[i] - k[i - 1]);
    return { 本数: k.length, 一番長い空き: Math.round(m) };
  });
  console.log('  ★打つ 前の 絵★ ' + 前の空き.本数 + '回 ／ ★一番 長い 空き ' + 前の空き.一番長い空き + ' ms★'
    + '（★これが 普段の 揺れです★）');

  /* ══ ★1マス 打ちます★ ══（★画面の 関数を 呼びます＝真似ません★）
       ★打つ 先は 遠い 空きマス★＝元の 数を 触りません（★どちらにしても 保存しません★） */
  await page.evaluate(() => {
    window.__刻 = [];
    window.__打ち始め = performance.now();
    window.__打ち終わり = null;
    /* ★台を 止めない為に 次の 絵で 打ちます★（打つ 前の 絵を 1枚 出させる） */
    requestAnimationFrame(() => {
      try { window.setCell(900, 30, '1'); } catch (e) { window.__打ちの落ち = String(e && e.message); }
      window.__打ち終わり = performance.now();
    });
  });

  /* ★台が 塞がって いる 間は ここも 返って きません★＝だから 短い 待ちを 重ねます */
  const 締め = Date.now() + 待つ秒 * 1000;
  let 済み = false;
  while (Date.now() < 締め) {
    await page.waitForTimeout(500);
    済み = await page.evaluate(() => window.__打ち終わり !== null).catch(() => false);
    if (済み) break;
  }

  const 出 = await page.evaluate(() => {
    const k = window.__刻; const 空き = [];
    for (let i = 1; i < k.length; i++) 空き.push(k[i] - k[i - 1]);
    空き.sort((a, b) => b - a);
    return {
      打ちの時間: window.__打ち終わり === null ? null : Math.round(window.__打ち終わり - window.__打ち始め),
      落ちた: window.__打ちの落ち || null,
      絵の本数: k.length,
      一番長い空き: 空き.length ? Math.round(空き[0]) : null,
      二番目: 空き.length > 1 ? Math.round(空き[1]) : null,
      半秒以上の空き: 空き.filter((x) => x >= 500).length,
      半秒以上の合計: Math.round(空き.filter((x) => x >= 500).reduce((a, b) => a + b, 0)),
    };
  });

  console.log('');
  console.log('  ══ ★1打ち目★ ══');
  console.log('    打ちに かかった 時間 ..... ' + (出.打ちの時間 === null
    ? '★' + 待つ秒 + '秒 待っても 返って きませんでした★' : 出.打ちの時間.toLocaleString() + ' ms'));
  if (出.落ちた) console.log('    ★打つと 落ちました★ ' + 出.落ちた);
  console.log('    ★画面が 固まって いた 一番 長い 所★ ... '
    + (出.一番長い空き === null ? '測れません' : '★' + 出.一番長い空き.toLocaleString() + ' ms★'));
  console.log('    その 次に 長い 所 ............. ' + (出.二番目 === null ? '-' : 出.二番目.toLocaleString() + ' ms'));
  console.log('    半秒 以上 止まった 回数 ....... ' + 出.半秒以上の空き + '回'
    + '（合わせて ' + 出.半秒以上の合計.toLocaleString() + ' ms）');
  console.log('    絵が 出た 回数 ................ ' + 出.絵の本数 + '回');
  console.log('');
  /* ══ ★★絵が 1回しか 出て いない 時に 「固まり 0 ms」と 書いては いけません★★ ══（2026-09-25）
       1回目に そう 書きました。★空きは 2回 以上 無いと 計算できません★ ので
       ★一番 強い 証し（絵が 1回）を 「0 ms」に 化けさせて いました★。
       ⇒★絵が 1回 以下 なら 固まりは 「打ちに かかった 時間 まるごと」です★
       ⇒★分母（絵が 何回 出たか）を 必ず 並べます★ */
  const 絵が足りない = 出.絵の本数 <= 1;
  const 固 = 絵が足りない ? (出.打ちの時間 || 0) : (出.一番長い空き || 0);
  if (絵が足りない) {
    console.log('  ★★絵は ' + 出.絵の本数 + '回しか 出て いません★★'
      + '＝★空きを 数える 事すら できません★');
    console.log('    ⇒★だから 固まりは 「打ちに かかった 時間 まるごと」で 数えます★'
      + '（★' + (出.打ちの時間 === null ? '返って こなかった' : 出.打ちの時間.toLocaleString() + ' ms') + "★）");
  }
  if (固 >= 2000) {
    console.log('  ⇒★★画面は ' + (固 / 1000).toFixed(1) + '秒 固まって います★★');
    console.log('    ＝★この間 「待って います」の 印は 1つも 出ません★（`book.html` に 0件）');
    console.log('    ⇒★お客さんは 「壊れた」と 思って 閉じます★＝★印が 要ります★');
  } else {
    console.log('  ⇒固まった 一番 長い 所は ' + 固 + ' ms ＝★普段の 揺れ（' + 前の空き.一番長い空き + ' ms）と 並べて 見て ください★');
    console.log('    （分母＝打った 後に 絵が ' + 出.絵の本数 + '回 出ました／打つ 前は ' + 前の空き.本数 + '回）');
  }
  console.log('  ★これは 「絵が 動いたか」で 数えて います★（印では ありません）');
} finally {
  await page.close().catch(() => {});
  await browser.close().catch(() => {});
  配信.閉じる();
}
