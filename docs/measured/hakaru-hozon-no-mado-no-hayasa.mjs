/* hakaru-hozon-no-mado-no-hayasa.mjs — ★保存の 窓が 出るまでの 時間★ 2026-09-22
 *
 *  ★★なぜ 測るか★★
 *    2026-09-22 に 私が 保存の 窓へ ★本の 中身を 数える★ 所を 足しました
 *    （`HonNoNakami.本の中身を数える`＝★全部の 板の 全部の 式を 舐めます★）。
 *    ⇒小さい 本（板 4枚）では 気に なりませんでした。
 *    ⇒★大きい 本で 測って いません★＝★「遅くない」と 言っては いけません★。
 *    ⇒★自分で 足した 重さは 自分で 測ります★。
 *
 *  ★★何を 何で 数えるか★★
 *    何で ... ★本物の webkit★（playwright）／★お客さんの 道★
 *             `#bookFileInput` に ファイルを 渡す＝画面の「読み込む」と 同じ
 *    どこを . `performance.now()` を ★ブラウザの 中で★ 取る（node 側の 時計では ない）
 *    3つ ... ①開く（読み込み終わるまで）
 *            ②★中身を 数える だけ★（`_本の中身を数えた物()` を 1回）
 *            ③★保存の 窓が 出るまで★（`saveOpenedBook()` ⇒ 窓が flex に なるまで）
 *    ⇒②が ③の 中で どれだけを 占めるかが ★私が 足した 重さ★。
 *
 *  ★★1バイトも 書きません★★
 *    ・`FileOut.deliver` を ★必ず 断る 物★に 差し替えてから 窓を 出します
 *    ・窓は ★[やめる]★ で 閉じます
 *    ・司さんの 実物を 材料に する 時も ★読むだけ★（playwright は 読むだけ）
 *
 *  ★★実物の 中身は 出しません★★（司さんの 決め）
 *    出すのは ★板の 枚数・式の マスの 数・関数の 名前と 回数・時間★ だけ。
 *    ★マスの 値・シート名は 出しません★。
 *
 *  走らせ方:
 *    node docs/measured/hakaru-hozon-no-mado-no-hayasa.mjs
 *    node docs/measured/hakaru-hozon-no-mado-no-hayasa.mjs "C:/(手元の道)/代行計算表2026.xlsb"
 *      ★手元の 道は 引数で 渡します★＝★試験に 焼き込みません★
 */
import path from 'node:path'; import http from 'node:http'; import fs from 'node:fs';
import { fileURLToPath } from 'node:url';
import { borrow, launch } from '../../scripts/_borrow-playwright.mjs';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '../..');
const 材料 = process.argv[2] || path.join(ROOT, 'tests/fixtures/cross-sheet-sample.xlsb');

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
const 大きさ = fs.statSync(材料).size;
console.log('★保存の 窓が 出るまでの 時間★');
console.log('  材料 ....... ' + path.basename(材料) + '（' + 大きさ.toLocaleString() + ' バイト）');
console.log('  ★1バイトも 書きません★（`FileOut.deliver` を 断る物に 替えて [やめる] を 押します）');

const wk = await borrow('hozon-no-mado-no-hayasa', 'webkit');
const browser = await launch('hozon-no-mado-no-hayasa', wk, {}, 'webkit');
const page = await browser.newPage({ viewport: { width: 1280, height: 900 } });
const 配信 = await 立てる(ROOT);

try {
  await page.goto(配信.url + '/book.html', { waitUntil: 'load', timeout: 120000 });
  await page.evaluate(() => {
    document.body.classList.remove('exally-locked');
    const ov = document.getElementById('loginOv');
    if (ov) { ov.classList.remove('open'); ov.style.display = 'none'; }
    /* ★先に 断る物に 替える★＝★測っている 間に 1本も 出さない★ */
    window.FileOut = { deliver: () => Promise.reject(new Error('★書き出しては いけません★')) };
    window.__t0 = performance.now();
  });

  await page.setInputFiles('#bookFileInput', 材料);
  /* ★★2026-09-24 ── ★時間の 指定が 渡って いませんでした★★（経営者1 が 実物で 割りました）
       この 待ちは ★（関数, 引数, 決め）の 3つ★ を 取ります。★2つ目は 「関数に 渡す 引数」★です。
       2つ目の 席に 決めを 置くと
       （★ここに 呼び名を 字で 書きません★＝★`tests/machi-no-kime.test.mjs` が 拾って しまう★）
       ★`{timeout:300000}` は 関数の 引数に なり、決めは 既定（30秒）★に なります。
       ⇒司さんの 実物（404 KB・板 15枚・マクロ入り）は ★30秒では 開き終わらず★
         `Timeout 30000ms exceeded` で 落ちました（★5分 待った つもりが 30秒★）。
       ⇒★`null` を 挟んで 3つ目に 渡します★
       ★証し★ ... 字には 300000 と 書いて あるのに 出しは ★30000ms★ と 言いました
       ★★私(Exally1)が 実物で 割り直しました★★＝★人の 言を 根拠に しない★
         2つ目に 決め ........ ★30,006 ms★（`Timeout 30000ms exceeded`）
         null を 挟んで 3つ目 .. ★ 1,009 ms★（`Timeout 1000ms exceeded`）
       ⇒★★同じ 書き方が 10本・18か所 在りました★★＝全部 直し、
         ★`tests/machi-no-kime.test.mjs`★ で 増えないように しました。 */
  await page.waitForFunction(() => {
    const sh = window.sheets && window.sheets[window.activeSheet];
    return !!(sh && sh.data && Object.keys(sh.data).length > 0);
  }, null, { timeout: 300000 });
  const 開いた = await page.evaluate(() => {
    const t = performance.now() - window.__t0;
    let 式 = 0, マス = 0;
    (window.sheets || []).forEach((sh) => {
      const d = (sh && sh.data) || {};
      Object.keys(d).forEach((k) => {
        マス += 1;
        const c = d[k];
        if (c && typeof c.f === 'string' && c.f.charAt(0) === '=') 式 += 1;
      });
    });
    return { 秒: t, 板: (window.sheets || []).length, マス: マス, 式: 式 };
  });
  console.log('');
  console.log('  ①開く ..... ' + Math.round(開いた.秒) + ' ms'
    + '（板 ' + 開いた.板 + '枚 ／ マス ' + 開いた.マス.toLocaleString()
    + ' ／ ★式の マス ' + 開いた.式.toLocaleString() + '★）');

  /* ② ★中身を 数える だけ★（3回 取って 真ん中） */
  const 数える = await page.evaluate(() => {
    const 出 = [];
    let 数 = null;
    for (let i = 0; i < 3; i++) {
      const a = performance.now();
      数 = window._本の中身を数えた物();
      出.push(performance.now() - a);
    }
    出.sort((x, y) => x - y);
    return { 秒: 出[1], みな: 出, 一文: 数 ? window.HonNoNakami.一文(数) : '(台が 無い)' };
  });
  console.log('  ②数える ... ' + Math.round(数える.秒) + ' ms'
    + '（3回 ' + 数える.みな.map((x) => Math.round(x)).join(' / ') + ' の 真ん中）');
  console.log('      ＝ ' + 数える.一文);

  /* ③ ★保存の 窓が 出るまで★（1マス 打ってから＝お客さんが 保存する 形） */
  await page.evaluate(() => {
    const sh = (window.sheets || [])[window.activeSheet || 0] || {};
    const d = sh.data || {};
    /* ★誰も 使って いない 所を 探して 打つ★（元の 式を 触らない） */
    let r = 200;
    while (d[r + ',0'] && r < 5000) r += 1;
    window.__uchi = r;
    window.setCell(r, 0, '1');
    window.__t1 = performance.now();
    window.__保存 = window.saveOpenedBook();
  });
  /* ★★2026-09-24 ── ★時間の 指定が 渡って いませんでした★★（経営者1 が 実物で 割りました）
       この 待ちは ★（関数, 引数, 決め）の 3つ★ を 取ります。★2つ目は 「関数に 渡す 引数」★です。
       2つ目の 席に 決めを 置くと
       （★ここに 呼び名を 字で 書きません★＝★`tests/machi-no-kime.test.mjs` が 拾って しまう★）
       ★`{timeout:300000}` は 関数の 引数に なり、決めは 既定（30秒）★に なります。
       ⇒司さんの 実物（404 KB・板 15枚・マクロ入り）は ★30秒では 開き終わらず★
         `Timeout 30000ms exceeded` で 落ちました（★5分 待った つもりが 30秒★）。
       ⇒★`null` を 挟んで 3つ目に 渡します★
       ★証し★ ... 字には 300000 と 書いて あるのに 出しは ★30000ms★ と 言いました
       ★★私(Exally1)が 実物で 割り直しました★★＝★人の 言を 根拠に しない★
         2つ目に 決め ........ ★30,006 ms★（`Timeout 30000ms exceeded`）
         null を 挟んで 3つ目 .. ★ 1,009 ms★（`Timeout 1000ms exceeded`）
       ⇒★★同じ 書き方が 10本・18か所 在りました★★＝全部 直し、
         ★`tests/machi-no-kime.test.mjs`★ で 増えないように しました。 */
  await page.waitForFunction(() => {
    const ov = document.getElementById('diffOverlay');
    return !!ov && ov.style.display === 'flex';
  }, null, { timeout: 300000 });
  const 窓 = await page.evaluate(() => ({
    秒: performance.now() - window.__t1,
    打った行: window.__uchi,
    文: (document.getElementById('diffBun') || {}).textContent || '',
    中身の文: (document.getElementById('diffNakami') || {}).textContent || '',
    畳み: [...document.querySelectorAll('#diffBody details.diffKuwashiku')]
      .map((x) => (x.querySelector('summary') || {}).textContent || ''),
  }));
  console.log('  ③窓が出る . ' + Math.round(窓.秒) + ' ms'
    + '（★うち 数えるのに ' + Math.round(数える.秒) + ' ms＝'
    + (窓.秒 > 0 ? Math.round(数える.秒 / 窓.秒 * 100) : 0) + '%★）');
  console.log('');
  console.log('★窓に 出た 字★（★マスの 値と シート名は 出しません★）');
  /* ★★1文は 出しません★★＝番地と 前後の 値が 入る＝★実物の 中身★（司さんの 決め）
       ★出て いるか だけ 数えます★ */
  console.log('  1文 ....... ' + (窓.文.length > 5 ? '★出て います★（' + 窓.文.length + '字・中身は 出しません）' : '★出て いません★'));
  console.log('  中身 ...... ' + 窓.中身の文);
  console.log('  畳み ...... ' + 窓.畳み.join(' / '));

  /* ★[やめる]★＝1本も 書き出さない */
  await page.evaluate(() => { document.getElementById('diffCancel').click(); });
  await page.evaluate(() => window.__保存);
  const 閉じた = await page.evaluate(() =>
    document.getElementById('diffOverlay').style.display !== 'flex');
  console.log('');
  console.log('  ★[やめる] で 閉じた★ ' + 閉じた + ' ／ ★書き出し 0本★');
} finally {
  await page.close().catch(() => {});
  await browser.close().catch(() => {});
  配信.閉じる();
}
