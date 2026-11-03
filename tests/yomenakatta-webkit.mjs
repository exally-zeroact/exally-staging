/* yomenakatta-webkit.mjs — ★ファイルを 読めなかった 時に 読み直すか★ 2026-09-25
 *
 *  ★★何が 在ったか（実測・お客さんの 道・実物）★★
 *    同じ 本の 2つの 写しで 数えました:
 *      `OneDrive` の 下 ... ★24回中 15回 断られる★（私 10中6・経営者1 9/14）
 *      `OneDrive` の 外 ... ★13回中 0回★（私 3中0・経営者1 0/7）
 *    断りの 字 ... `The object can not be found here.`（webkit）
 *                  `The I/O read operation failed.`（chromium）
 *    ＝★選んだ 後に ファイルの 実体が 変わった★ 時に ブラウザが 出す 物。
 *    ⇒★お客さんは OneDrive の ファイルを 開きます★＝★こちらで 受け止めます★
 *
 *  ★★この 試験が 見る 物★★
 *    ⑴★断られたら 読み直すか★（合わせて 3回まで）
 *    ⑵★読み直して 開けたら 普通に 開くか★（＝1回目の 断りを 引きずらない）
 *    ⑶★3回とも 駄目なら 日本語で 言うか★（★英語の 断りを 見せない★）
 *    ⑷★因を 書いて いないか★（「同期中」など ★測って いない 事★ を 言わない）
 *
 *  ★★正直に 書きます ── 断りは ★作り物★ です★★
 *    実物の 断りは ★2回に 1回しか 出ません★（上の 数）。
 *    ⇒★試験に 使えません★（★揺れる 見張りは 後ろの 段を 人質に 取る★）
 *    ⇒だから ★`BookOpen.openFile` を 包んで わざと 断らせます★。
 *    ⇒★断り その物は 実物で 見ました★。★受け止め方だけを ここで 見ます★。
 *
 *  走らせ方: node tests/yomenakatta-webkit.mjs
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

const ws = XLSX.utils.aoa_to_sheet([['あ', 1], ['い', 2]]);
const wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, ws, '板');
const 材料 = path.join(os.tmpdir(), 'exally-yomenakatta.xlsx');
fs.writeFileSync(材料, XLSX.write(wb, { bookType: 'xlsx', type: 'buffer' }));

console.log('[yomenakatta] ★ファイルを 読めなかった 時に 読み直すか★');

const wk = await borrow('yomenakatta', 'webkit');
const browser = await launch('yomenakatta', wk, {}, 'webkit');
const page = await browser.newPage({ viewport: { width: 1280, height: 900 } });
const 配信 = await 立てる(ROOT);

/** ★断る 回数を 決めて 包む★＝`BookOpen.openFile` の 前に 1枚 かぶせる */
async function 包む(何回断るか) {
  await page.evaluate((n) => {
    if (!window.__元のopenFile) window.__元のopenFile = window.BookOpen.openFile;
    window.__断った = 0;
    window.BookOpen.openFile = function (f) {
      if (window.__断った < n) {
        window.__断った++;
        /* ★実物と 同じ 言葉で 断ります★（webkit が 出した 字） */
        const e = new Error('The object can not be found here.');
        e.name = 'NotFoundError';
        return Promise.reject(e);
      }
      return window.__元のopenFile.call(window.BookOpen, f);
    };
  }, 何回断るか);
}
/* ★★知らせは `#toast` だけを 読みます★★（2026-09-25 ここで 1回 踏みました）
     はじめは 画面じゅうの `div,span` から 「読めませんでした」を 探しました
     ⇒★余白の 説明文（実Excel の 数が 外から 読めませんでした）を 拾いました★
     ⇒★『その字が 在る』と『その知らせが 出た』は 別★ */
const 知らせを読む = () => page.evaluate(() => {
  const el = document.getElementById('toast');
  return el ? String(el.textContent || '').slice(0, 200) : '(toast が 無い)';
});

try {
  await page.goto(配信.url + '/book.html', { waitUntil: 'load', timeout: 120000 });
  await page.evaluate(() => {
    document.body.classList.remove('exally-locked');
    const ov = document.getElementById('loginOv');
    if (ov) { ov.classList.remove('open'); ov.style.display = 'none'; }
    window.FileOut = { deliver: () => Promise.reject(new Error('★書き出しては いけません★')) };
  });
  /* ★台を 先に 読ませます★（`BookOpen` が 無いと 包めない） */
  await page.setInputFiles('#bookFileInput', 材料);
  await page.waitForFunction(() => !!window.BookOpen, null, { timeout: 120000 });
  await page.waitForFunction(() => {
    const ss = window.sheets || [];
    return ss.some((s) => s && s.data && Object.keys(s.data).length > 0);
  }, null, { timeout: 120000 });

  /* ══ ⑴⑵ 2回 断って 3回目で 開く ══ */
  await 包む(2);
  await page.evaluate(() => { window.sheets = [{ name: '空', data: {}, colW: {} }]; });
  await page.setInputFiles('#bookFileInput', 材料);
  await page.waitForFunction(() => {
    const ss = window.sheets || [];
    return ss.some((s) => s && s.data && Object.keys(s.data).length > 0);
  }, null, { timeout: 120000 }).catch(() => {});
  const 出1 = await page.evaluate(() => ({
    断った: window.__断った, 読み直した: window._読み直した回数,
    開いた: (window.sheets || []).some((s) => s && s.data && Object.keys(s.data).length > 0),
  }));
  console.log('      ── 実測 ── 断った ' + 出1.断った + '回 ／ 読み直した ' + 出1.読み直した
    + '回 ／ 開いた ' + 出1.開いた);
  T('★★2回 断られても 3回目で 開く★★（★読み直して いる★）',
    出1.開いた === true && 出1.読み直した === 2,
    '開いた ' + 出1.開いた + ' ／ 読み直した ' + 出1.読み直した + '回');

  T('★読み直しは 断られた 数だけ★（黙って 何回も 叩かない）',
    出1.読み直した === 出1.断った, '断った ' + 出1.断った + ' ／ 読み直した ' + 出1.読み直した);

  /* ══ ⑶⑷ 3回とも 断られる ══ */
  await 包む(99);
  await page.setInputFiles('#bookFileInput', 材料);
  await page.waitForFunction(() => window.__断った >= 3, null, { timeout: 120000 }).catch(() => {});
  await page.waitForTimeout(500);
  const 出2 = await page.evaluate(() => ({ 断った: window.__断った, 読み直した: window._読み直した回数 }));
  const 知らせ = await 知らせを読む();
  console.log('      ── 実測 ── 断った ' + 出2.断った + '回 ／ 読み直した ' + 出2.読み直した
    + '回 ／ 知らせ「' + 知らせ + '」');

  T('★3回で 止める★（際限なく 叩かない）', 出2.断った === 3, '断った ' + 出2.断った + '回');

  T('★★英語の 断りを お客さんに 見せない★★',
    知らせ.indexOf('The object') < 0 && 知らせ.indexOf('can not be found') < 0, 知らせ);

  T('★日本語で 「もう一度 押して ください」と 言う★',
    /読めませんでした/.test(知らせ) && /もう一度/.test(知らせ), 知らせ);

  T('★★測って いない 因を 書かない★★（「同期」「OneDrive」を 出さない）',
    知らせ.indexOf('同期') < 0 && 知らせ.indexOf('OneDrive') < 0, 知らせ);

  /* ══ ⑸ 台の 覚えが 断られた まま 残らない ══ */
  /* ★★覚えは 「本物の 断り」で 作ります★★（2026-09-25 ここでも 1回 踏みました）
       はじめは `_xlsxLoading` に ★外から 断られた 約束を 差し込みました★
       ⇒`_ensureXlsx` は ★覚えが 在れば それを 返すだけ★ なので
         ★捨てる 所を 1度も 通りません★＝★試験の 立て方の 誤り★
       ⇒★本当に 台を 読みに 行かせて、その 読み込みを 失敗させます★ */
  const 覚え = await page.evaluate(() => typeof window._ensureXlsx === 'function');
  T('★`_ensureXlsx` が 在る★', 覚え === true);
  const 捨てた = await page.evaluate(async () => {
    window._xlsxLoading = null;
    window.XLSX = undefined; window.XlsxIO = undefined;   /* 覚えの 近道を 塞ぐ */
    /* ★★取れない 道に します★★（2026-09-25 ここで もう 1回 踏みました）
         はじめは `_assetVer` を 変えて `?v=こわす` に しました
         ⇒★配る 側が `?` から 後ろを 捨てて いた★＝★普通に 取れて しまった★
         ⇒★読み込む 所その物を 断らせます★ */
    window._loadScript = function (src) {
      return Promise.reject(new Error(src + ' を読み込めません'));
    };
    try { await window._ensureXlsx(); } catch (e) { /* 断られる */ }
    await new Promise((r) => setTimeout(r, 300));
    return window._xlsxLoading === null || window._xlsxLoading === undefined;
  });
  T('★★断られた 覚えを 捨てる★★（★二度と 開けなく ならない★）',
    捨てた === true, '覚えが 断られた まま 残って います');
} finally {
  await page.close().catch(() => {});
  await browser.close().catch(() => {});
  配信.閉じる();
}

console.log('\n' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
