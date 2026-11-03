/* hakaru-idageta.mjs — ★`####` に なる マスを ★台の 口で★ 数える★ 2026-09-25
 *
 *  ★★なぜ 作り直すか★★
 *    `docs/measured/hakaru-mikire.mjs` は ★自分で 作った canvas★ で 測って いました。
 *    ＝★`cell.d` を そのまま 測る★＝★画面が 出す 字とは 別物★ の 事が 在ります。
 *    ⇒★それでは 「450マス」の 中身を 割れません★。
 *    ⇒★画面が 使って いる 口を そのまま 呼びます★
 *        `_字の元` / `_答えは字か` / `_入る字数` / `fmtForDisplay`
 *        `_マスの字大` / `_マスの書体` / `_数が入らないか` / `cW`
 *
 *  ★★経営者1 が 実Excel で 出した 数（2026-09-25・`8964ad3`）★★
 *    給料1/2/3 の 字の 在る マス ★2,039個★ ／ ★実Excel の `####` は 0個★
 *    ⇒★うちの `####` は 本物の 欠陥★（司さんの 言葉の 通り）
 *
 *  ★★経営者1 の 見立て（★私は まだ 確かめて いません★）★★
 *    「`measureText` の 前に `ctx.font` を その マスの 字に して いますか」
 *    ⇒★字を 読んだ 限りでは ★して います★★（`drawText` も 結合の 方も 先に `ctx.font`）
 *    ⇒だから ★別の 所★ だと 思って います。★それを この 道具で 割ります★。
 *
 *  ★★実物の 中身は 出しません★★（司さんの 決め）
 *    出すのは ★板の 名・列の 番号・幅・字の 大きさ・書式の 形・字数・点★ だけ。
 *    ★マスの 値そのものは 出しません★。
 *
 *  ★読むだけ★＝保存の 窓は 出しません。1バイトも 書きません。
 *
 *  走らせ方: node docs/measured/hakaru-idageta.mjs [材料の道]
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
console.log('★`####` に なる マスを 台の 口で 数える★');
console.log('  材料 ... ' + path.basename(材料) + '（' + fs.statSync(材料).size.toLocaleString() + ' バイト）');

const wk = await borrow('idageta', 'webkit');
const browser = await launch('idageta', wk, {}, 'webkit');
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
  /* ★★開かない 時に 「何も 分からない」で 終わらせない★★（2026-09-25）
       実物は ★開くのに 2分★（経営者1 の 実測）。
       私の 側では ★600秒・900秒で 落ちた 事が 4回★ 在ります（★因は 未測定★）。
       ⇒★待って いる 間の 中を 出します★＝★落ちた 時に 手掛かりが 残る★ */
  page.on('console', (m) => {
    const t = m.type();
    if (t === 'error' || t === 'warning') console.log('    [画面の 声/' + t + '] ' + String(m.text()).slice(0, 160));
  });
  page.on('pageerror', (e) => console.log('    [★画面が 落ちた★] ' + String(e.message).slice(0, 200)));
  /* ★★取れなかった 物の 名を 出す★★（2026-09-25）
       1回目の 見張りで 「Failed to load resource」だけ 出て ★何が 取れなかったか 分からず★。
       ⇒★URL と 番号を 出します★＝★読み込みが 止まる 因は だいたい ここ★ */
  page.on('requestfailed', (q) => console.log('    [★取れなかった★] ' + q.url().slice(0, 140)
    + ' ･･･ ' + ((q.failure() || {}).errorText || '(訳 不明)')));
  page.on('response', (r) => {
    if (r.status() >= 400) console.log('    [★番号 ' + r.status() + '★] ' + r.url().slice(0, 140));
  });
  const 始 = Date.now();
  const 見張り = setInterval(async () => {
    try {
      const n = await page.evaluate(() => {
        const ss = window.sheets || [];
        let 板 = ss.length, マス = 0;
        for (let i = 0; i < ss.length; i++) マス += Object.keys((ss[i] || {}).data || {}).length;
        /* ★★どこまで 進んだかを 段で 出す★★（2026-09-25）
             ①台を 読む（`_ensureXlsx`）･･･ `XLSX` `XlsxIO` `BookOpen` が 揃ったか
             ②本を 読む（`BookOpen.openFile`）･･･ `isOpened()`
             ③知らせ ･･･ ★「開けませんでした」が 出て いれば それは 止まりでは なく 断り★
           ⇒★止まって いる 段が 分かれば 直す 所が 決まります★ */
        const 台 = ['XLSX', 'XlsxIO', 'ZipSurgeon', 'XlsxEdit', 'XlsbEdit', 'XlsbJitai',
          'TableRefs', 'XlsxKazari', 'XlsxZukei', 'DiffPreview', 'BookOpen']
          .filter((k) => !window[k]);
        let 知らせ = '';
        document.querySelectorAll('div,span').forEach((el) => {
          const t = (el.textContent || '');
          if (t.indexOf('開けませんでした') >= 0 && t.length < 200) 知らせ = t.slice(0, 120);
        });
        return 板 + '枚 / ' + マス + 'マス'
          + ' ／ 揃って いない 台 ' + (台.length ? 台.join(',') : '0本')
          + ' ／ 開いた ' + (window.BookOpen && window.BookOpen.isOpened ? window.BookOpen.isOpened() : '(口なし)')
          + (知らせ ? ' ／ ★知らせ「' + 知らせ + '」★' : '');
      });
      console.log('    [' + Math.round((Date.now() - 始) / 1000) + '秒] ' + n);
    } catch (e) { /* 読めない 時は 黙る（落とさない） */ }
  }, 20000);
  await page.setInputFiles('#bookFileInput', 材料);
  /* ★★「開けませんでした」が 出たら すぐ 止めます★★（2026-09-25 ここで 1回 踏みました）
       私は 880秒 待ち続けて 「★アプリが 止まって いる★」と 報告する 所でした。
       実際は ★20秒の 時点で 画面は もう 断って いました★
         「開けませんでした：The object can not be found here.」
       ⇒★★止まって いたのは アプリでは なく 私の 測り道具★★。
       ⇒★『何も 起きない』と 『断られた のに 見て いない』は 別★。 */
  await page.waitForFunction(() => {
    const ss = window.sheets || [];
    for (let i = 0; i < ss.length; i++) {
      if (ss[i] && ss[i].data && Object.keys(ss[i].data).length > 0) return true;
    }
    let 断り = '';
    document.querySelectorAll('div,span').forEach((el) => {
      const t = (el.textContent || '');
      if (t.indexOf('開けませんでした') >= 0 && t.length < 200) 断り = t.slice(0, 160);
    });
    if (断り) throw new Error(断り);
    return false;
  }, null, { timeout: 900000 });
  clearInterval(見張り);
  await page.waitForTimeout(800);

  const 板の数 = await page.evaluate(() => (window.sheets || []).length);
  const 全 = { 見た: 0, 井桁: 0 };
  const 見本 = [];
  const 板ごと = [];

  for (let i = 0; i < 板の数; i++) {
    /* ★板を 切り替えます★＝`_マスの字大` など は `activeSheet` を 見る */
    await page.evaluate((n) => { window.switchSheet(n); }, i);
    await page.waitForTimeout(120);
    const 出 = await page.evaluate(() => {
      const sh = window.sheets[window.activeSheet];
      const d = sh.data || {};
      const g = window.ctx;
      let 見た = 0, 井桁 = 0;
      const 本 = [];
      for (const k in d) {
        const cell = d[k];
        if (!cell || cell.merged) continue;
        const raw = window._字の元(cell);
        if (!raw && raw !== 0) continue;
        見た++;
        const p = k.split(',');
        const c = +p[1];
        /* ★★結合した マスは 「くっついた 幅」で 測ります★★（2026-09-25 ここで 1回 踏みました）
             はじめは `cW(c)` だけを 見て いました ⇒ ★結合の 元が 21マス 出ました★。
             ⇒画面の 結合の 方（`_renderPass`）は `for(cc=c..c2) w+=cW(cc)` で 足して います。
             ⇒★1列ぶんで 測ると 入らないに 決まって います★＝★道具が 相手を 不利に して いた★ */
        let w = window.cW(c);
        if (cell.mergeEnd) {
          w = 0;
          for (let cc = c; cc <= cell.mergeEnd.c; cc++) w += window.cW(cc);
        }
        if (!(w > 0)) continue;
        /* ★★画面と 同じ 順で 同じ 口を 呼びます★★
             ①字体を 決める ②字を 作る ③入るか 判じる */
        const 大 = window._マスの字大(cell);
        const 書 = window._マスの書体(cell);
        g.font = (cell.italic ? 'italic ' : 'normal ') + (cell.bold ? 'bold ' : 'normal ')
          + 大 + 'px ' + 書;
        const 字か = window._答えは字か(cell);
        const display = 字か ? String(raw)
          : window.fmtForDisplay(raw, cell.numFmt, window._入る字数(w, raw, cell.numFmt));
        if (字か) continue;                       /* ★字は `####` に しない★ */
        if (!window._数が入らないか(raw, display, w)) continue;
        井桁++;
        if (本.length < 4) {
          本.push({
            板: sh.name, 行: +p[0], 列: c, 幅: Math.round(w * 100) / 100,
            結合: cell.mergeEnd ? (cell.mergeEnd.c - c + 1) + '列' : '-',
            字大: 大, 書体: String(書).slice(0, 24),
            書式: cell.numFmt || '(無し)',
            字数: String(display).length,
            点: Math.round(g.measureText(String(display)).width * 100) / 100,
            余白: window.マスの余白,
            /* ★値は 出しません★＝★形だけ★（数字を 9 に 置き換える） */
            形: String(display).replace(/[0-9]/g, '9'),
            一字: Math.round(g.measureText('9').width * 100) / 100,
          });
        }
      }
      return { 名: sh.name, 見た, 井桁, 本 };
    });
    全.見た += 出.見た; 全.井桁 += 出.井桁;
    if (出.井桁) { 板ごと.push(出); 見本.push(...出.本); }
  }

  console.log('');
  console.log('★数★ 板 ' + 板の数 + '枚 ／ 見た マス ' + 全.見た.toLocaleString()
    + ' ／ ★`####` に なる ' + 全.井桁.toLocaleString() + ' マス★');
  for (const b of 板ごと) console.log('  ' + String(b.名).slice(0, 12).padEnd(13)
    + ' 見た ' + String(b.見た).padStart(6) + ' ／ #### ' + String(b.井桁).padStart(5));
  console.log('');
  console.log('★見本（★値は 出しません・形だけ★）★');
  for (const x of 見本.slice(0, 12)) {
    console.log('  ' + String(x.板).slice(0, 8).padEnd(9)
      + ' 行' + String(x.行).padStart(4) + ' 列' + String(x.列).padStart(3)
      + ' ／ 幅 ' + String(x.幅).padStart(6) + '(' + x.結合 + ')' + ' 余白 ' + x.余白
      + ' ／ 字 ' + String(x.字大).padStart(3) + 'px'
      + ' ／ 一字 ' + String(x.一字).padStart(5) + 'px'
      + ' ／ 字数 ' + String(x.字数).padStart(3)
      + ' ／ 点 ' + String(x.点).padStart(7)
      + ' ／ 書式 ' + String(x.書式).slice(0, 14).padEnd(15)
      + ' ／ 形 ' + x.形);
  }
  console.log('');
  console.log('★見方★ ... ★点 > 幅 − 余白★ なら `####` に なります。');
  console.log('  ⇒★点が 幅を どれだけ 超えて いるか★／★一字あたり 何px か★ を 見ます。');
} finally {
  await page.close().catch(() => {});
  await browser.close().catch(() => {});
  配信.閉じる();
}
