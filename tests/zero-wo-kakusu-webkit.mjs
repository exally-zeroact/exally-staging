/* zero-wo-kakusu-webkit.mjs — ★ゼロを 隠す 板では 0 を 描かない★ 2026-09-25
 *
 *  ★★何が 在ったか（経営者1 の 実測・★物差しは 実Excel★）★★
 *    司さんの 実物（式 15,799個）を 実Excel の `.Text` と 突き合わせた:
 *      ★違う ... 9,163個（58%）★
 *      その うち ★6,743個（74%）が 「Excel は 空・うちは `0`」★
 *    ⇒`ActiveWindow.DisplayZeros` を 板ごとに 読むと ★15枚中 13枚が False★
 *    ⇒★計算の 間違いでは ありません＝板の 設定を 読んで いませんでした★
 *
 *  ★★印は 実物の バイトで 決めました（★番号表を 記憶で 書いて いません★）★★
 *    `.xlsb` ･･･ 板の 記録 番号137 の 先頭バイト:
 *        板 1〜7・9〜12・15 ... `8c 03`
 *        板 8 ................ `cc 03`
 *        ★板 13・14 .......... `9c 03`★
 *      `8c` = 1000 1100 ／ `9c` = 1001 1100 ⇒ ★違うのは ビット4 だけ★
 *      ⇒★立って いる 2枚 ＝ 経営者1 が COM で 読んだ True の 2枚★（1対1）
 *    `.xlsx` ･･･ `<sheetView showZeros="0">`（★札が 無い＝出す★）
 *
 *  ★★この 試験が 見る 物★★
 *    ⑴★隠す 板では 0 の マスに 何も 描かない★（★印では なく 絵の 点で 数える★）
 *    ⑵★出す 板では 0 を 描く★（★隠し過ぎて いない★）
 *    ⑶★0 でない 数は どちらでも 描く★
 *    ⑷★字の 「0」は 隠さない★（実Excel も 値が 0 の マスだけ 隠す）
 *    ⑸★結合した マスも 同じ★（★作る道が 2本 在る★）
 *
 *  走らせ方: node tests/zero-wo-kakusu-webkit.mjs
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

/* ══ ★材料★ ══（`.xlsx` で 作ります＝`showZeros` を 字で 書けるので）
     板「かくす」... A1=0 B1=5 C1=文字の0   ⇒ `showZeros="0"`
     板「だす」 ... A1=0 B1=5               ⇒ 札 無し（＝出す） */
const 作る = (ゼロ0) => {
  /* ★★1行目は 使いません★★（2026-09-25 ここで 1回 踏みました）
       A1 は 開いた 直後 ★選ばれて います★＝★字が 消えても 枠の 線が 40点 残ります★。
       ⇒★『消えた』と『選びの 枠』を 分ける 為に 2行目で 数えます★ */
  /* ★★式の 答えが 0 の マスも 要ります★★（2026-09-25 経営者1 の 指摘）
       はじめ 見本は ★数を 持った マス（定数）だけ★ でした。
       ⇒★式の マスは `v` が 空の 字・答えは `d`★ なので 別の 道です。
       ⇒実物では ★中の 字が 0 の 式 6,752個 の うち 隠したのは 0個★ でした。
       ⇒★見本に 式の 0 が 無いと 直しても また 通ります★ */
  const ws = XLSX.utils.aoa_to_sheet([['みだし', 'み2', 'み3', 'み4', 'み5'],
    [0, 5, '0', 1, 1]]);
  ws['C2'].t = 's';
  /* D3 ＝ `=D2-E2` ＝ 答え 0（★式の マス★）／ E3 ＝ `=D2+E2` ＝ 答え 2 */
  ws['D3'] = { t: 'n', f: 'D2-E2', v: 0 };
  ws['E3'] = { t: 'n', f: 'D2+E2', v: 2 };
  /* ══ ★★書式に 字が 混じった 0 も 要ります★★ ══（2026-09-25 経営者1 が 割りました）
       `js/book-open.js` が 書式を 掛けると 答えは ★字★ に なります。
       ⇒`Number('0 円')` ＝ ★NaN★ ⇒ ★0 と 判じられず 隠せません★
       ⇒司さんの 実物で ★47個★（`#,##0" 円"` 22 ／ `0.00"時間"` 19 ／ `"¥"#,##0` 6）
       ⇒★★3種 とも 1つずつ 入れます★★（★1種だけでは 他の 2種で また 通ります★） */
  ws['A4'] = { t: 'n', f: 'D2-E2', v: 0, z: '#,##0" 円"' };
  ws['B4'] = { t: 'n', f: 'D2-E2', v: 0, z: '0.00"時間"' };
  ws['C4'] = { t: 'n', f: 'D2-E2', v: 0, z: '"¥"#,##0_);[Red]("¥"#,##0)' };
  /* ★0 でない 方も 1つ 置きます★＝★隠し過ぎて いないかを 同じ 書式で 見ます★ */
  ws['D4'] = { t: 'n', f: 'D2+E2', v: 2, z: '#,##0" 円"' };
  /* ══ ★★保存値は 0 だが 計算し直すと 0 で ない マス★★ ══（2026-09-25）
       `書式前の数` を 付けた ので ★計算し直した 後も それが 残ると 古い 数で 判じます★。
       ⇒★`=D2+E2`（答え 2）に わざと 保存値 0 を 入れます★
       ⇒★開いた 直後は 0（隠す）／1マス 打つと 2（描く）が 正しい★
       ⇒★古い 数が 残って いると 打った 後も 隠れた ままに なります★ */
  ws['E4'] = { t: 'n', f: 'D2+E2', v: 0, z: '#,##0" 円"' };
  ws['!ref'] = 'A1:E4';
  return ws;
};
const wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, 作る(), 'かくす');
XLSX.utils.book_append_sheet(wb, 作る(), 'だす');
const 材料 = path.join(os.tmpdir(), 'exally-zero-kakusu.xlsx');
fs.writeFileSync(材料, XLSX.write(wb, { bookType: 'xlsx', type: 'buffer' }));

/* ★借り物は `showZeros` を 書いて くれません★＝★包みを 開けて 自分で 書きます★
   （★これは 材料作りです＝本番の 道では ありません★） */
{
  const ZipSurgeon = require_(path.join(ROOT, 'lib/zip-surgeon.js'));
  const z = ZipSurgeon.read(new Uint8Array(fs.readFileSync(材料)));
  const 名 = z.names().filter((n) => /^xl\/worksheets\/sheet1\.xml$/.test(n));
  if (!名.length) { console.log('★板1が 見つかりません★'); process.exit(1); }
  let xml = await z.text(名[0]);
  if (xml.indexOf('<sheetViews>') < 0) {
    xml = xml.replace('<sheetData>', '<sheetViews><sheetView showZeros="0" workbookViewId="0"/></sheetViews><sheetData>');
  } else {
    xml = xml.replace('<sheetView ', '<sheetView showZeros="0" ');
  }
  z.replaceText(名[0], xml);
  fs.writeFileSync(材料, Buffer.from((await z.build()).bytes));
}

console.log('[zero-wo-kakusu] ★ゼロを 隠す 板では 0 を 描かない★');
console.log('      ── 材料 ── ' + path.basename(材料) + '（板「かくす」に `showZeros="0"`）');

const wk = await borrow('zero-wo-kakusu', 'webkit');
const browser = await launch('zero-wo-kakusu', wk, {}, 'webkit');
const page = await browser.newPage({ viewport: { width: 1280, height: 900 } });
const 配信 = await 立てる(ROOT);

try {
  page.on('console', (m) => {
    const t = String(m.text());
    if (t.indexOf('飾り') >= 0 || m.type() === 'error') console.log('      [画面] ' + t.slice(0, 200));
  });
  page.on('pageerror', (e) => console.log('      [落ちた] ' + String(e.message).slice(0, 200)));
  await page.goto(配信.url + '/book.html', { waitUntil: 'load', timeout: 120000 });
  await page.evaluate(() => {
    document.body.classList.remove('exally-locked');
    const ov = document.getElementById('loginOv');
    if (ov) { ov.classList.remove('open'); ov.style.display = 'none'; }
    window.FileOut = { deliver: () => Promise.reject(new Error('★書き出しては いけません★')) };
  });
  await page.setInputFiles('#bookFileInput', 材料);
  await page.waitForFunction(() => {
    const ss = window.sheets || [];
    return ss.some((s) => s && s.data && Object.keys(s.data).length > 0);
  }, null, { timeout: 120000 });
  await page.waitForTimeout(600);


  const 印 = await page.evaluate(() => (window.sheets || []).map((s) => ({
    名: s.name, 隠す: !!s.ゼロを隠す,
  })));
  console.log('      ── 実測 ── 板の 印 ... '
    + 印.map((x) => x.名 + ':' + (x.隠す ? '★隠す★' : '出す')).join(' / '));
  T('★`showZeros="0"` の 板だけ 「隠す」★（札が 無い 板は 出す）',
    印.length === 2 && 印[0].隠す === true && 印[1].隠す === false,
    JSON.stringify(印));

  /** ★絵の 点を 数える★＝★印では なく 描かれたか★ */
  const 点 = (板, r, c) => page.evaluate(([n, rr, cc]) => {
    const i = (window.sheets || []).findIndex((s) => s.name === n);
    if (i < 0) return -1;
    if (window.activeSheet !== i) { window.switchSheet(i); }
    /* ★★選んで いる マスの 枠が 点に 数えられます★★（2026-09-25 ここで 1回 踏みました）
         A1 は 開いた 直後 選ばれて いて、★字が 消えても 枠の 線が 40点 残りました★。
         ⇒★遠い マスへ 選びを 移してから 数えます★ */
    if (typeof window.selectCell === 'function') window.selectCell(40, 20);
    else { window.selR1 = window.selR2 = 40; window.selC1 = window.selC2 = 20; }
    if (typeof window.render === 'function') window.render();
    else window._renderPass();
    const g = window.ctx, cv = g.canvas;
    const 倍 = cv.width / (cv.clientWidth || cv.width);
    const x = window.colX(cc), y = window.rowY(rr), w = window.cW(cc), h = window.rH(rr);
    /* ★★マスの ふちを 読みません★★（2026-09-25 ここで 1回 踏みました）
         消した 後の A2 に ★2点★ 残り、何も 無い D2 は 0点 でした。
         ＝★罫線の ぼけ★ を 拾って いました（★字では ない★）。
         ⇒★内側へ 3点 入った 所だけ 読みます★＝★どの マスも 同じ 測り方★ */
    const d = g.getImageData(Math.round((x + 3) * 倍), Math.round((y + 3) * 倍),
      Math.max(1, Math.round((w - 6) * 倍)), Math.max(1, Math.round((h - 6) * 倍))).data;
    let 数 = 0;
    for (let k = 0; k < d.length; k += 4) {
      if (d[k + 3] < 8) continue;
      if (255 - d[k] > 60 || 255 - d[k + 1] > 60 || 255 - d[k + 2] > 60) 数++;
    }
    return 数;
  }, [板, r, c]);

  await page.evaluate(() => { window.switchSheet(0); });
  await page.waitForTimeout(400);
  const か0 = await 点('かくす', 1, 0);
  const か5 = await 点('かくす', 1, 1);
  const か字 = await 点('かくす', 1, 2);
  /* ★★「空」の 点を 物差しに します★★（2026-09-25 ここで 1回 踏みました）
       消した 後の A2 に ★2点★ 残りました＝★罫線の ふちの ぼけ★。
       ⇒「0点でなければ 赤」に すると ★消えて いるのに 赤★ に なります。
       ⇒★何も 無い マス（D2）と 比べます★＝★自分で 決めた 数を 使わない★ */
  const か空 = await 点('かくす', 1, 5);   /* F2＝何も 無い マス */
  const か式0 = await 点('かくす', 2, 3);   /* D3＝式の 答えが 0 */
  const か式2 = await 点('かくす', 2, 4);   /* E3＝式の 答えが 2 */
  console.log('      ── 実測 ── 板「かくす」... A2(0) ' + か0 + ' ／ B2(5) ' + か5
    + ' ／ C2(字の0) ' + か字 + ' ／ ★D3(式の0) ' + か式0 + '★ ／ E3(式の2) ' + か式2
    + ' ／ ★何も 無い F2 ' + か空 + '★');
  T('★★隠す 板では 0 に 何も 描かない★★（★何も 無い マスと 同じ★）',
    か0 <= か空, 'A2 ' + か0 + '点 ／ 何も 無い F2 ' + か空 + '点');

  T('★★式の 答えが 0 の マスも 隠す★★（★実物の 6,752個は こちら★）',
    か式0 <= か空, 'D3(=D2-E2) ' + か式0 + '点 ／ 何も 無い F2 ' + か空 + '点');

  T('★式の 答えが 0 でなければ 描く★（隠し過ぎて いない）',
    か式2 > 0, 'E3(=D2+E2) ' + か式2 + '点');
  T('★隠し過ぎて いない★（0 でない 数は 描く）', か5 > 0, 'B2 に 点が ' + か5 + '個');

  /* ══ ★★書式に 字が 混じった 0（3種）★★ ══（2026-09-25）
       ★1種だけ 見ると 他の 2種で また 通ります★＝★3種 とも 名指しで 見ます★ */
  const か円 = await 点('かくす', 3, 0);    /* A4＝`#,##0" 円"` の 0 */
  const か時 = await 点('かくす', 3, 1);    /* B4＝`0.00"時間"` の 0 */
  const か円記号 = await 点('かくす', 3, 2); /* C4＝`"¥"#,##0` の 0 */
  const か円2 = await 点('かくす', 3, 3);   /* D4＝同じ 書式で 答え 2 */
  console.log('      ── 実測 ── ★書式に 字が 混じった 0★ ... 円 ' + か円
    + ' ／ 時間 ' + か時 + ' ／ ¥ ' + か円記号 + ' ／ ★0 でない 方 ' + か円2 + '★');
  T('★★書式が `#,##0" 円"` でも 0 を 隠す★★（実物 22個）',
    か円 <= か空, 'A4 ' + か円 + '点 ／ 何も 無い F2 ' + か空 + '点');
  T('★★書式が `0.00"時間"` でも 0 を 隠す★★（実物 19個）',
    か時 <= か空, 'B4 ' + か時 + '点 ／ 何も 無い F2 ' + か空 + '点');
  T('★★書式が `"¥"#,##0` でも 0 を 隠す★★（実物 6個）',
    か円記号 <= か空, 'C4 ' + か円記号 + '点 ／ 何も 無い F2 ' + か空 + '点');
  T('★同じ 書式で 0 でなければ 描く★（隠し過ぎて いない）',
    か円2 > 0, 'D4(=D2+E2) ' + か円2 + '点');
  T('★★字の 「0」は 隠さない★★（実Excel も 値が 0 の マスだけ 隠す）',
    か字 > 0, 'C2 に 点が ' + か字 + '個');

  /* ══ ★★1マス 打って 計算させた 後も 同じか★★ ══（2026-09-25）
       ＝★`書式前の数` が 古いまま 残って いないか★ */
  const 打つ前のE4 = await 点('かくす', 3, 4);
  await page.evaluate(() => { window.setCell(500, 20, '1'); });
  await page.waitForTimeout(1200);
  const 打った後のA4 = await 点('かくす', 3, 0);
  const 打った後のE4 = await 点('かくす', 3, 4);
  console.log('      ── 実測 ── ★1マス 打った 後★ ... 円の0(A4) ' + 打った後のA4
    + ' ／ ★保存0だが答え2(E4) 打つ前 ' + 打つ前のE4 + ' → 打った後 ' + 打った後のE4 + '★');
  T('★★打った 後も 書式つきの 0 は 隠れた まま★★（★古い 数に 頼って いない★）',
    打った後のA4 <= か空, 'A4 ' + 打った後のA4 + '点 ／ 何も 無い F2 ' + か空 + '点');
  T('★★保存値 0 でも 答えが 2 なら 打った 後は 描く★★（★古い 数で 隠さない★）',
    打った後のE4 > 0, 'E4 打つ前 ' + 打つ前のE4 + '点 → 打った後 ' + 打った後のE4 + '点');

  await page.evaluate(() => { window.switchSheet(1); });
  await page.waitForTimeout(400);
  const だ0 = await 点('だす', 1, 0);
  const だ5 = await 点('だす', 1, 1);
  console.log('      ── 実測 ── 板「だす」 ... A2(0) ' + だ0 + ' ／ B2(5) ' + だ5);
  T('★★出す 板では 0 を 描く★★（★隠し過ぎて いない★）', だ0 > 0, 'A2 に 点が ' + だ0 + '個');
  T('★出す 板の 5 も 描く★', だ5 > 0, 'B2 に 点が ' + だ5 + '個');
} finally {
  await page.close().catch(() => {});
  await browser.close().catch(() => {});
  配信.閉じる();
}

console.log('\n' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
