/* dasu-kobore-fusagi.mjs -- ★塞がれた 溢れ（#SPILL!）を そのまま 書き出す★（㊺）（2026-09-20）
 *
 *  ★★なぜ 要るか★★
 *    ㊹で 測れたのは ★ちゃんと 溢れた 物★だけです。
 *    ⇒★画面で 既に `#SPILL!` に なって いる 物を 書き出したら どう なるか★は ★1本も 測って いません★
 *    ⇒★ここが 一番 危ない★と 私（経営者1）が 見ました:
 *        ・`ref=` は 何に なるのか（★溢れて いないので 範囲が 無い★）
 *        ・`cm=` は 付くのか
 *        ・実Excel は 開いた 時に 何を 出すのか（★#SPILL! が 保たれるか／別の 誤りに なるか★）
 *
 *  ★★この 道具は 「出す」だけです★★
 *    ＝★実Excel で 開くのは 別の 道具★（`toru-kobore-fusagi-wo-excel-ni-hirakaseru.ps1`）
 *    ＝★1つの 道具に 2つの 仕事を させません★（門が 緩みます）
 *
 *  ★★お客さんの 道で 作ります★★
 *    `book.html:saveXlsx`（20026-20034行）と ★同じ 2行★を 呼びます:
 *        `_ensureXlsx()` ⇒ `XlsxIO.writeBook(GridXlsx.gridToBook(sheets))`
 *    ★打つのは `window.setCell`★＝★打鍵では ありません★（打鍵は `utte-osu-webkit.mjs` の役）
 *
 *  ★★置き場★★ `%TEMP%\exally-kobore-fusagi.xlsx`（★1本の 名★）
 *    ・★司さんの 実物の 名は 1文字も 在りません★
 *    ・★手元に 立てた 配信（127.0.0.1）だけ★を 開きます
 *
 *  ★門★
 *    ①`#SPILL!` に ★なって いなければ 走りません（exit 4）★
 *      ＝★塞がって いない 物を 「塞がった 物」として 出さない★
 *    ②書き出した 中身が 0バイトなら 落とす（exit 5）
 *
 *  使い方: node docs/measured/dasu-kobore-fusagi.mjs
 */
import fs from 'node:fs';
import http from 'node:http';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { borrow, launch, unmeasured } from '../../scripts/_borrow-playwright.mjs';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..', '..');
const TAG = 'kobore-fusagi';
const 置き場 = path.join(os.tmpdir(), 'exally-kobore-fusagi.xlsx');

/* ★立て方は 他の 道具と 同じ★（読むだけ・127.0.0.1 だけ） */
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

const wk = await borrow(TAG, 'webkit');
if (!wk) { unmeasured(TAG, 'webkit'); process.exit(0); }
const b = await launch(TAG, wk, {}, 'webkit');
if (!b) { unmeasured(TAG, 'webkit'); process.exit(0); }
const p = await b.newPage({ viewport: { width: 1280, height: 900 } });
const 配信 = await 立てる(ROOT);

console.log('');
console.log('[dasu-kobore-fusagi] ★塞がれた 溢れを そのまま 書き出す★（㊺）');
console.log('  ★どこ★ ... ' + 配信.url + '/book.html');

const 返 = await p.goto(配信.url + '/book.html', { waitUntil: 'load', timeout: 60000 });
console.log('  ★返事★ ... ' + (返 ? 返.status() : '(無し)'));
await p.waitForFunction(() => typeof window.setCell === 'function', null, { timeout: 30000 });

const 出 = await p.evaluate(async () => {
  /* ★包みを 外す★＝外さないと 画布が 0x0 で 何も 打てません */
  document.body.classList.remove('exally-locked');
  const ov = document.getElementById('loginOv');
  if (ov) ov.remove();

  const 読む = (r, c) => {
    const x = window.getCell(r, c);
    if (!x) return '';
    const v = (x.d !== undefined ? x.d : x.v);
    return (v === null || v === undefined) ? '' : String(v);
  };

  /* ★★先に 邪魔を 置き、後から 溢れる 式を 打つ★★＝これで #SPILL! に なります */
  window.setCell(1, 3, 'jama');          /* D2 */
  window.setCell(0, 3, '=SEQUENCE(3)');  /* D1 */
  const sh = window.sheets[window.activeSheet];
  if (typeof window.recalcSheet === 'function') window.recalcSheet(window.activeSheet, sh.data);

  const 画面 = [];
  for (let r = 0; r < 5; r += 1) 画面.push('D' + (r + 1) + '=' + 読む(r, 3));

  /* ★★本番の 書き出しと 同じ 2行★★（`book.html:saveXlsx` 20026 / 20034） */
  await window._ensureXlsx();
  const buf = window.XlsxIO.writeBook(window.GridXlsx.gridToBook(window.sheets));
  const u8 = (buf instanceof Uint8Array) ? buf : new Uint8Array(buf);
  return { 画面, 中身: Array.from(u8) };
});

await b.close();
配信.閉じる();

console.log('');
console.log('  ★★画面の 側★★');
for (const l of 出.画面) console.log('    ' + l);

/* ★★門①★★ ... ★#SPILL! に なって いなければ 出しません★ */
const D1 = 出.画面[0];
if (D1.indexOf('SPILL') < 0 && D1.indexOf('スピル') < 0) {
  console.log('');
  console.log('  ★★D1 が 塞がって いません ... ' + D1 + '★★');
  console.log('    ⇒★塞がって いない 物を 「塞がった 物」として 出しません★');
  process.exit(4);
}

/* ★★門②★★ */
if (!出.中身 || 出.中身.length === 0) {
  console.log('  ★★書き出した 中身が 0バイトです★★');
  process.exit(5);
}

fs.writeFileSync(置き場, Buffer.from(出.中身));
console.log('');
console.log('  ★★書き出しました★★ ... ' + 置き場 + '（' + 出.中身.length + ' バイト）');
console.log('');
console.log('  ★★断り★★');
console.log('    ・★`window.setCell` で 打って います★＝★打鍵では ありません★');
console.log('    ・★この 道具は 「出す」だけ★＝実Excel で 開くのは 別の 道具');
console.log('    ・★1つの 形（縦3・邪魔 1マス）だけ★');
