/* naraberu-jitsu-excel-to-exally.mjs
 *   ★実Excel の 絵と Exally の 絵を ★1枚に 並べる★★（2026-09-21）
 *
 *  ★★なぜ 要るか★★
 *    司さん「★全く 同じように 表示できるように しろや★」（イ）
 *    ⇒★今まで 一度も 「実Excel の 絵」と 並べて いません★
 *    ⇒数が 13/13 でも ★並べるまで 「同じ」とは 言えません★
 *
 *  ★★材料★★
 *    ①実Excel の 絵 ... `docs/measured/e/jitsu-excel-kazari3-2026-09-21.png`
 *        1209 x 567 点 ／ 範囲 A1:F12 ／ マスは 453.4 x 212.6 ポイント
 *        ★Excel 自身に 描かせた 物★（`CopyPicture` ⇒ 図に 貼る ⇒ `Export`）
 *        ⇒★窓の 縁・リボン・行列の 見出し・選んだ 印が 1つも 入って いません★
 *    ②Exally の 絵 ... ここで 撮ります
 *        ⇒★だから こちらも 見出しを 外して 同じ 範囲だけ 切ります★
 *        ⇒★リボンも 入れません★（キャンバスの 中だけ 切るので 入りません）
 *
 *  ★★大きさを 揃える★★（2026-09-21 数えました）
 *    実Excel  A1:F12 ＝ 453.4 x 212.6 ポイント ＝ ★604.5 x 283.5 点★（1pt ＝ 4/3点）
 *    Exally   A列 245 ＋ B〜F 72x5 ＝ ★605 点★ ／ 12行 x 24 ＝ ★288 点★
 *    ⇒★横の 差 0.5点 ／ 縦の 差 4.5点★（★1行あたり 0.375点★）
 *    ⇒並べる 時は ★実Excel の 絵の 幅（1209点）に 揃えます★
 *
 *  ★★見て いない 事★★
 *    ・★字の 形（書体）は 合わせて いません★＝Exally は 画面の 書体で 描きます
 *    ・★★右端の 濃い 青の 棒 ＝ 判子（hanko）★★（2026-09-21 並べて 分かりました）
 *      ＝経営者1 が 「未測定」と 書いて いた 物。★同じ 所に うちの 判子が 出ました★
 *      ＝色も 合いました（テーマの accent1 ＝ #156082）
 *      ⇒★絵の 範囲が A1:F12 なので 右端で 切れて います★
 *    ・★点を 1つずつ 突き合わせて いません★＝★目で 見る 為の 絵★です
 *
 *  使い方: node docs/measured/naraberu-jitsu-excel-to-exally.mjs
 */
import path from 'node:path';
import http from 'node:http';
import fs from 'node:fs';
import crypto from 'node:crypto';
import { fileURLToPath } from 'node:url';
import { borrow, launch } from '../../scripts/_borrow-playwright.mjs';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..', '..');
const TEMP = process.env.TEMP || process.env.TMP || '.';

const 実の絵 = 'docs/measured/e/jitsu-excel-kazari3-2026-09-21.png';
const 材料 = path.join(ROOT, 'tests/fixtures/kazari-hiraku3.xlsx');

function 立てる(root) {
  const 型 = {
    '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8',
    '.css': 'text/css; charset=utf-8', '.json': 'application/json; charset=utf-8',
    '.png': 'image/png', '.svg': 'image/svg+xml', '.ico': 'image/x-icon',
  };
  const s = http.createServer((q, r) => {
    const f = path.join(root, decodeURIComponent(String(q.url).split('?')[0]).replace(/^\/+/, ''));
    if (!f.startsWith(root) || !fs.existsSync(f) || fs.statSync(f).isDirectory()) {
      r.statusCode = 404; return r.end('no');
    }
    r.setHeader('content-type', 型[path.extname(f).toLowerCase()] || 'application/octet-stream');
    fs.createReadStream(f).pipe(r);
  });
  return new Promise((x) => s.listen(0, '127.0.0.1',
    () => x({ url: 'http://127.0.0.1:' + s.address().port, 閉じる: () => s.close() })));
}

console.log('');
console.log('[naraberu] ★実Excel の 絵と Exally の 絵を 1枚に 並べる★');

if (!fs.existsSync(path.join(ROOT, 実の絵))) {
  console.log('  ★実Excel の 絵が 有りません★ ' + 実の絵);
  console.log('    ⇒`docs/measured/toru-jitsu-excel-no-e.ps1` で 撮れます（COM が 要ります）');
  process.exit(2);
}
if (!fs.existsSync(材料)) { console.log('  ★材料が 無い★ ' + 材料); process.exit(2); }

const 実 = fs.readFileSync(path.join(ROOT, 実の絵));
console.log('      ＝ 実Excel の 絵 ' + 実.length + 'B ／ sha256 '
  + crypto.createHash('sha256').update(実).digest('hex'));

const wk = await borrow('naraberu', 'webkit');
const browser = await launch('naraberu', wk, {}, 'webkit');
const page = await browser.newPage({ viewport: { width: 1400, height: 900 } });
const 配信 = await 立てる(ROOT);

try {
  await page.goto(配信.url + '/book.html', { waitUntil: 'load', timeout: 60000 });
  await page.evaluate(() => {
    document.body.classList.remove('exally-locked');
    const ov = document.getElementById('loginOv');
    if (ov) { ov.classList.remove('open'); ov.style.display = 'none'; }
  });
  await page.setInputFiles('#bookFileInput', 材料);
  await page.waitForFunction(() => {
    const sh = (window.sheets || [])[window.activeSheet || 0];
    return !!(sh && sh.data && Object.keys(sh.data).length > 3);
  }, null, { timeout: 60000 }).catch(() => {});
  /* ★★選んだ 印を 外します★★（2026-09-21）
       ＝実Excel の 絵には ★選んだ 印が 入って いません★（経営者1 の 道具）
       ⇒うちの 絵にだけ ★青い 枠と 水色の 塗り★ が 入ると 並べた 時に 嘘に なります
       ⇒★ずっと 遠くの マスへ 寄せます★（A1:F12 の 外） */
  /* ★★選んだ マスを 絵の 外へ 動かします★★
       ＝`selR1` は `var` なので 窓（window）にも 出て います。
       ＝★但し 書き換えただけでは 描き直されません★
       ⇒★下の 方の マスを 「押して」 動かします★（★お客さんの 道★）
       ⇒2026-09-21 に 1回 踏みました＝★変数だけ 書き換えて 絵に 残って いました★ */
  await page.evaluate(() => {
    const cv = document.getElementById('grid-canvas');
    if (!cv || typeof window.colX !== 'function') return;
    const r = cv.getBoundingClientRect();
    /* ★A1:F12 の 外＝行 20 あたりを 押します★ */
    const x = r.left + window.colX(1) + 10;
    const y = r.top + window.rowY(20) + 5;
    for (const 種 of ['mousedown', 'mouseup', 'click']) {
      cv.dispatchEvent(new MouseEvent(種, { clientX: x, clientY: y, bubbles: true }));
    }
  });
  await page.waitForTimeout(300);

  /* ★知らせの 札が 消えるのを 待ちます★＝★絵に 写り込みます★ */
  await page.waitForFunction(() => {
    const t = document.querySelectorAll('.toast, #toast');
    for (const x of t) { if (x.offsetParent !== null) return false; }
    return true;
  }, null, { timeout: 20000 }).catch(() => {});

  const 出 = await page.evaluate(async (実の道) => {
    const cv = document.getElementById('grid-canvas');
    if (!cv) return { だめ: '(キャンバスが 無い)' };
    if (typeof window.colX !== 'function') return { だめ: '(colX が 無い)' };
    /* ★★切る 範囲＝A1:F12★★（実Excel の 絵と 同じ）
         ★見出しを 外します★＝`colX(0)` `rowY(0)` が マス A1 の 左上 */
    const 倍 = cv.width / cv.clientWidth;
    const x0 = window.colX(0), y0 = window.rowY(0);
    const x1 = window.colX(6), y1 = window.rowY(12);
    const w = Math.round((x1 - x0) * 倍), h = Math.round((y1 - y0) * 倍);
    if (w <= 0 || h <= 0) return { だめ: '(範囲が 0) w=' + w + ' h=' + h };

    const 切 = document.createElement('canvas');
    切.width = w; 切.height = h;
    切.getContext('2d').drawImage(cv, Math.round(x0 * 倍), Math.round(y0 * 倍), w, h, 0, 0, w, h);

    /* ★実Excel の 絵を 読み込みます★ */
    const 実絵 = await new Promise((ok, ng) => {
      const im = new Image();
      im.onload = () => ok(im);
      im.onerror = () => ng(new Error('実Excel の 絵を 読めません'));
      im.src = 実の道;
    });

    /* ★★1枚に 並べます★★（上＝実Excel ／ 下＝Exally・★同じ 幅に 揃える★） */
    const 幅 = 実絵.width;
    const 下高 = Math.round(h * (幅 / w));
    const 札 = 34;
    const 出 = document.createElement('canvas');
    出.width = 幅;
    出.height = 札 + 実絵.height + 札 + 下高 + 札;
    const g = 出.getContext('2d');
    g.fillStyle = '#FFFFFF'; g.fillRect(0, 0, 出.width, 出.height);
    g.imageSmoothingEnabled = true;

    const 書く = (字, y) => {
      g.fillStyle = '#1A2B22';
      g.font = 'bold 20px "Noto Sans JP",sans-serif';
      g.textBaseline = 'middle';
      g.fillText(字, 10, y + 札 / 2);
    };
    let y = 0;
    /* ★★測った 数を 絵の 中に 焼き込みます★★（札だけ 付けない） */
    書く('実Excel（16.0 build 20326）  A1:F12  ' + 実絵.width + ' x ' + 実絵.height + '点', y);
    y += 札;
    g.drawImage(実絵, 0, y);
    y += 実絵.height;
    書く('Exally  A1:F12  切った 元 ' + w + ' x ' + h + '点 ⇒ ' + 幅 + ' x ' + 下高 + '点に 伸ばした', y);
    y += 札;
    g.drawImage(切, 0, 0, w, h, 0, y, 幅, 下高);
    y += 下高;
    書く('★まだ 合って いない★ 書体／行の 高さ 1行 0.375点／右端の 濃い 青は ★判子★（両方 accent1 #156082）', y);

    return { だめ: '', 絵: 出.toDataURL('image/png'), w: w, h: h, 実w: 実絵.width, 実h: 実絵.height };
  }, 配信.url + '/' + 実の絵);

  if (出.だめ) throw new Error(出.だめ);

  const 中 = Buffer.from(String(出.絵).split(',')[1], 'base64');
  const 置き場 = path.join(TEMP, 'exally-narabeta.png');
  fs.writeFileSync(置き場, 中);
  console.log('      ＝ Exally を 切った 大きさ ' + 出.w + ' x ' + 出.h + '点');
  console.log('      ＝ 実Excel の 絵 ' + 出.実w + ' x ' + 出.実h + '点');
  console.log('');
  console.log('  ★並べた 絵★ ' + 置き場);
  console.log('    ' + 中.length + 'B ／ sha256 '
    + crypto.createHash('sha256').update(中).digest('hex'));
  console.log('');
  console.log('  ★これは 目で 見る 為の 絵です★＝★点を 1つずつ 突き合わせて いません★');
} catch (e) {
  console.log('  ★止まりました★ ' + String(e && e.message).slice(0, 200));
  process.exitCode = 1;
} finally {
  if (配信) 配信.閉じる();
  await browser.close();
}
