/* mikire-webkit.mjs — ★字は 隣が 空なら はみ出す（実Excel と 同じ）★ 2026-09-24
 *
 *  ★★司さんの 言葉★★
 *    「★読み込んだ ファイル 見切れとん とかも 自動で 調整しろ★」
 *    「★は？ Excel内で 見切れてない とこが 見切れとるけん いよんやろが★」
 *    ⇒★Excel から 外れろ では ない★＝★Excel に 合って いない＝不具合★。
 *
 *  ★★何が 在ったか（実測 2026-09-24）★★
 *    司さんの 実物（板 15枚・19,254マス）を お客さんの 道で 開いて 数えた:
 *      ★字が 切れて いる 171マス★ ／ ★その 171マス とも 右隣が 空★
 *    ＝★実Excel なら 1つ 残らず はみ出して 全部 見えて いる 所★。
 *    ★因★ ... 描く 前に `ctx.rect(x+1,y,w-2,h); ctx.clip();` で
 *              ★マスの 四角で 必ず 切って いた★。
 *              すぐ 下に 「文字は はみ出す（実Excel と 同じ）」と 書いて あったのに、
 *              ★その 上の 切り取りが 先に 効いて いた★。
 *
 *  ★★この 試験が 見る 物★★
 *    ★印(clip の 数)では なく ★絵★ で 数えます★（2026-09-08 の 決まり）
 *    ＝canvas の ★マスの 右隣の 場所★の 点を 読み、★背景で ない 点★を 数える。
 *      はみ出して いれば ★0では ない★／切れて いれば ★0★。
 *    ★わざと 壊して 確かめる★＝右隣に 字を 置いた 行は ★切れる★（実Excel も 切る）。
 *
 *  ★★材料を 作るのに 2回 踏みました★★（2026-09-24・★残します★）
 *    ⑴★はじめは 3行だけの 材料★ ⇒ ★直す前でも 緑★ に なりました。
 *        ＝★前に 描いた 道が 少ないと 切り取りが 広く、たまたま はみ出して いた★。
 *        ⇒★★だから 画面 1枚ぶん（26列×60行）を 先に 描かせます★★。
 *    ⑵★2組目を 62行目に 置いた★ ⇒ ★画面の 外で 1つも 描かれず 点 0★。
 *        ＝★「切れて いる」と 「そもそも 描かれて いない」は 別★。
 *        ⇒★見えて いる 行（20行目・y 478／窓の 高 596）へ 移しました★。
 *  ★★今は 直す前で 赤に なります★★（実測 2026-09-24）
 *    直した後 ... B20 の 場所の 点 ★105★（はみ出して いる）
 *    直す前 ..... B20 の 場所の 点 ★0★（切れて いる）⇒ ★赤 1件★
 *    ＝★1行目だけでは 割れません★（直す前でも はみ出す）。★20行目が 割れます★。
 *
 *  ★材料は この 場で 作ります★＝`%TEMP%`（★repo に 置かない★・作り直せる）
 *
 *  走らせ方: node tests/mikire-webkit.mjs
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

/* ══ ★材料★ ══
     1行目 ... A に 長い 字 ／ B は ★空★  ⇒ ★はみ出す★
     2行目 ... A に 長い 字 ／ B に 字     ⇒ ★切れる★（実Excel も 切る）
     3行目 ... A に 長い 数 ／ B は 空      ⇒ ★`####`★（数は はみ出さない）
     列の 幅は ★狭く★ 決める（入らない 事を 作る） */
const 長い字 = 'あいうえおかきくけこさしすせそたちつてと';
const ws = {
  A1: { t: 's', v: 長い字 },
  A2: { t: 's', v: 長い字 }, B2: { t: 's', v: 'じゃま' },
  A3: { t: 'n', v: 123456789012345 },
};
/* ★★見える 所を 埋めます★★（2026-09-24 ここで 1回 踏みました）
     はじめは 3行だけの 材料でした。
     ⇒★直す前でも はみ出して 緑に なりました★＝★3行では 再現しません★。
     ⇒因は `clip()` の 前に `beginPath()` が 無い 事＝★前に 描いた 道が 残る★。
     ⇒★★だから 前に 何かを 描かせてから 測ります★★（画面 1枚ぶんの マスを 置く）。 */
for (let r = 4; r <= 60; r++) {
  for (let c = 0; c < 26; c++) {
    ws[XLSX.utils.encode_cell({ r: r - 1, c: c })] = { t: 'n', v: r * (c + 1) };
  }
}
/* ★★画面に 見えて いる 所へ もう 1組★★（2026-09-24 ここでも 1回 踏みました）
     はじめは 62行目に 置きました ⇒★画面の 外で 1つも 描かれず 点 0★に なりました。
     ⇒★「切れて いる」と 「そもそも 描かれて いない」は 別★。
     ⇒★見えて いる 行（20行目あたり）へ 移します★ */
ws.A20 = { t: 's', v: 長い字 };
/* ★20行目は A だけ 残して 空ける★＝★はみ出す 場所を 作る★
     （消さないと C30 に 数が 在り、★実Excel も そこで 切ります★＝試験の 立て方の 誤り） */
for (let c = 1; c < 26; c++) delete ws[XLSX.utils.encode_cell({ r: 19, c: c })];
ws.A21 = { t: 's', v: 長い字 };
for (let c = 2; c < 26; c++) delete ws[XLSX.utils.encode_cell({ r: 20, c: c })];
ws.B21 = { t: 's', v: 'じゃま' };
ws['!ref'] = 'A1:Z63';
ws['!cols'] = [];
for (let c = 0; c < 26; c++) ws['!cols'].push({ width: 5 });
const wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, ws, '見切れ');
const 材料 = path.join(os.tmpdir(), 'exally-mikire.xlsx');
fs.writeFileSync(材料, XLSX.write(wb, { bookType: 'xlsx', type: 'buffer' }));

console.log('[mikire] ★字は 隣が 空なら はみ出す（実Excel と 同じ）★');
console.log('      ── 材料 ── ' + path.basename(材料) + '（列の 幅 5字・字 ' + 長い字.length + '文字）');

const wk = await borrow('mikire', 'webkit');
const browser = await launch('mikire', wk, {}, 'webkit');
const page = await browser.newPage({ viewport: { width: 1280, height: 900 } });
const 配信 = await 立てる(ROOT);

try {
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
    for (let i = 0; i < ss.length; i++) {
      if (ss[i] && ss[i].data && Object.keys(ss[i].data).length > 0) return true;
    }
    return false;
  }, null, { timeout: 120000 });
  await page.waitForTimeout(800);

  /* ★★絵を 読む★★＝マスの 右隣の 場所の 点を 数える（★印では なく 絵★） */
  const 出 = await page.evaluate(() => {
    /* ★★絵を 読む 前に 「どの 板を どの 倍で」 を 数える★★（2026-09-24 ここで 1回 踏みました）
         ・canvas は 1枚では ありません（★最初の 1枚を 取ると 別の 板を 読みます★）
         ・★点の 倍（devicePixelRatio）★ が 1で ない 時、
           `colX()` の 返す 点と `getImageData` の 点は ★別の 物差し★ です。
         ⇒★本体が 使って いる `ctx` の canvas★ を 取り、★倍を 測って 掛けます★ */
    const g = window.ctx;
    const cv = g.canvas;
    const 倍 = cv.width / (cv.clientWidth || cv.width);
    const 見る = (r, c) => {
      /* ★その マスの 「右隣」の 帯★を 読む */
      const x = window.colX(c + 1), y = window.rowY(r);
      const w = window.cW(c + 1), h = window.rH(r);
      if (!(w > 0 && h > 0)) return { 点: -1 };
      const d = g.getImageData(Math.round((x + 1) * 倍), Math.round((y + 2) * 倍),
        Math.max(1, Math.round((w - 2) * 倍)), Math.max(1, Math.round((h - 4) * 倍))).data;
      /* ★背景で ない 点★＝白(255,255,255)から 離れた 物（罫線の 薄い 灰も 拾わない よう 差 60） */
      let 数 = 0;
      for (let i = 0; i < d.length; i += 4) {
        if (d[i + 3] < 8) continue;
        if (255 - d[i] > 60 || 255 - d[i + 1] > 60 || 255 - d[i + 2] > 60) 数++;
      }
      return { 点: 数, 幅: w, 高: h, 倍: 倍 };
    };
    const 字 = (r, c) => {
      const sh = window.sheets[window.activeSheet];
      const cell = sh.data[r + ',' + c];
      return cell ? String(cell.d || cell.v || '') : '';
    };
    return {
      板: window.sheets[window.activeSheet].name,
      行20のy: window.rowY(19), 窓の高: window.wrapH,
      canvasの数: document.querySelectorAll('canvas').length,
      倍: 倍, canvas幅: cv.width, 見た目幅: cv.clientWidth,
      列の幅: window.cW(0),
      右が空: 見る(0, 0),        /* 1行目＝はみ出す はず */
      右に字: 見る(1, 0),        /* 2行目＝切れる はず */
      数の右: 見る(2, 0),        /* 3行目＝数は はみ出さない */
      B2の字: 字(1, 1),
      下の右が空: 見る(19, 0),   /* 20行目＝たくさん 描いた 後 */
      下の右に字: 見る(20, 0),   /* 21行目 */
    };
  });

  console.log('      ── 実測 ── 板 ' + 出.板 + ' ／ A列の 幅 ' + 出.列の幅 + '点'
    + ' ／ canvas ' + 出.canvasの数 + '枚 ／ 倍 ' + 出.倍 + '（' + 出.canvas幅 + '/' + 出.見た目幅 + '）');
  console.log('        1行目（右隣が 空）... B1 の 場所の 点 ' + 出.右が空.点);
  console.log('        2行目（右隣に 字）... B2 の 場所の 点 ' + 出.右に字.点 + '（B2 は 「' + 出.B2の字 + '」）');
  console.log('        3行目（数）......... B3 の 場所の 点 ' + 出.数の右.点);
  console.log('        20行目（★たくさん 描いた 後★）... B20 の 場所の 点 ' + 出.下の右が空.点
    + '（20行目の y ' + Math.round(出.行20のy) + ' ／ 窓の 高 ' + Math.round(出.窓の高) + '）');
  console.log('        21行目（右隣に 字）... B21 の 場所の 点 ' + 出.下の右に字.点);

  T('★★右隣が 空なら 字が はみ出して 出る★★（★実Excel と 同じ★・09-24 の 171マス）',
    出.右が空.点 > 0, 'B1 の 場所に 点が ' + 出.右が空.点 + '個＝★切れて います★');

  T('★右隣に 字が 在れば 切る★（実Excel も 切る＝重ねて 描かない）',
    出.右に字.点 > 0, 'B2 の 字が 出て いない');

  T('★★たくさん 描いた 後でも はみ出す★★（★ここが 直す前は 割れて いた 所★）',
    出.下の右が空.点 > 0, 'B20 の 場所に 点が ' + 出.下の右が空.点 + '個＝★切れて います★（★描かれて いないだけ かも★）');

  T('★★数は はみ出さない★★（入らなければ `####`＝実Excel と 同じ）',
    出.数の右.点 === 0, 'B3 の 場所に 点が ' + 出.数の右.点 + '個＝★数が はみ出して います★');

  /* ★★わざと 壊して 確かめる★★＝B1 に 字を 置くと ★はみ出しが 止まる★ */
  await page.evaluate(() => { window.setCell(0, 1, 'ふさぐ'); });
  /* ★★描き直しを 待ちます★★（2026-09-24 ここで 1回 踏みました）
       打った 直後に 点を 読んだら ★前の 絵★ を 読んで いました。
       ⇒★『打った』と『描かれた』は 別★ です。 */
  await page.waitForTimeout(600);
  const 壊した = await page.evaluate(() => {
    const g = window.ctx;
    const cv = g.canvas;
    const 倍 = cv.width / (cv.clientWidth || cv.width);
    /* ★B の 右（C）の 場所★を 見る＝B が 埋まったので ここには 何も 来ない はず */
    const x = window.colX(2), y = window.rowY(0), w = window.cW(2), h = window.rH(0);
    const d = g.getImageData(Math.round((x + 1) * 倍), Math.round((y + 2) * 倍),
      Math.max(1, Math.round((w - 2) * 倍)), Math.max(1, Math.round((h - 4) * 倍))).data;
    let 数 = 0;
    for (let i = 0; i < d.length; i += 4) {
      if (d[i + 3] < 8) continue;
      if (255 - d[i] > 60 || 255 - d[i + 1] > 60 || 255 - d[i + 2] > 60) 数++;
    }
    return 数;
  });
  console.log('        ★塞いだ 後★ ... C1 の 場所の 点 ' + 壊した);
  T('★★塞いだら はみ出しが 止まる★★（★隣を 上書きしない★）',
    壊した === 0, 'C1 の 場所に 点が ' + 壊した + '個＝★塞いでも はみ出して います★');
} finally {
  await page.close().catch(() => {});
  await browser.close().catch(() => {});
  配信.閉じる();
}

console.log('\n' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
