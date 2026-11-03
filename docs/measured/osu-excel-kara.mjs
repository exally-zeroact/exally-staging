/* osu-excel-kara.mjs — ★Excel が 作った ファイルを うちで 開いて 突き合わせる★（2026-09-11）
 *
 *  ★★向き★★（3つ あって ★③だけ 測って いませんでした★）
 *    ① 実Excel に 打たせて 答えを 合わせる … 前から やって いた
 *    ② うちが 書いた 物を Excel で 開き直す … 2026-09-10 に 足した
 *    ③★Excel が 書いた 物を うちで 開く★ … ★これ★（お客さんが 一番 よく やる 道）
 *
 *  ★★お客さんの 道で 開きます★★
 *    本物の ブラウザで book.html を 開き、★「Excelを読み込む」の 入口（bookFileInput）に
 *    実物の ファイルを 渡します★＝JS で 板を 作ったり しません。
 *
 *  ★見るのは 3つ★ ①式 ②出る字 ③答え（★答えだけ 見ない★）
 *  ★司さんの 実物には 触りません★＝実Excel に 作らせた 材料だけ。
 *
 *  使い方:
 *    ① pwsh -NoProfile -File docs/measured/toru-excel-kara.ps1   … 実Excel に 作らせる
 *    ② node docs/measured/osu-excel-kara.mjs                     … うちで 開いて 突き合わせ
 */
import fs from 'node:fs';
import http from 'node:http';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { borrow, launch, unmeasured } from '../../scripts/_borrow-playwright.mjs';

const ROOT = path.resolve(path.join(path.dirname(fileURLToPath(import.meta.url)), '..', '..'));
const ここ = path.join(ROOT, 'docs/measured');
const TAG = 'excel-kara';
const 材料 = path.join(ここ, 'excel-kara-2026-09-11.xlsx');
const 紙 = path.join(ここ, 'golden-excel-kara-2026-09-11.tsv');
const 出す先 = path.join(ここ, 'golden-excel-kara-awase-2026-09-11.tsv');

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

if (!fs.existsSync(材料) || !fs.existsSync(紙)) {
  console.log('  ★未測定★ 先に `pwsh -NoProfile -File docs/measured/toru-excel-kara.ps1` を 走らせて ください');
  process.exit(0);
}

/* ★実Excel の 正解を 読む★ */
const 実 = new Map();
for (const l of fs.readFileSync(紙, 'utf-8').split(/\r?\n/)) {
  if (!l || l.startsWith('#')) continue;
  const c = l.split('\t');
  if (c.length < 8) continue;
  /* ★2026-09-11 に 窓を 2つ 足しました★（tests/monosashi-mado.test.mjs が 求める 形）
       型   … String と Number を 取り違えない
       窓２ … 字が「0」に 見えた 時 実Excel に `=(A1)=0` を 打って 本当に 0 か 聞く
     ⇒★足した ので 「何を 見て いるか」は 8つ目に 動きました★ */
  実.set(c[0], { 入: c[1], 式: c[2], 字: c[3], 答: c[4], 型: c[5], 窓2: c[6], 何: c[7] });
}
if (!実.size) { console.error('★紙が 読めない★'); process.exit(2); }

const chromium = await borrow(TAG, 'chromium');
if (!chromium) { unmeasured(TAG, 'chromium'); process.exit(0); }
const browser = await launch(TAG, chromium, {}, 'chromium');
if (!browser) { unmeasured(TAG, 'chromium'); process.exit(0); }
const page = await browser.newPage({ viewport: { width: 1200, height: 700 } });
const 配信 = await 立てる(ROOT);

let 合 = 0, 違 = 0;
const 実物 = [];
const 行 = [];

try {
  await page.goto(配信.url + '/book.html', { waitUntil: 'load' });
  await page.evaluate(() => {
    document.body.classList.remove('exally-locked');
    const ov = document.getElementById('loginOv');
    if (ov) ov.style.display = 'none';
  });
  await page.waitForTimeout(500);

  /* ★お客さんの 入口に 実物の ファイルを 渡す★
     ★入口は わざと 隠して 在ります★（display:none）＝見えるまで 待つ 道では 掴めない。
     ⇒ 掴んでから 渡す（★押す 相手は お客さんと 同じ 入口★） */
  const 入口 = await page.waitForSelector('#bookFileInput', { state: 'attached', timeout: 15000 });
  await 入口.setInputFiles(材料);
  await page.waitForFunction(() => {
    try { return sheets && sheets[activeSheet] && Object.keys(sheets[activeSheet].data).length > 5; }
    catch (e) { return false; }
  }, null, { timeout: 30000 });
  await page.waitForTimeout(1500);

  /* ★★出る字は「画面が 本当に 描いた 字」を 取ります★★
     作り直すと ★物差しの 方が 間違う★（2026-09-11 に 実際 やらかしました）
       前の 版は 答えの 欄に ★出る字★を 入れて いて、
       A5（1,234,567.89 と 1234567.891）が ★0.001 の 差を すり抜けて 緑★に なりました
     ⇒ 描く 所（fillText）を 一度だけ 覗いて、★描かれた 字★を そのまま 使います */
  await page.evaluate(() => {
    window.__描いた = [];
    const 元 = CanvasRenderingContext2D.prototype.fillText;
    CanvasRenderingContext2D.prototype.fillText = function (t, x, y) {
      try { window.__描いた.push([String(t), x, y]); } catch (e) { /* 何も しない */ }
      return 元.apply(this, arguments);
    };
    window.__描き直す = () => { window.__描いた = []; render(); };
  });
  await page.evaluate(() => window.__描き直す());
  await page.waitForTimeout(300);

  const 読む = (a) => page.evaluate((n) => {
    const m = /^([A-Z]+)(\d+)$/.exec(n);
    let c = 0; for (const ch of m[1]) c = c * 26 + (ch.charCodeAt(0) - 64);
    c -= 1;
    const r = Number(m[2]) - 1;
    const 板 = sheets[activeSheet];
    const cell = 板.data[r + ',' + c];
    /* ★描かれた 字の うち、この マスの 箱に 入って いる 物★ */
    const x = colX(c), y = rowY(r), w = cW(c), h = rH(r);
    let 字 = '';
    for (const [t, tx, ty] of (window.__描いた || [])) {
      if (tx >= x - 1 && tx <= x + w + 1 && ty >= y - 1 && ty <= y + h + 1) { 字 = t; break; }
    }
    return {
      式: cell && cell.f ? String(cell.f) : '',
      /* ★答えは エンジンに 直接 聞く★
         ＝式の マスは cell.v を 持たない（★前の 版は ここが 空で 全部 NG に なった★） */
      値: (() => {
        try {
          const v = hf.getCellValue({ sheet: hf.getSheetId(板.name != null ? 板.name : hf.getSheetNames()[activeSheet]), row: r, col: c });
          if (v && v.type) return '#' + v.type;
          if (v == null) return cell && cell.v != null ? String(cell.v) : '';
          return String(v);
        } catch (e) { return cell && cell.v != null ? String(cell.v) : '(読めない)'; }
      })(),
      列の幅: cW(c),
      字: 字,
    };
  }, a);

  console.log('');
  console.log('[' + TAG + '] ★Excel が 作った ファイルを うちで 開いた★');

  const 数と読む = (s) => {
    const n = Number(String(s).replace(/,/g, ''));
    return Number.isFinite(n) ? n : null;
  };

  行.push('# ★Excel が 作った 物を うちで 開いて 突き合わせた★（2026-09-11）');
  行.push('#');
  行.push('# ★この 向きは 今日 初めて 測りました★（お客さんが 一番 よく やる 道）');
  行.push('# ★お客さんの 入口（Excelを読み込む）に 実物を 渡して います★');
  行.push('#');
  行.push(['# マス', '実Excel の 式', 'うちの 式', '実Excel の 答え', 'うちの 答え', '判じ', '何を 見て いるか'].join('\t'));

  for (const [マス, e] of 実) {
    const u = await 読む(マス);
    const 窓 = [];
    /* ①★答え★＝中の 数（★出る字と 混ぜない★）
       ★実Excel が String と 言って いる 物を 数として 比べない★
       ＝TEXT の 答え「3.00」を 数に すると 3 と 同じに なって しまい、
         ★桁が 落ちた 事を 見逃します★（記憶の 決まり … 中の数が同じは同じでない） */
    const a = 数と読む(e.答), b = 数と読む(u.値);
    let 答合う;
    if (e.型 === 'String') 答合う = String(e.答) === String(u.値);
    else if (a !== null && b !== null) 答合う = Math.abs(a - b) <= Math.max(1e-12, Math.abs(a) * 1e-9);
    else 答合う = String(e.答) === String(u.値);
    if (!答合う) 窓.push('答え Excel=' + e.答 + ' うち=' + u.値);
    /* ②★出る字★＝画面が 描いた 字（大文字小文字も 見る） */
    const 字合う = String(e.字) === String(u.字);
    if (!字合う) 窓.push('出る字 Excel=' + JSON.stringify(e.字) + ' うち=' + JSON.stringify(u.字));
    /* ③★式★＝Excel が 式を 持って いたのに うちが 落として いないか */
    const 式合う = !(e.式.charAt(0) === '=' && u.式.charAt(0) !== '=');
    if (!式合う) 窓.push('式が 落ちた Excel=' + e.式);

    const 判 = 窓.length ? '★' + 窓.length + 'つ 違う★' : '合った';
    if (!窓.length) 合++; else { 違++; 実物.push(マス + '（' + e.何 + '）\n    ' + 窓.join('\n    ')); }
    行.push([マス, e.式, u.式, e.答, u.値, e.字, u.字, 判, e.何].join('\t'));
    console.log('  ' + マス.padEnd(4) + (窓.length ? ' NG  ' : ' ok  ') + '答え ' + String(u.値).padEnd(22) + '出る字 ' + JSON.stringify(u.字));
  }
} catch (e) {
  違++;
  console.error('★落ちた … ' + e.message + '★');
} finally {
  await browser.close().catch(() => {});
  配信.閉じる();
}

行.push('#');
行.push('# ★★締め★★ 全 ' + (合 + 違) + '本 ／ 合った ' + 合 + ' ／ ★違う ' + 違 + '★');
fs.writeFileSync(出す先, 行.join('\n') + '\n', 'utf-8');
console.log('');
console.log('★★締め★★ 全 ' + (合 + 違) + '本 ／ 合った ' + 合 + ' ／ ★違う ' + 違 + '★');
if (実物.length) { console.log(''); console.log('★合わない 物の 実物★'); 実物.forEach((x) => console.log('  ' + x)); }
console.log('');
console.log('★書いた … ' + 出す先 + '★');
process.exit(違 ? 1 : 0);
