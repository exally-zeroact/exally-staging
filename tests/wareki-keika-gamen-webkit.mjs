/* wareki-keika-gamen-webkit.mjs ･･･ ★和暦・経過時間の 書式の マスが 画面で 実Excel と 同じ 字か★ 2026-10-04
 *
 *  ★★なぜ 要るか★★
 *    TEXT() の 和暦・経過時間は 10-02 に 作った（lib/shoshiki.js・真値 155/155）。
 *    ★画面は 別の 道★＝`fmtForDisplay` が 開いた 本で ★XLSX.SSF（借り物の 書式器）★を 使う。
 *    ⇒ lib の 段で 当てると ★㋒（実Excel の マスの .Text）と 48/155 しか 合わない★（10-02 測り）。
 *
 *  ★★物差し★★ ＝ ㋒ `docs/measured/golden-sel-shoshiki-wareki-keika-2026-10-02.tsv`
 *    経営者が 実Excel 16.0.20430 で マスに 書式を 付けて 取った ★画面の字（.Text）★ 155本
 *    ★負の 時刻 6本は ####★（列幅を 60 に しても）＝★# だけの 字なら 同じ★ と 数える
 *
 *  ★★測り方★★（★画面の 字は 画面の 関数で 作る＝`_gamen-no-michi.mjs`★）
 *    155組を ★値＋書式★ の マスに した `.xlsx` を 作り（列は 広く）、本物の 受け口で 開いて
 *    ★1マスずつ 画面の 字を 作って★ 突き合わせる。
 *
 *  ★走らせ方★: node tests/wareki-keika-gamen-webkit.mjs
 */
import path from 'node:path'; import http from 'node:http'; import fs from 'node:fs';
import os from 'node:os';
import { createRequire } from 'node:module';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { borrow, launch } from '../scripts/_borrow-playwright.mjs';
import { 道を確かめる, 道の字 } from '../docs/measured/_gamen-no-michi.mjs';

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

/* ══ ★物差し★ ══（1列目＝組・2列目＝値の式・3列目＝書式・4列目＝画面の字） */
/* ★10-04 に 指数の 形の 値 16組（組 E：1E-7・-5.55E-17・1E+21・0.1+0.2-0.3）を 足した 171組の 紙へ★
     （前の 155組は 1行も 変わって いない＝同じ 道具・同じ Excel） */
/* ★10-04 夕★ 日付の 上限（通し 2958465〜2958466.5）の 紙 2枚も 重ねて 読む（同じ 道具・同じ Excel）
     組 B ＝和暦・経過・yyyy/m/d の 境目 ／ 組 C ＝ほかの 日付・時刻の 書式 14 の 境目
     ★組 J（字の "1e-7"）は 別の 1件★＝ここでは 読まない（下の 字の マスで 別に 見る）
     ★同じ（値の式・書式）の 組は 1つに★ */
const 紙たち = [
  'docs/measured/golden-sel-shoshiki-shisuu-2026-10-04.tsv',
  'docs/measured/golden-sel-shoshiki-ji-to-sakaime-2026-10-04.tsv',
  'docs/measured/golden-sel-shoshiki-sakaime-hoka-2026-10-04.tsv',
];
const 紙 = 紙たち.map((p) => path.basename(p)).join(' ＋ ');
const 行 = [];
const 見た鍵 = new Set();
for (const p of 紙たち) {
  const L = fs.readFileSync(path.join(ROOT, p), 'utf8').split(/\r?\n/).filter((l) => l && l.charAt(0) !== '#');
  const 頭 = L.shift().split('\t');
  const 列 = (名) => 頭.indexOf(名);
  const c組 = 列('組'), c式 = 列('値の式'), c書 = 列('渡した書式(Local)'), c字 = 列('画面の字(.Text)');
  if (c組 < 0 || c式 < 0 || c書 < 0 || c字 < 0) { console.log('★物差しの 列が 読めない★ ' + p); process.exit(8); }
  for (const l of L) {
    const c = l.split('\t');
    if (c[c組] === 'J') continue;
    const 鍵 = c[c式] + '\t' + c[c書];
    if (見た鍵.has(鍵)) continue;
    見た鍵.add(鍵);
    行.push([c[c式], c[c書], c[c字]]);
  }
}
const c式 = 0, c書 = 1, c字 = 2;

/* ★値の式を 通し番号に★（DATE は 1900年の 起点 2つ＝通し 60 以前は 1日 前） */
function 通し(式) {
  const s = String(式).replace(/^=/, '');
  const d = /^DATE\((\d+),(\d+),(\d+)\)$/.exec(s);
  if (d) {
    let n = (Date.UTC(+d[1], +d[2] - 1, +d[3]) - Date.UTC(1899, 11, 30)) / 864e5;
    if (n < 61) n -= 1;
    return n;
  }
  if (!/^[-0-9.*/ eE+()]+$/.test(s)) return null;   /* ★指数（1E-7）と 足し算（0.1+0.2-0.3）も★ */
  return Function('return (' + s + ')')();
}
const 組 = 行.map((c) => ({ 式: c[c式], 書: c[c書], 字: c[c字], 値: 通し(c[c式]) }));
const 読めない = 組.filter((x) => x.値 === null || !isFinite(x.値));
console.log('[wareki-keika-gamen] ★和暦・経過時間の マスが 画面で 実Excel と 同じ 字か★');
console.log('  ★物差し★ ' + path.basename(紙) + ' … ' + 組.length + '組（値を 読めない ' + 読めない.length + '）');
if (組.length < 171 + 24 + 56 || 読めない.length) { console.log('★物差しが 足りない／読めない＝空振り★'); process.exit(8); }

/* ══ ★材料★ ══（1行1組・A列・列は 広く） */
const ws = {};
組.forEach((x, i) => { ws[XLSX.utils.encode_cell({ r: i, c: 0 })] = { t: 'n', v: x.値, z: x.書 }; });
/* ★字の "1e-7" の マス★（経営者の 叩き⑴）＝★台の 道に 乗らない★ こと（乗れば `[s]` で "0" に なる） */
const 字の行 = 組.length;
ws[XLSX.utils.encode_cell({ r: 字の行, c: 0 })] = { t: 's', v: '1e-7', z: '[s]' };
/* ★★組 J（10-04 夕・経営者の 真値）★★ 字の マスは 書式に 関わらず 字の まま
     `'1e-7`（字として 打った）と `="1e-7"`（答えが 字の 式）× [s]・[h]:mm:ss・ge.m.d・0.00・General
     ⇒ 実Excel は ★10組とも 「1e-7」★ ／ ★空の 字★は 空 */
const 組J = [];
{
  const L = fs.readFileSync(path.join(ROOT, 'docs/measured/golden-sel-shoshiki-ji-to-sakaime-2026-10-04.tsv'), 'utf8')
    .split(/\r?\n/).filter((l) => l && l.charAt(0) !== '#');
  const 頭 = L.shift().split('\t');
  const i組 = 頭.indexOf('組'), i式 = 頭.indexOf('値の式'), i書 = 頭.indexOf('渡した書式(Local)'), i字 = 頭.indexOf('画面の字(.Text)');
  for (const l of L) {
    const c = l.split('\t');
    if (c[i組] !== 'J') continue;
    組J.push({ 式: c[i式], 書: c[i書] === 'General' ? 'General' : c[i書], 字: c[i字] });
  }
}
if (組J.length !== 10) { console.log('★組 J が 10組で ない★ ' + 組J.length); process.exit(8); }
const J頭 = 字の行 + 1;
組J.forEach((x, k) => {
  const 式か = x.式.charAt(0) === '=';
  ws[XLSX.utils.encode_cell({ r: J頭 + k, c: 0 })] = 式か
    ? { t: 's', v: '1e-7', f: '"1e-7"', z: x.書 }
    : { t: 's', v: '1e-7', z: x.書 };
});
const 空の行 = J頭 + 組J.length;
ws[XLSX.utils.encode_cell({ r: 空の行, c: 0 })] = { t: 's', v: '', z: '0.00' };
ws['!ref'] = 'A1:A' + (空の行 + 1);
ws['!cols'] = [{ wch: 60 }];
const wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, ws, 'あ');
const 材料 = path.join(os.tmpdir(), 'exally-wareki-keika-gamen.xlsx');
fs.writeFileSync(材料, XLSX.write(wb, { bookType: 'xlsx', type: 'buffer' }));

const wk = await borrow('wareki-keika-gamen', 'webkit');
const browser = await launch('wareki-keika-gamen', wk, {}, 'webkit');
const page = await browser.newPage({ viewport: { width: 1280, height: 900 } });
const 配信 = await 立てる(ROOT);

try {
  page.on('pageerror', (e) => console.log('      [落ちた] ' + String(e.message).slice(0, 200)));
  await page.goto(配信.url + '/book.html', { waitUntil: 'load', timeout: 120000 });
  await page.evaluate(() => {
    document.body.classList.remove('exally-locked');
    const ov = document.getElementById('loginOv');
    if (ov) { ov.classList.remove('open'); ov.style.display = 'none'; }
  });
  await page.setInputFiles('#bookFileInput', 材料);
  await page.waitForFunction(() => {
    const ss = window.sheets || [];
    return ss.some((s) => s && s.data && Object.keys(s.data).length > 0);
  }, null, { timeout: 120000 });
  await page.waitForTimeout(600);
  await 道を確かめる(page);

  /* ★画面の 関数で 全部の 字を 作る★（`画面の字を作る` と 同じ 道を 1回の evaluate で） */
  const 出 = await page.evaluate((n) => {
    const i = (window.sheets || []).findIndex((s) => s.name === 'あ');
    window.switchSheet(i);
    const sh = window.sheets[i];
    const 字 = [];
    for (let r = 0; r < n; r++) {
      const cell = sh.data[r + ',0'];
      if (!cell) { 字.push('(マス無し)'); continue; }
      const raw = window._字の元(cell);
      if (window._ゼロを隠すか(cell, raw)) { 字.push(''); continue; }
      if (window._答えは字か(cell)) { 字.push(String(raw == null ? '' : raw)); continue; }
      字.push(String(window.fmtForDisplay(raw, cell.numFmt, window._入る字数(window.cW(0), raw, cell.numFmt))));
    }
    return 字;
  }, 組.length);

  let 合 = 0; const 違 = [];
  組.forEach((x, i) => {
    const got = 出[i];
    const ok = (/^#+$/.test(x.字) && /^#+$/.test(got)) || got === x.字;
    if (ok) 合++; else 違.push(x.式 + ' ' + x.書 + ' 画面=' + JSON.stringify(got) + ' 実Excel=' + JSON.stringify(x.字));
  });
  console.log('      ── 実測 ── 見た ' + 組.length + '組 ／ 合った ' + 合 + ' ／ 違う ' + 違.length);
  違.slice(0, 60).forEach((s) => console.log('         ' + s));   /* ★違いは 60件まで 全部 出す★（8件で 切ると 内訳が 割れない） */
  T('★★和暦・経過時間の 書式の マスが 実Excel の 画面の 字と 同じ★★（' + 合 + '/' + 組.length + '）', 違.length === 0,
    違.length + '組 違う');
  /* ★字の "1e-7" は 台の 道に 乗らない★（★今の 画面の 字が 実Excel と 同じかは 未測定＝別の 話★） */
  const 字の出 = await page.evaluate((r) => {
    const sh = window.sheets[window.activeSheet];
    const cell = sh.data[r + ',0'];
    const raw = window._字の元(cell);
    /* ★画面の 道の 順に★（`_答えは字か` が 先・10-04 夕 に 揃えた） */
    const 字か = window._答えは字か(cell);
    return { 型: typeof raw, 字か: 字か, 画面: 字か ? String(raw) : String(window.fmtForDisplay(raw, cell.numFmt, 30)) };
  }, 字の行);
  console.log('      ── 実測 ── 字の "1e-7"（[s]）... 型 ' + 字の出.型 + ' ／ 字か ' + 字の出.字か + ' ／ 画面 ' + JSON.stringify(字の出.画面));
  T('★字の "1e-7" は 台の 道に 乗らない★（台なら "0"）', 字の出.型 === 'string' && 字の出.画面 !== '0',
    JSON.stringify(字の出));
  /* ★★組 J＝字の マスの 画面の 字（画面の 関数の 道）★★ */
  const J出 = await page.evaluate(([頭, n, 空]) => {
    const sh = window.sheets[window.activeSheet];
    const 作る = (r) => {
      const cell = sh.data[r + ',0'];
      if (!cell) return '(マス無し)';
      const raw = window._字の元(cell);
      if (window._ゼロを隠すか(cell, raw)) return '';
      if (window._答えは字か(cell)) return String(raw == null ? '' : raw);
      return String(window.fmtForDisplay(raw, cell.numFmt, window._入る字数(window.cW(0), raw, cell.numFmt)));
    };
    const 出 = [];
    for (let k = 0; k < n; k++) 出.push(作る(頭 + k));
    return { J: 出, 空: 作る(空) };
  }, [J頭, 組J.length, 空の行]);
  const J違 = [];
  組J.forEach((x, k) => { if (J出.J[k] !== x.字) J違.push(x.式 + ' ' + x.書 + ' 画面=' + JSON.stringify(J出.J[k]) + ' 実Excel=' + JSON.stringify(x.字)); });
  console.log('      ── 実測 ── 組 J（字の マス）... 見た ' + 組J.length + ' ／ 合った ' + (組J.length - J違.length) + ' ／ 空の 字 ' + JSON.stringify(J出.空));
  J違.forEach((s) => console.log('         ' + s));
  T('★★字の マスは 字の まま（組 J ' + (組J.length - J違.length) + '/' + 組J.length + '）★★', J違.length === 0, J違.length + '組 違う');
  T('★空の 字の マスは 空★', J出.空 === '', JSON.stringify(J出.空));
} finally {
  await page.close().catch(() => {});
  await browser.close().catch(() => {});
  配信.閉じる();
}

console.log('');
console.log('  ★道★ ' + 道の字);
console.log('wareki-keika-gamen: ' + pass + ' 緑 / ' + fail + ' 赤');
process.exit(fail ? 1 : 0);
