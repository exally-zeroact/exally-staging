/* hozon-no-ayamari-wo-yomu-webkit.mjs ･･･ ★答えが 誤りで 保存された 式は 開いた だけで 誤りの 字★ 2026-10-02
 *
 *  ★★何が 起きて いたか（経営者 実測・司さんの 実物・10-02）★★
 *    開いた だけの 画面が 実Excel（Open だけ）と ★69個 違う★
 *    ＝69個 とも ★実Excel は #REF!／うちは 数★
 *    ★因★ ... 本の 記録11（答えが 誤り）69個 とも 0x17＝#REF!（記録を 数えた）
 *      ＝★実Excel は 計算して いない★（保存された 答えが #REF!）
 *      ＝借り物は ★`t:'e'`・`v`＝誤りの 番号 23・`w`＝'#REF!'★ で 渡す
 *      ＝うちは `v` を 持ち ★数の 23★ に 書式を 掛けて いた（`js/book-open.js` sheetToGrid）
 *
 *  ★★この 見張りが 見る 物★★
 *    ⑴★誤りの 7種 とも★ 開いた だけで 誤りの 字（★#NULL! は 番号 0＝落ちやすい★）
 *    ⑵★書式（円・時間・小数）を 通っても 字が 変わらない★
 *    ⑶★計算では なく 読んで いる★＝`=A1*10` に #REF! を 保存 ⇒ 開いた だけで #REF!
 *    ⑷★書き出しても 触らない 誤りの 式は 元の まま★（式も 誤りの 答えも）
 *  ★`.xlsx` で 作ります★＝借り物は `.xlsb` に 式の 記録を 書けません
 *    （★借り物が 渡す 形（`t:'e'`）は `.xlsx` も `.xlsb` も 同じ★＝読む 所は 1つ）
 *    ⇒★`.xlsb` の 書き出しは 実物で 経営者が 見ます（受け入れ ③）★
 *
 *  ★走らせ方★: node tests/hozon-no-ayamari-wo-yomu-webkit.mjs
 */
import path from 'node:path'; import http from 'node:http'; import fs from 'node:fs';
import os from 'node:os';
import { createRequire } from 'node:module';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { borrow, launch } from '../scripts/_borrow-playwright.mjs';
import { 道を確かめる, 道の字, 画面の字を作る } from '../docs/measured/_gamen-no-michi.mjs';

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

/* ══ ★材料★ ══ 1列目＝式／答えは 誤り（★式を 計算しても 同じ 誤り★＝打った 後も 動かない）
     ★最後の 1つだけ★ `=A1*10` に #REF! を 保存（★計算すれば 20＝読んで いるかの 証し★） */
const 誤り = [
  { 番地: 'B1', f: 'A1:A2 B1:B2', 番号: 0x00, 字: '#NULL!', z: 'General' },
  { 番地: 'B2', f: '1/0', 番号: 0x07, 字: '#DIV/0!', z: '#,##0"円"' },
  { 番地: 'B3', f: '"a"+1', 番号: 0x0F, 字: '#VALUE!', z: '0.00' },
  { 番地: 'B4', f: 'B99', 番号: 0x17, 字: '#REF!', z: '#,##0"円"' },
  { 番地: 'B5', f: 'NOSUCHFN()', 番号: 0x1D, 字: '#NAME?', z: '[h]:mm' },
  { 番地: 'B6', f: 'SQRT(-1)', 番号: 0x24, 字: '#NUM!', z: '0%' },
  { 番地: 'B7', f: 'NA()', 番号: 0x2A, 字: '#N/A', z: 'yyyy/m/d' },
  { 番地: 'B8', f: 'A1*10', 番号: 0x17, 字: '#REF!', z: '#,##0"円"', 読みの証し: true },
];
function 作る() {
  const ws = XLSX.utils.aoa_to_sheet([[2, null, 5], [3]]);   /* ★C1＝直す マス（誤りの 式は 見ない）★ */
  for (const e of 誤り) ws[e.番地] = { t: 'e', v: e.番号, f: e.f, z: e.z };
  ws['!ref'] = 'A1:C20';
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'あ');
  const 道 = path.join(os.tmpdir(), 'exally-hozon-no-ayamari.xlsx');
  fs.writeFileSync(道, XLSX.write(wb, { bookType: 'xlsx', type: 'buffer' }));
  /* ★材料が 本当に 誤りで 保存されて いるか★（★空振りの 緑を 作らない★） */
  const re = XLSX.read(fs.readFileSync(道), { type: 'buffer', cellFormula: true, cellNF: true });
  const 外れ = 誤り.filter((e) => { const c = re.Sheets['あ'][e.番地]; return !c || c.t !== 'e' || c.v !== e.番号 || !c.f; });
  if (外れ.length) { console.log('  ★★材料が 作れて いません＝空振り★★ ' + 外れ.map((e) => e.番地).join(',')); process.exit(8); }
  return 道;
}

console.log('[hozon-no-ayamari-wo-yomu] ★答えが 誤りで 保存された 式は 開いた だけで 誤りの 字★');
const 材料 = 作る();
const wk = await borrow('hozon-no-ayamari', 'webkit');
const browser = await launch('hozon-no-ayamari', wk, {}, 'webkit');
const page = await browser.newPage({ viewport: { width: 1280, height: 900 }, acceptDownloads: true });
const 配信 = await 立てる(ROOT);
const 落とし先 = fs.mkdtempSync(path.join(os.tmpdir(), 'exally-ayamari-'));

try {
  page.on('pageerror', (e) => console.log('      [落ちた] ' + String(e.message).slice(0, 200)));
  let 声 = '';
  page.on('console', (m) => { const t = String(m.text()); if (t.indexOf('開いた 直後の 計算') >= 0) 声 = t; });
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
  console.log('      [画面の 声] ' + 声.slice(0, 170));
  /* ★★開いた 直後に 計算して いたら ⑶ が 意味を 失います★★＝★段を 揃える★ */
  const 計算しない = await page.evaluate(() => window._開いた直後は計算しない === true);
  T('★段★ 開いた 直後は 計算して いない（旗 true）', 計算しない, '旗が true で ない＝⑴⑶は 計算の 字を 見て います');

  for (const e of 誤り) {
    const rc = XLSX.utils.decode_cell(e.番地);
    const 出 = await 画面の字を作る(page, 'あ', rc.r + ',' + rc.c);
    const 字 = String(出 && 出.字);
    T((e.読みの証し ? '★⑶ 計算でなく 読む★ ' : '★⑴⑵★ ') + e.番地 + ' 番号 0x' + e.番号.toString(16)
      + '（書式 ' + e.z + '）⇒ ' + e.字, 字 === e.字, '出た=' + JSON.stringify(字));
  }

  /* ══ ★⑷ 書き出し★ ══（★誤りの 式に つられない マスを 1つ 直す★＝お客さんの 道） */
  await page.evaluate(() => { window.setCell(0, 2, '7'); });   /* C1（★元から 在る マス＝`.xlsx` の 空きマスは 別の 制限★） */
  await page.waitForTimeout(1200);
  const 落とし物 = page.waitForEvent('download', { timeout: 120000 }).catch(() => null);
  await page.evaluate(() => { window.saveOpenedBook(); });
  await page.waitForTimeout(1200);
  const 窓 = await page.evaluate(() => {
    const go = document.getElementById('diffGo');
    return go && go.offsetWidth > 0 ? String(go.textContent || '').trim() : '';
  });
  console.log('      ── 実測 ── 窓の ボタン ... ' + JSON.stringify(窓));
  if (!窓) { console.log('  ★★窓が 出て いません＝⑷は 空振りです★★'); fail++; }
  else {
    await page.click('#diffGo');
    const dl = await 落とし物;
    const 帯 = await page.evaluate(() => {
      const t = document.querySelector('.toast, #toast, [class*="toast"]');
      return t ? String(t.textContent || '').trim().slice(0, 200) : '';
    });
    T('★⑷ 書き出せる★', !!dl, '帯 ' + JSON.stringify(帯));
    if (dl) {
      const 出た = path.join(落とし先, 'out.xlsx');
      await dl.saveAs(出た);
      const wb2 = XLSX.read(fs.readFileSync(出た), { type: 'buffer', cellFormula: true });
      const ws2 = wb2.Sheets['あ'];
      T('★⑷ 直した マスは 入った（C1=7）★', ws2 && ws2.C1 && String(ws2.C1.v) === '7',
        'C1=' + JSON.stringify(ws2 && ws2.C1));
      /* ★B8 は A1 に つられる＝打った 後に 計算されて 20 が 正しい（触らない 物では ない）★ */
      for (const e of 誤り.filter((x) => !x.読みの証し)) {
        const c = ws2 && ws2[e.番地];
        T('★⑷ 触らない 誤りの 式は 元の まま★ ' + e.番地 + ' ' + e.字,
          !!c && c.t === 'e' && c.v === e.番号 && String(c.f || '') === e.f,
          '出た=' + JSON.stringify(c));
      }
    }
  }
} finally {
  await page.close().catch(() => {});
  await browser.close().catch(() => {});
  配信.閉じる();
}

console.log('');
console.log('  ★道★ ' + 道の字);
console.log('hozon-no-ayamari-wo-yomu: ' + pass + ' 緑 / ' + fail + ' 赤');
process.exit(fail ? 1 : 0);
