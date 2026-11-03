/* hakaru-retsu-no-ten.mjs ･･･ ★本を 開いた 時の 列の 点（本番の cW）が 実Excel の px と 同じか★ 2026-10-04 夜
 *
 *  ★★なぜ★★
 *    列の 点は「一字の幅（本の 既定の 字体の 0）× ファイルの width」。一字の幅を ★お客さんの 機械で 測る★ので、
 *    Linux の WebKit では 74→93点に なり #### が 出なかった（§18）。直し①a＝★一字の幅を 実Excel の 数の 表から 引く★。
 *    その 受け入れ＝★本を 開いて 本番の cW を 読み、実Excel が 出した px と 1列ずつ 比べる★。
 *  ★★物差し★★ 経営者の 道具135（toru-tsukasa-no-hon-no-retsu-px.ps1）が 出す tsv
 *    「板(1から) ／ 列(0から) ／ 列幅(字) ／ px ／ 隠し」（★司さんの 本の 物は repo に 入れない＝公開 repo★）
 *  ★走らせ方★ node docs/measured/hakaru-retsu-no-ten.mjs --本 <xlsb/xlsx> --物差し <tsv> [--許し 0.5]
 *    本は %TEMP% に 写して 開く（★元の 本は 1バイトも 書かない★）。出すのは ★数と 違った 列の 番号・点だけ★（中身の 字は 出さない）
 */
import path from 'node:path'; import http from 'node:http'; import fs from 'node:fs'; import os from 'node:os';
import { fileURLToPath } from 'node:url';
import { borrow, launch } from '../../scripts/_borrow-playwright.mjs';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..', '..');
const 引数 = process.argv.slice(2);
const 取る = (名, 既定) => { const i = 引数.indexOf('--' + 名); return i >= 0 ? 引数[i + 1] : 既定; };
const 本 = 取る('本', null), 物差し = 取る('物差し', null), 許し = Number(取る('許し', '0.5'));
if (!本 || !物差し || !fs.existsSync(本) || !fs.existsSync(物差し)) { console.log('★--本 と --物差し を 渡して ください（在る 物）★'); process.exit(2); }

const L = fs.readFileSync(物差し, 'utf8').split(/\r?\n/).filter((l) => l && l.charAt(0) !== '#');
const 頭 = L.shift().split('\t');
const ix = (名) => { const i = 頭.indexOf(名); if (i < 0) { console.log('★物差しに 列が 無い★ ' + 名); process.exit(2); } return i; };
const i板 = ix('板(1から)'), i列 = ix('列(0から)'), i字 = ix('列幅(字)'), ipx = ix('px'), i隠 = ix('隠し');
const 列たち = L.map((l) => l.split('\t')).map((c) => ({ 板: Number(c[i板]) - 1, 列: Number(c[i列]), 字: Number(c[i字]), px: Number(c[ipx]), 隠: c[i隠] === 'True' }));

const 写し = path.join(os.tmpdir(), 'exally-retsu-no-ten' + path.extname(本));
fs.copyFileSync(本, 写し);
const 前の時刻 = fs.statSync(本).mtimeMs;

const 型 = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8', '.png': 'image/png', '.svg': 'image/svg+xml', '.ico': 'image/x-icon' };
const s = http.createServer((q, r) => {
  const f = path.join(ROOT, decodeURIComponent(String(q.url).split('?')[0]).replace(/^\/+/, ''));
  if (!f.startsWith(ROOT) || !fs.existsSync(f) || fs.statSync(f).isDirectory()) { r.statusCode = 404; return r.end('no'); }
  r.setHeader('content-type', 型[path.extname(f).toLowerCase()] || 'application/octet-stream');
  fs.createReadStream(f).pipe(r);
});
await new Promise((x) => s.listen(0, '127.0.0.1', x));
const wk = await borrow('retsu-no-ten', 'webkit');
const br = await launch('retsu-no-ten', wk, {}, 'webkit');
const p = await br.newPage({ viewport: { width: 1600, height: 900 } });
let 終わり値 = 0;
try {
  await p.goto('http://127.0.0.1:' + s.address().port + '/book.html', { waitUntil: 'load', timeout: 120000 });
  await p.evaluate(() => { document.body.classList.remove('exally-locked'); const o = document.getElementById('loginOv'); if (o) o.style.display = 'none'; });
  await p.setInputFiles('#bookFileInput', 写し);
  await p.waitForFunction(() => (window.sheets || []).some((x) => Object.keys(x.data || {}).length), null, { timeout: 300000 });
  await p.waitForTimeout(1500);
  const 出 = await p.evaluate((列たち) => {
    if (typeof window.cW !== 'function' || typeof window.switchSheet !== 'function') return { 誤り: '道が 無い（cW / switchSheet）' };
    const 前 = window.activeSheet;
    const 点 = 列たち.map((x) => {
      if (!window.sheets[x.板]) return null;
      window.activeSheet = x.板;   /* ★cW は 今の 板を 見る★（描かせない＝switchSheet は 呼ばない） */
      return window.cW(x.列) / (window.scale || 1);
    });
    window.activeSheet = 前;
    return { 点, 板の数: window.sheets.length, 拡大: window.scale };
  }, 列たち);
  if (出.誤り) { console.log('★' + 出.誤り + '★'); process.exit(8); }
  let 合 = 0; const 違 = []; const 差の箱 = {};
  列たち.forEach((x, k) => {
    const g = 出.点[k];
    if (g === null) { 違.push('板' + (x.板 + 1) + ' 列' + x.列 + ' 板が 無い'); return; }
    const 物 = x.隠 ? 0 : x.px;
    const d = g - 物;
    差の箱[d.toFixed(1)] = (差の箱[d.toFixed(1)] || 0) + 1;
    if (Math.abs(d) <= 許し) 合++; else 違.push('板' + (x.板 + 1) + ' 列' + x.列 + ' 字' + x.字 + ' 実Excel ' + 物 + 'px ／ うち ' + g);
  });
  console.log('[列の 点] 物差し ' + 列たち.length + '列 ／ 板 ' + 出.板の数 + '枚 ／ 拡大 ' + 出.拡大);
  console.log('  ★合う（±' + 許し + 'px）' + 合 + '/' + 列たち.length + '★');
  console.log('  差（うち－実Excel）の 分かれ ' + Object.keys(差の箱).sort((a, b) => a - b).map((k) => k + ':' + 差の箱[k]).join(' ／ '));
  違.slice(0, 40).forEach((t) => console.log('     ' + t));
  if (違.length > 40) console.log('     （残り ' + (違.length - 40) + '列は 同じ 形か 上の 分かれを 見る）');
  終わり値 = 合 === 列たち.length ? 0 : 1;
} finally { await p.close(); await br.close(); s.close(); try { fs.unlinkSync(写し); } catch (e) { /* 消せない 時は 残る */ } }
if (fs.statSync(本).mtimeMs !== 前の時刻) { console.log('★元の 本の 時刻が 変わった★'); process.exit(9); }
process.exit(終わり値);
