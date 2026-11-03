/* hakaru-doko-ga-omoi-ka.mjs
 *   -- ★111秒 固まって いる ── ★どの 関数か★★（114）（2026-09-25）
 *
 *  ★★なぜ この やり方に したか★★
 *    ★外から 1秒ごとに 観る★ やり方では ★何も 分かりませんでした★
 *    ＝節目 14個が ★全部 同じ 時刻（111,177 ms）★に 着きました
 *    ＝★私の 問い合わせ自体が 1回も 動けて いない★
 *    ＝★画面が 111秒 ずっと 1本の 処理で 塞がって います★（★細かい 段では ない★）
 *    ⇒★外から 観る 道は 塞がって います★
 *    ⇒★中で 誰が 使って いるかを 記録させます★（CPU の 記録）
 *
 *  ★★chromium を 使います★★
 *    ＝★記録の 口（CDP）は chromium に しか 在りません★
 *    ＝★webkit と 時間が 違うかも しれません★＝★それも 一緒に 出します★
 *    ＝★見たいのは 「どこが 重いか」★で ★秒数そのものでは ありません★
 *
 *  ★★出さない 物★★ ... ★マスの 値／シートの 名／式の 字★
 *  ★★1バイトも 書きません★★
 *
 *  使い方: node docs/measured/hakaru-doko-ga-omoi-ka.mjs "<xlsb>" [--秒 300]
 */
import fs from 'node:fs';
import path from 'node:path';
import http from 'node:http';
import crypto from 'node:crypto';
import { fileURLToPath } from 'node:url';
import { borrow, launch } from '../../scripts/_borrow-playwright.mjs';

const ここ = path.dirname(fileURLToPath(import.meta.url));
/* ★★どの 木を 測るか★★（2026-09-25・追記）
     ★なぜ 要るか★
       ★道具は 私の 木に 在り、測りたい 字は 相手の 木に 在る★事が 在ります
       （09-25 Exally1 の 直しは `origin/karimono-hazushi-dodai5` に 在り main に 無い）
     ⇒★`--元 <置き場>` で ★配る 木★を 指せる ように しました★
     ⇒★指さない 時は 今までと 同じ（道具の 木）★
     ★出しに 必ず どの 木を 測ったかを 書きます★＝[[feedback_doko_wo_kazoeta_ka_mo_kaku]] */
const ROOT = process.argv.includes('--元')
  ? path.resolve(process.argv[process.argv.indexOf('--元') + 1])
  : path.join(ここ, '..', '..');
const 材料 = process.argv[2];
const 秒 = process.argv.includes('--秒') ? Number(process.argv[process.argv.indexOf('--秒') + 1]) : 300;
/* ★★何回 取るか★★（2026-09-25・追記）
     ★なぜ 要るか★
       ★同じ 木・同じ 材料で 2回 取ったら 割合が 大きく 違いました★
         `lib/xlsx.full.min.js` ... ★33.0% → 1.1%★（どちらも 30.9秒）
       ⇒★1回の 記録を そのまま 出しては いけません★
       ⇒★自分が Exally1 に 「3回 回して 中ほどで 比べろ」と 言った のと 同じ事★
     ⇒★何回か 取って ★幅★も 一緒に 出します★ */
const 回 = process.argv.includes('--回') ? Number(process.argv[process.argv.indexOf('--回') + 1]) : 1;
if (!材料 || !fs.existsSync(材料)) { console.log('★材料が 在りません★'); process.exit(3); }
const 中 = fs.readFileSync(材料);
console.log('★材料★ ' + path.basename(材料) + '（' + 中.length.toLocaleString() + ' バイト）');
console.log('  sha256 ' + crypto.createHash('sha256').update(中).digest('hex'));
console.log('★★測る 木★★ ' + ROOT);
try {
  const { execSync } = await import('node:child_process');
  const h = execSync('git -C "' + ROOT + '" rev-parse --short HEAD', { encoding: 'utf8' }).trim();
  const d = execSync('git -C "' + ROOT + '" status --porcelain', { encoding: 'utf8' }).trim();
  console.log('  ★印★ ' + h + (d ? '（★手元に 未commit が ' + d.split('\n').length + '本 在ります★）' : '（手元は 綺麗）'));
} catch (e) { console.log('  ★印を 読めません ... ' + e.message + '★'); }

const 型 = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8', '.json': 'application/json', '.png': 'image/png',
  '.svg': 'image/svg+xml', '.ico': 'image/x-icon' };
const server = http.createServer((req, res) => {
  const u = decodeURIComponent(req.url.split('?')[0]);
  const p = path.join(ROOT, u === '/' ? 'book.html' : u.replace(/^\//, ''));
  if (!p.startsWith(ROOT) || !fs.existsSync(p) || fs.statSync(p).isDirectory()) { res.writeHead(404); res.end(); return; }
  res.writeHead(200, { 'Content-Type': 型[path.extname(p)] || 'application/octet-stream' });
  res.end(fs.readFileSync(p));
});
await new Promise((ok) => server.listen(0, '127.0.0.1', ok));
const もと = 'http://127.0.0.1:' + server.address().port;

const 全回 = [];
for (let 回目 = 1; 回目 <= 回; 回目++) {
const ck = await borrow('doko-ga-omoi', 'chromium');
const browser = await launch('doko-ga-omoi', ck, {}, 'chromium');
const page = await browser.newPage({ viewport: { width: 1280, height: 900 } });
await page.goto(もと + '/book.html', { waitUntil: 'load', timeout: 120000 });
await page.evaluate(() => {
  document.body.classList.remove('exally-locked');
  const ov = document.getElementById('loginOv');
  if (ov) { ov.classList.remove('open'); ov.style.display = 'none'; }
  window.FileOut = { deliver: () => Promise.reject(new Error('★書き出しては いけません★')) };
});

const cdp = await page.context().newCDPSession(page);
await cdp.send('Profiler.enable');
await cdp.send('Profiler.setSamplingInterval', { interval: 2000 });   /* 2ms ごと */
await cdp.send('Profiler.start');
console.log('★記録を 始めました★（CPU）');

const t0 = Date.now();
await page.setInputFiles('#bookFileInput', 材料);
let 開いた = 0;
for (let t = 0; t < 秒; t += 2) {
  await new Promise((r) => setTimeout(r, 2000));
  const い = await page.evaluate(() => {
    const ss = window.sheets || [];
    let m = 0; ss.forEach((sh) => { m += Object.keys((sh && sh.data) || {}).length; });
    return m;
  }).catch(() => -1);
  if (い > 0) { 開いた = Date.now() - t0; break; }
}
const { profile } = await cdp.send('Profiler.stop');
console.log('★★開くまで★★ ' + (開いた ? 開いた + ' ms' : '★開きません★'));

/* ── ★誰が 使ったかを 数えます★ ── */
const ふし = new Map();
profile.nodes.forEach((n) => ふし.set(n.id, n));
const 自分 = new Map();
(profile.samples || []).forEach((id) => 自分.set(id, (自分.get(id) || 0) + 1));
const 全 = (profile.samples || []).length || 1;
const 並 = [...自分.entries()].map(([id, n]) => {
  const x = ふし.get(id) || {};
  const f = x.callFrame || {};
  const 元 = String(f.url || '').replace(もと, '').replace(/^\//, '') || '(中)';
  return { 名: (f.functionName || '(名なし)') + ' @ ' + 元 + ':' + (f.lineNumber + 1), 数: n };
}).sort((a, b) => b.数 - a.数);

/* ★★関数ごと／ファイル別に 足す★★
     ★なぜ★ ... 同じ 関数が ★別の 呼ばれ方★で 何行にも 分かれて 出ます。
                ★上から 15行だけ 見ると 15行の 外が 落ちます★
                （09-25 に 実際に `名から番地` を 手で 足して 取りこぼしました）
     ⇒★機械に 足させます★ */
  const 関数ごと = new Map();
  const 本ごと = new Map();
  並.forEach((x) => {
    const 所 = x.名.split(' @ ')[1];
    関数ごと.set(x.名.split(' @ ')[0] + ' @ ' + 所, (関数ごと.get(x.名.split(' @ ')[0] + ' @ ' + 所) || 0) + x.数);
    const f = 所.split(':')[0];
    本ごと.set(f, (本ごと.get(f) || 0) + x.数);
  });
  全回.push({ ms: 開いた, 関数: 関数ごと, 本: 本ごと, 全: 全 });
  console.log('  ' + 回目 + '回目 ... ' + (開いた ? 開いた.toLocaleString() + ' ms' : '★開きません★')
    + '   （拾った 点 ' + 全.toLocaleString() + '）');
  await browser.close();
}
server.close();

/* ── ★何回ぶんかを まとめて 出します★ ── */
const 中央 = (a) => { const b = a.slice().sort((p, q) => p - q); return b[(b.length - 1) >> 1]; };
console.log('');
console.log('★★開くまで★★ ' + 全回.map((x) => x.ms.toLocaleString()).join(' / ')
  + ' ms（★中ほど ' + 中央(全回.map((x) => x.ms)).toLocaleString() + '★）');
const まとめ = (取る, 札) => {
  const 名々 = new Set();
  全回.forEach((x) => 取る(x).forEach((v, k) => 名々.add(k)));
  const 行 = [...名々].map((k) => {
    const a = 全回.map((x) => (取る(x).get(k) || 0) / x.全 * 100);
    return { k: k, 中: 中央(a), 小: Math.min.apply(null, a), 大: Math.max.apply(null, a) };
  }).sort((p, q) => q.中 - p.中);
  console.log('');
  console.log('★★' + 札 + '★★（' + 回 + '回・★中ほど（一番小さい〜一番大きい）★）');
  行.slice(0, 12).forEach((x) => {
    const 幅 = x.大 - x.小;
    console.log('  ' + x.中.toFixed(1).padStart(5) + '%  （' + x.小.toFixed(1) + '〜' + x.大.toFixed(1) + '%）'
      + (幅 >= 5 ? '  ★幅 ' + 幅.toFixed(1) + '＝1回では 言えません★' : '')
      + '  ' + x.k);
  });
};
まとめ((x) => x.関数, '関数ごと');
まとめ((x) => x.本, 'ファイル別');
console.log('');
console.log('★★言い落とさない 事★★');
console.log('  ・★記録を 取ると 少し 遅く なります★＝★この 秒数を 「直った 後の 目安」に しない★');
console.log('  ・★webkit に この 口は 在りません★＝★割合は chromium だけ／秒数は 両方★');
console.log('  ・★幅が 大きい 行は 1回の 記録では 言えません★（09-25 に 同じ 木で 33.0%→1.1% が 出た）');
