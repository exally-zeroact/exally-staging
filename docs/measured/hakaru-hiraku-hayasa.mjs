/* hakaru-hiraku-hayasa.mjs — ★本が 開くまでの 秒★ 2026-09-25
 *
 *  ★★なぜ 在るか★★
 *    経営者1 が 「2分の 正体」を 割りました（`docs/measured/hiraku-made-2fun-no-shoutai.md`）。
 *      CPU の 記録 ... ★87.4% が `lib/shiki-hyou.js` 1本★
 *      `名から番地` ... ★141,920,112 回★／★別々の 字は 53,365 種類★
 *    ⇒★覚えて おくだけで 59,471 ms → 32,078 ms（54%）★（向こうの 実測・chromium）
 *    ⇒★入れるのは 私の 持ち場★なので ★私の 台でも 前後を 測ります★。
 *
 *  ★★測り方の 決め 6つ★★（経営者1 と 2人で 決めました）
 *    ①★1回目を 物差しに しない★（★毎回 1回目が 一番 遅い★＝台が 温まって いない）
 *    ②★3回以上 回して 中ほどで 比べる★
 *    ③★同じ 台で 前後★（webkit 同士／chromium 同士）
 *    ④★★同じ 時間帯で 前後★★（2026-09-25・★私が 危うく 外す 所でした★）
 *        ＝同じ 木・同じ 材料でも ★22.4秒の 回と 24.4秒の 回★ が 在りました。
 *        ＝別の 時間帯の 数と 比べて 「遅く なった」と 読みかけました。
 *    ⑤★揺れ幅より 小さい 差を 「効いた」と 言わない★
 *    ⑥★★前と 後を 交互に 回す★★（経営者1 の 足し）
 *        ＝★時間帯の ずれを 2つで 分け合えます★＝★④を 機械が 守ります★
 *        ⇒`--木 <前の木> --木 <後の木>` で ★1回ずつ 交互に★ 回します。
 *
 *  ★★どこを 測るか★★
 *    `#bookFileInput` に 渡してから ★どれか 1枚でも 中身が 入るまで★。
 *    ＝★`activeSheet` の 1枚だけを 見ない★（経営者1 の 指摘・2026-09-24）
 *    ★断りが 出たら すぐ 止めます★（880秒 待って 嘘を 出さない・2026-09-25）
 *
 *  ★1バイトも 書きません★（保存の 窓は 出しません）
 *
 *  走らせ方:
 *    node docs/measured/hakaru-hiraku-hayasa.mjs <材料> [--回 3] [--台 webkit|chromium] [--札 なまえ]
 *    node docs/measured/hakaru-hiraku-hayasa.mjs <材料> --回 3 --木 <前の木> --木 <後の木>
 *      ⇒★交互に 回して 2つの 中ほどを 並べます★
 */
import path from 'node:path'; import http from 'node:http'; import fs from 'node:fs';
import crypto from 'node:crypto';
import { fileURLToPath } from 'node:url';
import { borrow, launch } from '../../scripts/_borrow-playwright.mjs';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '../..');
const 引 = process.argv.slice(2);
const 取る = (名, 既定) => {
  const i = 引.indexOf(名);
  return (i >= 0 && 引[i + 1]) ? 引[i + 1] : 既定;
};
const 材料 = 引.find((x) => x.charAt(0) !== '-' && 引[引.indexOf(x) - 1] !== '--回'
  && 引[引.indexOf(x) - 1] !== '--台' && 引[引.indexOf(x) - 1] !== '--札')
  || path.join(ROOT, 'tests/fixtures/cross-sheet-sample.xlsb');
const 回 = Number(取る('--回', '3'));
const 台 = 取る('--台', 'webkit');
const 札 = 取る('--札', '(札なし)');
/* ★★木を 2つ 渡せます★★＝★交互に 回して 時間帯の ずれを 分け合う★ */
const 木たち = [];
for (let i = 0; i < 引.length; i++) if (引[i] === '--木' && 引[i + 1]) 木たち.push(引[i + 1]);
if (!木たち.length) 木たち.push(ROOT);

function 立てる(root0) {
  /* ★★道を 揃えます★★（2026-09-25 ここで 1回 踏みました）
       `--木` に `C:/Users/...`（斜線）を 渡すと
       `path.join` が 作る `C:\Users\...`（逆斜線）と ★頭が 合いません★。
       ⇒`f.startsWith(root)` が いつも 偽 ⇒★全部 404★
       ⇒画面は 出るが `#bookFileInput` が 無い ⇒★『30秒 待って 見つからない』★
       ⇒★落ちる ので まだ 良い★（黙って 別の 物を 測るより ずっと 良い） */
  const root = path.resolve(root0);
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
const 本の印 = crypto.createHash('sha256')
  .update(fs.readFileSync(path.join(ROOT, 'lib/shiki-hyou.js'))).digest('hex').slice(0, 16);
console.log('★本が 開くまでの 秒★  札 ' + 札 + ' ／ 台 ' + 台 + ' ／ ' + 回 + '回');
console.log('  材料 ... ' + path.basename(材料) + '（' + fs.statSync(材料).size.toLocaleString() + ' バイト）');
console.log('  ★`lib/shiki-hyou.js` の 印★ sha256 ' + 本の印);

for (const k of 木たち) {
  if (!fs.existsSync(path.join(k, 'book.html'))) {
    console.log('★木に `book.html` が 有りません★ ' + k); process.exit(1);
  }
  console.log('  木 ' + k + ' ／ `lib/shiki-hyou.js` の 印 sha256 '
    + crypto.createHash('sha256').update(fs.readFileSync(path.join(k, 'lib/shiki-hyou.js')))
      .digest('hex').slice(0, 16));
}

const 種 = await borrow('hiraku-hayasa', 台);
const browser = await launch('hiraku-hayasa', 種, {}, 台);
const 配信たち = [];
for (const k of 木たち) 配信たち.push(await 立てる(k));
const 出た = 木たち.map(() => []);
let マス = 0, 板 = 0;

try {
  for (let i = 0; i < 回; i++) {
  for (let t = 0; t < 木たち.length; t++) {          /* ★交互に 回す★ */
    const page = await browser.newPage({ viewport: { width: 1400, height: 900 } });
    try {
      /* ★★「効いて いるか」を 画面の 声で 確かめます★★（2026-09-25）
           ＝★差が 出ない 時に 「効かなかった」と 「そもそも 効いて いない」は 別★
           ＝記憶「壊したのに 赤に ならない は まず 壊れて いるかを 見る」 */
      page.on('console', (m) => {
        const t2 = String(m.text());
        if (t2.indexOf('[Exally] 開いた') >= 0 || t2.indexOf('[Exally] 計算を') >= 0) {
          console.log('      [木' + (t + 1) + 'の 声] ' + t2.slice(0, 160));
        }
      });
      await page.goto(配信たち[t].url + '/book.html', { waitUntil: 'load', timeout: 180000 });
      await page.evaluate(() => {
        document.body.classList.remove('exally-locked');
        const ov = document.getElementById('loginOv');
        if (ov) { ov.classList.remove('open'); ov.style.display = 'none'; }
        window.FileOut = { deliver: () => Promise.reject(new Error('★書き出しては いけません★')) };
        window.__t0 = performance.now();
      });
      await page.setInputFiles('#bookFileInput', 材料);
      await page.waitForFunction(() => {
        const ss = window.sheets || [];
        for (let k = 0; k < ss.length; k++) {
          if (ss[k] && ss[k].data && Object.keys(ss[k].data).length > 0) return true;
        }
        const el = document.getElementById('toast');
        const t = el ? String(el.textContent || '') : '';
        if (t.indexOf('読めませんでした') >= 0 || t.indexOf('開けませんでした') >= 0) throw new Error(t.slice(0, 120));
        return false;
      }, null, { timeout: 900000 });
      const n = await page.evaluate(() => {
        const ss = window.sheets || [];
        let m = 0;
        for (let k = 0; k < ss.length; k++) m += Object.keys((ss[k] || {}).data || {}).length;
        return { 秒: performance.now() - window.__t0, マス: m, 板: ss.length };
      });
      出た[t].push(Math.round(n.秒)); マス = n.マス; 板 = n.板;
      console.log('    ' + (i + 1) + '回目 ／ 木' + (t + 1) + ' ... '
        + Math.round(n.秒).toLocaleString() + ' ms'
        + (i === 0 ? '  ★1回目は 物差しに しません★（台が 温まって いない）' : ''));
    } finally {
      await page.close().catch(() => {});
    }
  }
  }
} finally {
  await browser.close().catch(() => {});
  for (const h of 配信たち) h.閉じる();
}

console.log('');
console.log('  板 ' + 板 + '枚 ／ マス ' + マス.toLocaleString());
const 中ほど = [];
for (let t = 0; t < 木たち.length; t++) {
  const 並 = 出た[t].slice().sort((a, b) => a - b);
  中ほど.push(並[Math.floor(並.length / 2)]);
  console.log('  木' + (t + 1) + ' ' + 木たち[t]);
  console.log('    ★中ほど ' + 中ほど[t].toLocaleString() + ' ms★'
    + ' ／ 一番速い ' + 並[0].toLocaleString()
    + ' ／ 一番遅い ' + 並[並.length - 1].toLocaleString()
    + ' ／ ★揺れ幅 ' + (並[並.length - 1] - 並[0]).toLocaleString() + ' ms★');
}
if (木たち.length === 2) {
  const 差 = 中ほど[0] - 中ほど[1];
  const 幅 = Math.max(
    Math.max.apply(null, 出た[0]) - Math.min.apply(null, 出た[0]),
    Math.max.apply(null, 出た[1]) - Math.min.apply(null, 出た[1]));
  console.log('');
  console.log('  ★差 ' + 差.toLocaleString() + ' ms★'
    + ' ／ ★一番 大きい 揺れ幅 ' + 幅.toLocaleString() + ' ms★');
  console.log('  ⇒' + (Math.abs(差) > 幅
    ? '★★揺れ幅より 大きい＝本物の 差です★★'
    : '★★揺れ幅の 中＝差が 在るとは 言えません★★'));
}
console.log('  ★揺れ幅より 小さい 差を 「効いた」と 言いません★');
