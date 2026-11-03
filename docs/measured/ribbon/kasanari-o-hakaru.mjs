/* kasanari-o-hakaru.mjs — ★リボンの ボタンが 重なっていないかを 数える★（2026-09-07）
 *
 *  ★★数え方★★
 *    ★「在る／見えている」では 分からない★＝★押したら 何が 出るか★を 見る。
 *    ⇒ ボタンの ★真ん中★の 座標で `document.elementFromPoint` を 呼び、
 *      ★出てきた 物が そのボタン自身か★を 見る。
 *      別の 物が 出たら ★押しても 別の ボタンが 動く★＝重なり 1個。
 *
 *  ★★どこを 測るか★★
 *    ・ブラウザ 2つ … ★webkit（＝iPhone・iPad・Mac の Safari）★と ★本物の Chrome★
 *    ・幅 5通り ……… 1280 / 1440 / 1600 / 1920 / 2560
 *    ・タブは ★ホーム★（部品が 一番 多い）
 *    ⇒★★片方だけ 測ると「幅の せい」か「ブラウザの せい」か 分かれない★★
 *
 *  ★外へは 1回も 出ません★（127.0.0.1 だけ）
 *
 *  使い方: node docs/measured/ribbon/kasanari-o-hakaru.mjs [<repoの道>]
 *          （道を 渡すと ★直す前の repo★も 同じ 数え方で 測れる）
 */
import path from 'node:path'; import http from 'node:http'; import fs from 'node:fs';
import { fileURLToPath } from 'node:url';
import { borrow, launch } from '../../../scripts/_borrow-playwright.mjs';

const ここ = path.dirname(fileURLToPath(import.meta.url));
const 既定 = path.join(ここ, '..', '..', '..');
const ROOT = process.argv[2] ? path.resolve(process.argv[2]) : 既定;
const 札 = process.argv[3] || 'ima';
const 幅たち = [1280, 1440, 1600, 1920, 2560];
const 紙 = [];
const 言う = (s) => { 紙.push(s); console.log(s); };

function 立てる(root) {
  const 型 = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8',
    '.css': 'text/css; charset=utf-8', '.json': 'application/json; charset=utf-8',
    '.png': 'image/png', '.svg': 'image/svg+xml' };
  const s = http.createServer((q, r) => {
    const f = path.join(root, decodeURIComponent(String(q.url).split('?')[0]).replace(/^\/+/, ''));
    if (!f.startsWith(root) || !fs.existsSync(f) || fs.statSync(f).isDirectory()) { r.statusCode = 404; return r.end('no'); }
    r.setHeader('content-type', 型[path.extname(f).toLowerCase()] || 'application/octet-stream');
    fs.createReadStream(f).pipe(r);
  });
  return new Promise((x) => s.listen(0, '127.0.0.1', () => x({ url: 'http://127.0.0.1:' + s.address().port, 閉じる: () => s.close() })));
}

const 配信 = await 立てる(ROOT);
言う('# ★リボンの ボタンが 重なっていないかを 数えた★（' + new Date().toISOString().slice(0, 10) + '）');
言う('');
言う('★測った 物★ ' + ROOT);
言う('★数え方★ ボタンの ★真ん中★で elementFromPoint ⇒ ★別の 物が 出たら 重なり★');
言う('★外へは 1回も 出ていません（127.0.0.1 だけ）★');
言う('');

const 結果 = {};
for (const [名, kind, opts] of [['webkit', 'webkit', {}], ['Chrome', 'chromium', { channel: 'chrome' }]]) {
  const t = await borrow('kasanari', kind);
  const b = await launch('kasanari', t, opts, kind);
  結果[名] = {};
  for (const W of 幅たち) {
    const page = await b.newPage({ viewport: { width: W, height: 1000 } });
    await page.goto(配信.url + '/book.html', { waitUntil: 'load' });
    await page.evaluate(() => {
      document.body.classList.remove('exally-locked');
      const ov = document.getElementById('loginOv');
      if (ov) { ov.classList.remove('open'); ov.style.display = 'none'; }
    });
    await page.waitForTimeout(1800);
    const r = await page.evaluate(() => {
      const bs = [...document.querySelectorAll('#ribbon [data-act], #ribbon .rb-item')];
      let 見 = 0, 隠 = [], 画面の外 = 0;
      for (const b of bs) {
        const x = b.getBoundingClientRect();
        if (x.width < 1 || x.height < 1) continue;
        const cx = x.x + x.width / 2, cy = x.y + x.height / 2;
        /* ★★真ん中が 窓の 外に 在る 物は 数から 外す★★（2026-09-07 物差しを 直した）
           リボンの 帯は ★横に すべる★ので、右端の ボタンは
           ★左は 画面の 中／真ん中は 画面の 外★に なる事が 有る。
           その時 elementFromPoint は ★null★を 返すが、それは ★重なり では ない★
           （お客さんは 帯を すべらせて 押す）。
           ⇒★数えずに 別に 数える★＝★0 に 見せる為に 捨てない★ */
        if (cx < 0 || cx > window.innerWidth || cy < 0 || cy > window.innerHeight) { 画面の外++; continue; }
        見++;
        const e = document.elementFromPoint(cx, cy);
        if (e !== b && !b.contains(e)) {
          const 上 = e && e.closest ? e.closest('[data-act],.rb-item') : null;
          隠.push({ 自: b.getAttribute('data-act') || b.title || '', 上に: 上 ? (上.getAttribute('data-act') || 上.title || '') : '?' });
        }
      }
      /* ★組の 箱が 中身より 狭くないか★（重なりの 元） */
      let はみ出し = 0;
      for (const items of document.querySelectorAll('#ribbon .rb-items')) {
        const ib = items.getBoundingClientRect();
        let 右 = ib.left;
        for (const c of items.children) { const cb = c.getBoundingClientRect(); if (cb.width > 0 && cb.right > 右) 右 = cb.right; }
        はみ出し += Math.max(0, Math.round(右 - ib.right));
      }
      return { 見, 隠: 隠.length, 例: 隠.slice(0, 4), はみ出しの合計: はみ出し, 画面の外 };
    });
    結果[名][W] = r;
    言う('  ' + 名.padEnd(7) + ' 幅 ' + String(W).padStart(4)
      + ' … 見えた ' + String(r.見).padStart(3) + '個 ／ ★重なり ' + String(r.隠).padStart(3) + '個★'
      + ' ／ 箱から はみ出した 合計 ' + String(r.はみ出しの合計).padStart(4) + 'px'
      + ' ／ 真ん中が 画面の 外 ' + String(r.画面の外).padStart(2) + '個（帯を すべらせて 押す 物）');
    for (const x of r.例) 言う('        [' + x.自 + '] を 押すと → [' + x.上に + ']');
    if (W === 1920) {
      await page.screenshot({ path: path.join(ここ, 'ribbon-' + 名 + '-1920-' + 札 + '.png') });
      言う('        ★絵★ docs/measured/ribbon/ribbon-' + 名 + '-1920-' + 札 + '.png');
    }
    await page.close();
  }
  await b.close();
}
配信.閉じる();

言う('');
const 合 = (名) => 幅たち.reduce((a, W) => a + 結果[名][W].隠, 0);
言う('★まとめ★ webkit の 重なり 合計 ' + 合('webkit') + '個 ／ Chrome の 重なり 合計 ' + 合('Chrome') + '個'
  + '（幅 ' + 幅たち.length + '通り）');
fs.writeFileSync(path.join(ここ, 'kasanari-' + 札 + '.txt'), 紙.join('\n') + '\n', { encoding: 'utf-8' });
console.log('\n★書いた … docs/measured/ribbon/kasanari-' + 札 + '.txt★');
