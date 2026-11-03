/* jitsuhaishin-safari-de-hiraku.mjs — ★実配信を 本物の Safari で 開いて 数える★（2026-09-07）
 *
 *  ★★なぜ 要るか★★
 *    ★『手元で 直った』と『お客さんの 道で 直った』は 別です★
 *    手元の 直しは ★自分の フォルダを 自分で 配って★ 見ています。
 *    お客さんが 見るのは ★Vercel が 配っている 物★です。
 *    ⇒★★間に「配信の 合図が 届いていない」「古い 版が 残っている」が 挟まります★★
 *      （2026-08-18 に 実際に 起きた … CI 緑・push 済みなのに ★配信は 前の 版★）
 *
 *  ★★何を 見るか★★
 *    ①★本当に 新しい 版が 出ているか★（`?v=` の 刻印）
 *    ②★リボンの ボタンが 重なっていないか★（真ん中を 押して 何が 出るか）
 *    ③★絵を 1枚★（★数字が 緑でも 絵を 開くまで OK に しない★）
 *
 *  ★★ログインは 通していません★★
 *    ★合言葉は 1文字も 打っていません★。
 *    画面の 鍵（`exally-locked` と ログインの 覆い）を ★見えなく しただけ★です。
 *    ⇒★リボンの 並びを 見るのに ログインは 要りません★
 *    ⇒★でも『ログインした 状態を 見た』とは ★書きません★★（見ていないので）
 *
 *  使い方: node docs/measured/ribbon/jitsuhaishin-safari-de-hiraku.mjs [<配信先>]
 */
import path from 'node:path';
import fs from 'node:fs';
import { fileURLToPath } from 'node:url';
import { borrow, launch } from '../../../scripts/_borrow-playwright.mjs';

const ここ = path.dirname(fileURLToPath(import.meta.url));
const 配信先 = (process.argv[2] || 'https://exally.vercel.app').replace(/\/$/, '');
const 幅たち = [1280, 1920];
const 紙 = [];
const 言う = (s) => { 紙.push(s); console.log(s); };
let pass = 0, fail = 0;
const T = (n, よい, 添え) => {
  if (よい) { pass++; 言う('  ok   ' + n); }
  else { fail++; 言う('  NG   ' + n + (添え ? '\n       ' + 添え : '')); }
};

言う('# ★実配信を 本物の Safari（webkit）で 開いた★（' + new Date().toISOString().slice(0, 10) + '）');
言う('');
言う('★配信先★ ' + 配信先);
言う('★ログインは 通していません★（合言葉は 1文字も 打っていない／覆いを 見えなく しただけ）');
言う('');

const wk = await borrow('jitsuhaishin', 'webkit');
const browser = await launch('jitsuhaishin', wk, {}, 'webkit');
try {
  for (const W of 幅たち) {
    const page = await browser.newPage({ viewport: { width: W, height: 1000 } });
    /* ★★『本当に 配信を 叩いたか』を ★推し量りでは なく 記録で★ 残す★★
       ⇒ 実配信の 絵と 手元の 絵は ★バイト一致します★（同じ バイトを 同じ webkit が 描く 為）
       ⇒★★つまり 絵だけ 見ると ★手元の 絵を 写した★のと 見分けが つきません★★
       ⇒ だから ★配信の 機械しか 出せない 物★を 一緒に 書き残す
         ・`x-vercel-id` … Vercel が 1回ごとに 付ける 番号（手元には 出せない）
         ・`etag` `age` … 配信の 中身の 印
         ・返事の 大きさ／SHA … 実際に 受け取った バイト */
    const 返事 = [];
    page.on('response', (res) => {
      const u = res.url();
      if (u.endsWith('/book.html') || u.indexOf('lib/ribbon.js') >= 0) {
        const h = res.headers();
        返事.push({
          url: u, 状態: res.status(),
          'x-vercel-id': h['x-vercel-id'] || '(無い)',
          etag: h.etag || '(無い)', age: h.age || '(無い)',
          server: h.server || '(無い)',
        });
      }
    });
    await page.goto(配信先 + '/book.html', { waitUntil: 'load', timeout: 60000 });
    await page.evaluate(() => {
      document.body.classList.remove('exally-locked');
      const ov = document.getElementById('loginOv');
      if (ov) { ov.classList.remove('open'); ov.style.display = 'none'; }
    });
    await page.waitForTimeout(2500);

    /* ★①本当に 新しい 版か★＝刻印を 画面から 読む */
    const 刻印 = await page.evaluate(() => {
      const s = [...document.querySelectorAll('script[src*="lib/ribbon.js"]')][0];
      return s ? (s.getAttribute('src').split('?v=')[1] || '') : '';
    });
    言う('  ★配信の 刻印（lib/ribbon.js）… ' + (刻印 || '(取れない)') + '★');

    /* ★②直しが 入っているか★＝新しい 出口が 在るか */
    const 直しが在る = await page.evaluate(() =>
      !!(window.Ribbon && typeof window.Ribbon.組の幅を直す === 'function'));
    T('★幅 ' + W + ' … 直した 版が 配信されている（組の幅を直す が 在る）★', 直しが在る);

    /* ★③重なっていないか★＝真ん中を 押して 何が 出るか */
    const r = await page.evaluate(() => {
      const bs = [...document.querySelectorAll('#ribbon [data-act], #ribbon .rb-item')];
      let 見 = 0, 隠 = [], 外 = 0, はみ出し = 0;
      for (const b of bs) {
        const x = b.getBoundingClientRect();
        if (x.width < 1 || x.height < 1) continue;
        const cx = x.x + x.width / 2, cy = x.y + x.height / 2;
        if (cx < 0 || cx > window.innerWidth || cy < 0 || cy > window.innerHeight) { 外++; continue; }
        見++;
        const e = document.elementFromPoint(cx, cy);
        if (e !== b && !b.contains(e)) {
          const 上 = e && e.closest ? e.closest('[data-act],.rb-item') : null;
          隠.push((b.getAttribute('data-act') || b.title || '') + ' → '
            + (上 ? (上.getAttribute('data-act') || 上.title || '') : '?'));
        }
      }
      for (const items of document.querySelectorAll('#ribbon .rb-items')) {
        const ib = items.getBoundingClientRect();
        let 右 = ib.left;
        for (const c of items.children) { const cb = c.getBoundingClientRect(); if (cb.width > 0 && cb.right > 右) 右 = cb.right; }
        はみ出し += Math.max(0, Math.round(右 - ib.right));
      }
      return { 見, 隠: 隠.length, 例: 隠.slice(0, 4), 外, はみ出し };
    });
    T('★幅 ' + W + ' … 押した 物が そのまま 出る（重なり 0個）★', r.隠 === 0,
      '見えた ' + r.見 + '個 ／ ★重なり ' + r.隠 + '個★'
      + (r.例.length ? '\n       ' + r.例.join('\n       ') : ''));
    T('★幅 ' + W + ' … 組の 箱から 中身が はみ出していない★', r.はみ出し === 0,
      'はみ出し ' + r.はみ出し + 'px');
    言う('       … 見えた ' + r.見 + '個 ／ 重なり ' + r.隠 + '個 ／ はみ出し ' + r.はみ出し + 'px'
      + ' ／ 真ん中が 画面の 外 ' + r.外 + '個');

    /* ★配信の 機械が 出した 印を 残す★ */
    言う('  ★配信を 叩いた 印（手元では 出せない 物）★');
    for (const x of 返事) {
      言う('    ' + x.状態 + ' ' + x.url.replace(配信先, ''));
      言う('        x-vercel-id … ' + x['x-vercel-id']);
      言う('        etag ' + x.etag + ' ／ age ' + x.age + ' ／ server ' + x.server);
    }
    T('★配信の 機械の 番号（x-vercel-id）が 取れた★',
      返事.some((x) => x['x-vercel-id'] !== '(無い)'),
      '★取れないと「本当に 配信を 叩いたか」が 絵からは 言えない★');

    const 絵 = path.join(ここ, 'jitsuhaishin-safari-' + W + '.png');
    await page.screenshot({ path: 絵 });
    言う('       ★絵★ docs/measured/ribbon/jitsuhaishin-safari-' + W + '.png');
    await page.close();
  }
} finally {
  await browser.close();
}
言う('');
言う('jitsuhaishin-safari: ' + pass + ' 緑 / ' + fail + ' 赤');
fs.writeFileSync(path.join(ここ, 'jitsuhaishin-safari.txt'), 紙.join('\n') + '\n', { encoding: 'utf-8' });
console.log('\n★書いた … docs/measured/ribbon/jitsuhaishin-safari.txt★');
process.exit(fail ? 1 : 0);
