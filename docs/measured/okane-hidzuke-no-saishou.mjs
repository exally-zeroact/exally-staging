/* okane-hidzuke-no-saishou.mjs -- ★お金の 関数に 日付を 渡す 一番 小さい 形★（2026-09-18）
 *
 *  ★★これが 729本の 元でした★★
 *    `=COUPDAYS(DATE(2008,11,11),DATE(2021,3,1),2,0)` ... ★#VALUE!★
 *    `=COUPDAYS(39763,44256,2,0)` .................... ★180（合う）★
 *    ⇒★同じ 日・同じ 関数・★渡し方だけ★ 違う★
 *    ⇒★因は `lib/formula-kane-plug.js` の `数()`★
 *      借り物は 日付を `{ val: 39763, format: ... }` で 渡す
 *      ⇒`Number()` が NaN ⇒ `#VALUE!`
 *    ⇒★直し（`04f2555`）の 後は 5本とも 合います★
 *
 *  ★★なぜ 残すか★★
 *    ★一番 安い 切り分け★です（★1本だけ 形を 変えて 押す★）。
 *    ★同じ 型は また 起きます★（お金・割合・時刻の 包みも `val` を 持つ）。
 *
 *  ★★画面で 押して います★★（node の 台では ありません）
 *    ⇒`--どこ=<URL>` で テスト版でも 本番でも 押せます（★既定は テスト版★）
 *
 *  使い方:
 *    node docs/measured/okane-hidzuke-no-saishou.mjs
 *    node docs/measured/okane-hidzuke-no-saishou.mjs <URL>
 */
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

/* ★手元の 絶対の 道を 焼き込まない★（記憶の 決まり） */
const ここ = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(ここ, '..', '..');
const { borrow, launch } = await import(pathToFileURL(path.join(ROOT, 'scripts/_borrow-playwright.mjs')).href);

const URL = process.argv[2] || 'https://exally-git-karimono-hazushi-dodai5-exallysupoort-8848s-projects.vercel.app/book.html';
const 式たち = [
  ['=COUPDAYS(DATE(2008,11,11),DATE(2021,3,1),2,0)', '180'],
  ['=COUPNUM(DATE(2008,11,11),DATE(2021,3,1),2,0)', '25'],
  ['=ACCRINT(DATE(2008,3,1),DATE(2008,8,31),DATE(2008,5,1),0.1,1000,2,0)', '16.666666666666664'],
  ['=COUPDAYS(39763,44256,2,0)', '180（★数の 通し番号で★）'],
  ['=SUM(1,2,3)', '6（★対照★）'],
];

const wk = await borrow('okane4', 'webkit');
const browser = await launch('okane4', wk, {}, 'webkit');
const page = await browser.newPage({ viewport: { width: 1280, height: 800 } });
try {
  const 返 = await page.goto(URL, { waitUntil: 'load', timeout: 60000 });
  console.log('http ' + 返.status());
  await page.evaluate(() => {
    document.body.classList.remove('exally-locked');
    const ov = document.getElementById('loginOv');
    if (ov) { ov.classList.remove('open'); ov.style.display = 'none'; }
  });
  const 出 = await page.evaluate((問) => {
    const sh = window.sheets[window.activeSheet];
    return 問.map((f, i) => {
      window.setCell(30 + i, 9, f);
      if (typeof window.recalcSheet === 'function') window.recalcSheet(window.activeSheet, sh.data);
      const c = (sh.data || {})[(30 + i) + ',9'];
      if (!c) return '(マスが 無い)';
      return JSON.stringify({ v: c.v, d: c.d, f: c.f });
    });
  }, 式たち.map((x) => x[0]));
  for (let i = 0; i < 式たち.length; i++) {
    console.log('');
    console.log(式たち[i][0]);
    console.log('   実Excel ... ' + 式たち[i][1]);
    console.log('   出た .... ' + 出[i]);
  }
} finally {
  await browser.close().catch(() => {});
}
