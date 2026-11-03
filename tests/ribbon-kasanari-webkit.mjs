/* ribbon-kasanari-webkit.mjs — ★リボンの ボタンが 重なっていないか（本物の ブラウザ）★ 2026-09-07
 *
 *  ★★何が 在ったか（実測）★★
 *    ★webkit（＝iPhone・iPad・Mac の Safari）で ホームの リボンが 潰れていた★
 *      見えている ★55個★ のうち ★31個★が
 *      ★真ん中を 押すと 別の ボタンが 出る★（幅 1280〜2560 の 5通りとも 同じ）
 *      ⇒ 絵を 開くと ★字が 重なって 読めない★
 *      ⇒★本物の Chrome では 0個★＝★中身では なく ★幅の 数え方★の 違い★
 *
 *  ★★元★★
 *    `.rb-items` は ★縦に 積んで 段が 一杯に なったら 次の 列へ★（`flex-flow: column wrap`）。
 *    ★webkit は この 形の 箱の 幅を ★1列ぶんしか 数えない★★
 *      実測（幅1920・フォントの 組）… Chrome 箱428px／webkit 箱 ★80px★（中身は 428px）
 *    ⇒ 隣の 組が その ぶん 上に 乗る。
 *
 *  ★★直し方★★ `lib/ribbon.js` の `組の幅を直す()`
 *    ★並びは 変えない★（`column wrap` の まま）＝★Chrome の 見た目は 1つも 動かさない★
 *    ★ブラウザ自身が 出した 数★で 箱の 幅を 入れ直すだけ。
 *    ★大きさが 変わるたびに 数え直す★（隠れている 間に 描かれた 時も 拾う）
 *
 *  ★★この 試験が 見る 物★★
 *    ★「在る／見えている」では なく ★押したら 何が 出るか★★
 *    ⇒ ボタンの 真ん中で `document.elementFromPoint` を 呼ぶ。
 *    ⇒★真ん中が 窓の 外に 在る 物は 数えない★（帯は 横に すべる＝重なりでは ない）
 *      ★でも 0に 見せる為に 捨てない★＝別に 数えて 出す。
 *
 *  実測の 記録 … docs/measured/ribbon/kasanari-naosu-mae.txt ／ kasanari-naoshita-ato.txt
 *
 *  走らせ方: node tests/ribbon-kasanari-webkit.mjs [--self-test]
 */
import path from 'node:path'; import http from 'node:http'; import fs from 'node:fs';
import { fileURLToPath } from 'node:url';
import { borrow, launch } from '../scripts/_borrow-playwright.mjs';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const 自己試験 = process.argv.includes('--self-test');
const 幅たち = [1280, 1440, 1600, 1920, 2560];
let pass = 0, fail = 0;
const T = (n, よい, 添え) => {
  if (よい) { pass++; console.log('  ok   ' + n); }
  else { fail++; console.log('  NG   ' + n + (添え ? '\n       ' + 添え : '')); }
};
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

/* ★数え方は 1か所★＝本番でも 自己試験でも 同じ 物を 使う */
const 数える = () => {
  const bs = [...document.querySelectorAll('#ribbon [data-act], #ribbon .rb-item')];
  let 見 = 0, 隠 = [], 画面の外 = 0, はみ出し = 0;
  for (const b of bs) {
    const x = b.getBoundingClientRect();
    if (x.width < 1 || x.height < 1) continue;
    const cx = x.x + x.width / 2, cy = x.y + x.height / 2;
    if (cx < 0 || cx > window.innerWidth || cy < 0 || cy > window.innerHeight) { 画面の外++; continue; }
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
  return { 見, 隠: 隠.length, 例: 隠.slice(0, 4), 画面の外, はみ出し };
};

console.log('');
console.log('[ribbon-kasanari-webkit] ★リボンの ボタンが 重なっていないか★');

const wk = await borrow('ribbon-kasanari', 'webkit');
const browser = await launch('ribbon-kasanari', wk, {}, 'webkit');
const 配信 = await 立てる(ROOT);
try {
  for (const W of 幅たち) {
    const page = await browser.newPage({ viewport: { width: W, height: 1000 } });
    await page.goto(配信.url + '/book.html', { waitUntil: 'load' });
    /* ★鍵を 外す（ログインは 通していない）★＝リボンが 見えている 形に する */
    await page.evaluate(() => {
      document.body.classList.remove('exally-locked');
      const ov = document.getElementById('loginOv');
      if (ov) { ov.classList.remove('open'); ov.style.display = 'none'; }
    });
    await page.waitForTimeout(1800);
    const r = await page.evaluate(数える);
    T('★幅 ' + W + ' … 押した 物が そのまま 出る（重なり 0個）★', r.隠 === 0,
      '見えた ' + r.見 + '個 ／ ★重なり ' + r.隠 + '個★ ／ はみ出し ' + r.はみ出し + 'px'
      + (r.例.length ? '\n       ' + r.例.join('\n       ') : ''));
    T('★幅 ' + W + ' … 組の 箱から 中身が はみ出していない★', r.はみ出し === 0,
      'はみ出し ' + r.はみ出し + 'px');
    if (W === 1920) {
      T('★幅 1920 では 全部の 部品が 画面に 出ている★', r.見 >= 50 && r.画面の外 === 0,
        '見えた ' + r.見 + '個 ／ 画面の外 ' + r.画面の外 + '個');
    }

    if (自己試験 && W === 1920) {
      console.log('\n  [self-test] ★物差しが 効いているか★');
      /* ★★壊すのは ★画面（写し）★だけ★★＝repo の ファイルは 1バイトも 触らない
         直しが 入れた 幅を 剥がすと ★元の 壊れた 形★に 戻るはず */
      const 戻し = await page.evaluate(() => {
        let n = 0;
        for (const items of document.querySelectorAll('#ribbon .rb-items')) {
          if (items.style.width) { items.style.width = ''; n++; }
        }
        return n;
      });
      T('★直しが 実際に 幅を 入れていた（入れた 箱の 数 > 0）★', 戻し > 0, '入れていた 箱 ' + 戻し + '個');
      await page.waitForTimeout(200);
      const 壊 = await page.evaluate(数える);
      /* ★★『わざと 壊したら 赤が 1個 出た』で 信用しては いけない★★（会社の 決まり）
         ⇒ 0個なら 誰か 気づくが ★1個 出ると 誰も 疑わない★
           （代行請求 2026-09-05 … ★6個 出るはずが 1個★だった）
         ⇒ だから ★出るはずの 数★を 書いて 比べる。
         ★実測★ 幅1920 で 剥がすと ★重なり 31個／はみ出し 1022px★ に 戻る。
         ★下限は 20個★（部品の 数は 版で 少し 動くので 31 ちょうどには しない。
           ただし ★1個や 2個で 緑に ならない★所に 置く） */
      const 出るはず = 20;
      T('★★幅を 剥がすと ' + 出るはず + '個 以上 赤に なる★★'
        + '（★壊した 箱 ' + 戻し + '個 → 赤 ' + 壊.隠 + '個★）', 壊.隠 >= 出るはず,
        '★壊した 箱 ' + 戻し + '個 に 対して 赤が ' + 壊.隠 + '個 しか 出ない★'
        + '＝実測は 31個 だった ⇒ ★この 試験は もう 中身を 見ていない★');
      T('★★幅を 剥がすと はみ出しが 戻る（500px 以上）★★'
        + '（実測 1022px ／ 出た ' + 壊.はみ出し + 'px）', 壊.はみ出し >= 500,
        'はみ出しが ' + 壊.はみ出し + 'px しか 出ない＝★元に 戻っていない★');
      console.log('       … ★壊した 箱 ' + 戻し + '個★ ／ ★赤 ' + 壊.隠 + '個★'
        + ' ／ はみ出し ' + 壊.はみ出し + 'px（実測の 元は 31個／1022px）');
    }
    await page.close();
  }
} finally {
  配信.閉じる();
  await browser.close();
}
console.log('');
console.log('ribbon-kasanari-webkit: ' + pass + ' 緑 / ' + fail + ' 赤');
process.exit(fail ? 1 : 0);
