/* hakaru-shotai-ga-aru-ka.mjs
 *   -- ★その 書体が 本当に 在るか★（110）（2026-09-24）
 *
 *  ★★なぜ★★
 *    108で 測ったら
 *      `11px 游ゴシック` ... 6.118
 *      `11px sans-serif` ... ★6.118（同じ）★
 *    ⇒★游ゴシックが 無くて 別の 字で 測って いる★ 疑い
 *    ⇒★でも 「同じ数＝無い」とは 限りません★（★たまたま 同じ★も 在り得る）
 *    ⇒★在りもしない 名前★を 3つ 混ぜて ★見分けが つくかを 先に 確かめます★
 *
 *  ★★やり方★★
 *    ①`document.fonts.check()` で 訊く
 *    ②★幅を 測って 並べる★（在りもしない 名前＝必ず 逃げの 字に なる）
 *    ③★逃げの 字と 同じ 幅なら 「その 書体は 効いて いない」★
 *    ★①だけ では 足りません★（`check` は 嘘を 言う 事が 在る）
 *
 *  使い方: node docs/measured/hakaru-shotai-ga-aru-ka.mjs
 */
import fs from 'node:fs';
import path from 'node:path';
import http from 'node:http';
import { fileURLToPath } from 'node:url';
import { borrow, launch } from '../../scripts/_borrow-playwright.mjs';

const ここ = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(ここ, '..', '..');
const server = http.createServer((req, res) => {
  res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
  res.end('<!DOCTYPE html><meta charset="utf-8"><body>x</body>');
});
await new Promise((ok) => server.listen(0, '127.0.0.1', ok));
const もと = 'http://127.0.0.1:' + server.address().port;

const wk = await borrow('shotai', 'webkit');
const browser = await launch('shotai', wk, {}, 'webkit');
const page = await browser.newPage();
await page.goto(もと, { waitUntil: 'load', timeout: 60000 });

const 出 = await page.evaluate(() => {
  const c = document.createElement('canvas').getContext('2d');
  const 測る = (f, s) => { c.font = f; return c.measureText(s).width; };
  const 名たち = [
    '游ゴシック', 'Yu Gothic', 'YuGothic', 'Meiryo', 'MS PGothic', 'MS Gothic',
    'Calibri', 'Arial', 'sans-serif', 'serif', 'monospace',
    'ZZZ-在りもしない-1', 'ZZZ-nai-2', 'QQQQQQ-nonexistent-3',
  ];
  const 字 = '0';
  const 長い = '給料明細 1月分 合計 56,647 円';
  const 出 = {};
  名たち.forEach((な) => {
    出[な] = {
      零: 測る('11px "' + な + '"', 字),
      長: 測る('11px "' + な + '"', 長い),
      訊いた: (document.fonts && document.fonts.check) ? document.fonts.check('11px "' + な + '"') : null,
    };
  });
  return 出;
});

console.log('★書体が 本当に 在るか★（webkit）');
console.log('  名前'.padEnd(28) + '「0」の幅   長い字の幅   fonts.check');
const 逃げ = 出['ZZZ-在りもしない-1'];
Object.entries(出).forEach(([な, x]) => {
  const 同じ = Math.abs(x.零 - 逃げ.零) < 0.01 && Math.abs(x.長 - 逃げ.長) < 0.01;
  console.log('  ' + な.padEnd(26) + String(x.零.toFixed(3)).padStart(8)
    + String(x.長.toFixed(3)).padStart(12) + String(x.訊いた).padStart(12)
    + (同じ ? '   ★逃げの 字と 同じ＝効いて いない★' : ''));
});
console.log('');
const 無い3 = ['ZZZ-在りもしない-1', 'ZZZ-nai-2', 'QQQQQQ-nonexistent-3'].map((n) => 出[n].長);
const 割れる = Math.abs(無い3[0] - 無い3[1]) < 0.01 && Math.abs(無い3[1] - 無い3[2]) < 0.01;
console.log('★★道具の 確かめ★★ 在りもしない 名前 3つが 同じ 幅か ... ' + (割れる ? '★同じ（道具は 効いて います）★' : '★違う＝この 見分け方は 使えません★'));

await browser.close();
server.close();
