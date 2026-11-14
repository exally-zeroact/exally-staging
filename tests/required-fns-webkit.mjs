/* required-fns-webkit.mjs — ★起動時の「要る部品」の確かめを 本物の ブラウザで 見る★ 2026-11-14
 *
 *  ★なぜ★
 *    book.html の checkRequiredFns は 起動時に REQUIRED_FNS の 名が 全部 function かを 見て、
 *    ★無ければ 画面に 赤い 帯を 出す★（window.__EXALLY_MISSING__ に 名を 積む）。
 *    ところが それを 本物の ブラウザで 見る 試験は ★0本★だった。
 *    字の 見張り（tests/no-silent-optional.test.mjs）は「表に 載っているか」「どこかに 定義の 字が 在るか」までで、
 *    ★点つきの 名（FormulaSotoPlug.外へ出す）を 本当に 引けるか★は 字では 分からない（甘い 判定）。
 *
 *  ★見る事★
 *    ・開いた後に checkRequiredFns が 走った（window.__EXALLY_REQUIRED_CHECKED__）
 *    ・window.__EXALLY_MISSING__ が 空 で、赤い 帯（#exally-missing-fns）が DOM に 無い
 *      （★見えるかで なく DOM に 在るかで 見る★＝ログイン前は body.exally-locked で 隠れる）
 *    ・REQUIRED_FNS の 全部が window から 1段ずつ 辿って function
 *  ★--self-test★＝REQUIRED_FNS に 在りもしない 名（日本語・点つきの 末尾違い・頭が 無い）を 足した 写しを 配る
 *    → 帯が DOM に 出て、__EXALLY_MISSING__ に ★その3つだけ★が 載り、辿りで 投げない
 *
 *  外へ 出る 物（jsdelivr・supabase.co・字体）は page.route で 止める（手元と CI で 同じ 条件）。
 *  走らせ方: node tests/required-fns-webkit.mjs [--self-test]
 */
import path from 'node:path'; import http from 'node:http'; import fs from 'node:fs';
import { fileURLToPath } from 'node:url';
import { borrow, launch } from '../scripts/_borrow-playwright.mjs';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const 壊す = process.argv.includes('--self-test');
const まぼろし = ['まぼろしの部品', 'FormulaSotoPlug.まぼろしの口', 'まぼろしの頭.外へ出す'];
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
    const 名 = decodeURIComponent(String(q.url).split('?')[0]).replace(/^\/+/, '');
    const f = path.join(root, 名);
    if (!f.startsWith(root) || !fs.existsSync(f) || fs.statSync(f).isDirectory()) { r.statusCode = 404; return r.end('no'); }
    r.setHeader('content-type', 型[path.extname(f).toLowerCase()] || 'application/octet-stream');
    if (壊す && 名 === 'book.html') {
      const 元 = fs.readFileSync(f, 'utf8');
      const 印 = 'var REQUIRED_FNS = [';
      if (元.split(印).length !== 2) { r.statusCode = 500; return r.end('REQUIRED_FNS の 印が 1つで ない'); }
      return r.end(元.replace(印, 印 + まぼろし.map((n) => "'" + n + "', ").join('')));
    }
    fs.createReadStream(f).pipe(r);
  });
  return new Promise((x) => s.listen(0, '127.0.0.1', () => x({ url: 'http://127.0.0.1:' + s.address().port, 閉じる: () => s.close() })));
}

console.log('');
console.log('[required-fns-webkit] ★起動時の 要る部品の 確かめを 本物の ブラウザで 見る★' + (壊す ? '（--self-test＝在りもしない 名を 足した 写し）' : ''));
const wk = await borrow('required-fns', 'webkit');
const browser = await launch('required-fns', wk, {}, 'webkit');
const page = await browser.newPage({ viewport: { width: 1280, height: 800 } });
const 止めた = [];
await page.route(/^https?:\/\/(?!127\.0\.0\.1)/, (rt) => { 止めた.push(rt.request().url()); return rt.abort(); });
const 落ち = [];
page.on('pageerror', (e) => 落ち.push(String(e && e.message || e)));
const 配信 = await 立てる(ROOT);
try {
  await page.goto(配信.url + '/book.html', { waitUntil: 'load' });
  await page.waitForFunction(() => window.__EXALLY_REQUIRED_CHECKED__ === true, null, { timeout: 30000 }).catch(() => {});
  const 見 = await page.evaluate(() => {
    const 引く = (n) => { let v = window; for (const k of String(n).split('.')) { if (v == null) return undefined; v = v[k]; } return v; };
    const 表 = Array.isArray(window.REQUIRED_FNS) ? window.REQUIRED_FNS.slice() : null;
    return {
      走った: window.__EXALLY_REQUIRED_CHECKED__ === true,
      無い: Array.isArray(window.__EXALLY_MISSING__) ? window.__EXALLY_MISSING__.slice() : null,
      帯: !!document.getElementById('exally-missing-fns'),
      帯の字: (document.getElementById('exally-missing-fns') || {}).textContent || '',
      表,
      辿れない: 表 ? 表.filter((n) => typeof 引く(n) !== 'function') : null,
      点つき: 表 ? 表.filter((n) => n.includes('.')).length : 0,
      日本語: 表 ? 表.filter((n) => /[^\x00-\x7f]/.test(n)).length : 0,
    };
  });
  T('★checkRequiredFns が 走った★', 見.走った);
  T('★REQUIRED_FNS を 読めた（空振りで ない）★', 見.表 && 見.表.length >= 40, '数 ' + (見.表 && 見.表.length));
  if (!壊す) {
    T('★__EXALLY_MISSING__ が 空★', Array.isArray(見.無い) && 見.無い.length === 0, JSON.stringify(見.無い));
    T('★赤い 帯が DOM に 無い★', !見.帯, 見.帯の字);
    T('★REQUIRED_FNS の 全部が window から 辿って function（点つき ' + 見.点つき + '・日本語 ' + 見.日本語 + '）★',
      Array.isArray(見.辿れない) && 見.辿れない.length === 0 && 見.点つき >= 3 && 見.日本語 >= 15, JSON.stringify(見.辿れない));
  } else {
    const 無い = 見.無い || [];
    T('★足した 3つ（日本語・点つきの 末尾違い・頭が 無い）が __EXALLY_MISSING__ に 載る★',
      まぼろし.every((n) => 無い.includes(n)), JSON.stringify(無い));
    T('★それ以外は 載らない（本物の 名を 無いと 言わない）★', 無い.length === まぼろし.length, JSON.stringify(無い));
    T('★赤い 帯が DOM に 出て 名が 書いてある★', 見.帯 && まぼろし.every((n) => 見.帯の字.includes(n)), 見.帯の字);
  }
  T('★確かめの 辿りで ページが 落ちていない★', !落ち.some((m) => /まぼろし|_要る部品を引く|checkRequiredFns/.test(m)), 落ち.join(' ／ '));
  console.log('       ... 外へ 出る 読み込みを 止めた 数 ' + 止めた.length + ' ／ ページの 落ち ' + 落ち.length + '件');
} finally {
  配信.閉じる();
  await browser.close();
}
console.log('');
console.log('required-fns-webkit' + (壊す ? ' --self-test' : '') + ': ' + pass + ' 緑 / ' + fail + ' 赤');
process.exit(fail ? 1 : 0);
