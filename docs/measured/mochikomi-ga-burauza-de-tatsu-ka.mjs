/* mochikomi-ga-burauza-de-tatsu-ka.mjs
 *   -- ★持ち込みの 画面が ★本物の ブラウザで★ 立つか★（96）（2026-09-21）
 *
 *  ★★なぜ★★
 *    試験は 6緑0赤 でしたが、それは ★字を 読んだだけ★です。
 *    ⇒★読み込む と 登録される は 別★（記憶の 決まり）
 *    ⇒★立てて、実物の ファイルを 入れて、画面に 何が 出るかを 見ます★
 *
 *  ★★お客さんの 道で 入れます★★
 *    `#finput` に ファイルを 入れる（★画面の ボタンと 同じ 入口★）
 *    ⇒`Mochikomi.つなぐ()` が 付けた 聞き耳が 動きます
 *    ⇒★JS で 中を 直に 作りません★
 *
 *  ★★見る 物★★
 *    ①落ちて いないか（言づて・読み込めなかった 台）
 *    ②★1文★が 出て いるか
 *    ③★ドロップダウン★が 出て いるか（中身の 行も 数える）
 *    ④★絵★（png）＝★数が 緑でも 絵を 見るまで OK に しない★
 *
 *  使い方: node docs/measured/mochikomi-ga-burauza-de-tatsu-ka.mjs [--もと <xlsx>]
 */
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import http from 'node:http';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';
import { borrow, launch } from '../../scripts/_borrow-playwright.mjs';

const ここ = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(ここ, '..', '..');
const require_ = createRequire(path.join(ROOT, 'package.json'));
const TEMP = os.tmpdir();

const い = process.argv.indexOf('--もと');
const 元 = い > 0 ? process.argv[い + 1] : path.join(ROOT, 'tests/fixtures/kazari-graph.xlsx');
if (!fs.existsSync(元)) { console.log('★材料が 在りません★ ... ' + 元); process.exit(3); }

const 型 = {
  '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8', '.json': 'application/json', '.png': 'image/png',
  '.svg': 'image/svg+xml', '.ico': 'image/x-icon', '.xlsx': 'application/octet-stream',
};
const server = http.createServer((req, res) => {
  const u = decodeURIComponent(req.url.split('?')[0]);
  const p = path.join(ROOT, u === '/' ? 'mochikomi.html' : u.replace(/^\//, ''));
  if (!p.startsWith(ROOT) || !fs.existsSync(p) || fs.statSync(p).isDirectory()) { res.writeHead(404); res.end(); return; }
  res.writeHead(200, { 'Content-Type': 型[path.extname(p)] || 'application/octet-stream' });
  res.end(fs.readFileSync(p));
});
await new Promise((ok) => server.listen(0, '127.0.0.1', ok));
const 港 = server.address().port;
const 元URL = 'http://127.0.0.1:' + 港;
console.log('★立てました★ ' + 元URL);

/* ★playwright は ★借り方の 台★から 借ります★（自前で 道を 書かない）
     ＝`scripts/_borrow-playwright.mjs`（借り先の 名簿を 1か所に する 為に 在る）
     ★手元の 絶対パスを 焼き込みません★（記憶の 決まり）＝★別の 木で 走らせても 動く★ */
const ck = await borrow('mochikomi', 'chromium');
const browser = await launch('mochikomi', ck, {}, 'chromium');
/* ★落ちて くる ファイルを 受け取る には 先に 言って おく 必要が 在ります★
     ＝ を 立てないと ★押しても 何も 来ません★（2026-09-22 踏みました） */
const page = await browser.newPage({ viewport: { width: 1000, height: 900 }, acceptDownloads: true });
const 言づて = [], 落ち = [];
page.on('console', (m) => { if (m.type() === 'error') 言づて.push(m.text()); });
page.on('pageerror', (e) => 落ち.push(String(e && e.message ? e.message : e)));
page.on('requestfailed', (r) => 落ち.push('取れません ' + r.url()));

let 緑 = 0; const 赤 = [];
const 見る = (名, 条件, 出し) => {
  if (条件) { 緑 += 1; console.log('  ok   ' + 名 + (出し ? ' ... ' + 出し : '')); }
  else { 赤.push(名); console.log('  ★赤★ ' + 名 + (出し ? ' ... ' + 出し : '')); }
};

await page.goto(元URL + '/mochikomi.html', { waitUntil: 'load', timeout: 60000 });

/* ★ログインの 扉は 閉まったまま です★（Supabase が 無いので 開きません）
   ⇒★測る 為に 開けるのでは なく 「閉まって いる」事を 先に 数えます★ */
const 隠れ = await page.evaluate(() => {
  const a = document.getElementById('app');
  return { hidden: !!(a && a.hidden), 台: { BookOpen: !!window.BookOpen, HonNoNakami: !!window.HonNoNakami,
    Mochikomi: !!window.Mochikomi, FileOut: !!window.FileOut, ExcelVersion: !!window.ExcelVersion,
    Vba: !!window.Vba, XLSX: !!window.XLSX, XlsxIO: !!window.XlsxIO } };
});
console.log('');
console.log('★台が 窓に 居るか★');
Object.keys(隠れ.台).forEach((k) => 見る('  ' + k, 隠れ.台[k]));
見る('★ログインが 済むまで 中身は 隠れて いる★', 隠れ.hidden, 'app.hidden=' + 隠れ.hidden);

/* ★中身を 見るには 扉を 開ける 必要が 在ります★
   ＝★ここで 開けるのは 「測る ため」だけ★で、★作りは 1文字も 変えません★ */
/* ★★2026-09-21 ── ★数は 緑なのに 絵は ログインの 覆いだけ★でした★★
     ＝`#app` を 出しても ★覆いが 上に 乗った まま★
     ＝DOM は 読めるので ★16緑0赤★に なり、★絵を 開くまで 気づきません★
     ⇒★覆いも どけてから 撮ります★（★作りは 1文字も 変えません★） */
await page.evaluate(() => {
  document.getElementById('app').hidden = false;
  document.querySelectorAll('.login-ov, .login-wrap, #loginOverlay').forEach((e) => { e.style.display = 'none'; });
  /* ★名前で 当てずっぽうに 消さない★＝残って いたら 下で 数えます */
});
const 覆い = await page.evaluate(() => {
  const み = [...document.body.children].filter((e) => {
    if (e.id === 'app') return false;
    const s = getComputedStyle(e);
    return s.display !== 'none' && s.position === 'fixed';
  });
  return み.map((e) => e.className || e.tagName);
});
if (覆い.length) console.log('★まだ 上に 乗って いる 物★ ' + 覆い.join(' / '));

const 入 = await page.$('#finput');
await 入.setInputFiles(元);
await page.waitForFunction(() => {
  const k = document.getElementById('kekka');
  return k && !k.hidden;
}, null, { timeout: 60000 }).catch(() => {});

const 出 = await page.evaluate(() => {
  const t = (id) => (document.getElementById(id) || {}).textContent || '';
  const ds = [...document.querySelectorAll('#kuwashiku details')].map((d) => ({
    見出し: (d.querySelector('summary') || {}).textContent || '',
    行: d.querySelectorAll('li').length,
    注: (d.querySelector('.chu') || {}).textContent || '',
  }));
  return { 出た: !(document.getElementById('kekka') || {}).hidden, 名: t('fname'), 目: t('fmeta'),
    一文: t('hitokoto'), 畳み: ds, 知らせ: t('shirase'),
    保存できる: !(document.getElementById('hozon') || {}).disabled };
});

console.log('');
console.log('★画面に 出た 物★');
console.log('  名 ......... ' + 出.名);
console.log('  目 ......... ' + 出.目);
console.log('  ★1文★ .... ' + 出.一文);
出.畳み.forEach((d) => console.log('  畳み ....... ' + d.見出し + '（行 ' + d.行 + '）' + (d.注 ? ' 注=' + d.注.slice(0, 40) : '')));
if (出.知らせ) console.log('  知らせ ..... ' + 出.知らせ);

console.log('');
見る('★結果の 箱が 出た★', 出.出た);
見る('★1文が 出た★', 出.一文.length > 5, 出.一文);
見る('★ドロップダウンが 2つ★', 出.畳み.length === 2, '出た ' + 出.畳み.length + '個');
見る('★関数の 行が 1本 以上★', (出.畳み[0] || {}).行 > 0, '行 ' + ((出.畳み[0] || {}).行));
見る('★保存が 押せる★', 出.保存できる);
見る('★落ちて いない★', 落ち.length === 0, 落ち.slice(0, 3).join(' / '));
見る('★言づての 赤が 無い★', 言づて.length === 0, 言づて.slice(0, 3).join(' / '));

/* ══ ★★「1バイトも 変えません」を ★押して★ 数えます★★ ══（2026-09-22）
     ★画面に 大きさを 並べて 出す 形には して ありましたが、
       ★実際に 押して 測って いませんでした★＝★言っただけ★です。
     ⇒★お客さんの 道で 押します★（`#hozon` を click）
     ⇒落ちて くる バイト列を 受け取り、★元の ファイルと 1バイトずつ 突き合わせ★ */
const 落ち先 = path.join(TEMP, 'exally-mochikomi-otoshi');
fs.mkdirSync(落ち先, { recursive: true });
let 出たファイル = null;
/* ★★押してから 「来たか」を 見るのでは 遅い★★（2026-09-22 踏みました）
     ＝`page.on('download')` は 後から 呼ばれるので、★押した 直後に 数えると まだ 空★です。
     ⇒★押すのと 同時に 待つ★（`waitForEvent` と `click` を 並べる） */
const [落ちた] = await Promise.all([
  page.waitForEvent('download', { timeout: 60000 }).catch(() => null),
  page.click('#hozon'),
]);
if (落ちた) {
  const さき = path.join(落ち先, 落ちた.suggestedFilename());
  try { await 落ちた.saveAs(さき); 出たファイル = さき; }
  catch (e) { 落ち.push('落とせません ' + e.message); }
}
await page.waitForFunction(() => {
  const e = document.getElementById('hozon-kekka');
  return e && !e.hidden;
}, null, { timeout: 60000 }).catch(() => {});
const 保存の字 = await page.evaluate(() => (document.getElementById('hozon-kekka') || {}).textContent || '');
console.log('');
console.log('★保存を 押した 後★ ' + 保存の字);
if (出たファイル && fs.existsSync(出たファイル)) {
  const も = fs.readFileSync(元);
  const で = fs.readFileSync(出たファイル);
  const 同 = も.length === で.length && も.equals(で);
  console.log('  元 ' + も.length + 'B ／ 出た ' + で.length + 'B');
  見る('★★出た ファイルが 元と 1バイトも 違わない★★', 同,
    (同 ? '同じ' : '★違います★ 差 ' + (で.length - も.length) + 'B'));
} else {
  見る('★ファイルが 落ちて きた★', false, '★落ちて きません★（押しても 出て いない）');
}

/* ══ ★★畳みは 「印」では なく ★高さ★ で 数えます★★ ══（2026-09-22）
     ＝`open` 属性が 無くても ★中の 部品が 自分で `display` を 持つと 見えたまま★
     ＝記憶「★見た目の 見張りは 印(class)が 付いたかで 緑に するな★」
     ⇒★畳んだ 時 中身の 高さが 0★／★開けたら 0より 大きい★ を 数えます */
const 畳み = await page.evaluate(() => {
  const ds = [...document.querySelectorAll('#kuwashiku details')];
  const 高さ = (d) => [...d.children].filter((e) => e.tagName !== 'SUMMARY')
    .reduce((a, e) => a + e.getBoundingClientRect().height, 0);
  const 閉 = ds.map((d) => ({ 見出し: (d.querySelector('summary') || {}).textContent || '', 高: 高さ(d) }));
  ds.forEach((d) => { d.open = true; });
  const 開 = ds.map((d) => 高さ(d));
  ds.forEach((d) => { d.open = false; });
  return { 閉, 開 };
});
console.log('');
console.log('★畳みの 高さ★');
畳み.閉.forEach((x, i) => console.log('  ' + x.見出し + ' ... 畳んだ ' + Math.round(x.高) + 'px ／ 開けた ' + Math.round(畳み.開[i]) + 'px'));
見る('★★畳んだら 中身の 高さが 0★★', 畳み.閉.every((x) => x.高 === 0),
  '畳んだ 時の 高さ ' + 畳み.閉.map((x) => Math.round(x.高)).join(' / '));
見る('★開けたら 高さが 出る★', 畳み.開.some((h) => h > 0),
  '開けた 時の 高さ ' + 畳み.開.map((h) => Math.round(h)).join(' / '));

const 絵 = path.join(TEMP, 'exally-mochikomi.png');
await page.screenshot({ path: 絵, fullPage: false });
const st = fs.statSync(絵);
console.log('');
console.log('★絵★ ' + 絵);
console.log('  ' + st.size + 'B ／ sha256 ' + require_('node:crypto').createHash('sha256').update(fs.readFileSync(絵)).digest('hex'));
console.log('  ⇒★数が 緑でも 絵を 開いて 見るまで OK に しません★');

await browser.close();
server.close();
console.log('');
console.log('mochikomi-burauza: ' + 緑 + ' 緑 / ' + 赤.length + ' 赤');
process.exit(赤.length ? 1 : 0);
