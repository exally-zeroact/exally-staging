/* utte-osu-webkit.mjs -- ★お客さんと 同じに ★キーボードで 打つ★★（2026-09-19）
 *
 *  ★★なぜ 要るか（★決まりの 紙に 書いて ありました★）★★
 *    `team/global-rules.md` §11②
 *      「実アプリを Playwright/jsdom で 動かし、★全タブ・全ボタン・全入力を
 *        実際に 押す/打ち込む★」
 *    ⇒★★repo の 見張りは ★1本も 打って いません★★（2026-09-19 に 数えた）
 *       `osu-kami-webkit.mjs` ／ `shiki-wo-osu-webkit.mjs` ほか 全部 `window.setCell(...)`
 *    ⇒★記憶「JSで イベントを 投げて 出した 物は ★お客さんの 道では ない★」★
 *    ⇒★★だから これを 書きました★★
 *
 *  ★★何を して いるか★★
 *    ①`#grid-canvas` を クリックして 焦点を 当てる（★お客さんと 同じ★）
 *    ②Ctrl+Home で A1 へ
 *    ③矢印で 目当ての マスへ
 *    ④★`page.keyboard.type(式)` で 1文字ずつ 打つ★
 *    ⑤Enter で 確定（★選びは 1つ 下へ 動きます★）
 *
 *  ★★見て いない 事（★正直に★）★★
 *    ・★出た 字は 画（canvas）からは 読めません★
 *      ⇒★板の `d`（画に 描く 字）を 読んで います★
 *      ⇒★★「打つ」は 本物／「読む」は 板から★★
 *    ・★分母は 小さいです★（★2,116本は `osu-kami-webkit.mjs` の 役★）
 *      ＝★これは 「打った 時にも 同じ 答えに なるか」を 見る 道具★
 *    ・★書式・罫線・溢れの 見た目は 見て いません★
 *
 *  使い方:
 *    node docs/measured/utte-osu-webkit.mjs --どこ=<URL>
 */
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const ここ = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(ここ, '..', '..');
/* ★repo の 共通の 借り口を 使います★（★作る前に 探せ★） */
const { borrow, launch } = await import(
  pathToFileURL(path.join(ROOT, 'scripts/_borrow-playwright.mjs')).href);

const 引数 = process.argv.slice(2);
const 取る = (名) => {
  const a = 引数.find((x) => x.startsWith('--' + 名 + '='));
  return a ? a.slice(名.length + 3) : null;
};
const どこ = 取る('どこ');
if (!どこ) {
  console.log('★使い方★ node docs/measured/utte-osu-webkit.mjs --どこ=<URL>');
  process.exit(2);
}

/* ══ ★打つ 式と 実Excel の 答え★ ══
   ★出どころは 全部 `docs/measured/golden-*.tsv`★＝★手で 作った 数では ありません★ */
const 問い = [
  { 式: '=1+1', 正: '2', 訳: '一番 易しい 物（★打てて いるかの 対照★）' },
  { 式: '=SUM(1,2,3)', 正: '6', 訳: '足し算' },
  { 式: '=PERMUT(5,2)', 正: '20', 訳: 'golden-kansuu-7kaime の 対照' },
  { 式: '=COUPDAYS(DATE(2008,11,11),DATE(2021,3,1),2,0)', 正: '180',
    訳: '★お金に DATE() を 渡す★（09-18 に 729本 直した 所）' },
  { 式: '=COUPDAYS(39763,44256,2,0)', 正: '180', 訳: '同じ 日を 数で 渡す（★上と 同じに なる はず★）' },
  { 式: '=ODDLPRICE(DATE(2009,3,10),DATE(2009,8,31),DATE(2008,11,30),0.045,0.05,100,4,1)',
    正: '99.7764763580019', 訳: 'ODDL（★今日 出した 物★）' },
  { 式: '=ODDLYIELD(DATE(2009,3,10),DATE(2009,8,31),DATE(2008,11,30),0.045,99.7764763580019,100,4,1)',
    正: '0.05', 訳: 'ODDL の 利回り' },
  { 式: '=ODDFPRICE(DATE(2009,3,10),DATE(2015,2,28),DATE(2008,12,5),DATE(2009,8,31),0.045,0.05,100,4,1)',
    正: '96.1767213744142', 訳: '★ODDF（★出す 前は #NAME? の はず★）★' },
  { 式: '=ODDFPRICE(DATE(2009,3,10),DATE(2015,2,28),DATE(2008,11,29),DATE(2009,8,31),0.045,0,100,4,1)',
    正: '126.88994565217391', 訳: '★ODDF の ★間の 1日★（今日 一番 難しかった 所）★' },
  { 式: '=CONVERT(1,"m","cm")', 正: '100', 訳: 'convert' },
  { 式: '=TEXT(1234.5,"#,##0.00")', 正: '1,234.50', 訳: '★書式（SheetJS 側）★' },
  { 式: '=LET(x,1,x+1)', 正: '2', 訳: 'JS層が 受ける 物' },
];

/* ★打ち始める マス★（★空いて いる 所★） */
const 始まり行 = 200;
const 列 = 9;   /* J列 */

/* ★★元から 在る 誤り（★名指しで 許す★）★★
   ＝`osu-kami-webkit.mjs` と ★同じ 2件★（★名簿を 2つ 持たない★）
     ①WebKit が `interactive-widget` を 知らない だけ（★Chrome では 出ません★）
     ②Vercel の ★仮の 配信だけ★に 出る `vercel.live` の 中の 字
       ⇒★本番（exally.vercel.app）には 在りません★（2026-09-18 実測 0件）
   ⇒★★消えたら 赤に します★★（＝許しを 外す） */
const 元から在る誤り = [
  'Viewport argument key "interactive-widget" not recognized',
  'navigator.storage.persisted',
];
const 窓の誤り = [];
const 元から在る = [];
const 誤りを分ける = (s2) => {
  if (元から在る誤り.some((x) => s2.indexOf(x) >= 0)) 元から在る.push(s2);
  else 窓の誤り.push(s2);
};
const wk = await borrow('utte-osu', 'webkit');
const b = await launch('utte-osu', wk, {}, 'webkit');
const p = await b.newPage({ viewport: { width: 1280, height: 900 } });
p.on('console', (m) => { if (m.type() === 'error') 誤りを分ける('console.error: ' + String(m.text()).slice(0, 160)); });
p.on('pageerror', (e) => 誤りを分ける('pageerror: ' + String(e.message).slice(0, 160)));

console.log('');
console.log('[utte-osu-webkit] ★お客さんと 同じに キーボードで 打つ★');
console.log('  ★どこ★ ... ' + どこ);

const 返事 = await p.goto(どこ, { waitUntil: 'load', timeout: 60000 });
console.log('  ★返事★ ... ' + (返事 ? 返事.status() : '(無し)'));
await p.waitForFunction(() => typeof window.getCell === 'function', null, { timeout: 30000 });

/* ★★ログインの 覆いを どける★★
   ＝`osu-kami-webkit.mjs` と ★同じ 手★（★お客さんは ここで ログインします★）
   ⇒★★これは 「打つ 道」の 話では ありません／覆いが 在ると 画が 0x0 です★★ */
await p.evaluate(() => {
  document.body.classList.remove('exally-locked');
  const ov = document.getElementById('loginOv');
  if (ov) { ov.classList.remove('open'); ov.style.display = 'none'; }
});
await p.waitForFunction(() => {
  const c = document.getElementById('grid-canvas');
  return c && c.getBoundingClientRect().width > 100;
}, null, { timeout: 30000 });

/* ①canvas を クリック（★お客さんと 同じ★） */
await p.click('#grid-canvas');
/* ②A1 へ */
await p.keyboard.press('Control+Home');
await p.waitForTimeout(120);
/* ③目当ての マスへ（★矢印だけで 動かす＝お客さんと 同じ★） */
for (let i = 0; i < 列; i += 1) await p.keyboard.press('ArrowRight');
/* ★行は 0 から 数えます★（A1 が 0）＝★矢印の 回数＝行の 番号★ */
for (let i = 0; i < 始まり行; i += 1) await p.keyboard.press('ArrowDown');
await p.waitForTimeout(150);

const 居場所 = await p.evaluate(() => ({ r: window.selR1, c: window.selC1 }));
console.log('  ★矢印だけで 着いた 所★ ... 行 ' + 居場所.r + ' / 列 ' + 居場所.c
  + '（★狙い 行 ' + 始まり行 + ' / 列 ' + 列 + '★）');
if (居場所.r !== 始まり行 || 居場所.c !== 列) {
  console.log('  ★★狙った マスに 着いて いません＝この 回は 捨てます★★');
  await b.close();
  process.exit(3);
}

const 出 = [];
for (let i = 0; i < 問い.length; i += 1) {
  const 行 = 始まり行 + i;
  /* ★★1文字ずつ 打つ★★（★これが 本物の 入力★） */
  await p.keyboard.type(問い[i].式, { delay: 4 });
  await p.keyboard.press('Enter');
  await p.waitForTimeout(90);
  const よみ = await p.evaluate((a) => {
    const c = window.getCell(a.r, a.c);
    return { f: (c && c.f) || null, d: c ? String(c.d !== undefined ? c.d : c.v) : '(マスが 無い)' };
  }, { r: 行, c: 列 });
  出.push({ ...問い[i], 行, ...よみ });
}

await b.close();

const 数か = (x) => /^[-+]?[0-9]*[.]?[0-9]+([eE][-+]?[0-9]+)?$/.test(x);
const 合うか = (a, b2) => {
  if (String(a).trim() === String(b2).trim()) return true;
  if (数か(a) && 数か(b2)) {
    const x = Number(a); const y = Number(b2);
    return Math.abs(x - y) <= Math.max(1e-9, Math.abs(y) * 1e-9);
  }
  return false;
};

console.log('');
console.log('  ★★打った 式 ... ' + 出.length + '本★★（★これが 分母★）');
console.log('');
let 合 = 0;
let 打てず = 0;
for (const x of 出) {
  const 打てた = x.f === x.式;
  if (!打てた) 打てず += 1;
  const 同 = 合うか(x.d, x.正);
  if (同) 合 += 1;
  console.log('  ' + (同 ? 'o ' : '★X★') + ' ' + x.式.slice(0, 66));
  console.log('      出た ' + String(x.d).slice(0, 26) + ' ／実Excel ' + String(x.正).slice(0, 26)
    + (打てた ? '' : '  ★★打った 字が マスに 入って いません（' + String(x.f).slice(0, 40) + '）★★'));
  console.log('      ' + x.訳);
}

console.log('');
console.log('  ★★合った ' + 合 + ' / ' + 出.length + '★★');
console.log('  ★★打った 字が そのまま 入らなかった ... ' + 打てず + '本★★'
  + '（★1本でも 出たら 打つ 道が 壊れて います★）');
console.log('  ★★窓の 誤り（★新しい 物だけ★） ... ' + 窓の誤り.length + '件★★');
console.log('    ★元から 在る（名指しで 許した） ... ' + 元から在る.length + '件★');
if (元から在る.length === 0) {
  console.log('    ★★許した 物が 1件も 出ませんでした＝許しを 外して ください★★');
}
for (const e of 窓の誤り.slice(0, 6)) console.log('    ・' + e);
console.log('');
console.log('  ★★断り★★');
console.log('    ・★打つのは 本物（矢印と 1文字ずつの 打鍵だけ）★');
console.log('    ・★読むのは 板の `d`★＝★画（canvas）の 字は 読んで いません★');
console.log('    ・★分母は 小さいです★＝2,116本は `osu-kami-webkit.mjs` の 役');
