/* cell-honmono-de-osu.mjs — ★CELL を ★本物の Chrome★ で お客さんと 同じ 押し方で 押す★（2026-09-07）
 *
 *  ★★なぜ 要るか★★
 *    CELL は ★そのマスが どう 見えているか★を 返します。
 *    ⇒ 幅も 表示形式も そろえも ★画面が 持っている 物★
 *    ⇒ 部品（lib）と エンジンだけ 緑でも
 *      ★book.html の 繋ぎ（見た目を 渡す 所）が 間違っていたら 気づけない★
 *      （`マス.numFmt` を `マス.fmt` と 書き間違えても ★試験は 全部 緑★に なる）
 *    ⇒★★だから 本物の ブラウザで ★お客さんと 同じ 押し方★で 確かめる★★
 *
 *  ★★お客さんと 同じ 道で 動かす★★
 *    ・マスは ★本物の マウスで 押す★（canvas の 座標）
 *    ・値と 式は ★数式バーに 打つ★（`#formula-input` → Enter）
 *    ・表示形式・そろえは ★リボンの ボタンを 本物の マウスで 押す★
 *    ⇒★JS で イベントを 投げたり 関数を 直に 呼んだりは していません★
 *
 *  ★★なぜ Chrome か（webkit では ない 訳）★★
 *    ★webkit では ホームの リボンの ボタンが 重なっていて 押せません★
 *      実測 … 見えている 55個 のうち ★31個★が
 *             ★真ん中を 押すと 別の ボタンが 出る★（幅 1280〜2560 の 5通りとも 同じ）
 *      ★本物の Chrome では 0個★＝重なっていない
 *      ★これは CELL とは 別の 話★＝★本番（origin/main）でも 同じ★なので
 *        ★別の 件として 指示役へ 上げました（この 回では 直していません）★
 *    ⇒ ここでは ★リボンが ちゃんと 並ぶ Chrome★で 押します
 *
 *  ★答え★ 実Excel に 打たせた 物と 同じ（docs/measured/kansuu46/golden-cell*.tsv）
 *    `#,##0` → `,0` ／ `¥#,##0` → `C0` ／ 右そろえの 字 → `"` ／ 何もしない 列の 幅 → 8
 *
 *  ★外へは 1回も 出ません★（127.0.0.1 だけ）
 *
 *  使い方: node docs/measured/kansuu46/cell-honmono-de-osu.mjs
 */
import path from 'node:path'; import http from 'node:http'; import fs from 'node:fs';
import { fileURLToPath } from 'node:url';
import { borrow, launch } from '../../../scripts/_borrow-playwright.mjs';

const ここ = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(ここ, '..', '..', '..');
const 出す = path.join(ここ, 'cell-honmono-de-osu.txt');
const 紙 = [];
let pass = 0, fail = 0;
const 言う = (s) => { 紙.push(s); console.log(s); };
const T = (n, よい, 添え) => {
  if (よい) { pass++; 言う('  ok   ' + n); }
  else { fail++; 言う('  NG   ' + n + (添え ? '\n       ' + 添え : '')); }
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

言う('# ★CELL を 本物の Chrome で 押した★（' + new Date().toISOString().slice(0, 10) + '）');
言う('');
言う('★外へは 1回も 出ていません（127.0.0.1 だけ）★');
言う('★押し方★ マスは 本物の マウス／値と 式は 数式バー／表示形式は リボンの ボタン');
言う('');

const cr = await borrow('cell', 'chromium');
const browser = await launch('cell', cr, { channel: 'chrome' }, 'chromium');
const page = await browser.newPage({ viewport: { width: 1920, height: 1080 } });
const 配信 = await 立てる(ROOT);
try {
  await page.goto(配信.url + '/book.html', { waitUntil: 'load' });
  /* ★何で 出したか★＝手元・★鍵を 外した（ログインは 通していない）★ */
  await page.evaluate(() => {
    document.body.classList.remove('exally-locked');
    const ov = document.getElementById('loginOv');
    if (ov) { ov.classList.remove('open'); ov.style.display = 'none'; }
  });
  await page.waitForTimeout(1500);

  T('★画面が CELL の 部品を 読み込んでいる★',
    await page.evaluate(() => typeof window.FormulaCell === 'object'
      && typeof window.FormulaCellPlug === 'object'));

  /* ★リボンが 重なっていないか 先に 数える★（重なっていたら 押しても 別の 物が 出る） */
  const 重なり = await page.evaluate(() => {
    const bs = [...document.querySelectorAll('#ribbon [data-act], #ribbon .rb-item')];
    let 見 = 0, 隠 = 0;
    for (const b of bs) {
      const x = b.getBoundingClientRect();
      if (x.width < 1 || x.height < 1) continue;
      if (x.y < 0 || x.y > window.innerHeight || x.x < 0 || x.x > window.innerWidth) continue;
      見++;
      const e = document.elementFromPoint(x.x + x.width / 2, x.y + x.height / 2);
      if (e !== b && !b.contains(e)) 隠++;
    }
    return { 見, 隠 };
  });
  T('★リボンが 重なっていない（押した 物が そのまま 出る）★', 重なり.隠 === 0,
    '見えた ' + 重なり.見 + '個 ／ 重なり ' + 重なり.隠 + '個');

  const 枠 = await page.evaluate(() => {
    const c = document.getElementById('grid-canvas');
    const r = c.getBoundingClientRect();
    return { x: r.left, y: r.top, HDR_W: window.HDR_W, HDR_H: window.HDR_H,
      ROW_H: window.ROW_H, COL_W: window.COL_W };
  });
  T('★画面の 枠の 数を 読めた★', !!(枠 && 枠.COL_W > 0 && 枠.ROW_H > 0), JSON.stringify(枠));

  async function マスを押す(行番, 列) {
    const x = 枠.x + 枠.HDR_W + 列 * 枠.COL_W + 枠.COL_W / 2;
    const y = 枠.y + 枠.HDR_H + 行番 * 枠.ROW_H + 枠.ROW_H / 2;
    await page.mouse.click(x, y);
    await page.waitForTimeout(150);
  }
  async function 打つ(行番, 列, 字) {
    await マスを押す(行番, 列);
    await page.fill('#formula-input', String(字));
    await page.press('#formula-input', 'Enter');
    await page.waitForTimeout(300);
  }
  async function 画面の字(行番, 列) {
    return page.evaluate(([r, c]) => {
      const d = (window.sheets[window.activeSheet].data) || {};
      const m = d[r + ',' + c];
      if (!m) return '';
      return String(m.d !== undefined && m.d !== '' ? m.d : (m.v === undefined ? '' : m.v));
    }, [行番, 列]);
  }
  async function リボンを押す(名) {
    const 印 = page.locator('#ribbon [title*="' + 名 + '"], #ribbon [data-act*="' + 名 + '"]');
    const n = await 印.count();
    for (let i = 0; i < n; i++) {
      const b = 印.nth(i);
      if (await b.isVisible()) { await b.click(); await page.waitForTimeout(300); return true; }
    }
    return false;
  }

  /* ① 数を 打って 桁区切りを 当てる → CELL("format") は ,0 */
  await 打つ(0, 0, '1234.5678');                       /* A1 */
  await マスを押す(0, 0);
  T('★リボンの「桁区切り」を 本物の マウスで 押せた★', await リボンを押す('桁区切り'));
  T('★押した 結果が 画面に 入った（表示形式 #,##0）★',
    await page.evaluate(() => ((window.sheets[window.activeSheet].data || {})['0,0'] || {}).numFmt === '#,##0'));
  await 打つ(0, 1, '=CELL("format",A1)');              /* B1 */
  T('★★=CELL("format",A1) が ,0★★（実Excel と 同じ）',
    (await 画面の字(0, 1)) === ',0', '出 ' + await 画面の字(0, 1));

  /* ② 通貨（¥）を 当てる → C0 */
  await 打つ(1, 0, '1000');                            /* A2 */
  await マスを押す(1, 0);
  T('★リボンの「通貨」を 本物の マウスで 押せた★', await リボンを押す('通貨表示形式'));
  await 打つ(1, 1, '=CELL("format",A2)');              /* B2 */
  T('★★=CELL("format",A2) が C0★★（¥ が 通貨・$ では ない）',
    (await 画面の字(1, 1)) === 'C0', '出 ' + await 画面の字(1, 1));

  /* ③ 右そろえの 字 → prefix は " */
  await 打つ(2, 0, 'あいう');                          /* A3 */
  await マスを押す(2, 0);
  T('★リボンの「右揃え」を 本物の マウスで 押せた★', await リボンを押す('右揃え'));
  await 打つ(2, 1, '=CELL("prefix",A3)');              /* B3 */
  T('★★=CELL("prefix",A3) が "★★（右そろえの 字）',
    (await 画面の字(2, 1)) === '"', '出 ' + JSON.stringify(await 画面の字(2, 1)));

  /* ④ 触っていない 所＝実Excel の まっさらな シートと 同じ */
  await 打つ(3, 1, '=CELL("width",A1)');               /* B4 */
  T('★=CELL("width",A1) が 8★（何も していない 列）',
    (await 画面の字(3, 1)) === '8', '出 ' + await 画面の字(3, 1));
  await 打つ(4, 1, '=CELL("address",A1)');             /* B5 */
  T('★=CELL("address",A1) が $A$1★', (await 画面の字(4, 1)) === '$A$1',
    '出 ' + await 画面の字(4, 1));
  await 打つ(5, 1, '=CELL("type",A3)');                /* B6 */
  T('★=CELL("type",A3) が l★（字）', (await 画面の字(5, 1)) === 'l',
    '出 ' + await 画面の字(5, 1));
  await 打つ(6, 1, '=CELL("type",D9)');                /* B7 */
  T('★=CELL("type",D9) が b★（空）', (await 画面の字(6, 1)) === 'b',
    '出 ' + await 画面の字(6, 1));

  /* ★★繋ぎが 効いているか★★＝見た目を 渡していなければ 表示形式は 既定＝G に なる */
  T('★★book.html の 繋ぎが 効いている（G では なく ,0 が 出た）★★',
    (await 画面の字(0, 1)) !== 'G',
    '★G が 出た＝画面の 見た目が 渡っていない★');

  await page.screenshot({ path: path.join(ここ, 'cell-honmono-de-osu.png') });
  言う('');
  言う('★絵★ docs/measured/kansuu46/cell-honmono-de-osu.png（お客さんに 出る 画面）');
} finally {
  配信.閉じる();
  await browser.close();
}
言う('');
言う('cell-honmono: ' + pass + ' 緑 / ' + fail + ' 赤');
fs.writeFileSync(出す, 紙.join('\n') + '\n', { encoding: 'utf-8' });
console.log('\n★書いた … docs/measured/kansuu46/cell-honmono-de-osu.txt★');
process.exit(fail ? 1 : 0);
