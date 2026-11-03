/* isref-mada-ga-moreru-ka.mjs — ★作り物の 名前が お客さんの 外へ 漏れないか★（2026-09-07）
 *
 *  ★★なぜ 要るか★★
 *    ISREF の「まだ 出せない 形」を `ISREF.MADA()` に 置き換えました。
 *    ⇒★これは ★Excel に 無い 名前★です★
 *    ⇒★★もし これが お客さんの ファイルに 残ったら
 *       ★開くたびに 壊れた 式が 出る★＝★私たちが 壊した 事に なります★★★
 *    ⇒★『たぶん 漏れません』では 出しません★（指示役 2026-09-07）
 *
 *  ★★お客さんの 道で 3か所 見ます★★
 *    ①★画面の 数式バー★ … 打った 式が ★元の まま★ 見えるか
 *    ②★★Excel に 書き出した .xlsx の 中身★★ ←★一番 大事★
 *       ⇒ 実際に 書き出して ★ファイルを 開いて 式を 読む★
 *    ③★控え（cell.f）★ … 画面が 持っている 式が 元の ままか
 *       （AI にも 書き出しにも ここが 渡る）
 *
 *  ★★押し方は お客さんと 同じ★★
 *    ・マスは ★本物の マウス★／式は ★数式バー★に 打つ
 *    ・書き出しは ★「Excelに書き出す」を 本物の マウスで 押す★
 *    ⇒★JS で 関数を 直に 呼びません★
 *
 *  ★外へは 1回も 出ません★（127.0.0.1 だけ）
 *
 *  使い方: node docs/measured/kansuu46/isref-mada-ga-moreru-ka.mjs
 */
import path from 'node:path'; import http from 'node:http'; import fs from 'node:fs';
import os from 'node:os';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';
import { borrow, launch } from '../../../scripts/_borrow-playwright.mjs';

const ここ = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(ここ, '..', '..', '..');
const require_ = createRequire(path.join(ROOT, 'package.json'));
const 出す = path.join(ここ, 'isref-mada-ga-moreru-ka.txt');
const 紙 = [];
const 言う = (s) => { 紙.push(s); console.log(s); };
let pass = 0, fail = 0;
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

/* ★打つ 式★＝★まだ 出せない 形★（ここが `ISREF.MADA` に 化ける） */
const 打つ式 = [
  '=ISREF(INDEX(A1:B2,1,1))',
  '=ISREF(INDIRECT("A1"))',
  '=IF(ISREF(SUM(A1:A2)),1,2)',
  '=ISREF(A1)',                 /* ★これは 出せる 形★（TRUE に なる） */
];

言う('# ★作り物の 名前（ISREF.MADA）が お客さんの 外へ 漏れないか★（'
  + new Date().toISOString().slice(0, 10) + '）');
言う('');
言う('★外へは 1回も 出ていません（127.0.0.1 だけ）★');
言う('★押し方★ マスは 本物の マウス／式は 数式バー／書き出しは 本物の マウスで ボタン');
言う('');

const cr = await borrow('isref-mada', 'chromium');
const browser = await launch('isref-mada', cr, { channel: 'chrome' }, 'chromium');
const page = await browser.newPage({ viewport: { width: 1600, height: 1000 }, acceptDownloads: true });
const 配信 = await 立てる(ROOT);
const 仮 = fs.mkdtempSync(path.join(os.tmpdir(), 'isref-'));
try {
  await page.goto(配信.url + '/book.html', { waitUntil: 'load' });
  await page.evaluate(() => {
    document.body.classList.remove('exally-locked');
    const ov = document.getElementById('loginOv');
    if (ov) { ov.classList.remove('open'); ov.style.display = 'none'; }
  });
  await page.waitForTimeout(1500);

  const 枠 = await page.evaluate(() => {
    const c = document.getElementById('grid-canvas');
    const r = c.getBoundingClientRect();
    return { x: r.left, y: r.top, HDR_W: window.HDR_W, HDR_H: window.HDR_H, ROW_H: window.ROW_H, COL_W: window.COL_W };
  });
  const マスを押す = async (行, 列) => {
    await page.mouse.click(枠.x + 枠.HDR_W + 列 * 枠.COL_W + 枠.COL_W / 2,
      枠.y + 枠.HDR_H + 行 * 枠.ROW_H + 枠.ROW_H / 2);
    await page.waitForTimeout(120);
  };

  /* ★中身を 置く★（INDEX が 何かを 指せる ように） */
  for (let r = 0; r < 2; r++) {
    for (let c = 0; c < 2; c++) {
      await マスを押す(r, c);
      await page.fill('#formula-input', String((r + 1) * 10 + (c + 1)));
      await page.press('#formula-input', 'Enter');
      await page.waitForTimeout(150);
    }
  }
  /* ★式を 打つ★（D列に 縦に） */
  for (let i = 0; i < 打つ式.length; i++) {
    await マスを押す(i, 3);
    await page.fill('#formula-input', 打つ式[i]);
    await page.press('#formula-input', 'Enter');
    await page.waitForTimeout(250);
  }

  /* ── ①画面の 数式バー ── */
  const バー = [];
  for (let i = 0; i < 打つ式.length; i++) {
    await マスを押す(i, 3);
    await page.waitForTimeout(150);
    バー.push(await page.inputValue('#formula-input'));
  }
  言う('  ★①数式バーに 出る 字★');
  for (let i = 0; i < 打つ式.length; i++) 言う('    打った ' + 打つ式[i] + '  →  出た ' + バー[i]);
  T('★①数式バーが 元の 式の まま（作り物の 名前が 出ない）★',
    バー.every((s, i) => s === 打つ式[i]),
    バー.map((s, i) => (s === 打つ式[i] ? '' : '★' + 打つ式[i] + ' → ' + s + '★')).filter(Boolean).join('\n       '));

  /* ── ③控え（画面が 持っている 式）＝AI にも 書き出しにも ここが 渡る ── */
  const 控え = await page.evaluate(() => {
    const d = window.sheets[window.activeSheet].data || {};
    const 出 = [];
    for (let r = 0; r < 4; r++) { const m = d[r + ',3']; if (m) 出.push({ f: m.f, d: m.d, v: m.v }); }
    return 出;
  });
  言う('');
  言う('  ★③画面が 持っている 式と 出た 答え★');
  for (const x of 控え) 言う('    式 ' + x.f + '  →  画面 ' + JSON.stringify(x.d !== undefined ? x.d : x.v));
  T('★③控えに 作り物の 名前が 無い★', !控え.some((x) => /ISREF\.MADA/i.test(String(x.f))),
    JSON.stringify(控え.map((x) => x.f)));
  T('★★まだ 出せない 形は #NAME? に なっている（黙って 逆の 答えを 出していない）★★',
    控え.slice(0, 3).every((x) => /#NAME\?/.test(String(x.d !== undefined ? x.d : x.v))),
    控え.map((x) => x.f + ' → ' + (x.d !== undefined ? x.d : x.v)).join('\n       '));
  T('★出せる 形（=ISREF(A1)）は TRUE のまま★',
    控え.length > 3 && /TRUE|true/i.test(String(控え[3].d !== undefined ? 控え[3].d : 控え[3].v)),
    JSON.stringify(控え[3]));

  /* ── ②★書き出した .xlsx の 中身★（一番 大事） ── */
  const 待つ = page.waitForEvent('download', { timeout: 60000 });
  const 押せた = await (async () => {
    const 印 = page.locator('button, a').filter({ hasText: 'Excelに書き出す' });
    const n = await 印.count();
    for (let i = 0; i < n; i++) {
      const b = 印.nth(i);
      if (await b.isVisible()) { await b.click(); return true; }
    }
    return false;
  })();
  T('★「Excelに書き出す」を 本物の マウスで 押せた★', 押せた);
  let 道 = null;
  try {
    const dl = await 待つ;
    道 = path.join(仮, dl.suggestedFilename() || 'out.xlsx');
    await dl.saveAs(道);
  } catch (e) { 道 = null; }
  T('★書き出した ファイルを 受け取れた★', !!道 && fs.existsSync(道), String(道));

  if (道 && fs.existsSync(道)) {
    const XLSX = require_(path.join(ROOT, 'lib/xlsx.full.min.js'));
    const wb = XLSX.read(fs.readFileSync(道), { type: 'buffer' });
    const 式たち = [];
    for (const 名 of wb.SheetNames) {
      const ws = wb.Sheets[名];
      for (const a of Object.keys(ws)) {
        if (a.charAt(0) === '!') continue;
        if (ws[a] && ws[a].f) 式たち.push(名 + '!' + a + '  =' + ws[a].f);
      }
    }
    言う('');
    言う('  ★②書き出した .xlsx の 中の 式（' + 式たち.length + '本）★');
    for (const x of 式たち) 言う('    ' + x);
    T('★★書き出した .xlsx に 作り物の 名前が 1つも 無い★★',
      !式たち.some((s) => /ISREF\.MADA/i.test(s)),
      '★お客さんの ファイルに 化け物が 残ります★\n       ' + 式たち.filter((s) => /ISREF\.MADA/i.test(s)).join('\n       '));
    T('★書き出した .xlsx に 打った 式が 残っている（空では ない）★', 式たち.length >= 4,
      '式 ' + 式たち.length + '本＝★少なすぎる＝物差しが 空洞★');
    /* ★生の 中身も 見る★（SheetJS が 読み飛ばした 所に 隠れていないか） */
    const 生 = fs.readFileSync(道);
    const 生に在る = 生.indexOf(Buffer.from('ISREF.MADA', 'utf-8')) >= 0;
    T('★★ファイルの 生の バイトにも 作り物の 名前が 無い★★', !生に在る,
      '★ZIP の 中の どこかに 残っています★');
    言う('    ★ファイルの 大きさ … ' + 生.length + 'バイト★');
  }
} finally {
  配信.閉じる();
  await browser.close();
  try { fs.rmSync(仮, { recursive: true, force: true }); } catch (e) { /* 手元の 仮 */ }
}
言う('');
言う('isref-mada: ' + pass + ' 緑 / ' + fail + ' 赤');
fs.writeFileSync(出す, 紙.join('\n') + '\n', { encoding: 'utf-8' });
console.log('\n★書いた … docs/measured/kansuu46/isref-mada-ga-moreru-ka.txt★');
process.exit(fail ? 1 : 0);
