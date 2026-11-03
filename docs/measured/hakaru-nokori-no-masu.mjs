/* hakaru-nokori-no-masu.mjs — ★名指しの マスを 開いて 中を 見る★ 2026-09-25
 *
 *  ★★なぜ 要るか★★
 *    実Excel と 違う ★116個★ の うち
 *      ★68個★ ･･･ 実Excel は `#REF!`・うちは 数（★お客さんは 間違いに 気付けません★）
 *      ★47個★ ･･･ 実Excel は 空・うちは `0 円` の ような 書式つきゼロ
 *    ⇒★私の 数え方では 47個が 出ず 9個でした★（★板が 違いました★）
 *    ⇒経営者1 が ★居場所だけ★ を くれました（★値は 1つも 出して いません★）
 *    ⇒★名指しで 開くのが 一番 早い★
 *
 *  ★★出す 物★★（★司さんの 実物の 中身は 出しません★）
 *    ・板の 名 ／ `行,列` ／ ★値の 形（数字は 9 に 伏せる）★
 *    ・`numFmt` ／ `ゼロを隠す` ／ `_ゼロを隠すか` の 答え ／ 画面に 出る 字の 形
 *    ・★式は 形だけ★（数字を 9 に 伏せる）
 *
 *  ★読むだけ★＝1バイトも 書きません。
 *
 *  走らせ方:
 *    node docs/measured/hakaru-nokori-no-masu.mjs <材料> 板|行,列 [板|行,列 ...]
 */
import path from 'node:path'; import http from 'node:http'; import fs from 'node:fs';
import { fileURLToPath } from 'node:url';
import { borrow, launch } from '../../scripts/_borrow-playwright.mjs';
import { 道を確かめる, 道の字, 効いた数を出す } from './_gamen-no-michi.mjs';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '../..');
const 引 = process.argv.slice(2);
const 材料 = 引[0];
const 場所 = 引.slice(1);
if (!材料 || !場所.length) {
  console.log('使い方: node docs/measured/hakaru-nokori-no-masu.mjs <材料> 板|行,列 ...');
  process.exit(1);
}

function 立てる(root) {
  const 型 = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8',
    '.css': 'text/css; charset=utf-8', '.json': 'application/json; charset=utf-8',
    '.png': 'image/png', '.svg': 'image/svg+xml', '.ico': 'image/x-icon' };
  const s = http.createServer((q, r) => {
    const f = path.join(root, decodeURIComponent(String(q.url).split('?')[0]).replace(/^\/+/, ''));
    if (!f.startsWith(root) || !fs.existsSync(f) || fs.statSync(f).isDirectory()) { r.statusCode = 404; return r.end('no'); }
    r.setHeader('content-type', 型[path.extname(f).toLowerCase()] || 'application/octet-stream');
    fs.createReadStream(f).pipe(r);
  });
  return new Promise((x) => s.listen(0, '127.0.0.1', () => x({
    url: 'http://127.0.0.1:' + s.address().port, 閉じる: () => s.close(),
  })));
}

if (!fs.existsSync(材料)) { console.log('★材料が 有りません★ ' + 材料); process.exit(1); }
console.log('★名指しの マスを 開いて 中を 見る★');
console.log('  材料 ... ' + path.basename(材料));
console.log('  見る マス ... ' + 場所.length + '個');

const wk = await borrow('nokori', 'webkit');
const browser = await launch('nokori', wk, {}, 'webkit');
const page = await browser.newPage({ viewport: { width: 1400, height: 900 } });
const 配信 = await 立てる(ROOT);

try {
  await page.goto(配信.url + '/book.html', { waitUntil: 'load', timeout: 180000 });
  await page.evaluate(() => {
    document.body.classList.remove('exally-locked');
    const ov = document.getElementById('loginOv');
    if (ov) { ov.classList.remove('open'); ov.style.display = 'none'; }
    window.FileOut = { deliver: () => Promise.reject(new Error('★書き出しては いけません★')) };
  });
  await page.setInputFiles('#bookFileInput', 材料);
  await page.waitForFunction(() => {
    const ss = window.sheets || [];
    for (let k = 0; k < ss.length; k++) {
      if (ss[k] && ss[k].data && Object.keys(ss[k].data).length > 0) return true;
    }
    return false;
  }, null, { timeout: 900000 });
  await page.waitForTimeout(600);

  /* ★★決め（2026-09-25・2人で 決めました）★★
       ★画面の 事を 測る 時は 画面の 関数を 呼ぶ。真似ない★
       ⇒★要る 関数が 1つでも 無ければ 数を 出さずに 止まります★ */
  await 道を確かめる(page);

  let 隠した数 = 0;
  for (const 所 of 場所) {
    const 割 = String(所).split('|');
    const 板名 = 割[0], 印 = 割[1];
    const 出 = await page.evaluate(([な, k]) => {
      const i = (window.sheets || []).findIndex((s) => s.name === な);
      if (i < 0) return { だめ: '板が 見つかりません' };
      /* ★板を 合わせます★＝`_ゼロを隠すか` も `cW` も `activeSheet` を 見ます */
      window.switchSheet(i);
      const sh = window.sheets[i];
      const cell = (sh.data || {})[k];
      if (!cell) return { だめ: 'マスが 空（`data` に 無い）', 隠す板: !!sh.ゼロを隠す };
      const raw = window._字の元(cell);
      const 生 = String(raw === undefined || raw === null ? '' : raw);
      const 隠す = window._ゼロを隠すか(cell, raw);
      const w = window.cW(+String(k).split(',')[1]);
      let 画面 = '';
      if (隠す) 画面 = '(隠す＝何も 描かない)';
      else if (window._答えは字か(cell)) 画面 = 生;
      else 画面 = String(window.fmtForDisplay(raw, cell.numFmt, window._入る字数(w, raw, cell.numFmt)));
      return {
        隠す板: !!sh.ゼロを隠す,
        vの型: typeof cell.v, vが空: cell.v === '',
        式か: !!cell.f, d字: !!cell.d字,
        生形: 生.replace(/[0-9]/g, '9'), 生長: 生.length,
        ぴったり0: Number(生) === 0 && 生.trim() !== '',
        書式: cell.numFmt || '(無し)',
        隠すと判じた: !!隠す,
        画面形: String(画面).replace(/[0-9]/g, '9'),
        式形: cell.f ? String(cell.f).replace(/[0-9]/g, '9').slice(0, 90) : '',
        幅: w,
      };
    }, [板名, 印]);
    console.log('');
    console.log('  ════ ' + 所);
    if (出.だめ) { console.log('    ★' + 出.だめ + '★' + (出.隠す板 !== undefined ? ' ／ 隠す板 ' + 出.隠す板 : '')); continue; }
    console.log('    隠す板 ' + 出.隠す板 + ' ／ 式か ' + 出.式か + ' ／ d字 ' + 出.d字
      + ' ／ vの型 ' + 出.vの型 + '(空' + 出.vが空 + ')');
    console.log('    生の 形 ... ' + 出.生形 + '（長さ ' + 出.生長 + '／ぴったり0 ' + 出.ぴったり0 + '）');
    console.log('    書式 ..... ' + 出.書式 + ' ／ 列の 幅 ' + 出.幅);
    if (出.隠すと判じた) 隠した数++;
    console.log('    ★隠すと 判じた ... ' + 出.隠すと判じた + '★');
    console.log('    ★画面の 形 ....... ' + 出.画面形 + '★');
    if (出.式形) console.log('    式の 形 ... ' + 出.式形);
  }
  console.log('');
  効いた数を出す('隠すと 判じた', 隠した数, 場所.length);
} finally {
  await page.close().catch(() => {});
  await browser.close().catch(() => {});
  配信.閉じる();
}
