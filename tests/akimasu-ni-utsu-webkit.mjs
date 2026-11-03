/* akimasu-ni-utsu-webkit.mjs ･･･ ★空いて いる マスに 打って 書き出せるか★ 2026-09-26
 *
 *  ★★何が 起きて いたか（経営者1 が お客さんの 道で 押しました・司さんの 実物）★★
 *    ★空いて いる マスに 数を 打って 「書き出す」を 押すと 1冊も 出ません★
 *      帯 ･･･「このファイルは直せません（給料1：
 *            ★そのセルが元のファイルにありません: 1,8★）」
 *    ⇒★本番と 同じ 木でも 同じ＝元から です★
 *    ⇒★★＝お客さんが する 一番 ふつうの 事が 出来ませんでした★★
 *    ⇒司さんの 決め ア「★全部 保存しろや、断る 理由が なんか あるんか★」に 直に 当たる
 *
 *  ★★直し★★
 *    `lib/xlsb-edit.js` に `マスを足す` を 足した
 *      ・行の 頭（25バイト）は ★元の 板の 行の 頭を 写して 行番号だけ 書き換える★
 *        ＝残り 21バイト（高さ・書式 等）を ★当て推量で 作らない★
 *        ＝`xlsb板を作る` が 前から 同じ 手（★前例に 合わせる★）
 *      ・数の マス（12バイト）＝列(4) ＋ 組の番号(4) ＋ RK(4)／組の 番号は 0（飾り 無し）
 *      ・★行は 行番号の 順／マスは 列の 順★ に 入れる
 *      ・★使った 範囲（記録148）も 広げる★
 *      ・★行の 頭が 1本も 無い 板では 足さない★（写す 見本が 無い）
 *
 *  ★★この 見張りが 見る 物★★
 *    ⑴★行は 在るが マスが 無い 所★（見本の B7）に 打って 書き出せるか
 *    ⑵★行ごと 無い 所★（見本の 9行目）に 打って 書き出せるか
 *    ⑶★打った 値が 本の 中に 入って いるか★
 *    ⑷★元から 在った マスが 1つも 変わって いないか★（★足したのに 壊して いない★）
 *    ⇒★実Excel を 使いません★（借り物で 開き直して 読む）
 *
 *  ★材料★ `tests/fixtures/book-open-sample.xlsb`（★repo に 在る 物★）
 *    A1:E1 見出し ／ A2:E6 商品5行 ／ A7 合計 ／ D7 `=SUM(D2:D6)`
 *    ⇒★B7 は 行が 在って マスが 無い★／★8行目から 行ごと 無い★
 *
 *  ★走らせ方★ node tests/akimasu-ni-utsu-webkit.mjs
 */
import path from 'node:path'; import http from 'node:http'; import fs from 'node:fs';
import os from 'node:os';
import { createRequire } from 'node:module';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { borrow, launch } from '../scripts/_borrow-playwright.mjs';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(pathToFileURL(path.join(ROOT, 'package.json')));
const XLSX = require_(path.join(ROOT, 'lib/xlsx.full.min.js'));

let pass = 0, fail = 0;
const T = (n, ok, m) => {
  if (ok) { pass++; console.log('  ok   ' + n); }
  else { fail++; console.log('  NG   ' + n + (m ? '\n       ' + m : '')); }
};

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

const 材料 = path.join(ROOT, 'tests/fixtures/book-open-sample.xlsb');

/** ★本を 開いて 「番地 → 値」に する★（★借り物で 直に 読む＝画面を 通しません★） */
function 値を並べる(道) {
  const wb = XLSX.read(fs.readFileSync(道), { type: 'buffer', cellFormula: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const 出 = {};
  Object.keys(ws).forEach((a1) => {
    if (a1.charAt(0) === '!') return;
    const c = ws[a1];
    出[a1] = (c && c.v !== undefined && c.v !== null) ? String(c.v) : '';
  });
  return 出;
}

/* ★打つ 所（★見本の 形から 決めました★）★
     ⑴B7 ･･･ ★行 7 は 在る（A7 と D7）／B7 は 無い★
     ⑵A9 ･･･ ★9行目は 行ごと 無い★ */
const 場所 = [
  /* ★★B7 は 元から 「空の 記録」が 在りました★★（2026-09-26 わざと 戻して 分かりました）
       ＝`locate` は 空の マス（記録1）も 拾う ので ★「無い」では ありません★
       ⇒★前の 形でも B7 は 書き出せて いました★
       ⇒★★本当に 無いのは 「行ごと 無い」所です★★（下の A9）
       ⇒★だから B7 は 「壊して いないか」を 見る 為に 残します★ */
  { 名: '行は 在るが 空の 記録', 行: 6, 列: 1, 番地: 'B7', 値: '11' },
  { 名: '行ごと 無い', 行: 8, 列: 0, 番地: 'A9', 値: '22' },
];

console.log('[空きマスに 打つ] ★空いて いる マスに 打って 書き出せるか★');
const 元 = 値を並べる(材料);
console.log('  ★材料★ ' + path.basename(材料) + '（マス ' + Object.keys(元).length + '個）');
for (const 所 of 場所) {
  if (元[所.番地] !== undefined) {
    console.log('  ★★' + 所.番地 + ' が すでに 在ります＝この 見張りは 空振りです★★');
    process.exit(8);
  }
}

const wk = await borrow('akimasu-ni-utsu', 'webkit');
const browser = await launch('akimasu-ni-utsu', wk, {}, 'webkit');
const 配信 = await 立てる(ROOT);
const 落とし先 = fs.mkdtempSync(path.join(os.tmpdir(), 'exally-akimasu-'));

try {
  for (const 所 of 場所) {
    const page = await browser.newPage({ viewport: { width: 1280, height: 900 }, acceptDownloads: true });
    try {
      page.on('pageerror', (e) => console.log('      [落ちた] ' + String(e.message).slice(0, 200)));
      await page.goto(配信.url + '/book.html', { waitUntil: 'load', timeout: 120000 });
      await page.evaluate(() => {
        document.body.classList.remove('exally-locked');
        const ov = document.getElementById('loginOv');
        if (ov) { ov.classList.remove('open'); ov.style.display = 'none'; }
      });
      await page.setInputFiles('#bookFileInput', 材料);
      await page.waitForFunction(() => {
        const ss = window.sheets || [];
        return ss.some((s) => s && s.data && Object.keys(s.data).length > 0);
      }, null, { timeout: 120000 });
      await page.waitForTimeout(600);

      const 無い = await page.evaluate(() => ['setCell', 'saveOpenedBook'].filter((n) => typeof window[n] !== 'function'));
      if (無い.length) { console.log('  ★★測れません★★ 在りません ... ' + 無い.join(' / ')); process.exit(8); }

      console.log('      ══ ★' + 所.名 + '★（' + 所.番地 + ' に ' + 所.値 + '）══');
      await page.evaluate(([r, c, v]) => { window.setCell(r, c, v); }, [所.行, 所.列, 所.値]);
      await page.waitForTimeout(1000);

      const 落とし物 = page.waitForEvent('download', { timeout: 120000 }).catch(() => null);
      await page.evaluate(() => { window.saveOpenedBook(); });
      await page.waitForTimeout(1200);
      const 窓 = await page.evaluate(() => {
        const go = document.getElementById('diffGo');
        return go && go.offsetWidth > 0 ? String(go.textContent || '').trim() : '';
      });
      if (!窓) { console.log('        ★★窓が 出て いません＝この 見張りは 空振りです★★'); process.exit(8); }
      console.log('        窓 ... ' + JSON.stringify(窓));
      await page.click('#diffGo');

      const dl = await 落とし物;
      const 帯 = await page.evaluate(() => {
        const t = document.querySelector('.toast, #toast, [class*="toast"]');
        return t ? String(t.textContent || '').trim().slice(0, 200) : '';
      });
      console.log('        落とし物 ... ' + (dl ? '★来た★' : '★来ません★')
        + ' ／ 帯 ' + JSON.stringify(帯.slice(0, 110)));
      T('★★' + 所.名 + ' に 打っても 書き出せる★★（' + 所.番地 + '）', !!dl,
        '帯 ' + JSON.stringify(帯.slice(0, 160)));
      if (!dl) continue;

      const 出た = path.join(落とし先, '出-' + 所.番地 + '.xlsb');
      await dl.saveAs(出た);
      const 後 = 値を並べる(出た);
      console.log('        出た 本 ... ' + fs.statSync(出た).size.toLocaleString() + ' バイト'
        + ' ／ マス ' + Object.keys(後).length + '個');
      console.log('        ' + 所.番地 + ' ... ' + JSON.stringify(後[所.番地]));
      T('★★打った 値が 本の 中に 入って いる★★（' + 所.番地 + ' ＝ ' + 所.値 + '）',
        String(後[所.番地]) === String(Number(所.値)), 所.番地 + ' ' + JSON.stringify(後[所.番地]));

      /* ★元から 在った マスを 壊して いないか★（★足したのに 他を 動かして いない★） */
      const 壊れた = Object.keys(元).filter((k) => String(元[k]) !== String(後[k] === undefined ? '' : 後[k]));
      console.log('        元から 在った マスで 変わった 物 ... ' + 壊れた.length + '個'
        + (壊れた.length ? '  ' + 壊れた.slice(0, 8).join(' / ') : ''));
      T('★★元から 在った マスを 1つも 壊して いない★★',
        壊れた.length === 0, '変わった ' + 壊れた.length + '個 ' + 壊れた.slice(0, 8).join(' / '));
    } finally {
      await page.close().catch(() => {});
    }
  }
} finally {
  await browser.close().catch(() => {});
  配信.閉じる();
}

console.log('');
console.log('akimasu-ni-utsu: ' + pass + ' 緑 / ' + fail + ' 赤');
process.exit(fail ? 1 : 0);
