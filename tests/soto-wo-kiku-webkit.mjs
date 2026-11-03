/* soto-wo-kiku-webkit.mjs ･･･ ★外へ つながる 式が 在る 本だけ お客さんに 訊く★ 2026-09-26
 *
 *  ★★なぜ 要るか★★
 *    `lib/formula-soto.js` の 頭に こう 書いて あります:
 *      「もらった `.xlsx` の 式は ★式のまま★ 入る＝★開いた だけで 走る★」
 *      「★画面から 直に 外へ 出さない／行き先は うちが 決める★」
 *    ⇒★知らない 人の 本を 開いた だけで 外へ 通信します★
 *    ⇒司さん 09-25「★おすすめで 直せ★」／経営者1 の 推し ⑶
 *    ⇒★★その 4つが 在る 本だけ お客さんに 訊く★★
 *      （`WEBSERVICE` ／ `STOCKHISTORY` ／ `TRANSLATE` ／ `DETECTLANGUAGE`）
 *
 *  ★★この 見張りが 見る 物★★
 *    ⑴★4つが 0個の 本では 窓を 出さない★（★司さんの 実物は 0個★＝★邪魔を しない★）
 *    ⑵★1個でも 在れば 窓を 出す★／★字に 「インターネット」と 書いて ある★
 *    ⑶★「つながない」を 押すと 1回も 外へ 出ない★（★便の 呼ばれた 数が 0★）
 *    ⑷★「つないで よい」を 押すと 外へ 出る★（★同じ 本・同じ 道★）
 *    ⇒★本物の インターネットへは 出しません★＝`道具を入れる` で 作り物を 入れて 数えます
 *
 *  ★走らせ方★ node tests/soto-wo-kiku-webkit.mjs
 */
import path from 'node:path'; import http from 'node:http'; import fs from 'node:fs';
import os from 'node:os';
import { createRequire } from 'node:module';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { borrow, launch } from '../scripts/_borrow-playwright.mjs';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(pathToFileURL(path.join(ROOT, 'package.json')));
const XLSX = require_(path.join(ROOT, 'lib/xlsx.full.min.js'));
const F = require_(path.join(ROOT, 'lib/formula-soto.js'));

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

/* ★4つの 名は `lib/formula-soto.js` から 貰います★＝★ここで 書き足しません★
     （★名簿を 2か所に 持つと 片方だけ 増えます★） */
const 四つ = (typeof F.足した名前 === 'function' ? F.足した名前() : null);

/* ══ ★材料★ ══（`.xlsx`・★答えを わざと 保存して おきます★）
     ⑴外 無し ... `=A1*2` だけ
     ⑵外 在り ... `=WEBSERVICE("https://example.com/a.txt")` を 1本 */
function 作る(外を入れるか) {
  const ws = XLSX.utils.aoa_to_sheet([[2, null, null]]);
  ws['B1'] = { t: 'n', f: 'A1*2', v: 4 };
  if (外を入れるか) ws['C1'] = { t: 's', f: 'WEBSERVICE("https://example.com/a.txt")', v: 'まえ' };
  ws['!ref'] = 'A1:C1';
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'あ');
  const 道 = path.join(os.tmpdir(), 'exally-soto-' + (外を入れるか ? 'ari' : 'nashi') + '.xlsx');
  fs.writeFileSync(道, XLSX.write(wb, { bookType: 'xlsx', type: 'buffer' }));
  return 道;
}
const 外無し = 作る(false);
const 外在り = 作る(true);

console.log('[外へ つなぐか 訊く] ★4つが 在る 本だけ 訊く★');
console.log('  ★4つの 名（`lib/formula-soto.js` から）★ ... '
  + (四つ ? 四つ.join(' / ') : '★取れません★'));
if (!四つ || 四つ.length !== 4) {
  console.log('  ★★4つの 名が 取れません＝この 見張りは 空振りです★★');
  process.exit(8);
}

const wk = await borrow('soto-wo-kiku', 'webkit');
const browser = await launch('soto-wo-kiku', wk, {}, 'webkit');
const 配信 = await 立てる(ROOT);

/** ★本を 開いて 窓の 姿と 外へ 出た 数を 返す★
 *  押し ... null＝押さない ／ 'よい'＝つないで よい ／ 'だめ'＝つながない */
async function 開いて見る(本, 押し) {
  const page = await browser.newPage({ viewport: { width: 1280, height: 900 } });
  /* ══ ★★通信そのものを 数えます★★ ══（2026-09-26 ★数え方を 1回 間違えました★）
       ★初めは `道具を入れる` で 作り物を 入れて 数えました★
       ⇒★本を 開くと 画面が `つなぐ(...)` で 本物の 道具に 上書きします★
       ⇒★★私の 作り物は 使われず 「0回」と 出ました★★（★出て いたのに★）
       ⇒★だから 「網の 口」で 数えます＝上書きされません★
       ⇒★本物の インターネットへは 出しません＝ここで 止めて 作り物を 返します★ */
  let 外へ出た = 0;
  await page.route('**/api/soto**', async (route) => {
    外へ出た++;
    await route.fulfill({ status: 200, contentType: 'application/json',
      body: JSON.stringify({ ok: true, text: '作り物' }) });
  });
  await page.route('**/api/ai**', async (route) => {
    外へ出た++;
    await route.fulfill({ status: 200, contentType: 'application/json',
      body: JSON.stringify({ ok: true, 言葉: '作り物' }) });
  });
  try {
    page.on('pageerror', (e) => console.log('      [落ちた] ' + String(e.message).slice(0, 200)));
    await page.goto(配信.url + '/book.html', { waitUntil: 'load', timeout: 120000 });
    await page.evaluate(() => {
      document.body.classList.remove('exally-locked');
      const ov = document.getElementById('loginOv');
      if (ov) { ov.classList.remove('open'); ov.style.display = 'none'; }
      /* ★数えるのは 網の 口です★（上の 断りを 見て ください） */
    });
    await page.setInputFiles('#bookFileInput', 本);
    await page.waitForFunction(() => {
      const ss = window.sheets || [];
      return ss.some((s) => s && s.data && Object.keys(s.data).length > 0);
    }, null, { timeout: 120000 });
    await page.waitForTimeout(900);

    const 窓 = await page.evaluate(() => {
      const ov = document.getElementById('diffOverlay');
      const go = document.getElementById('diffGo');
      const head = document.getElementById('diffHead');
      const body = document.getElementById('diffBody');
      const 見える = !!(ov && ov.offsetWidth > 0 && ov.offsetHeight > 0);
      return {
        見える: 見える,
        頭: head ? String(head.textContent || '').trim() : '',
        字: body ? String(body.textContent || '').trim() : '',
        ボタン: go ? String(go.textContent || '').trim() : '',
        許し: (window.FormulaSotoPlug && typeof FormulaSotoPlug.外へ出してよいか === 'function')
          ? FormulaSotoPlug.外へ出してよいか() : 'わからない',
      };
    });

    if (押し === 'よい') await page.click('#diffGo');
    if (押し === 'だめ') await page.click('#diffCancel');
    if (押し) await page.waitForTimeout(1500);

    const 後 = await page.evaluate(() => {
      const sh = (window.sheets || [])[0] || {};
      const c = (sh.data || {})['0,2'] || null;
      return {
        許し: (window.FormulaSotoPlug && typeof FormulaSotoPlug.外へ出してよいか === 'function')
          ? FormulaSotoPlug.外へ出してよいか() : 'わからない',
        /* ★どこで 止まって いるかを 見る 為★ */
        マスの式: c ? String(c.f || '') : '（無し）',
        マスの答え: c ? String(c.d === undefined ? '' : c.d) : '（無し）',
        覚えの数: (window.FormulaSotoPlug && typeof FormulaSotoPlug.覚えの数 === 'function')
          ? FormulaSotoPlug.覚えの数() : -1,
      };
    });
    後.外へ出た = 外へ出た;   /* ★網の 口で 数えた 数★ */
    return { 窓: 窓, 後: 後 };
  } finally {
    await page.close().catch(() => {});
  }
}

try {
  /* ══ ⑴4つが 0個の 本 ══ */
  const な = await 開いて見る(外無し, null);
  console.log('      ── 実測 ── ★外 0個の 本★ ... 窓が 見える ' + な.窓.見える
    + ' ／ 許し ' + JSON.stringify(な.窓.許し) + ' ／ 外へ出た ' + な.後.外へ出た + '回');
  T('★★4つが 0個の 本では 窓を 出さない★★（★司さんの 実物は 0個＝邪魔を しない★）',
    な.窓.見える === false, '窓 ' + JSON.stringify(な.窓.頭));
  T('★4つが 0個の 本では 1回も 外へ 出ない★', な.後.外へ出た === 0, String(な.後.外へ出た));

  /* ══ ⑵4つが 在る 本＝窓が 出るか ══ */
  const あ = await 開いて見る(外在り, null);
  console.log('      ── 実測 ── ★外 1個の 本★ ... 窓が 見える ' + あ.窓.見える);
  console.log('        窓の 頭 ... ' + JSON.stringify(あ.窓.頭));
  console.log('        ボタン ... ' + JSON.stringify(あ.窓.ボタン));
  T('★★4つが 在る 本では 窓を 出す★★', あ.窓.見える === true, '窓 ' + JSON.stringify(あ.窓.頭));
  T('★窓の 字に 「インターネット」と 書いて ある★（★言わずに 訊かない★）',
    あ.窓.字.indexOf('インターネット') >= 0, '字 ' + JSON.stringify(あ.窓.字.slice(0, 120)));
  T('★答える 前は 許して いない★（`null`）', あ.窓.許し === null, JSON.stringify(あ.窓.許し));
  T('★答える 前は 1回も 外へ 出て いない★', あ.後.外へ出た === 0, String(あ.後.外へ出た));

  /* ══ ⑶「つながない」を 押す ══ */
  const だ = await 開いて見る(外在り, 'だめ');
  console.log('      ── 実測 ── ★つながない を 押した★ ... 許し ' + JSON.stringify(だ.後.許し)
    + ' ／ 外へ出た ' + だ.後.外へ出た + '回');
  T('★★「つながない」を 押すと 1回も 外へ 出ない★★', だ.後.外へ出た === 0, String(だ.後.外へ出た));
  T('★「つながない」を 押すと 許しは false★', だ.後.許し === false, JSON.stringify(だ.後.許し));

  /* ══ ⑷「つないで よい」を 押す ══ */
  const よ = await 開いて見る(外在り, 'よい');
  console.log('      ── 実測 ── ★つないで よい を 押した★ ... 許し ' + JSON.stringify(よ.後.許し)
    + ' ／ 外へ出た ' + よ.後.外へ出た + '回 ／ 覚え ' + よ.後.覚えの数 + '件');
  console.log('        マスの 式 ... ' + JSON.stringify(よ.後.マスの式));
  console.log('        マスの 答え ... ' + JSON.stringify(よ.後.マスの答え));
  T('★「つないで よい」を 押すと 許しは true★', よ.後.許し === true, JSON.stringify(よ.後.許し));
  T('★★「つないで よい」を 押すと 外へ 出る★★（★門が 締まった ままでは ない★）',
    よ.後.外へ出た >= 1, String(よ.後.外へ出た));
} finally {
  await browser.close().catch(() => {});
  配信.閉じる();
}

console.log('');
console.log('soto-wo-kiku: ' + pass + ' 緑 / ' + fail + ' 赤');
process.exit(fail ? 1 : 0);
