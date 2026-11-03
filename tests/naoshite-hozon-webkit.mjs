/* naoshite-hozon-webkit.mjs ･･･ ★お客さんが 入力値を 直して 書き出せるか★ 2026-09-25
 *
 *  ★★何が 起きて いたか（経営者1 が お客さんの 道で 押しました・司さんの 実物）★★
 *    ★入力値を 1つ 直して 「書き出す」を 押すと 出ません★
 *      窓 ･･･「この297か所を直して 書き出す」⇒ 押す
 *      帯 ･･･「書き出しませんでした／
 *            ★この数式セルの答えの形は まだ直せません（記録 8）★」
 *    ⇒★本番と 同じ 木（15cc377）でも 1字も 違わず 出ません★＝★元から です★
 *    ⇒★★＝「直して 保存する」が 出来ませんでした★★（司さんの 決め ア に 直に 当たる）
 *
 *  ★★因★★
 *    記録 8 ＝ `BrtFmlaString`（★答えが 字の 式★）
 *    ＝数を 入れると 記録の 長さが 変わる ので 書き換えられない
 *    ⇒`lib/xlsb-edit.js` が ★投げて いました★ ⇒ ★1マスの 為に 本 1冊が 出ない★
 *
 *  ★★直し★★
 *    ⑴★投げずに 「触れなかった」と 数えて 先へ 進む★
 *    ⑵★触れなかった 物が 在れば `workbook.bin` に 「開いたら 全部 計算しろ」の 印を 立てる★
 *      （記録157 の 26バイト目 ビット0 ＝ 経営者1 が 実Excel で 7冊 作って 決めた 所）
 *    ⇒★実Excel が 開いた 時に 計算し直す ので 古い 答えは 直ります★
 *
 *  ★★見る 所を 分けて います（★どちらも 要ります★）★★
 *    ★この 見張り（お客さんの 道）★
 *      ＝画面で 1マス 打って 窓の ボタンを 押し、★本が 出るか／打った 値が 入ったか★
 *      ＝★見本の E2 は 「･･･円」＝数に 見えない ので 書く 手前で 外れます★
 *        ⇒★だから ここでは 記録8 の 段に 入りません★
 *    ★`tests/book-open.test.mjs`（単体）★
 *      ＝★記録8 の マスを 数に して 保存を 呼ぶ★
 *      ＝★本が 出るか／記録8 を 1バイトも 触って いないか／印を 立てたか★
 *    ⇒★★「触れなかった 時に 印を 立てる」は 単体が 見ます★★
 *    ⇒★★「お客さんが 直して 出せる」は こちらが 見ます★★
 *
 *  ★★この 見張りが 見る 物★★
 *    ⑴★答えが 字の 式が 在る 本で、入力値を 直して 書き出せるか★
 *    ⑵★打った 値が 本の 中に 入って いるか★
 *    ⑶★「開いたら 全部 計算しろ」の 印が 立って いるか★
 *    ⇒★`.xlsb` で 見ます★（★司さんの 実物が `.xlsb`★）
 *
 *  ★走らせ方★ node tests/naoshite-hozon-webkit.mjs
 */
import path from 'node:path'; import http from 'node:http'; import fs from 'node:fs';
import os from 'node:os';
import { createRequire } from 'node:module';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { borrow, launch } from '../scripts/_borrow-playwright.mjs';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(pathToFileURL(path.join(ROOT, 'package.json')));
const XLSX = require_(path.join(ROOT, 'lib/xlsx.full.min.js'));
const JT = require_(path.join(ROOT, 'lib/xlsb-jitai.js'));
const XE = require_(path.join(ROOT, 'lib/xlsb-edit.js'));
const ZipSurgeon = require_(path.join(ROOT, 'lib/zip-surgeon.js'));

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

/* ══ ★材料★ ══（★`.xlsb` で 作ります＝記録 8 は `.xlsb` の 話★）
     A1 = 1（★お客さんが 直す マス★）
     B1 = `=A1*2`  ･･･ ★答えが 数の 式（記録 9）★
     C1 = `=TEXT(A1,"0")` ･･･ ★答えが 字の 式（記録 8）★  ←★ここが 前は 全部を 止めて いた★
     ⇒★どちらも A1 に つられます★ */
/* ══ ★★見張りの 材料が まだ 作れません★★ ══（2026-09-25）
     借り物（SheetJS）は ★`.xlsb` に 式の 記録を 書きません★
     ⇒作った 見本は ★記録8 が 0個・記録9 も 0個★ ＝★この 見張りは 空振りに なります★
     ⇒★記録の 形を 自分で 作るのは しません★（★当て推量で 記録を 作らない★）
     ⇒★★実Excel で 作った 小さい 見本を 経営者1 に 頼みました★★
        （A1=1 ／ B1=`=A1*2` ／ C1=`=TEXT(A1,"0")` の `.xlsb` 1本）
     ⇒それが 来るまで は ★`--本 <道>` で 実物を 当てて 確かめます★
     ⇒★だから まだ `tests/run.js` に 入れて いません★（★空振りの 緑を 作らない★） */
const 引数 = process.argv.slice(2);
const 取る = (名, 既定) => { const i = 引数.indexOf('--' + 名); return i >= 0 ? 引数[i + 1] : 既定; };
const 外の本 = 取る('本', null);
/* ★★repo に 材料が 在りました★★（2026-09-25・`tests/book-open.test.mjs` が 使って います）
     `tests/fixtures/book-open-sample.xlsb` ･･･ ★記録8 が 5個・記録9 が 7個★
     ⇒★作る 前に 探せ＝借り物では 作れない 物が すでに 在りました★
     ⇒★`--本` で 実物も 当てられます★（司さんの 実物で 確かめる 時） */
const 見本の道 = path.join(ROOT, 'tests/fixtures/book-open-sample.xlsb');
const 材料 = 外の本 ? 外の本 : 見本の道;
if (外の本 && !fs.existsSync(外の本)) { console.log('★本が 在りません★ ' + 外の本); process.exit(2); }
/* ★見本では B2（行1・列1）を 打ちます★＝★D2 と D7 と E2 に つられます★
     E2 は `=TEXT(D2,"#,##0")&"円"` ＝★答えが 字の 式（記録8）★ */
const 打つ行 = Number(取る('行', 1)), 打つ列 = Number(取る('列', 1));
const 打つ値 = 取る('値', '5');
/* ★★`--足す` ･･･ 元の 値に 足した 数を 打ちます★★（2026-09-25）
     ＝経営者1 の 再現が これ（★元の 値 ＋ 1★）
     ＝★決め打ちの 数を 入れると つられる 所が 変わります★ */
const 足す = 引数.indexOf('--足す') >= 0 ? Number(取る('足す', '1')) : null;

/** ★`.xlsb` の 記録を 数える★（★借り物を 通しません＝自前で 歩きます★） */
async function 記録を数える(道) {
  const z = ZipSurgeon.read(new Uint8Array(fs.readFileSync(道)));
  const 板 = z.names().filter((n) => /^xl[/]worksheets[/]sheet[0-9]+[.]bin$/.test(n));
  const 出 = { 記録8: 0, 記録9: 0, 印: null };
  /* ★記録を 歩くのは `XlsbEdit.parse`★＝★在る 物を 呼びます（作って いません）★ */
  for (const n of 板) {
    const b = await z.bytes(n);
    const r = XE.parse(b instanceof Uint8Array ? b : new Uint8Array(b), XE.SHAPE.sheet);
    if (!r.ok) continue;
    r.recs.forEach((x) => { if (x.id === XE.R.FMLA_STRING) 出.記録8++; if (x.id === XE.R.FMLA_NUM) 出.記録9++; });
  }
  const wb = z.names().filter((n) => /^xl[/]workbook[.]bin$/.test(n))[0];
  if (wb) {
    const b2 = await z.bytes(wb);
    const r2 = XE.parse(b2 instanceof Uint8Array ? b2 : new Uint8Array(b2), XE.SHAPE.workbook);
    if (r2.ok) 出.印 = JT.開いたら全部計算するか(r2.recs);
  }
  return 出;
}

console.log('[直して 保存] ★お客さんが 入力値を 直して 書き出せるか★');
/* ★★要る 物が 無ければ 数を 出さずに 止まります★★（2026-09-25 の 決め） */
{
  const 足りない = [];
  if (typeof XE.parse !== 'function') 足りない.push('XlsbEdit.parse');
  if (typeof JT.開いたら全部計算するか !== 'function') 足りない.push('XlsbJitai.開いたら全部計算するか');
  if (typeof XE.全部計算の印を立てる !== 'function') 足りない.push('XlsbEdit.全部計算の印を立てる');
  if (足りない.length) { console.log('  ★★測れません★★ 在りません ... ' + 足りない.join(' / ')); process.exit(8); }
}
const 元 = await 記録を数える(材料);
console.log('  ★材料★ ' + path.basename(材料)
  + '（★答えが 字の 式（記録8） ' + 元.記録8 + '個★ ／ 答えが 数の 式（記録9） ' + 元.記録9 + '個'
  + ' ／ 印 ' + JSON.stringify(元.印) + '）');
if (元.記録8 < 1) {
  console.log('  ★★答えが 字の 式が 0個＝この 見張りは 空振りです★★');
  process.exit(8);
}

const wk = await borrow('naoshite-hozon', 'webkit');
const browser = await launch('naoshite-hozon', wk, {}, 'webkit');
const page = await browser.newPage({ viewport: { width: 1280, height: 900 }, acceptDownloads: true });
const 配信 = await 立てる(ROOT);
const 落とし先 = fs.mkdtempSync(path.join(os.tmpdir(), 'exally-naoshite-'));

try {
  let 帯 = '';
  let 触れなかった = null;
  page.on('console', (m) => {
    const t = String(m.text());
    const g = /触れなかった 式 ([0-9]+)個/.exec(t);
    if (g) 触れなかった = Number(g[1]);
  });
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

  /* ══ ★入力値を 1つ 直します★ ══（★画面の `setCell`＝お客さんの 道★） */
  /* ★打つ 所は 札で 変えられます★（実物では 空きマスを 避ける 為） */
  const 打った = await page.evaluate(([r, c, v, 足], ) => {
    const sh = window.sheets[window.activeSheet];
    const 元 = sh && sh.data ? sh.data[r + ',' + c] : null;
    let 字 = v;
    if (足 !== null) {
      /* ★元の 値に 足します★＝★決め打ちの 数を 入れない★ */
      const 生 = 元 ? (元.v !== undefined && 元.v !== null && 元.v !== '' ? 元.v : 元.d) : null;
      const n = Number(String(生).replace(/[^0-9.eE+-]/g, ''));
      if (!isFinite(n)) return { だめ: '元の 値が 数では ありません', 元: String(生).slice(0, 12) };
      字 = String(n + 足);
    }
    window.setCell(r, c, 字);
    return { 打った字: 字 };
  }, [打つ行, 打つ列, 打つ値, 足す]);
  if (打った.だめ) { console.log('  ★★打てません★★ ' + 打った.だめ); process.exit(8); }
  /* ★★打った 字は 出しません★★＝★司さんの 実物の 中身です★（決め ②）
       ⇒★桁数だけ 出します★ */
  console.log('      ── 実測 ── 打った 字 ... ★' + String(打った.打った字).length + '桁★'
    + '（★中身は 出しません＝司さんの 実物です★）');
  await page.waitForTimeout(1200);

  const 落とし物 = page.waitForEvent('download', { timeout: 120000 }).catch(() => null);
  await page.evaluate(() => { window.saveOpenedBook(); });
  await page.waitForTimeout(1200);
  const 窓 = await page.evaluate(() => {
    const go = document.getElementById('diffGo');
    return go && go.offsetWidth > 0 ? String(go.textContent || '').trim() : '';
  });
  console.log('      ── 実測 ── 窓の ボタン ... ' + JSON.stringify(窓));
  if (!窓) { console.log('  ★★窓が 出て いません＝この 見張りは 空振りです★★'); process.exit(8); }
  await page.click('#diffGo');

  const dl = await 落とし物;
  帯 = await page.evaluate(() => {
    const t = document.querySelector('.toast, #toast, [class*="toast"]');
    return t ? String(t.textContent || '').trim().slice(0, 200) : '';
  });
  console.log('      ── 実測 ── 落とし物 ... ' + (dl ? '★来た★' : '★来ません★')
    + ' ／ 帯 ' + JSON.stringify(帯.slice(0, 120)));

  T('★★入力値を 直しても 書き出せる★★（★前は 記録 8 で 1冊 まるごと 出なかった★）',
    !!dl, '帯 ' + JSON.stringify(帯.slice(0, 160)));
  if (!dl) throw new Error('落とし物が 来ません');

  const 出た = path.join(落とし先, 'out.xlsb');
  await dl.saveAs(出た);
  const 後 = await 記録を数える(出た);
  console.log('      ── 実測 ── 出た 本 ... ' + fs.statSync(出た).size.toLocaleString() + ' バイト'
    + ' ／ 記録8 ' + 後.記録8 + '個 ／ 記録9 ' + 後.記録9 + '個 ／ ★印 ' + JSON.stringify(後.印) + '★');

  /* ★打った 値が 入って いるか★（★借り物で 開き直して 読みます★） */
  const wb2 = XLSX.read(fs.readFileSync(出た), { type: 'buffer' });
  const 番地 = XLSX.utils.encode_cell({ r: 打つ行, c: 打つ列 });
  const ws2 = wb2.Sheets[wb2.SheetNames[0]];
  const 打ったマス = ws2 ? ws2[番地] : null;
  console.log('      ── 実測 ── 出た 本の ' + 番地 + ' ... ★' + String(打ったマス && 打ったマス.v).length + '桁★'
    + '（打った 字と 同じか ... ★' + (String(打ったマス && 打ったマス.v) === String(Number(打った.打った字)) ? 'はい' : 'いいえ') + '★）');
  T('★★打った 値が 本の 中に 入って いる★★（' + 番地 + '）',
    !!打ったマス && String(打ったマス.v) === String(Number(打った.打った字)),
    番地 + ' ... ★桁数 ' + String(打ったマス && 打ったマス.v).length + '★（中身は 出しません）');

  /* ══ ★★印は 「触れなかった 式が 在る 時だけ」立てます★★ ══（2026-09-25）
       ＝★1つも 触れなかった 物が 無ければ 立てません★（★要らない 印を 立てない★）
       ＝画面が 声に 出す 数（`触れなかった 式 N個`）で 分けます
       ⇒★声が 来なければ 数を 出さずに 止まります★ */
  console.log('      ── 実測 ── ★触れなかった 式★ ... ' + JSON.stringify(触れなかった) + '個');
  if (触れなかった === null) {
    console.log('  ★★画面の 声が 来ません＝この 検査は 空振りです★★');
    process.exit(8);
  }
  /* ══ ★★触れなかった 所は お客さんに 言うか★★ ══（2026-09-25・経営者1 の ④）
       ＝★黙って いると お客さんは 「全部 直った」と 思います★
       ＝記憶「ボタンの『出せます』と 押した後の 門は 同じ物を 見る」の 家 */
  if (触れなかった > 0) {
    T('★★触れなかった 所を 帯で 言う★★（★黙って 「全部 直った」に しない★）',
      帯.indexOf(String(触れなかった) + 'か所は 形が 違う') >= 0,
      '帯 ' + JSON.stringify(帯.slice(0, 200)));
  }
  if (触れなかった > 0) {
    T('★★触れなかった 式が 在る 時は 「開いたら 全部 計算しろ」の 印を 立てる★★',
      後.印 === true, '印 ' + JSON.stringify(後.印) + '（元は ' + JSON.stringify(元.印) + '／触れなかった ' + 触れなかった + '個）');
  } else {
    T('★★触れなかった 式が 0個なら 印は 立てない★★（★要らない 印を 立てない★）',
      後.印 === 元.印, '印 ' + JSON.stringify(後.印) + '（元は ' + JSON.stringify(元.印) + '）');
  }
} finally {
  await page.close().catch(() => {});
  await browser.close().catch(() => {});
  配信.閉じる();
}

console.log('');
console.log('naoshite-hozon: ' + pass + ' 緑 / ' + fail + ' 赤');
process.exit(fail ? 1 : 0);
