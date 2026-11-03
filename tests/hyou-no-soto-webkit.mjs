/* hyou-no-soto-webkit.mjs — ★本物の ブラウザで 実際に 開いて 押す★ 2026-09-04
 *
 *  ★なぜ 要るか（司さん「やることやれや」）★
 *    jsdom と 字合わせでは ★画面に 本当に 出るか★は 分からない。
 *    ★実Excel に 作らせた 見本の ブック★を 本物の ブラウザで 開いて
 *    ①知らせが 出るか ②ボタンが 出るか ③窓に ★直し方★が 出るか を 見る。
 *
 *  ★見本★ tests/fixtures/hyou-no-soto-sample.xlsx（tools/make-mihon.ps1 が 実Excel で 作った）
 *    表2つ（R8.8＝A1:C5 ／ R8.9＝A8:C12）／R8.9 の C列が R8.8 の その行を 指す
 *    ★Excel 自身も 答えを 出せず 値は 0★（エラーは 出ない＝客は 気づけない）
 *
 *  ★測れない時は 赤にせず「未測定」と はっきり言う★（0件と 混ぜない）
 *
 *  走らせ方: node tests/hyou-no-soto-webkit.mjs [--self-test]
 */
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import http from 'node:http';
import fs from 'node:fs';
import { borrow, launch, unmeasured } from '../scripts/_borrow-playwright.mjs';

/** ★その場で 立てる 小さい 配信★（読むだけ・127.0.0.1 だけ） */
function 立てる(root) {
  const 型 = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8',
    '.mjs': 'text/javascript; charset=utf-8', '.css': 'text/css; charset=utf-8',
    '.json': 'application/json; charset=utf-8', '.png': 'image/png', '.svg': 'image/svg+xml' };
  const s = http.createServer((req, res) => {
    const 道 = decodeURIComponent(String(req.url).split('?')[0]);
    const f = path.join(root, 道.replace(/^\/+/, ''));
    if (!f.startsWith(root) || !fs.existsSync(f) || fs.statSync(f).isDirectory()) { res.statusCode = 404; return res.end('no'); }
    res.setHeader('content-type', 型[path.extname(f).toLowerCase()] || 'application/octet-stream');
    fs.createReadStream(f).pipe(res);
  });
  return new Promise((r) => s.listen(0, '127.0.0.1', () => {
    r({ url: 'http://127.0.0.1:' + s.address().port, 閉じる: () => s.close() });
  }));
}

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const TAG = 'hyou-no-soto';
const 見本 = path.join(ROOT, 'tests/fixtures/hyou-no-soto-sample.xlsx');

let pass = 0, fail = 0;
const T = (n, よい, 添え) => {
  if (よい) { pass++; console.log('  ok   ' + n); }
  else { fail++; console.log('  NG   ' + n + (添え ? '\n       ' + 添え : '')); }
};

if (!fs.existsSync(見本)) {
  console.log('  ★未測定★ 見本が 在りません: ' + 見本);
  process.exit(1);
}

const webkit = await borrow(TAG, 'webkit');
const browser = await launch(TAG, webkit, {}, 'webkit');
const page = await browser.newPage({ viewport: { width: 1280, height: 800 } });

console.log('');
console.log('[' + TAG + '] ★本物の ブラウザで 見本を 開いて 押す★');

const 配信 = await 立てる(ROOT);
try {
  await page.goto(配信.url + '/book.html', { waitUntil: 'load' });
  /* ★★何で 出したかを 必ず 書く★★（会社の 決まり）
     ・手元（127.0.0.1 に 立てた 配信）／本番では ない
     ・★鍵を 外した★ … body.exally-locked を 外している＝★ログインは 通っていない★
       （お客さんは ログインして 入る。ここは 中の 画面を 見る為だけ）
     ・★ファイルは 本物の 口に 渡している★（画面の「開く」が 呼ぶ input と 同じ）
     ・★ボタンは 本物の マウスで 押している★（JS で イベントを 投げていない） */
  await page.evaluate(() => {
    document.body.classList.remove('exally-locked');
    const ov = document.getElementById('loginOv');
    if (ov) { ov.classList.remove('open'); ov.style.display = 'none'; }
  });
  await page.setInputFiles('#bookFileInput', 見本);

  /* ★出るまで 待つ★（読み込みは 非同期・描き終わるまで 待つ）
     ★「何か 出た」で 止めない★＝★2本目の ボタンの 字が 出るまで★ 待つ */
  let 札 = '';
  for (let i = 0; i < 200; i++) {
    札 = await page.evaluate(() => {
      const b = document.getElementById('shindanBtn');
      return (b && !b.hidden) ? (b.textContent || '') : '';
    });
    if (札) break;
    await page.waitForTimeout(100);
  }
  T('★開いただけで 知らせの ボタンが 出る（聞かれる前に こちらから 言う）★', !!札, '出ていない');
  T('★ボタンに 何か所かが 出る（4か所）★', /4か所/.test(札), '札=' + JSON.stringify(札));

  const 中 = await page.evaluate(() => {
    const 拾う = (id) => (document.getElementById(id) || {}).textContent || '';
    return { 知らせ: document.body.innerText || '', 札: 拾う('shindanBtn') };
  });
  T('★知らせに 題が 出る★', 中.知らせ.indexOf('ここは ずっと 0 のままです') >= 0);
  T('★知らせに「0 のまま」が 出る（気づけない事を 言う）★', 中.知らせ.indexOf('0 のまま') >= 0);

  /* ★本物の マウスで 押す★ */
  await page.click('#shindanBtn');
  await page.waitForTimeout(300);
  const 窓 = await page.evaluate(() => {
    const ov = document.getElementById('dgOverlay');
    if (!ov || ov.style.display !== 'flex') return null;
    return {
      題: (document.getElementById('dgTitle') || {}).textContent || '',
      本文: (document.getElementById('dgBody') || {}).textContent || '',
      つぎ: (document.getElementById('dgNext') || {}).textContent || '',
      一覧: (document.getElementById('dgList') || {}).innerText || '',
      行数: (document.getElementById('dgList') || {}).children.length,
    };
  });
  T('★押したら 窓が 開く★', !!窓, '開かない');
  if (窓) {
    T('★窓の 題が 2本目★', 窓.題.indexOf('ずっと 0 のまま') >= 0, 窓.題);
    T('★★窓に 直し方が 出る（R8.8[@標本七ｈ] → R8.9[@標本七ｈ]）★★',
      窓.一覧.indexOf('R8.8 を R8.9 に 変える') >= 0,
      窓.一覧.slice(0, 200));
    T('★一覧が 何を 何に 変えるかを 言う★', 窓.一覧.indexOf('に 変える') >= 0, 窓.一覧.slice(0, 120));
    T('★一覧が 列の 名前で どこかを 言う★', 窓.一覧.indexOf('標本七ｈ の 所') >= 0, 窓.一覧.slice(0, 160));
    T('★客に 中の言葉を 出さない（[@ ／ IFERROR）★',
      窓.本文.indexOf('[@') < 0 && 窓.本文.indexOf('IFERROR') < 0
      && 窓.一覧.indexOf('[@') < 0, (窓.本文 + ' || ' + 窓.一覧).slice(0, 200));
    T('★場所が 出る（計算 の C9）★', 窓.一覧.indexOf('C9') >= 0, 窓.一覧.slice(0, 120));
    T('★4か所とも 出る★', 窓.行数 === 4, '行数=' + 窓.行数);
    T('★どこを 押して 何を 変えるかまで 言う（聞かない）★',
      窓.つぎ.indexOf('式の欄') >= 0 && 窓.つぎ.indexOf('表の名前だけ') >= 0, 窓.つぎ);
  }

  /* ★押したら その場所へ 飛ぶ★ */
  await page.evaluate(() => { const l = document.getElementById('dgList'); if (l && l.children[0]) l.children[0].click(); });
  await page.waitForTimeout(300);
  const 行き先 = await page.evaluate(() => {
    const ov = document.getElementById('dgOverlay');
    return { 窓が閉じた: !ov || ov.style.display === 'none',
      場所: (document.getElementById('cell-addr') || {}).value || '',
      式: (document.getElementById('formula-input') || {}).value || '' };
  });
  T('★選ぶと 窓が 閉じる★', 行き先.窓が閉じた);
  T('★選んだ 場所へ 飛ぶ（C9）★', String(行き先.場所).toUpperCase().indexOf('C9') >= 0,
    '場所=' + JSON.stringify(行き先.場所));
  /* ★式の 欄は 前から A1 の 形で 出る★（表の 参照は 全部 A1 に 直してから 計算する）
     ＝★今回の 直しで そう なった のでは ない★（[@数] も B9 に なっている）。
     ★表の 名前は 診断の 窓が 出す★＝どこを どう 直すかは そこで 分かる。 */
  T('★飛んだ 先に エラーの 印が 見える（実Excel も 答えを 出せない 所）★',
    行き先.式.indexOf('#VALUE!') >= 0, '式=' + JSON.stringify(行き先.式).slice(0, 140));

  /* ★★答えが 実Excel と 同じか★★（2026-09-04＝★絵を 開いて 見つけた★）
     直す前 … 実Excel は ★0★ なのに うちは ★#ERROR★（式ごと 読めず IFERROR まで 道連れ）
     ⇒★数字は 全部 緑だった★＝★絵を 開くまで 分からなかった★ */
  const 画面の値 = await page.evaluate(() => {
    const 出 = {};
    for (const a of ['C9', 'C10', 'C11', 'C12']) {
      const r = +a.slice(1) - 1;
      const cell = (window.sheets[window.activeSheet].data || {})[r + ',2'] || {};
      出[a] = (cell.d !== undefined ? cell.d : cell.v);
    }
    return 出;
  });
  T('★★答えが 実Excel と 同じ（4つとも 0／#ERROR に しない）★★',
    ['C9', 'C10', 'C11', 'C12'].every((a) => 画面の値[a] === 0 || 画面の値[a] === '0'),
    JSON.stringify(画面の値));
  /* ★絵を 撮る★＝★数字が 全部 緑でも 絵を 開いて 見るまで OK を 出さない★（会社の 決まり） */
  if (process.env.EXALLY_SHOT) {
    await page.click('#shindanBtn');
    await page.waitForTimeout(300);
    await page.screenshot({ path: process.env.EXALLY_SHOT });
    console.log('       … 絵 ' + process.env.EXALLY_SHOT);
  }
} finally {
  配信.閉じる();
  await browser.close();
}

console.log('');
console.log(TAG + '-webkit: ' + pass + ' 緑 / ' + fail + ' 赤');
process.exit(fail ? 1 : 0);
