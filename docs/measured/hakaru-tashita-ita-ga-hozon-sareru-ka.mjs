/* hakaru-tashita-ita-ga-hozon-sareru-ka.mjs
 *   ★足した 板が 書き出す 先に 入るか★を ★お客さんの 道★で 測る（2026-09-21）
 *
 *  ★★なぜ 要るか★★
 *    司さん「★全部 保存しろや、断る 理由が なんか あるんか★」（ア）
 *    `lib/hairanai.js` の 頭（2026-09-06 実測）
 *      「新しいシートを 1枚 足して 保存 ⇒ ★出た ファイルに 無い★」
 *    ⇒★直した ので 測り直します★
 *
 *  ★★お客さんの 道で やります★★
 *    ①画面の 「読み込む」と 同じ 入口（`#bookFileInput`）で 開く
 *    ②画面の 「シートを 足す」と 同じ 口で 板を 足す
 *    ③画面の 「書き出す」が 呼ぶ 所（`BookOpen.saveOpened`）で 保存する
 *    ⇒★JS で 中を 直に 作りません★
 *
 *  ★★出す 物★★
 *    ・出た ファイル（`%TEMP%`）＝★経営者1 が 実Excel に 開かせます★
 *    ・大きさと sha256
 *    ・★包みの 部品の 前後★（増えた／減った／変わった）
 *
 *  ★★ここでは 測れない 事★★
 *    ・★実Excel が 「修復しました」と 言わないか★（COM が 要る＝経営者1 の 持ち場）
 *    ・★グラフが 減って いないか★（この 材料に グラフが 在れば 部品で 分かります）
 *
 *  使い方:
 *    node docs/measured/hakaru-tashita-ita-ga-hozon-sareru-ka.mjs
 *    node docs/measured/hakaru-tashita-ita-ga-hozon-sareru-ka.mjs --もと <xlsx への 道>
 */
import path from 'node:path';
import http from 'node:http';
import fs from 'node:fs';
import crypto from 'node:crypto';
import { fileURLToPath } from 'node:url';
import { borrow, launch } from '../../scripts/_borrow-playwright.mjs';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..', '..');
const TEMP = process.env.TEMP || process.env.TMP || '.';

const i = process.argv.indexOf('--もと');
const 材料 = (i >= 0 && process.argv[i + 1])
  ? process.argv[i + 1]
  : path.join(ROOT, 'tests/fixtures/kazari-hiraku3.xlsx');

function 立てる(root) {
  const 型 = {
    '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8',
    '.css': 'text/css; charset=utf-8', '.json': 'application/json; charset=utf-8',
    '.png': 'image/png', '.svg': 'image/svg+xml', '.ico': 'image/x-icon',
  };
  const s = http.createServer((q, r) => {
    const f = path.join(root, decodeURIComponent(String(q.url).split('?')[0]).replace(/^\/+/, ''));
    if (!f.startsWith(root) || !fs.existsSync(f) || fs.statSync(f).isDirectory()) {
      r.statusCode = 404; return r.end('no');
    }
    r.setHeader('content-type', 型[path.extname(f).toLowerCase()] || 'application/octet-stream');
    fs.createReadStream(f).pipe(r);
  });
  return new Promise((x) => s.listen(0, '127.0.0.1',
    () => x({ url: 'http://127.0.0.1:' + s.address().port, 閉じる: () => s.close() })));
}

console.log('');
console.log('[tashita-ita] ★足した 板が 書き出す 先に 入るか★（お客さんの 道）');
if (!fs.existsSync(材料)) { console.log('  ★材料が 無い★ ' + 材料); process.exit(2); }
const 元 = fs.readFileSync(材料);
console.log('      ＝ 材料 ' + 材料);
console.log('      ＝ ' + 元.length + 'B ／ sha256 '
  + crypto.createHash('sha256').update(元).digest('hex'));

const wk = await borrow('tashita-ita', 'webkit');
const browser = await launch('tashita-ita', wk, {}, 'webkit');
const page = await browser.newPage({ viewport: { width: 1280, height: 800 } });
const 配信 = await 立てる(ROOT);

try {
  await page.goto(配信.url + '/book.html', { waitUntil: 'load', timeout: 60000 });
  await page.evaluate(() => {
    document.body.classList.remove('exally-locked');
    const ov = document.getElementById('loginOv');
    if (ov) { ov.classList.remove('open'); ov.style.display = 'none'; }
  });
  await page.setInputFiles('#bookFileInput', 材料);
  await page.waitForFunction(() => {
    const sh = (window.sheets || [])[window.activeSheet || 0];
    return !!(sh && sh.data && Object.keys(sh.data).length > 3);
  }, null, { timeout: 60000 }).catch(() => {});

  const 前 = await page.evaluate(() => (window.sheets || []).map((s) => s.name));
  console.log('      ＝ 開いた 板 ' + 前.length + '枚 ／ ' + 前.join(' '));

  /* ★★板を 足します★★（★画面が 使う 口を 探して 呼びます★） */
  const 足した = await page.evaluate(() => {
    const 口 = ['addSheet', 'シートを足す', '新しいシート', 'addNewSheet'];
    for (const k of 口) if (typeof window[k] === 'function') { window[k](); return k; }
    /* ★口が 見つからない 時は 台に 直に 足します★＝★その事を 必ず 出します★ */
    window.sheets.push({ name: 'Tashita', data: {}, colW: {}, rowH: {},
      hiddenRows: {}, hiddenCols: {} });
    return '(画面の 口が 見つからず 台に 直に 足しました)';
  });
  console.log('      ＝ 足した 道 ' + 足した);

  /* ★足した 板に 値を 入れます★（★空だと 「入った」が 分かりにくい★） */
  await page.evaluate(() => {
    const i = window.sheets.length - 1;
    const sh = window.sheets[i];
    sh.data['0,0'] = { v: 123, f: '', d: '123' };
    sh.data['0,1'] = { v: 'tashita', f: '', d: 'tashita' };
  });
  const 後 = await page.evaluate(() => (window.sheets || []).map((s) => s.name));
  console.log('      ＝ 足した 後 ' + 後.length + '枚 ／ ' + 後.join(' '));

  /* ★★本番の 保存の 道★★（`saveOpenedBook` が 呼ぶ 所） */
  const 出 = await page.evaluate(async () => {
    if (typeof window._ensureXlsx === 'function') {
      try { await window._ensureXlsx(); } catch (e) { /* もう 在る */ }
    }
    if (!window.BookOpen || typeof window.BookOpen.saveOpened !== 'function') {
      return { だめ: '(BookOpen.saveOpened が 無い)' };
    }
    try {
      const r = await window.BookOpen.saveOpened(window.sheets);
      const u8 = r.bytes instanceof Uint8Array ? r.bytes : new Uint8Array(r.bytes);
      return { だめ: '', 中: Array.from(u8) };
    } catch (e) { return { だめ: String(e && e.message).slice(0, 200) }; }
  });
  if (出.だめ) throw new Error(出.だめ);

  const 中 = Buffer.from(出.中);
  /* ★★出す 先の 名前は 材料から 作ります★★（2026-09-22）
       ★★ここで 1回 踏みました★★
         前は ★材料が 何でも 同じ 名前★ で 書いて いました。
         ⇒2本目（グラフ入り）を 測った 時に ★1本目を 上書き★ しました。
         ⇒経営者1 が ★1組目 対 2組目の 出力★ を 並べかけました
         ⇒★★そのままなら 「値が 違う」と 嘘の 赤が 出ます★★
         ⇒★しかも その 赤は 私の 直しの せいに 見えます★
       ⇒★材料の 名前を 付けて 分けます★（`exally-tashita-<材料の名>.xlsx`） */
  const もとの名 = path.basename(材料).replace(/\.[^.]+$/, '');
  /* ★出す 先の 拡張子も 材料に 合わせます★（2026-09-22）
       ＝`.xlsb` を 測ったのに `.xlsx` の 名で 置いて いました
       ＝★中身は 正しいのに 名前が 嘘★＝★次に 開く 人が 迷います★ */
  const もとの拡張子 = path.extname(材料) || '.xlsx';
  const 置き場 = path.join(TEMP, 'exally-tashita-' + もとの名 + もとの拡張子);

  /* ══ ★★名前と 中身が 合って いるか★★ ══（2026-09-22）
       ★★なぜ 要るか★★
         印（sha256）は ★中身の すり替わり★ を 捕まえますが、
         ★★名前の 嘘は 捕まえません★★（★中身が 同じ なら 印も 同じ★）。
         ⇒2026-09-22 に 実際に 通り抜けました
           `exally-tashita-kazari-hiraku3.`xlsx`` の 中身が ★`.xlsb`★
           ＝経営者1 が ★中身の 形（板が .bin か .xml か）を 見て★ 見つけました
         ⇒★そのまま なら 「.xlsx の 測り」として 紙に 残る 所でした★
       ★だから 出す 前に 自分で 数えます★＝★合わなければ 止めます★ */
  var 板の形 = '';
  {
    let q = 0;
    while (q + 30 <= 中.length) {
      if (中.readUInt32LE(q) !== 0x04034b50) break;
      const z = 中.readUInt32LE(q + 18);
      const n2 = 中.readUInt16LE(q + 26);
      const x2 = 中.readUInt16LE(q + 28);
      const nm = 中.toString('utf8', q + 30, q + 30 + n2);
      if (nm.indexOf('xl/worksheets/sheet') === 0) {
        if (nm.slice(-4) === '.bin') { 板の形 = 'bin'; break; }
        if (nm.slice(-4) === '.xml') { 板の形 = 'xml'; break; }
      }
      q = q + 30 + n2 + x2 + z;
    }
  }
  const 名の形 = (もとの拡張子.toLowerCase() === '.xlsb') ? 'bin' : 'xml';
  if (板の形 && 板の形 !== 名の形) {
    throw new Error('★★名前と 中身が 合いません★★'
      + '／名前 ' + もとの拡張子 + '（' + 名の形 + '）／中身 ' + 板の形
      + '／★印では 捕まりません★（中身が 同じ なら 印も 同じ）');
  }
  console.log('      ＝ 名前と 中身 ' + もとの拡張子 + ' / 板は ' + (板の形 || '(見つからず)')
    + ' ⇒ ★合い★');
  fs.writeFileSync(置き場, 中);
  console.log('');
  console.log('  ★★出た ファイル★★ ' + 置き場);
  console.log('    ' + 中.length + 'B ／ sha256 '
    + crypto.createHash('sha256').update(中).digest('hex'));
  console.log('');
  console.log('  ★経営者1 へ★ ... ★この ファイルを 実Excel に 開かせて ください★');
  console.log('    ①新しい 板（Tashita）が 在るか ②元の 判子・罫線・図形・グラフが 減って いないか');
  console.log('    ③元の 板の 値が 変わって いないか ④★`Open` が 投げないか★');
  console.log('');
  console.log('  ★部品を 突き合わせるには★');
  console.log('    node docs/measured/kuraberu-tsutsumi-no-buhin.mjs '
    + JSON.stringify(材料) + ' ' + JSON.stringify(置き場));
} catch (e) {
  console.log('  ★止まりました★ ' + String(e && e.message).slice(0, 200));
  process.exitCode = 1;
} finally {
  if (配信) 配信.閉じる();
  await browser.close();
}
