/* hakaru-kakikata-no-kazu.mjs — ★書式が どれだけ 付いて いるか／隠れなかった 0 は 何か★ 2026-09-25
 *
 *  ★★なぜ 要るか（経営者1 の 実測・物差しは 実Excel）★★
 *    表示が 違う ★2,420個★（9,163 から 減った 後）の 内訳:
 *      ★㋐桁の 区切りが 無い ... 1,645個★
 *      ★㋒日付が 生の 数字 .....   459個★
 *      ★㋑小数の 桁が 違う .....   224個★
 *        ㋓その他 ..............    45個
 *    ＋★中の 字が 0 なのに 隠れて いない 9個★
 *
 *  ★★この 道具が 数える 物★★（★直す 前に 数える★）
 *    ⑴★マスに 書式（`numFmt`）が 付いて いるか★
 *        ＝付いて いなければ `General` 扱い＝★カンマも 丸めも 効きません★
 *        ＝`.xlsb` は ★借り物が 書式を くれません★（字体と 同じ 形）
 *    ⑵★書式の 字ごとの 数★（★中身は 出しません★＝書式の 字だけ）
 *    ⑶★隠れなかった 0 の `d` と `d字`★（★経営者1 の 頼み★）
 *
 *  ★お客さんの 道で 数えます★＝`#bookFileInput` に 渡す／★1バイトも 書きません★
 *  ★マスの 値は 出しません★（★書式の 字・型・数だけ★）
 *
 *  走らせ方: node docs/measured/hakaru-kakikata-no-kazu.mjs [材料]
 */
import path from 'node:path'; import http from 'node:http'; import fs from 'node:fs';
import { fileURLToPath } from 'node:url';
import { borrow, launch } from '../../scripts/_borrow-playwright.mjs';
import { 道を確かめる, 効いた数を出す } from './_gamen-no-michi.mjs';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '../..');
const 材料 = process.argv[2] || path.join(ROOT, 'tests/fixtures/cross-sheet-sample.xlsb');

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
console.log('★書式が どれだけ 付いて いるか★');
console.log('  材料 ... ' + path.basename(材料) + '（' + fs.statSync(材料).size.toLocaleString() + ' バイト）');

const wk = await borrow('kakikata', 'webkit');
const browser = await launch('kakikata', wk, {}, 'webkit');
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
    const el = document.getElementById('toast');
    const t = el ? String(el.textContent || '') : '';
    if (t.indexOf('読めませんでした') >= 0 || t.indexOf('開けませんでした') >= 0) throw new Error(t.slice(0, 120));
    return false;
  }, null, { timeout: 900000 });
  await page.waitForTimeout(600);

  /* ★★決め（2026-09-25・2人で 決めました）★★
       ★要る 関数が 1つでも 無ければ 数を 出さずに 止まります★ */
  await 道を確かめる(page);

  const 板の数 = await page.evaluate(() => (window.sheets || []).length);
  const 合 = { 式: 0, 書式あり: 0, 書式なし: 0, 隠さなかった0: 0, 式にREF: 0 };
  const REF本 = [];
  const 書式ごと = {};
  const 残り0 = [];
  const 描 = { 見た: 0, 生にカンマ: 0, 描きにカンマ: 0, 生と描きが違う: 0, 見本: [], 画面が0: 0, 零本: [] };

  for (let i = 0; i < 板の数; i++) {
    /* ★板を 開く★＝`activeSheet` の 代入だけでは 台に 流れません */
    await page.evaluate((n) => { window.switchSheet(n); }, i);
    await page.waitForTimeout(60);
    const 出 = await page.evaluate(() => {
      const sh = window.sheets[window.activeSheet];
      const d = sh.data || {};
      const 数 = { 式: 0, 書式あり: 0, 書式なし: 0, 隠さなかった0: 0, 式にREF: 0 };
      const 書 = {};
      const 残 = [];
      const 残REF = [];
      /* ★★「画面に 出る 字」は 描く 時に 作られます★★（2026-09-25）
           `_字の元` は ★書式を 掛ける 前の 値★ です。
           画面は `fmtForDisplay(raw, cell.numFmt, _入る字数(w, raw, cell.numFmt))` を 描きます。
           ⇒★どちらを 取るかで 「カンマが 無い」に なるか 変わります★ */
      const 描 = { 見た: 0, 生にカンマ: 0, 描きにカンマ: 0, 生と描きが違う: 0, 見本: [], 画面が0: 0, 零本: [] };
      for (const k in d) {
        const cell = d[k];
        if (!cell || cell.merged) continue;
        if (!cell.f) continue;                       /* ★式の マスだけ★（経営者1 と 同じ 分母） */
        数.式++;
        const f = cell.numFmt;
        if (f && f !== 'General') { 数.書式あり++; 書[f] = (書[f] || 0) + 1; }
        else { 数.書式なし++; 書[f ? 'General' : '(無し)'] = (書[f ? 'General' : '(無し)'] || 0) + 1; }
        /* ★画面に 出る 字を 作って 数える★ */
        try {
          const p = k.split(',');
          const w = window.cW(+p[1]);
          const r0 = window._字の元(cell);
          const 生 = String(r0 === undefined || r0 === null ? '' : r0);
          const 描き = window._答えは字か(cell) ? 生
            : String(window.fmtForDisplay(r0, cell.numFmt, window._入る字数(w, r0, cell.numFmt)));
          描.見た++;
          if (生.indexOf(',') >= 0) 描.生にカンマ++;
          if (描き.indexOf(',') >= 0) 描.描きにカンマ++;
          if (生 !== 描き) {
            描.生と描きが違う++;
            if (描.見本.length < 4) 描.見本.push({
              生: 生.replace(/[0-9]/g, '9'), 描き: 描き.replace(/[0-9]/g, '9'), 書式: cell.numFmt || '(無し)',
            });
          }
          /* ★★画面が 「0」に 見える のに 隠れて いない マス★★（2026-09-25）
               経営者1 の 実測 ... 実Excel は 空・うちは `0 円` `0.00時間` `¥0` ＝★47個★
               ⇒`Number(raw)===0` では 拾えて いない＝★raw が ぴったり 0 では ない★ かも
               ⇒★raw の 形と 書式と 板を 数えます★（★中身は 9 に 伏せます★） */
          const 数だけ = 描き.replace(/[^0-9.\-]/g, '');
          if (描き !== '' && 数だけ !== '' && Number(数だけ) === 0
              && !window._ゼロを隠すか(cell, r0)) {
            描.画面が0++;
            if (描.零本.length < 10) 描.零本.push({
              板: sh.name, 隠す板: !!sh.ゼロを隠す,
              raw形: 生.replace(/[0-9]/g, '9'), raw長: 生.length,
              画面: 描き.replace(/[0-9]/g, '9'),
              書式: cell.numFmt || '(無し)',
              ぴったり0: Number(生) === 0,
            });
          }
        } catch (e) { /* 1マス 読めなくても 続ける */ }
        /* ★★式の 中に `#REF!` が 在る マス★★（2026-09-25）
             経営者1 の 実測 ... ★実Excel が `#REF!`・うちは 数 ＝ 68個★
             ⇒★数が 出る＝お客さんは 間違いに 気付けません★＝一番 重い
             ⇒★式の 字に `#REF!` が 在るか★ を 先に 数えます */
        if (String(cell.f).indexOf('#REF!') >= 0) {
          数.式にREF++;
          if (残REF.length < 8) 残REF.push({
            板: sh.name, 印: k,
            式形: String(cell.f).replace(/[0-9]/g, '9').slice(0, 60),
            画面: String(window._答えは字か(cell) ? window._字の元(cell)
              : window.fmtForDisplay(window._字の元(cell), cell.numFmt, 30)).replace(/[0-9]/g, '9').slice(0, 20),
          });
        }
        /* ★中の 字が 0 なのに 隠れなかった マス★ */
        const raw = window._字の元(cell);
        const s = String(raw === undefined || raw === null ? '' : raw).trim();
        if (s !== '' && Number(s) === 0 && isFinite(Number(s))) {
          if (!window._ゼロを隠すか(cell, raw)) {
            数.隠さなかった0++;
            if (残.length < 12) {
              残.push({
                板: sh.name, 印: k,
                d: s.replace(/[0-9]/g, '9'),        /* ★形だけ★ */
                d字: !!cell.d字,
                書式: cell.numFmt || '(無し)',
                隠す板: !!sh.ゼロを隠す,
                vの型: typeof cell.v, vが空: cell.v === '',
              });
            }
          }
        }
      }
      return { 数, 書, 残, 残REF, 描き: 描 };
    });
    合.式 += 出.数.式; 合.書式あり += 出.数.書式あり;
    合.書式なし += 出.数.書式なし; 合.隠さなかった0 += 出.数.隠さなかった0;
    for (const k of Object.keys(出.書)) 書式ごと[k] = (書式ごと[k] || 0) + 出.書[k];
    残り0.push(...出.残);
    合.式にREF += 出.数.式にREF;
    if (REF本.length < 8) REF本.push(...出.残REF.slice(0, 3));
    描.見た += 出.描き.見た; 描.生にカンマ += 出.描き.生にカンマ;
    描.描きにカンマ += 出.描き.描きにカンマ; 描.生と描きが違う += 出.描き.生と描きが違う;
    if (描.見本.length < 6) 描.見本.push(...出.描き.見本.slice(0, 2));
    描.画面が0 += 出.描き.画面が0;
    if (描.零本.length < 10) 描.零本.push(...出.描き.零本.slice(0, 3));
  }

  console.log('');
  効いた数を出す('書式が 付いて いる', 合.書式あり, 合.式);
  効いた数を出す('画面が 0 に 見えるのに 隠れて いない', 描.画面が0, 描.見た);
  効いた数を出す('式の 字に `#REF!` が 在る', 合.式にREF, 合.式);
  console.log('★式の マス ' + 合.式.toLocaleString() + '個★（板 ' + 板の数 + '枚）');
  console.log('  ★書式が 付いて いる ... ' + 合.書式あり.toLocaleString() + '個★');
  console.log('  ★書式が 無い（General 扱い）... ' + 合.書式なし.toLocaleString() + '個★');
  console.log('    ⇒★書式が 無ければ カンマも 丸めも 効きません★');
  console.log('');
  console.log('★書式の 字ごと（多い順・上位 20）★（★中身は 出しません★）');
  const 並 = Object.keys(書式ごと).map((k) => ({ 書: k, 数: 書式ごと[k] }))
    .sort((a, b) => b.数 - a.数).slice(0, 20);
  for (const x of 並) console.log('    ' + String(x.数).padStart(6) + '  ' + x.書);
  console.log('');
  console.log('★中の 字が 0 なのに 隠れなかった ... ' + 合.隠さなかった0 + '個★');
  for (const x of 残り0) {
    console.log('    ' + String(x.板).slice(0, 8).padEnd(9) + ' ' + String(x.印).padEnd(10)
      + ' d形 ' + String(x.d).padEnd(12) + ' d字 ' + String(x.d字).padEnd(6)
      + ' 隠す板 ' + String(x.隠す板).padEnd(6) + ' vの型 ' + x.vの型 + '/空' + x.vが空
      + ' 書式 ' + x.書式);
  }
  console.log('');
  console.log('★★式の 字に `#REF!` が 在る ... ' + 合.式にREF + '個★★');
  for (const x of REF本) console.log('    ' + String(x.板).slice(0,8).padEnd(9) + ' ' + String(x.印).padEnd(10)
    + ' 画面 ' + String(x.画面).padEnd(14) + ' 式形 ' + x.式形);
  console.log('');
  console.log('★★画面が 0 に 見える のに 隠れて いない ... ' + 描.画面が0 + '個★★');
  for (const x of 描.零本) console.log('    ' + String(x.板).slice(0,8).padEnd(9)
    + ' 隠す板 ' + String(x.隠す板).padEnd(6) + ' raw形 ' + String(x.raw形).padEnd(20)
    + ' 長さ ' + String(x.raw長).padStart(3) + ' ぴったり0 ' + String(x.ぴったり0).padEnd(6)
    + ' 画面 ' + String(x.画面).padEnd(12) + ' 書式 ' + x.書式);
  console.log('');
  console.log('★★『書式を 掛ける 前』と『画面に 出る 字』★★');
  console.log('    見た 式 .............. ' + 描.見た.toLocaleString());
  console.log('    ★生（`_字の元`）に カンマ ... ' + 描.生にカンマ.toLocaleString() + '個★');
  console.log('    ★画面に 出る 字に カンマ ... ' + 描.描きにカンマ.toLocaleString() + '個★');
  console.log('    生と 画面が 違う ....... ' + 描.生と描きが違う.toLocaleString() + '個');
  for (const x of 描.見本) console.log('      生 ' + String(x.生).padEnd(16)
    + ' ⇒ 画面 ' + String(x.描き).padEnd(18) + ' 書式 ' + x.書式);

  /* ══ ★★その 書式で 何と 出るか★★ ══（2026-09-25）
       書式は ★15,799個中 15,245個に 付いて います★。
       ⇒それでも カンマが 出て いない ⇒★書式を 読む 所の 話★ と 見て 数えます。
       ★材料の 値では なく 決めた 数を 通します★＝★実物の 中身は 出しません★ */
  const 試し = await page.evaluate((書たち) => {
    const 出 = [];
    for (const f of 書たち) {
      let a = '', b = '', c = '';
      try { a = String(window.fmtForDisplay(1234567, f, 30)); } catch (e) { a = '★断り★ ' + e.message.slice(0, 40); }
      try { b = String(window.fmtForDisplay(1234.5678, f, 30)); } catch (e) { b = '★断り★'; }
      try { c = String(window.fmtForDisplay(-1234, f, 30)); } catch (e) { c = '★断り★'; }
      出.push({ 書: f, 千二百三十四万: a, 小数: b, 負: c });
    }
    return 出;
  }, 並.map((x) => x.書));
  console.log('');
  console.log('★その 書式に 1234567 ／ 1234.5678 ／ -1234 を 通すと★');
  for (const x of 試し) {
    console.log('    ' + String(x.書).slice(0, 30).padEnd(31)
      + ' ' + String(x.千二百三十四万).slice(0, 18).padEnd(19)
      + ' ' + String(x.小数).slice(0, 18).padEnd(19)
      + ' ' + String(x.負).slice(0, 16));
  }
} finally {
  await page.close().catch(() => {});
  await browser.close().catch(() => {});
  配信.閉じる();
}
