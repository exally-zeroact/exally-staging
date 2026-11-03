/* doko-wo-sawatta-bun-webkit.mjs — ★どこを 触ったか＝1文 ＋ ドロップダウンで 詳しく★ 2026-09-22
 *
 *  ★★司さんの 注文（ア）★★
 *    「★全部 保存しろや★・★Excel のように 保存★
 *      （★どこを 触ったか どんな 関数や マクロが 組まれてるかは
 *        簡潔に 文に して ドロップダウンで 詳しく★）」
 *
 *  ★★形は 経営者1 の 注文で 決めました★★（2026-09-21）
 *    「★「どこを 触ったか」は ★私が 実Excel で 割れる 形★に して ください★
 *      ＝「A1 を 3 から 42 に した」の ように ★マスと 前後の 値★が 在れば
 *        ★私が 開いて 突き合わせられます★」
 *
 *  ★★この 試験が 見る 物★★（2つ。★どちらも 欠けたら 赤★）
 *    ① ★言葉（`DiffPreview.文`）★・★残りの 数は `userCount` から 取る★
 *         ★2026-09-22 ここで 1回 踏みました★
 *           `sheets[].rows` は ★先頭 3行だけ★（この 台の 元からの 決め）。
 *           見えて いる 分だけ 数えたので ★4つ 直したのに 「ほか 1か所」★ と 出ました。
 *           ⇒★数は 必ず 数えた 側から 取ります★
 *    ② ★絵（本物の webkit）★・★畳んだ 中身が 本当に 高さ 0 か★
 *         ★2026-09-08 Rakunally の 請求書で 実測★
 *           `<details>` を 畳んでも 中の 部品が 自分で display を 持つと
 *           ★open=false のまま 幅358×高49 で 見えたまま★ でした。
 *           ⇒★「畳んだ」は 印(open 属性)では なく ★高さ★で 数えます★
 *
 *  ★★お客さんの 道で 出します★★
 *    `#bookFileInput` に ファイルを 渡す＝★画面の 「読み込む」が 押す 物と 同じ★。
 *    そのあと ★保存の 窓を 出して [やめる] を 押します★
 *    ＝★1本も 書き出しません★（ファイルは 作られない）。
 *    ただし ★マスに 打つ 所だけは `setCell()` を 呼びます★
 *      ＝★字を 打つ 真似では なく 台の 口です★（正直に 書く）。
 *
 *  走らせ方: node tests/doko-wo-sawatta-bun-webkit.mjs
 */
import path from 'node:path'; import http from 'node:http'; import fs from 'node:fs';
import { createRequire } from 'node:module';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { borrow, launch } from '../scripts/_borrow-playwright.mjs';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(pathToFileURL(path.join(ROOT, 'package.json')));
const DiffPreview = require_(path.join(ROOT, 'lib/diff-preview.js'));

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

/* ══════ ① 言葉 ══════ */
console.log('★① 言葉（`DiffPreview.文`）★');

/* ★直した 数を 決めて 作る★＝`build()` に 通す（★窓と 同じ 台★） */
function 作る(直した数, 波及数) {
  const ch = {}, base = {}, edited = {};
  for (let i = 0; i < 直した数; i++) {
    ch['0,' + i] = i + 1; base['計算|0,' + i] = 9; edited['計算|0,' + i] = { beforeF: null };
  }
  for (let j = 0; j < (波及数 || 0); j++) { ch['5,' + j] = 100 + j; base['計算|5,' + j] = 7; }
  return DiffPreview.build({
    sheets: [{ name: '計算', data: {} }],
    changedCells: () => ch, base, edited, format: (v) => String(v),
  });
}

{
  const p = 作る(1, 1);
  const 出 = DiffPreview.文(p);
  console.log('      ── 実測 ── ' + 出);
  T('1つ 直した＋波及1 ＝ マスと 前後の 値が 出る',
    出 === '計算!A1 を 9 から 1 に しました。つられて 1か所が 変わります', 出);
}
{
  const p = 作る(2, 0);
  const 出 = DiffPreview.文(p);
  console.log('      ── 実測 ── ' + 出);
  T('2つ 直した ＝ 2つとも 出る（「ほか」は 付かない）',
    出 === '計算!A1 を 9 から 1 に しました／計算!B1 を 9 から 2 に しました', 出);
}
{
  /* ★★これが 2026-09-22 に 踏んだ 穴★★
       `rows` は 先頭 3行だけ ⇒ 見えて いる 分を 数えると 「ほか 1か所」に なった */
  const p = 作る(4, 0);
  const 出 = DiffPreview.文(p);
  console.log('      ── 実測 ── ' + 出 + '（rows ' + p.sheets[0].rows.length
    + '行 / userCount ' + p.userCount + '）');
  T('★4つ 直した ＝「ほか 2か所」★（★rows(3行)では なく userCount(4)で 数える★）',
    出 === '計算!A1 を 9 から 1 に しました／計算!B1 を 9 から 2 に しました（ほか 2か所）', 出);
}
{
  const p = 作る(18, 0);
  const 出 = DiffPreview.文(p);
  console.log('      ── 実測 ── ' + 出);
  T('18つ 直した ＝「ほか 16か所」（★数は いくつでも 合う★）',
    出.indexOf('（ほか 16か所）') > 0, 出);
}
{
  /* 前の 値を 控えて いない＝★黙って 0や 空に しない★ */
  const p = DiffPreview.build({
    sheets: [{ name: '計算', data: {} }],
    changedCells: () => ({ '0,0': 42 }), base: {}, edited: { '計算|0,0': { beforeF: null } },
    format: (v) => String(v),
  });
  const 出 = DiffPreview.文(p);
  console.log('      ── 実測 ── ' + 出);
  T('前が 分からない ＝「分かりません」と 出す（黙って 埋めない）',
    出 === '計算!A1 を 分かりません から 42 に しました', 出);
}
{
  const p = DiffPreview.build({
    sheets: [{ name: '計算', data: {} }], changedCells: () => ({}), base: {}, edited: {},
  });
  const 出 = DiffPreview.文(p);
  console.log('      ── 実測 ── ' + 出);
  T('何も 直して いない ＝「変える所はありません」', 出 === '変える所はありません', 出);
}
{
  /* ★人が 直した 所が 1行も 出て いない 時★（波及が 先頭を 埋めた 等）
     ⇒★無い事に しない★＝数だけ 言う */
  const ch = {}, base = {}, edited = {};
  for (let i = 0; i < 9; i++) { ch['0,' + i] = i; base['計算|0,' + i] = 0; }
  ch['0,20'] = 5; base['計算|0,20'] = 1; edited['計算|0,20'] = { beforeF: null };
  const p = DiffPreview.build({
    sheets: [{ name: '計算', data: {} }], changedCells: () => ch, base, edited,
    format: (v) => String(v), maxRows: 0,
  });
  const 出 = DiffPreview.文(p);
  console.log('      ── 実測 ── ' + 出 + '（rows 0行 / userCount ' + p.userCount + '）');
  T('★1行も 出せなかった 時も 数は 言う★（無い事に しない）',
    出.indexOf('あなたが 直した所 1か所') === 0, 出);
}

{
  /* ══ ★★「うち 別のシートが N か所」★★ ══（2026-09-22）
       経営者1 が 実Excel で 割った 中身（`f876f43`）
         `4月!C4` 1 ⇒ 99 で つられて 4か所。★半分の 2か所が 別の 板★。
       ⇒★同じ 板だけ 見て いたら 「2か所」と 言う 所でした★
       ★★板ごとの 数は 先頭3行では なく 全部から 数える★★
         ＝`rows` は 先頭3行だけ ⇒★行から 数えたら また 嘘に なります★ */
  const ch1 = { '0,0': 99, '3,4': 99000 };                 /* 4月 ... 人 1 ／ つられ 1 */
  const ch2 = { '3,1': 186000, '6,1': 527000 };            /* まとめ ... ★全部 つられ★ */
  const p = DiffPreview.build({
    sheets: [{ name: '4月', data: {} }, { name: 'まとめ', data: {} }],
    changedCells: (sh) => (sh.name === '4月' ? ch1 : ch2),
    base: { '4月|0,0': 1, '4月|3,4': 1000, 'まとめ|3,1': 88000, 'まとめ|6,1': 429000 },
    edited: { '4月|0,0': { beforeF: null } }, format: (v) => String(v),
  });
  const 出 = DiffPreview.文(p);
  console.log('      ── 実測 ── ' + 出 + '（つられ ' + p.spreadCount
    + ' / うち 別の板 ' + p.otherSheetCount + '）');
  T('★つられた 3か所の うち 別のシートが 2か所★（実Excel の 割りと 同じ 形）',
    出 === '4月!A1 を 1 から 99 に しました。つられて 3か所が 変わります（うち 別のシートが 2か所）', 出);
}
{
  /* ★別の 板が 0の 時は 言わない★＝★字を 長くする 分の 値打ちが 無い★ */
  const ch = { '0,0': 99, '3,4': 99000 };
  const p = DiffPreview.build({
    sheets: [{ name: '4月', data: {} }],
    changedCells: () => ch, base: { '4月|0,0': 1, '4月|3,4': 1000 },
    edited: { '4月|0,0': { beforeF: null } }, format: (v) => String(v),
  });
  const 出 = DiffPreview.文(p);
  console.log('      ── 実測 ── ' + 出 + '（うち 別の板 ' + p.otherSheetCount + '）');
  T('★別のシートが 0の 時は 付けない★（0か所 と 書かない）',
    出 === '4月!A1 を 1 から 99 に しました。つられて 1か所が 変わります', 出);
}

/* ══════ ② 絵（本物の webkit・お客さんの 道） ══════ */
console.log('\n★② 絵（本物の webkit で 開いて 保存の 窓を 出す）★');

const 材料 = path.join(ROOT, 'tests/fixtures/cross-sheet-sample.xlsb');
if (!fs.existsSync(材料)) { console.log('  ★材料が 有りません★ ' + 材料); process.exit(1); }

const wk = await borrow('doko-wo-sawatta-bun', 'webkit');
const browser = await launch('doko-wo-sawatta-bun', wk, {}, 'webkit');
const page = await browser.newPage({ viewport: { width: 1280, height: 900 } });
const 配信 = await 立てる(ROOT);

try {
  const 返 = await page.goto(配信.url + '/book.html', { waitUntil: 'load', timeout: 60000 })
    .catch((e) => ({ エラー: e.message }));
  if (!返 || 返.エラー) throw new Error('★開けませんでした★ ' + (返 && 返.エラー));

  await page.evaluate(() => {
    document.body.classList.remove('exally-locked');
    const ov = document.getElementById('loginOv');
    if (ov) { ov.classList.remove('open'); ov.style.display = 'none'; }
  });

  /* ★お客さんの 入口★ */
  await page.setInputFiles('#bookFileInput', 材料);
  await page.waitForFunction(() => {
    const sh = window.sheets && window.sheets[window.activeSheet];
    return !!(sh && sh.data && Object.keys(sh.data).length > 3);
  }, null, { timeout: 60000 });

  /* ★1マス 打つ★（★台の 口＝`setCell`★。字を 打つ 真似では ありません） */
  const 打てた = await page.evaluate(() => {
    if (typeof window.setCell !== 'function') return '(setCell が 無い)';
    const i = (window.sheets || []).findIndex((s) => s.name === '4月');
    if (i < 0) return '(4月 が 無い)';
    window.activeSheet = i;
    window.setCell(3, 2, '99');
    return 'ok';
  });
  if (打てた !== 'ok') throw new Error('★打てませんでした★ ' + 打てた);

  /* ★保存の 窓を 出す★（★[やめる] を 押すので 1本も 書き出しません★） */
  await page.evaluate(() => {
    window.FileOut = { deliver: () => Promise.reject(new Error('★書き出しては いけません★')) };
    window.__保存 = window.saveOpenedBook();
  });
  await page.waitForFunction(() => {
    const ov = document.getElementById('diffOverlay');
    return !!ov && ov.style.display === 'flex';
  }, null, { timeout: 60000 });

  /* ★畳んだ ままの 姿を 測る★（★触る 前★） */
  const 畳んだ時 = await page.evaluate(() => {
    const 高 = (el) => (el ? Math.round(el.getBoundingClientRect().height) : -1);
    const bun = document.getElementById('diffBun');
    const d = document.querySelector('#diffBody details.diffKuwashiku');
    const sm = d && d.querySelector('summary');
    const 中 = d ? Array.prototype.filter.call(d.children, (x) => x.tagName !== 'SUMMARY') : [];
    const nk = document.getElementById('diffNakami');
    const みな = [...document.querySelectorAll('#diffBody details.diffKuwashiku')];
    const 中みな = みな.map((x) => Array.prototype.filter
      .call(x.children, (y) => y.tagName !== 'SUMMARY').reduce((t, y) => t + 高(y), 0));
    return {
      文: bun ? bun.textContent : '(diffBun が 無い)',
      文の高さ: 高(bun),
      中身の文: nk ? nk.textContent : '(diffNakami が 無い)',
      中身の文の高さ: 高(nk),
      畳みの数: みな.length,
      見出したち: みな.map((x) => (x.querySelector('summary') || {}).textContent || ''),
      畳みの中みな: 中みな,
      畳みが在る: !!d,
      open: d ? d.open : null,
      summaryの高さ: 高(sm),
      中の数: 中.length,
      中の高さ: 中.reduce((a, x) => a + 高(x), 0),
      本文の字: (document.getElementById('diffBody') || {}).textContent || '',
    };
  });
  console.log('      ── 実測 ── 文「' + 畳んだ時.文 + '」／ 文の高さ ' + 畳んだ時.文の高さ
    + 'px ／ 畳み open=' + 畳んだ時.open + ' summary ' + 畳んだ時.summaryの高さ
    + 'px ／ 中 ' + 畳んだ時.中の数 + '個 合計 ' + 畳んだ時.中の高さ + 'px');

  T('★1文が 画面に 出ている★（字が 在る＋高さが 0でない）',
    畳んだ時.文の高さ > 0 && /を .* から .* に しました|変える所はありません|あなたが 直した所/.test(畳んだ時.文),
    '高さ ' + 畳んだ時.文の高さ + 'px ／ 字「' + 畳んだ時.文 + '」');

  T('★1文に マスと 前後の 値が 入って いる★（実Excel で 突き合わせられる 形）',
    /![A-Z]+\d+ を .+ から .+ に しました/.test(畳んだ時.文), 畳んだ時.文);

  T('詳しくは ドロップダウンに 入って いる（`<details>` が 在る・閉じて いる）',
    畳んだ時.畳みが在る && 畳んだ時.open === false,
    '在る=' + 畳んだ時.畳みが在る + ' open=' + 畳んだ時.open);

  T('★畳みの 見出しは 見えて いる★（押せないと 開けない）',
    畳んだ時.summaryの高さ > 0, 'summary ' + 畳んだ時.summaryの高さ + 'px');

  T('★★畳んだ 中身は 高さ 0★★（★印(open)では なく 絵で 数える★・2026-09-08 の 穴）',
    畳んだ時.中の数 > 0 && 畳んだ時.中の高さ === 0,
    '中 ' + 畳んだ時.中の数 + '個 合計 ' + 畳んだ時.中の高さ + 'px');

  /* ══ ★★何が 組まれて いるか★★══（2026-09-22・司さんの ア の 後半）
       「★どんな 関数や マクロが 組まれてるかは 簡潔に 文に して ドロップダウンで 詳しく★」
       ★★台は 作り直して いません★★＝`lib/hon-no-nakami.js`（経営者1 の 物）を 呼ぶだけ */
  console.log('      ── 実測 ── 中身の文「' + 畳んだ時.中身の文 + '」（高さ '
    + 畳んだ時.中身の文の高さ + 'px）／ 畳み ' + 畳んだ時.畳みの数 + '個 '
    + 畳んだ時.見出したち.join(' / ') + ' ／ 畳んだ 中 ' + 畳んだ時.畳みの中みな.join(',') + 'px');

  T('★何が 組まれて いるかの 1文が 出ている★（板の 枚数・関数・マクロ）',
    畳んだ時.中身の文の高さ > 0 && /板が \d+枚/.test(畳んだ時.中身の文)
      && /関数|式は/.test(畳んだ時.中身の文) && /マクロ/.test(畳んだ時.中身の文),
    '高さ ' + 畳んだ時.中身の文の高さ + 'px ／ 字「' + 畳んだ時.中身の文 + '」');

  T('★ドロップダウンは 3つ★（どこを触ったか ／ 関数 ／ マクロ）',
    畳んだ時.畳みの数 === 3, '出た ' + 畳んだ時.畳みの数 + '個: ' + 畳んだ時.見出したち.join(' / '));

  T('★★3つとも 畳んだ 中身が 高さ 0★★（1つでも 開いていたら 赤）',
    畳んだ時.畳みの中みな.length === 3
      && 畳んだ時.畳みの中みな.every((x) => x === 0),
    畳んだ時.畳みの中みな.join(' / ') + 'px');

  /* ★開けたら 出る★（畳んだ ままで 中身が 死んで いないか） */
  const 開けた時 = await page.evaluate(() => {
    const d = document.querySelector('#diffBody details.diffKuwashiku');
    d.querySelector('summary').click();
    const 中 = Array.prototype.filter.call(d.children, (x) => x.tagName !== 'SUMMARY');
    return {
      open: d.open,
      中の高さ: 中.reduce((a, x) => a + Math.round(x.getBoundingClientRect().height), 0),
      字: d.textContent,
    };
  });
  console.log('      ── 実測 ── 開けたら open=' + 開けた時.open + ' 中 ' + 開けた時.中の高さ + 'px');
  T('★開けたら 1マスずつ 出る★（高さが 0でない・番地が 在る）',
    開けた時.open === true && 開けた時.中の高さ > 0 && /[A-Z]+\d+/.test(開けた時.字),
    'open=' + 開けた時.open + ' 中 ' + 開けた時.中の高さ + 'px');

  /* ★[やめる] を 押す★＝★1本も 書き出さない★ */
  await page.evaluate(() => { document.getElementById('diffCancel').click(); });
  await page.evaluate(() => window.__保存);
  const 窓が閉じた1 = await page.evaluate(() =>
    document.getElementById('diffOverlay').style.display !== 'flex');
  T('[やめる] で 窓が 閉じる（★1本も 書き出さない★）', 窓が閉じた1);

  /* ══ ★★マクロ入りの 本でも 言えるか★★ ══（2026-09-22）
       ★★ここまでの 材料は マクロが 0本★★＝「マクロは ありません」しか 出て いません。
       ⇒「マクロが 在る 時に 本数を 言えるか」は ★別に 測らないと 分かりません★。
       ⇒`tests/fixtures/vba-sample.xlsm` で もう 1回 通します。
       ★同じ 画面を 使い回します★＝お客さんも 続けて 2冊 開きます。 */
  const 材料2 = path.join(ROOT, 'tests/fixtures/vba-sample.xlsm');
  if (!fs.existsSync(材料2)) throw new Error('★材料が 有りません★ ' + 材料2);
  await page.setInputFiles('#bookFileInput', 材料2);
  await page.waitForFunction(() => {
    const c = window.BookOpen && window.BookOpen.current();
    return !!(c && String(c.name || '').indexOf('vba-sample') >= 0);
  }, null, { timeout: 60000 });
  /* ★誰も 使って いない マスに 打つ★（元の 式を 触らない） */
  await page.evaluate(() => { window.setCell(50, 10, '1'); });
  await page.evaluate(() => {
    window.FileOut = { deliver: () => Promise.reject(new Error('★書き出しては いけません★')) };
    window.__保存2 = window.saveOpenedBook();
  });
  await page.waitForFunction(() => {
    const ov = document.getElementById('diffOverlay');
    return !!ov && ov.style.display === 'flex';
  }, null, { timeout: 60000 });
  const マクロ本 = await page.evaluate(() => {
    const 高 = (el) => (el ? Math.round(el.getBoundingClientRect().height) : -1);
    const みな = [...document.querySelectorAll('#diffBody details.diffKuwashiku')];
    const マ = みな.find((x) => /マクロ/.test((x.querySelector('summary') || {}).textContent || ''));
    return {
      中身の文: (document.getElementById('diffNakami') || {}).textContent || '',
      畳みの数: みな.length,
      マクロの見出し: マ ? (マ.querySelector('summary') || {}).textContent : '(無い)',
      マクロの行: マ ? マ.querySelectorAll('div > div').length : -1,
      畳んだ高さ: みな.map((x) => Array.prototype.filter
        .call(x.children, (y) => y.tagName !== 'SUMMARY').reduce((t, y) => t + 高(y), 0)),
    };
  });
  console.log('      ── 実測 ── .xlsm 中身の文「' + マクロ本.中身の文 + '」／ 畳み '
    + マクロ本.畳みの数 + '個 ／ ' + マクロ本.マクロの見出し
    + ' ／ 畳んだ 高さ ' + マクロ本.畳んだ高さ.join(',') + 'px');

  /* ★★1文の 数と 畳みの 数が 揃って いるか★★（2026-09-22）
       ★★2026-09-22 に ここで 実際に 食い違いが 出ました★★
         1文 「マクロが ★4★本」／畳み 「マクロ（★5★件）」
         ＝1文は ★かたまり(モジュール)★、畳みは ★手続き(Sub/Function)★ を 数えて いた。
         ＝★どちらも 正しい 数★。★言葉が 揃って いないだけ★。
         ⇒経営者1 が 台を 直し（`74ad9e8`）
           1文 「マクロが 4かたまり（手続き 5本）」／畳み 「マクロの 手続き（5件）」
       ★★私は ここで 1つ 見立てを 外しました★★
         「言葉を 変えても 私の 試験は 赤に ならない」と 便りに 書きましたが
         ⇒★実際は 赤に なりました★（`/マクロが [0-9]+本/` が 外れた）。
         ⇒★字の 形を 見る 門は 相手が 言葉を 直した 日に 割れます★。
       ⇒★★だから 字の 形では なく 「数が 揃って いるか」を 見ます★★
         ＝★言葉が また 変わっても 割れず、食い違いだけ 捕まえます★ */
  const 手続きの数 = /手続き ([0-9]+)本/.exec(マクロ本.中身の文);
  const 畳みの件数 = /（([0-9]+)件）/.exec(マクロ本.マクロの見出し);
  T('★★マクロ入りの 本で 数を 言える★★（「ありません」で 済ませない）',
    /マクロ/.test(マクロ本.中身の文) && /[0-9]/.test(マクロ本.中身の文)
      && !/マクロは ありません/.test(マクロ本.中身の文), マクロ本.中身の文);

  T('★★1文の 数と 畳みの 数が 揃って いる★★（09-22 に 4 と 5 が 並んだ）',
    !!手続きの数 && !!畳みの件数 && 手続きの数[1] === 畳みの件数[1],
    '1文「' + マクロ本.中身の文 + '」／畳み「' + マクロ本.マクロの見出し + '」');

  T('★マクロの ドロップダウンに 中身が 在る★（0件で 畳みだけ 出さない）',
    マクロ本.マクロの行 > 0, マクロ本.マクロの見出し + ' 行 ' + マクロ本.マクロの行);

  T('★マクロ入りでも 畳みは 全部 高さ 0★',
    マクロ本.畳んだ高さ.length > 0 && マクロ本.畳んだ高さ.every((x) => x === 0),
    マクロ本.畳んだ高さ.join(' / ') + 'px');

  await page.evaluate(() => { document.getElementById('diffCancel').click(); });
  await page.evaluate(() => window.__保存2);
  const 窓が閉じた = await page.evaluate(() =>
    document.getElementById('diffOverlay').style.display !== 'flex');
  T('2冊目でも [やめる] で 窓が 閉じる（★2冊 続けて 開いても 1本も 書き出さない★）', 窓が閉じた);
} finally {
  await page.close().catch(() => {});
  await browser.close().catch(() => {});
  配信.閉じる();
}

console.log('\n' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
