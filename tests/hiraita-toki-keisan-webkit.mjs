/* hiraita-toki-keisan-webkit.mjs ･･･ ★理由が 1つも 無い 本は 開いた 直後に 計算しない★ 2026-09-25
 *
 *  ★★なぜ 要るか★★
 *    司さんの 実物は 開くのに ★23.7秒★。★読み込みは 1〜2%★ で
 *    残りは ★開いた 直後の 計算★（経営者1 の CPU の 記録・87%が `lib/shiki-hyou.js`）。
 *    ★実Excel は 保存済みの 答えを 持って いて、開いた だけでは 計算しません★。
 *    ⇒理由が 1つも 無い 本は 計算しません ⇒★23.7秒 → 1.7秒★（前後 交互・webkit）
 *
 *  ★★この 試験の 骨（★印では なく 出る 字で 数えます★）★★
 *    ★★答えを わざと 違えて 保存した 式★★ を 入れます:
 *      A2 = 2 ／ A3 = `=A2*10` で ★保存された 答えは 999★（★正しくは 20★）
 *    ⇒★計算しなければ 999 が 出ます★（実Excel と 同じ）
 *    ⇒★計算すれば 20 に なります★
 *    ＝★★どちらに 転んでも 出る 字で 分かります★★
 *      （★止め過ぎ・止め足らず の 両方で 赤に なる★）
 *
 *  ★★見る 物 5つ★★
 *    ⑴★理由が 無い 本 ･･･ 999 の まま★（＝計算して いない）
 *    ⑵★その 本で 1マス 打つと 20 に なる★（＝★止めたのは 「開いた 直後」だけ★）
 *    ⑶★`=TODAY()` が 在る 本 ･･･ 20★（㋐日で 変わる 関数）
 *    ⑷★答えが 保存されて いない 式が 在る 本 ･･･ 20★（㋑）
 *    ⑸★`fullCalcOnLoad="1"` の 本 ･･･ 20★（㋓印）
 *
 *  ★走らせ方★: node tests/hiraita-toki-keisan-webkit.mjs
 */
import path from 'node:path'; import http from 'node:http'; import fs from 'node:fs';
import os from 'node:os';
import { createRequire } from 'node:module';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { borrow, launch } from '../scripts/_borrow-playwright.mjs';
import { 道を確かめる, 道の字, 画面の字を作る } from '../docs/measured/_gamen-no-michi.mjs';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(pathToFileURL(path.join(ROOT, 'package.json')));
const XLSX = require_(path.join(ROOT, 'lib/xlsx.full.min.js'));
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

/* ══ ★材料★ ══（★`.xlsx` で 作ります＝札を 字で 書けるので★）
     A1 = 見出し（字）
     A2 = 2
     A3 = `=A2*10` ･･･ ★保存された 答えは 999（わざと 違えて います・正しくは 20）★
     A4 = `=TODAY()`（★㋐の 本だけ★）
     A5 = `=A2+1` ･･･ ★答えを 保存しない（㋑の 本だけ）★ */
function 作る(なに) {
  const ws = XLSX.utils.aoa_to_sheet([['みだし'], [2], [null], [null], [null]]);
  /* ★A2（＝2）を 掛けます★＝計算すれば ★20★（★`A1*10` に すると `#VALUE!` に なり
       「999 で ない」しか 言えません＝★当てる 数を 決められる 形に します★） */
  ws['A3'] = { t: 'n', f: 'A2*10', v: 999 };   /* ★999 は わざと（正しくは 20）★ */
  if (なに === '日で変わる') ws['A4'] = { t: 'n', f: 'TODAY()', v: 46000 };
  if (なに === '答えが無い') ws['A5'] = { t: 'n', f: 'A2+1' };
  ws['!ref'] = 'A1:A5';
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'あ');
  const 道 = path.join(os.tmpdir(), 'exally-keisan-' + なに + '.xlsx');
  fs.writeFileSync(道, XLSX.write(wb, { bookType: 'xlsx', type: 'buffer' }));
  return 道;
}

/* ★借り物は 「答えを 書かない」も `fullCalcOnLoad` も 書いて くれません★
   ⇒★包みを 開けて 自分で 直します★（★これは 材料作りです＝本番の 道では ありません★）
   ⇒★直せて いなければ その場で 落とします★（＝★空振りの 緑を 作らない★） */
async function 手で直す(道, なに) {
  const z = ZipSurgeon.read(new Uint8Array(fs.readFileSync(道)));
  if (なに === '答えが無い') {
    const n = z.names().filter((x) => /^xl\/worksheets\/sheet1\.xml$/.test(x))[0];
    if (!n) throw new Error('★板1が 見つかりません★');
    let xml = await z.text(n);
    if (!/<c r="A5"/.test(xml)) throw new Error('★A5 が 書かれて いません＝この 検査が 空振り★');
    xml = xml.replace(/(<c r="A5"[^>]*>)(<f>[^<]*<\/f>)<v>[^<]*<\/v>/, '$1$2');
    if (/<c r="A5"[^>]*><f>[^<]*<\/f><v>/.test(xml)) throw new Error('★A5 の 答えを 消せて いません＝この 検査が 空振り★');
    z.replaceText(n, xml);
  }
  if (なに === '印') {
    const n = z.names().filter((x) => /^xl\/workbook\.xml$/.test(x))[0];
    if (!n) throw new Error('★workbook.xml が 見つかりません★');
    let xml = await z.text(n);
    if (/<calcPr[\s/>]/.test(xml)) xml = xml.replace(/<calcPr/, '<calcPr fullCalcOnLoad="1" ');
    else xml = xml.replace('</workbook>', '<calcPr fullCalcOnLoad="1"/></workbook>');
    if (!/fullCalcOnLoad="1"/.test(xml)) throw new Error('★印を 立てられて いません＝この 検査が 空振り★');
    z.replaceText(n, xml);
  }
  fs.writeFileSync(道, Buffer.from((await z.build()).bytes));
  return 道;
}

console.log('[hiraita-toki-keisan] ★理由が 1つも 無い 本は 開いた 直後に 計算しない★');
console.log('  ★骨★ A3 = `=A2*10` ／ ★保存された 答えは 999（正しくは 20）★');
console.log('        ⇒★計算しなければ 999 の まま／計算すれば 20★＝出る 字で 数えます');

const 本たち = [
  { なに: '理由なし', 計算するか: false, 訳: '★理由が 1つも 無い（㋐0 ㋑0 ㋓false）★' },
  { なに: '日で変わる', 計算するか: true, 訳: '㋐`=TODAY()` が 在る' },
  { なに: '答えが無い', 計算するか: true, 訳: '㋑A5 の 答えが 保存されて いない' },
  { なに: '印', 計算するか: true, 訳: '㋓`fullCalcOnLoad="1"`' },
];
for (const 本 of 本たち) 本.道 = await 手で直す(作る(本.なに), 本.なに);

const wk = await borrow('hiraita-toki-keisan', 'webkit');
const browser = await launch('hiraita-toki-keisan', wk, {}, 'webkit');
const 配信 = await 立てる(ROOT);
let 道を見た = false;

try {
  for (const 本 of 本たち) {
    const page = await browser.newPage({ viewport: { width: 1280, height: 900 } });
    try {
      page.on('pageerror', (e) => console.log('      [落ちた] ' + String(e.message).slice(0, 200)));
      let 声 = '';
      page.on('console', (m) => { const t = String(m.text()); if (t.indexOf('開いた 直後の 計算') >= 0) 声 = t; });
      await page.goto(配信.url + '/book.html', { waitUntil: 'load', timeout: 120000 });
      await page.evaluate(() => {
        document.body.classList.remove('exally-locked');
        const ov = document.getElementById('loginOv');
        if (ov) { ov.classList.remove('open'); ov.style.display = 'none'; }
        window.FileOut = { deliver: () => Promise.reject(new Error('★書き出しては いけません★')) };
      });
      await page.setInputFiles('#bookFileInput', 本.道);
      await page.waitForFunction(() => {
        const ss = window.sheets || [];
        return ss.some((s) => s && s.data && Object.keys(s.data).length > 0);
      }, null, { timeout: 120000 });
      await page.waitForTimeout(600);

      /* ★画面の 事を 測る 時は 画面の 関数を 呼ぶ。真似ない★（2026-09-25 の 決め） */
      if (!道を見た) { await 道を確かめる(page); 道を見た = true; }

      const 出 = await 画面の字を作る(page, 'あ', '2,0');   /* ★A3＝3行1列＝0から 2,0★ */
      const 字 = String(出 && 出.字);
      console.log('      ── ' + 本.なに + ' ── ' + 本.訳);
      console.log('         画面に 出た 字 ... ' + JSON.stringify(字)
        + '（' + (本.計算するか ? '★20 に なる はず★' : '★999 の まま の はず★') + '）');
      if (声) console.log('         [画面の 声] ' + 声.slice(0, 170));
      T('★' + 本.なに + ' ⇒ ' + (本.計算するか ? '計算する（20）' : '計算しない（999 の まま）') + '★',
        字 === (本.計算するか ? '20' : '999'),
        '出た=' + JSON.stringify(字) + ' ／ 声=' + 声);

      if (本.なに === '理由なし') {
        /* ★★止めて いるのは 「開いた 直後」だけ★★
           ＝★1マスでも 打ったら いつも通りに 戻ります★
           ＝★ここが 割れると 「打っても 直らない 本」を 客に 出します★ */
        await page.evaluate(() => { window.setCell(9, 0, '1'); });
        await page.waitForTimeout(1200);
        const 後 = await 画面の字を作る(page, 'あ', '2,0');
        const 後字 = String(後 && 後.字);
        console.log('         ★1マス 打った 後★ ... ' + JSON.stringify(後字) + '（★20 に なる はず★）');
        T('★★1マス 打つと いつも通りに 戻る（999 → 20）★★', 後字 === '20',
          '打った 後=' + JSON.stringify(後字));
      }
    } finally {
      await page.close().catch(() => {});
    }
  }
} finally {
  await browser.close().catch(() => {});
  配信.閉じる();
}

console.log('');
console.log('  ★道★ ' + 道の字);
console.log('hiraita-toki-keisan: ' + pass + ' 緑 / ' + fail + ' 赤');
process.exit(fail ? 1 : 0);
