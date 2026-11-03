/* hakaru-mikire.mjs — ★見切れて いる マスを 数える★ 2026-09-24
 *
 *  ★★司さんの 言葉★★
 *    「★読み込んだ ファイル 見切れとん とかも 自動で 調整しろ★」
 *    「★は？ Excel内で 見切れてない とこが 見切れとるけん いよんやろが★」
 *    ⇒★Excel から 外れろ では ない★＝★Excel に 合って いない＝不具合★。
 *
 *  ★★2通りの 見切れを 分けて 数えます★★（★混ぜると 直し方を 間違えます★）
 *    ⑴★数が `####` に なって いる★
 *        ＝`_数が入らないか()` が 真 ⇒ `_井桁で埋める()`
 *        ＝★列が 狭いと 早く `####` に なる★
 *    ⑵★字が 途中で 切れて いる★
 *        ＝字の 点が ★マスの 幅★ より 広い
 *        ＝★実Excel は 右隣が 空なら はみ出させます★（切りません）
 *        ⇒だから ★右隣が 空か★ も 一緒に 数えます（★空なのに 切れて いる＝Excel と 違う★）
 *
 *  ★★お客さんの 道で 測ります★★
 *    `#bookFileInput` に 渡す＝画面の「読み込む」と 同じ。
 *    ★1バイトも 書きません★（保存の 窓は 出しません）。
 *
 *  ★★実物の 中身は 出しません★★（司さんの 決め）
 *    出すのは ★板の 名・列の 番号・数★ だけ。★マスの 値は 出しません★。
 *
 *  走らせ方: node docs/measured/hakaru-mikire.mjs [材料の道]
 */
import path from 'node:path'; import http from 'node:http'; import fs from 'node:fs';
import { fileURLToPath } from 'node:url';
import { borrow, launch } from '../../scripts/_borrow-playwright.mjs';

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
console.log('★見切れて いる マスを 数える★');
console.log('  材料 ... ' + path.basename(材料) + '（' + fs.statSync(材料).size.toLocaleString() + ' バイト）');

const wk = await borrow('mikire', 'webkit');
const browser = await launch('mikire', wk, {}, 'webkit');
const page = await browser.newPage({ viewport: { width: 1400, height: 900 } });
const 配信 = await 立てる(ROOT);

try {
  await page.goto(配信.url + '/book.html', { waitUntil: 'load', timeout: 120000 });
  await page.evaluate(() => {
    document.body.classList.remove('exally-locked');
    const ov = document.getElementById('loginOv');
    if (ov) { ov.classList.remove('open'); ov.style.display = 'none'; }
    window.FileOut = { deliver: () => Promise.reject(new Error('★書き出しては いけません★')) };
  });
  await page.setInputFiles('#bookFileInput', 材料);
  /* ★★どれか 1枚でも 中身が 入れば 開けた★★（2026-09-24）
       前は `sheets[activeSheet]` の 1枚だけを 見て いて、
       ★開いて いるのに 待ち続ける★ 事が 在りました（経営者1 の 指摘）。 */
  await page.waitForFunction(() => {
    const ss = window.sheets || [];
    for (let i = 0; i < ss.length; i++) {
      if (ss[i] && ss[i].data && Object.keys(ss[i].data).length > 0) return true;
    }
    return false;
  }, null, { timeout: 600000 });

  const 出 = await page.evaluate(() => {
    const ss = window.sheets || [];
    const 板 = [];
    let 井桁 = 0, 字切れ = 0, 右空で字切れ = 0, なお切れる = 0, 見たマス = 0;
    for (let i = 0; i < ss.length; i++) {
      const sh = ss[i];
      const d = (sh && sh.data) || {};
      const colW = sh.colW || {};
      const 既定 = sh.既定の列幅 || 72;
      let 板井桁 = 0, 板字 = 0, 板右空 = 0, 板なお = 0, 板マス = 0;
      const 列ごと = {};
      for (const k in d) {
        const cell = d[k];
        if (!cell) continue;
        const raw = (cell.v !== undefined && cell.v !== null && cell.v !== '') ? cell.v : cell.d;
        if (raw === undefined || raw === null || raw === '') continue;
        const p = k.split(',');
        const r = +p[0], c = +p[1];
        if (cell.merged) continue;                    /* 結合の 中は 別の 話 */
        板マス++; 見たマス++;
        const w = colW[c] || 既定;
        /* ★描く 所と 同じ 字で 測る★ */
        const 大 = (typeof window._マスの字大 === 'function') ? window._マスの字大(cell) : 11;
        const 書 = (typeof window._マスの書体 === 'function') ? window._マスの書体(cell) : '游ゴシック';
        const cv = document.createElement('canvas').getContext('2d');
        cv.font = (cell.bold ? 'bold ' : 'normal ') + 大 + 'px ' + 書;
        const 字 = String(cell.d !== undefined && cell.d !== '' ? cell.d : raw);
        const 点 = cv.measureText(字).width;
        const 余白 = (typeof window.マスの余白 === 'number') ? window.マスの余白 : 3;
        const 入らない = 点 > (w - 余白);
        if (!入らない) continue;
        const 数か = (typeof raw === 'number')
          || (typeof raw === 'string' && /^-?\d+(\.\d+)?([eE][+-]?\d+)?$/.test(raw));
        if (数か) { 板井桁++; 井桁++; } else {
          板字++; 字切れ++;
          const 右 = d[r + ',' + (c + 1)];
          const 右が空 = !右 || ((右.v === undefined || 右.v === null || 右.v === '')
            && (右.d === undefined || 右.d === ''));
          if (右が空) { 板右空++; 右空で字切れ++; }
          /* ★★はみ出しても なお 切れるか★★（2026-09-24 の 直しの 後）
               ＝右（左）の ★空いた マス★ を 足して いって、それでも 足りない 物を 数える。
               ＝★画面の 端で 止まる 分は ここでは 見て いません★（★未測定★）。 */
          let 広さ = w, cc = c + 1;
          for (;;) {
            const 隣 = d[r + ',' + cc];
            const 空 = !隣 || ((隣.v === undefined || 隣.v === null || 隣.v === '')
              && (隣.d === undefined || 隣.d === ''));
            if (!空) break;
            const ww = colW[cc] || 既定;
            if (!(ww > 0)) break;
            広さ += ww; cc++;
            if (広さ - 余白 >= 点) break;
            if (cc > c + 60) break;
          }
          if (広さ - 余白 < 点) { 板なお++; なお切れる++; }
        }
        列ごと[c] = (列ごと[c] || 0) + 1;
      }
      if (板井桁 || 板字) {
        const 並び = Object.keys(列ごと).map((c) => ({ 列: +c, 数: 列ごと[c] }))
          .sort((a, b) => b.数 - a.数).slice(0, 5);
        板.push({ 名: sh.name, マス: 板マス, 井桁: 板井桁, 字: 板字, 右空: 板右空, なお: 板なお, 多い列: 並び });
      }
    }
    return {
      板の数: ss.length, 見たマス, 井桁, 字切れ, 右空で字切れ, なお切れる, 板,
      標準の列幅: (ss[0] || {}).既定の列幅,
      余白: window.マスの余白,
    };
  });

  console.log('');
  console.log('★数★ 板 ' + 出.板の数 + '枚 ／ 見た マス ' + 出.見たマス.toLocaleString()
    + ' ／ 標準の列幅 ' + 出.標準の列幅 + '点 ／ マスの余白 ' + 出.余白 + '点');
  console.log('  ⑴★数が `####` に なる★ ......... ' + 出.井桁.toLocaleString() + ' マス');
  console.log('  ⑵★字が 切れる★ ................. ' + 出.字切れ.toLocaleString() + ' マス');
  console.log('      うち ★右隣が 空★ ........... ' + 出.右空で字切れ.toLocaleString() + ' マス'
    + '（★実Excel は ここを はみ出させます＝切りません★）');
  console.log('      ★★はみ出しても なお 切れる★★ .. ' + 出.なお切れる.toLocaleString() + ' マス'
    + '（★隣に 物が 在って 伸ばせない＝実Excel も 切ります★）');
  console.log('');
  console.log('★板ごと★（多い 順に 5列まで・★マスの 値は 出しません★）');
  for (const b of 出.板.sort((a, b) => (b.井桁 + b.字) - (a.井桁 + a.字))) {
    console.log('  ' + String(b.名).slice(0, 12).padEnd(13)
      + ' マス ' + String(b.マス).padStart(6)
      + ' ／ #### ' + String(b.井桁).padStart(5)
      + ' ／ 字切れ ' + String(b.字).padStart(5)
      + '（右空 ' + String(b.右空).padStart(5) + ' ／ なお ' + String(b.なお).padStart(4) + '）'
      + ' ／ 多い列 ' + b.多い列.map((x) => x.列 + ':' + x.数).join(' '));
  }
  /* ══ ★★絵を 撮る★★ ══（★数が 緑でも 絵を 開いて 見るまで OKを 出さない★）
       `--絵=<置き場>` で 板を 1枚 撮ります。
       ★★repo には 入れません★★＝司さんの 実物の 中身が 写るので 置き場は 引数で 渡します。 */
  const 絵の指定 = process.argv.find((x) => x.indexOf('--絵=') === 0);
  if (絵の指定) {
    const 置き場 = 絵の指定.slice('--絵='.length);
    const 板名 = (process.argv.find((x) => x.indexOf('--板=') === 0) || '--板=').slice('--板='.length);
    const 選んだ = await page.evaluate((な) => {
      const ss = window.sheets || [];
      let i = ss.findIndex((x) => x && x.name === な);
      if (i < 0) i = 0;
      window.switchSheet(i);
      return (ss[i] || {}).name;
    }, 板名);
    await page.waitForTimeout(1500);
    /* ★★知らせの 板を どける★★＝★絵に 写ると 表が 隠れます★
         ＝★消すのでは なく 隠すだけ★（お客さんの 画面を 作り変えない） */
    await page.evaluate(() => {
      document.querySelectorAll('div').forEach((el) => {
        const st = getComputedStyle(el);
        if (st.position !== 'fixed' && st.position !== 'absolute') return;
        const b = el.getBoundingClientRect();
        if (b.width > 300 && b.height > 200 && b.top > 200 && b.top < 700) el.style.visibility = 'hidden';
      });
    });
    await page.waitForTimeout(300);
    const 出先 = 置き場 + '/mikire-' + 選んだ + '.png';
    /* ★表の 所だけ★＝リボンや 知らせを 入れない */
    await page.screenshot({ path: 出先, clip: { x: 0, y: 300, width: 1400, height: 560 } });
    const 中 = fs.readFileSync(出先);
    console.log('');
    console.log('★絵を 撮りました★ ' + 出先);
    console.log('  板 ' + 選んだ + ' ／ ' + 中.length.toLocaleString() + ' バイト ／ sha256 '
      + (await import('node:crypto')).createHash('sha256').update(中).digest('hex'));
  }
} finally {
  await page.close().catch(() => {});
  await browser.close().catch(() => {});
  配信.閉じる();
}
