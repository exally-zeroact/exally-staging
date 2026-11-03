/* osu-jitsubutsu-deruji.mjs — ★司さんの実物の「出る字」を うちと 突き合わせる★（2026-09-11）
 *
 *  ★★なぜ★★
 *    今まで 実物で 合わせたのは ★答え★だけ（2026-08-29 … 19,323/19,323 一致）。
 *    ★出る字は 一度も 突き合わせて いません★。
 *    2026-09-11 に 小さい ファイル 19本で 測ったら ★出る字だけで 本当の 穴が 4つ★ 出ました。
 *
 *  ★★お客さんの 道で 開きます★★
 *    本物の ブラウザで book.html を 開き、★「Excelを読み込む」の 入口★に 写しを 渡す。
 *    出る字は ★画面が 本当に 描いた 字（fillText）★を 取ります＝★作り直しません★。
 *
 *  ★司さんの 実物には 触りません★＝写し（scratchpad）だけ。
 *  ★この 道具は 自分の 台を 建てません★＝計算するのは 画面の 中の 本番の エンジン。
 *
 *  使い方:
 *    ① pwsh -NoProfile -File docs/measured/toru-jitsubutsu-deruji.ps1
 *    ② node docs/measured/osu-jitsubutsu-deruji.mjs
 */
import fs from 'node:fs';
import http from 'node:http';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { borrow, launch, unmeasured } from '../../scripts/_borrow-playwright.mjs';

const ROOT = path.resolve(path.join(path.dirname(fileURLToPath(import.meta.url)), '..', '..'));
const TAG = 'jitsubutsu-deruji';
const 写し = 'C:/Users/zeroa/AppData/Local/Temp/claude/C--WINDOWS-System32-WindowsPowerShell-v1-0/5b4e50e6-20a1-4af5-8ffb-8b6d6ca3f52b/scratchpad/jitsubutsu.xlsb';
const 板名 = process.argv[2] || '';
const 紙 = path.join(ROOT, 'docs/measured/golden-jitsubutsu-deruji-2026-09-11.tsv');
const 出す先 = path.join(ROOT, 'docs/measured/golden-jitsubutsu-deruji-awase-2026-09-11.tsv');

function 立てる(root) {
  const 型 = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8',
    '.mjs': 'text/javascript; charset=utf-8', '.css': 'text/css; charset=utf-8',
    '.json': 'application/json; charset=utf-8', '.png': 'image/png', '.svg': 'image/svg+xml' };
  const s = http.createServer((req, res) => {
    const 道 = decodeURIComponent(String(req.url).split('?')[0]);
    const f = path.join(root, 道.replace(/^[/]+/, ''));
    if (!f.startsWith(root) || !fs.existsSync(f) || fs.statSync(f).isDirectory()) { res.statusCode = 404; return res.end('no'); }
    res.setHeader('content-type', 型[path.extname(f).toLowerCase()] || 'application/octet-stream');
    fs.createReadStream(f).pipe(res);
  });
  return new Promise((r) => s.listen(0, '127.0.0.1', () => {
    r({ url: 'http://127.0.0.1:' + s.address().port, 閉じる: () => s.close() });
  }));
}

if (!fs.existsSync(紙) || !fs.existsSync(写し)) {
  console.log('  ★未測定★ 先に `pwsh -NoProfile -File docs/measured/toru-jitsubutsu-deruji.ps1` を 走らせて ください');
  process.exit(0);
}

let 紙の板 = '';
const 実 = [];
for (const l of fs.readFileSync(紙, 'utf-8').split(/\r?\n/)) {
  { const m = /^# ★板★ … (.+)$/.exec(l); if (m) 紙の板 = m[1].trim(); }
  if (!l || l.startsWith('#')) continue;
  const c = l.split('\t');
  if (c.length < 5) continue;
  実.push({ マス: c[0], 字: c[1], 型: c[2], 幅: c[3], 書式: c[4] });
}

const chromium = await borrow(TAG, 'chromium');
if (!chromium) { unmeasured(TAG, 'chromium'); process.exit(0); }
const browser = await launch(TAG, chromium, {}, 'chromium');
if (!browser) { unmeasured(TAG, 'chromium'); process.exit(0); }
const page = await browser.newPage({ viewport: { width: 1600, height: 900 } });
const 配信 = await 立てる(ROOT);

let 合 = 0, 違 = 0;
const 行 = [];
const 実物 = [];

try {
  await page.goto(配信.url + '/book.html', { waitUntil: 'load' });
  await page.evaluate(() => {
    document.body.classList.remove('exally-locked');
    const ov = document.getElementById('loginOv');
    if (ov) ov.style.display = 'none';
  });
  await page.waitForTimeout(600);
  const 入口 = await page.waitForSelector('#bookFileInput', { state: 'attached', timeout: 15000 });
  await 入口.setInputFiles(写し);
  await page.waitForFunction(() => {
    try { return Object.keys(sheets[activeSheet].data).length > 50; } catch (e) { return false; }
  }, null, { timeout: 90000 });
  await page.waitForTimeout(4000);
  /* ★板を 選ぶ★＝お客さんと 同じ 道（タブを 押す） */
  if (板名) {
    const 出来た = await page.evaluate((n) => {
      /* ★番号でも 指せます★＝名前は 貝殻で 文字化けする事が 在る */
      const i = /^[0-9]+$/.test(String(n)) ? (Number(n) - 1) : sheets.findIndex((s) => s.name === n);
      if (i < 0 || i >= sheets.length) return false;
      if (typeof switchSheet === 'function') switchSheet(i);
      else { activeSheet = i; if (typeof loadSheetIntoEngine === 'function') loadSheetIntoEngine(i); render(); }
      return true;
    }, 板名);
    if (!出来た) throw new Error('★板が 無い … ' + 板名 + '★');
    await page.waitForTimeout(4000);
  }
  /* ★★同じ 板を 見て いるかを 先に 確かめます★★（2026-09-11）
     ★違う 板を 比べて いたら 数は 全部 無意味★＝
     実際 ★板の 番号が ずれて 933本 違う★ように 見えました。
     ⇒★紙に 書いて ある 板の 名前と 突き合わせます★ */
  const うちの板 = await page.evaluate(() => sheets[activeSheet].name);
  if (紙の板 && String(うちの板) !== String(紙の板)) {
    throw new Error('★違う 板を 見て います★  実Excel「' + 紙の板 + '」／うち「' + うちの板 + '」'
      + '  ⇒★数を 出す 前に 止めます★');
  }
  console.log('★板★ ' + うちの板 + '（実Excel と 同じ）');

  /* ★★描く 所を 覗きます★★（作り直した 字では なく ★描かれた 字★）
     ★★`render()` は その場で 描きません★★（2026-09-11 ここで つまずきました）
       次の 描き直しを ★予約するだけ★＝同じ 一息で 読むと ★0本★に なります
       ⇒★予約して から 待って、それから 読みます★
     ★画面に 出て いる 所しか 描かれません★
       ⇒★画面を 動かしながら 何度も 描いて 集めます★（お客さんと 同じ 動かし方） */
  /* ★★どの マスを 描いた 時の 字かを 直に 取ります★★（2026-09-11）
     ★★前の 版は 座標で 当てて いて、取りこぼしました★★
       画面は 描く 途中で ★自分で スクロールの 値を 動かして 戻します★（固定・分割の 為）。
       ⇒ 描いた 時と 読む 時で ずれ、★売上表で 933マスが「空」に 見えました★。
       ⇒★絵を 撮ったら ちゃんと 出て いました★＝★物差しの 方が 間違い★。
     ⇒★`drawText(r, c)` を 包んで、その 中で 出た 字を そのまま 受けます★
       ＝★座標の 当て推量が 1つも 要りません★ */
  await page.evaluate(() => {
    window.__字 = {};
    const 元描く = window.drawText;
    const 元字 = CanvasRenderingContext2D.prototype.fillText;
    let 今 = null;
    window.__座 = [];
    CanvasRenderingContext2D.prototype.fillText = function (t, x, y) {
      try {
        /* ★★最後に 描かれた 字を 取ります★★（2026-09-11）
           ★画面は 先に ふつうの マスとして 描き、その上から 結合の 分で 塗り直します★
           ⇒ 最初の 字を 取ると ★結合が `########` に 見えます★（実際 21本 そう 見えた）
           ⇒★お客さんが 見るのは 最後に 描かれた 物★ */
        if (今) { window.__字[今] = String(t); }
        else {
          /* ★★`drawText` を 通らない 描き所が 在ります★★（2026-09-11 実測）
             ★結合した マスは `_renderPass` の 中で 直に 描いて います★
             ⇒ その分は ★場所で 拾います★（絶対の 位置＝画面の ずれを 足し戻す）
           ★★`xToC`/`yToR` で マスを 当てる のは やめました★★（2026-09-11）
             見出しの 字まで 拾って ★合う 数が 大きく 減りました★
             （板12 … 5,437 → 3,775）⇒★元に 戻しました★ */
          window.__座.push([String(t), x + scrollLeft - HDR_W, y + scrollTop - HDR_H]);
        }
      } catch (e) { /* 何も しない */ }
      return 元字.apply(this, arguments);
    };
    window.drawText = function (r, c) {
      const 前 = 今;
      今 = r + ',' + c;
      try { return 元描く.apply(this, arguments); } finally { 今 = 前; }
    };
    /* ★動かす 順番を 先に 作ります★＝`render()` は ★予約だけ★なので
       まとめて 呼ぶと ★最後の 1回しか 描かれません★（実測 … 239マスしか 取れなかった）
       ⇒★1歩 動かして 1回 待つ★を 繰り返します */
    window.__順番 = () => {
      const sh = sheets[activeSheet];
      let maxR = 0, maxC = 0;
      for (const k in sh.data) { const p = k.split(','); if (+p[0] > maxR) maxR = +p[0]; if (+p[1] > maxC) maxC = +p[1]; }
      const 縦 = []; { let y = 0, r = 0; while (r <= maxR) { 縦.push(y); let d = 0, m = 0;
        while (r + m <= maxR && d < 300) { d += (sh.rowH[r + m] || ROW_H); m++; } y += d; r += m; } }
      const 横 = []; { let x = 0, c = 0; while (c <= maxC) { 横.push(x); let d = 0, n = 0;
        while (c + n <= maxC && d < 400) { d += (sh.colW[c + n] || COL_W); n++; } x += d; c += n; } }
      const 出 = [];
      for (const y of 縦) for (const x of 横) 出.push([y, x]);
      return 出;
    };
    window.__一歩 = (y, x) => { scrollTop = y; scrollLeft = x; render(); };
  });

  /* ★1歩 動かして 1回 待つ★（`render()` は 予約だけ） */
  const 順 = await page.evaluate(() => window.__順番());
  console.log('★動かす 回数 … ' + 順.length + '★');
  for (let i = 0; i < 順.length; i++) {
    await page.evaluate(([y, x]) => window.__一歩(y, x), 順[i]);
    await page.waitForTimeout(45);
  }
  await page.waitForTimeout(500);
  const 集め = await page.evaluate(() => window.__字);
  console.log('★描かれた マス … ' + Object.keys(集め).length + '★');

  const 場所 = (n) => {
    const m = /^([A-Z]+)([0-9]+)$/.exec(n);
    let c = 0; for (const ch of m[1]) c = c * 26 + (ch.charCodeAt(0) - 64);
    return { r: Number(m[2]) - 1, c: c - 1 };
  };
  const 描いた = {};
  const まだ = [];
  for (const e of 実) {
    const p = 場所(e.マス);
    const v = 集め[p.r + ',' + p.c];
    描いた[e.マス] = v === undefined ? '' : v;
    if (v === undefined) まだ.push({ 名: e.マス, r: p.r, c: p.c });
  }
  /* ★★取れなかった 分は その マスに 寄って もう一度★★
     ★空を「うちは 何も 出さない」と 決めつけない★＝
     2026-09-11 に ★933マスを 穴だと 見誤り★ました（絵を 撮ったら 出て いた） */
  if (まだ.length) {
    console.log('★もう一度 寄って 測る … ' + まだ.length + 'マス★');
    for (const x of まだ) {
      await page.evaluate(([r, c]) => {
        const sh = sheets[activeSheet];
        let y = 0; for (let i = 0; i < r; i++) y += (sh.rowH[i] || ROW_H);
        let X = 0; for (let j = 0; j < c; j++) X += (sh.colW[j] || COL_W);
        window.__一歩(Math.max(0, y - ROW_H), Math.max(0, X - COL_W));
      }, [x.r, x.c]);
      await page.waitForTimeout(45);
      const v = await page.evaluate(([k, r, c]) => {
        if (window.__字[k] !== undefined) return window.__字[k];
        /* ★結合した マス＝場所で 拾う★（その マスの 箱に 入って いる 字） */
        const sh = sheets[activeSheet];
        let y = 0; for (let i = 0; i < r; i++) y += (sh.rowH[i] || ROW_H);
        let X = 0; for (let j = 0; j < c; j++) X += (sh.colW[j] || COL_W);
        const w = (sh.colW[c] || COL_W), h = (sh.rowH[r] || ROW_H);
        for (const [t, tx, ty] of window.__座) {
          if (tx >= X - 2 && tx <= X + w + 2 && ty >= y - 2 && ty <= y + h + 2) return t;
        }
        return undefined;
      }, [x.r + ',' + x.c, x.r, x.c]);
      if (v !== undefined) 描いた[x.名] = v;
    }
  }

  行.push('# ★司さんの実物の「出る字」を うちと 突き合わせた★（2026-09-11）');
  行.push('#');
  行.push('# ★今まで 実物で 合わせたのは 答えだけ★（08-29 … 19,323/19,323）');
  行.push('#   ★出る字は 今日 初めて 突き合わせました★');
  行.push('# ★画面が 本当に 描いた 字（fillText）★を 取って います＝作り直して いません');
  行.push('#');
  行.push(['# マス', '実Excel の 出る字', 'うちの 出る字', '判じ', '型', '列の点', '書式'].join('\t'));

  for (const e of 実) {
    const u = 描いた[e.マス] === undefined ? '(描かれない)' : 描いた[e.マス];
    const よい = String(u) === String(e.字);
    if (よい) 合++; else { 違++; if (実物.length < 40) 実物.push(e.マス + ' 実Excel="' + e.字 + '" うち="' + u + '"（' + e.型 + '／幅' + e.幅 + '／' + e.書式 + '）'); }
    行.push([e.マス, e.字, u, よい ? '合った' : '★違う★', e.型, e.幅, e.書式].join('\t'));
  }
} catch (e) {
  違++;
  console.error('★落ちた … ' + e.message + '★');
} finally {
  await browser.close().catch(() => {});
  配信.閉じる();
}

行.push('#');
行.push('#');
行.push('# ★★まだ 取れない 物（★物差しの 限り★・アプリの 穴では ない）★★');
行.push('#   ★結合した マス★ … 画面は 先に ふつうの マスとして `########` を 描き、');
行.push('#     その上から 結合の 分で 塗り直します。');
行.push('#     `drawText` を 包む この 取り方では ★下の `########` を 拾って しまいます★。');
行.push('#   ★絵で 確かめました★ … 実際は `2026年1月` `640,098 円` と 正しく 出て います。');
行.push('#   ⇒★合わない と 出たら まず 絵を 撮る事★（2026-09-11 に 2回 助かりました）');
行.push('# ★★締め★★ 全 ' + (合 + 違) + '本 ／ 合った ' + 合 + ' ／ ★違う ' + 違 + '★');
fs.writeFileSync(出す先, 行.join('\n') + '\n', 'utf-8');
console.log('');
console.log('★★締め★★ 全 ' + (合 + 違) + '本 ／ 合った ' + 合 + ' ／ ★違う ' + 違 + '★');
if (実物.length) { console.log(''); console.log('★合わない 物（先頭 ' + 実物.length + '本）★'); 実物.forEach((x) => console.log('  ' + x)); }
console.log('');
console.log('★書いた … ' + 出す先 + '★');
process.exit(違 ? 1 : 0);
