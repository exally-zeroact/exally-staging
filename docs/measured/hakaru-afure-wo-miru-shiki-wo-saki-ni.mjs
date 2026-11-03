/* hakaru-afure-wo-miru-shiki-wo-saki-ni.mjs ･･･ ★お客さんが 1マスずつ 打つ 道で、溢れを 見る 式を 先に 打つと 画面は どう なるか★ 2026-10-05
 *
 *  ★★なぜ★★ 台の 見張り（shiki-hyou-matomete ④）で、台に 1つずつ 打つ 形は
 *    F1 `=E2+E3` を 先に 打ち、後から E1 `=A1:A3`（E2・E3 に 溢れる）を 打つと ★F1 が 0 の まま★だった（実Excel は 5）。
 *    経営者の 問い：お客さんが 1マスずつ 打つ 道（setCell）でも 出るか。
 *  ★測り方★ 新しい 本で 画面の 関数 setCell を 呼ぶ（真似ない）。打つ 順を 2通り：
 *    ㋐ 材料 → E1 → F1（溢れが 先）／㋑ F1 → 材料 → E1（溢れを 見る 式が 先）
 *    出すのは 画面の 字（sheets[0].data の d）と 描いた字（render の fillText）
 *  ★走らせ方★ node docs/measured/hakaru-afure-wo-miru-shiki-wo-saki-ni.mjs [--台 chromium|webkit]
 */
import path from 'node:path'; import http from 'node:http'; import fs from 'node:fs';
import { fileURLToPath } from 'node:url';
import { borrow, launch } from '../../scripts/_borrow-playwright.mjs';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..', '..');
const 引数 = process.argv.slice(2);
const 取る = (名, 既定) => { const i = 引数.indexOf('--' + 名); return i >= 0 ? 引数[i + 1] : 既定; };
const 台 = 取る('台', 'chromium');

const 型 = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8', '.png': 'image/png', '.svg': 'image/svg+xml', '.ico': 'image/x-icon' };
const s = http.createServer((q, r) => {
  const f = path.join(ROOT, decodeURIComponent(String(q.url).split('?')[0]).replace(/^\/+/, ''));
  if (!f.startsWith(ROOT) || !fs.existsSync(f) || fs.statSync(f).isDirectory()) { r.statusCode = 404; return r.end('no'); }
  r.setHeader('content-type', 型[path.extname(f).toLowerCase()] || 'application/octet-stream');
  fs.createReadStream(f).pipe(r);
});
await new Promise((x) => s.listen(0, '127.0.0.1', x));
const wk = await borrow('afure-wo-miru', 台);
const br = await launch('afure-wo-miru', wk, {}, 台);
let 終わり値 = 0;
try {
  for (const [名, 順] of [['㋐ 溢れが 先', [[0, 0, '1'], [1, 0, '2'], [2, 0, '3'], [0, 4, '=A1:A3'], [0, 5, '=E2+E3']]],
                          ['㋑ 溢れを 見る 式が 先', [[0, 5, '=E2+E3'], [0, 0, '1'], [1, 0, '2'], [2, 0, '3'], [0, 4, '=A1:A3']]]]) {
    const p = await br.newPage({ viewport: { width: 1280, height: 800 } });
    await p.goto('http://127.0.0.1:' + s.address().port + '/book.html', { waitUntil: 'load', timeout: 120000 });
    await p.evaluate(() => { document.body.classList.remove('exally-locked'); const o = document.getElementById('loginOv'); if (o) o.style.display = 'none'; });
    await p.waitForFunction(() => typeof window.setCell === 'function' && (window.sheets || []).length > 0, null, { timeout: 60000 });
    const 出 = await p.evaluate(async (順) => {
      window.switchSheet(0);
      /* ★★条件で 待つ★★（2026-10-05・経営者の 叩き）＝打った 後の 計算し直し（_scheduleRecalc→recalcSheet）は 150ms 後に 走る。
           決まった 時間では 待たない＝★打ち終わった 後に 始まった recalcSheet が 終わるまで★ 待つ（上限 30秒）*/
      const 元の再計算 = window.recalcSheet; const 終わり時刻 = [];
      window.recalcSheet = function () { const t0 = performance.now(); try { return 元の再計算.apply(this, arguments); } finally { 終わり時刻.push([t0, performance.now()]); } };
      for (const [r, c, v] of 順) window.setCell(r, c, v);
      const 打ち終わり = performance.now();
      while (!終わり時刻.some(([t0]) => t0 >= 打ち終わり) && performance.now() - 打ち終わり < 30000) await new Promise((ok) => setTimeout(ok, 20));
      const 待った = Math.round(performance.now() - 打ち終わり), 条件 = 終わり時刻.some(([t0]) => t0 >= 打ち終わり);
      window.recalcSheet = 元の再計算;
      await new Promise((ok) => requestAnimationFrame(() => requestAnimationFrame(ok)));
      const d = window.sheets[0].data;
      const 字 = (k) => (d[k] ? String(d[k].d === undefined ? '' : d[k].d) : '（無い）');
      /* ★描いた字★ */
      const 型 = window.CanvasRenderingContext2D.prototype, 元 = 型.fillText, 描 = {};
      型.fillText = function (t, x, y) {
        try { if (this.canvas === window.ctx.canvas) { const m = this.getTransform(); const X = (m.a * x + m.e) / m.a, Y = (m.d * y + m.f) / m.d;
          if (X > window.HDR_W && Y > window.HDR_H) { const k = window.yToR(Y) + ',' + window.xToC(X); 描[k] = (描[k] || '') + String(t); } } } catch (e) { /* 取らない */ }
        return 元.apply(this, arguments);
      };
      try { window.render(); await new Promise((ok) => requestAnimationFrame(() => requestAnimationFrame(ok))); } finally { 型.fillText = 元; }
      return { 待った, 条件, E1: 字('0,4'), E2: 字('1,4'), E3: 字('2,4'), F1: 字('0,5'), 描いたF1: 描['0,5'] || '', 描いたE2: 描['1,4'] || '' };
    }, 順);
    console.log('  ' + 名 + ' ... 画面の 字 E1=' + 出.E1 + ' E2=' + 出.E2 + ' E3=' + 出.E3 + ' ★F1=' + 出.F1 + '★ ／ 描いた F1=' + JSON.stringify(出.描いたF1) + ' E2=' + JSON.stringify(出.描いたE2) + '（実Excel F1=5）／ 待った ' + 出.待った + 'ms' + (出.条件 ? '' : ' ★条件 立たず＝測って いない★'));
    if (出.F1 !== '5' || 出.描いたF1 !== '5' || !出.条件) 終わり値 = 1;
    await p.close();
  }
} finally { await br.close(); s.close(); }
console.log(終わり値 ? '★実Excel と 違う 形が 在る★' : '★どちらの 順でも F1＝5（実Excel と 同じ）★');
process.exit(終わり値);
