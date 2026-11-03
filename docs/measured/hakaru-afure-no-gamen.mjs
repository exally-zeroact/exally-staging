/* hakaru-afure-no-gamen.mjs ･･･ ★新しい 本に 溢れる 式を 打った 時、画面の 字が 実Excel と 同じか★ 2026-10-05
 *
 *  ★★なぜ★★ 新しい 本で `=A1:A3` を 打つと 画面は #VALUE!（E2・E3 に 溢れない）／溢れを 見る 式を 先に 打つと F1 が 0 の まま。
 *    本番 14955ae・03ddb8d とも 同じ（前から 在る）。司さんの 本は 溢れ 0 なので 受け入れでは 1度も 見えない 形。
 *  ★★物差し★★ 経営者の 道具137 の 紙 golden-jitsu-excel-no-afure-2026-10-05.tsv（Excel 16.0.20430・新しい 本・.Formula2・31行・9組）
 *  ★★測り方★★ 組ごとに 新しい 頁で 新しい 本。A1:A3＝1,2,3 を 打ち、組の 式（と 値）を 画面の 関数 setCell で 打つ（真似ない）。
 *    打つ 順 ･･･ 紙の 順。但し 組の 名に「見る式が先」が 在れば F1 を 先に 打つ。
 *    比べる 物 ･･･ ★画面の 字（sheets[0].data の d）★と★描いた 字（render の fillText）★を 実Excel の .Text と
 *  ★走らせ方★ node docs/measured/hakaru-afure-no-gamen.mjs [--台 chromium|webkit]
 */
import path from 'node:path'; import http from 'node:http'; import fs from 'node:fs';
import { fileURLToPath } from 'node:url';
import { borrow, launch } from '../../scripts/_borrow-playwright.mjs';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..', '..');
const 引数 = process.argv.slice(2);
const 取る = (名, 既定) => { const i = 引数.indexOf('--' + 名); return i >= 0 ? 引数[i + 1] : 既定; };
const 台 = 取る('台', 'chromium');
/* ★--壊す★ 画面の `_画面の誤りの字` を 素通しに する（#SPILL! を #スピル! に 換えない）＝見張りの 自己試験 用 */
const 壊す = 引数.includes('--壊す');
const 紙 = path.join(ROOT, 取る('紙', 'docs/measured/golden-jitsu-excel-no-afure-2026-10-05.tsv'));

const L = fs.readFileSync(紙, 'utf8').split(/\r?\n/).filter((l) => l && l.charAt(0) !== '#');
const 頭 = L.shift().split('\t');
const ix = (n) => { const i = 頭.indexOf(n); if (i < 0) { console.log('★紙に 列が 無い★ ' + n); process.exit(2); } return i; };
const i組 = ix('組'), i番 = ix('番地'), i式 = ix('式(Formula2)'), i字 = ix('画面の字(.Text)');
const 行 = L.map((l) => l.split('\t')).map((c) => ({ 組: c[i組], 番: c[i番], 式: c[i式] || '', 字: c[i字] || '' }));
const 組たち = [...new Set(行.map((x) => x.組))];
/* ★道具137 の 組の 定義★ `@{ 名 = '名前'; 打つ = @(@('E2', '9'), ...); 読む = ... }` から 名 と 打つ 順を 拾う */
const 道具の順 = new Map();
for (const l of fs.readFileSync(path.join(ROOT, 'docs/measured/toru-jitsu-excel-no-afure.ps1'), 'utf8').split(/\r?\n/)) {
  const m = /@\{\s*名\s*=\s*'([^']+)';\s*打つ\s*=\s*@\((.*)\);\s*読む/.exec(l);
  if (!m) continue;
  /* ★一重引用符 '...' と 二重引用符 "..."（PowerShell の `n＝改行・`t＝タブ）の 両方を 読む★（2026-10-05・組 R） */
  道具の順.set(m[1], [...m[2].matchAll(/@\('([A-Z]+\d+)',\s*(?:'((?:[^']|'')*)'|"((?:[^"`]|`.)*)")\)/g)]
    .map((x) => [x[1], x[2] !== undefined ? x[2].replace(/''/g, "'") : x[3].replace(/`n/g, '\n').replace(/`t/g, '\t').replace(/`r/g, '\r').replace(/``/g, '`')]));
}
if (道具の順.size !== 組たち.length) { console.log('★道具137 の 組 ' + 道具の順.size + ' ／ 紙の 組 ' + 組たち.length + '＝数が 合わない★'); process.exit(2); }
const 番を = (a) => { const m = /^([A-Z]+)(\d+)$/.exec(a); let c = 0; for (const ch of m[1]) c = c * 26 + ch.charCodeAt(0) - 64; return [Number(m[2]) - 1, c - 1]; };

const 型 = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8', '.png': 'image/png', '.svg': 'image/svg+xml', '.ico': 'image/x-icon' };
const s = http.createServer((q, r) => {
  const f = path.join(ROOT, decodeURIComponent(String(q.url).split('?')[0]).replace(/^\/+/, ''));
  if (!f.startsWith(ROOT) || !fs.existsSync(f) || fs.statSync(f).isDirectory()) { r.statusCode = 404; return r.end('no'); }
  r.setHeader('content-type', 型[path.extname(f).toLowerCase()] || 'application/octet-stream');
  fs.createReadStream(f).pipe(r);
});
await new Promise((x) => s.listen(0, '127.0.0.1', x));
console.log('[溢れの 画面] 紙 ' + path.basename(紙) + ' ／ ' + 行.length + '行・' + 組たち.length + '組 ／ 台 ' + 台);
const wk = await borrow('afure-no-gamen', 台);
const br = await launch('afure-no-gamen', wk, {}, 台);
let 合 = 0, 違 = 0, 描合 = 0, 待ちが立たない = 0, 受けない = 0;
const 待ち = [];
const 違い = [];
try {
  for (const 組 of 組たち) {
    const 組の行 = 行.filter((x) => x.組 === 組);
    /* ★打つ 順は 物差しを 取った 道具137 の 組の 定義を そのまま 読む★（2026-10-05）
         紙には 最後の 形しか 無い（J「9 を 打って 消す」の 消す が 紙に 出ない）＝自分で 順を 決めると 素通りで 合う */
    const 打つ = 道具の順.get(組);
    if (!打つ) { console.log('★道具137 に 組 ' + 組 + ' の 打つ 順が 無い★'); process.exit(2); }
    const 順 = [['A1', '1'], ['A2', '2'], ['A3', '3']].concat(打つ).map(([a, v]) => [...番を(a), v]);
    const 見る = 組の行.map((x) => [x.番, ...番を(x.番)]);
    const p = await br.newPage({ viewport: { width: 1280, height: 800 } });
    await p.goto('http://127.0.0.1:' + s.address().port + '/book.html', { waitUntil: 'load', timeout: 120000 });
    await p.evaluate(() => { document.body.classList.remove('exally-locked'); const o = document.getElementById('loginOv'); if (o) o.style.display = 'none'; });
    await p.waitForFunction(() => typeof window.setCell === 'function' && (window.sheets || []).length > 0, null, { timeout: 60000 });
    const 出 = await p.evaluate(async ([順, 見る, 壊す]) => {
      window.switchSheet(0);
      if (壊す) window._画面の誤りの字 = function (t) { return t; };
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
      const 型 = window.CanvasRenderingContext2D.prototype, 元 = 型.fillText, 描 = {};
      型.fillText = function (t, x, y) {
        try { if (this.canvas === window.ctx.canvas) { const m = this.getTransform(); const X = (m.a * x + m.e) / m.a, Y = (m.d * y + m.f) / m.d;
          if (X > window.HDR_W && Y > window.HDR_H) { const k = window.yToR(Y) + ',' + window.xToC(X); 描[k] = (描[k] || '') + String(t); } } } catch (e) { /* 取らない */ }
        return 元.apply(this, arguments);
      };
      try { window.render(); await new Promise((ok) => requestAnimationFrame(() => requestAnimationFrame(ok))); } finally { 型.fillText = 元; }
      return { 待った, 条件, 行: 見る.map(([a, r, c]) => ({ 番: a, 字: (function (t) { return typeof window._画面の誤りの字 === 'function' ? window._画面の誤りの字(t) : t; })(d[r + ',' + c] ? String(d[r + ',' + c].d === undefined ? '' : d[r + ',' + c].d) : ''),   /* ★中の 値は #SPILL! の まま＝見える 字に 換えて 比べる（画面の 関数を 呼ぶ）★ */ 描: 描[r + ',' + c] || '' })) };
    }, [順, 見る, 壊す]);
    待ち.push(組.slice(0, 1) + ' ' + 出.待った + 'ms' + (出.条件 ? '' : '★条件 立たず★'));
    if (!出.条件) 待ちが立たない++;
    for (const x of 組の行) {
      /* ★実Excel が 式を 受けなかった 行は 比べない★（タブ 入りの 式など・合う／違う に 数えない） */
      if (x.字.indexOf('Excel が 式を 受けない') >= 0) { 受けない++; continue; }
      const g = 出.行.find((o) => o.番 === x.番);
      if (g.字 === x.字) 合++; else { 違++; 違い.push(組 + ' ' + x.番 + ' 画面=' + JSON.stringify(g.字) + ' 実Excel=' + JSON.stringify(x.字)); }
      if (g.描 === x.字) 描合++;
    }
    await p.close();
  }
} finally { await br.close(); s.close(); }
console.log('  待った（打ち終わり → 打った 後の 計算し直しが 終わる まで）' + 待ち.join(' ／ '));
if (待ちが立たない) console.log('  ★' + 待ちが立たない + '組で 計算し直しが 終わらない まま 数えた＝その 組は 測って いない★');
const 分母 = 行.length - 受けない;
console.log('  ★画面の 字 ' + 合 + '/' + 分母 + ' ／ 描いた 字 ' + 描合 + '/' + 分母 + '★' + (受けない ? '（実Excel が 式を 受けない ' + 受けない + '行は 比べない）' : ''));
違い.forEach((t) => console.log('     ' + t));
process.exit((違 || 待ちが立たない) ? 1 : 0);
