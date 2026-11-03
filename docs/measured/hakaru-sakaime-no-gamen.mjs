/* hakaru-sakaime-no-gamen.mjs ･･･ ★実Excel が 字を 出す 一番 狭い 列（境目）で うちの 画面も 同じ所で # に 切り替わるか★ 2026-10-04 夜
 *
 *  ★★なぜ★★
 *    Linux の WebKit で 列が 93点（Windows 74点）に なり、#### が 出ない 組が 7つ 出た（§18）。
 *    直し（一字の 幅を 実Excel の 表から 引く）が ★機械に よらず 同じ所で # に 切り替わるか★ は、
 *    幅を 3つ（4／8.44／60）当てる だけでは 分からない ⇒ ★組ごとの 境目の 両側★で 見る。
 *  ★★物差し★★ 経営者の 道具134 の 紙 golden-ji-wo-dasu-ichiban-semai-haba-2026-10-04.tsv（Excel 16.0.20430・游ゴシック 11）
 *    境目 ･･･「入る 一番 狭い 列(px)」で 幅60 の 字と 同じ 字 ／「入らない 一番 広い 列(px)」で ★字では ない★
 *    いつも# ･･･ 幅60 の 字が # ＝ 485px（≒幅60）でも # のまま
 *  ★★測り方★★（hakaru-haba-no-gamen と 同じ＝★本番の render() を 呼び fillText を 覗く★）
 *    ★列の 点は 本番の 入れ物 sh.colW に 直に 置く★＝★一字の幅（列の 決め方）を 通さない★
 *      ⇒ この 道具が 見るのは ★「同じ 点の 列なら 同じ所で # に なるか」＝字の 幅の 側★だけ
 *    1組ずつ ･･･ A1 と C1 に その組の マスを 写し、A 列＝入る 点・C 列＝入らない 点、描かせて 取る
 *  ★入らない 側の「空」★ ･･･ 実Excel は 狭い 列で # も 出さず .Text が 空（経営者 実測 0.1字・0.5字）。
 *    ★空を「合う」に しない★ ＝ 型「入らない側 空」で 別に 数える（合う／違う と 混ぜない）
 *  ★走らせ方★ node docs/measured/hakaru-sakaime-no-gamen.mjs [--紙 <tsv>]
 *  ★出すのは 数と 違った 組★（作り物の 値＝司さんの 本は 使わない）
 */
import path from 'node:path'; import http from 'node:http'; import fs from 'node:fs'; import os from 'node:os';
import { createRequire } from 'node:module';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { borrow, launch } from '../../scripts/_borrow-playwright.mjs';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..', '..');
const require_ = createRequire(pathToFileURL(path.join(ROOT, 'package.json')));
const XLSX = require_(path.join(ROOT, 'lib/xlsx.full.min.js'));
const 引数 = process.argv.slice(2);
const 取る = (名, 既定) => { const i = 引数.indexOf('--' + 名); return i >= 0 ? 引数[i + 1] : 既定; };
const 紙 = path.join(ROOT, 取る('紙', 'docs/measured/golden-ji-wo-dasu-ichiban-semai-haba-2026-10-04.tsv'));
/* ★--大きさ★ 物差しの 本の 既定の 字の 大きさ（pt）＝紙と 揃える（11／9／12）
   ★--許し★ 境目から 何px 離して 描かせるか（0＝両側 ぴったり／1＝±1px 内）。入る 側は 入る＋許し、入らない 側は 出－許し */
const 大きさ = Number(取る('大きさ', '11'));
const 許し = Number(取る('許し', '0'));

function 通し(式) {
  const s = String(式).replace(/^=/, '');
  const d = /^DATE\((\d+),(\d+),(\d+)\)$/.exec(s);
  if (d) { let n = (Date.UTC(+d[1], +d[2] - 1, +d[3]) - Date.UTC(1899, 11, 30)) / 864e5; if (n < 61) n -= 1; return n; }
  if (!/^[-0-9.*/ eE+()]+$/.test(s)) return null;
  return Function('return (' + s + ')')();
}
const L = fs.readFileSync(紙, 'utf8').split(/\r?\n/).filter((l) => l && l.charAt(0) !== '#');
const 頭 = L.shift().split('\t');
const 列 = (名) => { const i = 頭.indexOf(名); if (i < 0) { console.log('★紙に 列が 無い★ ' + 名); process.exit(2); } return i; };
const i組 = 列('組'), i式 = 列('値の式'), i書 = 列('書式'), i判 = 列('判じ'), i入 = 列('入る一番狭い(px)'), i出 = 列('入らない一番広い(px)'), i字 = 列('幅60の字');
const 組 = L.map((l) => l.split('\t')).map((c) => ({ 組: c[i組], 式: c[i式], 書: c[i書], 判: c[i判],
  入: Number(c[i入]), 出: Number(c[i出]), 字: c[i字], 値: 通し(c[i式]) }));
const 読めない = 組.filter((x) => x.値 === null || !isFinite(x.値));
const 使う = 組.filter((x) => x.値 !== null && isFinite(x.値) && (x.判 === '境目' || /^#+$/.test(x.字)));
console.log('[境目の 画面] 紙 ' + path.basename(紙) + ' ... ' + 組.length + '組 ／ 使う ' + 使う.length
  + '（境目 ' + 使う.filter((x) => x.判 === '境目').length + '・いつも# ' + 使う.filter((x) => x.判 !== '境目').length + '）／ 値を 読めない ' + 読めない.length);
if (使う.length + 読めない.length !== 組.length) { console.log('★使わない 組が 在る（判じが 境目でも # でも ない）★'); process.exit(2); }

/* ★材料★＝組 k を A(2k+1) に 置く だけ（描く 時に A1・C1 へ 写す） */
const ws = {};
使う.forEach((x, k) => { ws[XLSX.utils.encode_cell({ r: k + 1, c: 0 })] = { t: 'n', v: x.値, z: x.書 }; });
ws['!ref'] = 'A1:C' + (使う.length + 1);
const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'あ');
const 材料 = path.join(os.tmpdir(), 'exally-sakaime-no-gamen.xlsx');
fs.writeFileSync(材料, XLSX.write(wb, { bookType: 'xlsx', type: 'buffer' }));

const 型 = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8', '.png': 'image/png', '.svg': 'image/svg+xml', '.ico': 'image/x-icon' };
const s = http.createServer((q, r) => {
  const f = path.join(ROOT, decodeURIComponent(String(q.url).split('?')[0]).replace(/^\/+/, ''));
  if (!f.startsWith(ROOT) || !fs.existsSync(f) || fs.statSync(f).isDirectory()) { r.statusCode = 404; return r.end('no'); }
  r.setHeader('content-type', 型[path.extname(f).toLowerCase()] || 'application/octet-stream');
  fs.createReadStream(f).pipe(r);
});
await new Promise((x) => s.listen(0, '127.0.0.1', x));
const wk = await borrow('sakaime-no-gamen', 'webkit');
const br = await launch('sakaime-no-gamen', wk, {}, 'webkit');
/* ★--倍★ 画面の 細かさ（deviceScaleFactor）。経営者の 物差しは ★200% の 画面の Excel★（GetDpiForSystem＝192・2026-10-04 夜）
   ⇒ 比べる 時は ★--倍 2★ で 揃える（既定 1＝前の 測りと 同じ） */
const 倍 = Number(取る('倍', '1'));
const p = await br.newPage({ viewport: { width: 1600, height: 900 }, deviceScaleFactor: 倍 });
let 終わり値 = 0;
try {
  await p.goto('http://127.0.0.1:' + s.address().port + '/book.html', { waitUntil: 'load', timeout: 120000 });
  await p.evaluate(() => { document.body.classList.remove('exally-locked'); const o = document.getElementById('loginOv'); if (o) o.style.display = 'none'; });
  await p.setInputFiles('#bookFileInput', 材料);
  await p.waitForFunction(() => (window.sheets || []).some((x) => Object.keys(x.data || {}).length), null, { timeout: 120000 });
  await p.waitForTimeout(800);
  const 無い = await p.evaluate(() => ['render', 'switchSheet', 'yToR', 'xToC', 'cW']
    .filter((x) => typeof window[x] !== 'function').concat(window.ctx ? [] : ['ctx']));
  if (無い.length) { console.log('★道が 無い★ ' + 無い.join(' / ')); process.exit(8); }
  await p.evaluate((n) => { window.__測る大きさ = n; }, 大きさ);
  const 出 = await p.evaluate(async (L) => {
    window.switchSheet(0);
    const sh = window.sheets[window.activeSheet];
    const 元の = L.map((_, k) => sh.data[(k + 1) + ',0']);
    if (元の.some((x) => !x)) return { 誤り: '材料の マスが 読めない' };
    for (const k of Object.keys(sh.data)) delete sh.data[k];
    /* ★物差しの 本と 字体を 揃える★（経営者の 本＝Normal を 游ゴシック 11）。
       SheetJS が 書く 材料は 既定が Calibri ⇒ 板の 既定を 游ゴシック 11 に し、マスの 字体は 外す */
    sh.既定の字体名 = '游ゴシック'; sh.既定の字大 = window.__測る大きさ;
    for (const x of 元の) { delete x.fontName; delete x.fontSize; }
    const 型 = window.CanvasRenderingContext2D.prototype;
    const 元 = 型.fillText;
    const 二回待つ = () => new Promise((ok) => requestAnimationFrame(() => requestAnimationFrame(ok)));
    const 結果 = [];
    for (let k = 0; k < L.length; k++) {
      sh.data['0,0'] = JSON.parse(JSON.stringify(元の[k]));
      sh.data['0,2'] = JSON.parse(JSON.stringify(元の[k]));
      sh.colW[0] = L[k][0]; sh.colW[1] = 30; sh.colW[2] = L[k][1];
      const 描いた = { A: [], C: [] }; let 字体 = '';
      型.fillText = function (t, x, y) {
        try {
          if (this.canvas === window.ctx.canvas) {
            const m = this.getTransform ? this.getTransform() : null;
            const X = m ? m.a * x + m.e : x, Y = m ? m.d * y + m.f : y;
            const dpr = m && m.a ? m.a : 1;
            if (X / dpr > window.HDR_W && Y / dpr > window.HDR_H && window.yToR(Y / dpr) === 0) {
              const c = window.xToC(X / dpr);
              if (c === 0) { 描いた.A.push(String(t)); 字体 = this.font; } else if (c === 2) 描いた.C.push(String(t));
            }
          }
        } catch (e) { /* 取れない 物は 取らない */ }
        return 元.apply(this, arguments);
      };
      try { window.render(); await 二回待つ(); } finally { 型.fillText = 元; }
      /* ★Excel の 字を うちの 筆で 測った 幅★（境目の ずれを 出す 為・描く 物は 変えない） */
      const 筆 = document.createElement('canvas').getContext('2d');
      筆.font = 字体 || window.ctx.font;
      結果.push({ A: 描いた.A.join(''), C: 描いた.C.join(''), 回A: 描いた.A.length, 字体, 点A: window.cW(0), 点C: window.cW(2),
        字幅: 筆.measureText(String(L[k][2])).width,
        /* ★丸めない 大きさ（11pt＝14.667px）で 測った 幅★（①b の 見当＝描く 物は 変えない） */
        字幅丸めず: (() => { const f = 筆.font; 筆.font = (字体 || window.ctx.font).replace(/^(\D*?)\d+(\.\d+)?px/, '$1' + (window.__測る大きさ * 96 / 72) + 'px'); const w = 筆.measureText(String(L[k][2])).width; 筆.font = f; return w; })() });
    }
    return { 結果 };
  }, 使う.map((x) => [x.判 === '境目' ? x.入 + 許し : 485, x.判 === '境目' ? Math.max(1, x.出 - 許し) : 485, x.字]));
  if (出.誤り) { console.log('★' + 出.誤り + '★'); process.exit(8); }
  const 数 = {}; const 違 = {}; let 字体たち = new Set();
  const 足す = (型名, t) => { 数[型名] = (数[型名] || 0) + 1; if (t) (違[型名] = 違[型名] || []).push(t); };
  使う.forEach((x, k) => {
    const g = 出.結果[k]; 字体たち.add(g.字体);
    const 名 = x.組 + ' ' + x.式 + ' ' + x.書;
    if (x.判 !== '境目') {
      if (/^#+$/.test(g.A)) 足す('いつも# 合う'); else 足す('いつも# 違う', 名 + ' 画面=' + JSON.stringify(g.A));
      return;
    }
    if (g.点A !== x.入 + 許し || g.点C !== Math.max(1, x.出 - 許し)) { 足す('★列の 点が 置けて いない★', 名 + ' 点 ' + g.点A + '/' + g.点C); return; }
    /* 入る 側 */
    if (g.A === x.字) 足す('入る側 合う');
    else if (/^#+$/.test(g.A)) 足す('入る側 うち#（Excel 字）', 名 + ' ' + x.入 + 'px Excel=' + JSON.stringify(x.字));
    else 足す('入る側 字が違う', 名 + ' ' + x.入 + 'px 画面=' + JSON.stringify(g.A) + ' Excel=' + JSON.stringify(x.字));
    /* 入らない 側 */
    if (/^#+$/.test(g.C)) 足す('入らない側 合う（#）');
    else if (g.C === '') 足す('入らない側 空', 名 + ' ' + x.出 + 'px');
    else 足す('入らない側 うち字（Excel 入らない）', 名 + ' ' + x.出 + 'px 画面=' + JSON.stringify(g.C));
  });
  const 境 = 使う.filter((x) => x.判 === '境目').length;
  const 両方 = 使う.filter((x, k) => x.判 === '境目' && 出.結果[k].A === x.字 && /^#+$/.test(出.結果[k].C)).length;
  console.log('  描いた 字体 ' + [...字体たち].join(' ／ '));
  console.log('  ★境目 ' + 境 + '組の うち 両側とも 合う ' + 両方 + '★ ／ いつも# 合う ' + (数['いつも# 合う'] || 0) + '/' + (使う.length - 境));
  for (const k of Object.keys(数)) console.log('  ' + k + ' ' + 数[k]);
  /* ★ずれ★＝Excel の 境目（入る 一番 狭い px）－ うちの 筆で 測った 字の 幅。
     一定なら「余白」の 違い、散るなら 字の 幅（字体の 測り方）の 違い */
  const ずれ = 使う.map((x, k) => [x, 出.結果[k]]).filter(([x]) => x.判 === '境目').map(([x, g]) => ({ x, d: x.入 - g.字幅 }));
  const 箱 = {};
  for (const { d } of ずれ) { const b = Math.floor(d); 箱[b] = (箱[b] || 0) + 1; }
  console.log('  ★ずれ（Excel の 入る px － うちの 字の 幅）の 分かれ★ ' + Object.keys(箱).map(Number).sort((a, b) => a - b).map((b) => b + '〜' + (b + 1) + 'px:' + 箱[b]).join(' ／ '));
  const ds = ずれ.map((z) => z.d).sort((a, b) => a - b);
  if (ds.length) console.log('  最小 ' + ds[0].toFixed(2) + ' ／ 中 ' + ds[Math.floor(ds.length / 2)].toFixed(2) + ' ／ 最大 ' + ds[ds.length - 1].toFixed(2) + '（うちの 判じは 字の 幅 ＋ 余白3 で 入る）');
  if (引数.includes('--余白を探す')) {
    /* ★1つの 余白 m で「字の 幅 ＋ m ≦ 列」と 判じたら 境目の 両側を 何組 満たすか★（当てはめず 数える だけ） */
    const 境目 = 使う.map((x, k) => [x, 出.結果[k]]).filter(([x]) => x.判 === '境目');
    for (const [名, 鍵] of [['15px（今）', '字幅'], ['14.667px（丸めない）', '字幅丸めず']]) {
      const 行 = [];
      for (let m = 2; m <= 7.001; m += 0.25) {
        const n = 境目.filter(([x, g]) => g[鍵] + m <= x.入 + 許し && !(g[鍵] + m <= x.出 - 許し)).length;
        行.push(m.toFixed(2) + ':' + n);
      }
      console.log('  ★余白 m ごとの 両側 合う 数（' + 名 + '・許し ' + 許し + 'px）★ ' + 行.join(' '));
    }
  }
  if (引数.includes('--ずれを全部')) ずれ.forEach(({ x, d }) => console.log('     ずれ ' + d.toFixed(2) + ' ' + x.組 + ' ' + x.式 + ' ' + x.書 + ' ' + JSON.stringify(x.字) + ' 入る ' + x.入 + 'px'));
  for (const k of Object.keys(違)) { console.log('  ── ' + k + ' ' + 違[k].length + '組'); 違[k].forEach((t) => console.log('     ' + t)); }
  終わり値 = (両方 === 境 && (数['いつも# 合う'] || 0) === 使う.length - 境) ? 0 : 1;
} finally { await p.close(); await br.close(); s.close(); }
process.exit(終わり値);
