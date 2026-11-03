/* kakareta-ji-webkit.mjs ･･･ ★画面に 実際に 描かれた 字（render() の fillText）が 実Excel と 同じか★ 2026-10-04
 *
 *  ★★なぜ 要るか★★
 *    画面の 関数を 並べて 呼ぶ 測りは、★描く 段の 抜け★を 見ない（経営者の 叩き⑴）。
 *    ⇒ render() を 呼び fillText を 覗いて 初めて ★書式付きの 値 0 が 描かれて いない★（`if(!raw) return;`）が 見えた。
 *  ★★見る 物★★
 *    ⑴ ★ゼロを 隠さない 板★の 書式付き 0 は 描く（[h]:mm → 0:00 ／ [s] → 0 ／ ge.m.d;@ → M33.1.0 ／ 0.00E+00 → 0.00E+00）
 *       ＝実Excel の 真値（経営者 道具132・幅 60・組 K／S／X）
 *    ⑵ ★ゼロを 隠す 板★（showZeros="0"）の 書式付き 0 は 描かない（司さんの 本の 3マス＝DisplayZeros False・実Excel も 空）
 *    ⑶ 幅 8.44 の 166組（経営者の 物差し）を 描かれた 字で 突き合わせる＝`docs/measured/hakaru-haba-no-gamen.mjs` を 走らせる
 *  ★--self-test★ ... `_数が入らないか` を いつも 偽に して ⑶ が 赤に なるか（経営者の 叩き⑷）
 *  ★走らせ方★: node tests/kakareta-ji-webkit.mjs [--self-test]
 */
import path from 'node:path'; import http from 'node:http'; import fs from 'node:fs'; import os from 'node:os';
import { spawnSync } from 'node:child_process';
import { createRequire } from 'node:module';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { borrow, launch } from '../scripts/_borrow-playwright.mjs';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(pathToFileURL(path.join(ROOT, 'package.json')));
const XLSX = require_(path.join(ROOT, 'lib/xlsx.full.min.js'));
const ZipSurgeon = require_(path.join(ROOT, 'lib/zip-surgeon.js'));
const 壊す = process.argv.includes('--self-test');

let pass = 0, fail = 0;
const T = (n, ok, m) => {
  if (ok) { pass++; console.log('  ok   ' + n); }
  else { fail++; console.log('  NG   ' + n + (m ? '\n       ' + m : '')); }
};
console.log('[kakareta-ji] ★画面に 描かれた 字が 実Excel と 同じか★' + (壊す ? '（--self-test＝わざと 壊す）' : ''));

const 道具 = path.join(ROOT, 'docs/measured/hakaru-haba-no-gamen.mjs');
const 紙 = 'docs/measured/golden-sel-shoshiki-haba-hyoujun-2026-10-04.tsv';
/* ★⑶ の 条件★（2026-10-04 夜・経営者と 決めた）
     ・物差しは ★200% の 画面の Excel★（経営者の 機械）＝★--倍 2★ で 揃える
     ・★境目の 紙★（道具134・游ゴシック 11）で 組ごとに「列が 境目から 1px 以内か」を 分ける
     ・★遠い 組は 全部 合う★（違う 0）
     ・★近い 組の 違いは OS ごとの 記録まで★＝字の 幅を Excel と 同じに できない（±0.6px）ので 0 は 約束しない。
       ★増えたら 赤★。下がったら 記録を 下げる（★上げない★）。記録の 無い OS は 未測定＝赤 */
const 境目 = 'docs/measured/golden-ji-wo-dasu-ichiban-semai-haba-2026-10-04.tsv';
const 三の引数 = ['--幅', '8.44', '--紙', 紙, '--倍', '2', '--境目', 境目];
const 近いの記録 = { win32: 2, linux: 2 };   /* ★2026-10-04 夜 実測★ win32＝手元 WebKit・linux＝CI（使い捨ての枝 run 37192014268） */

if (壊す) {
  /* ★先に 借りる★（10-04 夜・CI で 赤）＝毎回の CI（ci.yml）には webkit が 無い。
       借りられない 時は `borrow` が ★未測定★ の 声を 出して 緑で 終わる 決まり（ほかの webkit の 見張りと 同じ）。
       前は 子の 道具だけが 未測定で 0 を 返し、★「壊したのに 赤に ならない」と 自己試験が 赤★ に なって いた
       （手元には webkit が 在るので 総なめでは 見えなかった） */
  await borrow('kakareta-ji', 'webkit');
  /* ★わざと 壊す★＝`_数が入らないか` を いつも 偽 ⇒ Excel の # が うちで 字に なる ⇒ 赤 で なければ 見張りは 死んで いる */
  const r = spawnSync(process.execPath, [道具, ...三の引数, '--壊す'], { cwd: ROOT, encoding: 'utf8' });
  const 違 = /遠い 違う (\d+)/.exec(r.stdout || '');
  console.log('      ── 壊した 時 ── exit ' + r.status + ' ／ ' + (違 ? '遠い 違う ' + 違[1] : '出しが 読めない'));
  T('★わざと 壊すと 赤（遠い 違う 1 以上）★', 違 && Number(違[1]) > 0, (r.stdout || '') + (r.stderr || ''));
  console.log('\nkakareta-ji --self-test: ' + pass + ' 緑 / ' + fail + ' 赤');
  process.exit(fail ? 1 : 0);
}

/* ══ ⑴⑵ ★材料★ 板「みせる」（隠さない）と 板「かくす」（showZeros="0"）に 同じ 書式付き 0 ══ */
const 組 = [
  { 書: '[h]:mm', 字: '0:00' }, { 書: '[s]', 字: '0' }, { 書: 'ge.m.d;@', 字: 'M33.1.0' }, { 書: '0.00E+00', 字: '0.00E+00' },
];
function 作る() {
  const ws = {};
  組.forEach((x, i) => { ws[XLSX.utils.encode_cell({ r: i * 2, c: 0 })] = { t: 'n', v: 0, z: x.書 }; });
  ws['!ref'] = 'A1:A' + (組.length * 2);
  ws['!cols'] = [{ wch: 30 }];
  return ws;
}
const wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, 作る(), 'かくす');
XLSX.utils.book_append_sheet(wb, 作る(), 'みせる');
/* ★⑷ 結合した マス★（10-04 夕・経営者の「① の 描く 版」で 結合の 頭に ★同じ 字が 2回★／★#### ★ が 出た）
     狭い 列 3本（各 5字）を 結合し 日付（yyyy"年"m"月"）＝1マス分 では 入らない・結合した 幅 なら 入る
     ⇒★結合の 頭に 1回 だけ「2026年1月」★（1周目の 1マス分の 「####」や 2回目を 描かない） */
{
  const ws = {};
  ws.A1 = { t: 'n', v: 46023, z: 'yyyy"年"m"月"' };
  ws.B1 = { t: 'z' }; ws.C1 = { t: 'z' };
  ws['!ref'] = 'A1:C1';
  ws['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 2 } }];
  ws['!cols'] = [{ wch: 5 }, { wch: 5 }, { wch: 5 }];
  XLSX.utils.book_append_sheet(wb, ws, 'けつごう');
}
const 材料 = path.join(os.tmpdir(), 'exally-kakareta-ji.xlsx');
fs.writeFileSync(材料, XLSX.write(wb, { bookType: 'xlsx', type: 'buffer' }));
{
  /* ★借り物は showZeros を 書かない★＝包みを 開けて 書く（材料作り・zero-wo-kakusu-webkit と 同じ 手） */
  const z = ZipSurgeon.read(new Uint8Array(fs.readFileSync(材料)));
  const 名 = z.names().filter((n) => /^xl\/worksheets\/sheet1\.xml$/.test(n));
  if (!名.length) { console.log('★板1が 見つかりません★'); process.exit(8); }
  let xml = await z.text(名[0]);
  xml = xml.indexOf('<sheetViews>') < 0
    ? xml.replace('<sheetData>', '<sheetViews><sheetView showZeros="0" workbookViewId="0"/></sheetViews><sheetData>')
    : xml.replace('<sheetView ', '<sheetView showZeros="0" ');
  if (!/showZeros="0"/.test(xml)) { console.log('★showZeros を 書けない＝空振り★'); process.exit(8); }
  z.replaceText(名[0], xml);
  fs.writeFileSync(材料, Buffer.from((await z.build()).bytes));
}

const 型 = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8', '.png': 'image/png', '.svg': 'image/svg+xml', '.ico': 'image/x-icon' };
const s = http.createServer((q, r) => {
  const f = path.join(ROOT, decodeURIComponent(String(q.url).split('?')[0]).replace(/^\/+/, ''));
  if (!f.startsWith(ROOT) || !fs.existsSync(f) || fs.statSync(f).isDirectory()) { r.statusCode = 404; return r.end('no'); }
  r.setHeader('content-type', 型[path.extname(f).toLowerCase()] || 'application/octet-stream');
  fs.createReadStream(f).pipe(r);
});
await new Promise((x) => s.listen(0, '127.0.0.1', x));
const wk = await borrow('kakareta-ji', 'webkit');
const br = await launch('kakareta-ji', wk, {}, 'webkit');
const p = await br.newPage({ viewport: { width: 1280, height: 900 } });
try {
  await p.goto('http://127.0.0.1:' + s.address().port + '/book.html', { waitUntil: 'load', timeout: 120000 });
  await p.evaluate(() => { document.body.classList.remove('exally-locked'); const o = document.getElementById('loginOv'); if (o) o.style.display = 'none'; });
  await p.setInputFiles('#bookFileInput', 材料);
  await p.waitForFunction(() => (window.sheets || []).some((x) => Object.keys(x.data || {}).length), null, { timeout: 120000 });
  await p.waitForTimeout(800);
  /* ★描かれた 字を 板ごとに 取る★（render() を 呼び fillText を 覗く・取る だけ） */
  const 板の字 = await p.evaluate(async (n) => {
    const 出 = {};
    for (let i = 0; i < window.sheets.length; i++) {
      window.switchSheet(i);
      const 描いた = {};
      const 回 = {};
      /* ★結合の 中で 描かれた 字は 結合の 頭に 付ける★（中央揃えの 字は 中の マスの 座標から 描き始める） */
      const 頭へ = {};
      const dt = window.sheets[i].data || {};
      for (const k of Object.keys(dt)) {
        const me = dt[k] && dt[k].mergeEnd; if (!me) continue;
        const [r0, c0] = k.split(',').map(Number);
        for (let rr = r0; rr <= me.r; rr++) for (let cc = c0; cc <= me.c; cc++) 頭へ[rr + ',' + cc] = k;
      }
      const 型 = window.CanvasRenderingContext2D.prototype;
      const 元 = 型.fillText;
      型.fillText = function (t, x, y) {
        try {
          if (this.canvas === window.ctx.canvas) {
            const m = this.getTransform(); const d = m.a || 1;
            const X = (m.a * x + m.e) / d, Y = (m.d * y + m.f) / d;
            if (X > window.HDR_W && Y > window.HDR_H) {
              let k = window.yToR(Y) + ',' + window.xToC(X);
              if (頭へ[k]) k = 頭へ[k];
              描いた[k] = (描いた[k] || '') + String(t);
              回[k] = (回[k] || 0) + 1;
            }
          }
        } catch (e) { /* 取れない 物は 取らない */ }
        return 元.apply(this, arguments);
      };
      try { window.render(); await new Promise((ok) => requestAnimationFrame(() => requestAnimationFrame(ok))); }
      finally { 型.fillText = 元; }
      出[window.sheets[i].name] = { 隠す: !!window.sheets[i].ゼロを隠す, 字: Array.from({ length: n }, (_, k) => 描いた[(k * 2) + ',0'] || ''),
        頭: { 字: 描いた['0,0'] || '', 回: 回['0,0'] || 0 } };
    }
    return 出;
  }, 組.length);
  console.log('      ── 実測 ── みせる（隠す ' + 板の字['みせる'].隠す + '）' + JSON.stringify(板の字['みせる'].字)
    + ' ／ かくす（隠す ' + 板の字['かくす'].隠す + '）' + JSON.stringify(板の字['かくす'].字));
  T('★材料の 段★ 「かくす」は ゼロを 隠す 板／「みせる」は 隠さない', 板の字['かくす'].隠す && !板の字['みせる'].隠す);
  const 結 = 板の字['けつごう'].頭;
  console.log('      ── 実測 ── 結合の 頭 … 描いた ' + JSON.stringify(結.字) + ' ／ fillText ' + 結.回 + '回');
  T('⑷ 結合の 頭は 結合した 幅で 1回 だけ 描く（2026年1月）', 結.字 === '2026年1月' && 結.回 === 1, JSON.stringify(結));
  組.forEach((x, k) => {
    T('⑴ 隠さない 板の 書式付き 0 を 描く ' + x.書 + ' ⇒ ' + x.字, 板の字['みせる'].字[k] === x.字, '描いた=' + JSON.stringify(板の字['みせる'].字[k]));
    T('⑵ 隠す 板の 書式付き 0 は 描かない ' + x.書, 板の字['かくす'].字[k] === '', '描いた=' + JSON.stringify(板の字['かくす'].字[k]));
  });
} finally { await p.close(); await br.close(); s.close(); }

/* ══ ⑶ 幅 8.44 の 166組（経営者の 物差し）＝描かれた 字で ══ */
{
  const r = spawnSync(process.execPath, [道具, ...三の引数], { cwd: ROOT, encoding: 'utf8' });
  const m = /見た (\d+) ／ 合った (\d+) ／ 違う (\d+)★（描かれなかった (\d+)）/.exec(r.stdout || '');
  const k = /遠い 違う (\d+) ／ 近い 違う (\d+)/.exec(r.stdout || '');
  const 記録 = Object.prototype.hasOwnProperty.call(近いの記録, process.platform) ? 近いの記録[process.platform] : null;
  console.log('      ── 実測 ── 幅 8.44（倍2）... ' + (m ? '見た ' + m[1] + ' ／ 合った ' + m[2] + ' ／ 違う ' + m[3] + ' ／ 描かれなかった ' + m[4] : '出しが 読めない')
    + ' ／ ' + (k ? '遠い 違う ' + k[1] + ' ／ 近い 違う ' + k[2] : '境目の 数が 読めない') + ' ／ ' + process.platform + ' の 記録 ' + 記録);
  /* ★出しを 切らない★（10-04 夜＝末尾 400字で 切って 29組の 内訳が 見えなかった） */
  const 出し = (r.stdout || '') + (r.stderr || '');
  T('⑶ 幅 8.44 の 166組が 全部 描かれ、境目から 1px より 遠い 組は 全部 合う', !!(m && m[1] === '166' && m[4] === '0' && k && k[1] === '0'), 出し);
  T('⑶ 境目から 1px 以内の 違いが ' + process.platform + ' の 記録（' + 記録 + '）より 増えて いない', !!(記録 !== null && k && Number(k[2]) <= 記録),
    記録 === null ? '★' + process.platform + ' の 記録が 無い＝未測定（緑に しない）★\n' + 出し : 出し);
}

console.log('\nkakareta-ji: ' + pass + ' 緑 / ' + fail + ' 赤');
process.exit(fail ? 1 : 0);
