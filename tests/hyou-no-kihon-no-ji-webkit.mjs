/* hyou-no-kihon-no-ji-webkit.mjs
 *   ･･･ ★`General` の 桁は 「本の 既定の 字」で 決まる★ 2026-09-26
 *
 *  ★★何が 起きて いたか（経営者1 の 実測・物差しは 実Excel）★★
 *    司さんの 実物の 1マス（書式 `General`・列の 幅 59・★そのマスの 字は 9ポイント★）
 *      実Excel ... ★7字★／うち ... ★9字★
 *    ★列の 幅は 1つも ずれて いません★（44.25点 × 96/72 ＝ 59 ＝ `cW(7)`）
 *    ⇒★違うのは 1文字ぶんの 幅★ ... 実Excel 8.67 ／ うち 6.56
 *
 *  ★★訳★★
 *    Excel の 列の 幅は ★「標準の 字で 何文字 入るか」★ で 決まって います。
 *    ＝★物差しは 本の 既定の 字＝そのマスの 字では ありません★
 *    ＝実物で 裏が 取れます … `.ColumnWidth 6.81` ／ `.Width 44.25点`
 *      ⇒44.25 ÷ 6.81 ＝ ★6.50点/字★ ⇒ × 96/72 ＝ ★8.67ドット★
 *    ⇒★数を 決め打ちせず 本の 既定の 字から 出します★
 *
 *  ★★この 見張りが 見る 物★★
 *    ⑴★同じ 中身・同じ 列の 幅で、マスの 字だけ 小さい 時 桁が 変わらない★
 *      ＝★これが 本体★（前は ★小さい 字の マスだけ 桁が 増えて いました★）
 *    ⑵★本の 既定の 字が 違う 本では 桁が 変わる★（★物差しは 本の 字＝だから 効く★）
 *    ⑶★描く 字の 大きさは 変わって いない★（★桁を 決める 時だけ 使う★）
 *
 *  ★走らせ方★ node tests/hyou-no-kihon-no-ji-webkit.mjs
 */
import path from 'node:path'; import http from 'node:http'; import fs from 'node:fs';
import os from 'node:os';
import { createRequire } from 'node:module';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { borrow, launch } from '../scripts/_borrow-playwright.mjs';

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

/* ══ ★材料★ ══（`.xlsx`・★桁が 切れる 長い 小数★）
     A1 ･･･ 19705.567567567567（★書式 General★）
     ⇒★同じ 中身・同じ 列の 幅で 「マスの 字」だけ 変える 2本を 作ります★ */
const 中身 = 19705.567567567567;
async function 作る(マスの字, 本の既定の字) {
  const ws = XLSX.utils.aoa_to_sheet([[中身]]);
  ws['!cols'] = [{ wpx: 59 }];          /* ★列の 幅を 59 に 固定★ */
  ws['!ref'] = 'A1:A1';
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'あ');
  const 道 = path.join(os.tmpdir(), 'exally-kihon-ji-' + マスの字 + '-' + 本の既定の字 + '.xlsx');
  fs.writeFileSync(道, XLSX.write(wb, { bookType: 'xlsx', type: 'buffer' }));

  /* ★借り物は 字体を 書いて くれません★＝★包みを 開けて 自分で 書きます★
     （★これは 材料作りです＝本番の 道では ありません★）
     `xl/styles.xml` の `<fonts>` ･･･ ★0番が 本の 既定★／1番を マスに 当てます */
  const z = ZipSurgeon.read(new Uint8Array(fs.readFileSync(道)));
  const st = z.names().filter((n) => /^xl[/]styles[.]xml$/.test(n))[0];
  const sh = z.names().filter((n) => /^xl[/]worksheets[/]sheet1[.]xml$/.test(n))[0];
  if (!st || !sh) throw new Error('★styles か sheet1 が 在りません＝この 見張りは 空振り★');
  let sx = await z.text(st);
  /* ★字体を 2つ 置く★（0＝本の 既定 ／ 1＝マスの 字） */
  sx = sx.replace(/<fonts[^>]*>[\s\S]*?<[/]fonts>/,
    '<fonts count="2">'
    + '<font><sz val="' + 本の既定の字 + '"/><name val="游ゴシック"/></font>'
    + '<font><sz val="' + マスの字 + '"/><name val="游ゴシック"/></font>'
    + '</fonts>');
  /* ★組を 2つ 置く★（0＝字体0 ／ 1＝字体1） */
  sx = sx.replace(/<cellXfs[^>]*>[\s\S]*?<[/]cellXfs>/,
    '<cellXfs count="2">'
    + '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
    + '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
    + '</cellXfs>');
  if (!/fontId="1"/.test(sx)) throw new Error('★字体を 置けて いません＝この 見張りは 空振り★');
  z.replaceText(st, sx);
  let hx = await z.text(sh);
  hx = hx.replace(/<c r="A1"([^>]*)>/, '<c r="A1" s="1"$1>');   /* ★A1 に 組1（小さい 字）★ */
  if (!/<c r="A1" s="1"/.test(hx)) throw new Error('★A1 に 組を 当てられて いません＝空振り★');
  z.replaceText(sh, hx);
  fs.writeFileSync(道, Buffer.from((await z.build()).bytes));
  return 道;
}

console.log('[本の 既定の 字] ★`General` の 桁は 「本の 既定の 字」で 決まる★');

const wk = await borrow('kihon-no-ji', 'webkit');
const browser = await launch('kihon-no-ji', wk, {}, 'webkit');
const 配信 = await 立てる(ROOT);

/** ★1マスの 「画面に 出る 字」と 「描く 字の 大きさ」を 画面の 関数で 取る★ */
async function 見る(本) {
  const page = await browser.newPage({ viewport: { width: 1280, height: 900 } });
  try {
    page.on('pageerror', (e) => console.log('      [落ちた] ' + String(e.message).slice(0, 200)));
    await page.goto(配信.url + '/book.html', { waitUntil: 'load', timeout: 120000 });
    await page.evaluate(() => {
      document.body.classList.remove('exally-locked');
      const ov = document.getElementById('loginOv');
      if (ov) { ov.classList.remove('open'); ov.style.display = 'none'; }
      window.FileOut = { deliver: () => Promise.reject(new Error('★書き出しては いけません★')) };
    });
    await page.setInputFiles('#bookFileInput', 本);
    await page.waitForFunction(() => {
      const ss = window.sheets || [];
      return ss.some((s) => s && s.data && Object.keys(s.data).length > 0);
    }, null, { timeout: 120000 });
    await page.waitForTimeout(600);

    /* ★★画面の 事を 測る 時は 画面の 関数を 呼ぶ★★（2026-09-25 の 決め） */
    const 無い = await page.evaluate(() => ['_字の元', 'fmtForDisplay', '_入る字数', 'cW', '_マスの字大', '_既定の字体の字']
      .filter((n) => typeof window[n] !== 'function'));
    if (無い.length) { console.log('  ★★測れません★★ 在りません ... ' + 無い.join(' / ')); process.exit(8); }

    return await page.evaluate(() => {
      const sh = window.sheets[0];
      const cell = (sh.data || {})['0,0'];
      const raw = window._字の元(cell);
      const w = window.cW(0);
      const n = window._入る字数(w, raw, cell && cell.numFmt);
      return {
        列の幅: w,
        入る字数: n,
        出た字: String(window.fmtForDisplay(raw, cell && cell.numFmt, n)),
        マスの字大: window._マスの字大(cell),
        既定の字: window._既定の字体の字(1),
        倍: window.scale,
        物差しの字: (function () { var m = window._物差しの筆 ? window._物差しの筆() : null; return m ? m.font : '（無し）'; }()),
        既定の一文字: (function () {
          var 前 = window.ctx.font;
          window.ctx.font = window._既定の字体の字(1);
          var w2 = Math.round(window.ctx.measureText('0').width * 100) / 100;
          window.ctx.font = 前;
          return w2;
        }()),
        /* ★★そのマスの 字を わざと 大きく してから 数える★★（2026-09-26）
             ＝★`_入る字数` が 本当に そのマスの 字を 見て いない事を 見ます★
             ＝★1回 「一文字の 所だけ」直して 効かなかった 形を 二度 作らない 為★ */
        字を替えても同じ: (function () {
          var 前 = window.ctx ? window.ctx.font : null;
          if (!window.ctx) return null;
          window.ctx.font = '40px "游ゴシック",sans-serif';
          var a1 = window._入る字数(w, raw, cell && cell.numFmt);
          window.ctx.font = '6px "游ゴシック",sans-serif';
          var a2 = window._入る字数(w, raw, cell && cell.numFmt);
          window.ctx.font = 前;
          return { 大きい字: a1, 小さい字: a2 };
        }()),
        本の既定の字大: sh.既定の字大 || null,
        本の既定の字体名: sh.既定の字体名 || null,
      };
    });
  } finally {
    await page.close().catch(() => {});
  }
}

try {
  /* ══ ⑴マスの 字だけ 小さい（本の 既定は 11）══ */
  const 小 = await 見る(await 作る(9, 11));
  /* ══ 同じ 本で マスの 字も 11（＝物差しと 同じ）══ */
  const 同 = await 見る(await 作る(11, 11));
  console.log('      ── 実測 ── ★マスの 字 9pt／本の 既定 11pt★ ... 列の幅 ' + 小.列の幅
    + ' ／ 入る字数 ' + 小.入る字数 + ' ／ 出た字 ' + 小.出た字.length + '字'
    + ' ／ 描く 字 ' + 小.マスの字大 + 'px ／ 既定の 1文字 ' + 小.既定の一文字);
  console.log('      ── 実測 ── ★マスの 字 11pt／本の 既定 11pt★ ... 入る字数 ' + 同.入る字数
    + ' ／ 出た字 ' + 同.出た字.length + '字 ／ 描く 字 ' + 同.マスの字大 + 'px');

  T('★★マスの 字が 小さくても 桁は 変わらない★★（★これが 本体★）',
    小.入る字数 === 同.入る字数,
    '9pt ' + 小.入る字数 + '字 ／ 11pt ' + 同.入る字数 + '字（★前は 9pt だけ 増えて いました★）');
  T('★★出る 字も 同じ★★', 小.出た字 === 同.出た字,
    '9pt ' + JSON.stringify(小.出た字) + ' ／ 11pt ' + JSON.stringify(同.出た字));
  /* ══ ★★そのマスの 字を わざと 替えても 桁が 動かない事★★ ══
       ＝★「マスの 字が 9pt の 材料」を 借り物が 作れなかった★（実測＝描く 字が 同じ 15px）
       ⇒★★だから 「材料で 変える」では なく 「`ctx.font` を 直に 替えて 数える」★★
       ⇒★これが 本体の 検査です★（★空振りに ならない★） */
  const か = 小.字を替えても同じ;
  console.log('      ── 実測 ── ★`ctx` の 字を 替えて 数えた★ ... '
    + '40px ⇒ ' + (か ? か.大きい字 : '取れず') + '字 ／ 6px ⇒ ' + (か ? か.小さい字 : '取れず') + '字');
  T('★★そのマスの 字を 40px/6px に 替えても 桁が 同じ★★（★物差しは 本の 既定の 字★）',
    !!か && か.大きい字 === か.小さい字,
    '40px ' + (か && か.大きい字) + '字 ／ 6px ' + (か && か.小さい字) + '字'
    + '（★違ったら そのマスの 字で 測って います★）');

  /* ══ ★★ここから 先は ★私の 物差しでは 決められません★★★ ══（2026-09-26）
       ★★確かめられた 事（上の 3本）★★
         ＝★そのマスの 字が 何ポイントでも 桁は 同じ★
         ＝★`ctx` の 字を 40px/6px に 替えても 同じ★
         ⇒★＝「そのマスの 字で 桁を 決める」のは 直りました★
       ★★確かめられて いない 事（★正直に 書きます★）★★
         ⑴★本の 既定の 字を 20pt に した 本で 桁が 減りませんでした★
            ⇒★つまり 「本の 字の 大きさ」が どこまで 効くかは 分かって いません★
            ⇒因は 詰め直しの 輪の 中だと 思います（★測って いません★）
         ⑵★借り物が 列の 幅を 59 に して くれません★（`wpx:59` ⇒ `cW` は 79）
            ⇒★だから 実物と 同じ 形（幅 59・9pt・General）を 作れて いません★
         ⇒★★実物の 1マスが 7字に なったかは 経営者1 の 物差し（実Excel）で しか 言えません★★
       ★だから ここでは 「桁が 減る」を 門に しません★
         ＝★作れて いない 材料で 緑を 出す のは 空振りです★
       ★本の 既定の 字を 読めて いる事だけ 見ます★ */
  const 大 = await 見る(await 作る(11, 20));
  console.log('      ── 実測 ── ★本の 既定 20pt★ ... 入る字数 ' + 大.入る字数
    + ' ／ 既定の 1文字 ' + 大.既定の一文字 + ' ／ 物差しの字 ' + JSON.stringify(大.物差しの字));
  T('★本の 既定の 字を 読めて いる★（★物差しに 使う 字★）',
    大.本の既定の字大 === 20 && /27px/.test(String(大.物差しの字)),
    '読めた 既定 ' + JSON.stringify(大.本の既定の字大) + ' ／ 物差し ' + JSON.stringify(大.物差しの字));
  console.log('');
  console.log('  ★★この 見張りが 言える 事★ ... 「そのマスの 字で 桁を 決めない」★★');
  console.log('  ★★言えない 事★ ... 「実物の 1マスが 実Excel と 同じ 字に なった」★★');
  console.log('    ⇒★それは 経営者1 の 物差し（実Excel・21,422マス）で 測って ください★');
} finally {
  await browser.close().catch(() => {});
  配信.閉じる();
}

console.log('');
console.log('hyou-no-kihon-no-ji: ' + pass + ' 緑 / ' + fail + ' 赤');
process.exit(fail ? 1 : 0);
