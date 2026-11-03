/* hakaru-kazari-ga-gamen-made-todoku-ka.mjs
 *   ★実Excel が 作った 飾り付きの ファイルを ★お客さんの 道★で 開いて
 *     ★どの 飾りが 画面まで 届くか★ を 数える★（2026-09-21）
 *
 *  ★★なぜ 要るか★★
 *    司さん「★全く 同じように 表示できるように しろや★
 *            ★何の ために 細胞レベルまで 分析したんど★」（イ）
 *
 *    今 在る 紙 `golden-uketotta-kazari3-excel-2026-09-20.tsv` が 見て いるのは
 *    ★★ファイルの 中に 残って いるか★★ だけです（実Excel に 開かせて 読んだ）。
 *    ⇒★画面に 出て いるかは 1つも 測って いません★
 *    ⇒★「消えて いない」と「見えて いる」は 別です★
 *
 *  ★★これは 見張りでは ありません＝測り道具です★★
 *    ★材料が repo の 外に 在る★（`%TEMP%` の xlsx）ので 見張りに できません。
 *    ⇒★数えて 紙に する のが 仕事★。見張りは その 後です。
 *
 *  ★★材料★★（★実Excel が COM で 作った 物★）
 *    既定 ... `%TEMP%\exally-tameshi-hiraku3.xlsx`
 *    作り方 ... `docs/measured/tsukuru-tameshi-hiraku3-kazari-to-kobore3.ps1`
 *              （★この 道具は COM を 使います＝ここでは 走らせません★）
 *    ★手元の 決め打ちの 道を 焼き込みません★＝`--もと <path>` で 変えられます
 *
 *  ★★材料に 何が 入って いるか★★（上の ps1 の 80〜104行目 を そのまま 写した）
 *    罫線   ... `A1:C3` の 周り（LineStyle 1 / Weight 3）＋ `B2` の 下（LineStyle 1 / Weight 4）
 *    塗り   ... `B1` ＝ 65535（COM の 数＝BGR）＝★黄色★
 *    字     ... `A1` 太字 ／ `A1` 色 255（COM の 数＝BGR）＝★赤★
 *    表示の形 ... `B1` ＝ `#,##0` ／ `A3` ＝ `0.0%`
 *    繋げたマス ... `A5:C5`
 *    図形   ... 1つ（四角・左320 上20 幅60 高60・ポイント）
 *    列の幅 ... A列 ＝ 30
 *    溢れ   ... `C1`縦3 ／ `E1`横3 ／ `A10`2次元 2x3
 *
 *  ★★お客さんの 道で 開きます★★
 *    `#bookFileInput` に ファイルを 入れます（★画面の ボタンと 同じ 入口★）
 *    ＝`onPickBookFile()` が 動きます。
 *    ★JS で 中を 直に 作りません★（★それは お客さんの 道では ありません★）
 *
 *  ★★出す 物★★
 *    ①飾りごとに ★台（`sheets[].data`）に 届いたか★
 *    ②★絵★（png）＝★数が 緑でも 絵を 見るまで OK に しない★
 *
 *  使い方:
 *    node docs/measured/hakaru-kazari-ga-gamen-made-todoku-ka.mjs
 *    node docs/measured/hakaru-kazari-ga-gamen-made-todoku-ka.mjs --もと <xlsx への 道>
 */
import path from 'node:path';
import http from 'node:http';
import fs from 'node:fs';
import crypto from 'node:crypto';
import { fileURLToPath } from 'node:url';
import { borrow, launch } from '../../scripts/_borrow-playwright.mjs';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..', '..');
const TEMP = process.env.TEMP || process.env.TMP || '.';

/* ★手元の 道を 焼き込みません★＝渡されたら そちらを 使います */
const i = process.argv.indexOf('--もと');
const 材料 = (i >= 0 && process.argv[i + 1])
  ? process.argv[i + 1]
  : path.join(TEMP, 'exally-tameshi-hiraku3.xlsx');

function 立てる(root) {
  const 型 = {
    '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8',
    '.css': 'text/css; charset=utf-8', '.json': 'application/json; charset=utf-8',
    '.png': 'image/png', '.svg': 'image/svg+xml', '.ico': 'image/x-icon',
  };
  const s = http.createServer((q, r) => {
    const f = path.join(root, decodeURIComponent(String(q.url).split('?')[0]).replace(/^\/+/, ''));
    if (!f.startsWith(root) || !fs.existsSync(f) || fs.statSync(f).isDirectory()) {
      r.statusCode = 404; return r.end('no');
    }
    r.setHeader('content-type', 型[path.extname(f).toLowerCase()] || 'application/octet-stream');
    fs.createReadStream(f).pipe(r);
  });
  return new Promise((x) => s.listen(0, '127.0.0.1',
    () => x({ url: 'http://127.0.0.1:' + s.address().port, 閉じる: () => s.close() })));
}

console.log('');
console.log('[kazari-gamen] ★飾りが ★画面まで★ 届くか★');
console.log('      ＝ 材料 ' + 材料);

if (!fs.existsSync(材料)) {
  console.log('  ★材料が 有りません★');
  console.log('    ⇒`docs/measured/tsukuru-tameshi-hiraku3-kazari-to-kobore3.ps1` で 作れます');
  console.log('      （★COM を 使います＝実Excel が 1つも 動いて いない 時だけ★）');
  console.log('    ⇒★0件を 緑に しません★');
  process.exit(2);
}
const 中 = fs.readFileSync(材料);
console.log('      ＝ ' + 中.length + 'B ／ sha256 ' + crypto.createHash('sha256').update(中).digest('hex'));

const wk = await borrow('kazari-gamen', 'webkit');
const browser = await launch('kazari-gamen', wk, {}, 'webkit');
const page = await browser.newPage({ viewport: { width: 1280, height: 800 } });
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

  /* ★★お客さんの 入口★★＝画面の 「読み込む」ボタンが 押す 物と 同じ */
  const ある = await page.evaluate(() => !!document.getElementById('bookFileInput')
    && typeof window.onPickBookFile === 'function');
  if (!ある) throw new Error('★お客さんの 入口（#bookFileInput / onPickBookFile）が 有りません★');
  console.log('  ok   ★お客さんの 入口が 在る★');

  await page.setInputFiles('#bookFileInput', 材料);
  /* ★読み終わるまで 待ちます★（SheetJS は 押してから 読み込まれる＝時間が 要る） */
  await page.waitForFunction(() => {
    const sh = window.sheets && window.sheets[window.activeSheet];
    return !!(sh && sh.data && Object.keys(sh.data).length > 3);
  }, null, { timeout: 60000 }).catch(() => {});

  const 出 = await page.evaluate(() => {
    const sh = (window.sheets || [])[window.activeSheet || 0] || {};
    const d = sh.data || {};
    const 読 = (r, c) => d[r + ',' + c] || null;
    return {
      マスの数: Object.keys(d).length,
      A1: 読(0, 0), B1: 読(0, 1), B2: 読(1, 1), A3: 読(2, 0), A5: 読(4, 0),
      C1: 読(0, 2), E1: 読(0, 4), A10: 読(9, 0),
      colW: sh.colW || null,
      rowH: sh.rowH || null,
      merges: sh.merges || sh.繋げたマス || null,
      /* ★★置いた 図形は ★シートが★ 持って います★★（`sheets[i].objects`）
           `book.html:14692 物の箱()` ＝ `s0.objects`
         ★★2026-09-21 ここで 1回 踏みました★★
           私は `SheetObjects.形たち` を 数えて 「図形が 1つ 届いた」と 出しました。
           ⇒`形たち` は ★Exally が 描ける 形の 種類の 名簿★（四角・角丸四角四角・角丸四角 ほか7種）
           ⇒★★開いた ファイルの 図形では ありません★★
           ⇒★絵を 見たら 判子が 1つも 無く、そこで 気づきました★
           ⇒★数は 正しく 測れて いました＝測る 先が 違った★ */
      図形: (function () {
        const sh0 = (window.sheets || [])[window.activeSheet || 0] || {};
        const 箱 = sh0.objects;
        return {
          置いた数: Array.isArray(箱) ? 箱.length : '(objects が 無い)',
          中身: Array.isArray(箱) ? 箱.slice(0, 2) : null,
          描ける種類: (window.SheetObjects && Array.isArray(window.SheetObjects.形たち))
            ? window.SheetObjects.形たち.length : '(分からない)',
        };
      }()),
      シートの鍵: Object.keys(sh),
      /* ★★板が 2枚 以上の 時に 「どの 板に 出たか」を 見ます★★（2026-09-21）
           ＝`.xlsb` の 板の 名前 ⇒ 部品名 を 引き違えると
             ★判子が 別の 板に 出ます★（★消えるのでは なく ずれる★） */
      板たち: (window.sheets || []).map(function (x) {
        return { 名: x.name,
          A1: (x.data && x.data['0,0']) ? String(x.data['0,0'].d) : '',
          図形: (x.objects || []).map(function (o) { return o.名 + '/' + (o.種類 || '?'); }) };
      }),
      当て方: (function () {
        try {
          var c = (window.BookOpen && window.BookOpen.current) ? window.BookOpen.current() : null;
          return c ? (c.図形の当て方 || '(空)') : '(開いて いない)';
        } catch (e) { return '(取れない)'; }
      }()),
    };
  });

  const ま = (x) => (x ? JSON.stringify(x) : '(空)');
  console.log('');
  console.log('  ★台（`sheets[].data`）に 何が 届いたか★  マス ' + 出.マスの数 + '個');
  console.log('    A1  (太字・赤・罫線) = ' + ま(出.A1));
  console.log('    B1  (黄・#,##0)      = ' + ま(出.B1));
  console.log('    B2  (下の 罫線 太4)  = ' + ま(出.B2));
  console.log('    A3  (0.0%)           = ' + ま(出.A3));
  console.log('    A5  (A5:C5 を 繋げた)= ' + ま(出.A5));
  console.log('    C1  (溢れ 縦3)       = ' + ま(出.C1));
  console.log('    E1  (溢れ 横3)       = ' + ま(出.E1));
  console.log('    A10 (溢れ 2次元)     = ' + ま(出.A10));
  console.log('    列の幅  = ' + ま(出.colW));
  console.log('    行の高さ= ' + ま(出.rowH));
  console.log('    繋げた  = ' + ま(出.merges));
  console.log('    図形    = ' + ま(出.図形));
  console.log('    シートの 鍵 = ' + (出.シートの鍵 || []).join(' '));
  console.log('    ★板の 当て方★ = ' + 出.当て方);
  (出.板たち || []).forEach(function (x, i) {
    console.log('    板' + (i + 1) + ' 名=' + x.名 + ' A1=' + x.A1
      + ' 図形=' + (x.図形.length ? x.図形.join(',') : '(無し)'));
  });

  /* ★★飾りごとに 届いたか★★（★1つずつ 名指し★） */
  const A1 = 出.A1 || {}, B1 = 出.B1 || {}, B2 = 出.B2 || {}, A3 = 出.A3 || {};
  const 判じ = [
    ['字（A1）太字',    !!A1.bold],
    ['字（A1）色 赤',   !!A1.color],
    ['罫線（A1 の 周り）', !!A1.border],
    ['罫線（B2 の 下・太さ4）', !!B2.border],
    ['塗り（B1）黄',    !!B1.bgColor],
    ['表示の形（B1）#,##0', !!B1.numFmt],
    ['表示の形（A3）0.0%',  !!A3.numFmt],
    /* ★★繋げたマスは シートでは なく ★マスが★ 持って います★★（`mergeEnd`）
         ＝2026-09-21 に ★私が `sh.merges` を 見て 「届かない」と 出しました★
         ＝★探す 所を 間違えると 本番の 穴に 見えます★ */
    ['繋げたマス（A5:C5）', !!(出.A5 && 出.A5.mergeEnd)],
    ['列の幅（A列＝30）',   !!(出.colW && Object.keys(出.colW).length)],
    ['図形（判子）1つ',     !!(出.図形 && typeof 出.図形.置いた数 === 'number' && 出.図形.置いた数 > 0)],
    ['溢れ 縦（C1:C3）',    !!(出.C1 && 出.C1.f)],
    ['溢れ 横（E1:G1）',    !!(出.E1 && 出.E1.f)],
    ['溢れ 2次元（A10）',   !!(出.A10 && 出.A10.f)],
  ];
  let 届 = 0;
  console.log('');
  console.log('  ★★飾りが 台に 届いたか★★');
  for (const [な, ok] of 判じ) {
    if (ok) 届++;
    console.log('    ' + (ok ? '★届いた★  ' : '★届かない★') + ' ' + な);
  }
  console.log('');
  console.log('  ＝ ★届いた ' + 届 + ' / ' + 判じ.length + '★');
  console.log('  ★★但し これは 「台に 在るか」です★★');
  console.log('    ＝★描いて いるか は 下の 絵を 見て ください★');

  /* ══ ★★画素の 色を 測る★★ ══（2026-09-21）
       ★「印が 付いたか」では なく 「絵が 変わったか」★（記憶の 決まり）
       ⇒キャンバスから ★そのマスの 真ん中の 色★ を 読みます。
       ★待ってから 撮ります★＝知らせの 札が 消えるのを 待つ（札は 絵に 写ります） */
  await page.waitForFunction(() => {
    const t = document.querySelectorAll('.toast, #toast');
    for (const x of t) { if (x.offsetParent !== null) return false; }
    return true;
  }, null, { timeout: 15000 }).catch(() => {});

  const 画素 = await page.evaluate(() => {
    const cv = document.getElementById('grid-canvas');
    if (!cv) return '(キャンバスが 無い)';
    const g = cv.getContext('2d');
    /* ★画面の 点と キャンバスの 点は 倍率が 違います★（Retina 等）
         ★clientWidth が 0 の 時は 測れません★＝★1 と 決めつけません★ */
    if (!cv.clientWidth || !cv.width) return '(キャンバスの 大きさが 取れない) w=' + cv.width + ' cw=' + cv.clientWidth;
    const 倍 = cv.width / cv.clientWidth;
    const 読 = (r, c) => {
      if (typeof window.colX !== 'function') return '(colX が 無い)';
      const x = window.colX(c) + window.cW(c) / 2;
      const y = window.rowY(r) + window.rH(r) / 2;
      const d = g.getImageData(Math.round(x * 倍), Math.round(y * 倍), 1, 1).data;
      const h = (n) => ('0' + n.toString(16).toUpperCase()).slice(-2);
      return '#' + h(d[0]) + h(d[1]) + h(d[2]);
    };
    return { A1: 読(0, 0), B1: 読(0, 1), C1: 読(0, 2), A4: 読(3, 0), D1: 読(0, 3) };
  });
  console.log('');
  console.log('  ★★画素の 色★★（★マスの 真ん中★）');
  console.log('    B1（黄の はず #FFFF00） = ' + (画素 && 画素.B1));
  console.log('    A1（塗り 無し）          = ' + (画素 && 画素.A1));
  console.log('    C1（塗り 無し）          = ' + (画素 && 画素.C1));
  console.log('    D1（塗り 無し・対照）    = ' + (画素 && 画素.D1));
  console.log('    A4（飾り 無し・対照）    = ' + (画素 && 画素.A4));

  const 絵 = path.join(TEMP, 'exally-kazari-gamen.png');
  await page.screenshot({ path: 絵, fullPage: false });
  const 絵中 = fs.readFileSync(絵);
  console.log('');
  console.log('  ★絵★ ' + 絵);
  console.log('    ' + 絵中.length + 'B ／ sha256 ' + crypto.createHash('sha256').update(絵中).digest('hex'));
} catch (e) {
  console.log('  ★止まりました★ ' + String(e && e.message).slice(0, 200));
  process.exitCode = 1;
} finally {
  if (配信) 配信.閉じる();
  await browser.close();
}
