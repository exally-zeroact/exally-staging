/* kazoeru-hiraita-ato-keisan-ga-iruka.mjs
 *   -- ★「開いた 直後に 計算が 要るか」を ★場合ごとに 数える★★（117）（2026-09-25）
 *
 *  ★★なぜ★★
 *    司さん 09-25「必要ならいるし必要ないならいらんやろが／どんな状況でいるかいらんか言えや」
 *    ＝★「どっちにしますか」と 訊いたのが 間違い★。★場合を 挙げて、この 本が どれかを 数える★
 *
 *  ★★場合★★
 *    ★計算が いる★
 *      ㋐★開いた 日で 答えが 変わる 関数★（TODAY NOW RAND RANDBETWEEN OFFSET INDIRECT CELL INFO ...）
 *      ㋑★式は 在るのに 答えが 保存されて いない マス★
 *      ㋒★Excel 以外の 道具が 作った 本★（答えが 入って いない／古い 事が 在る）
 *      ㋓★「開いたら 全部 計算しろ」の 印が 本に 立って いる★（`calcPr` の `fullCalcOnLoad`）
 *      ㋔★お客さんが 1マス 打った 後★ ... ★これは 「開いた 直後」では ない＝この 道具は 見ない★
 *    ★計算が いらん★
 *      ㋐〜㋓ が ★1つも 無い★＝★計算しても 保存されて いる 答えと 同じ 物が 出るだけ★
 *
 *  ★★一番 大事な 物★★
 *    ★Exally が 計算した 答えと 本に 保存されて いる 答えが ★何個 違うか★★
 *      ・★0個★なら ⇒ この 本で 計算は ★1つも 新しい 物を 生んで いない★
 *      ・★違う★なら ⇒ ★保存が 古い★か ★Exally が 間違って いる★の どちらか
 *                    ＝★どちらでも 「計算を やめる」判断の 前に 割らないと いけない★
 *
 *  ★★出さない 物★★ ... ★マスの 値／シートの 名／式の 字★（★数と 関数名だけ★）
 *
 *  ★★1バイトも 書きません★★（`FileOut.deliver` は 拒む 代物に 差し替え）
 *
 *  使い方: node docs/measured/kazoeru-hiraita-ato-keisan-ga-iruka.mjs "<xlsb|xlsx>" [--秒 300]
 */
import fs from 'node:fs';
import path from 'node:path';
import http from 'node:http';
import crypto from 'node:crypto';
import { fileURLToPath } from 'node:url';
import { borrow, launch } from '../../scripts/_borrow-playwright.mjs';

const ここ = path.dirname(fileURLToPath(import.meta.url));
/* ★どの 木を 測るか★（`--元 <置き場>`／指さない 時は 道具の 木）
     ＝★道具は 私の 木・測りたい 字は 相手の 木★ が 普通です（114・115 と 同じ 口） */
const ROOT = process.argv.includes('--元')
  ? path.resolve(process.argv[process.argv.indexOf('--元') + 1])
  : path.join(ここ, '..', '..');
const 材料 = process.argv[2];
const 秒 = process.argv.includes('--秒') ? Number(process.argv[process.argv.indexOf('--秒') + 1]) : 300;
/* ★★物差しを 実Excel に 取り替える★★（2026-09-25・追記）
     ★なぜ★
       09-25 に 私は ★借り物（SheetJS）の `w`★ を 物差しに して ★3,403個 違う★と 出しました。
       ⇒★但し 借り物は 曜日 `aaa` を 出せて いませんでした★＝★Exally の 方が 合って いた★
       ⇒★物差しが 一部 壊れて いる 数でした★
     ⇒★実Excel が 出した 字（道具 120 の TSV）を 物差しに します★
     ⇒★実Excel が 正★なので「Exally の 方が 正しい」は 原理上 起きません
     ★TSV の 形★ ... `板の名<TAB>行<TAB>列<TAB>出た字`（行・列は 0から） */
const 実道 = process.argv.includes('--実Excel')
  ? process.argv[process.argv.indexOf('--実Excel') + 1] : null;
/* ★★`--例 N` ＝ 違った マスの ★居場所★を N個 出す★★（2026-09-25・既定は 0）
     ★なぜ 要るか★ ... 作り手が ★その マスを 名指しで 開く★ 為
     ★出すのは 居場所だけ★ ... 板の 名 ＋ `行,列`（0から）★／★値は 1つも 出しません★
     ★★これを 付けた 回の 出しは repo に 入れません★★（★板の 名が 入ります★）
     ⇒既定 0 ＝★紙に 残す 回は 居場所も 出ません★ */
const 例数 = process.argv.includes('--例')
  ? Number(process.argv[process.argv.indexOf('--例') + 1]) : 0;
if (!材料 || !fs.existsSync(材料)) { console.log('★材料が 在りません★'); process.exit(3); }
const 中 = fs.readFileSync(材料);
console.log('★材料★ ' + path.basename(材料) + '（' + 中.length.toLocaleString() + ' バイト）');
console.log('  sha256 ' + crypto.createHash('sha256').update(中).digest('hex'));
console.log('★★測る 木★★ ' + ROOT);
try {
  const { execSync } = await import('node:child_process');
  const h = execSync('git -C "' + ROOT + '" rev-parse --short HEAD', { encoding: 'utf8' }).trim();
  const d = execSync('git -C "' + ROOT + '" status --porcelain', { encoding: 'utf8' }).trim();
  console.log('  ★印★ ' + h + (d ? '（★手元に 未commit が 在ります★）' : '（手元は 綺麗）'));
} catch (e) { console.log('  ★印を 読めません ... ' + e.message + '★'); }

/* ★★開いた 日で 答えが 変わる 関数★★
     ★なぜ 名簿を 手で 持つか★
       ＝★台に 「揮発するか」を 聞く 口が 無い★（借り物の 中）
       ⇒★名簿は ★実Excel の 決まり★を 写した 物＝★私の 思いつきでは ない★
       ⇒★足りない かも しれない★＝★「0個」を 「絶対に 無い」と 読んでは いけない★ */
const 日で変わる = [
  'TODAY', 'NOW', 'RAND', 'RANDBETWEEN', 'RANDARRAY',
  'OFFSET', 'INDIRECT', 'CELL', 'INFO',
];

/* ══ ①本を 直に 読む（★計算を 一切 させずに★） ══ */
const XLSX = (await import(path.join(ROOT, 'lib', 'xlsx.full.min.js').replace(/\\/g, '/')).catch(() => null))
  || (await import('xlsx').catch(() => null));
let 台 = XLSX && (XLSX.default || XLSX);
if (!台 || !台.read) {
  /* ★台は 画面向けの 字＝node で import 出来ない 事が 在る★
     ⇒★その 時は 字を 読んで その場で 動かす★（★repo は 触らない★） */
  const 字 = fs.readFileSync(path.join(ROOT, 'lib', 'xlsx.full.min.js'), 'utf8');
  const g = {};
  // eslint-disable-next-line no-new-func
  new Function('global', 'self', 'window', 字).call(g, g, g, g);
  台 = g.XLSX;
}
if (!台 || !台.read) { console.log('★★台を 読めません★★'); process.exit(4); }

const 本 = 台.read(中, { type: 'buffer', cellFormula: true, cellStyles: false, bookVBA: false });
const 保存 = {};          /* 板|番地 → 保存されて いた 値（★生の 数★） */
const 保存字 = {};        /* 板|番地 → ★Excel が 表示して いた 字★（SheetJS の `w`） */
let 式の数 = 0, 値なし式 = 0, マスの数 = 0;
const 使われた日で変わる = {};
(本.SheetNames || []).forEach((名) => {
  const sh = 本.Sheets[名] || {};
  Object.keys(sh).forEach((番) => {
    if (番.charAt(0) === '!') return;
    const c = sh[番];
    マスの数++;
    if (c && c.f) {
      式の数++;
      const 大 = String(c.f).toUpperCase();
      日で変わる.forEach((k) => {
        if (new RegExp('(^|[^A-Z0-9_.])' + k + '\\s*\\(').test(大)) {
          使われた日で変わる[k] = (使われた日で変わる[k] || 0) + 1;
        }
      });
      if (c.v === undefined || c.v === null) 値なし式++;
      /* ★★鍵は 画面の 側と 同じ 形に する★★（2026-09-25・★1回 嘘を 出した★）
           ・本の 側 ... `A1`
           ・画面の 側 ... `行,列`（0から）＝`js/book-open.js:615  data[rc.r + ',' + rc.c] = cell`
           ⇒★揃えないと 1つも 噛み合わず 「違い 0個」が 出る★ */
      else {
        const rc = 台.utils.decode_cell(番);
        const 鍵 = 名 + '|' + rc.r + ',' + rc.c;
        保存[鍵] = c.v;
        /* ★★Excel 自身が 表示して いた 字★★（SheetJS の `w`）
             ★なぜ 要るか★
               画面の 側（`d`）は ★書式を 掛けた 後の 字★です。
               生の 数（`v`）と 比べると ★1,053個が 「日付らしい 字」★ で 止まり
               ★合って いるのか いないのか 分かりません★（09-25 実際に そう なった）
             ⇒★字 vs 字★で 比べます＝[[feedback_naka_no_kazu_ga_onaji_wa_onaji_de_nai]]
             ⇒★`w` が 無い マスも 在ります★＝★その 分は 分母から 外して 数えます★ */
        if (c.w !== undefined && c.w !== null) 保存字[鍵] = String(c.w);
      }
    }
  });
});
console.log('');
console.log('★★①本を 直に 読んだ 数（★計算させて いません★）★★');
console.log('  板 ' + (本.SheetNames || []).length + '枚 ／ マス ' + マスの数.toLocaleString()
  + '個 ／ ★式 ' + 式の数.toLocaleString() + '個★');
console.log('  ㋑★式は 在るのに 答えが 保存されて いない マス★ ... ★' + 値なし式.toLocaleString() + '個★');

console.log('');
console.log('★★㋐開いた 日で 答えが 変わる 関数★★（★名簿 ' + 日で変わる.length + '個を 探しました★）');
const 日鍵 = Object.keys(使われた日で変わる);
if (!日鍵.length) console.log('  ★0個★（★但し 名簿に 無い 物は 見つかりません★）');
else 日鍵.sort((a, b) => 使われた日で変わる[b] - 使われた日で変わる[a])
  .forEach((k) => console.log('  ★' + k.padEnd(12) + String(使われた日で変わる[k]).padStart(6) + '個★'));

/* ══ ★実Excel の 字を 読む★ ══ */
const 実字 = {};
let 実行数 = 0;
if (実道) {
  if (!fs.existsSync(実道)) { console.log('★★--実Excel の ファイルが 在りません★★'); process.exit(7); }
  const 生 = fs.readFileSync(実道, 'utf8');
  生.split('\n').forEach((行) => {
    if (!行) return;
    const p2 = 行.split('\t');
    if (p2.length !== 4) return;
    実字[p2[0] + '|' + p2[1] + ',' + p2[2]] = p2[3];
    実行数++;
  });
  console.log('');
  console.log('★★実Excel の 字★★ ' + 実行数.toLocaleString() + ' マス（' + path.basename(実道) + '）');
  if (実行数 === 0) { console.log('★★1行も 読めません★★'); process.exit(7); }
}

/* ══ ②「開いたら 全部 計算しろ」の 印 ══ */
console.log('');
console.log('★★㋓「開いたら 全部 計算しろ」の 印★★');
let 印 = '★見つかりません★';
try {
  const z = 台.read(中, { type: 'buffer', bookFiles: true });
  const 名々 = Object.keys(z.files || {});
  const w = 名々.filter((n) => /workbook\.(xml|bin)$/i.test(n));
  console.log('  本の 台帳 ... ' + (w.length ? w.join(' / ') : '★無し★'));
  const x = 名々.filter((n) => /workbook\.xml$/i.test(n));
  if (x.length) {
    const s = String(z.files[x[0]].content || '');
    const m = /<calcPr[^>]*>/i.exec(s);
    印 = m ? m[0] : '★`calcPr` が 在りません＝印は 立って いません★';
  } else {
    印 = '★`.bin` の 台帳＝この 道具では 読めません（未測定）★';
  }
} catch (e) { 印 = '★読めません ... ' + e.message + '★'; }
console.log('  ' + 印);

/* ══ ③Exally に 計算させて 突き合わせる ══ */
const 型 = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8', '.json': 'application/json', '.png': 'image/png',
  '.svg': 'image/svg+xml', '.ico': 'image/x-icon' };
const server = http.createServer((req, res) => {
  const u = decodeURIComponent(req.url.split('?')[0]);
  const p = path.join(ROOT, u === '/' ? 'book.html' : u.replace(/^\//, ''));
  if (!p.startsWith(ROOT) || !fs.existsSync(p) || fs.statSync(p).isDirectory()) { res.writeHead(404); res.end(); return; }
  res.writeHead(200, { 'Content-Type': 型[path.extname(p)] || 'application/octet-stream' });
  res.end(fs.readFileSync(p));
});
await new Promise((ok) => server.listen(0, '127.0.0.1', ok));
const もと = 'http://127.0.0.1:' + server.address().port;

const ck = await borrow('keisan-iruka', 'chromium');
const browser = await launch('keisan-iruka', ck, {}, 'chromium');
const page = await browser.newPage({ viewport: { width: 1280, height: 900 } });
await page.goto(もと + '/book.html', { waitUntil: 'load', timeout: 120000 });
await page.evaluate(() => {
  document.body.classList.remove('exally-locked');
  const ov = document.getElementById('loginOv');
  if (ov) { ov.classList.remove('open'); ov.style.display = 'none'; }
  window.FileOut = { deliver: () => Promise.reject(new Error('★書き出しては いけません★')) };
});
console.log('');
console.log('★★③Exally に 計算させて います★★（最大 ' + 秒 + '秒）');
const t0 = Date.now();
await page.setInputFiles('#bookFileInput', 材料);
let 開いた = 0;
for (let t = 0; t < 秒; t += 2) {
  await new Promise((r) => setTimeout(r, 2000));
  const い = await page.evaluate(() => {
    const ss = window.sheets || [];
    let m = 0; ss.forEach((sh) => { m += Object.keys((sh && sh.data) || {}).length; });
    return m;
  }).catch(() => -1);
  if (い > 0) { 開いた = Date.now() - t0; break; }
}
console.log('  開くまで ' + (開いた ? 開いた.toLocaleString() + ' ms' : '★開きません★'));
if (!開いた) { await browser.close(); server.close(); console.log('★★開かないので 突き合わせ 出来ません★★'); process.exit(5); }

/* ★★板は ★お客さんと 同じ やり方で 開きます★★★（2026-09-25・★3回目の 直し★）
     ★1回目★ ... `cell.d` を 読んだ ⇒★描く 所の 直しが 見えない★
     ★2回目★ ... `window.activeSheet = i` を 入れた ⇒★足りない★
     ★因（`book.html` の `switchSheet`）★
         activeSheet = idx
         ★loadSheetIntoEngine(idx)★   ... 板を 台に 流す
         ★recalcSheet(idx, ...)★       ... 計算し直す
         updateBar(); render()
       ⇒★代入だけでは 「板を 開いた」事に なりません★
       ⇒★`cW(列)` も 開いて いない 板の 幅を 見ます★（`General` は 幅で 桁が 変わる）
     ⇒★`switchSheet(i)` を 呼びます★＝★お客さんが 板の 札を 押すのと 同じ★
     ⇒★1板ずつ 別の 呼びに します★（1回で 15板 回すと 待ちきれません） */
const 板の数 = await page.evaluate(() => (window.sheets || []).length);
const 道しらべ = await page.evaluate(() => {
  const 要 = [['_字の元', window._字の元], ['_ゼロを隠すか', window._ゼロを隠すか],
    ['_答えは字か', window._答えは字か], ['fmtForDisplay', window.fmtForDisplay],
    ['_入る字数', window._入る字数], ['cW', window.cW], ['switchSheet', window.switchSheet]];
  const 欠 = 要.filter((x) => typeof x[1] !== 'function').map((x) => x[0]);
  return 欠.length ? '★画面の 字は 測れません★（欠け＝' + 欠.join(' ') + '）'
    : '客の道（switchSheet→_字の元→_ゼロを隠すか→_答えは字か→fmtForDisplay(_入る字数(cW))）';
});
const 計算 = {};
const 生の答え = {};
const 判じ = { 見た: 0, 隠した: 0, ゼロの式: 0, ゼロで隠した: 0, ゼロだが隠さない: 0, 作れない: 0, rawの型: {} };
const 使った道 = 道しらべ;
console.log('  ★どの 道で 字を 取ったか★ ' + 使った道);
if (String(使った道).indexOf('測れません') >= 0) {
  console.log('★★画面の 字を 作れない 木です＝数を 出しません★★');
  await browser.close(); server.close(); process.exit(8);
}
for (let i = 0; i < 板の数; i++) {
  const 一枚 = await page.evaluate((idx) => {
    window.switchSheet(idx);            /* ★お客さんが 札を 押すのと 同じ★ */
    const sh = window.sheets[idx];
    const 名 = (sh && sh.name) || '';
    const d = (sh && sh.data) || {};
    const 画 = {}; const 生 = {};
    const 数 = { 見た: 0, 隠した: 0, ゼロの式: 0, ゼロで隠した: 0, ゼロだが隠さない: 0, 作れない: 0, rawの型: {} };
    Object.keys(d).forEach((k) => {
      const c = d[k];
      if (!c || typeof c !== 'object') return;
      if (c.f === undefined || c.f === null || c.f === '') return;
      const 生値 = (c.v !== undefined && c.v !== null && c.v !== '') ? c.v : c.d;
      生[名 + '|' + k] = 生値;
      let raw = null;
      try { raw = window._字の元(c); } catch (e) { raw = null; }
      if (!raw) { 画[名 + '|' + k] = ''; return; }
      let 隠 = false;
      try { 隠 = !!window._ゼロを隠すか(c, raw); } catch (e) { 隠 = false; }
      数.見た++;
      if (隠) 数.隠した++;
      const r2 = Number(String(raw).trim());
      if (Number.isFinite(r2) && r2 === 0) {
        数.ゼロの式++;
        if (隠) 数.ゼロで隠した++; else 数.ゼロだが隠さない++;
      }
      数.rawの型[typeof raw] = (数.rawの型[typeof raw] || 0) + 1;
      if (隠) { 画[名 + '|' + k] = ''; return; }
      let 字 = '';
      try {
        if (window._答えは字か(c)) 字 = String(raw);
        else {
          const 列 = Number(String(k).split(',')[1]);
          字 = String(window.fmtForDisplay(raw, c.numFmt, window._入る字数(window.cW(列), raw, c.numFmt)));
        }
      } catch (e) { 字 = '★作れません★'; 数.作れない++; }
      画[名 + '|' + k] = 字;
    });
    return { 画: 画, 生: 生, 数: 数 };
  }, i);
  Object.assign(計算, 一枚.画);
  Object.assign(生の答え, 一枚.生);
  ['見た', '隠した', 'ゼロの式', 'ゼロで隠した', 'ゼロだが隠さない', '作れない']
    .forEach((x) => { 判じ[x] += 一枚.数[x]; });
  Object.keys(一枚.数.rawの型).forEach((t) => { 判じ.rawの型[t] = (判じ.rawの型[t] || 0) + 一枚.数.rawの型[t]; });
}
console.log('  ★開いた 板★ ' + 板の数 + '枚（★1枚ずつ `switchSheet` で 開きました★）');
if (判じ) {
  console.log('');
  console.log('★★「ゼロを 隠す」の 判じ（★式の マスだけ★）★★');
  console.log('  判じた 式 .............. ' + 判じ.見た.toLocaleString() + '個');
  console.log('  ★隠すと 判じた★ ....... ★' + 判じ.隠した.toLocaleString() + '個★');
  console.log('  中の 字が 0 の 式 ...... ' + 判じ.ゼロの式.toLocaleString() + '個');
  console.log('    うち 隠した .......... ' + 判じ.ゼロで隠した.toLocaleString() + '個');
  console.log('    ★うち 隠さなかった★ . ★' + 判じ.ゼロだが隠さない.toLocaleString() + '個★');
  console.log('  ★`_字の元` が 返す 型★ ' + JSON.stringify(判じ.rawの型));
  console.log('  ★画面の 字を 作れなかった 数★ ' + (判じ.作れない || 0).toLocaleString() + '個');
  if (判じ.ゼロだが隠さない > 0 && 判じ.ゼロで隠した === 0) {
    console.log('  ⇒★★式の マスでは 1つも 隠れて いません★★');
  }
}
await browser.close();
server.close();

/* ── ★突き合わせ★ ── */
/* ★★字と 数を 跨いで 比べる★★
     ＝★計算の 側は いつも 字★（`_台の値を字に`）／★保存の 側は 数★
     ⇒★両方が 数として 読めるなら 数で 比べる★
     ⇒★でなければ 字で 比べる★
     ★ここで 緩めすぎると 嘘の 緑に なる★ので ★丸めの 幅だけ★ 許す */
const 数に = (v) => {
  if (typeof v === 'number') return Number.isFinite(v) ? v : null;
  if (typeof v !== 'string') return null;
  const t = v.trim();
  if (t === '') return null;
  const n = Number(t);
  return Number.isFinite(n) ? n : null;
};
const 同じ字 = (a, b) => {
  if (a === b) return true;
  const x = 数に(a), y = 数に(b);
  if (x !== null && y !== null) {
    if (x === y) return true;
    const d = Math.abs(x - y), s = Math.max(Math.abs(x), Math.abs(y));
    return d <= (s > 0 ? s * 1e-10 : 1e-10);          /* ★丸めの 幅だけ 許す★ */
  }
  if (a === undefined || a === null) return b === undefined || b === null || b === '';
  if (b === undefined || b === null) return a === '';
  return String(a) === String(b);
};
let 突き合わせた = 0, 同じ = 0, 違う = 0, 保存だけ = 0, 計算だけ = 0;
const 違いの形 = {};
const 違いの中身 = {};   /* ★中身は 出さず 形だけ★ */
const 長さの例 = {};
Object.keys(保存).forEach((k) => {
  if (!(k in 生の答え)) { 保存だけ++; return; }
  突き合わせた++;
  if (同じ字(保存[k], 生の答え[k])) 同じ++;
  else {
    違う++;
    const f = (v) => (v === undefined || v === null) ? '空'
      : (typeof v === 'number') ? '数'
        : (typeof v === 'string' && v.charAt(0) === '#') ? '誤'
          : (typeof v === 'boolean') ? '真偽' : '字';
    const 形 = '保存=' + f(保存[k]) + ' → 計算=' + f(生の答え[k]);
    違いの形[形] = (違いの形[形] || 0) + 1;
    /* ★★中身を 出さずに 「形」だけ 分ける★★（2026-09-25）
         ＝★2,910個が 「数 → 字」で 止まった＝それだけでは 何も 分からない★
         ＝★値そのものは 出しません★＝★どんな 字の 並びかだけ★ */
    const v = 生の答え[k];
    let 札;
    if (v === undefined || v === null) 札 = '①計算側が 無い';
    else if (typeof v === 'string' && v === '') 札 = '②計算側が ★空の 字★';
    else if (typeof v !== 'string') 札 = '⑧字で ない（' + typeof v + '）';
    else if (/^[#]/.test(v)) 札 = '③計算側が 誤り（#...）';
    else if (/^-?[0-9,]+(\.[0-9]+)?$/.test(v)) 札 = '④数字と カンマだけ';
    else if (/[/\-]/.test(v) && /[0-9]/.test(v)) 札 = '⑤数字と 斜線か 横棒（日付らしい）';
    else if (/^(TRUE|FALSE)$/.test(v)) 札 = '⑥真偽';
    else if (/[0-9]/.test(v)) 札 = '⑦数字を 含む 別の 字';
    else 札 = '⑨数字を 含まない 字';
    違いの中身[札] = (違いの中身[札] || 0) + 1;
    if (!長さの例[札]) 長さの例[札] = { 最短: v == null ? 0 : String(v).length, 最長: 0 };
    const L = v == null ? 0 : String(v).length;
    if (L < 長さの例[札].最短) 長さの例[札].最短 = L;
    if (L > 長さの例[札].最長) 長さの例[札].最長 = L;
  }
});
Object.keys(生の答え).forEach((k) => { if (!(k in 保存)) 計算だけ++; });

console.log('');
console.log('★★★保存されて いた 答え vs Exally が 計算した 答え★★★');
console.log('  ★これは ★答えそのもの★で 比べます（★画面で 隠すかどうかは 関わりません★）★');
/* ★★突き合わせが 0件なら 赤★★（2026-09-25・★この 門が 無くて 嘘を 出した★）
     ＝鍵の 形が 合って いないと ★「違い 0個」＝「全部 同じ」★に 見える
     ＝[[feedback_hakaru_dougu_ga_kaeshita_0_wo_shinjiruna]] を 自分で 踏んだ */
console.log('  突き合わせた 式 ...... ' + 突き合わせた.toLocaleString() + '個');
console.log('  ★同じ★ ............. ★' + 同じ.toLocaleString() + '個★'
  + (突き合わせた ? '（' + (同じ / 突き合わせた * 100).toFixed(2) + '%）' : ''));
console.log('  ★違う★ ............. ★' + 違う.toLocaleString() + '個★'
  + (突き合わせた ? '（' + (違う / 突き合わせた * 100).toFixed(2) + '%）' : ''));
console.log('  保存に だけ 在る ..... ' + 保存だけ.toLocaleString() + '個（★Exally が 式と 見て いない★）');
console.log('  計算に だけ 在る ..... ' + 計算だけ.toLocaleString() + '個（★保存に 答えが 無い＝㋑★）');
if (違う) {
  console.log('');
  console.log('  ★違いの 形★');
  Object.keys(違いの形).sort((a, b) => 違いの形[b] - 違いの形[a])
    .forEach((k) => console.log('    ' + k.padEnd(28) + String(違いの形[k]).padStart(7) + '個'));
  console.log('');
  console.log('  ★計算側の 字は どんな 形か★（★中身は 出しません★）');
  Object.keys(違いの中身).sort((a, b) => 違いの中身[b] - 違いの中身[a])
    .forEach((k) => console.log('    ' + k.padEnd(34) + String(違いの中身[k]).padStart(7) + '個'
      + '   字数 ' + 長さの例[k].最短 + '〜' + 長さの例[k].最長));
}

if (突き合わせた === 0) {
  console.log('');
  console.log('★★★突き合わせが 0件です＝鍵が 噛み合って いません★★★');
  console.log('  ⇒★「違い 0個」を 「全部 同じ」と 読んでは いけません★');
  process.exit(6);
}

/* ══ ★★字 vs 字★★ ══
     ★物差しは 2つ 在ります★
       ・`--実Excel` を 渡した ⇒★実Excel が 出した 字★（★これが 正★）
       ・渡さない          ⇒借り物（SheetJS）の `w`（★一部 壊れて います★）
     ⇒★どちらを 使ったかを 出しに 必ず 書きます★ */
const 物差し = 実道 ? 実字 : 保存字;
const 物差しの名 = 実道 ? '★実Excel が 出した 字★' : '借り物（SheetJS）の `w`（★一部 壊れて います★）';
let 字突き = 0, 字同じ = 0, 字違う = 0, 字なし = 0;
const 字違いの形 = {};
const 字違いの訳 = { 物差しが空: 0, 空とゼロ: 0, 空と非ゼロ: 0, うちが空: 0, 両方字あり: 0 };
const 形の訳 = { 日付: 0, カンマ: 0, 小数: 0, その他: 0 };
const 居場所 = {};
Object.keys(保存).forEach((k) => {
  if (!(k in 計算)) return;
  if (!(k in 物差し)) { 字なし++; return; }
  字突き++;
  const a = 物差し[k];
  const b = (計算[k] === undefined || 計算[k] === null) ? '' : String(計算[k]);
  if (a === b) { 字同じ++; return; }
  /* ★前後の 空白だけの 違いは 同じと 見ます★（Excel は 書式で 右寄せの 空きを 入れる） */
  if (a.trim() === b.trim()) { 字同じ++; return; }
  字違う++;
  const 形 = (a.trim().replace(/[0-9]/g, '9').slice(0, 12)) + ' → ' + (b.trim().replace(/[0-9]/g, '9').slice(0, 12));
  字違いの形[形] = (字違いの形[形] || 0) + 1;
  /* ★★因ごとに 分ける★★（2026-09-25）
       ★実測★ ... 15枚中 ★13枚★が `DisplayZeros = False`（★ゼロを 隠して います★）
       ⇒★Excel が 空で Exally が 0 を 出して いる★分は ★その 1つの 因★
       ⇒★因ごとに 分けないと 「9,163個 直せ」に 見えます★ */
  /* ★居場所を 集めます（★値は 入れません★）★ */
  const 札 = (a.trim() === '')
    ? (Number(b.trim().replace(/[^0-9.+-]/g, '')) === 0 ? '実Excelは空・うちは書式つきゼロ' : '実Excelは空・うちは字')
    : (b.trim() === '' ? 'うちが空' : (/^#/.test(a.trim()) ? '★実Excelが誤り・うちは字★' : '両方字あり'));
  if (!居場所[札]) 居場所[札] = [];
  if (居場所[札].length < 例数) 居場所[札].push(k);
  if (a.trim() === '') {
    字違いの訳.物差しが空++;
    const n2 = Number(b.trim());
    if (Number.isFinite(n2) && n2 === 0) 字違いの訳.空とゼロ++;
    else 字違いの訳.空と非ゼロ++;
  } else if (b.trim() === '') {
    字違いの訳.うちが空++;
  } else {
    字違いの訳.両方字あり++;
    /* ★★2,373個を ★形ごとに★ 分ける★★（2026-09-25）
         ★なぜ★ ... 「2,373個 直せ」では ★どこから 手を 付けるか 決まりません★
         ★分け方（★Excel の 字を 見て 決めます★）★
           ㋐カンマ ... Excel に `,` が 在り うちに 無い
           ㋑小数   ... 小数点の 後の 桁数が 違う
           ㋒日付   ... Excel に `/` か `-` が 在り うちに 無い（生の 数字）
           ㋓その他
         ★1つの マスが 2つに 当たる 事が 在ります★⇒★先に 当たった 1つだけ 数えます★
         ⇒★だから 4つの 合計は 2,373 に なります★ */
    const A = a.trim(), B = b.trim();
    const 小数桁 = (x) => { const i = x.indexOf('.'); return i < 0 ? -1 : x.length - i - 1; };
    if (/[/-]/.test(A) && !/[/-]/.test(B) && /^[0-9.]+$/.test(B)) 形の訳.日付++;
    else if (A.indexOf(',') >= 0 && B.indexOf(',') < 0) 形の訳.カンマ++;
    else if (小数桁(A) !== 小数桁(B)) 形の訳.小数++;
    else 形の訳.その他++;
  }
});
console.log('');
console.log('★★★Excel が 表示して いた 字 vs Exally が 出した 字★★★');
console.log('  ★物差し★ ' + 物差しの名);
if (String(使った道).indexOf('測れません') >= 0) {
  console.log('  ★★画面の 字を 作れない 木です＝この 突き合わせは 出しません★★');
  console.log('  ⇒' + 使った道);
  process.exit(8);
}
console.log('  突き合わせた ......... ' + 字突き.toLocaleString() + '個'
  + '（★物差しに 無くて 比べられない ' + 字なし.toLocaleString() + '個は 分母の 外★）');
console.log('  ★同じ★ ............. ★' + 字同じ.toLocaleString() + '個★'
  + (字突き ? '（' + (字同じ / 字突き * 100).toFixed(2) + '%）' : ''));
console.log('  ★違う★ ............. ★' + 字違う.toLocaleString() + '個★'
  + (字突き ? '（' + (字違う / 字突き * 100).toFixed(2) + '%）' : ''));
if (字違う) {
  console.log('');
  console.log('  ★★違いの 訳（★因ごとに 分けました★）★★');
  console.log('    ★Excel は 空・うちは ★0★★ ................ ★' + 字違いの訳.空とゼロ.toLocaleString() + '個★'
    + '   ⇒★因は 「ゼロを 表示しない」（実測 15枚中 13枚）★');
  console.log('    Excel は 空・うちは ★0 でない 字★ ......... ' + 字違いの訳.空と非ゼロ.toLocaleString() + '個'
    + '   ⇒★別の 因★');
  console.log('    うちが 空・Excel は 字 .................. ' + 字違いの訳.うちが空.toLocaleString() + '個');
  console.log('    ★両方 字が 在るが 違う★ ................. ★' + 字違いの訳.両方字あり.toLocaleString() + '個★'
    + '   ⇒★書式（桁の 区切り・小数・日付）★');
  console.log('');
  console.log('  ★★「両方 字が 在るが 違う」の 内訳（★1マス 1つだけ 数えます★）★★');
  console.log('    ㋒★日付が 生の 数字★ ........ ★' + 形の訳.日付.toLocaleString() + '個★');
  console.log('    ㋐★桁の 区切りが 無い★ ...... ★' + 形の訳.カンマ.toLocaleString() + '個★');
  console.log('    ㋑★小数の 桁が 違う★ ........ ★' + 形の訳.小数.toLocaleString() + '個★');
  console.log('    ㋓その他 .................... ' + 形の訳.その他.toLocaleString() + '個');
  console.log('    （合計 ' + (形の訳.日付 + 形の訳.カンマ + 形の訳.小数 + 形の訳.その他).toLocaleString()
    + '個 ＝ 両方 字が 在るが 違う ' + 字違いの訳.両方字あり.toLocaleString() + '個）');
  console.log('');
  console.log('  ★違いの 形（★数字は 9 に 伏せて います＝中身は 出しません★）★ 多い順 12');
  Object.keys(字違いの形).sort((a, b) => 字違いの形[b] - 字違いの形[a]).slice(0, 12)
    .forEach((k) => console.log('    ' + k.padEnd(30) + String(字違いの形[k]).padStart(7) + '個'));
}

if (例数 > 0) {
  console.log('');
  console.log('★★違った マスの 居場所（★値は 1つも 出して いません★・各 ' + 例数 + '個まで）★★');
  console.log('  ★★この 出しは repo に 入れません（板の 名が 入って います）★★');
  Object.keys(居場所).forEach((k) => {
    console.log('  ' + k);
    居場所[k].forEach((x) => console.log('    ' + x));
  });
}

console.log('');
console.log('★★この 本は どの 場合か★★');
const いる理由 = [];
if (日鍵.length) いる理由.push('㋐開いた 日で 答えが 変わる 関数が ' + 日鍵.length + '種類 在る');
if (値なし式) いる理由.push('㋑答えが 保存されて いない 式が ' + 値なし式.toLocaleString() + '個 在る');
if (字違う) いる理由.push('★Excel の 字と Exally の 字が ' + 字違う.toLocaleString()
  + '個 違う＝どちらが 正しいかを 先に 割る★');
if (!いる理由.length) {
  console.log('  ★㋐㋑ とも 0 ／ 保存と 計算も 全部 同じ★');
  console.log('  ⇒★この 本では 「開いた 直後の 計算」は 新しい 物を 1つも 生んで いません★');
} else {
  いる理由.forEach((x) => console.log('  ★' + x + '★'));
}
console.log('');
console.log('★★言い落とさない 事★★');
console.log('  ・★㋐の 名簿は 手で 持って います（' + 日で変わる.length + '個）＝★名簿に 無い 物は 見つかりません★');
console.log('  ・★㋓の 印は `.bin` の 台帳だと この 道具では 読めません＝未測定★');
console.log('  ・★これは 1冊の 本の 数です＝★他の 本では 別の 数★');
