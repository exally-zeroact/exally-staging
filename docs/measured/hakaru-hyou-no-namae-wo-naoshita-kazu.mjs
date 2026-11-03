/* hakaru-hyou-no-namae-wo-naoshita-kazu.mjs ･･･ ★表の 名前を 直した マスの 数★ 2026-09-25
 *
 *  ★★なぜ 要るか（経営者1 の 宿題）★★
 *    経営者1 が 実Excel で 取り直しました:
 *      給料3 の 4マス ･･･ ★実Excel の 字は `#REF!`／本に 保存された 値は 数★
 *      板3 まるごと 654個 ･･･ ★実Excel の 字が 誤り 68個／本の 保存値が 誤り 0個★
 *      15枚 15,799個 ･･･ ★式の 字に `#REF!` が 在るのは 1個だけ★
 *                          ★なのに 実Excel は 69個 誤りを 出して いる★
 *                          ⇒★68個は 式の 字からは 拾えない★
 *    ⇒経営者1 の 見立て（★未測定★）
 *      「★借り物が 表の 名前を 捨てる★ ⇒ `c.f` から `#REF!` が 消える」
 *      「★`lib/table-refs.js` が 生きた A1 の 範囲に 直す★ ⇒ うちは 数を 出す」
 *      「＝★直しが 効きすぎて いる★」
 *    ⇒★これは 私の 持ち場です★（経営者1 は 実Excel 側しか 測れません）
 *
 *  ★★何を 数えるか★★
 *    ⑴`TableRefs.resolve` が ★直した マスの 数★（`stats` と `fixes` を そのまま 出す）
 *    ⑵★その 番地の 一覧★（板ごとの 数）
 *    ⑶`--番地` で 渡した マスが ★直された 中に 入って いるか★
 *    ⇒★これで 「68個が 直しの 産物か」が 決まります★
 *
 *  ★出すのは 数と 番地だけ★（★式の 字も 値も 1つも 出しません★）
 *
 *  ★走らせ方★
 *    node docs/measured/hakaru-hyou-no-namae-wo-naoshita-kazu.mjs "<本の 道>" [--番地 "板|行,列" ･･･]
 */
import path from 'node:path'; import http from 'node:http'; import fs from 'node:fs';
import { fileURLToPath } from 'node:url';
import { borrow, launch } from '../../scripts/_borrow-playwright.mjs';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..', '..');
const 引数 = process.argv.slice(2);
const 札の値 = ['台', '番地表'];
const 素 = (() => {
  const 出 = [];
  for (let i = 0; i < 引数.length; i++) {
    const a = 引数[i];
    if (a.slice(0, 2) === '--') { if (札の値.indexOf(a.slice(2)) >= 0) i++; continue; }
    出.push(a);
  }
  return 出;
})();
const 本 = 素[0];
const 見る番地 = [];
for (let i = 0; i < 引数.length; i++) if (引数[i] === '--番地' && 引数[i + 1]) 見る番地.push(引数[i + 1]);
/* ★★`--番地表 <紙>` ･･･ 1行 1番地（★4個では 分母が 小さすぎる★）★★ */
{
  const 紙 = 取る0('番地表', null);
  if (紙) {
    if (!fs.existsSync(紙)) { console.log('★番地表が 在りません★ ' + 紙); process.exit(2); }
    io読み(紙).forEach((x) => 見る番地.push(x));
  }
}
const 取る = (名, 既定) => { const i = 引数.indexOf('--' + 名); return i >= 0 ? 引数[i + 1] : 既定; };
function 取る0(名, 既定) { const i = 引数.indexOf('--' + 名); return i >= 0 ? 引数[i + 1] : 既定; }
function io読み(紙) {
  /* ★逆斜線を 1つも 書きません★＝★便りや heredoc で 落ちて 別の 物に なる★
       （2026-09-25 ここで 1回 踏んだ ･･･ `\r` が 本物の CR に なって 正規表現が 割れた） */
  const 改行 = String.fromCharCode(10);
  const 戻り = String.fromCharCode(13);
  return fs.readFileSync(紙, "utf8").split(改行)
    .map((x) => x.split(戻り).join("").trim())
    .filter((x) => x && x.indexOf("|") > 0);
}
const 台 = 取る('台', 'webkit');
if (!本 || !fs.existsSync(本)) { console.log('★本が 在りません★ ' + 本); process.exit(2); }

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

console.log('[表の 名前を 直した 数] ★出すのは 数と 番地だけ＝式の 字も 値も 出しません★');
console.log('  ★本★ ' + path.basename(本) + '（★読むだけ／1バイトも 書きません★）／★台★ ' + 台);
if (見る番地.length) console.log('  ★重なりを 見る 番地★ ' + 見る番地.length + '個');

const wk = await borrow('hyou-no-namae', 台);
const browser = await launch('hyou-no-namae', wk, {}, 台);
const page = await browser.newPage({ viewport: { width: 1280, height: 900 } });
const 配信 = await 立てる(ROOT);

try {
  page.on('pageerror', (e) => console.log('      [落ちた] ' + String(e.message).slice(0, 200)));
  await page.goto(配信.url + '/book.html', { waitUntil: 'load', timeout: 180000 });
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
  }, null, { timeout: 180000 });
  await page.waitForTimeout(800);

  /* ★★要る 物が 無ければ 数を 出さずに 止まります★★（2026-09-25 の 決め） */
  /* ★★`fixes` は 画面に 残りません★★（`BookOpen.current()` の 鍵に 在りません）
       ★残るのは `tableRefs`（数だけ）★。
       ⇒だから ★★同じ 関数（`TableRefs.resolve`）を 同じ 材料で もう 一度 呼びます★★
       ＝★真似ません＝お客さんの 道が 呼ぶ 物 そのものです★
       ＝★材料も 同じ（`BookOpen.current().bytes`）★
       ⇒★残って いる 数と 合うか を 先に 見ます★（合わなければ 空振り） */
  const 無い = await page.evaluate(() => {
    const 足りない = [];
    if (!window.BookOpen || typeof window.BookOpen.current !== 'function') 足りない.push('BookOpen.current');
    if (!window.TableRefs || typeof window.TableRefs.resolve !== 'function') 足りない.push('TableRefs.resolve');
    if (!window.ZipSurgeon) 足りない.push('ZipSurgeon');
    if (!window.XLSX) 足りない.push('XLSX');
    return 足りない;
  });
  if (無い.length) { console.log('  ★★測れません★★ 在りません ... ' + 無い.join(' / ')); process.exit(8); }

  const 出 = await page.evaluate(async (見る番地) => {
    const cur = window.BookOpen.current();
    /* ★開いた 時に 作られた 物を そのまま 読みます★＝★もう 一度 走らせません★
         （もう 一度 走らせると ★測る 為に 別の 回を 作る★ ことに なります） */
    const 残った数 = (cur && cur.tableRefs) || null;
    if (!cur || !cur.bytes) return { fixes無し: true, stats: null, 断り: null, 鍵: Object.keys(cur || {}) };
    const wb = window.XLSX.read(cur.bytes, { type: 'array', cellFormula: true, cellStyles: true, bookVBA: true });
    const r = await window.TableRefs.resolve(cur.bytes, cur.kind, wb, window.ZipSurgeon);
    if (!r || !r.ok) return { fixes無し: true, stats: (r && r.stats) || null, 断り: (r && r.why) || null, 鍵: Object.keys(cur) };
    const fixes = r.fixes || {};
    const stats = r.stats || null;
    const 断り = (r.断り && r.断り.length) ? r.断り.join(' / ') : null;
    const 番地 = Object.keys(fixes);
    const 板ごと = {};
    番地.forEach((k) => { const な = String(k).split('|')[0]; 板ごと[な] = (板ごと[な] || 0) + 1; });
    const 重なり = 見る番地.map((k) => ({ 番地: k, 直された: Object.prototype.hasOwnProperty.call(fixes, k) }));
    return { 直した数: 番地.length, 板ごと: 板ごと, stats: stats, 断り: 断り, 重なり: 重なり, 残った数: 残った数 };
  }, 見る番地);

  console.log('');
  if (出.fixes無し) {
    console.log('  ★★`tableFixes` が 在りません★★＝★この 道具は 空振りです★');
    console.log('    `BookOpen.current()` の 鍵 ... ' + (出.鍵 || []).join(' / '));
    if (出.stats) console.log('    `tableStats` ... ' + JSON.stringify(出.stats));
    process.exit(8);
  }
  console.log('  ══ ★`TableRefs` が 直した マス★ ══');
  console.log('    ★' + 出.直した数.toLocaleString() + '個★');
  if (出.stats) console.log('    `stats` ... ' + JSON.stringify(出.stats));
  console.log('    ★開いた 時に 残った 数★ ... ' + JSON.stringify(出.残った数));
  if (出.残った数 && 出.stats && 出.残った数.fixed !== undefined && 出.残った数.fixed !== 出.stats.fixed) {
    console.log('    ★★合いません＝この 道具は 空振りです★★ 開いた時 ' + 出.残った数.fixed + ' ／ 今 ' + 出.stats.fixed);
    process.exit(8);
  }
  console.log('    ⇒★開いた 時と 同じ 数です★＝★空振りでは ありません★');
  if (出.断り) console.log('    断り ..... ' + String(出.断り).slice(0, 300));
  console.log('    ── 板ごと ──');
  Object.keys(出.板ごと).sort((a, b) => 出.板ごと[b] - 出.板ごと[a]).forEach((な) => {
    console.log('      ' + な + ' ... ' + 出.板ごと[な].toLocaleString() + '個');
  });

  if (出.重なり && 出.重なり.length) {
    console.log('');
    console.log('  ══ ★渡した 番地が 直されて いるか★ ══');
    /* ★20個を 超えたら 1つずつ 出しません（★出しを 溢れさせない★）★ */
    if (出.重なり.length <= 20) {
      出.重なり.forEach((x) => {
        console.log('    ' + x.番地 + ' ... ' + (x.直された ? '★直された★' : '直されて いない'));
      });
    } else {
      const 外 = 出.重なり.filter((x) => !x.直された);
      console.log('    ★直されて いない 番地★ ... ' + 外.length + '個'
        + (外.length ? '  ' + 外.map((x) => x.番地).slice(0, 20).join(' / ') : ''));
      const 行ごと = {};
      出.重なり.filter((x) => x.直された).forEach((x) => {
        const 行 = x.番地.split('|')[0] + '|' + String(x.番地.split('|')[1]).split(',')[0] + '行';
        行ごと[行] = (行ごと[行] || 0) + 1;
      });
      console.log('    ★直された 番地（行ごと）★');
      Object.keys(行ごと).forEach((k) => console.log('      ' + k + ' ... ' + 行ごと[k] + '個'));
    }
    const 何個 = 出.重なり.filter((x) => x.直された).length;
    console.log('    ★' + 何個 + '個 / ' + 出.重なり.length + '個 が 直された マス★');
    /* ══ ★★「半分」と 言うのは 66/69 でも 0/69 でも 同じ 字に なります★★ ══（2026-09-25 1回 踏んだ）
         ⇒★割合を 出して 段で 言い分けます★（★分母は いつも 隣に 置きます★） */
    const 割 = 出.重なり.length ? Math.round((何個 / 出.重なり.length) * 1000) / 10 : 0;
    console.log('    ⇒★' + 割 + '%★ が 直した 中');
    if (何個 === 出.重なり.length) {
      console.log('    ⇒★★全部＝経営者1 の 見立て（直しが 効きすぎ）で 説明できます★★');
    } else if (何個 === 0) {
      console.log('    ⇒★★1つも 在りません＝経営者1 の 見立ては 外れです★★');
    } else if (割 >= 90) {
      console.log('    ⇒★★大半＝経営者1 の 見立てで 説明できます★★');
      console.log('      ＝★但し 残り ' + (出.重なり.length - 何個) + '個は 別の 訳です★（★0個に しない★）');
    } else {
      console.log('    ⇒★訳が 2つ 以上 在ります＝1つでは 説明できません★');
    }
  }
  console.log('  ★式の 字も 値も 1つも 出して いません★（数と 番地だけ）');
} finally {
  await page.close().catch(() => {});
  await browser.close().catch(() => {});
  配信.閉じる();
}
