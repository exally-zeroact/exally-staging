/* hakaru-nama-no-byte-ni-ayamari-sanshou-ga-aru-ka.mjs
 *   ･･･ ★生の バイトの 式に 「誤りの 参照」が 在るか★ 2026-09-25
 *
 *  ★★なぜ 要るか（経営者1 の 宿題の 最後の 1本）★★
 *    実Excel が `#REF!` を 出す マス ★69個★（経営者1 の 実測・物差しは 実Excel の `.Text`）
 *    その うち ★66個は `lib/table-refs.js` が 直した 中に 在りました★（私の 実測）
 *    外の 3個 ･･･ SUM / SUM / SUBTOTAL ＝★範囲を まとめる 関数★
 *      ★2個は 壊れた かたまりの 中の 行を 足して います★（Excel 75行・76行）
 *      ⇒★つまり 69個の うち 68個は 1つの 訳で 説明が つきます★
 *    ⇒★但し `table-refs.js` は `refused 0`＝★1つも 「直せない」と 断って いません★★
 *    ⇒経営者1「★`.xlsb` の 生の バイトから 式を 読め★」
 *
 *  ★★何を 数えるか★★
 *    `.xlsb` の 式は ★Ptg（トークン）の 並び★で 入って います。
 *    その 中に ★「壊れた 参照」だけの ための トークン★が 在ります:
 *      `PtgRefErr`     0x2A / 0x4A / 0x6A
 *      `PtgAreaErr`    0x2B / 0x4B / 0x6B
 *      `PtgRefErr3d`   0x3C / 0x5C / 0x7C
 *      `PtgAreaErr3d`  0x3D / 0x5D / 0x7D
 *    ⇒★これが 在れば 「実Excel が `#REF!` を 出す」と 本の 中で 決まって います★
 *    ⇒★在るのに うちが 数を 出して いるなら ★それが 直し過ぎの 正体★です★
 *
 *  ★★番号表を 記憶で 書いて いません★★（★記憶の 決まり★）
 *    ＝★数えた 結果で 決めます★。渡した 番地（実Excel が 誤りを 出す）と
 *      ★渡して いない マス★の 両方で 数え、★差が 立つか★を 見ます。
 *    ⇒★差が 立たなければ 「この 印では 決められません」と 出して 止めます★
 *
 *  ★出すのは 数だけ★（★式の 字も 値も 1つも 出しません★）
 *
 *  ★走らせ方★
 *    node docs/measured/hakaru-nama-no-byte-ni-ayamari-sanshou-ga-aru-ka.mjs "<本>" --番地表 <紙>
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
const 取る = (名, 既定) => { const i = 引数.indexOf('--' + 名); return i >= 0 ? 引数[i + 1] : 既定; };
const 本 = 素[0];
const 台 = 取る('台', 'webkit');
const 番地表 = 取る('番地表', null);
if (!本 || !fs.existsSync(本)) { console.log('★本が 在りません★ ' + 本); process.exit(2); }
if (!番地表 || !fs.existsSync(番地表)) { console.log('★番地表が 在りません★ ' + 番地表); process.exit(2); }

/* ★逆斜線を 1つも 書きません★＝★便りや heredoc で 落ちて 別の 物に なる★ */
const 印の番地 = (() => {
  const 改行 = String.fromCharCode(10);
  const 戻り = String.fromCharCode(13);
  return fs.readFileSync(番地表, 'utf8').split(改行)
    .map((x) => x.split(戻り).join('').trim())
    .filter((x) => x && x.indexOf('|') > 0);
})();

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

console.log('[生の バイトの 誤り参照] ★出すのは 数だけ＝式の 字も 値も 出しません★');
console.log('  ★本★ ' + path.basename(本) + '（★読むだけ／1バイトも 書きません★）／★台★ ' + 台);
console.log('  ★印の 番地（実Excel が 誤りを 出す）★ ' + 印の番地.length + '個');

const wk = await borrow('nama-no-byte', 台);
const browser = await launch('nama-no-byte', wk, {}, 台);
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
  const 無い = await page.evaluate(() => {
    const 足りない = [];
    const T = window.TableRefs;
    if (!T) 足りない.push('TableRefs');
    else {
      if (typeof T._walkRecords !== 'function') 足りない.push('TableRefs._walkRecords');
      if (typeof T._rgceRange !== 'function') 足りない.push('TableRefs._rgceRange');
      if (typeof T._partToSheetName !== 'function') 足りない.push('TableRefs._partToSheetName');
    }
    if (!window.ZipSurgeon) 足りない.push('ZipSurgeon');
    if (!window.BookOpen || typeof window.BookOpen.current !== 'function') 足りない.push('BookOpen.current');
    return 足りない;
  });
  if (無い.length) { console.log('  ★★測れません★★ 在りません ... ' + 無い.join(' / ')); process.exit(8); }

  const 出 = await page.evaluate(async (印の番地) => {
    const T = window.TableRefs;
    const cur = window.BookOpen.current();
    if (!cur || !cur.bytes) return { だめ: '★開いた 本の バイトが 在りません★' };
    if (String(cur.kind) !== 'xlsb') return { だめ: '★`.xlsb` では ありません ... ' + cur.kind + '★' };
    const zip = window.ZipSurgeon.read(cur.bytes);
    const 部品 = zip.names().filter((n) => /^xl[/]worksheets[/]sheet[0-9]+[.]bin$/.test(n));
    if (!部品.length) return { だめ: '★板の `.bin` が 在りません★' };
    const 板の名 = await T._partToSheetName(zip, true);

    /* ★壊れた 参照だけの ための トークン★（★記憶で 書かず 下で 差を 見ます★） */
    const 誤り印 = [0x2A, 0x4A, 0x6A, 0x2B, 0x4B, 0x6B, 0x3C, 0x5C, 0x7C, 0x3D, 0x5D, 0x7D];
    const u32 = (b, p) => (b[p] | (b[p + 1] << 8) | (b[p + 2] << 16) | (b[p + 3] << 24)) >>> 0;
    const colName = (c) => { let s = ''; c += 1; while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = (c - 1 - m) / 26; } return s; };

    const 印 = {};
    印の番地.forEach((k) => { 印[k] = true; });

    const 数 = { 式: 0, 誤り印あり: 0, 印の中: 0, 印の中で誤り印あり: 0, 印の外: 0, 印の外で誤り印あり: 0, 範囲取れず: 0 };
    const 印の中で無い番地 = [];

    for (const pn of 部品) {
      const sn = 板の名[pn];
      if (!sn) continue;
      const b = await zip.bytes(pn);
      const recs = T._walkRecords(b);
      if (!recs) continue;
      let rw = 0;
      for (const r of recs) {
        if (r.id === 0) { rw = u32(b, r.start); continue; }
        const 配列の式 = (r.id === 426);
        if (r.id !== 8 && r.id !== 9 && r.id !== 10 && r.id !== 11 && !配列の式) continue;
        const rg = T._rgceRange(r.id, b, r.start, r.len);
        if (!rg) { 数.範囲取れず++; continue; }
        数.式++;
        const 行 = 配列の式 ? u32(b, r.start) : rw;
        const c = 配列の式 ? u32(b, r.start + 8) : u32(b, r.start);
        const key = sn + '|' + 行 + ',' + c;
        /* ★トークンの 並びを 頭から 読まずに 「その バイトが 在るか」だけ 見ます★
             ＝★これは 荒い 見方です★＝だから ★印の 中と 外で 差が 立つか★で 値打ちを 決めます */
        let ある = false;
        for (let p = rg.s; p < rg.s + rg.len; p++) if (誤り印.indexOf(b[p]) >= 0) { ある = true; break; }
        if (ある) 数.誤り印あり++;
        if (印[key]) {
          数.印の中++;
          if (ある) 数.印の中で誤り印あり++; else 印の中で無い番地.push(key);
        } else {
          数.印の外++;
          if (ある) 数.印の外で誤り印あり++;
        }
      }
    }
    return { 数: 数, 印の中で無い番地: 印の中で無い番地.slice(0, 12), 部品: 部品.length };
  }, 印の番地);

  if (出.だめ) { console.log('  ★★測れません★★ ' + 出.だめ); process.exit(8); }

  const n = 出.数;
  const 割 = (a, b) => (b ? Math.round((a / b) * 1000) / 10 : 0);
  console.log('');
  console.log('  ══ ★生の バイトで 数えた★ ══（板の 部品 ' + 出.部品 + '本）');
  console.log('    式の 記録 .............. ' + n.式.toLocaleString() + '個'
    + (n.範囲取れず ? '（★範囲が 取れなかった ' + n.範囲取れず + '個★）' : ''));
  console.log('    ★誤り参照の バイトが 在る★ ... ' + n.誤り印あり.toLocaleString() + '個'
    + '（' + 割(n.誤り印あり, n.式) + '%）');
  console.log('');
  console.log('    ── ★印の 中（実Excel が 誤りを 出す）★ ──');
  console.log('      見た ' + n.印の中 + '個 ／ ★誤り参照あり ' + n.印の中で誤り印あり + '個（' + 割(n.印の中で誤り印あり, n.印の中) + '%）★');
  console.log('    ── 印の 外 ──');
  console.log('      見た ' + n.印の外.toLocaleString() + '個 ／ 誤り参照あり ' + n.印の外で誤り印あり.toLocaleString()
    + '個（' + 割(n.印の外で誤り印あり, n.印の外) + '%）');
  if (出.印の中で無い番地.length) {
    console.log('    ★印の 中で 誤り参照が 無かった 番地★（頭 12個）');
    console.log('      ' + 出.印の中で無い番地.join(' / '));
  }

  console.log('');
  const 中 = 割(n.印の中で誤り印あり, n.印の中);
  const 外 = 割(n.印の外で誤り印あり, n.印の外);
  console.log('  ★差★ 印の中 ' + 中 + '% ／ 印の外 ' + 外 + '%');
  if (n.印の中 === 0) {
    console.log('  ⇒★★印の 番地が 1つも 当たって いません＝この 道具は 空振りです★★');
  } else if (中 >= 90 && 外 <= 10) {
    console.log('  ⇒★★差が 立ちます＝この バイトで 「実Excel が #REF! を 出す」を 先に 知れます★★');
    console.log('    ⇒★`table-refs.js` が ここで 断れば 69個は 消える 見込みです★（★まだ 直して いません★）');
  } else {
    console.log('  ⇒★★差が 立ちません＝この 印では 決められません★★');
    console.log('    ＝★荒い 見方（バイトが 在るかだけ）では 足りません★');
    console.log('    ⇒★トークンの 並びを 頭から 読む 所から 作る 事に なります★（★安請け合いしません★）');
  }
  console.log('  ★式の 字も 値も 1つも 出して いません★（数だけ）');
} finally {
  await page.close().catch(() => {});
  await browser.close().catch(() => {});
  配信.閉じる();
}
