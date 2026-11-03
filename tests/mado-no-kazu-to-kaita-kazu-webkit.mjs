/* mado-no-kazu-to-kaita-kazu-webkit.mjs
 *   ･･･ ★窓に 出る 数と 本当に 書き込まれた 数が 合うか★ 2026-09-25
 *
 *  ★★なぜ 要るか（★今日 実物で 出ました★）★★
 *    経営者1 が お客さんの 道で 押しました（司さんの 実物）:
 *      ★1マスだけ 打って 「書き出す」を 押す★
 *      ⇒窓 ･･･ 「★この3044か所を直して 書き出す★」
 *      ⇒★実際に 書き込まれたのは 1個だけ★（★2通りで 数えて 1個・式は 0個★）
 *    ⇒★★窓の 数が 嘘でした★★
 *    ⇒★お客さんは 「3,044か所 直される」と 読んで 「やめる」を 押します★
 *    ⇒★＝出せる 物が 出せなく なります★
 *
 *  ★★因（字を 読んで 決めました）★★
 *    ⑴`js/book-open.js` の 控えは ★「計算し直した 直後」に 取り直す★ 決まり。
 *       ⇒開いた 直後に 計算しなく なった ので ★控えが 「本の 値」に なりました★
 *       ⇒1打ちで 旗が 下り、その後 計算し直すと ★答えが 控えと 違う★
 *       ⇒★うちの 答えが 全部 「客が 変えた」に 数えられました★
 *    ⑵数える 時と 書く 時で ★物差しが 2つ★（★本番にも 在ります★）
 *       比べる ･･･ `normForCompare`（`,` `¥` 空白を 外す）
 *       書く ..... `isNaN(Number(val))` なら 飛ばす（★外さない★）
 *       ⇒★「1,234 円」は 「変わった」に 数えられ、書く 時は 飛ばされます★
 *
 *  ★★当てる 数は 3個です★★（★1個 では ありません★）
 *    打った A2 ＋ ★つられた B2（=A2*2）と C2（=A2+1）★
 *    ＝★つられた 分は 書かないと いけません★（別の 板の 合計が 古いまま 保存される）
 *    ⇒★壊れて いる 時は 81個（式 80個 ぜんぶ）★
 *
 *  ★★この 見張りが 見る 物★★
 *    ⑴★1マス 打った 時、窓に 出る 数が 1か★
 *    ⑵★書き出した 本を 開き直して 「変わった 値」が 1個か★
 *    ⑶★窓の 数と 変わった 数が 同じか★  ←★これが 一番 効きます★
 *    ⇒★実Excel を 使いません★（★借り物で 開き直して 値を 比べます★）
 *
 *  ★走らせ方★ node tests/mado-no-kazu-to-kaita-kazu-webkit.mjs
 */
import path from 'node:path'; import http from 'node:http'; import fs from 'node:fs';
import os from 'node:os';
import { createRequire } from 'node:module';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { borrow, launch } from '../scripts/_borrow-playwright.mjs';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(pathToFileURL(path.join(ROOT, 'package.json')));
const XLSX = require_(path.join(ROOT, 'lib/xlsx.full.min.js'));

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

/* ══ ★材料★ ══（★式の マスを たくさん 入れます★＝そこが 数に 混ざる 所）
     A列 ... 打ち替える 定数（★お客さんが 打つ マス★）
     B列 ... `=A*2` の 式（★書式つき★＝`#,##0" 円"`）
     C列 ... `=A+1` の 式（書式 無し）
     ⇒★式の 答えが 控えと 違う 形を わざと 作ります★ */
const 行数 = 40;
const 作る = () => {
  const 表 = [['あ', 'い', 'う']];
  for (let i = 0; i < 行数; i++) 表.push([i + 1, null, null]);
  const ws = XLSX.utils.aoa_to_sheet(表);
  for (let i = 0; i < 行数; i++) {
    const r = i + 2;
    /* ★答えを わざと 違えて 保存します★＝計算し直すと 変わる（＝控えと 違う 形） */
    ws['B' + r] = { t: 'n', f: 'A' + r + '*2', v: 999, z: '#,##0" 円"' };
    ws['C' + r] = { t: 'n', f: 'A' + r + '+1', v: 888 };
  }
  ws['!ref'] = 'A1:C' + (行数 + 1);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'あ');
  const 道 = path.join(os.tmpdir(), 'exally-mado-kazu.xlsx');
  fs.writeFileSync(道, XLSX.write(wb, { bookType: 'xlsx', type: 'buffer' }));
  return 道;
};
const 材料 = 作る();

/* ══ ★★当てる 数を 材料から 出します★★ ══（2026-09-25 経営者1 の 注文）
     ＝★「3個」を 決め打ちに すると 材料を 変えた 日に 黙って 嘘に なります★
     ＝★打つ マス（A2）を 指して いる 式を 数えて 1 を 足します★
     ⇒★材料の 指紋（マスの 数・式の 数）も 出しに 書きます★ */
const 打つ番地 = 'A2';   /* ★行1・列0（0から）★ */
const 材料の姿 = (() => {
  const wb = XLSX.read(fs.readFileSync(材料), { type: 'buffer', cellFormula: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  let マス = 0, 式 = 0, つられる = 0;
  Object.keys(ws).forEach((a1) => {
    if (a1.charAt(0) === '!') return;
    マス++;
    const f = ws[a1] && ws[a1].f;
    if (typeof f !== 'string' || f === '') return;
    式++;
    /* ★打つ マスを 名指しで 指して いる 式★（`A2` の 直後が 数字なら 別の 番地） */
    const re = new RegExp(打つ番地 + '(?![0-9])');
    if (re.test(f)) つられる++;
  });
  return { マス: マス, 式: 式, つられる: つられる, 当てる: 1 + つられる };
})();

/** ★本を 開いて 「板名|行,列 → 値」に する★（★借り物で 直に 読む＝画面を 通しません★） */
function 値を並べる(道) {
  const wb = XLSX.read(fs.readFileSync(道), { type: 'buffer', cellFormula: true });
  const 出 = {};
  wb.SheetNames.forEach((な) => {
    const ws = wb.Sheets[な];
    Object.keys(ws).forEach((a1) => {
      if (a1.charAt(0) === '!') return;
      const c = ws[a1];
      const rc = XLSX.utils.decode_cell(a1);
      出[な + '|' + rc.r + ',' + rc.c] = (c && c.v !== undefined && c.v !== null) ? String(c.v) : '';
    });
  });
  return 出;
}

console.log('[窓の 数と 書いた 数] ★1マス 打ったら 窓と 本の 数が 合うか★');
console.log('  ★材料★ ' + path.basename(材料)
  + '（マス ' + 材料の姿.マス + '個 ／ 式 ' + 材料の姿.式 + '個 ／ ★答えを わざと 違えて 保存★）');
console.log('  ★当てる 数★ ... 打つ ' + 打つ番地 + ' 1個 ＋ ★つられる ' + 材料の姿.つられる + '個★'
  + ' ＝ ★' + 材料の姿.当てる + '個★'
  + '（★材料から 出して います＝決め打ちでは ありません★）');
if (材料の姿.つられる < 1) { console.log('  ★★つられる 式が 0個＝この 見張りは 空振りです★★'); process.exit(8); }

const wk = await borrow('mado-kazu', 'webkit');
const browser = await launch('mado-kazu', wk, {}, 'webkit');
const page = await browser.newPage({ viewport: { width: 1280, height: 900 }, acceptDownloads: true });
const 配信 = await 立てる(ROOT);
const 落とし先 = fs.mkdtempSync(path.join(os.tmpdir(), 'exally-mado-'));

try {
  page.on('pageerror', (e) => console.log('      [落ちた] ' + String(e.message).slice(0, 200)));
  await page.goto(配信.url + '/book.html', { waitUntil: 'load', timeout: 120000 });
  await page.evaluate(() => {
    document.body.classList.remove('exally-locked');
    const ov = document.getElementById('loginOv');
    if (ov) { ov.classList.remove('open'); ov.style.display = 'none'; }
  });
  await page.setInputFiles('#bookFileInput', 材料);
  await page.waitForFunction(() => {
    const ss = window.sheets || [];
    return ss.some((s) => s && s.data && Object.keys(s.data).length > 0);
  }, null, { timeout: 120000 });
  await page.waitForTimeout(600);

  /* ★★要る 物が 無ければ 数を 出さずに 止まります★★（2026-09-25 の 決め） */
  const 無い = await page.evaluate(() => ['setCell', 'saveOpenedBook'].filter((n) => typeof window[n] !== 'function'));
  if (無い.length) { console.log('  ★★測れません★★ 在りません ... ' + 無い.join(' / ')); process.exit(8); }

  /* ══ ★1マスだけ 打ちます★ ══（★画面の `setCell`＝お客さんの 道★） */
  await page.evaluate(() => { window.setCell(1, 0, '7'); });
  await page.waitForTimeout(1200);

  /* ══ ★「書き出す」を 押します★ ══（★窓の ボタンまで 押します★） */
  const 落とし物 = page.waitForEvent('download', { timeout: 120000 }).catch(() => null);
  await page.evaluate(() => { window.saveOpenedBook(); });
  await page.waitForTimeout(1500);

  /* ★窓が 出て いるか＝出て いなければ この 見張りは 空振り★ */
  const 窓 = await page.evaluate(() => {
    const ov = document.getElementById('diffOverlay');
    const go = document.getElementById('diffGo');
    return {
      在る: !!(ov && go),
      見えるか: !!(ov && ov.offsetWidth > 0 && ov.offsetHeight > 0),
      字: go ? String(go.textContent || '').trim() : '',
    };
  });
  console.log('      ── 実測 ── 窓 ... ' + (窓.在る ? '在る' : '★在りません★')
    + ' ／ 見えるか ' + 窓.見えるか + ' ／ ボタンの 字 ' + JSON.stringify(窓.字));
  if (!窓.在る || !窓.見えるか) {
    console.log('  ★★窓が 出て いません＝この 見張りは 空振りです★★');
    console.log('  ⇒★1マス 打ったら 窓が 出る はずです★（0か所なら 窓は 出ません）');
    process.exit(8);
  }

  /* ★ボタンの 字から 数を 取ります★（★お客さんが 読む 字★） */
  const 窓の数 = (() => {
    const m = /([0-9][0-9,]*)\s*か所/.exec(窓.字);
    return m ? Number(m[1].replace(/,/g, '')) : null;
  })();
  console.log('      ── 実測 ── ★窓に 出た 数 ... ' + 窓の数 + 'か所★');
  /* ══ ★★正しい 数は 3個です★★ ══（2026-09-25 見込みを 1回 間違えました）
       打った A2 ＋ ★つられた B2（=A2*2）と C2（=A2+1）★ ＝ 3個
       ＝★つられた 分は 書かないと いけません★（実Excel も そう します）
       ＝★はじめ 「1個」に して 赤に しました＝つられる 分を 数えて いませんでした★
       ⇒★★壊れて いる 時は 81個 に なります★★（式 80個 ぜんぶ）
       ⇒★だから 「1」でも 「81」でも なく 「3」で 当てます★ */
  T('★★窓に 出る 数は ' + 材料の姿.当てる + 'か所★★（打った 1個 ＋ つられた ' + 材料の姿.つられる + '個）',
    窓の数 === 材料の姿.当てる, 'ボタンの 字 ' + JSON.stringify(窓.字) + '（★81 なら 控えの 取り直しが 走って いません★）');

  await page.click('#diffGo');
  const dl = await 落とし物;
  if (!dl) {
    console.log('  ★★落とし物が 来ません＝この 見張りは 空振りです★★');
    process.exit(8);
  }
  const 出た = path.join(落とし先, 'out.xlsx');
  await dl.saveAs(出た);
  console.log('      ── 実測 ── 落とし物 ... ' + fs.statSync(出た).size.toLocaleString() + ' バイト');

  /* ══ ★開き直して 値を 突き合わせます★ ══（★実Excel を 使いません★） */
  const 前 = 値を並べる(材料);
  const 後 = 値を並べる(出た);
  const 変わった = [];
  Object.keys(前).forEach((k) => { if (String(前[k]) !== String(後[k] === undefined ? 前[k] : 後[k])) 変わった.push(k); });
  Object.keys(後).forEach((k) => { if (前[k] === undefined) 変わった.push(k + '（増えた）'); });
  console.log('      ── 実測 ── ★本の 中で 変わった 値 ... ' + 変わった.length + '個★'
    + (変わった.length ? '  ' + 変わった.slice(0, 8).join(' / ') : ''));

  T('★★本の 中で 変わった 値は ' + 材料の姿.当てる + '個★★（打った 1個 ＋ つられた ' + 材料の姿.つられる + '個）',
    変わった.length === 材料の姿.当てる, '変わった ' + 変わった.length + '個 ' + 変わった.slice(0, 8).join(' / '));
  T('★★窓の 数と 本の 中で 変わった 数が 同じ★★（★ここが 一番 効きます★）',
    窓の数 === 変わった.length, '窓 ' + 窓の数 + ' ／ 本 ' + 変わった.length);
} finally {
  await page.close().catch(() => {});
  await browser.close().catch(() => {});
  配信.閉じる();
}

console.log('');
console.log('mado-no-kazu-to-kaita-kazu: ' + pass + ' 緑 / ' + fail + ' 赤');
process.exit(fail ? 1 : 0);
