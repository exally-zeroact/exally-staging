/* osu-kami-webkit.mjs -- ★紙の 式を ★お客さんの 道★で まとめて 押す★（2026-09-18）
 *
 *  ★訳★
 *    `tests/shiki-wo-osu-webkit.mjs` は ★14本★です。
 *    ＝★紙に 在る 式の 0.14%★＝★「14本 合った」は「完璧」では ありません★。
 *    ⇒★紙から 機械で 拾って 本数を 増やします★（★手で 書き写しません★）
 *
 *  ★どう 速くするか★
 *    ★1本ごとに 窓と 行き来しない★＝★全部 まとめて 1回で 押します★
 *
 *
 *  ★★★一番 大きい 弱み（★先に 書きます★）★★★
 *    ★材料は 1組だけです★（8枠目・9枠目の 紙の 頭の 物）。
 *    ★でも 紙は 6枚 在り、★紙ごとに 材料が 違います★★。
 *    ⇒★★材料が 違う 紙の 行は「合わない」と 出ます＝★欠陥では ありません★★★
 *    ⇒★今 分かって いる 分（4本）★
 *        =FORECAST(6,B1:B5,A1:A5) / =FORECAST.LINEAR(...) / =LOOKUP(3,A1:A5,B1:B5)
 *          ... その 紙の B は ★2,4,6,8,10★（ここでは B3 が =1/0）
 *        =BYROW(A1:B2,LAMBDA(r,SUM(r)))
 *          ... その 紙の A1:B2 は ★1,2／2,4★
 *    ⇒★★直し方 ... 紙ごとに 材料を 持たせる★★（★まだ して いません★）
 *
 *  ★見て いない 事★
 *    ・★溢れ（スピル）の 2つ目 以降★は 見て いません（左上だけ）
 *    ・★書き出し・読み込みの 道★は 見て いません
 *    ・★紙は 6枚★です（8/9/7枠目・ODDL 6回目・ODDF 部品・ODDF 46個）
 *
 *  使い方:
 *    node <この道具>                      ... 手元
 *    node <この道具> --どこ=<URL>         ... テスト版／本番
 */
import path from 'node:path';
import http from 'node:http';
import fs from 'node:fs';
import { fileURLToPath, pathToFileURL } from 'node:url';

/* ★手元の 絶対の 道を 焼き込まない★（記憶の 決まり） */
const ここ = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(ここ, '..', '..');
const { borrow, launch } = await import(pathToFileURL(path.join(ROOT, 'scripts/_borrow-playwright.mjs')).href);

const 口 = (process.argv.find((a) => a.startsWith('--どこ=')) || '--どこ=手元').split('=').slice(1).join('=');
const 手元か = (口 === '手元');

/* ★8枠目・9枠目の 紙の 頭に 書いて ある 材料★（★既定★） */
const 既定の材料 = [
  ['A1', 1], ['A2', 2], ['A3', 3], ['A4', 4], ['A5', 5],
  ['B1', 1], ['B2', 2], ['B3', '=1/0'], ['B4', 4], ['B5', 5],
  ['C1', 1], ['C2', 3], ['C3', 5], ['C4', 7], ['C5', 9],
  ['D1', 9], ['D2', 7], ['D3', 5], ['D4', 3], ['D5', 1],
  ['F1', 1], ['G1', 2], ['F2', 10], ['G2', 20],
];

/* ★★式の 列・答えの 列は ★紙ごとに 名指しします★★
   ＝★機械に 探させようと しましたが ★9枠目の 紙で 外れました★★
     （並びが「式 ／ ★見込み★ ／ .Value2 ／ 答え」で、★見込みを 答えと 読んで いました★）
   ＝★飾り（★）の 割合で 見分ける 案も 13% で 通って しまいました★
   ⇒★★探す 決まりを 作らず 1枚ずつ 目で 見て 名指しします★★
   ⇒★門★ ... 各紙に ★対照★を 1本 持ち、合わなければ ★その場で 止めます★ */
const 紙たち = [
  { 名: 'golden-kansuu-8kaime-2026-09-18.tsv', 式列: 1, 答列: 2,
    対照: { 式: '=PERMUT(0,0)', 答: '1' } },
  { 名: 'golden-kansuu-9kaime-2026-09-18.tsv', 式列: 2, 答列: 4,
    対照: { 式: '=PERMUT(5,2)', 答: '20' } },
  { 名: 'golden-kansuu-7kaime-2026-09-18.tsv', 式列: 1, 答列: 2,
    対照: { 式: '=PERMUT(5,2)', 答: '20' } },
  { 名: 'golden-oddl-6kaime-2026-09-18.tsv', 式列: 1, 答列: 2,
    対照: { 式: '=ODDLPRICE(DATE(2009,3,10),DATE(2009,8,31),DATE(2008,11,30),0,0.05,100,4,0)',
           答: '97.696364140993609' } },
  { 名: 'golden-oddf-buhin-2026-09-16.tsv', 式列: 3, 答列: 4,
    対照: { 式: '=COUPDAYBS(DATE(2008,11,11),DATE(2021,3,1),2,0)', 答: '70' } },
  { 名: 'golden-oddf-to-46ko-2026-09-16.tsv', 式列: 2, 答列: 3, 対照: null,
    /* ★★この 紙だけ 材料が 違います★★（出どころ ... kansuu46-no-dodai.mjs の 材料()）
       ＝A1:A5 = 1,2,3,4,5 ／ ★B1:B5 = 2,4,6,8,10★
       ＝裏取り ... =FORECAST(6,B1:B5,A1:A5) が 12（★B = 2x でしか 12に ならない★） */
    材料: { B1: 2, B2: 4, B3: 6, B4: 8, B5: 10 } },
  { 名: 'golden-oddf-2kaime-2026-09-16.tsv', 式列: 1, 答列: 2,
    対照: { 式: '=ODDFPRICE(DATE(2009,7,2),DATE(2013,1,1),DATE(2009,1,1),DATE(2010,1,1),0.06,0.05,100,2,1)',
           答: '103.09945989078082' } },
  { 名: 'golden-oddf-3kaime-2026-09-16.tsv', 式列: 1, 答列: 2,
    対照: { 式: '=ODDFPRICE(DATE(2009,1,2),DATE(2013,1,1),DATE(2009,1,1),DATE(2009,7,1),0.06,0.05,100,2,1)',
           答: '-2146826252' } },
  { 名: 'golden-oddf-4kaime-2026-09-17.tsv', 式列: 1, 答列: 2,
    対照: { 式: '=ODDFPRICE(DATE(2009,3,1),DATE(2013,1,1),DATE(2009,1,1),DATE(2009,7,1),0,0.05,100,2,0)',
           答: '-2146826252' } },
  { 名: 'golden-oddf-5kaime-2026-09-17.tsv', 式列: 1, 答列: 2,
    対照: { 式: '=ODDFPRICE(DATE(2009,3,1),DATE(2013,1,1),DATE(2009,1,1),DATE(2009,7,1),0.06,0.03,100,2,0)',
           答: '-2146826252' } },
  /* ★★ここから 下は ★マスを 1つも 指しません★★（材料が 要りません） */
  { 名: 'kansuu46/golden-kane-2026-09-07.tsv', 式列: 1, 答列: 2,
    対照: { 式: '=ACCRINT(DATE(2008,3,1),DATE(2008,8,31),DATE(2008,5,1),0.1,1000,2,0)',
           答: '16.666666666666664' } },
  { 名: 'kansuu46/golden-convert-2026-09-07.tsv', 式列: 1, 答列: 2,
    対照: { 式: '=CONVERT(1,"g","g")', 答: '1' } },
  { 名: 'kansuu46/golden-convert2-2026-09-07.tsv', 式列: 1, 答列: 2,
    対照: { 式: '=CONVERT(1,"kft","ft")', 答: '#N/A' } },
  { 名: 'kansuu46/golden-convert3-2026-09-07.tsv', 式列: 1, 答列: 2,
    対照: { 式: '=CONVERT(1,"kBTU","BTU")', 答: '#N/A' } },
  { 名: 'kansuu46/golden-filterxml-2026-09-07.tsv', 式列: 1, 答列: 2,
    対照: { 式: '=FILTERXML("<a><b>1</b>","//b")', 答: '#VALUE!' } },
  { 名: 'kansuu46/golden-isomitted-2026-09-08.tsv', 式列: 1, 答列: 2,
    対照: { 式: '=LAMBDA(x,y,ISOMITTED(y))(1,2)', 答: 'False' } },
  /* ★★入れなかった 紙と その 訳（★書いて 残します★）★★
     `kansuu46/golden-cell6` `golden-cell7` ... =CELL("format",A1)
        ＝★マスに 付いた 表示形式で 答えが 変わります★＝★板では 作れません★
     `golden-hoyuu-27` ... =A1+A2-0.3 ＝★材料が 要り、1列目が 紙の 名前★
     `golden-86-karimono`（1,998本）`golden-346`（3,593本）
        ＝★マス参照が 多く 材料が 要ります★＝★次に 足します★ */
];


/* ══ ★★関数ごとの 数（★どこが 一番 大きいか を 出す★）★★ ══ */
function 関数名(式) {
  const m = /^=([A-Z0-9_.]+)\(/.exec(式);
  return m ? m[1] : '(不明)';
}

/* ★★真偽の 字だけ 大小を 揃えてから 比べます★★（2026-09-18・Exally1 の 決め）
   ★訳★ ... 紙の 列は `.Value2`＝`False` ／ ★画面に 出るのは `FALSE`★
           ＝★お客さんが 見るのは 画面です★（`.Value2` は 測る 道具の 都合）
   ⇒★台は `FALSE` の まま★／★突き合わせる 側で 揃えます★
   ★★真偽だけです★★＝★他の 字の 大小は 揃えません★（★本物の 違いを 隠さない 為★） */
const 真偽をそろえる = (x) => {
  const s = String(x).trim();
  if (/^(TRUE|FALSE)$/i.test(s)) return s.toUpperCase();
  return s;
};

const 裸 = (s) => String(s === undefined ? '' : s).replace(/★/g, '').replace(/`/g, '').trim();

/* ★実Excel の .Value2 は 誤りを 負の 数で 返します★（記憶の 決まり） */
const 誤りの数 = {
  '-2146826281': '#DIV/0!', '-2146826252': '#NUM!', '-2146826246': '#N/A',
  '-2146826273': '#VALUE!', '-2146826265': '#REF!', '-2146826259': '#NAME?',
  '-2146826288': '#NULL!', '-2146826238': '#CALC!', '-2146826245': '#SPILL!',
};

function 答になるか(a) {
  if (a === '') return false;
  /* ★Excel 自身が 式を 受け付けなかった 行は 答えでは ありません★ */
  if (/打てません|受け付けません|HRESULT/.test(a)) return false;
  return true;
}

const 出し先 = (process.argv.find((a) => a.startsWith('--出し=')) || '').split('=').slice(1).join('=');
const 問い = [];
const 紙ごと = [];
for (const p of 紙たち) {
  const 生 = fs.readFileSync(path.join(ここ, p.名), 'utf-8').replace(/^\uFEFF/, '');
  const 行たち = 生.split(/\r?\n/)
    .filter((l) => l && !l.startsWith('#') && l.includes('\t'))
    .map((l) => l.split('\t'));
  /* ★★門 ... 対照が 合わなければ 列が ずれて います★★ */
  if (p.対照) {
    const r = 行たち.find((c) => 裸(c[p.式列]) === p.対照.式);
    if (!r) { console.log('  NG   ★' + p.名 + ' に 対照の 式が ありません★'); process.exit(9); }
    if (裸(r[p.答列]) !== p.対照.答) {
      console.log('  NG   ★' + p.名 + ' の 答えの 列が ずれて います★'
        + ' ... 対照 ' + p.対照.式 + ' の 答えは ' + p.対照.答
        + ' の はずが [' + 裸(r[p.答列]) + ']');
      process.exit(9);
    }
  }
  let n = 0;
  for (const c of 行たち) {
    const 式 = 裸(c[p.式列]);
    let 答 = 裸(c[p.答列]);
    if (!式.startsWith('=')) continue;
    if (!答になるか(答)) continue;
    if (誤りの数[答]) 答 = 誤りの数[答];
    問い.push({ 紙: p.名, 材料: p.材料 || null, 式, 答 });
    n += 1;
  }
  紙ごと.push(p.名 + ' ... ' + n + '本（式 ' + p.式列 + '列目 ／ 答え ' + p.答列 + '列目'
    + (p.対照 ? ' ／ ★対照 ok★' : ' ／ ★対照 なし★') + '）');
}

console.log('');
console.log('[osu-kami-webkit] ★紙の 式を お客さんの 道で 押す★');
console.log('  ★どこ★ ... ' + (手元か ? '手元' : 口));
for (const r of 紙ごと) console.log('  ★紙★ ... ' + r);
console.log('  ★★拾った 式 ... ' + 問い.length + '本★★（★これが 分母★）');

function 立てる(root) {
  const 型 = {
    '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8',
    '.css': 'text/css; charset=utf-8', '.json': 'application/json; charset=utf-8',
    '.png': 'image/png', '.svg': 'image/svg+xml', '.ico': 'image/x-icon',
  };
  const s = http.createServer((q, r) => {
    const f = path.join(root, decodeURIComponent(String(q.url).split('?')[0]).replace(/^\/+/, ''));
    if (!f.startsWith(root) || !fs.existsSync(f) || fs.statSync(f).isDirectory()) { r.statusCode = 404; return r.end('no'); }
    r.setHeader('content-type', 型[path.extname(f).toLowerCase()] || 'application/octet-stream');
    fs.createReadStream(f).pipe(r);
  });
  return new Promise((x) => s.listen(0, '127.0.0.1', () => x({ url: 'http://127.0.0.1:' + s.address().port, 閉じる: () => s.close() })));
}

const 時計 = Date.now();
const wk = await borrow('osu-kami', 'webkit');
const browser = await launch('osu-kami', wk, {}, 'webkit');
const page = await browser.newPage({ viewport: { width: 1280, height: 800 } });

/* ★★窓の 誤りを 拾います★★（2026-09-18・★繋ぐ 時の 門★）
   ＝★読み込んだ だけで 投げる 物は ここにしか 出ません★
     （例）`process.env` を ブラウザで 読む ／ `root.Bahttext` が `undefined`
   ＝★node では 1件も 出ません★＝★だから 窓で 拾います★ */
/* ★★元から 在る 誤り（★名指しで 許す★）★★
   ＝2026-09-18 に ★本番（exally.vercel.app）で 実測★した 1件
   ＝`<meta name="viewport">` の `interactive-widget` を WebKit が 知らない だけ
   ＝★計算にも 画面にも 出ません★／★Chrome では 出ません★
   ⇒★★繋いだ 後に 増えた 誤りだけを 見る 為に 名指しで 除きます★★
   ⇒★消えたら 赤に します★（＝許しを 外す） */
const 元から在る誤り = [
  'Viewport argument key "interactive-widget" not recognized',
  /* ★2026-09-18 ... ★仮の 配信（preview）だけに 出ます★
     ＝Vercel が 足す `vercel.live/_next-live/feedback/feedback.js` の 中
     ＝★本番（exally.vercel.app）には この 行が 在りません★（実測 0件）
     ★裏取り★ ... `test` 枝（★台を 読み込んで いない★）でも ★同じ 1件★
       ⇒★★台を 足した せいでは ありません★★
       ⇒★危うく「㋐が 誤りを 増やした」と 出す 所でした★
     ＝★うちの 字では ありません★／★お客さんには 出ません★ */
  "navigator.storage.persisted",
];
const 窓の誤り = [];
const 元から在る = [];
const 誤りを分ける = (s2) => {
  if (元から在る誤り.some((x) => s2.indexOf(x) >= 0)) 元から在る.push(s2);
  else 窓の誤り.push(s2);
};
page.on('pageerror', (e) => 誤りを分ける('pageerror: ' + String(e.message).slice(0, 160)));
page.on('console', (m) => {
  if (m.type() === 'error') 誤りを分ける('console.error: ' + String(m.text()).slice(0, 160));
});
const 配信 = 手元か ? await 立てる(ROOT) : null;
const 住所 = 手元か ? (配信.url + '/book.html') : 口;
let 終わり = 1;
try {
  const 返 = await page.goto(住所, { waitUntil: 'load', timeout: 60000 }).catch((e) => ({ エラー: e.message }));
  if (!返 || 返.エラー || (返.status && 返.status() >= 400)) {
    console.log('  NG   ★開けませんでした★ ' + (返 && 返.エラー ? 返.エラー : ('http ' + (返 && 返.status && 返.status()))));
    throw new Error('★開けないので 止めます★');
  }
  console.log('  ok   ★開けた★ http ' + 返.status());
  await page.evaluate(() => {
    document.body.classList.remove('exally-locked');
    const ov = document.getElementById('loginOv');
    if (ov) { ov.classList.remove('open'); ov.style.display = 'none'; }
  });
  const 読み = await page.evaluate(() => {
    const s = Array.from(document.querySelectorAll('script[src]')).map((x) => x.getAttribute('src') || '');
    return {
      全: s.length,
      hf: s.filter((u) => /hyperformula/i.test(u)).length,
      shiki: s.filter((u) => /shiki-/.test(u)).length,
    };
  });
  console.log('  ★script src ' + 読み.全 + '本 ／ hyperformula ' + 読み.hf + '本 ／ shiki- ' + 読み.shiki + '本★');
  /* ★★頁の 大きさと 開くまでの 秒★★（★18本 増えると ここに 出ます★） */
  const 大きさ = await page.evaluate(() => {
    const n = performance.getEntriesByType('navigation')[0];
    const r = performance.getEntriesByType('resource');
    return {
      頁: n ? Math.round(n.transferSize || 0) : 0,
      全部: r.reduce((a, x) => a + (x.transferSize || 0), (n ? (n.transferSize || 0) : 0)),
      本数: r.length,
      秒: n ? Math.round(n.loadEventEnd) : 0,
    };
  });
  console.log('  ★頁 ' + 大きさ.頁 + ' バイト ／ 全部で ' + 大きさ.全部 + ' バイト（' + 大きさ.本数 + '本）'
    + ' ／ 開くまで ' + 大きさ.秒 + ' ミリ秒★');

  /* ★★1回で まとめて 押します★★（★紙ごとに 材料を 入れ直します★） */
  const 出 = await page.evaluate(({ 材, 問 }) => {
    const 番地 = (s) => {
      const m = /^([A-Z]+)(\d+)$/.exec(s);
      let c = 0;
      for (const ch of m[1]) c = c * 26 + (ch.charCodeAt(0) - 64);
      return { r: Number(m[2]) - 1, c: c - 1 };
    };
    const sh = window.sheets[window.activeSheet];
    const 材を入れる = (上書き) => {
      for (const kv of 材) { const a = 番地(kv[0]); window.setCell(a.r, a.c, kv[1]); }
      if (上書き) for (const k of Object.keys(上書き)) { const a = 番地(k); window.setCell(a.r, a.c, 上書き[k]); }
    };
    const 答 = [];
    const 行 = 20;              /* ★材料の 下に 置きます★ */
    let 今の紙 = null;
    for (let i = 0; i < 問.length; i++) {
      if (問[i].紙 !== 今の紙) { 今の紙 = 問[i].紙; 材を入れる(問[i].材料); }
      let v;
      try {
        window.setCell(行, 9, 問[i].式);
        if (typeof window.recalcSheet === 'function') window.recalcSheet(window.activeSheet, sh.data);
        const c = (sh.data || {})[行 + ',9'];
        v = !c ? null : (c.d !== undefined ? c.d : c.v);
      } catch (e) { 答.push({ 転: String(e.message).slice(0, 60) }); continue; }
      答.push({ 値: (v === undefined || v === null) ? '' : String(v) });
    }
    return 答;
  }, { 材: 既定の材料, 問: 問い.map((q) => ({ 紙: q.紙, 材料: q.材料, 式: q.式 })) });

  let 合 = 0, 違 = 0, 空 = 0, 転 = 0;
  const 外れ = [];
  const 関数ごと = new Map();
  let 日付を渡すと = 0;   /* ★`DATE(` が 在り 出たのが #VALUE! の 本数★ */
  const 印 = (式, どう) => {
    const n = 関数名(式);
    if (!関数ごと.has(n)) 関数ごと.set(n, { 合: 0, 違: 0 });
    関数ごと.get(n)[どう] += 1;
  };
  for (let i = 0; i < 問い.length; i++) {
    const q = 問い[i];
    const o = 出[i] || {};
    if (o.転) { 転 += 1; 外れ.push(q.式 + ' => ★転んだ★ ' + o.転); continue; }
    if (o.値 === '' || o.値 === undefined) { 空 += 1; 外れ.push(q.式 + ' => ★空っぽ★（はず ' + q.答 + '）'); continue; }
    const 数どうし = /^[-+]?[0-9]*[.]?[0-9]+([eE][-+]?[0-9]+)?$/.test(q.答) && /^[-+]?[0-9]*[.]?[0-9]+([eE][-+]?[0-9]+)?$/.test(o.値);
    const 同 = (真偽をそろえる(o.値) === 真偽をそろえる(q.答))
      || (数どうし && Math.abs(Number(o.値) - Number(q.答)) <= Math.max(1e-9, Math.abs(Number(q.答)) * 1e-9));
    if (同) { 合 += 1; 印(q.式, '合'); }
    else {
      違 += 1; 印(q.式, '違');
      if (q.式.indexOf('DATE(') >= 0 && String(o.値).trim() === '#VALUE!') 日付を渡すと += 1;
      外れ.push(q.式 + ' => 出た [' + o.値 + '] ／実Excel [' + q.答 + ']'); }
  }
  console.log('');
  console.log('  ★★合った ' + 合 + ' / ' + 問い.length + '★★ ／ 違った ' + 違
    + ' ／ ★空っぽ ' + 空 + '★ ／ 転んだ ' + 転);
  console.log('  ★合った 割合 ... ' + (問い.length ? (100 * 合 / 問い.length).toFixed(1) : '0') + '%★');
  if (出し先) {
    const 行ごと = 問い.map((q, i) => q.式 + '\t' + String((出[i] || {}).値 === undefined ? '(空)' : (出[i] || {}).値) + '\t' + q.答);
    fs.writeFileSync(出し先, 行ごと.join('\n') + '\n', 'utf-8');
    console.log('  ★1行ずつ 出しました ... ' + 出し先 + '（' + 行ごと.length + '行）★');
  }

  /* ══ ★★1本の 式が 後ろを 止めないか★★（2026-09-18・Exally1 の 注文）
       ★訳★ ... 借り物が `Value of the formula cell is not computed.` を ★投げます★
                `recalcSheet` が 受けて いないと ★そこで 繰り返しが 止まり★
                ★後ろの マスが 全部 古い まま★に なります。
                ⇒★誤りに ならない／画面も 崩れない／数だけ 違う★
       ★★私の 分母（紙 2,116本）では 出ません★★
         ＝1本ずつ 同じ マスに 打って いるから
         ＝★同じ 板に 2本 並べないと 出ません★
       ⇒★だから 名指しの 対照を 1組 置きます★ */
  /* ★★受け止めます★★＝★測る 道具が 測る 相手で 死んでは いけません★
     （2026-09-18 ... ★この 対照を 足した その日に 私が 死にました★） */
  /* ★★2026-09-19 足した ── ★裏取りを その場で 押す★★★
     ＝下の「DATE() を 渡すと #VALUE!」の 裏取りは ★字を 焼き込んで いました★
     ＝お金の 直し（`包みを外す`）が 入った 後も ★毎回 同じ 嘘を 印字して いました★
     ⇒★★測り道具が 自分で 押して 自分で 読む★★＝★飾りの 字に 頼らない★
     ⇒★合う／合わないの 判じは しません★＝★出た 字を そのまま 並べます★ */
  let 止める式 = { 投げる: '(押せず)', 次: '(押せず)', その次: '(押せず)', 投げた: null,
    日付渡し: '(押せず)', 数渡し: '(押せず)' };
  try {
    止める式 = await page.evaluate(() => {
    const sh = window.sheets[window.activeSheet];
    window.setCell(60, 9, '=CONVERT(,"m","cm")');   /* ★投げる 式★ */
    window.setCell(61, 9, '=1+1');                  /* ★その 後ろ★ */
    window.setCell(62, 9, '=SUM(1,2,3)');           /* ★もっと 後ろ★ */
    /* ★裏取り★ ... ★同じ 日・同じ 関数・渡し方だけ 違う 2本★ */
    window.setCell(63, 9, '=COUPDAYS(DATE(2008,11,11),DATE(2021,3,1),2,0)');
    window.setCell(64, 9, '=COUPDAYS(39763,44256,2,0)');
    if (typeof window.recalcSheet === 'function') window.recalcSheet(window.activeSheet, sh.data);
    const 読む = (r) => {
      const c = (sh.data || {})[r + ',9'];
      if (!c) return '(マスが 無い)';
      return String(c.d !== undefined ? c.d : c.v);
    };
      return { 投げる: 読む(60), 次: 読む(61), その次: 読む(62), 投げた: null,
        日付渡し: 読む(63), 数渡し: 読む(64) };
    });
  } catch (e) {
    止める式.投げた = String(e.message).slice(0, 120);
  }
  const 止まっていない = (止める式.次 === '2' && 止める式.その次 === '6');
  console.log('');
  console.log('  ★★1本の 式が 後ろを 止めて いないか★★');
  console.log('    `=CONVERT(,"m","cm")` ... ' + 止める式.投げる + '（実Excel `#N/A`）');
  console.log('    その 後ろ `=1+1` ..... ' + 止める式.次 + '（★2 の はず★）');
  console.log('    もっと 後ろ `=SUM(1,2,3)` ... ' + 止める式.その次 + '（★6 の はず★）');
  if (止める式.投げた) {
    console.log('    ★★窓の 中で 投げました★★ ... ' + 止める式.投げた);
    console.log('    ⇒★★これ 自体が「止まって いる」印です★★');
  }
  console.log('    ⇒ ' + (止まっていない
    ? '★止まって いません★'
    : '★★止まって います＝後ろの マスが 古い まま です★★'));

  console.log('');
  console.log('  ★★窓の 誤り（★新しい 物だけ★） ... ' + 窓の誤り.length + '件★★'
    + '（★読み込んだ だけで 投げる 物は ここにしか 出ません★）');
  console.log('    ★元から 在る（名指しで 許した） ... ' + 元から在る.length + '件★');
  if (元から在る.length === 0) {
    console.log('    ★★許した 物が 1件も 出ませんでした＝許しを 外して ください★★');
  }
  for (const e of 窓の誤り.slice(0, 12)) console.log('    ・' + e);
  if (窓の誤り.length > 12) console.log('    ...（残り ' + (窓の誤り.length - 12) + '件）');
  console.log('');
  console.log('  ★★★DATE() を 渡すと #VALUE! ... ' + 日付を渡すと + '本★★★'
    + '（★合わない ' + 外れ.length + '本の うち★）');
  /* ★★ここは ★この 配信で 今 押した 字★です（★焼き込みでは ありません★）★★ */
  console.log('    ★裏取り（★この 配信で 今 押しました★）★');
  console.log('      日付で 渡す =COUPDAYS(DATE(2008,11,11),DATE(2021,3,1),2,0) ... '
    + 止める式.日付渡し);
  console.log('      数で 渡す   =COUPDAYS(39763,44256,2,0) ................... '
    + 止める式.数渡し + '（実Excel ★180★）');
  /* ★★2026-09-19 ── ★足した その場で わざと 壊して 出た 欠陥★★★
     ＝`test` 枝（★直しが 入って いない★）で 押したら ★窓が 投げて 両方 `(押せず)`★
     ＝それを 見比べて ★「変わりません」＝偽の 緑★ を 出しました
     ⇒★★押せたかを ★先に★ 見る★★＝★「同じ」より 先に「読めたか」★ */
  const 押せた = (x) => x !== '(押せず)' && x !== '(マスが 無い)';
  console.log('    ⇒ ' + (!押せた(止める式.日付渡し) || !押せた(止める式.数渡し)
    ? '★★裏取りが 押せて いません＝この 2行は 何も 言って いません★★'
      + (止める式.投げた ? '（窓が 投げた ... ' + 止める式.投げた + '）' : '')
    : 止める式.日付渡し === 止める式.数渡し
      ? '★★渡し方で 変わりません★★'
      : '★★同じ 日・同じ 関数なのに ★渡し方だけ★で 変わって います★★'));
  console.log('');
  console.log('  ★★関数ごと ... 合わない 数が 多い 順（上 20）★★');
  console.log('    関数              合った  合わない');
  const 並び = [...関数ごと.entries()]
    .map(([n, v]) => ({ n, ...v }))
    .filter((x) => x.違 > 0)
    .sort((a, b) => b.違 - a.違)
    .slice(0, 20);
  for (const x of 並び) {
    console.log('    ' + x.n.padEnd(18) + String(x.合).padStart(5) + String(x.違).padStart(10));
  }
  console.log('');
  console.log('  ★★合わない 全部（' + 外れ.length + '本）★★');
  for (const s of 外れ.slice(0, 40)) console.log('    ・' + s);
  if (外れ.length > 40) console.log('    ...（★残り ' + (外れ.length - 40) + '本は 切りました★）');
  終わり = 0;
} finally {
  await browser.close().catch(() => {});
  if (配信) 配信.閉じる();
  console.log('');
  console.log('  ★掛かった 秒 ... ' + ((Date.now() - 時計) / 1000).toFixed(1) + '秒★');
}
process.exit(終わり);
