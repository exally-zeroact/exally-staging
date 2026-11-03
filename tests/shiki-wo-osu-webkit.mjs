/* shiki-wo-osu-webkit.mjs — ★お客さんと 同じ 道で 式を 打って 答えを 読む★（2026-09-18）
 *
 *  ★★なぜ 要るか（★決まりです★）★★
 *    `team/global-rules.md` §11 ②
 *      「★Claude Code が 実 UI で 全ボタン・全パターンを 操作する★」
 *      「★計算 lib 緑＝『ボタンが 押せる／画面が 動く／配線されてる』は 保証しない★」
 *    ⇒★★これまで「マスに 式を 打って 答えを 読む」は ★0本★でした★★
 *    ⇒★node で 出した 数は 画面の 数では ありません★（記憶の 決まり）
 *
 *  ★★★今 これが 測って いる のは「借り物」です★★★（2026-09-18 実測）
 *    ★画面の `shiki-` は ★0本★★＝★私が 書いた 台は 1つも 読まれて いません★
 *    ⇒★★だから ここで 合う 物は ★借り物（HyperFormula ＋ JS層）が 答えて います★★★
 *    ⇒★★「私が 今日 書いた 13個が お客さんに 届いたか」は ★まだ 0本 測れて いません★★★
 *
 *  ★★★繋ぐ 前の 姿（★これを 守ります★）★★★
 *    ★14本 中 12本 合う／★2本 違う★★
 *      `=PERCENTRANK({1;3;5;7;9},4)`   … 画面 ★0.5★／★実Excel 0.375★／うちの台 ★0.375★
 *      `=AGGREGATE(19,6,{1;2;4;5},1)`  … 画面 ★#VALUE!★／★実Excel 1.25★／うちの台 ★1.25★
 *    ⇒★★借り物が 2本 間違えて います＝★客に 出る 欠陥★★★
 *    ⇒★★繋げば 直る はずの 2本です★★
 *
 *  ★★この 見張りの 使い方★★
 *    既定 …………… ★繋ぐ 前の 姿と 同じか★（★今は 緑★）
 *    `--終わりの線` … ★14本 とも 合う ＋ hyperformula 0本★（★今は 赤＝当たり前★）
 *    ⇒★★繋いだら 既定が 赤に なります＝それが 狙いです★★
 *
 *  ★★どこを 開くかを 口（引数）で 選べます★★
 *    `--どこ=手元`     … 手元の ファイルを 小さい 配信で 出す（★既定★）
 *    `--どこ=<URL>`    … テスト版／本番 の 住所を そのまま
 *    ⇒★★同じ 道具で 3つとも 押せます★★＝★「テスト版だけ 違う」が すぐ 出ます★
 *    ⇒★鍵が 要るのは ★最後の 1回だけ★に なります★
 *
 *  ★★出す 物★★
 *    ・★分母つき★（◯ / ◯本）
 *    ・★空っぽの 控え★（★答えが 空の 物を 数える＝0件を 緑に しない★）
 *    ・★掛かった 秒★
 *    ・★`hyperformula` が 何本 読まれたか★／★`shiki-` が 何本★（★終わりの 線★）
 *
 *  ★★門★★
 *    ①★開けなかったら「開けなかった」と 書く★（★0件を 緑に しない★）
 *    ②★本数 決め打ち★（式の 数が 変わったら 赤）
 *    ③★1本ずつ 受け止める★（1本 転んでも 全部 止めない）
 *    ④★★読み込んだ 本数を 画面から 数える★★（★私の 見立てでは ない★）
 *    ⑤★秒が 毎回 同じなら 上限か 空振りを 疑う★
 *
 *  走らせ方:
 *    node tests/shiki-wo-osu-webkit.mjs
 *    node tests/shiki-wo-osu-webkit.mjs --どこ=https://……/book.html
 */
import path from 'node:path'; import http from 'node:http'; import fs from 'node:fs';
import { fileURLToPath } from 'node:url';
import { borrow, launch } from '../scripts/_borrow-playwright.mjs';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
let pass = 0, fail = 0;
const T = (n, よい, 添え) => {
  if (よい) { pass++; console.log('  ok   ' + n); }
  else { fail++; console.log('  NG   ' + n + (添え ? '\n       ' + 添え : '')); }
};

/* ★どこを 開くか★ */
const 口 = (process.argv.find((a) => a.startsWith('--どこ=')) || '--どこ=手元').split('=').slice(1).join('=');
const 手元か = (口 === '手元');

/* ══ ★★押す 式（★ここが 分母★）★★ ══
   ★今日 書いた 13個を 中心に、★お客さんの 道で 本当に 動くか★ を 見ます★
   ★答えは ★字で★ 突き合わせます★（★出た 字が 同じ＝同じ★） */
const 式たち = [
  { 式: '=SUM(1,2,3)',                 答: '6',       訳: '★対照★ 前から 動く 物' },
  { 式: '=PERMUT(5,2)',                答: '20',      訳: '順列（09-18）' },
  { 式: '=PERMUTATIONA(5,2)',          答: '25',      訳: '順列（09-18）' },
  { 式: '=XMATCH(5,{1;3;5;7;9})',      答: '3',       訳: 'XMATCH（09-18）' },
  { 式: '=XMATCH(4,{1;3;5;7;9},-1)',   答: '2',       訳: 'XMATCH 次に小さい' },
  { 式: '=PERCENTRANK({1;3;5;7;9},4)', 答: '0.375',   訳: 'PERCENTRANK（09-18）' },
  { 式: '=UNICODE(ASC(UNICHAR(65313)))', 答: '65',    訳: 'ASC（09-18）' },
  { 式: '=UNICODE(DBCS(UNICHAR(65)))', 答: '65313',   訳: 'DBCS（09-18）' },
  { 式: '=LENB(UNICHAR(12354))',       答: '2',       訳: 'LENB（09-18）' },
  { 式: '=LEN(MIDB(UNICHAR(12354)&"A",2,2))', 答: '2', 訳: '★外した 1本★ MIDB' },
  { 式: '=TEXTAFTER("a-b-c","-")',     答: 'b-c',     訳: 'TEXTAFTER（09-18）' },
  { 式: '=TEXTBEFORE("a-b-c","-",2)',  答: 'a-b',     訳: 'TEXTBEFORE（09-18）' },
  { 式: '=AGGREGATE(9,6,{1;2;4;5})',   答: '12',      訳: 'AGGREGATE（09-18）' },
  { 式: '=AGGREGATE(19,6,{1;2;4;5},1)', 答: '1.25',   訳: '★外した 1本★ 機能19＝EXC' },
];
const 式の本数 = 14;

/* ══ ★★繋ぐ 前の 姿（★2026-09-18 に 実測して 固定★）★★ ══
     ★枝の 頭★ 172247a ／ ★刻印★ ?v=17fe12a2
     ★訳★ … ★後から 辻褄を 合わせられない 形に します★
     ★借り物（HyperFormula ＋ JS層）が 答えた 姿★です */
const 繋ぐ前 = {
  合: 12,
  違: 2,
  shiki: 0,
  違う式: ['=PERCENTRANK({1;3;5;7;9},4)', '=AGGREGATE(19,6,{1;2;4;5},1)'],
};
/* ══ ★★㋐（読み込むだけ）の 後の 姿★★ ══（2026-09-18・★実測して 固定★）
     ★★㋐で 答えが 変わったら 赤★★＝★読み込むだけ なのに 変わるのは おかしい★
     ★`shiki-` は 9本★（`lib/bessel.js` は 名前が 違うので 別に 数えます）
     ★`hyperformula` は 1本の まま★＝★外すのは また 別の 1押し★
     ★★なぜ 2つ 持つか★★
       経営者1 の 門は「`shiki-` が 1本 以上 なら 緑」でした
       ⇒★★㋐だけで 緑に なります／でも 答えは 1つも 変わって いません★★
       ⇒★記憶「入れて 落ちないかで 測るな＝引いて 正しい 答えが 出るかで 測れ」★
       ⇒★だから ★答えの 姿★を 決め打ちに します★ */
const 読み込んだ後 = {
  合: 12,
  違: 2,
  shiki: 9,
  hyperformula: 1,
};
/* == ★★㋑（計算を 台に 回した）後の 姿★★ ==（2026-09-18・★実測して 固定★）
     ★ここから 答えが 変わります★＝★㋐との 違いは そこです★
     ★`shiki-` が 10本★（`lib/shiki-ita-awase.js` が 増えた＝★板の 写しを 合わせる 手★）
     ★`hyperformula` は 1本の まま★＝★外すのは また 別の 1押し★
     ★★門の 向きを 変えました★★（2026-09-18）
       前 ... 「★繋ぐ 前の 姿と 同じ★」（★繋いだら 赤に なる★＝手を 止める 門）
       今 ... 「★台に 回した 後の 姿と 同じ★」＋★下がって いない★
       ⇒★向きを 変える 時も わざと 壊して 確かめました★
         （`--わざと壊す` で 決め打ちを 1つ ずらすと 赤に なります）
     ★★下がる 0本★★＝★合った 本数が ★㋐の 12 より 減ったら 赤★★ */
const 台に回した後 = {
  合: 14,
  違: 0,
  shiki: 10,
  hyperformula: 1,
  違う式: [],
};
/* ★門を わざと 壊す★（★足しただけでは 効かない＝門は 引き継がれない★） */
if (process.argv.includes('--わざと壊す')) {
  台に回した後.合 = 13;
  console.log('  ★★わざと 壊して います★★（決め打ちの 合を 14 -> 13 に しました）');
}
const 終わりの線か = process.argv.includes('--終わりの線');

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
  return new Promise((x) => s.listen(0, '127.0.0.1', () => x({ url: 'http://127.0.0.1:' + s.address().port, 閉じる: () => s.close() })));
}

console.log('');
console.log('[shiki-wo-osu-webkit] ★お客さんと 同じ 道で 式を 打って 答えを 読む★');
console.log('  ★どこ★ … ' + (手元か ? '手元（file を 小さい 配信で 出す）' : 口));

/* ★②本数の 門★ */
if (式たち.length !== 式の本数) {
  console.log('  NG   ★' + 式の本数 + '本の はずが ' + 式たち.length + '本です★');
  process.exit(4);
}
console.log('  ★押す 式 … ' + 式たち.length + '本★（決め打ち ' + 式の本数 + '本）');

const 時計 = Date.now();
const wk = await borrow('shiki-wo-osu', 'webkit');
const browser = await launch('shiki-wo-osu', wk, {}, 'webkit');
const page = await browser.newPage({ viewport: { width: 1280, height: 800 } });

/* ══ ★★窓の 誤りを 数える★★ ══（2026-09-18・経営者1 の 注文）
     ★なぜ★ ... ★読み込んだ だけで 投げる 物は ★ここにしか 出ません★★
       ＝node の 試験は 1件も 出しません（`require` は 通る）
       ＝2026-09-18 に 私が 直した 3つ（`process.env` / `Bahttext` / `FormulaFilterxml`）は
         ★全部 ここに 出る 形★でした
     ★元から 在る 物は 名指しで 許します★（★数で 許すと 新しい 物が 隠れます★） */
const 許す誤り = [
  /* ★`<meta viewport>` の 鍵を WebKit が 知らない だけ★＝★計算にも 画面にも 出ません★
     ＝経営者1 が 本番でも 同じ 1件を 見て います（2026-09-18） */
  'interactive-widget',
];
const 窓の誤り = [];
page.on('pageerror', (e) => 窓の誤り.push('pageerror: ' + String(e && e.message).slice(0, 120)));
page.on('console', (m) => {
  if (m.type() !== 'error') return;
  窓の誤り.push('console.error: ' + String(m.text()).slice(0, 120));
});
const 配信 = 手元か ? await 立てる(ROOT) : null;
const 住所 = 手元か ? (配信.url + '/book.html') : 口;
let 開けた = false;
try {
  /* ★①開けなかったら「開けなかった」と 書く★ */
  const 返 = await page.goto(住所, { waitUntil: 'load', timeout: 60000 }).catch((e) => ({ エラー: e.message }));
  if (!返 || 返.エラー || (返.status && 返.status() >= 400)) {
    T('★開けた★', false, '★開けませんでした★ ' + (返 && 返.エラー ? 返.エラー : ('http ' + (返 && 返.status && 返.status()))));
    throw new Error('★開けないので ここで 止めます★（★0件を 緑に しません★）');
  }
  開けた = true;
  T('★開けた★', true, 'http ' + 返.status());

  /* ★鍵を 外す（ログインは 通して いません）★ */
  await page.evaluate(() => {
    document.body.classList.remove('exally-locked');
    const ov = document.getElementById('loginOv');
    if (ov) { ov.classList.remove('open'); ov.style.display = 'none'; }
  });

  /* ★④読み込んだ 本数を ★画面から★ 数える★（★私の 見立てでは ない★） */
  const 読み = await page.evaluate(() => {
    const s = Array.from(document.querySelectorAll('script[src]')).map((x) => x.getAttribute('src') || '');
    return {
      全: s.length,
      hf: s.filter((u) => /hyperformula/i.test(u)).length,
      shiki: s.filter((u) => /\/shiki-|^shiki-|lib\/shiki-/.test(u)).length,
    };
  });
  console.log('  ★画面が 読み込んだ script src … ' + 読み.全 + '本★');
  console.log('    ★hyperformula … ' + 読み.hf + '本★（★終わりの 線＝0本★）');
  console.log('    ★shiki- ……… ' + 読み.shiki + '本★（★終わりの 線＝1本 以上★）');

  T('★式を 打つ 口が 在る★', await page.evaluate(() => typeof window.setCell === 'function'));

  /* == ★★JS層が 投げた 時に 後ろの マスが 止まらないか★★ ==（2026-09-18）
       `=CONVERT(,"m","cm")`（★1つ目の 引数が 空★）
         ★JS層★ ... 借り物が `Value of the formula cell is not computed.` を ★投げます★
         ★前★ ... `recalcSheet` が 受けて いなかった
                 ⇒★繰り返しが 途中で 止まり／後ろの マスが 古い まま★
                 ⇒★画面は 空★（実Excel は `#N/A`）
         ★今★ ... ★投げたら 「JS層は 答えない」★ ⇒ 台が `#N/A`
       ★★後ろの マスも 一緒に 見ます★★＝★止まって いない 事を 数で 出します★ */
  {
    const 出 = await page.evaluate(() => {
      const sh = window.sheets[window.activeSheet];
      const 段 = [];
      const 試 = (名, f) => { try { f(); 段.push(名 + ':ok'); } catch (e) { 段.push(名 + ':NG ' + String(e && e.message).slice(0, 80)); } };
      試('setCell(CONVERT)', () => window.setCell(60, 10, '=CONVERT(,"m","cm")'));
      試('setCell(2)', () => window.setCell(61, 10, '2'));
      試('setCell(K62*3)', () => window.setCell(62, 10, '=K62*3'));
      試('recalcSheet', () => { if (typeof window.recalcSheet === 'function') window.recalcSheet(window.activeSheet, sh.data); });
      const 読 = (r, c) => {
        const x = (sh.data || {})[r + ',' + c];
        if (!x) return '(無い)';
        const v = (x.d !== undefined ? x.d : x.v);
        return (v === undefined || v === null) ? '' : String(v);
      };
      return { 投げる: 読(60, 10), 後ろ: 読(62, 10), 段: 段.join(' | ') };
    });
    console.log('  ★=CONVERT(,"m","cm") ... "' + 出.投げる + '"（★#N/A が 正★）'
      + ' ／ ★後ろの マス "' + 出.後ろ + '"（★6 が 正★）');
    console.log('     ★段★ ' + 出.段);
    T('★★投げる 式が 実Excel と 同じ #N/A★★（★前は 空★）',
      出.投げる === '#N/A', '出た "' + 出.投げる + '"');
    T('★★投げても 後ろの マスが 止まらない★★（=K62*3 ⇒ 6）',
      出.後ろ === '6', '出た "' + 出.後ろ + '"');
    T('★★どの 段も 投げない★★（★`recalcSheet` が 途中で 止まらない★）',
      出.段.indexOf(':NG') < 0, 出.段);
  }

  /* == ★★2つの 道具が 割れた 7個★★ ==（2026-09-19）
       ★私の 門（`tsunagu-mon`）★ ... 台に 無い ★25個★（★台が 名前を 持つか★）
       ★経営者1 の 道具★ ......... 台に 無い ★27個★（★製品が 答えるか★）
       ★差の 中身★ ... ★台は 知って いるのに 向こうの 道具は 「無い」と 言う 7個★
         BYCOL BYROW LAMBDA LET AREAS BINOM.DIST.RANGE PHONETIC
       ⇒★★どちらが 正しいかは ★お客さんの 画面★で 決まります★★
       ⇒★だから ここで 押します★（★名簿の 突き合わせでは 決めない★）
       ★答えは 実Excel で 確かめた 物だけ 決め打ちに して います★ */
  {
    const 出 = await page.evaluate(() => {
      const sh = window.sheets[window.activeSheet];
      const 押 = (r, f) => {
        window.setCell(r, 11, f);
        return null;
      };
      window.setCell(50, 10, '1'); window.setCell(50, 11, '2');
      window.setCell(51, 10, '10'); window.setCell(51, 11, '20');
      const 組 = [
        ['LET', '=LET(x,2,x*3)'],
        ['LAMBDA', '=LAMBDA(x,x*2)(4)'],
        ['AREAS', '=AREAS(K51:L52)'],
        ['BINOM.DIST.RANGE', '=BINOM.DIST.RANGE(10,0.5,3)'],
        ['PHONETIC', '=PHONETIC(K51)'],
        ['BYROW', '=BYROW(K51:L52,LAMBDA(r,SUM(r)))'],
        ['BYCOL', '=BYCOL(K51:L52,LAMBDA(c,SUM(c)))'],
      ];
      const 結 = {};
      for (let i = 0; i < 組.length; i++) {
        const r = 55 + i;
        window.setCell(r, 13, 組[i][1]);
        結[組[i][0]] = null;
      }
      if (typeof window.recalcSheet === 'function') window.recalcSheet(window.activeSheet, sh.data);
      for (let i = 0; i < 組.length; i++) {
        const x = (sh.data || {})[(55 + i) + ',13'];
        const v = x ? (x.d !== undefined ? x.d : x.v) : '(無い)';
        結[組[i][0]] = (v === undefined || v === null) ? '' : String(v);
      }
      return 結;
    });
    console.log('  ★2つの 道具が 割れた 7個 ... ' + JSON.stringify(出));
    const 名前が通らない = Object.keys(出).filter((k) => String(出[k]).indexOf('#NAME?') >= 0);
    T('★★7個とも お客さんの 画面で `#NAME?` に ならない★★（★台は 知って いる★）',
      名前が通らない.length === 0, '名前が 通らない ... ' + 名前が通らない.join(' '));
    T('★★=LET(x,2,x*3) が 6★★', 出.LET === '6', '出た "' + 出.LET + '"');
    T('★★=LAMBDA(x,x*2)(4) が 8★★', 出.LAMBDA === '8', '出た "' + 出.LAMBDA + '"');
    T('★★=AREAS(K51:L52) が 1★★', 出.AREAS === '1', '出た "' + 出.AREAS + '"');
  }

  /* == ★★台の 溢れが マスに 並ぶか★★ ==（2026-09-18・経営者1 の 実測から）
       `=BYROW(F1:G2,LAMBDA(r,SUM(r)))` ... ★台は 3 と 30 を 出して います★
         ★前★ ... 溢れは いつも 借り物に 落として いた ⇒ 借り物は BYROW を 知らない ⇒ `#NAME?`
         ★今★ ... ★借り物が 誤りを 返した 時だけ 台の 溢れを 使います★
                  ⇒★下がる 事が ありません★（★悪い 方を 良い 方に 換えるだけ★）
       ★横にも 縦にも 並ぶ かを 両方 見ます★（BYROW は 縦／BYCOL は 横） */
  {
    const 出 = await page.evaluate(() => {
      const sh = window.sheets[window.activeSheet];
      window.setCell(70, 5, '1'); window.setCell(70, 6, '2');
      window.setCell(71, 5, '10'); window.setCell(71, 6, '20');
      window.setCell(70, 8, '=BYROW(F71:G72,LAMBDA(r,SUM(r)))');   /* I71 ... 縦に 3, 30 */
      window.setCell(73, 8, '=BYCOL(F71:G72,LAMBDA(c,SUM(c)))');   /* I74 ... 横に 11, 22 */
      window.setCell(75, 8, '=SORT(F71:G71)');                     /* 借り物が 答える 溢れ */
      if (typeof window.recalcSheet === 'function') window.recalcSheet(window.activeSheet, sh.data);
      const 読 = (r, c) => {
        const x = (sh.data || {})[r + ',' + c];
        if (!x) return '(無い)';
        const v = (x.d !== undefined ? x.d : x.v);
        return (v === undefined || v === null) ? '' : String(v);
      };
      return {
        byrow頭: 読(70, 8), byrow次: 読(71, 8),
        bycol頭: 読(73, 8), bycol次: 読(73, 9),
        sort頭: 読(75, 8), sort次: 読(75, 9),
      };
    });
    console.log('  ★BYROW ... "' + 出.byrow頭 + '" / "' + 出.byrow次 + '"（★3 と 30 が 正★）');
    console.log('  ★BYCOL ... "' + 出.bycol頭 + '" / "' + 出.bycol次 + '"（★11 と 22 が 正★）');
    console.log('  ★SORT  ... "' + 出.sort頭 + '" / "' + 出.sort次 + '"（★借り物の 溢れ・1 と 2★）');
    T('★★BYROW の 頭が 3★★（★前は #NAME?★）', 出.byrow頭 === '3', '出た "' + 出.byrow頭 + '"');
    T('★★BYROW が ★縦に★ 並ぶ（次が 30）★★', 出.byrow次 === '30', '出た "' + 出.byrow次 + '"');
    T('★★BYCOL の 頭が 11★★', 出.bycol頭 === '11', '出た "' + 出.bycol頭 + '"');
    T('★★BYCOL が ★横に★ 並ぶ（次が 22）★★', 出.bycol次 === '22', '出た "' + 出.bycol次 + '"');
    T('★★借り物が 答えられる 溢れは 今まで通り★★（=SORT ⇒ 1 と 2）',
      出.sort頭 === '1' && 出.sort次 === '2', JSON.stringify(出));
  }

  /* == ★★JS層が 引数の 数違いを 飲み込んで いた 1本★★ ==（2026-09-18）
       `=MAP(A1:A5,LAMBDA(x,y,x))`（★口が 2つ／MAP は 1つしか 渡さない★）
         ★実Excel★ ... `#VALUE!`（`golden-kansuu-9kaime-2026-09-18.tsv` 55行目）
         ★前の うち★ ... `#ERROR!`（★実Excel に 無い 字★／★JS層から★）
       ⇒★JS層が 答えない ように し、台に 回しました★
       ★`#ERROR`（`!` 無し）は 借り物から／`#ERROR!`（`!` 有り）は JS層から★ */
  {
    const 出 = await page.evaluate(() => {
      const sh = window.sheets[window.activeSheet];
      for (let i = 0; i < 5; i++) window.setCell(95 + i, 0, String(i + 1));
      window.setCell(95, 3, '=MAP(A96:A100,LAMBDA(x,y,x))');
      if (typeof window.recalcSheet === 'function') window.recalcSheet(window.activeSheet, sh.data);
      const c = (sh.data || {})['95,3'];
      const v = c ? (c.d !== undefined ? c.d : c.v) : '(無い)';
      return String(v === undefined || v === null ? '' : v);
    });
    console.log('  ★=MAP(A96:A100,LAMBDA(x,y,x)) ... "' + 出 + '"（★#VALUE! が 正★）');
    T('★★引数の 数違いの MAP が 実Excel と 同じ #VALUE!★★（★前は #ERROR!★）',
      出 === '#VALUE!', '出た "' + 出 + '"');
    T('★★`#ERROR!`（実Excel に 無い 字）を 出して いない★★',
      出.indexOf('#ERROR') < 0, '出た "' + 出 + '"');
  }

  /* == ★★台が 知らない 式は 借り物に 落ちるか★★ ==（2026-09-18・㋑⑶）
       ★訳★ … 台に 無い 関数は ★27個★ 残って います。
              落ちる 道が 無いと ★その場で `#NAME?`★ に なり ★今 出て いる 答えが 消えます★
       ★`tsunagu-mon` は 「落とす 道が 在るか」を 字で 見ます★
       ⇒★★ここでは ★効いて いるか★を 実物の ブラウザで 見ます★★
         （記憶「入れて 落ちないかで 測るな＝引いて 正しい 答えが 出るかで 測れ」）
       ★見方★ … `_台に聞く` が ★知らない 式で null／知って いる 式で 中身★を 返すか */
  {
    const 落とす = await page.evaluate(() => {
      if (typeof window._台に聞く !== 'function') return { 口が無い: true };
      const 知らない = window._台に聞く(window.activeSheet, 90, 0, '=WEBSERVICE("x")');
      const 知って = window._台に聞く(window.activeSheet, 91, 0, '=SUM(1,2,3)');
      const 溢れ = window._台に聞く(window.activeSheet, 92, 0, '=SORT({3;1;2})');
      return {
        知らない: 知らない === null,
        知って: !!(知って && 知って.字 === '6'),
        溢れ: 溢れ === null,
      };
    });
    console.log('  ★落とす 道 … 知らない式 ' + 落とす.知らない + ' ／ 知って いる式 '
      + 落とす.知って + ' ／ 溢れ ' + 落とす.溢れ + '★');
    T('★★台に 無い 関数は 借り物に 落ちる★★（=WEBSERVICE ⇒ 台は 答えない）',
      落とす.知らない === true, JSON.stringify(落とす));
    T('★★台が 知って いる 式は 台が 答える★★（=SUM(1,2,3) ⇒ 6）',
      落とす.知って === true, JSON.stringify(落とす));
    T('★★溢れも 借り物に 落ちる★★（=SORT ⇒ 台は 答えない＝並べるのは 借り物）',
      落とす.溢れ === true, JSON.stringify(落とす));
  }

  /* == ★★打った 直後の マスが 新しい 答えか★★ ==（2026-09-18・経営者1 の 注文）
       ★訳★ … `setCellFormula` に 入った 時 `data` は ★まだ 古い★
              ⇒★合わせる → その マスを 打つ★の 順で なければ ★1マス 古いまま 答えます★
              ⇒★そして 写しの 門は ★緑のまま★です★（合わせた 直後の 食い違いは 0）
       ⇒★★ここでしか 出ません★★ */
  {
    const 新しさ = await page.evaluate(() => {
      const sh = window.sheets[window.activeSheet];
      const 読む = (r, c) => {
        const x = (sh.data || {})[r + ',' + c];
        if (!x) return '(無い)';
        const v = (x.d !== undefined ? x.d : x.v);
        return (v === undefined || v === null) ? '' : String(v);
      };
      window.setCell(80, 0, '5');                 /* A81 = 5 */
      window.setCell(80, 1, '=A81*2');            /* B81 = =A81*2 */
      const 打った直後 = 読む(80, 1);
      window.setCell(80, 0, '7');                 /* A81 を 7 に */
      if (typeof window.recalcSheet === 'function') window.recalcSheet(window.activeSheet, sh.data);
      const 直した後 = 読む(80, 1);
      return { 打った直後, 直した後 };
    });
    console.log('  ★打った 直後の B81 … "' + 新しさ.打った直後 + '"（★10 が 正★）'
      + ' ／ A81 を 7 に した 後 … "' + 新しさ.直した後 + '"（★14 が 正★）');
    T('★★打った 直後の マスが 新しい 答え★★（A81=5 ⇒ B81 ==A81*2 ⇒ 10）',
      新しさ.打った直後 === '10', '出た "' + 新しさ.打った直後 + '"');
    T('★★元の マスを 直したら 追いかける★★（A81=7 ⇒ B81 ⇒ 14）',
      新しさ.直した後 === '14', '出た "' + 新しさ.直した後 + '"');
  }

  /* ══ ★★1本ずつ 打って 読む★★ ══ */
  let 合 = 0, 違 = 0, 空 = 0, 転 = 0;
  const 外れ = [];
  for (let i = 0; i < 式たち.length; i++) {
    const x = 式たち[i];
    let 出;
    try {
      出 = await page.evaluate(({ r, f }) => {
        /* ★★呼び方は 本番と 同じに します★★（2026-09-18 に 間違えました）
             `recalcSheet` は ★引数 2つ★（`recalcSheet(sheetIdx, data)`）
             ⇒引数 無しで 呼ぶと `Object.keys(data)` で 転びます
             ⇒★本番の 呼び方★ … `recalcSheet(activeSheet, sheets[activeSheet].data)`
               （book.html:2010 ／ 11967 ／ 15830 と 同じ） */
        const sh = window.sheets[window.activeSheet];
        window.setCell(r, 0, f);
        if (typeof window.recalcSheet === 'function') window.recalcSheet(window.activeSheet, sh.data);
        const d = sh.data || {};
        const c = d[r + ',0'];
        if (!c) return { 空: true };
        const v = (c.d !== undefined ? c.d : c.v);
        return { 値: (v === undefined || v === null) ? '' : String(v) };
      }, { r: i, f: x.式 });
    } catch (e) {
      /* ★③1本ずつ 受け止める★ */
      転++; 外れ.push(x.式 + ' ⇒ ★転んだ★ ' + String(e.message).slice(0, 60));
      continue;
    }
    if (!出 || 出.空 || 出.値 === '') { 空++; 外れ.push(x.式 + ' ⇒ ★空っぽ★'); continue; }
    if (出.値 === x.答) 合++;
    else { 違++; 外れ.push(x.式 + ' ⇒ 出た「' + 出.値 + '」／はず「' + x.答 + '」'); }
  }

  console.log('  ★★合った ' + 合 + ' / ' + 式たち.length + '★★'
    + ' ／ 違った ' + 違 + ' ／ ★空っぽ ' + 空 + '★ ／ 転んだ ' + 転);
  for (const s of 外れ.slice(0, 8)) console.log('       ・' + s);
  T('★★空っぽが 0件★★（★0件を 緑に しない★）', 空 === 0, '空っぽ ' + 空 + '件');
  T('★★対照（=SUM(1,2,3)）が 6★★', 合 > 0 && !外れ.some((s) => s.startsWith('=SUM(')));

  if (終わりの線か) {
    /* ★★終わりの 線★★（★今は 赤＝当たり前★） */
    T('★★' + 式たち.length + '本 とも 合う★★', 合 === 式たち.length,
      '合 ' + 合 + ' ／ 違 ' + 違 + ' ／ 空 ' + 空 + ' ／ 転 ' + 転);
    T('★★hyperformula が 0本★★', 読み.hf === 0, 読み.hf + '本');
    T('★★shiki- が 1本 以上★★', 読み.shiki >= 1, 読み.shiki + '本');
  } else {
    /* ★★繋ぐ 前の 姿と 同じか★★（★姿が 変わったら 赤＝それが 狙い★） */
    console.log('  ★姿の 移り変わり★ 繋ぐ前 合' + 繋ぐ前.合 + '/違' + 繋ぐ前.違
      + ' → ㋐ 合' + 読み込んだ後.合 + '/違' + 読み込んだ後.違
      + ' → ★㋑ 合' + 合 + '/違' + 違 + '★');
    T('★★下がって いない（★合った 本数が ㋐の ' + 読み込んだ後.合 + ' 以上★）★★',
      合 >= 読み込んだ後.合, '合 ' + 合 + ' ／ ㋐ ' + 読み込んだ後.合);
    T('★★台に 回した 後の 姿と 同じ（合 ' + 台に回した後.合 + ' / ' + 式たち.length + '）★★',
      合 === 台に回した後.合 && 違 === 台に回した後.違,
      '★姿が 変わりました★ 合 ' + 合 + '（決め打ち ' + 台に回した後.合 + '）／違 ' + 違 + '（決め打ち ' + 台に回した後.違 + '）'
      + '\n       ⇒★繋いだ 後なら これが 狙いです★／★繋いで いない なら 止まって 訳を 探す★');
    const 違名 = 外れ.map((s) => s.split(' ⇒ ')[0]);
    T('★★借り物が 間違える 2本が ★直った★★★（★台が 答えて います★）',
      台に回した後.違う式.every((f) => 違名.includes(f)) && 違名.length === 台に回した後.違,
      '出た ' + JSON.stringify(違名) + ' ／ 繋ぐ 前 ' + JSON.stringify(繋ぐ前.違う式));
    /* ★★㋐が 済んだ 姿★★（2026-09-18）＝★増えても 減っても 赤★ */
    T('★★`shiki-` が ' + 台に回した後.shiki + '本★★（★減っても 増えても 赤★）',
      読み.shiki === 台に回した後.shiki,
      読み.shiki + '本（決め打ち ' + 台に回した後.shiki + '本 ／ ㋐ ' + 読み込んだ後.shiki
      + '本 ／ 繋ぐ 前 ' + 繋ぐ前.shiki + '本）');
    T('★★`hyperformula` は まだ ' + 読み込んだ後.hyperformula + '本★★（★外すのは また 別の 1押し★）',
      読み.hf === 読み込んだ後.hyperformula,
      読み.hf + '本');
    /* ★★窓の 誤り★★（★許した 物を 引いてから 数えます★） */
    const 新しい誤り = 窓の誤り.filter((s) => !許す誤り.some((k) => s.indexOf(k) >= 0));
    console.log('  ★窓の 誤り ... 全部 ' + 窓の誤り.length + '件'
      + ' ／ ★許した 物を 引いて ' + 新しい誤り.length + '件★');
    for (const s of 窓の誤り) console.log('       ・' + s);
    T('★★窓の 誤り（新しい 物）が 0件★★（★読み込んだ だけで 投げる 物は ここにしか 出ません★）',
      新しい誤り.length === 0, 新しい誤り.join('\n       '));
  }
} catch (e) {
  if (開けた) { fail++; console.log('  NG   ★途中で 止まりました★ ' + String(e.message).slice(0, 120)); }
} finally {
  if (配信) 配信.閉じる();
  await browser.close();
}

const 秒 = ((Date.now() - 時計) / 1000).toFixed(1);
console.log('');
console.log('  ★掛かった 秒 … ' + 秒 + '秒★');
console.log('  ★秒が 毎回 同じなら ★上限に 当てて いる★ か ★空振り★ を 疑う★');
console.log('shiki-wo-osu-webkit: ' + pass + ' 緑 / ' + fail + ' 赤');
process.exit(fail ? 1 : 0);
