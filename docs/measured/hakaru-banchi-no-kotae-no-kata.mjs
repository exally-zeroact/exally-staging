/* hakaru-banchi-no-kotae-no-kata.mjs ･･･ ★この 番地の 答えは 何型か★ 2026-09-25
 *
 *  ★★なぜ 要るか★★
 *    経営者1 と 私で ★同じ マスが 違って 見えて います★（2026-09-25）
 *      経営者1 の 物差し（実Excel の `.Text`）... ★数が 出る★
 *    ★★測った 答え（2026-09-25・15枚 まるごと）★★
 *      ★15,798個 / 15,799個 が 同じ★＝★描く 側は 1つも 化けさせて いません★
 *      違ったのは ★給料表|272,7 の 1個だけ★（11字 → 9字・`General` の 桁の 話）
 *    ⇒★`#REF!` は ★私の 側にも 1個も 出ません★（654個・15,799個 とも 誤り 0個）★
 *      私の 画面 ........................... ★`#REF!` が 出る★
 *    ⇒経営者1「★どちらかの 取り方が 違います★」
 *    ⇒経営者1 は ★実Excel から 取り直します★
 *    ⇒★私は 本の 中の バイトと 画面を 数えます★（★重ならない 半分★）
 *
 *  ★★出すのは 型と 数だけです★★（★司さんの 実物の 中身は 1字も 出しません★）
 *    ⑴本の 中に 保存されて いる 答えが ★誤り／数／字／無し★ の どれか
 *    ⑵誤りなら ★どの 誤りか★（`#REF!` `#VALUE!` ･･･ ＝★これは 中身では なく 印です★）
 *    ⑶式に 出て くる ★関数の 名前★（★番地や 数は 出しません★）
 *    ⑷画面が 出す 字が ★誤りか／誤りでないか★（★字そのものは 出しません★）
 *    ⇒★これで 「どちらの 取り方が 違うか」が 決まります★
 *
 *  ★★この 道具で 分かる 事と 分からない 事★★（★2026-09-25 言葉を 1回 間違えた★）
 *    `cell.d` を 「本に 保存された 答え」と 書いて いました。★違います★。
 *    `js/book-open.js` が ★書式を 掛けて 作った 字★ です（`XLSX.SSF.format`）。
 *    ⇒★★分かる 事★★ ･･･ ★描く 側が それを 変えて いないか★
 *    ⇒★★分からない 事★★ ･･･ ★book-open の 書式の 掛け方が 実Excel と 同じか★
 *    ⇒★そして 実Excel の 「保存された 字」は 本の 中に 在りません★
 *       ＝本に 在るのは ★答えの 値★ だけ。字は ★Excel も うちも 後で 作ります★
 *       ⇒★★だから 表示の 違いは 「計算の 違い」では なく 「書式の 違い」です★★
 *
 *  ★走らせ方★
 *    node docs/measured/hakaru-banchi-no-kotae-no-kata.mjs "<本の 道>" "板|行,列" ["板|行,列" ...]
 */
import path from 'node:path'; import http from 'node:http'; import fs from 'node:fs';
import { fileURLToPath } from 'node:url';
import { borrow, launch } from '../../scripts/_borrow-playwright.mjs';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..', '..');
const 引数 = process.argv.slice(2);
/* ★★`--台 webkit` の 「webkit」を 番地と 取り違えて いました★★（2026-09-25 1回 踏んだ）
     ⇒★札の 次の 1個は 札の 値＝素から 外します★ */
const 札の値 = ['台', '板'];
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
const 番地たち = 素.slice(1);
const 取る = (名, 既定) => { const i = 引数.indexOf('--' + 名); return i >= 0 ? 引数[i + 1] : 既定; };
const 台 = 取る('台', 'webkit');
/* ★`--まとめだけ` ･･･ 1つずつ 出さない（★自分で 数える のを やめる＝道具に 出させる★）★ */
const まとめだけ = 引数.indexOf('--まとめだけ') >= 0;
/* ══ ★★`--打ってから` ･･･ 1マス 打って 計算させた 後で 数えます★★ ══（2026-09-25）
     ★★なぜ★★
       `lib/table-refs.js` は 消された 表を `#REF!` に 置き換えて います。
       ⇒★台に 渡る 式には `#REF!` が 在る（69個の うち 67個）★
       ⇒★なのに 画面は 数を 出す＝計算して いない から 保存値を 出して いる★
       ⇒★★では 計算させたら 実Excel と 同じ `#REF!` に なるのか★★
     ＝★打つのは 画面の `setCell`＝お客さんの 1打ち目と 同じ 道★
     ＝★遠い 空きマスに 打ちます（元の 数を 触りません・保存しません）★ */
const 打ってから = 引数.indexOf('--打ってから') >= 0;
if (!本 || !fs.existsSync(本)) { console.log('★本が 在りません★ ' + 本); process.exit(2); }
/* ★★`--板 <名>` ･･･ その 板の 式の マスを 全部 見ます★★（2026-09-25）
     ＝★4個だけでは 分母が 小さい★＝★板 まるごとで 数えられる ように します★ */
const 板まるごと = 取る('板', null);
/* ★`--全部` ･･･ 15枚 まるごと（★分母を 小さくしない★）★ */
const 全部見る = 引数.indexOf('--全部') >= 0;
if (!番地たち.length && !板まるごと && !全部見る) { console.log('★番地か --板 を 下さい★（例 "給料3|71,4" ／ --板 給料3）'); process.exit(2); }

const 誤りの印 = ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA', '#SPILL!', '#CALC!'];

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

console.log('[番地の 答えの 型] ★出すのは 型と 数だけ＝中身は 1字も 出しません★');
console.log('  ★本★ ' + path.basename(本) + '（★読むだけ／1バイトも 書きません★）／★台★ ' + 台);
console.log('  ★番地★ ' + 番地たち.length + '個');

const wk = await borrow('banchi-no-kata', 台);
const browser = await launch('banchi-no-kata', wk, {}, 台);
const page = await browser.newPage({ viewport: { width: 1280, height: 900 } });
const 配信 = await 立てる(ROOT);

try {
  let 声 = '';
  page.on('console', (m) => { const t = String(m.text()); if (t.indexOf('開いた 直後の 計算') >= 0) 声 = t; });
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
  if (声) console.log('  [画面の 声] ' + 声);

  /* ★★要る 関数が 無ければ 数を 出さずに 止まります★★（2026-09-25 の 決め） */
  const 無い = await page.evaluate(() => ['switchSheet', '_字の元', '_答えは字か', 'fmtForDisplay', '_入る字数', 'cW']
    .filter((n) => typeof window[n] !== 'function'));
  if (無い.length) { console.log('  ★★測れません★★ 在りません ... ' + 無い.join(' / ')); process.exit(8); }

  /* ★`--板` なら その 板の 式の マスを 全部 集めます★ */
  let 見る番地 = 番地たち;
  if (全部見る) {
    見る番地 = await page.evaluate(() => {
      const 出 = [];
      (window.sheets || []).forEach((s) => {
        const d = (s && s.data) || {};
        for (const k in d) { const c = d[k]; if (c && typeof c.f === 'string' && c.f.charAt(0) === '=') 出.push(s.name + '|' + k); }
      });
      return 出;
    });
    console.log('  ★--全部★ ... 板 ' + (await page.evaluate(() => (window.sheets || []).length))
      + '枚 ／ 式の マス ★' + 見る番地.length.toLocaleString() + '個★');
  }
  if (板まるごと) {
    見る番地 = await page.evaluate((な) => {
      const i = (window.sheets || []).findIndex((s) => s.name === な);
      if (i < 0) return null;
      const d = (window.sheets[i].data) || {};
      const 出 = [];
      for (const k in d) { const c = d[k]; if (c && typeof c.f === 'string' && c.f.charAt(0) === '=') 出.push(な + '|' + k); }
      return 出;
    }, 板まるごと);
    if (見る番地 === null) { console.log('  ★★板が 在りません★★ ' + 板まるごと); process.exit(8); }
    console.log('  ★--板 ' + 板まるごと + '★ ... 式の マス ★' + 見る番地.length.toLocaleString() + '個★');
  }
  if (打ってから) {
    const 無2 = await page.evaluate(() => (typeof window.setCell === 'function' ? [] : ['setCell']));
    if (無2.length) { console.log('  ★★測れません★★ 在りません ... setCell'); process.exit(8); }
    const t0 = Date.now();
    await page.evaluate(() => { window.setCell(900, 30, '1'); });
    await page.waitForTimeout(1500);
    console.log('  ★1マス 打って 計算させた★ ... ' + (Date.now() - t0).toLocaleString() + ' ms');
  }

  const 出 = await page.evaluate(([番地たち, 誤りの印]) => {
    const 型を見る = (x) => {
      if (x === undefined) return '無し';
      if (x === null) return '無し';
      if (typeof x === 'number') return '数';
      if (typeof x === 'boolean') return '真偽';
      const s = String(x);
      if (s === '') return '空の字';
      for (const e of 誤りの印) if (s === e) return '誤り ' + e;
      if (s.charAt(0) === '#') return '誤りっぽい 字';
      /* ★★式の マスの `d` は ★書式を 掛けた 字★ です★★（経営者1 の 指摘・2026-09-25）
           ＝だから 「字」だけでは ★数なのか 文字なのか 分かりません★
           ⇒★数に 見えるか どうかを 分けます★（★値そのものは 出しません★） */
      const 裸 = s.replace(/[,\s¥￥円%]/g, '');
      if (裸 !== '' && isFinite(Number(裸))) return '数に 見える 字（' + s.length + '字）';
      return '数に 見えない 字（' + s.length + '字）';
    };
    return 番地たち.map((まる) => {
      const [な, k] = String(まる).split('|');
      const i = (window.sheets || []).findIndex((s) => s.name === な);
      if (i < 0) return { 番地: まる, 出来ず: '★板が 在りません★' };
      window.switchSheet(i);
      const cell = ((window.sheets[i] || {}).data || {})[k];
      if (!cell) return { 番地: まる, 出来ず: '★マスが 在りません★' };
      const raw = window._字の元(cell);
      let 画面;
      if (window._答えは字か(cell)) 画面 = String(raw === undefined || raw === null ? '' : raw);
      else {
        const w = window.cW(+String(k).split(',')[1]);
        画面 = String(window.fmtForDisplay(raw, cell.numFmt, window._入る字数(w, raw, cell.numFmt)));
      }
      /* ══ ★★式が 指して いる 行の 幅（数だけ）★★ ══（2026-09-25）
           経営者1 の 69個の うち ★66個は 74〜79行（0から）の べた塗り★でした。
           外の 3個は ★SUM / SUM / SUBTOTAL★＝★範囲を まとめる 関数★。
           ⇒★その 範囲が 壊れた 66個を 含むなら 訳は 1つで 足ります★
           ⇒★出すのは 行の 番号だけ＝式の 字も 値も 出しません★ */
      /* ★関数の 名前だけ 拾います★＝番地・数・字は 1つも 出しません */
      const 名 = [];
      const f = typeof cell.f === 'string' ? cell.f : '';
      const 行たち = [];
      {
        const re2 = new RegExp('[A-Z]+\\$?([0-9]+)', 'g');
        let m2;
        while ((m2 = re2.exec(f))) 行たち.push(parseInt(m2[1], 10));
      }

      const re = /([A-Z][A-Z0-9._]*)\s*\(/g;
      let m;
      while ((m = re.exec(f))) if (名.indexOf(m[1]) < 0) 名.push(m[1]);
      return {
        番地: まる,
        式か: f.charAt(0) === '=',
        式の長さ: f.length,
        関数の名: 名,
        保存の答えの型: 型を見る(cell.d),
        vの型: 型を見る(cell.v),
        画面の型: 型を見る(画面),
        書式が在るか: !!cell.numFmt,
        /* ★★直した 後の 式に `#REF!` が 在るか★★（2026-09-25）
             `lib/table-refs.js` は ★消された 表（id = 0xFFFFFFFF）を `#REF!` に 置き換えます★
             （断りに `給料3!B75 =INDEX(#REF!, MATCH(...))` と 名指しで 書いて あります）
             ⇒★台に 渡る 式には `#REF!` が 在る★
             ⇒★それでも 画面が 数を 出すのは 「計算しない」から★＝★保存値を 出して いる★
             ⇒★だから ここを 数えれば 「計算しない」との 関わりが 決まります★ */
        式にREF: f.indexOf('#REF!') >= 0,
        指す行の一番小さい: 行たち.length ? Math.min.apply(null, 行たち) : null,
        指す行の一番大きい: 行たち.length ? Math.max.apply(null, 行たち) : null,
      };
    });
  }, [見る番地, 誤りの印]);

  console.log('');
  /* ★板 まるごとの 時は 1つずつ 出しません（★出しを 溢れさせない★）★ */
  ((板まるごと || 全部見る || まとめだけ) ? [] : 出).forEach((x) => {
    console.log('  ══ ' + x.番地 + ' ══');
    if (x.出来ず) { console.log('    ' + x.出来ず); return; }
    console.log('    式か ................... ' + (x.式か ? '★式★（' + x.式の長さ + '字）' : '定数'));
    console.log('    関数の 名 .............. ' + (x.関数の名.length ? x.関数の名.join(' / ') : '（無し）'));
    console.log('    ★book-open が 作った 答え★ ... ' + x.保存の答えの型);
    console.log('    `v` の 型 .............. ' + x.vの型);
    console.log('    ★画面が 出す 字★ ......... ' + x.画面の型);
    console.log('    書式が 在るか .......... ' + (x.書式が在るか ? '在る' : '無い'));
    console.log('    ★直した 後の 式に `#REF!` が 在るか★ ... ' + (x.式にREF ? '★在る★' : '無い'));
    console.log('    ★式が 指す 行（Excel の 数え方）★ ... '
      + (x.指す行の一番小さい === null ? '（無し）'
        : x.指す行の一番小さい + '行 〜 ' + x.指す行の一番大きい + '行'));
  });

  const 誤り = 出.filter((x) => !x.出来ず && String(x.画面の型).slice(0, 2) === '誤り').length;
  const 保存も誤り = 出.filter((x) => !x.出来ず && String(x.保存の答えの型).slice(0, 2) === '誤り').length;
  console.log('');
  console.log('  ★まとめ★ 見た ' + 出.length + '個 ／ ★画面が 誤り ' + 誤り + '個★ ／ ★book-open の 答えも 誤り ' + 保存も誤り + '個★');
  /* ══ ★★誤りが 0個の 時に 何も 言わない のは 黙って いるのと 同じ★★ ══（2026-09-25 1回 踏んだ）
       ⇒★「本の 答え」と 「画面の 字」が 同じ 型か を いつも 出します★ */
  const 見た = 出.filter((x) => !x.出来ず);
  const 同じ = 見た.filter((x) => x.保存の答えの型 === x.画面の型).length;
  /* ★板 まるごと・全部の 時も この 数は 出します★ */
  {
    const 見た2 = 出.filter((x) => !x.出来ず);
    const REF = 見た2.filter((x) => x.式にREF).length;
    console.log('  ★直した 後の 式に `#REF!` が 在る★ ... ★' + REF + '個 / ' + 見た2.length + '個★');
  }
  console.log('  ★book-open の 答えと 画面の 字が 同じ 型★ ... ★' + 同じ + '個 / ' + 見た.length + '個★'
    + (見た.length && 同じ === 見た.length ? '  ★★＝画面は book-open の 答えを そのまま 出して います★★' : ''));
  if (見た.length && 同じ === 見た.length) {
    console.log('  ⇒★描く 側では 化けて いません★＝★違いが 在るなら 書式を 掛ける 所（book-open）か 物差しの 取り方です★');
  } else if (見た.length) {
    console.log('  ⇒★★book-open と 画面が 違います＝描く 側で 変わって います★★');
    見た.filter((x) => x.保存の答えの型 !== x.画面の型).slice(0, 8).forEach((x) => {
      console.log('      ' + x.番地 + ' ... book-open ' + x.保存の答えの型 + ' → 画面 ' + x.画面の型);
    });
  }
  console.log('  ★中身は 1字も 出して いません★（型と 数と 関数の 名だけ）');
} finally {
  await page.close().catch(() => {});
  await browser.close().catch(() => {});
  配信.閉じる();
}
