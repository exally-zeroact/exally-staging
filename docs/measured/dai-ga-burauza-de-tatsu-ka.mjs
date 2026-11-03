/* dai-ga-burauza-de-tatsu-ka.mjs — ★台が ★本物の ブラウザ★で 立つか★（2026-09-18）
 *
 *  ★★なぜ 繋ぐ 前に これを 走らせるか★★
 *    経営者1「★私の `osu-dai-webkit` は この 順で 読んで 板が 立ちませんでした★
 *              ＝足りない 本が 在るかも しれません／先に 手元で 1回 開いて ください」
 *    ⇒★★`book.html` を 触る 前に ★同じ 読み方★で 1回 立ててみます★★
 *    ⇒★記憶「★読み込む ≠ 登録される★」★
 *
 *  ★★読む 順（★依存を 数えて 決めました★・当て推量では ありません）★★
 *    各 `lib/shiki-*.js` の `root.XXX` を 全部 出して 依存を 引きました
 *      shiki-keisan  ... 無し
 *      shiki-kiru    ... 無し
 *      shiki-katachi ... 無し
 *      shiki-sansho  ... 無し
 *      shiki-afure   ... ShikiKeisan
 *      shiki-basho   ... ShikiKeisan
 *      shiki-kansuu  ... ShikiKeisan / ShikiBasho / Shoshiki
 *      shiki-tsunagi ... ShikiKeisan / ShikiAfure / Formula* / Bahttext / Bessel
 *      shiki-hyou    ... 上の 全部 ＋ ShikiKiru / ShikiSansho / ShikiKatachi / ShikiKansuu / ShikiTsunagi
 *    ⇒★経営者1 の 順には ★`shiki-kiru` と `shiki-sansho` が 入って いませんでした★★
 *      ＝★だから 板が 立たなかった 見込み★（★この 道具で 確かめます★）
 *
 *  ★★先に 見つけた 2つ（★ブラウザでしか 出ません★）★★
 *    ⑴`lib/shiki-tsunagi.js` が `root.FormulaFilterxml`（小文字 x）を 読んで いる
 *       ⇒本物は `root.FormulaFilterXml`（★大文字 X★）＝★ブラウザでは `undefined`★
 *       ⇒node の `require` では 通るので ★今まで 出ませんでした★
 *    ⑵`lib/bessel.js` が `book.html` に ★1本も 在りません★
 *       ⇒`shiki-tsunagi` は `root.Bessel` を 読みます
 *    ⇒★★どちらも「読み込む ≠ 登録される」です★★
 *
 *  ★★この 道具が 見て いない 事★★
 *    ・★`book.html` では ありません★（★同じ 本を 同じ 順で 読むだけの 白紙★）
 *    ・★お客さんの 道（`convertFormula` / JS層）は 通って いません★
 *
 *  使い方: node docs/measured/dai-ga-burauza-de-tatsu-ka.mjs
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { borrow, launch } from 'file:///C:/Users/zeroa/exally-prod/scripts/_borrow-playwright.mjs';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '../..');

/* ★読む 順★（★依存から 決めた★／★`book.html` に 足す 順も これ★） */
const 順 = [
  /* ★先に 要る 物（`book.html` に 既に 在る 物も 並べます）★ */
  'lib/shoshiki.js',
  'lib/bahttext.js',
  'lib/bessel.js',                 /* ★★book.html に 在りません＝足す 要る★★ */
  'lib/formula-extra.js',
  'lib/formula-kane.js',
  'lib/formula-filterxml.js',
  'lib/formula-cell.js',
  'lib/formula-yosoku.js',
  'lib/formula-nokori.js',
  /* ★台★ */
  'lib/shiki-keisan.js',
  'lib/shiki-kiru.js',
  'lib/shiki-katachi.js',
  'lib/shiki-sansho.js',
  'lib/shiki-afure.js',
  'lib/shiki-basho.js',
  'lib/shiki-kansuu.js',
  'lib/shiki-tsunagi.js',
  'lib/shiki-hyou.js',
];

let 緑 = 0, 赤 = 0;
const T = (n, よい, 添) => {
  if (よい) { 緑++; console.log('  ok   ' + n); }
  else { 赤++; console.log('  NG   ' + n + (添 ? '\n       ' + 添 : '')); }
};

console.log('');
console.log('[dai-ga-burauza-de-tatsu-ka] ★台が 本物の ブラウザで 立つか★');
console.log('  ★読む 本 ... ' + 順.length + '本★');

for (const p of 順) {
  if (!fs.existsSync(path.join(ROOT, p))) { T('★' + p + ' が 在る★', false); }
}
if (赤) { console.log('\n★本が 足りないので 止めます★'); process.exit(1); }

const 中身 = 順.map((p) => ({ 道: p, 字: fs.readFileSync(path.join(ROOT, p), 'utf-8') }));

const wk = await borrow('dai-tatsu', 'webkit');
const browser = await launch('dai-tatsu', wk, {}, 'webkit');
const page = await browser.newPage();
try {
  await page.setContent('<!doctype html><meta charset="utf-8"><title>dai</title>');
  const 転んだ = [];
  for (const x of 中身) {
    try { await page.addScriptTag({ content: x.字 }); }
    catch (e) { 転んだ.push(x.道 + ' ... ' + String(e.message).slice(0, 80)); }
  }
  T('★★' + 順.length + '本 とも 読み込めた★★', 転んだ.length === 0, 転んだ.join('\n       '));

  /* ★★「読み込む」と「登録される」は 別★★＝★名前が 窓に 在るか★ */
  const 名 = await page.evaluate(() => {
    const 要る = ['Shoshiki', 'BahtText', 'Bessel', 'FormulaExtra', 'FormulaKane',
      'FormulaFilterXml', 'FormulaCell', 'FormulaYosoku', 'FormulaNokori',
      'ShikiKeisan', 'ShikiKiru', 'ShikiKatachi', 'ShikiSansho', 'ShikiAfure',
      'ShikiBasho', 'ShikiKansuu', 'ShikiTsunagi', 'ShikiHyou'];
    const 出 = {};
    for (const n of 要る) 出[n] = (typeof window[n] !== 'undefined');
    return 出;
  });
  const 無い = Object.keys(名).filter((k) => !名[k]);
  T('★★名前が 全部 窓に 在る★★（★読み込む ≠ 登録される★）', 無い.length === 0, '無い ... ' + 無い.join(' '));

  /* ★★皮が 中身を 掴めたか★★（★小文字 x の 件★） */
  const 皮 = await page.evaluate(() => {
    if (typeof window.ShikiTsunagi !== 'object') return { 口: false };
    const n = (typeof window.ShikiTsunagi.名前たち === 'function') ? window.ShikiTsunagi.名前たち() : [];
    return { 口: true, 数: n.length, FILTERXML: n.indexOf('FILTERXML') >= 0, BESSELJ: n.indexOf('BESSELJ') >= 0 };
  });
  T('★皮の 口が 在る★', 皮.口);
  T('★★皮が 65個 出せる★★', 皮.数 === 65, '出た ' + 皮.数 + '個');
  T('★★皮が FILTERXML を 持って いる★★（★小文字 x の 件★）', !!皮.FILTERXML);
  T('★皮が BESSELJ を 持って いる★（★bessel.js の 件★）', !!皮.BESSELJ);

  /* ★★板が 立って 式が 通るか★★ */
  const 押 = await page.evaluate(() => {
    try {
      const h = window.ShikiHyou.表();
      for (let i = 1; i <= 5; i++) h.打つ('A' + i, String(i));
      const 出 = {};
      const 試 = (f) => { h.打つ('J1', f); return String(h.字('J1')); };
      出.SUM = 試('=SUM(A1:A5)');
      出.MAP = 試('=SUM(MAP(A1:A5,LAMBDA(x,x*2)))');
      出.LAMBDA = 試('=LAMBDA(x,x*2)(3)');
      出.ODDF = 試('=ODDFPRICE(DATE(2009,3,1),DATE(2013,1,1),DATE(2009,1,1),DATE(2010,1,1),0.06,0.05,100,2,1)');
      出.FILTERXML = 試('=FILTERXML("<r><b>1</b><b>2</b></r>","//b")');
      出.台が知る = Object.keys(window.ShikiKansuu.表 || {}).length;
      return 出;
    } catch (e) { return { 誤: String(e.message).slice(0, 140) }; }
  });
  if (押.誤) {
    T('★★板が 立つ★★', false, '★投げました★ ' + 押.誤);
  } else {
    T('★★板が 立つ★★', true, '台が 知る ' + 押.台が知る + '個');
    T('★=SUM(A1:A5) が 15★', 押.SUM === '15', '出た「' + 押.SUM + '」');
    T('★=SUM(MAP(...)) が 30★', 押.MAP === '30', '出た「' + 押.MAP + '」');
    T('★=LAMBDA(x,x*2)(3) が 6★', 押.LAMBDA === '6', '出た「' + 押.LAMBDA + '」');
    T('★★=ODDFPRICE が #NAME? で ない★★（★皮に 載って いません＝#NAME? が 正★）',
      押.ODDF === '#NAME?', '出た「' + 押.ODDF + '」');
    T('★★=FILTERXML が 通る★★（★ブラウザの XPath★）', 押.FILTERXML === '1',
      '出た「' + 押.FILTERXML + '」');
  }
} catch (e) {
  console.log('  NG   ★途中で 止まりました★ ' + String(e.message).slice(0, 140));
  赤++;
} finally {
  await browser.close();
}

console.log('');
console.log('★言えない 事★');
console.log('  ・★`book.html` では ありません★（★同じ 本を 同じ 順で 読むだけの 白紙★）');
console.log('  ・★お客さんの 道（convertFormula / JS層）は 通って いません★');
console.log('');
console.log('dai-ga-burauza-de-tatsu-ka: ' + 緑 + ' 緑 / ' + 赤 + ' 赤');
process.exit(赤 ? 1 : 0);
