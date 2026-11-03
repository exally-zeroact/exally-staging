/* xpath-name-honmono-de-osu.mjs — ★本物の ブラウザで `name()` が 読めるか★（2026-09-18）
 *
 *  ★★なぜ★★
 *    経営者1 が「★`//*[name()='b']` が 台で #VALUE!★＝繋ぐ 前の 関門」と 出しました。
 *    ★でも `lib/formula-filterxml.js` は ★自分で XPath を 作って いません★★
 *      ＝`document.evaluate`（★ブラウザの 本物★）に そのまま 渡して います。
 *    ⇒★★だとすると 読めない のは ★渡した 相手★の 話です★★
 *
 *  ★★測って 分かれました★★
 *    jsdom ........ `//*[name()='b']` -> ★#VALUE!★（★jsdom が `name()` を 持って いない★）
 *    ★本物（WebKit）★ -> ★["1","2","3"]★（★読めます★）
 *    ⇒★★台の 欠陥では ありません★★＝★測り道具（jsdom）の 限りでした★
 *    ⇒★実Excel も `FILTERXML(...)` -> 1 ／ `ROWS(...)` -> 3★＝★合います★
 *
 *  ★★今日 これで 2回目です★★
 *    1回目 ... 経営者1 の `osu-dai-dake.mjs` が node で 押して
 *            `FILTERXML` 37本 全部 #VALUE!（★道具に DOMParser が 無い★）
 *    2回目 ... ★今★（jsdom に `name()` が 無い）
 *    ⇒★★「台が 違う」と 言う 前に ★測り道具が その 式を 押せるか★を 見る★★
 *
 *  使い方: node docs/measured/xpath-name-honmono-de-osu.mjs
 */
import { borrow, launch } from 'file:///C:/Users/zeroa/exally-prod/scripts/_borrow-playwright.mjs';
const wk = await borrow('xpath-name', 'webkit');
const browser = await launch('xpath-name', wk, {}, 'webkit');
const page = await browser.newPage();
const 出 = await page.evaluate(() => {
  const xml = "<r><b>1</b><b>2</b><b>3</b><c id='x'>A</c><c id='y'>B</c><d><e>ne</e></d></r>";
  const doc = new DOMParser().parseFromString(xml, 'text/xml');
  const 試 = (xp) => {
    try {
      const r = doc.evaluate(xp, doc, null, XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null);
      const o = [];
      for (let i = 0; i < r.snapshotLength; i++) o.push(r.snapshotItem(i).textContent);
      return o;
    } catch (e) { return '★投げた★ ' + String(e.message).slice(0, 60); }
  };
  return {
    b: 試('//b'),
    name: 試("//*[name()='b']"),
    pos: 試('//b[position()>1]'),
  };
});
console.log(JSON.stringify(出, null, 1));
await browser.close();
