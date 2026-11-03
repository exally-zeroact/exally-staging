/* hoshi-kazoeru.mjs — ★お客さんの 画面に 出る 字に ★ が 混じっていないか 数える★（2026-09-06・2026-10-08 に 広げた）
 *
 *  ★★なぜ★★
 *    ★ は ★私たちの 便りの 印★（司さん・指示役・私の やりとり）。
 *    ★お客さんには 壊れた 字に 見えます★。
 *    2026-09-06 に 指示役から 指摘され、書き出しの 窓の 字から 12個以上 抜いた。
 *    同じ日に Rakunally も ★23件★ 直している（会社ぜんぶの 決まり）。
 *
 *  ★★2026-10-08 に 広げた 訳（Rakunally の 引き継ぎ＝司さん 指示・対立役が 叩いた）★★
 *    前の 門は「口の 第1引数の ★同じ行★の 文字列」だけを 見て、★HTML の 本文は 見ていなかった★。
 *    ⇒ 門は 0本で 緑なのに、お客さんの 画面に ★ が 出て いた（数えた）：
 *       ・book.html の 窓の 説明の 本文 ･･･ 41節（例「★読むだけです。ここでは マクロを 動かしません★」）
 *       ・口の 引数が 行を またぐ／三項演算子の 形 ･･･ 8本（口の 呼び出し 380か所の うち 門が 拾えて いたのは 370）
 *    また ★逃がし（逆斜線 u 2605）と 実体参照（&#9733; など）★を 読み解かず、書かれたら 素通りする 形だった。
 *
 *  ★★見る 物（2026-10-08 から）★★
 *    ①口（showToast ／ notify ／ 窓の副題）の ★第1引数 まるごと★（括弧の 対を 数えて 取る＝行を またいでも 三項でも）
 *      ・中の 文字列（'...' "..." `...` の 地の 字）を ★JS の 逃がしを 解いて★ から 見る（逆斜線が 奇数個の 時だけ 逃がし）
 *      ・notify ／ showToast は ★innerHTML に 入れる★（book.html の notify）＝★実体参照も 解く★
 *        窓の副題 は ★textContent★＝実体参照は ★解かない★（そのまま 字で 出る）
 *    ②HTML の 見える 字（repo 直下の *.html ★全部★＝名前で 選ばない）
 *      ・本文（タグの 外の 字）と 属性の 字（style／class／id／href／src／on... は 除く）
 *      ・★隠れて いる 所（hidden・display:none）も 数える★＝窓は JS が 開くので お客さんに 見える
 *        （41節は ★全部 最初は 隠れた 窓の 中★だった＝除くと 0本に なる）
 *      ・除くのは script ／ style ／ HTML の 注記 だけ
 *      ・HTML の 読み方は jsdom（実体参照の 読み解きを 正しく 1回だけ＝&amp;#9733; は ★ に しない）
 *
 *    ③★JS の 字の塊 まるごと★（2026-10-09 に 広げた＝JSの字の星）
 *      ・注記の 外の '...' "..." `...` を ★全部★（JS が 画面へ 直に 書く 字・lib が 返して 画面が 出す 字・throw の 字 も 入る）
 *        前は ①口だけ＝textContent の 右側 42・innerHTML の 右側 31・lib の 訳 など ★214個★ が 門の 外で 客に 出て いた
 *      ・切り方は scripts/lib/chuki.mjs の 字の塊を拾う（注記を 外すのと 同じ 読み方＝正規表現を 知っている）
 *        ★2026-10-09 に 構文解析器（acorn 8.16.0・repo には 入れて いない）と 突き合わせて ★入り 269／269 一致★
 *        自前で 切ると 正規表現の 中の ` で 3,600行を 1つの 字と 飲み込み 88個が 1個に 化けた（対立役が 数えた）
 *      ・★読めない 字（引用符が 閉じない）／行を またぐ テンプレート（飲み込みの 疑い・今 4本 全部 1行）は 赤★
 *      ・除く 物＝console の 第1引数（機械が 毎回 確かめる）と 白名簿（下・数か 指紋が 違えば 赤）だけ
 *        「どの html も js も 読み込まない js」も ★白名簿に 名指しした 物だけ★ 除く（読み込まれたら 赤）
 *      ・逃がし（★）も 実体参照（&#9733;）も 解いて 見る
 *      ・読み込まれるのに 見て いない js／字で 無い 読み込み（_loadScript(名) の 変数）が 在れば 赤
 *    ④on 属性の 中の JS・<template> の 中身・CSS の content・manifest.json・fromCharCode(9733)（どれも 今 0）
 *    ★2026-10-09 本番前の 対立役が W1〜W7 の 7つの 穴を 作って 見せた（どれも 緑で 通った）＝全部 自己試験の 歯に した
 *
 *  ★★まだ 見て いない 物（2026-10-09 に 数えた）★★
 *    ・★AI に 渡す 字★ ･･･ api/claude.js が 読む prompt/ 6本の ★ 396・lib/formula-extra.js の 説明 41（make-prompt で prompt/kansuu.md へ）・
 *      lib/formula-soto.js の 頼み文 2・api/claude.js の 頼み文 6。AI の 答えは book.html の fetch から 画面に 出る（★を 抜く 所 0）。
 *      ★AI が ★ を 真似て 答えるかは 見立て（AI を 呼んで いない）★＝別件（棚）
 *    ・api/ の 下 ･･･ 見る ファイルの 外（★入りは 頼み文 6 だけ・客へ 返す 誤りの 字に ★ 0＝res.json 11か所）
 *    ・倉庫（Supabase）の 中身・画像 ･･･ 見て いない（見立て）
 *    ・似た 字 ☆（9）※（20）･･･ ★の 役が 移っても 見ない
 *    ・★2026-10-09 叩き直しの 対立役が 作った 残りの 抜け道（どれも 前から 緑・今の repo に 当てはまる 物 0）★
 *      A innerHTML に 入れた 属性の 中の 実体参照（title="&#9733;"）／B 割った 実体参照（'&#97' + '33;'）／
 *      C on 属性の 中の 2段の 実体参照（&amp;#9733;）／J CSS 変数を 通した content（--zz:"\2605"）／
 *      K JS が 書く CSS の 字（st.textContent='.a:after{content:...}'）／M 字の 決め打ちで ない fromCharCode（0x02605・9732+1・apply）／
 *      L manifest.json で ない 名前の 札（app.webmanifest）／D・E import './x.js'・export * from（module の script は 今 0本）
 *    ・★4回目の 対立役が 数えた 残り（今の repo に 当てはまる 物 どれも 0件）★
 *      M1 css の 中の @import で 読む css／M2 JS が 足す link（l.href='a.css'）＝読む css を 探すのは 置き場の 名と <link href> だけ
 *      M9 setAttribute('src', ...)／M10 new window.Worker(...)・new self.Worker(v)＝読み込みとして 拾わない
 *      止めすぎ M18 名前の 頭が 日本語（var 元src = 'x.js'）／M23 メソッドの 定義 { _loadScript(src){} }
 *    ・★5回目の 対立役が 数えた 残り（今の repo に 当てはまる 物 どれも 0件）★
 *      C1 content:"a;★"・content:"}★"（字の 中の ; } で 切れる）／C6 css の 逃がしで 書いた 名前 cont\65nt
 *      止めすぎ M8・O3〜O5 テンプレートの 字の 中の 'import(x)'・'_loadScript(v)'（テンプレートは 丸ごと 空けない＝わざと）／
 *      O1・O2 html の 注記や script の 注記の 中に 書いた <style>...★...</style>（<style> の 中身を 生で 拾う ため）
 *    ・★6回目の 対立役が 数えた 残り（今の repo に 当てはまる 物 どれも 0件）★
 *      S1・S2 テンプレートの ${...} の 中に ` が 在る（${"`"}・/`/g）と chuki が テンプレートの 終わりを 読み違え、
 *        後ろの " ' の 組が 入れ替わって import(v) を 空ける（★を 見る 道は 抜けない・引用符が 奇数なら「読めない 字」で 赤）。
 *        直すには chuki の 字の終わり に ${ の 深さが 要る＝自前の 深さ読みで 3回 抜けたので 今は 名指しだけ（${ を 含む テンプレートは 今 0本）
 *      読み込みの 形で 拾わない 物：_loadScript?.(v)・_loadScript.call(null, v)・別名（var L = _loadScript; L(v)）・
 *        置き場の 字＋変数で 名前を 作る _loadScript('extra/' + n + '.js')（白名簿に 無い 置き場の js を 読める 唯一の 素通り）。
 *        2026-10-09 に grep で 0件（_loadScript('...' + の 11件は 全部 'lib/x.js' + v＝名前は 字で 全部 見える）
 *    ⇒★「この 門が 0本」≠「お客さんの 画面に ★ が 無い」★
 *
 *  使い方: node scripts/hoshi-kazoeru.mjs           ... 数える（直さない）
 *          node scripts/hoshi-kazoeru.mjs --外す    ... 見つけた ★ を その 場で 外す（字は 残す）
 *  ★見張り★ tests/hoshi-nashi.test.mjs（--self-test の 歯も そちら）
 */
import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import { createRequire } from 'node:module';
import { fileURLToPath, pathToFileURL } from 'node:url';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const { 注記を外す, 字の塊を拾う } = await import(pathToFileURL(path.join(ROOT, 'scripts/lib/chuki.mjs')).href);
const require_ = createRequire(path.join(ROOT, 'package.json'));
const { JSDOM } = require_('jsdom');

/* ★口★＝名前と 出し方（innerHTML か textContent か）
     book.html の notify は `t.innerHTML = msg`・showToast は その 別名・窓の副題 は `e.textContent` */
export const 口の出し方 = { showToast: 'html', notify: 'html', 窓の副題: 'text' };

/* ══ JS の 字の 読み方 ══════════════════════════════════════ */
/** ★JS の 逃がしを 解く★（★ ／ \u{2605} ／ \x41 ／ \n など。逆斜線が 奇数個の 時だけ 逃がし） */
export function JSの逃がしを解く(生) {
  let 出 = '';
  for (let i = 0; i < 生.length; i++) {
    const c = 生[i];
    if (c !== '\\') { 出 += c; continue; }
    const d = 生[i + 1];
    if (d === undefined) { 出 += c; continue; }
    if (d === 'u' && 生[i + 2] === '{') {
      const j = 生.indexOf('}', i + 3);
      if (j > 0) { 出 += String.fromCodePoint(parseInt(生.slice(i + 3, j), 16)); i = j; continue; }
    }
    if (d === 'u' && /^[0-9a-fA-F]{4}$/.test(生.slice(i + 2, i + 6))) { 出 += String.fromCharCode(parseInt(生.slice(i + 2, i + 6), 16)); i += 5; continue; }
    if (d === 'x' && /^[0-9a-fA-F]{2}$/.test(生.slice(i + 2, i + 4))) { 出 += String.fromCharCode(parseInt(生.slice(i + 2, i + 4), 16)); i += 3; continue; }
    const 一 = { n: '\n', t: '\t', r: '\r', b: '\b', f: '\f', v: '\v', 0: '\0' };
    出 += Object.prototype.hasOwnProperty.call(一, d) ? 一[d] : d;   /* \\ は \ ／ \' は ' */
    i++;
  }
  return 出;
}

let _道具 = null;
/** ★実体参照を 解く★（ブラウザと 同じ 読み方＝jsdom。&amp;#9733; は 1回だけ 解いて 「&#9733;」の 字） */
export function 実体参照を解く(s) {
  if (!_道具) _道具 = new JSDOM('<!DOCTYPE html><div></div>').window.document.querySelector('div');
  _道具.innerHTML = String(s);
  return _道具.textContent;
}

/** ★文字列の 塊を 読み進める★＝i は 引用符の 位置。{ 終わり, 地 } を 返す（地＝テンプレートの ${...} を 除いた 字） */
function 文字列を読む(src, i) {
  const q = src[i];
  let j = i + 1, 地 = '';
  while (j < src.length) {
    const c = src[j];
    if (c === '\\') { 地 += c + (src[j + 1] || ''); j += 2; continue; }
    if (c === q) return { 終わり: j + 1, 地 };
    if (q === '`' && c === '$' && src[j + 1] === '{') {
      /* ${ ... } を 飛ばす（中の 文字列や 括弧も 数える） */
      let 深 = 1; j += 2;
      while (j < src.length && 深 > 0) {
        const e = src[j];
        if (e === "'" || e === '"' || e === '`') { j = 文字列を読む(src, j).終わり; continue; }
        if (e === '{') 深++;
        else if (e === '}') 深--;
        j++;
      }
      continue;
    }
    地 += c; j++;
  }
  return { 終わり: j, 地 };
}

/** ★口の 第1引数 まるごと★＝括弧の 対を 数えて 取る。中の 文字列の 塊を 返す [{ 開始, 終わり, 地 }] */
export function 第1引数の文字列(src, 開き) {
  const 塊 = [];
  let j = 開き + 1, 深 = 0;
  while (j < src.length) {
    const c = src[j];
    if (c === "'" || c === '"' || c === '`') {
      const r = 文字列を読む(src, j);
      塊.push({ 開始: j, 終わり: r.終わり, 地: r.地 });
      j = r.終わり; continue;
    }
    if (c === '(' || c === '[' || c === '{') 深++;
    else if (c === ')' || c === ']' || c === '}') { if (深 === 0) break; 深--; }
    else if (c === ',' && 深 === 0) break;
    j++;
  }
  return 塊;
}

/** ★口に 渡る 字★（注記を 外した 字を 渡す＝位置は 元の 字と 同じ）
 *  @param {{復号?: boolean}} opt  復号＝false は 歯の 比べ 用（逃がしも 実体参照も 解かない） */
export function 口に渡る字(注記なしの字, opt = {}) {
  const 復号 = opt.復号 !== false;
  const 出 = [];
  for (const k of Object.keys(口の出し方)) {
    /* ★`\b` は ASCII の 区切りしか 見ない★＝日本語の 名前（窓の副題）の 前では 効かない ⇒ ASCII の 時だけ */
    const 頭 = /^[A-Za-z_$]/.test(k) ? '\\b' : '';
    const re = new RegExp(頭 + k + '\\s*\\(', 'g');
    let m;
    while ((m = re.exec(注記なしの字))) {
      const 開き = m.index + m[0].length - 1;
      const 塊 = 第1引数の文字列(注記なしの字, 開き);
      if (!塊.length) continue;
      let 字 = 塊.map((x) => (復号 ? JSの逃がしを解く(x.地) : x.地)).join('');
      if (復号 && 口の出し方[k] === 'html') 字 = 実体参照を解く(字);
      出.push({ 口: k, 字, 塊 });
    }
  }
  return 出;
}

/** ★★ が 混じっている 物だけ 返す */
export function 星入り(渡る字) {
  return 渡る字.filter((x) => x.字.indexOf('★') >= 0);
}

/* ══ HTML の 見える 字 ══════════════════════════════════════ */
const 見えない属性 = /^(style|class|id|href|src)$|^on/i;
/** ★HTML の 見える 字★＝本文（script／style の 外）と 属性（style／class／id／href／src／on... を 除く）
 *  ★隠れて いる 所も 数える★（窓は JS が 開く）。返す 形 [{ 種, 字, 開始, 終わり }]（位置は 元の 字）
 *  @param {{復号?: boolean}} opt  復号＝false は 歯の 比べ 用（元の 字の まま＝実体参照を 解かない） */
export function HTMLの見える字(html, opt = {}) {
  const 復号 = opt.復号 !== false;
  const dom = new JSDOM(html, { includeNodeLocations: true });
  const doc = dom.window.document;
  const 出 = [];
  const 歩く = (n) => {
    if (n.nodeType === 3) {
      const p = n.parentNode && n.parentNode.nodeName;
      if (p === 'SCRIPT' || p === 'STYLE') return;
      const 位 = dom.nodeLocation(n);
      if (!位) return;
      出.push({ 種: '本文', 字: 復号 ? n.data : html.slice(位.startOffset, 位.endOffset), 開始: 位.startOffset, 終わり: 位.endOffset });
      return;
    }
    if (n.nodeType === 1) {
      const 位 = dom.nodeLocation(n);
      for (const a of Array.from(n.attributes || [])) {
        if (見えない属性.test(a.name)) continue;
        const 属位 = 位 && 位.attrs && 位.attrs[a.name];
        const 生 = 属位 ? html.slice(属位.startOffset, 属位.endOffset) : a.value;
        出.push({ 種: '属性 ' + a.name, 字: 復号 ? a.value : 生, 開始: 属位 ? 属位.startOffset : -1, 終わり: 属位 ? 属位.endOffset : -1 });
      }
    }
    for (const c of Array.from(n.childNodes || [])) 歩く(c);
    /* ★<template> の 中身は childNodes に 居ない★（.content の 側）＝JS が 写して 出す（今 0個） */
    if (n.nodeName === 'TEMPLATE' && n.content) for (const c of Array.from(n.content.childNodes)) 歩く(c);
  };
  歩く(doc);
  dom.window.close();
  return 出;
}

/* ══ 見る ファイル（★名前で 選ばない★） ═══════════════════════ */
/* ★借り物★＝他人の 字（文字コード表の 本物の 記号が 入る）。★名指しで 除く★
     （前は「.min. と いう 名前」で 除いて いた＝自前の ファイルを .min.js と 名付けると 黙って 外れる・2026-10-09 対立役） */
export const 借り物 = ['hyperformula.full.min.js', 'lib/xlsx.full.min.js'];
export function 見るファイル(根 = ROOT) {
  const 出 = { html: [], js: [] };
  const 借り = (p) => 借り物.indexOf(path.relative(根, p).split(path.sep).join('/')) >= 0;
  for (const f of fs.readdirSync(根).sort()) {
    const p = path.join(根, f);
    if (!fs.statSync(p).isFile()) continue;
    if (/\.html$/i.test(f)) 出.html.push(p);
    else if (/\.(js|mjs)$/.test(f) && !借り(p)) 出.js.push(p);
  }
  const 潜る = (d) => {
    if (!fs.existsSync(d)) return;
    for (const f of fs.readdirSync(d).sort()) {
      const p = path.join(d, f);
      if (fs.statSync(p).isDirectory()) 潜る(p);
      else if (/\.(js|mjs)$/.test(f) && !借り(p)) 出.js.push(p);
    }
  };
  潜る(path.join(根, 'js'));
  潜る(path.join(根, 'lib'));
  return 出;
}

/* ══ JS の 字 まるごと（2026-10-09 に 広げた） ═══════════════════ */
/* ★口ごとに 足す 形（黒名簿）は 終わらない★＝textContent・innerHTML・lib が 返す 訳・throw の 字...
     ⇒ ★注記の 外の 字の塊（'...' "..." `...`）は 全部 見る★。客に 出ないと ★機械が 毎回 確かめた★ 物と、
       白名簿に 名指しした 物だけ 除く。
   ★白名簿★＝1件ずつ「置き場・種・数・指紋・訳」。★数か 指紋が 合わなければ 赤★（足されても 消えても 入れ替えても 黙らない）
     指紋＝その 置き場・種の ★入りの 字の塊を 並びの 順に つないだ sha256 の 頭16字
     （数だけ だと「AI への 字を 1つ 消して 客の 字を 1つ 足す」が 緑で 通った＝2026-10-09 対立役が 作った） */
export const 白名簿 = [
  { 名: 'lib/symbols.js', 種: '字そのもの', 数: 1, 指紋: '6dc562f4adea13d4',
    訳: '記号を 入れる 窓の 並び＝★ は 記号 そのもの（お客さんが 選んで 入れる 字）' },
  { 名: 'lib/ribbon-spec.js', 種: '字そのもの', 数: 1, 指紋: '6dc562f4adea13d4',
    訳: '表示 タブ「ユーザー設定のビュー」の 印（lib/ribbon.js が rb-ico に 描く）。★実Excel の 印の 形は 未測定★' },
  { 名: 'lib/formula-extra.js', 種: 'AIへ', 数: 41, 指紋: '5f62a90ab6a429a7',
    訳: '形で動く の 説明＝scripts/make-prompt.mjs が prompt/kansuu.md に 入れて ★AI に 渡す★（画面へ 直に 出す 口は 0＝2026-10-09 対立役が 辿った）。AI の 道は 別件（下の「まだ 見て いない 物」）' },
  { 名: 'lib/formula-soto.js', 種: 'AIへ', 数: 2, 指紋: '323a019940d31c52',
    訳: 'AI への 頼み文（訳して／何語か）。AI の 道は 別件' },
  { 名: 'lib/yakusoku-daicho.js', 種: '読み込まれていない', 数: 9, 指紋: '499fb5051e20c777',
    訳: '約束の 台帳＝tests/yakusoku.test.mjs だけが 読む（どの html も js も 読み込まない＝機械が 毎回 確かめる）' },
];
export const 白の指紋 = (塊ら) => crypto.createHash('sha256').update(塊ら.map((x) => x.生).join('\u0000')).digest('hex').slice(0, 16);

/** ★読み込まれる js★（src= ・ _loadScript(...) ・ import ・ new Worker ・ importScripts）＝repo の 中の 物を 相対の 名で
 *  ★html だけで なく js も 見る★（js が js を 読む 道＝2026-10-09 対立役が 作った 穴。今 0本）
 *  引用符は ' " ` の どれでも（テンプレートの _loadScript(`...`) が 素通り した＝同じ 日） */
export function 読み込まれる物(根 = ROOT) {
  const 出 = new Set();
  const 名簿 = 見るファイル(根);
  for (const p of 名簿.html.concat(名簿.js)) {
    const s = 注記を外す(fs.readFileSync(p, 'utf8'), { html: /\.html$/i.test(p) });
    /* 頭は「名前の 字で ない 字」か 行頭（my_loadScript(...) を 読み込みと 取らない）。★後読みは 使わない★
       「.」を 許さないのは import だけ（X.import(d)）。el.src= ・window._loadScript ・self.importScripts は 普通の 形
       （「.」を 全部 外したら この 3形が 緑に なった＝2026-10-09 3回目の 対立役） */
    const re = /(?:(?:^|[^\w$])(?:src\s*=\s*|_loadScript\s*\(\s*|new\s+Worker\s*\(\s*|importScripts\s*\(\s*)|(?:^|[^\w$.])(?:import\s*\(\s*|import\s+[^'"`;]*?from\s*))["'`]([^"'`?#]+\.m?js)/g;
    let m;
    while ((m = re.exec(s))) {
      if (/^(https?:)?\/\//.test(m[1])) continue;
      /* 頭が / の 物は 頁の 根から（Windows の 道として 解かない＝2026-10-09 対立役） */
      const 字 = m[1].replace(/^\/+/, '');
      const 基 = /\.html$/i.test(p) || !/^\.\.?\//.test(字) ? 根 : path.dirname(p);
      出.add(path.relative(根, path.resolve(基, 字)).split(path.sep).join('/'));
    }
  }
  return 出;
}


/** ★名前で 読み込む 所が 字で ない★（_loadScript(名) の 様に 変数で 渡す）＝何を 読むか 字から 決まらない＝赤（今 0） */
export function 字で無い読み込み(根 = ROOT) {
  const 出 = [];
  const 名簿 = 見るファイル(根);
  for (const p of 名簿.html.concat(名簿.js)) {
    /* ★字の 中身を 空けて から 見る★（'import(x)' の 様な 字の 中を 読み込みと 取らない＝2026-10-09 対立役） */
    const 生 = fs.readFileSync(p, 'utf8');
    const html = /\.html$/i.test(p);
    let 空けた = 生;
    /* ★空けるのは ' と " の 字だけ★。テンプレート（`...`）は ${...} の 中が 動く コードなので ★丸ごと 残す★
         （${...} を 自前の 読み方で 分けたら 字・注記・正規表現の 中の { } " / で 3回 続けて 抜け道を 作った＝2026-10-09 3〜5回目の 対立役。
           repo の テンプレートは 4本・全部 1行＝止めすぎ（テンプレートの 地の 字の 'import(x)'）に 倒れる 方を 選ぶ） */
    for (const x of 字の塊を拾う(生, { html }).塊.reverse()) {
      if (x.生[0] === '`') continue;
      空けた = 空けた.slice(0, x.開始 + 1) + x.生.slice(1, -1).replace(/[^\n]/g, ' ') + 空けた.slice(x.終わり - 1);
    }
    const s = 注記を外す(空けた, { html });
    const re = /(?:(?:^|[^\w$])(?:_loadScript|new\s+Worker|importScripts)|(?:^|[^\w$.])import)\s*\(\s*(?!["'`])([^)\s]{1,40})/g;
    let m;
    while ((m = re.exec(s))) {
      /* 定義（function _loadScript(src)）は 読み込みで ない。★後読み (?<!) は 使わない★（tests/no-lookbehind） */
      if (/\bfunction\s*$/.test(s.slice(Math.max(0, m.index - 20), m.index + 1))) continue;
      if (/^function\b|^\)/.test(m[1])) continue;
      出.push({ 名: path.relative(根, p).split(path.sep).join('/'), 字: m[0].slice(0, 60), 行: s.slice(0, m.index).split('\n').length });
    }
  }
  return 出;
}

/** ★JS の 字の塊の ★★＝[{ 名, 種, 生, 開始, 終わり }]（種＝赤／console／読み込まれていない／字そのもの／AIへ）と 読めない 所 */
export function JSの字の星(根 = ROOT) {
  const 名簿 = 見るファイル(根);
  const 読まれる = 読み込まれる物(根);
  const 名 = (p) => path.relative(根, p).split(path.sep).join('/');
  const 出 = [], 読めない = [], 行またぎ = [];
  let 塊の数 = 0;
  for (const p of 名簿.html.concat(名簿.js)) {
    const 生 = fs.readFileSync(p, 'utf8');
    const html = /\.html$/i.test(p);
    const r = 字の塊を拾う(生, { html });
    塊の数 += r.塊.length;
    for (const k of r.読めない) 読めない.push({ 名: 名(p), 行: 生.slice(0, k).split('\n').length });
    for (const x of r.塊) {
      /* ★テンプレートが 行を またいだら 赤★＝正規表現の 中の ` を 字の 始まりと 取り違えて 何千行も 飲み込んだ 疑い
           （今 テンプレートは 4本・全部 1行＝2026-10-09 に 数えた。飲み込みは「閉じない」には 出ない） */
      if (x.生[0] === '`' && x.生.indexOf('\n') >= 0) 行またぎ.push({ 名: 名(p), 行: 生.slice(0, x.開始).split('\n').length });
      /* ★逃がし（★）も 実体参照（&#9733;・&starf;）も 解いて 見る★＝innerHTML に 入れば ★ に なる
           （実体参照を 解かず &#9733; が 緑で 通った＝2026-10-09 対立役が 作った。& の 在る 字だけ 解く） */
      const 解いた = JSの逃がしを解く(x.生);
      if (解いた.indexOf('★') < 0 && !(解いた.indexOf('&') >= 0 && 実体参照を解く(解いた).indexOf('★') >= 0)) continue;
      let 種 = '赤';
      const 読まれない = !html && !読まれる.has(名(p));
      const 白 = 白名簿.find((w) => w.名 === 名(p) && (w.種 !== '読み込まれていない' || 読まれない));
      /* 前の 字は ★元の 字★で 見る（注記を 外すと 絵文字で 長さが 変わり 位置が ずれる）。
         console の 前は 名前の 字で ない 事（myconsole.log が 通った＝2026-10-09） */
      if (/(?:^|[^\w$.])console\.\w+\s*\(\s*$/.test(生.slice(Math.max(0, x.開始 - 40), x.開始))) 種 = 'console';
      else if (白) 種 = 白.種;
      出.push({ 名: 名(p), 種, 生: x.生, 開始: x.開始, 終わり: x.終わり, 行: 生.slice(0, x.開始).split('\n').length });
    }
  }
  /* ★白名簿の 数★＝名指しの 数と 今の 数が 違えば 腐って いる */
  const 白の数 = 白名簿.map((w) => {
    const 塊ら = 出.filter((x) => x.名 === w.名 && x.種 === w.種);
    return { ...w, 今: 塊ら.length, 今の指紋: 白の指紋(塊ら) };
  });
  return { 出, 読めない, 行またぎ, 白の数, 読まれる, 塊の数 };
}

/** ★全部 数える★＝[{ 名, 置き場, 種, 字, 開始, 終わり }]（★ 入りだけ）と 数 */
export function 全部数える(根 = ROOT) {
  const 名簿 = 見るファイル(根);
  const 赤 = [];
  const 数 = { 口の字: 0, 本文: {}, 属性: 0, ファイル: 名簿.html.length + 名簿.js.length };
  const 名 = (p) => path.relative(根, p).split(path.sep).join('/');
  for (const p of 名簿.html.concat(名簿.js)) {
    const 生 = fs.readFileSync(p, 'utf8');
    const 外した = 注記を外す(生);
    const 渡る = 口に渡る字(外した);
    数.口の字 += 渡る.length;
    /* ★口の 塊は「字」で 持つ★（2026-10-08）＝注記を 外すと 4バイトの 字（絵文字）で 長さが 変わる 事が ある（book.html は 7字 短い）
         ⇒ 位置は 元の 字と 合わない。外す 時は 字で 探す */
    for (const x of 星入り(渡る)) 赤.push({ 名: 名(p), 置き場: '口 ' + x.口, 字: x.字, 文字列: x.塊.map((b) => 外した.slice(b.開始, b.終わり)) });
    if (/\.html$/i.test(p)) {
      const 見える = HTMLの見える字(生);
      数.本文[名(p)] = 見える.filter((x) => x.種 === '本文').length;
      数.属性 += 見える.filter((x) => x.種 !== '本文').length;
      for (const x of 見える) if (x.字.indexOf('★') >= 0) 赤.push({ 名: 名(p), 置き場: x.種, 字: x.字, 塊: [{ 開始: x.開始, 終わり: x.終わり }] });
      /* ★on 属性の 中の JS★（onclick="x.textContent='★'"）＝<script> の 外なので JS の 字の 見方に 入らない（今 0） */
      for (const x of on属性の星(生)) 赤.push({ 名: 名(p), 置き場: 'on 属性 ' + x.名, 字: x.字 });
    }
  }
  /* ── JS の 字 まるごと（2026-10-09） ── */
  const J = JSの字の星(根);
  数.JSの字の星 = J.出.length;
  数.字の塊 = J.塊の数;
  数.白の数 = J.白の数;
  数.JSの種 = {};
  for (const x of J.出) 数.JSの種[x.種] = (数.JSの種[x.種] || 0) + 1;
  for (const x of J.出) if (x.種 === '赤') 赤.push({ 名: x.名, 置き場: 'JS の 字 ' + x.行 + '行', 字: x.生, 塊: [{ 開始: x.開始, 終わり: x.終わり }] });
  for (const x of J.読めない) 赤.push({ 名: x.名, 置き場: '読めない 字 ' + x.行 + '行', 字: '（引用符が 閉じない＝正規表現と 取り違えた 疑い）' });
  for (const x of J.行またぎ) 赤.push({ 名: x.名, 置き場: '行を またぐ テンプレート ' + x.行 + '行', 字: '（飲み込みの 疑い）' });
  for (const w of J.白の数) if (w.今 !== w.数 || w.今の指紋 !== w.指紋) 赤.push({ 名: w.名, 置き場: '白名簿の 数', 字: w.種 + '＝名指し ' + w.数 + '（' + w.指紋 + '） ／ 今 ' + w.今 + '（' + w.今の指紋 + '）' });
  for (const x of 字で無い読み込み(根)) 赤.push({ 名: x.名, 置き場: '字で 無い 読み込み ' + x.行 + '行', 字: x.字 });
  /* ★字の 形で ない ★★（String.fromCharCode(9733) ・ fromCodePoint(0x2605)）＝字の塊に 出ない（今 0） */
  for (const p of 名簿.html.concat(名簿.js)) {
    const s = 注記を外す(fs.readFileSync(p, 'utf8'), { html: /\.html$/i.test(p) });
    for (const m of s.matchAll(/fromCharCode|fromCodePoint/g)) {
      const 引数 = s.slice(m.index, m.index + 80);
      if (/\(\s*(9733|0x2605|0X2605)\b/.test(引数)) 赤.push({ 名: 名(p), 置き場: '字の 形で ない ★', 字: 引数.slice(0, 40) });
    }
  }
  /* ★CSS の content★（html の <style>・style 属性・*.css）と ★manifest.json の 字★（入れた 時の 名前に 出る）＝今 0 */
  /* ★注記を 外して から★・★名前の 頭で 切って★ 見る（justify-content: や 注記の textContent: で 赤に なった＝2026-10-09 対立役） */
  /* *.css は ★ブロックの 注記 だけ★ が 注記（// は url(//...) の 中に 在る＝JS の 読み方で 外すと 行の 残りが 消えた＝2026-10-09 3回目の 対立役）
       ⇒ chuki の <style> の 読み方（ブロックの 注記だけ）を 借りる */
  /* ★<!-- と --> は css では 捨てられる 印 だけ＝間の 決まりは 生きる★（chuki の html の 読み方は <style> の 中の <!-- --> も 空けて
       *.css と <style> の 間の ★ を 見落とした＝2026-10-09 4回目の 対立役）⇒ css は ブロックの 注記だけを 自分で 外す。
       html は 今まで 通り（属性・script の 字の 中の css も 見る）に 加えて <style> の 中身を 生の まま css として 見る */
  /* 字（' "）は 飛ばして から ブロックの 注記を 外す（content:"/*" の 字の 中を 注記と 取り、間の ★ を 消した＝2026-10-09 5回目の 対立役） */
  const css注記を外す = (s) => s.replace(/("(?:\\[\s\S]|[^"\\\n])*"|'(?:\\[\s\S]|[^'\\\n])*')|\/\*[\s\S]*?\*\//g, (m, 字) => 字 || m.replace(/[^\n]/g, ' '));
  const CSSの字 = (s, html) => html
    ? 注記を外す(s, { html: true }) + '\n' + [...s.matchAll(/<style\b[^>]*>([\s\S]*?)<\/style>/gi)].map((m) => css注記を外す(m[1])).join('\n')
    : css注記を外す(s);
  const CSSの星 = (s, html) => (CSSの字(s, html).match(/(?:^|[^\w-])content\s*:[^;}]*/gi) || []).filter((c) => c.indexOf('★') >= 0 || /\\0*2605/i.test(c));
  const css置き場 = [];
  /* 見る 所は 配られる 置き場（直下・css/・js/・lib/）だけ＝tests の 下の 見本や 手元の 報告の css を 拾わない */
  const 潜るcss = (d) => { if (!fs.existsSync(d)) return; for (const f of fs.readdirSync(d)) { const q = path.join(d, f); if (fs.statSync(q).isDirectory()) { if (f !== 'node_modules' && f[0] !== '.') 潜るcss(q); } else if (/\.css$/i.test(f)) css置き場.push(q); } };
  for (const f of fs.readdirSync(根)) if (/\.css$/i.test(f) && fs.statSync(path.join(根, f)).isFile()) css置き場.push(path.join(根, f));
  for (const d of ['css', 'js', 'lib']) 潜るcss(path.join(根, d));
  /* ★html の <link href> が 読む css も★（置き場の 名で 決めない＝assets/a.css が 緑に なった＝2026-10-09 3回目の 対立役） */
  for (const p of 名簿.html) {
    for (const m of 注記を外す(fs.readFileSync(p, 'utf8'), { html: true }).matchAll(/<link\b[^>]*\bhref\s*=\s*["']?([^"'\s>?#]+\.css)/gi)) {
      if (/^(https?:)?\/\//.test(m[1])) continue;
      const q = path.resolve(根, m[1].replace(/^\/+/, ''));
      if (fs.existsSync(q) && css置き場.indexOf(q) < 0) css置き場.push(q);
    }
  }
  for (const p of css置き場.concat(名簿.html)) for (const c of CSSの星(fs.readFileSync(p, 'utf8'), /\.html$/i.test(p))) 赤.push({ 名: 名(p), 置き場: 'CSS の content', 字: c.slice(0, 60) });
  /* ★白名簿の「読み込まれていない」は 名前の 側から 引く★＝その 名前（拡張子 無し）が 注記の 外に 1つでも 出たら 赤
       （拾い方を 足し続けても _loadScript('lib/' + 名)・src=名・_loadScript (名) が 抜けた＝2026-10-09 対立役） */
  for (const w of 白名簿.filter((x) => x.種 === '読み込まれていない')) {
    const 素 = path.basename(w.名).replace(/\.m?js$/, '');
    for (const p of 名簿.html.concat(名簿.js)) {
      if (名(p) === w.名) continue;
      if (注記を外す(fs.readFileSync(p, 'utf8'), { html: /\.html$/i.test(p) }).indexOf(素) >= 0) 赤.push({ 名: 名(p), 置き場: '読み込まれていない 筈の 名前が 出ている', 字: 素 + '（' + w.名 + '）' });
    }
  }
  const 札 = path.join(根, 'manifest.json');
  if (fs.existsSync(札)) {
    const 歩く = (v) => { if (typeof v === 'string') { if (v.indexOf('★') >= 0) 赤.push({ 名: 'manifest.json', 置き場: 'manifest の 字', 字: v }); } else if (v && typeof v === 'object') for (const k of Object.keys(v)) 歩く(v[k]); };
    歩く(JSON.parse(fs.readFileSync(札, 'utf8')));
  }
  数.css = css置き場.length;
  /* ★読み込まれるのに 見て いない js★＝置き場の 名で 見る 範囲を 決めて いるので、新しい 置き場は 黙って 外れる */
  const 見る名 = new Set(見るファイル(根).js.map(名));
  for (const f of J.読まれる) if (!見る名.has(f) && 借り物.indexOf(f) < 0) 赤.push({ 名: f, 置き場: '読み込まれるのに 見て いない', 字: '' });
  数.読み込まれる = J.読まれる.size;
  return { 赤, 数 };
}

/** ★on 属性の 中の ★★（JS として 動く 字）。逃がしも 解く */
export function on属性の星(html) {
  const dom = new JSDOM(html);
  const 出 = [];
  for (const el of Array.from(dom.window.document.querySelectorAll('*'))) {
    for (const a of Array.from(el.attributes)) if (/^on/i.test(a.name) && JSの逃がしを解く(a.value).indexOf('★') >= 0) 出.push({ 名: a.name, 字: a.value });
  }
  dom.window.close();
  return 出;
}

/* ★★取り込まれた 時は 走らない★★（2026-09-06 実測で 踏んだ・記憶 feedback_global_tool_must_not_judge_itself_by_argv）
     tests/hoshi-nashi.test.mjs が この ファイルを import する ＝★自分が 直に 走らされた 時だけ 動く★ */
const 直に走った = process.argv[1] && pathToFileURL(process.argv[1]).href === import.meta.url;

if (直に走った) {
  const { 赤, 数 } = 全部数える();
  console.log('\n[hoshi-kazoeru] お客さんの 画面に 出る 字の ★ を 数える');
  console.log('  見た ファイル ' + 数.ファイル + '本 ／ 口の 字 ' + 数.口の字 + '本 ／ HTML の 本文 ' + JSON.stringify(数.本文) + ' ／ 属性 ' + 数.属性);
  console.log('  JS の 字の塊 ' + 数.字の塊 + '個 ／ その中の ★入り ' + 数.JSの字の星 + '個 ' + JSON.stringify(数.JSの種) + ' ／ html が 読み込む js ' + 数.読み込まれる + '本');
  console.log('  ★ が 混じっている 物 ･･･ ★' + 赤.length + '★');
  for (const x of 赤.slice(0, 60)) console.log('    ' + x.名 + ' [' + x.置き場 + '] ' + x.字.replace(/\s+/g, ' ').trim().slice(0, 70));
  if (process.argv.includes('--外す')) {
    /* ★見つけた ★ を その 場で 外す★（字は 残す）。位置は 元の 字＝後ろから 消す */
    /* ①HTML の 見える 字＝jsdom が 元の 字で 読んだ 位置（後ろから）②口の 文字列＝字で 探す（位置は 使わない） */
    const 名ら = [...new Set(赤.map((x) => x.名))];
    for (const f of 名ら) {
      const p = path.join(ROOT, f);
      let s = fs.readFileSync(p, 'utf8');
      const 前の長さ = s.length;
      const 範囲 = 赤.filter((x) => x.名 === f && x.塊).flatMap((x) => x.塊).filter((b) => b.開始 >= 0).sort((a, b) => b.開始 - a.開始);
      for (const b of 範囲) s = s.slice(0, b.開始) + s.slice(b.開始, b.終わり).split('★').join('') + s.slice(b.終わり);
      for (const x of 赤.filter((y) => y.名 === f && y.文字列)) {
        for (const 元 of x.文字列) { if (元.indexOf('★') >= 0) s = s.split(元).join(元.split('★').join('')); }
      }
      fs.writeFileSync(p, s);
      console.log('  外した ' + f + ' ･･･ ★ ' + (前の長さ - s.length) + '個');
    }
  }
  console.log('\n  ★この 門が 見て いない 置き場が 在ります★（上の 注記の「まだ 見て いない 物」）');
}
