/* hoshi-nashi.test.mjs — ★お客さんの 画面に 出る 字に ★ を 使わない★（2026-09-06・2026-10-08 に 広げた）
 *
 *  ★★なぜ★★
 *    ★ は ★私たちの 便りの 印★（司さん・指示役・私の やりとり／コメント／書類）。
 *    ★お客さんには 壊れた 字に 見えます★。
 *    2026-09-06 ... 書き出しの 窓の 字に 12個以上 出ていて 指示役に 止められた。
 *    同じ日に Rakunally も ★23件★ 直している（会社ぜんぶの 決まり）。
 *
 *  ★★上限を 付けない（0本）★★（2026-09-06 指示役の 裁定）
 *    ★上限が 死ぬ 道は 2つ★ ①下げれば 通る ②減らす 途中の 数が 天井に なる ⇒★0本★
 *
 *  ★★2026-10-08 に 広げた★★（Rakunally の 引き継ぎ＝司さん 指示・対立役が 叩いた）
 *    前は「口を 通る 字の ★同じ行★」だけで、★HTML の 本文を 見ず★、門は 緑の まま
 *    お客さんの 画面に ★ が 49本 出て いた（窓の 説明 41節＋行を またぐ 口 8本）。
 *    ★この 見張りが 守る 範囲★ ＝ scripts/hoshi-kazoeru.mjs の 注記（口を 通る 字・HTML の 中の 字・
 *      lib が 返して 画面が 出す 字は まだ・行を またぐ 連結は 括弧の 対で 取る）
 *    ★まだ 見て いない 置き場が 在る★（JS が 直に 書く 字 47・lib が 返す 字 8 など＝道具の 注記に 数）
 *    ⇒★「0本に なったから 終わり」に しない★
 *
 *  使い方: node tests/hoshi-nashi.test.mjs
 *          node tests/hoshi-nashi.test.mjs --self-test
 *          node tests/hoshi-nashi.test.mjs --list   ... ★赤の 一覧★
 */
import fs from 'node:fs';
import os from 'node:os';
import { fileURLToPath, pathToFileURL } from 'node:url';
import path from 'node:path';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const { 注記を外す } = await import(pathToFileURL(path.join(ROOT, 'scripts/lib/chuki.mjs')).href);
const 道具 = await import(pathToFileURL(path.join(ROOT, 'scripts/hoshi-kazoeru.mjs')).href);
const { 口に渡る字, 星入り, HTMLの見える字, 全部数える, 見るファイル, 白名簿 } = 道具;

let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ✓ ' + n); } catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + (e && e.message)); } };

/* ★JS の 字の 歯★＝仮の repo（html ＋ lib）を 作って ★本物の 全部数える★ に 通す（2026-10-09）
     返す＝赤の 置き場の 頭（「JS の 字」「読めない 字」...）の 数。白名簿の「数が 違う」は 仮の repo では 必ず 出るので 除く */
const 仮のrepo = (files) => {
  const d = fs.mkdtempSync(path.join(os.tmpdir(), 'hoshi-'));
  for (const [f, s] of Object.entries(files)) { fs.mkdirSync(path.dirname(path.join(d, f)), { recursive: true }); fs.writeFileSync(path.join(d, f), s); }
  try { return 全部数える(d).赤.filter((x) => x.置き場 !== '白名簿の 数'); } finally { fs.rmSync(d, { recursive: true, force: true }); }
};
const 頁 = (js, 読む = []) => '<!DOCTYPE html><html><body><p>あ</p>' + 読む.map((f) => '<script src="' + f + '?v=1"></script>').join('') + '<script>' + js + '</script></body></html>';
const JSの赤 = (files, 頭) => 仮のrepo(files).filter((x) => x.置き場.indexOf(頭) === 0).length;

/* ★歯★＝本物の 抽出の 道（口に渡る字・HTMLの見える字）を 通して 星が 在るか */
const 口の星 = (src, opt) => 星入り(口に渡る字(注記を外す(src), opt)).length;
const HTMLの星 = (html, opt) => HTMLの見える字(html, opt).filter((x) => x.字.indexOf('★') >= 0).length;

if (process.argv.includes('--self-test')) {
  console.log('\n[hoshi-nashi --self-test] ★本物の 抽出の 道を 通る 入力で 歯を 確かめる★');
  const 赤 = (n, v) => T(n + ' ⇒ 赤', () => { if (!(v > 0)) throw new Error('見つけて いない（' + v + '）'); });
  const 緑 = (n, v) => T(n + ' ⇒ 赤に しない', () => { if (v !== 0) throw new Error('拾いすぎ（' + v + '）'); });
  /* ── 口 ── */
  赤('口 notify(\'★あ★\')', 口の星("notify('★あ★');"));
  緑('口 notify(\'あ\')', 口の星("notify('あ');"));
  緑('口で ない console.log(\'★\')', 口の星("console.log('★あ★');"));
  緑('注記の 中の ★', 口の星("/* notify('★') */ notify('あ'); // notify('★')"));
  赤('窓の副題 も 口', 口の星("窓の副題('★', 'name-insert');"));
  赤('三項 演算子で 行を またぐ（2026-10-08 に 8本 素通りして いた 形）', 口の星("notify(x\n  ? '★あ' : 'い');"));
  赤('改行して + で つなぐ', 口の星("notify('あ'\n  + '★い');"));
  赤('テンプレート 文字列', 口の星('notify(`★${x}`);'));
  赤('JS の 逃がし \\u2605（逆斜線 1つ）', 口の星("notify('\\u2605あ');"));
  赤('JS の 逃がし \\u{2605}', 口の星("notify('\\u{2605}あ');"));
  緑('逃がしの 逆斜線を 逃がした \\\\u2605（字は「逆斜線 u2605」）', 口の星("notify('\\\\u2605');"));
  赤('逆斜線 3つ（逆斜線＋★）', 口の星("notify('\\\\\\u2605');"));
  赤('notify は innerHTML＝実体参照 &#9733; も ★', 口の星("notify('&#9733;あ');"));
  緑('窓の副題 は textContent＝&#9733; は 字の まま', 口の星("窓の副題('&#9733;');"));
  緑('&amp;#9733; は ★ に ならない（1回だけ 解く）', 口の星("notify('&amp;#9733;');"));
  /* ── HTML ── */
  赤('本文 <div>★注意</div>', HTMLの星('<div>★注意</div>'));
  赤('隠れた 窓の 中（hidden）＝JS が 開く ので 見える', HTMLの星('<div hidden><p>★注意</p></div>'));
  赤('display:none の 中も 見える', HTMLの星('<div style="display:none">★注意</div>'));
  緑('HTML の 注記 <!--★-->', HTMLの星('<!--★--><div>あ</div>'));
  緑('script の 中', HTMLの星("<script>var a='★';</script><div>あ</div>"));
  緑('style の 中', HTMLの星('<style>.a:after{content:"★"}</style><div>あ</div>'));
  赤('実体参照 &#9733;', HTMLの星('<div>&#9733;注意</div>'));
  赤('実体参照 &#x2605; ／ 頭に 0 ／ ; 無し ／ 名前 &starf;', HTMLの星('<div>&#x2605;&#09733;&#9733 &starf;</div>'));
  緑('&amp;#9733; は 字で「&#9733;」', HTMLの星('<div>&amp;#9733;</div>'));
  赤('属性 title="★"', HTMLの星('<div title="★あ">あ</div>'));
  赤('名簿に 無い 属性も 見る（data-tip="★"）', HTMLの星('<div data-tip="★あ">あ</div>'));
  緑('style 属性の 中', HTMLの星('<div style="--x:\'★\'">あ</div>'));
  赤('textarea と option の 字も 見える', HTMLの星('<textarea>★</textarea><select><option>★</option></select>'));
  /* ── 読み解きを 外すと 答えが 変わる（★復号が 効いて いる 証し★＝同じ 入力を 両方に 通す） ── */
  T('★口：逃がしを 解かないと \\u2605 を 見逃す（解く 1 ／ 解かない 0）★', () => {
    const 解 = 口の星("notify('\\u2605');"), 素 = 口の星("notify('\\u2605');", { 復号: false });
    if (!(解 === 1 && 素 === 0)) throw new Error('解く ' + 解 + ' ／ 解かない ' + 素);
  });
  T('★HTML：実体参照を 解かないと &#9733; を 見逃す（解く 1 ／ 解かない 0）★', () => {
    const 解 = HTMLの星('<div>&#9733;</div>'), 素 = HTMLの星('<div>&#9733;</div>', { 復号: false });
    if (!(解 === 1 && 素 === 0)) throw new Error('解く ' + 解 + ' ／ 解かない ' + 素);
  });
  /* ── JS の 字 まるごと（2026-10-09）＝仮の repo を 本物の 全部数える に 通す ── */
  赤('textContent の 右側', JSの赤({ 'book.html': 頁("el.textContent = '★あ';") }, 'JS の 字'));
  赤('innerHTML の 右側が 行を またぐ', JSの赤({ 'book.html': 頁("el.innerHTML = '<b>'\n  + '★あ</b>';") }, 'JS の 字'));
  赤('lib が 返す 訳（読み込まれて いる）', JSの赤({ 'book.html': 頁('', ['lib/p.js']), 'lib/p.js': "function f(){ return { なぜ: '★あ' }; }" }, 'JS の 字'));
  赤('どの html も 読み込まない js でも 白名簿に 無ければ 赤（黙って 除かない）', JSの赤({ 'book.html': 頁(''), 'lib/p.js': "function f(){ return { なぜ: '★あ' }; }" }, 'JS の 字'));
  緑('白名簿の「読み込まれていない」（約束の 台帳）は 読み込まれない 間は 除く', JSの赤({ 'book.html': 頁(''), 'lib/yakusoku-daicho.js': "var a = '★あ';" }, 'JS の 字'));
  赤('白名簿の「読み込まれていない」でも 読み込まれたら 赤', JSの赤({ 'book.html': 頁('', ['lib/yakusoku-daicho.js']), 'lib/yakusoku-daicho.js': "var a = '★あ';" }, 'JS の 字'));
  /* ── 2026-10-09 本番前の 対立役が 作った 穴（W1〜W7）＝どれも 前は 緑で 通った ── */
  赤('W1 実体参照 &#9733; の 字（innerHTML に 入れば ★）', JSの赤({ 'book.html': 頁("el.innerHTML = '&#9733;あ';") }, 'JS の 字'));
  緑('W1 &amp;#9733; は ★ に ならない', JSの赤({ 'book.html': 頁("el.innerHTML = '&amp;#9733;';") }, 'JS の 字'));
  赤('W1b String.fromCharCode(9733)', JSの赤({ 'book.html': 頁('el.textContent = String.fromCharCode(9733);') }, '字の 形で ない'));
  赤('W2 テンプレートの _loadScript(`lib/yakusoku-daicho.js`)', JSの赤({ 'book.html': 頁('_loadScript(`lib/yakusoku-daicho.js`);'), 'lib/yakusoku-daicho.js': "var a = '★あ';" }, 'JS の 字'));
  赤('W2b 変数で 渡す _loadScript(n)', JSの赤({ 'book.html': 頁("var n = 'lib/x.js'; _loadScript(n);") }, '字で 無い 読み込み'));
  緑('W2b _loadScript の 定義（function _loadScript(src)）は 読み込みで ない', JSの赤({ 'book.html': 頁('function _loadScript(src){ return src; }') }, '字で 無い 読み込み'));
  赤('W3 js が js を 読む（import(\'./q.js\')）', JSの赤({ 'book.html': 頁('', ['lib/p.js']), 'lib/p.js': 'imp' + "ort('./yakusoku-daicho.js');" /* ★字を 割る★＝tests/refs-resolve が 本物の import と 読む */, 'lib/yakusoku-daicho.js': "var a = '★あ';" }, 'JS の 字'));
  赤('W5 myconsole.log(\'★\') は console で ない', JSの赤({ 'book.html': 頁("myconsole.log('★あ');") }, 'JS の 字'));
  赤('W6 CSS の content:"★"（<style>）', JSの赤({ 'book.html': '<!DOCTYPE html><style>.a:after{content:"★"}</style><body><p>い</p></body>' }, 'CSS の content'));
  赤('W6 CSS の content:"\\2605"（*.css）', JSの赤({ 'book.html': 頁(''), 'css/a.css': '.a::before { content: "\\2605"; }' }, 'CSS の content'));
  /* ── 2026-10-09 叩き直しの 対立役：止めすぎ 4形（緑の 歯）と 名前の 側から 引く 3形（赤の 歯） ── */
  緑('止めすぎ：justify-content の 行の 注記の ★', JSの赤({ 'book.html': '<!DOCTYPE html><style>.z{justify-content:center /* ★中央 */;}</style><body><p>い</p></body>' }, 'CSS の content'));
  緑('止めすぎ：script の 注記の textContent: ★', JSの赤({ 'book.html': 頁('/* el.textContent: ★ここ */ var a = 1;') }, 'CSS の content'));
  緑('止めすぎ：X.import(d)・my_loadScript(x)・字の 中の import(x)', JSの赤({ 'book.html': 頁("X.import(d); my_loadScript(x); var t = 'imp' + 'ort(x)'; var u = 'a import(x) b';") }, '字で 無い 読み込み'));
  緑('止めすぎ：頭が / の src="/lib/p.js" は 見て いる js', JSの赤({ 'book.html': '<!DOCTYPE html><body><p>い</p><script src="/lib/p.js"></script></body>', 'lib/p.js': "var a = 'あ';" }, '読み込まれるのに 見て いない'));
  赤('F 名前を 足して 読む _loadScript(\'lib/\' + \'yakusoku-daicho.js\')', JSの赤({ 'book.html': 頁("_loadScript('lib/' + 'yakusoku-daicho.js' + v);"), 'lib/yakusoku-daicho.js': "var a = '★あ';" }, '読み込まれていない 筈の'));
  赤('P 空白の 在る _loadScript (\'lib/yakusoku-daicho.js\')', JSの赤({ 'book.html': 頁("_loadScript ('lib/yakusoku-daicho.js');"), 'lib/yakusoku-daicho.js': "var a = '★あ';" }, 'JS の 字'));
  赤('H 引用符の 無い <script src=lib/yakusoku-daicho.js>', JSの赤({ 'book.html': '<!DOCTYPE html><body><p>い</p><script src=lib/yakusoku-daicho.js></script></body>', 'lib/yakusoku-daicho.js': "var a = '★あ';" }, '読み込まれていない 筈の'));
  緑('名前が 注記に 在るだけ なら 赤に しない', JSの赤({ 'book.html': 頁('/* lib/yakusoku-daicho.js は 試験だけが 読む */ var a = 1;'), 'lib/yakusoku-daicho.js': "var a = '★あ';" }, '読み込まれていない 筈の'));
  /* ── 2026-10-09 3回目の 対立役：止めすぎを 直した 時に 開けた 抜け道（4d50d18 で 緑・c6f2d98 で 赤だった） ── */
  赤('A el.src = \'extra/x.js\'（読み込みの 一番 普通の 形）', JSの赤({ 'book.html': 頁("var el = document.createElement('script'); el.src = 'extra/x.js';"), 'extra/x.js': "var a='あ';" }, '読み込まれるのに 見て いない'));
  赤('A window._loadScript(\'extra/x.js\')', JSの赤({ 'book.html': 頁("window._loadScript('extra/x.js');"), 'extra/x.js': "var a='あ';" }, '読み込まれるのに 見て いない'));
  赤('A self.importScripts(\'extra/x.js\')', JSの赤({ 'book.html': 頁('', ['lib/w.js']), 'lib/w.js': "self.importScripts('extra/x.js');", 'extra/x.js': "var a='あ';" }, '読み込まれるのに 見て いない'));
  赤('A window._loadScript(v)', JSの赤({ 'book.html': 頁('window._loadScript(v);') }, '字で 無い 読み込み'));
  赤('A self.importScripts(v)', JSの赤({ 'book.html': 頁('self.importScripts(v);') }, '字で 無い 読み込み'));
  赤('B テンプレートの ${import(v)} は 動く コード', JSの赤({ 'book.html': 頁('var t = `${imp' + 'ort(v)}`;') }, '字で 無い 読み込み'));
  赤('B テンプレートの a${_loadScript(v)}b', JSの赤({ 'book.html': 頁('var t = `a${_loadScript(v)}b`;') }, '字で 無い 読み込み'));
  赤('C <link href="assets/a.css"> が 読む css', JSの赤({ 'book.html': '<!DOCTYPE html><link rel="stylesheet" href="assets/a.css"><body><p>い</p></body>', 'assets/a.css': '.a:after{content:"★"}' }, 'CSS の content'));
  赤('D css の url(//...) の 後ろの content（// は css の 注記で ない）', JSの赤({ 'book.html': 頁(''), 'css/a.css': '.a{background:url(//cdn.example/a.png)} .b:after{content:"★"}' }, 'CSS の content'));
  /* ── 4回目の 対立役：css の <!-- --> は 注記で ない（間の 決まりは 生きる）／テンプレートの ${...} の 中の 字と 注記の } ── */
  赤('M4 *.css の <!-- ... --> の 間の content', JSの赤({ 'book.html': 頁(''), 'css/a.css': '<!-- .a:after{content:"★"} -->' }, 'CSS の content'));
  赤('M3 <style> の 中の <!-- ... --> の 間の content', JSの赤({ 'book.html': '<!DOCTYPE html><style><!-- .a:after{content:"★"} --></style><body><p>い</p></body>' }, 'CSS の content'));
  赤("M6 ${...} の 中の 字の '}' で 深さを 数え違えない", JSの赤({ 'book.html': 頁("var t = `${f('}') + imp" + "ort(v)}`;") }, '字で 無い 読み込み'));
  赤('M6b ${...} の 中の 注記の } で 深さを 数え違えない', JSの赤({ 'book.html': 頁('var t = `${ /* } */ imp' + 'ort(v) }`;') }, '字で 無い 読み込み'));
  /* M8（テンプレートの 中の 字の 'import(x)' を 拾う 止めすぎ）は ★わざと 残す★＝テンプレートは 丸ごと 空けない（道具の 注記） */
  赤('M5 css の 字の 中の /* と */ で 間の ★ を 消さない', JSの赤({ 'book.html': 頁(''), 'css/a.css': '.q:after{content:"/*"} .b:after{content:"★"} .z:after{content:"*/"}' }, 'CSS の content'));
  赤('N2 ${...} の 中の 正規表現の " の 後ろの import(v)', JSの赤({ 'book.html': 頁('var t = `${ s.replace(/"/g, 1) + imp' + 'ort(v) }`;') }, '字で 無い 読み込み'));
  赤("N2b ${...} の 中の 正規表現の ' の 後ろの _loadScript(v)", JSの赤({ 'book.html': 頁("var t = `${ /'/.test(a) ? _loadScript(v) : 0 }`;") }, '字で 無い 読み込み'));
  赤('N3 ${...} の 中の 正規表現の // の 後ろの import(v)', JSの赤({ 'book.html': 頁('var t = `${ u.replace(/\\/\\//g, 1) + imp' + 'ort(v) }`;') }, '字で 無い 読み込み'));
  緑('D css の /* */ の 中の content は 見ない', JSの赤({ 'book.html': 頁(''), 'css/a.css': '/* .b:after{content:"★"} */ .c{color:red}' }, 'CSS の content'));
  赤('W7 manifest.json の 名前',JSの赤({ 'book.html': 頁(''), 'manifest.json': '{"name":"★Exally"}' }, 'manifest の 字'));
  T('★W4 白名簿の 物を 1つ 消して 客の 字を 1つ 足すと 赤（数は 同じでも 指紋で 止める）★', () => {
    const d = fs.mkdtempSync(path.join(os.tmpdir(), 'hoshi-'));
    try {
      fs.mkdirSync(path.join(d, 'lib'));
      fs.writeFileSync(path.join(d, 'book.html'), 頁('', ['lib/formula-soto.js']));
      fs.writeFileSync(path.join(d, 'lib/formula-soto.js'), "var a = '★訳した 文だけ'; document.title = '★客';");
      const 赤ら = 全部数える(d).赤;
      const 白 = 赤ら.find((x) => x.名 === 'lib/formula-soto.js' && x.置き場 === '白名簿の 数');
      if (!白 || !/今 2/.test(白.字)) throw new Error('見つけて いない ' + JSON.stringify(赤ら.map((x) => x.置き場 + ' ' + x.字)));
    } finally { fs.rmSync(d, { recursive: true, force: true }); }
  });
  緑('console の 第1引数', JSの赤({ 'book.html': 頁("console.log('★あ');") }, 'JS の 字'));
  緑('注記の 中の ★', JSの赤({ 'book.html': 頁("/* el.textContent = '★' */ // '★'\nvar a = 'い';") }, 'JS の 字'));
  赤('JS の 逃がし \\u2605', JSの赤({ 'book.html': 頁("el.textContent = '\\u2605あ';") }, 'JS の 字'));
  赤('throw の 字', JSの赤({ 'book.html': 頁("throw new Error('★あ');") }, 'JS の 字'));
  赤('★正規表現の 中の ` の 何十行も 先の ★（book.html:3516 と 同じ 形＝自前の 切り方が 3,600行 飲み込んだ）★',
    JSの赤({ 'book.html': 頁("t = t.replace(/`([^`]+)`/g, '<i>$1</i>');\n" + 'var x = 1;\n'.repeat(40) + "el.textContent = '★あ';") }, 'JS の 字'));
  緑('正規表現の 中の ` で 読み損ねない（読めない 0）', JSの赤({ 'book.html': 頁("t = t.replace(/`([^`]+)`/g, 'a'); var b = /\"/g; var c = /'/g;") }, '読めない 字'));
  赤('閉じない 引用符は 読めない 字＝赤', JSの赤({ 'book.html': 頁("var a = 'あ\nvar b = 1;") }, '読めない 字'));
  赤('行を またぐ テンプレート＝飲み込みの 疑い で 赤', JSの赤({ 'book.html': 頁('var a = `あ\nい`;') }, '行を またぐ テンプレート'));
  赤('on 属性の 中の JS', JSの赤({ 'book.html': '<!DOCTYPE html><body><button onclick="this.textContent=\'★\'">あ</button><p>い</p></body>' }, 'on 属性'));
  赤('<template> の 中身', 仮のrepo({ 'book.html': '<!DOCTYPE html><body><p>い</p><template><div>★あ</div></template></body>' }).filter((x) => x.置き場 === '本文').length);
  赤('読み込まれるのに 見て いない 置き場（parts/）', JSの赤({ 'book.html': 頁('', ['parts/x.js']), 'parts/x.js': "var a='あ';" }, '読み込まれるのに 見て いない'));
  赤('名前が .min. でも 借り物で なければ 見る', JSの赤({ 'book.html': 頁('', ['lib/a.min.js']), 'lib/a.min.js': "el.textContent='★';" }, 'JS の 字'));
  T('★白名簿の 物の 数が 変わると 赤（symbols.js に ★ を もう1つ 足す）★', () => {
    const 赤ら = (() => {
      const d = fs.mkdtempSync(path.join(os.tmpdir(), 'hoshi-'));
      try {
        fs.mkdirSync(path.join(d, 'lib'));
        fs.writeFileSync(path.join(d, 'book.html'), 頁('', ['lib/symbols.js']));
        fs.writeFileSync(path.join(d, 'lib/symbols.js'), "var a = ['★', '★'];");
        return 全部数える(d).赤;
      } finally { fs.rmSync(d, { recursive: true, force: true }); }
    })();
    const 白 = 赤ら.find((x) => x.名 === 'lib/symbols.js' && x.置き場 === '白名簿の 数');
    if (!白 || !/今 2/.test(白.字)) throw new Error('見つけて いない ' + JSON.stringify(赤ら.map((x) => x.置き場)));
  });
  T('★白名簿は 1件ずつ 名・種・数・訳 を 持つ（訳の 無い 除外を 作らない）★', () => {
    for (const w of 白名簿) if (!(w.名 && w.種 && w.数 > 0 && w.訳 && w.訳.length >= 10)) throw new Error(JSON.stringify(w));
  });
  T('★★上限を 持っていない（0本しか 通らない）★★', () => {
    const 素 = 注記を外す(fs.readFileSync(fileURLToPath(import.meta.url), 'utf8'));
    if (/赤\.length\s*[<>]=?\s*[1-9]/.test(素)) throw new Error('★上限を 作っています★');
  });
  T('★守る 範囲を 書いてある（まだ 見て いない 置き場と 数を 名乗る）', () => {
    const 本文 = fs.readFileSync(path.join(ROOT, 'scripts/hoshi-kazoeru.mjs'), 'utf8');
    for (const 要る of ['まだ 見て いない 物', 'JS が 画面へ 直に 書く 字', 'lib が 返して 画面が 出す 字', '隠れて いる 所', 'AI に 渡す 字']) {
      if (本文.indexOf(要る) < 0) throw new Error('書いていない: ' + 要る);
    }
  });
  console.log('\n' + pass + ' passed, ' + fail + ' failed');
  process.exit(fail ? 1 : 0);
}

const { 赤, 数 } = 全部数える(ROOT);

if (process.argv.includes('--list')) {
  console.log('\n[hoshi-nashi --list] ★ が 混じっている 字');
  for (const x of 赤) console.log('  ' + x.名 + ' [' + x.置き場 + '] ' + x.字.replace(/\s+/g, ' ').trim().slice(0, 100));
  console.log('\n合計 ' + 赤.length + '本');
  process.exit(0);
}

console.log('\n[hoshi-nashi] お客さんの 画面に 出る 字に ★ を 使っていないか');
T('★見る ファイルを 名前で 選んで いない（repo 直下の HTML 全部・js／lib の 下 全部）', () => {
  const 名簿 = 見るファイル(ROOT);
  const 直下 = fs.readdirSync(ROOT).filter((f) => /\.html$/i.test(f));
  if (名簿.html.length !== 直下.length || !直下.length) throw new Error('HTML ' + 名簿.html.length + ' ／ 直下 ' + 直下.length);
});
T('★検査が 空振りしていない（口を 通る 字 と HTML の 本文を 実際に 読んでいる）', () => {
  if (数.口の字 < 100) throw new Error('口を 通る 字が 少なすぎます: ' + 数.口の字);
  for (const [f, n] of Object.entries(数.本文)) if (!(n > 0)) throw new Error(f + ' の 本文が 読めて いない（' + n + '節）');
  /* ★JS の 字の塊を 実際に 読んで いる★（2026-10-09 の 数＝塊 21,859・html が 読み込む js 122） */
  if (!(数.字の塊 > 10000)) throw new Error('JS の 字の塊が 少なすぎます: ' + 数.字の塊);
  if (!(数.読み込まれる > 50)) throw new Error('html が 読み込む js が 少なすぎます: ' + 数.読み込まれる);
});
T('★白名簿の 物が 今も その 数だけ 在る（無く なった 除外を 残さない）★', () => {
  const 違う = 数.白の数.filter((w) => w.今 !== w.数);
  if (違う.length) throw new Error(違う.map((w) => w.名 + ' ' + w.種 + ' 名指し ' + w.数 + ' ／ 今 ' + w.今).join('／'));
});
T('★★お客さんの 画面に 出る 字に ★ が 0本★★', () => {
  if (赤.length) {
    const 束 = {};
    for (const x of 赤) 束[x.名 + ' ' + x.置き場.split(' ')[0]] = (束[x.名 + ' ' + x.置き場.split(' ')[0]] || 0) + 1;
    throw new Error('★' + 赤.length + '本 出ています★\n   ' + Object.keys(束).map((k) => k + ' ... ' + 束[k] + '本').join('\n   ')
      + '\n   → 一覧は  node tests/hoshi-nashi.test.mjs --list');
  }
});
console.log('\n── 実測 ──');
console.log('  見た ファイル ' + 数.ファイル + '本 ／ 口を 通る 字 ' + 数.口の字 + '本 ／ HTML の 本文 ' + JSON.stringify(数.本文) + ' ／ 属性 ' + 数.属性);
console.log('  JS の 字の塊 ' + 数.字の塊 + '個 ／ ★入り ' + 数.JSの字の星 + '個 ' + JSON.stringify(数.JSの種) + ' ／ html が 読み込む js ' + 数.読み込まれる + '本');
console.log('  ★ が 混じっている ... ' + 赤.length + '本');
console.log('\n' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
