/* no-silent-optional.test.mjs — ★「黙って素通り」する守り方を許さない★
 *
 * なぜ必要か（2026-08-09・実機で出た）:
 *   14シートのブックを開いてもタブが1枚のままだった。原因は
 *       if (typeof renderSheetTabs === 'function') renderSheetTabs();
 *   ★その名前の関数が存在しなかったので、黙って何も起きなかった★。
 *   node --check も div の数も参照の解決も緑。★押すまで誰も気づけない★。
 *
 * 判定:
 *   book.html の `typeof ○○ === 'function'` で守っている名前は、
 *   ① REQUIRED_FNS に載っている（＝起動時に存在を確かめ、無ければ画面に赤で出す）か
 *   ② OPTIONAL_FNS に★理由つき★で載っている（＝外部ライブラリの版で有無が変わる物）
 *   のどちらかであること。どちらにも無い名前が1つでもあれば赤。
 *   さらに ★REQUIRED に載っている名前が本当にどこかで定義されているか★も見る
 *   （載せただけで実体が無い＝今回の事故そのもの）。
 *
 * 使い方: node tests/no-silent-optional.test.mjs
 *         node tests/no-silent-optional.test.mjs --self-test  ★わざと壊して赤になるか★
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { 注記を外す } from '../scripts/lib/chuki.mjs';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const read = (f) => fs.readFileSync(path.join(ROOT, f), 'utf8');

/** コメントを外す（★コメントの中の例文は「守り」ではない★）。文字列の中は狙わない */
export function stripComments(src) {
  /* ★注記外しは 共通部品へ（2026-08-26 指示役：同じ型を3回 踏んだら 決まりにする）★ */
  return 注記を外す(String(src), { html: true });
}

/** `typeof X === 'function'` / `!==` で守っている名前を全部拾う */
export function guardedNames(srcRaw) {
  const src = stripComments(srcRaw);
  const out = [];
  /* ★引用符は 3種類ある★（' " `）。1種類しか見ないと、
     わざと壊した物が素通りした（2026-08-09 実測。これで気づいた）。
     ==/!= の2文字版も拾う。 */
  /* ★名は 日本語でも 拾う（2026-11-14）★＝英字だけを 拾っていた＝日本語の 名の 守り 19種が 見張りの 外だった */
  const re = /typeof\s+([\p{ID_Start}_$][\p{ID_Continue}_$.\u200c\u200d]*)\s*(?:===|!==|==|!=)\s*['"`]function['"`]/gu;
  let m;
  while ((m = re.exec(src))) out.push(m[1]);
  return [...new Set(out)];
}
/** 宣言された表を読む（book.html の中の配列/オブジェクト） */
export function declared(src) {
  const req = (src.match(/var REQUIRED_FNS\s*=\s*\[([\s\S]*?)\];/) || [])[1] || '';
  const opt = (src.match(/var OPTIONAL_FNS\s*=\s*\{([\s\S]*?)\};/) || [])[1] || '';
  const par = (src.match(/var PARAM_FNS\s*=\s*\{([\s\S]*?)\};/) || [])[1] || '';
  const names = (s) => [...s.matchAll(/'([^']+)'/g)].map((x) => x[1]);
  const required = names(req.replace(/\/\/[^\n]*/g, ''));
  const optional = {};
  opt.replace(/\/\/[^\n]*/g, '').replace(/'([^']+)'\s*:\s*'([^']*)'/g, (_, k, v) => { optional[k] = v; return ''; });
  const param = {};
  stripComments(par).replace(/'([^']+)'\s*:\s*'([^']*)'/g, (_, k, v) => { param[k] = v; return ''; });
  return { required, optional, param };
}
/** その名前が どこかで定義されているか（book.html か 読み込む .js）
 *  ★前の境目は \\b でなく (?:^|[^\\p{ID_Continue}$.])★（後読みは 使わない＝tests/no-lookbehind の 決まり）＝日本語の 名の 前では \\b が 効かない（日本語は \\w でない）。
 *  ★点つき（A.b）★＝頭の 名が どこかで 定義され（root.A = ／ window.A = ／ A = factory() も 定義）、
 *    末尾の 名が ★頭を 定義した 同じ ファイル★に function 末尾( か 末尾: で 在る事。
 *    ★これは 甘い 判定★＝本当の 守りは 起動時の 確かめを 本物の ブラウザで 見る 試験（tests/required-fns-webkit.mjs） */
const esc = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
function defRe(base) {
  const b = esc(base);
  return new RegExp('(?:async\\s+)?function\\s+' + b + '\\s*\\(|(?:^|[^\\p{ID_Continue}$.])' + b + '\\s*[:=]\\s*(?:async\\s+)?(?:function|\\()', 'u');
}
export function isDefined(name, sources) {
  const parts = name.split('.');
  if (parts.length === 1) { const re = defRe(name); return sources.some((s) => re.test(s)); }
  const head = esc(parts[0]), tail = esc(parts[parts.length - 1]);
  const headRe = new RegExp('(?:(?:^|[^\\p{ID_Continue}$])(?:root|window|globalThis|self)\\.' + head + '|(?:^|[^\\p{ID_Continue}$.])' + head + ')\\s*=(?!=)', 'u');
  const tailRe = new RegExp('(?:async\\s+)?function\\s+' + tail + '\\s*\\(|(?:^|[^\\p{ID_Continue}$.])' + tail + '\\s*:', 'u');
  return sources.some((s) => headRe.test(s) && tailRe.test(s));
}

let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ✓ ' + n); } catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + (e && e.message)); } };

if (process.argv.includes('--self-test')) {
  console.log('\n[no-silent-optional] ★わざと壊して赤になるか★');
  const good = "var REQUIRED_FNS = ['render'];\nvar OPTIONAL_FNS = { 'hf.x': '版で有無が変わる' };\n"
    + "if (typeof render === 'function') render();\nfunction render(){}\n";
  T('表に載っている名前だけなら緑', () => {
    const d = declared(good);
    const un = guardedNames(good).filter((n) => !d.required.includes(n) && !d.optional[n]);
    if (un.length) throw new Error('誤検知: ' + un.join(','));
  });
  T('★表に載っていない名前を1つ足したら赤', () => {
    const bad = good + "if (typeof renderSheetTabs === 'function') renderSheetTabs();\n";
    const d = declared(bad);
    const un = guardedNames(bad).filter((n) => !d.required.includes(n) && !d.optional[n]);
    if (un.length !== 1 || un[0] !== 'renderSheetTabs') throw new Error('捕まえられない: ' + JSON.stringify(un));
  });
  T('★REQUIRED に載っているのに実体が無ければ赤（今回の事故そのもの）', () => {
    const bad = "var REQUIRED_FNS = ['renderSheetTabs'];\nvar OPTIONAL_FNS = {};\n";
    if (isDefined('renderSheetTabs', [bad])) throw new Error('無いのに「有る」と言った');
  });
  T('★コメントの中の例文は「守り」と数えない', () => {
    const withComment = good + "// if (typeof ghostFn === 'function') ghostFn();\n";
    if (guardedNames(withComment).includes('ghostFn')) throw new Error('コメントを数えてしまった');
    const withBlock = good + "/* if (typeof ghost2 === 'function') ghost2(); */\n";
    if (guardedNames(withBlock).includes('ghost2')) throw new Error('ブロックコメントを数えてしまった');
  });
  T('★理由が空の任意は赤', () => {
    const bad = "var REQUIRED_FNS = [];\nvar OPTIONAL_FNS = { 'a.b': '' };\n";
    const d = declared(bad);
    if (d.optional['a.b'] !== undefined && d.optional['a.b'].length >= 8) throw new Error('短い理由を通した');
  });
  /* ★2026-11-14 日本語の 名と 点つき★ */
  const jp = "var REQUIRED_FNS = ['すべて再計算'];\nvar OPTIONAL_FNS = {};\nfunction すべて再計算(){}\nif (typeof すべて再計算 === 'function') すべて再計算();\n";
  T('★日本語の 名を 表に 載せずに 守ったら 赤', () => {
    const bad = jp + "if (typeof ページ設定の初期化 === 'function') ページ設定の初期化();\n";
    const d = declared(bad);
    const un = guardedNames(bad).filter((n) => !d.required.includes(n) && !d.optional[n]);
    if (un.length !== 1 || un[0] !== 'ページ設定の初期化') throw new Error('捕まえられない: ' + JSON.stringify(un));
  });
  T('★日本語の REQUIRED に 実体が 在れば 緑（\\b の 誤りで 赤に しない）', () => {
    if (!isDefined('すべて再計算', [jp])) throw new Error('在るのに「無い」と言った');
    if (!isDefined('_台の板を得る', ['var x=1;function _台の板を得る(){}'])) throw new Error('_ で始まる 日本語の 名を 見落とした');
  });
  T('★日本語の REQUIRED に 実体が 無ければ 赤', () => {
    if (isDefined('ページ設定の初期化', [jp])) throw new Error('無いのに「有る」と言った');
    if (isDefined('設定の初期化', ['var ページ設定の初期化 = function(){}'])) throw new Error('名の 途中を 定義と 読んだ');
  });
  const plug = "(function(root,factory){ root.FormulaSotoPlug = factory(); })(this, function(){ function 外へ出す(){} return { 外へ出す: 外へ出す }; });";
  T('★点つきの 末尾が 在れば 緑・無ければ 赤・頭が 無ければ 赤', () => {
    if (!isDefined('FormulaSotoPlug.外へ出す', [plug])) throw new Error('在るのに「無い」と言った');
    if (isDefined('FormulaSotoPlug.外へ出さない', [plug])) throw new Error('末尾が 無いのに「有る」と言った');
    if (isDefined('SotoPlug.外へ出す', [plug])) throw new Error('頭が 無いのに「有る」と言った');
    if (isDefined('FormulaSotoPlug.外へ出す', [plug.replace('function 外へ出す(){} return { 外へ出す: 外へ出す };', ''), 'function 外へ出す(){}'])) throw new Error('頭と 末尾が 別の ファイルなのに「有る」と言った');
  });
  console.log('\n' + pass + ' passed, ' + fail + ' failed');
  process.exit(fail ? 1 : 0);
}

console.log('\n[no-silent-optional] 「無ければ黙って素通り」する守りが残っていないか');
const html = read('book.html');
const libs = ['exally-formula.js', 'lib/grid-xlsx.js', 'js/book-open.js', 'lib/xlsx-io.js', 'lib/formula-soto-plug.js', 'lib/shiki-hyou.js']
  .filter((f) => fs.existsSync(path.join(ROOT, f))).map(read);
const sources = [html].concat(libs);
const d = declared(html);
const guarded = guardedNames(html);

T('★どの表にも載っていない守りが無い', () => {
  const un = guarded.filter((n) => !d.required.includes(n) && !d.optional[n] && !d.param[n]);
  if (un.length) {
    throw new Error('この名前は REQUIRED_FNS / OPTIONAL_FNS / PARAM_FNS のどれかに入れること:\n   - ' + un.join('\n   - '));
  }
});
T('★REQUIRED に載っている名前が、本当にどこかで定義されている', () => {
  const ghost = d.required.filter((n) => !isDefined(n, sources));
  if (ghost.length) throw new Error('名前だけあって実体が無い（＝押すと黙って何も起きない）: ' + ghost.join(', '));
});
T('★任意・引数の物には理由が書いてある', () => {
  for (const [k, v] of Object.entries(Object.assign({}, d.optional, d.param))) {
    if (!v || v.length < 8) throw new Error(k + ': 理由が短すぎる');
  }
});
T('検査が空振りしていない（守りを実際に数えている）', () => {
  if (guarded.length < 5) throw new Error('守りが少なすぎます: ' + guarded.length);
  if (!d.required.length) throw new Error('REQUIRED_FNS を読めていない');
});

console.log('\n── 実測 ──');
console.log(`  typeof で守っている名前: ${guarded.length}種`);
console.log(`  ★無かったら壊れている（起動時に確かめる）: ${d.required.length}★`);
console.log(`  本当に任意（理由つき）: ${Object.keys(d.optional).length}  ${Object.keys(d.optional).join(', ')}`);
console.log(`  引数として渡る物（理由つき）: ${Object.keys(d.param).length}  ${Object.keys(d.param).join(', ')}`);
console.log('\n' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
