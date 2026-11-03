/* om-gap.mjs — ★Excel の オブジェクトモデルと うちの 差を 数える★ 2026-08-30
 *
 *  ★司さんの 方針★「細胞分解レベルまで 網羅して 把握してから 持ち込む」
 *  ⇒ ★リボンの ボタン（297）では なく、Excel が 出来る事（17,828）で 測る★
 *
 *  ★数え方（★見た目で 決めつけない★）★
 *    ・Excel 側 … docs/excel-objectmodel-2026-08-30.tsv（EXCEL.EXE の 型ライブラリ）
 *    ・うち 側 … lib/*.js ＋ book.html の 中に ★その 名前が 出てくるか★
 *      ※ ★これは「在るかも」の 目安★＝★出来る の 証明では ない★
 *      　（本当の 印は 押して 効くか。ここでは ★どこを 掘るか の 地図★を 作る）
 *
 *  走らせ方: node tools/om-gap.mjs [型の名前]
 */
import fs from 'node:fs';
import path from 'node:path';

const ROOT = path.join(path.dirname(new URL(import.meta.url).pathname.replace(/^\/([A-Za-z]:)/, '$1')), '..');
const NL = String.fromCharCode(10), TAB = String.fromCharCode(9);

const 行 = fs.readFileSync(path.join(ROOT, 'docs/excel-objectmodel-2026-08-30.tsv'), 'utf8')
  .split(NL).filter((l) => l.trim() && !l.startsWith('#')).map((l) => l.split(TAB));
if (行.length < 20000) { console.log('★読めていません★ ' + 行.length + '行'); process.exit(1); }

/* うちの ソース（画面＋部品）を 1つに */
const libs = fs.readdirSync(path.join(ROOT, 'lib')).filter((f) => f.endsWith('.js'));
const うち = [
  fs.readFileSync(path.join(ROOT, 'book.html'), 'utf8'),
  ...libs.map((f) => fs.readFileSync(path.join(ROOT, 'lib', f), 'utf8')),
].join(NL);
if (うち.length < 100000) { console.log('★うちの ソースが 読めていません★'); process.exit(1); }

/* ★I付きは 同じ物の 影★＝除く／定数の 束も 別に 数える */
const 影か = (t) => /^I[A-Z]/.test(t);
const 本体 = 行.filter((c) => c[0] !== 'TKIND_ENUM' && !影か(c[1]));

const 型ごと = new Map();
for (const c of 本体) {
  if (!型ごと.has(c[1])) 型ごと.set(c[1], new Set());
  型ごと.get(c[1]).add(c[3]);
}

/** その 名前が うちの ソースに 出てくるか（★英字の 切れ目で 見る★）
 *  ★正規表現を 組み立てない★＝名前に . や ( が 混ざると 壊れるので
 *  ★探して 前後の 1文字を 自分で 見る★（2026-08-30 実際に 壊れた） */
const 語の字 = (c) => c !== undefined
  && ((c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z') || (c >= '0' && c <= '9') || c === '_');
function 在るか(名) {
  if (名.length < 3) return false;                 /* 短すぎる 名前は 数えない */
  let i = うち.indexOf(名);
  while (i >= 0) {
    if (!語の字(うち[i - 1]) && !語の字(うち[i + 名.length])) return true;
    i = うち.indexOf(名, i + 1);
  }
  return false;
}

const 的 = process.argv[2];
const 並び = [...型ごと.entries()]
  .filter(([t]) => !的 || t === 的)
  .sort((a, b) => b[1].size - a[1].size);

let 全体 = 0, 当たり = 0;
const 表 = [];
for (const [t, s] of 並び) {
  const 名たち = [...s];
  const 有 = 名たち.filter(在るか);
  全体 += 名たち.length; 当たり += 有.length;
  表.push({ 型: t, 全: 名たち.length, 在: 有.length, 無: 名たち.length - 有.length });
}

console.log('★Excel の 出来る事（I付きの 影と 定数を 除く）= ' + 全体 + '個★');
console.log('★うちの ソースに 名前が 出てくる = ' + 当たり + '個（'
  + (当たり / 全体 * 100).toFixed(1) + '%）★');
console.log('★出てこない = ' + (全体 - 当たり) + '個★');
console.log('  ※ ★これは「在るかも」の 目安★＝押して 効く 証明では ない');
console.log('');
console.log('── 大きい 型 上位25（全 / 名前が 出る / 出ない）──');
表.slice(0, 25).forEach((v) => console.log('  ' + v.型.padEnd(22)
  + String(v.全).padStart(5) + String(v.在).padStart(6) + String(v.無).padStart(7)));

if (的) {
  const s = 型ごと.get(的);
  if (!s) { console.log('★その 型は 無い★'); process.exit(1); }
  const 名たち = [...s].sort();
  console.log('\n★' + 的 + ' の 出てこない 物★');
  名たち.filter((n) => !在るか(n)).forEach((n) => console.log('  ・' + n));
}
