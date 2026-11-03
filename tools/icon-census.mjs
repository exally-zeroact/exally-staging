/* icon-census.mjs — ★リボンの 印を 全部 並べて 種類ごとに 数える★ 2026-08-30
 *
 *  ★なぜ★（監査役の 指示）
 *    「表示しない」の 印が ★🙈（猿の顔）★＝業務ソフトに 顔・動物は ふざけて 見える。
 *    ★でも 1つだけ 直すと ちぐはぐに なる★
 *    ⇒ ★先に 288個 全部を 並べて 数え、揃え方を 決めてから 一度に 直す★。
 *
 *  ★数え方（字の 番号で 分ける・見た目で 決めつけない）★
 *    ・顔/人/動物/食べ物 … U+1F300〜U+1F9FF の うち その 帯に 入る 物
 *    ・色つきの 絵文字   … Emoji_Presentation の 帯（U+1F000〜U+1FAFF）
 *    ・記号（線画）      … U+2000〜U+2BFF（矢印・幾何・その他の 記号）
 *    ・日本語/英数       … かな・漢字・ラテン
 *
 *  走らせ方: node tools/icon-census.mjs
 */
import fs from 'node:fs';
import path from 'node:path';

const ROOT = path.join(path.dirname(new URL(import.meta.url).pathname.replace(/^\/([A-Za-z]:)/, '$1')), '..');
const spec = fs.readFileSync(path.join(ROOT, 'lib/ribbon-spec.js'), 'utf8');

/* 部品ごとの 印を 取り出す
   ★行ごとに 読む★＝1つの 大きい 正規表現は
   ★入れ子の 括弧で 切れて 取りこぼす★
   （08-30 実測＝288個あるのに 286個しか 拾えなかった）。
   ★拾えた数が 合わなければ その場で 赤にする★。 */
const 部品 = [];
for (const 行 of spec.split(String.fromCharCode(10))) {
  const 頭 = /\{ t: '([^']*)', g: '([^']*)', p: '([^']*)', a: (.*)\},?\s*$/.exec(行.trim());
  if (!頭) continue;
  const a = 頭[4].trim() === 'null' ? null : 頭[4];
  const icon = a && /icon: '([^']*)'/.exec(a);
  const act = a && /act: '([^']*)'/.exec(a);
  const 借り = a && /取り込む: '([^']*)'/.exec(a);
  部品.push({ t: 頭[1], g: 頭[2], p: 頭[3], act: act ? act[1] : null,
    icon: icon ? icon[1] : null, 借り: 借り ? 借り[1] : null });
}
/* ★空振りを その場で 止める★＝行の数と 拾った数が 合うか */
const 行の数 = (spec.match(/\{ t: '/g) || []).length;
if (部品.length !== 行の数) {
  console.log('★数え方が 取りこぼしています★ 行=' + 行の数 + ' 拾えた=' + 部品.length);
  process.exit(1);
}

/** その 字の 種類（★番号で 分ける★） */
function 種類(c) {
  const n = c.codePointAt(0);
  if (n >= 0x1F600 && n <= 0x1F64F) return '顔';
  if (n >= 0x1F900 && n <= 0x1F9FF) return '顔/人';
  if (n >= 0x1F400 && n <= 0x1F4FF) return '物/動物';   /* 動物・道具が 混ざる 帯 */
  if (n >= 0x1F300 && n <= 0x1F3FF) return '自然/物';
  if (n >= 0x1F500 && n <= 0x1F5FF) return '記号/道具';
  if (n >= 0x1F000 && n <= 0x1FAFF) return 'その他の 絵文字';
  if (n >= 0x2600 && n <= 0x27BF) return '記号（絵文字に なりやすい）';
  if (n >= 0x2000 && n <= 0x2BFF) return '記号（線画）';
  if (n >= 0x3040 && n <= 0x30FF) return 'かな';
  if (n >= 0x4E00 && n <= 0x9FFF) return '漢字';
  if (n >= 0xFF00 && n <= 0xFFEF) return '全角の 英数記号';
  if (n < 0x180) return '英数';
  return 'その他（U+' + n.toString(16).toUpperCase() + '）';
}

/* ★顔・人・動物・食べ物★＝業務ソフトに ふさわしくない と 言われた 帯 */
function ふざけて見えるか(c) {
  const n = c.codePointAt(0);
  return (n >= 0x1F600 && n <= 0x1F64F)       /* 顔 */
      || (n >= 0x1F900 && n <= 0x1F9FF)       /* 人・体 */
      || (n >= 0x1F400 && n <= 0x1F43F)       /* 動物 */
      || (n >= 0x1F32D && n <= 0x1F37F)       /* 食べ物・飲み物 */
      || (n >= 0x1F950 && n <= 0x1F96F);      /* 食べ物 */
}

const 結んだ = 部品.filter((v) => v.icon);
const 数 = {};
const 危ない = [];
const 印ごと = {};

for (const v of 結んだ) {
  for (const c of [...v.icon]) {
    const k = 種類(c);
    数[k] = (数[k] || 0) + 1;
    if (ふざけて見えるか(c)) 危ない.push(v.t + '|' + v.g + '|' + v.p + ' … ' + c
      + '（U+' + c.codePointAt(0).toString(16).toUpperCase() + '）');
  }
  印ごと[v.icon] = (印ごと[v.icon] || 0) + 1;
}

const 借りた = 部品.filter((v) => v.借り);
console.log('★リボンの 部品 = ' + 部品.length + '個★');
console.log('  結んだ = ' + 部品.filter((v) => v.act).length + '個'
  + '（印を 持つ ' + 結んだ.length + '個 ＋ ★画面の 部品を 借りる ' + 借りた.length + '個★）');
借りた.forEach((v) => console.log('    借り … ' + v.t + '|' + v.g + '|' + v.p + ' → ' + v.借り));
console.log('★印の 種類 = ' + Object.keys(印ごと).length + '種★\n');

console.log('★字の 種類ごとの 数★');
Object.keys(数).sort((a, b) => 数[b] - 数[a])
  .forEach((k) => console.log('  ' + String(数[k]).padStart(4) + '  ' + k));

console.log('\n★顔・人・動物・食べ物 = ' + 危ない.length + '個★');
危ない.forEach((v) => console.log('  ★NG★ ' + v));

/* 同じ 印を いくつの 部品で 使い回しているか（多いと 見分けが つかない） */
const 使い回し = Object.keys(印ごと).filter((k) => 印ごと[k] >= 4)
  .sort((a, b) => 印ごと[b] - 印ごと[a]);
console.log('\n★同じ 印を 4個以上で 使い回している = ' + 使い回し.length + '種★');
使い回し.forEach((k) => console.log('  ' + String(印ごと[k]).padStart(3) + '個  ' + k));

/* 印が 2文字以上（組み合わせ）の 物 */
const 合わせ = Object.keys(印ごと).filter((k) => [...k].length >= 2);
console.log('\n★2文字 以上の 印 = ' + 合わせ.length + '種★');
合わせ.forEach((k) => console.log('  ' + k));

process.exit(危ない.length ? 1 : 0);
