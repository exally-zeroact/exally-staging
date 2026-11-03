/* kami-ga-nai-kansuu.mjs — ★実Excel の 紙が 無い 関数の 名前を 出す★（2026-09-16）
 *
 *  ★★なぜ 要るか★★
 *    ★台が 知らない 関数★を 書く 時、★紙（golden-*.tsv）が 在れば 字だけで 書けます★。
 *    ★紙が 無い 物は 実Excel を 叩かないと 書けません★。
 *    ⇒★実Excel を 開くのは 重い★（★Quit まで 120〜300秒★・機械の 順番を 待つ）
 *    ⇒★★1回の 枠で まとめて 聞く 為に、先に 名前を 出します★★
 *
 *  ★★この 道具は 名簿の 引き算だけ★★
 *    ・実Excel を 叩きません
 *    ・借り物を 立てません（★名前を 見るだけ★）
 *    ⇒★軽い★
 *
 *  ★数え方★
 *    ①★実Excel に 在る 名前★ … `docs/measured/excel-functions-*.txt`
 *    ②★台が 出せる 名前★ … `lib/shiki-kansuu.js` ＋ 皮（`lib/formula-*.js`）
 *    ③★紙に 在る 名前★ … `docs` の tsv の 中で 実際に 押されて いる 関数名
 *    ⇒★★①にあって ②に無く ③にも無い★★ ＝ ★実Excel を 叩かないと 書けない 物★
 *
 *  使い方: node docs/measured/kami-ga-nai-kansuu.mjs
 */
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '../..');
const require_ = createRequire(path.join(ROOT, 'package.json'));

/* ══ ①実Excel に 在る 名前 ══ */
const 実道 = fs.readdirSync(path.join(ROOT, 'docs/measured'))
  .filter((f) => /^excel-functions-.*\.txt$/.test(f)).sort();
if (!実道.length) { console.error('★excel-functions-*.txt が 在りません★'); process.exit(2); }
const 実Excel = new Set(
  fs.readFileSync(path.join(ROOT, 'docs/measured', 実道[実道.length - 1]), 'utf-8')
    .split(/\r?\n/).map((s) => s.trim()).filter((s) => s && !s.startsWith('#'))
    .map((s) => s.split(/[\s\t]/)[0].toUpperCase())
);

/* ══ ②台が 出せる 名前 ══
     ★★自分で Object.keys を 探さない★★
       2026-09-16 に 私は `K.関数たち || K.関数 || K` と 書いて
       ★「台が 出せる 39個」という 嘘の 数を 出しました★（本当は 394個）。
     ⇒★★既に 在る 道具と 同じ 読み方を する★★
       ＝`docs/measured/nokori-wo-kansuu-goto-ni-waru.mjs` と 同じ
         土台 … `K.表`／皮 … `T.名前たち()`
     ★記憶★「★読み込む≠登録される★＝『◆個 できる』を module を 読み込んで 数えるな」 */
const K = require_(path.join(ROOT, 'lib/shiki-kansuu.js'));
const T = require_(path.join(ROOT, 'lib/shiki-tsunagi.js'));
const H = require_(path.join(ROOT, 'lib/shiki-hyou.js'));
/* ★★板に 書いた 物（先に 計算しない 形）も 数える★★（2026-09-18）
     ＝LET／LAMBDA は ★木を 知らないと 書けない★ので `lib/shiki-hyou.js` に 在ります
     ＝★ここが 見て いないと「書いたのに 数が 減らない」に なります★
     ＝★2026-09-18 に 実際に なりました★（LET を 書いても 35 の まま）
     ⇒★`tests/tsunagu-mon.test.mjs` と ★同じ 読み方★に 揃えて あります★
       （★2つの 道具の 数が 合うか★を あちらが 見て います） */
const 特別名 = H.特別な形 || [];
const 土台名 = Object.keys(K.表 || {});
const 皮名 = T.名前たち();
const 台 = new Set([...土台名, ...皮名, ...特別名].map((x) => String(x).toUpperCase()));
if (台.size < 300) {
  console.error('★★台が ' + 台.size + '個しか 読めません★★（★読み方が 壊れて います★）');
  console.error('  ★読めない 数で 引き算を すると ★嘘の 一覧★が 出ます★');
  process.exit(3);
}
/* ══ ③紙に 在る 名前 ══ */
const 紙 = new Set();
const 歩く = (d) => {
  for (const e of fs.readdirSync(d, { withFileTypes: true })) {
    const p = path.join(d, e.name);
    if (e.isDirectory()) { 歩く(p); continue; }
    if (!/\.(tsv|csv)$/.test(e.name)) continue;
    if (/-awase-|^cases-/.test(e.name)) continue;   /* ★突き合わせの 出しは 紙では ない★ */
    const s = fs.readFileSync(p, 'utf-8');
    for (const m of s.matchAll(/(?:^|[=\s,(+\-*/&<>^])([A-Z][A-Z0-9._]{1,30})\s*\(/g)) 紙.add(m[1]);
  }
};
歩く(path.join(ROOT, 'docs'));

/* ══ 引き算 ══ */
const 台にない = [...実Excel].filter((n) => !台.has(n)).sort();
const 紙もない = 台にない.filter((n) => !紙.has(n));
const 紙はある = 台にない.filter((n) => 紙.has(n));

console.log('');
console.log('★★実Excel の 紙が 無い 関数★★（★名簿の 引き算だけ＝実Excel を 叩いて いません★）');
console.log('');
console.log('  ★実Excel に 在る★ ……………… ' + 実Excel.size + '個（' + 実道[実道.length - 1] + '）');
console.log('  ★台が 出せる★ ………………… ' + 台.size + '個'
  + '（土台 ' + 土台名.length + '個 ＋ 皮 ' + 皮名.length + '個 ＋ 板 ' + 特別名.length + '個）');
console.log('  ★紙に 出て くる★ ……………… ' + 紙.size + '個');
console.log('');
console.log('  ★★台に 無い★★ ………………… ' + 台にない.length + '個');
console.log('    ・★紙が 在る（字だけで 書ける）★ … ' + 紙はある.length + '個');
console.log('    ・★★紙が 無い（実Excel が 要る）★★ … ' + 紙もない.length + '個');
console.log('');
console.log('★★実Excel を 叩かないと 書けない 物（★この 名前を 1回の 枠で 聞きます★）★★');
for (let i = 0; i < 紙もない.length; i += 6) {
  console.log('  ' + 紙もない.slice(i, i + 6).map((s) => s.padEnd(22)).join(''));
}
console.log('');
/* ★★何を 何で 数えたか★★（2026-09-16 に 足しました）
     ★なぜ★ 私は この 道具の「56個」と
       `nokori-wo-kansuu-goto-ni-waru.mjs` の「56個（252通り）」を ★混ぜました★。
     ★★両方 56 ですが ★別の 物を 数えて います★★
       ・この 道具 ……… ★実Excel 519 の 名簿★で 台に 無い うち ★紙が 在る 物★
       ・向こうの 道具 … ★紙の 式★に 出て くる 関数の うち 台が 知らない 物
     ⇒★★司さんに 出る のは ★519 の 名簿で 数えた 数★★★
       ＝★お客さんが 打てる 関数の 数★
     ⇒★通り（式の 形）は ★中の 進み具合★★＝★混ぜません★ */
console.log('');
console.log('★★何を 何で 数えたか★★');
console.log('  この 道具の 分母 … ★実Excel 519個の 名簿★'
  + '（docs/measured/excel-functions-*.txt）');
console.log('  ⇒★★司さんに 出す 「あと 何個」は この 数★★ … ★' + 台にない.length + '個★');
console.log('');
console.log('  ★混ぜて は いけない 数★');
console.log('    `nokori-wo-kansuu-goto-ni-waru.mjs` の 「◆通り」 … ★紙の 式を 数えた 数★');
console.log('      ＝★中の 進み具合★（★お客さんが 打てる 関数の 数では ありません★）');
console.log('    ★2026-09-16 に 私は この 2つを 混ぜました★（どちらも 56 だった）');
console.log('');
console.log('★見て いない 事★');
console.log('  ・★「台が 出せる」は ★名前の 一覧★です★（★正しく 答えるかは 別★）');
console.log('  ・★「紙に 出て くる」は 式の 中の 名前です★（★答えが 控えて 在るかは 別★）');
console.log('    ⇒★その 関数の 紙が 「押せる 形か」までは 見て いません★');
