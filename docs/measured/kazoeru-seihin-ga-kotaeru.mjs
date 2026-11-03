/* ★★製品が 実Excel の 関数を 何個 名乗るか を 数える★★（2026-09-15 経営者1）
 *
 * ★訳★ … `docs/measured/karimono-keshi-no-ookisa.md` が ★445／74★ と 書いて いたが 違った。
 *
 * ★★この 道具が 守る 決まりは 4つ★★
 *  ①★数えるのは「書いた 物」では なく「お客さんの 画面に 届いて いる 物」★
 *  ②★★読み込む ＝ 登録される、では ない★★
 *      `lib/formula-*-plug.js` は ★`つなぐ(H, 中身)` を 呼ぶまで 0個★。
 *      `exally-formula.js` ★だけ★ 尻尾で 自分を 登録する（`registerExallyFunctions`）
 *      ⇒★1本だけ 自己登録するので「読み込むだけ」の 道具でも 動いて いるように 見える★
 *      ⇒★★2026-09-15 経営者1 は これで 65個 取りこぼし「424」と 出した（正しくは 489）★★
 *  ③★建て方を 自分で 書かない★＝`docs/measured/honban-no-michi.mjs` の `建てる()` を 呼ぶ
 *      （★`つなぐ` の 口は 8本とも 形が 違う★＝写し取ると すぐ ずれる）
 *  ④★★「名乗る」は「正しく 答える」では ない★★
 *      例＝お金の 21個は `DATE()` で 日付を 渡すと `#VALUE!`（棚㉘）。
 *
 * ★★見て いない 事★★
 *   ・ここは ★node の 台★＝★画面の 数では ない★（`建てる()` の 断りと 同じ）
 *   ・★名前が 在るか★だけを 見る。★答えが 合うかは 見て いない★
 *
 * ★走らせ方★ … repo の 根で `node docs/measured/kazoeru-seihin-ga-kotaeru.mjs`
 */
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { 建てる, ROOT, 本番のプラグイン数 } from './honban-no-michi.mjs';

const 読む = (p) => fs.readFileSync(path.join(ROOT, p), 'utf8');
const require_ = createRequire(path.join(ROOT, 'package.json'));

/* ── ①実Excel の 名簿 ── */
const 実Excel名簿 = 'docs/measured/excel-functions-2026-09-06.txt';
const 実 = new Set(読む(実Excel名簿).split(/\r?\n/).map((s) => s.trim()).filter(Boolean));

/* ── ②お客さんの 画面（book.html）を 見る ── */
const html = 読む('book.html');
const 全src = [...html.matchAll(/<script[^>]+src="([^"?]+)/g)].map((m) => m[1]);
const 皮の本数 = 全src.filter((p) => /\/shiki-/.test(p)).length;
const つなぐ回数 = (html.match(/Plug\.つなぐ\(/g) || []).length;

/* ── ★門は 使う 前に 置く★（09-15 … 空の book.html で ★門より 前に 落ちた★） ── */
for (const [名, 値, 単] of [
  ['実Excel の 名簿', 実.size, '行'],
  ['book.html の script src', 全src.length, '本'],
  ['book.html の つなぐ 呼び出し', つなぐ回数, 'か所'],
]) {
  if (!値) {
    console.error(`★★道具が 壊れて います★★ … ${名} が ★0${単}★（根 … ${ROOT}）`);
    process.exit(1);
  }
}

/* ── ③段1＝★借り物だけ★（`建てる()` を 呼ぶ 前に 取る） ── */
const HFns = require_(path.join(ROOT, 'hyperformula.full.min.js'));
const 段1 = new Set(HFns.HyperFormula.getRegisteredFunctionNames('enGB'));

/* ── ④段3＝★本番と 同じ 並びで つないだ 後★ ── */
const 道 = await 建てる();
const 段3 = new Set(道.HF0.getRegisteredFunctionNames('enGB'));

/* ── ⑤JS層（式が 借り物へ 行く 前に 横取りする 名簿） ── */
const 素 = 読む('exally-formula.js');
const jm = 素.match(/var _jsSet\s*=\s*\{([\s\S]*?)\};/);
const JS層 = new Set(jm ? [...jm[1].matchAll(/([A-Z][A-Z0-9._]*)\s*:\s*1/g)].map((x) => x[1]) : []);
const 段4 = new Set([...段3, ...JS層]);

/* ── ⑥★別の 道でも 数えて 突き合わせる★（片方が 壊れても 気づけるように）
       plug の 名簿（implementedFunctions）を ★字で★ 読む。
       ★正規表現で 閉じ波括弧まで 取らない★＝逆斜線が 落ちやすい（09-15 に 2回 転んだ）
       ⇒★見つけた 所から 波括弧を 数えて 閉じまで 取る★ */
const plug名簿 = [];
for (const f of fs.readdirSync(path.join(ROOT, 'lib')).filter((x) => /^formula-.*-plug\.js$/.test(x))) {
  const 中 = 読む(`lib/${f}`);
  let 名 = [];
  const i = 中.search(/implementedFunctions\s*=\s*\{/);
  if (i >= 0) {
    const 始 = 中.indexOf('{', i);
    let 深 = 0, 終 = -1;
    for (let k = 始; k < 中.length; k++) {
      if (中[k] === '{') 深++;
      else if (中[k] === '}') { 深--; if (深 === 0) { 終 = k; break; } }
    }
    if (終 > 始) 名 = [...中.slice(始, 終).matchAll(/['"]([A-Z][A-Z0-9._]*)['"]\s*:/g)].map((x) => x[1]);
  }
  plug名簿.push({ f, 名 });
}
const 字で読んだplug = new Set();
for (const p of plug名簿) p.名.forEach((n) => 字で読んだplug.add(n));

/* ── ⑦数える（実Excel の 名前だけ） ── */
const 実のみ = (s) => [...s].filter((x) => 実.has(x));
const 借り物のみ = 実のみ(段1);
const 名乗る = 実のみ(段4);
const 横取り済 = 借り物のみ.filter((x) => JS層.has(x));
const 消えた = [...段1].filter((x) => !段3.has(x));

/* ── ⑧JS層の 中で 借り物の 表を 読む 物（★名前は 自前でも 外せない★） ── */
const 頭 = [...素.matchAll(/\bfunction (_js[A-Za-z0-9_]+)\s*\(/g)].map((x) => ({ 名: x[1], i: x.index }));
const 触る = [];
頭.forEach((h, k) => {
  const 中 = 素.slice(h.i, k + 1 < 頭.length ? 頭[k + 1].i : 素.length);
  if (/_hf\b/.test(中)) 触る.push(h.名);
});

/* ── ⑨出す（★何を 何本 読んだかを 数の 隣に 書く★） ── */
const 行 = [];
行.push('★★製品が 実Excel の 関数を 何個 名乗るか★★（★本番の 建て方＝honban-no-michi.mjs の 建てる()★）');
行.push('');
行.push(`  ★読んだ 物★ ${実Excel名簿} … ★${実.size}行★`);
行.push(`               book.html … ★1本★（script src ★${全src.length}本★／つなぐ ★${つなぐ回数}か所★）`);
行.push(`               つないだ plug … ★${道.積んだ}本★（book.html が 読む 数 ★${本番のプラグイン数()}本★）`);
行.push(`               lib の plug の 名簿を 字でも 読んだ … ★${plug名簿.length}本★`);
行.push('');
行.push(`  ★①借り物だけ★ …………………………… ★${借り物のみ.length}★（素の 名前 ${段1.size}個）`);
行.push(`  ★②③＋自己登録＋つなぐ 8本★ … +${実のみ(段3).length - 借り物のみ.length} ⇒ ★${実のみ(段3).length}★`);
for (const p of plug名簿) {
  const 名 = p.f.replace('formula-', '').replace('-plug.js', '');
  行.push(`       ${名.padEnd(10)} 名簿 ${String(p.名.length).padStart(2)}個`);
}
行.push(`  ★④＋JS層 _jsSet★ ……………………… +${名乗る.length - 実のみ(段3).length} ⇒ ★${名乗る.length}★（名簿 ${JS層.size}個）`);
行.push(`  −消えた 名前 …………………………………… ${消えた.length}個`);
行.push('');
行.push(`  ★★名乗る ＝ ${名乗る.length}個 ／ 名乗らない ＝ ${実.size - 名乗る.length}個★★`);
行.push(`  ★★「名乗る」は「正しく 答える」では ない★★（お金の 21個は DATE() で #VALUE!・棚㉘）`);
行.push(`  ★名乗らない ${実.size - 名乗る.length}個★ … ${[...実].filter((x) => !段4.has(x)).sort().join(' ')}`);
行.push('');
/* ★★台が 出せる 数も ここで 出します★★（2026-09-16）
     ★なぜ★ この 道具は ★製品（借り物）だけ★を 数えて いました。
       ⇒★台の 数は その たび 別の 所で 数えられ★、
         2026-09-16 に ★「台が 知る 269個」と いう 間違った 数★が 出ました
         （`K.表` を 取らず ★当てで 一番 大きい 塊を 選んで いた★）。
       ★私も 同じ 穴を 踏みました★（`K.関数たち || K.関数 || K` で 39個）。
     ⇒★★二つの 数を 隣に 並べて 出す★★＝★誰も 数え直さなくて 済む★
     ★読み方は 1か所に 寄せます★
       土台 … `K.表` ／ 皮 … `T.名前たち()`
       （`nokori-wo-kansuu-goto-ni-waru.mjs` と 同じ） */
const 土台表 = require_(path.join(ROOT, 'lib/shiki-kansuu.js'));
const 皮表 = require_(path.join(ROOT, 'lib/shiki-tsunagi.js'));
const 土台名 = Object.keys((土台表 && 土台表.表) || {});
const 皮名 = (皮表 && typeof 皮表.名前たち === 'function') ? 皮表.名前たち() : [];
const 台名 = new Set([...土台名, ...皮名].map((x) => String(x).toUpperCase()));
行.push(`  ★台（lib/shiki-*）が 出せる 名前 ＝ ${台名.size}個★（土台 ${土台名.length}個 ＋ 皮 ${皮名.length}個）`);
行.push(`     ★実Excel ${実.size}個の うち 台に 無い ＝ ${[...実].filter((x) => !台名.has(x)).length}個★`);
行.push(`     ★★この 数を 別の 所で 数え直さないで ください★★（★当てで 選ぶと 外れます★）`);
行.push(`  ★自前の 台（lib/shiki-*）が 効いて いる ＝ ${皮の本数 === 0 ? '0個' : '要 確認'}★`);
行.push(`     訳 … book.html の script src に shiki-* が ★${皮の本数}本★`);
行.push('');
/* ★★この 2つを 並べて 出す★★（2026-09-16・経営者1）
   ★訳★ … 前は 下の 数だけ 出して いて、★「あと 何個 書けば よいか」と 読まれた★。
   ★「あと 何個」は 上の「台に 無い」の 方★＝★経営者1 が 365 と 出して 取り下げた★。
   ⇒★消さずに 並べ、★何の 数か★を 名指しで 書く★（消すと 次の 人が また 同じ 数を 出す） */
行.push(`  ★★あと 何個 書けば 借り物を 外せるか ＝ ${[...実].filter((x) => !台名.has(x)).length}個★★（＝上の「台に 無い」）`);
行.push(`  （参考）★借り物に しか 実装が 無い 名前 ＝ ${借り物のみ.length - 横取り済.length}個★`);
行.push(`     ＝借り物の 実Excel名 ${借り物のみ.length} − JS層が 既に 横取り ${横取り済.length}`);
行.push(`     ★★これは「あと 何個」では ありません★★（JS層は 中で 借り物の 表を 読んで います）`);
行.push(`  ★★ただし この 数は「名前の 数」★★`);
行.push(`     exally-formula.js の _js* 定義 ★${頭.length}本★ の うち`);
行.push(`     ★★${触る.length}本が _hf（借り物の 表）を 読む★★ … ${触る.join(' ')}`);
行.push(`     ⇒★「借り物にしか 実装が 無い 数」≠「借り物を 外せる 数」★`);
console.log(行.join('\n'));

/* ── ⑩★自分で 自分を 見張る★ ── */
let 赤 = 0;
const 弾く = (訳) => { console.error(`★★道具が 壊れて います★★ … ${訳}`); 赤++; };
/* ★★台の 読み方が 壊れて いないか★★（2026-09-16）
     ★読めない 数で 引き算を すると ★嘘の 一覧★が 出ます★
     （2026-09-16 に 実際に 2人が 別々に 踏みました：394→2 08 と 394→39） */
if (!土台表 || !土台表.表) 弾く('`K.表` が 取れません（★土台の 読み方が 変わった★）');
if (台名.size < 300) 弾く(`台が ${台名.size}個しか 読めません（★300個 未満＝読み方が 壊れて います★）`);
if (!皮名.length) 弾く('`T.名前たち()` が 空です（★皮の 読み方が 変わった★）');
if (段1.size === 0) 弾く('借り物の 登録が ★0個★');
if (道.積んだ !== 本番のプラグイン数()) 弾く(`つないだ ${道.積んだ}本 ≠ book.html が 読む ${本番のプラグイン数()}本`);
/* ★2つの 道の 突き合わせ★ … 字で 読んだ 名簿が 本当に 登録されて いるか */
const 未登録 = [...字で読んだplug].filter((x) => !段3.has(x));
if (未登録.length) 弾く(`plug の 名簿に 在るのに 登録されて いない … ${未登録.length}個（${未登録.slice(0, 8).join(' ')}）`);
if (赤) process.exit(1);
console.log('');
console.log(`  ★見張り★ つないだ ${道.積んだ}本＝book.html ${本番のプラグイン数()}本 ／ plug 名簿 ${字で読んだplug.size}個 とも 登録済 ⇒ ★落ち 0★`);
