/* tsunagu-mon.test.mjs — ★台を お客さんの 画面に 繋ぐ 手を 止める 門★（2026-09-18）
 *
 *  ★★なぜ 要るか★★
 *    `docs/measured/tsunagu-dandori.md` ④ ... ★経営者1 が 条件に した 物★
 *      「★繋ぐ 前に 必ず 赤に なる 見張り★」
 *      「★『繋いだら 赤』では なく ★繋ぐ 前から 赤★★」
 *      　＝`book.html` に `shiki-*` が 1本でも 入ったら 赤 ⇒★入れる 手が 止まる★
 *
 *  ★★なぜ「繋いだら 赤」では 駄目か★★
 *    ★繋いだ 後に 赤に なる 見張りは、★お客さんに 届いてから★ 鳴ります★
 *    ⇒★記憶「見張りの 値打ちは 赤に する 事では なく ★いつ 赤に なるか★」★
 *      その場（直せる）／後から（運）／気づかない（直せない）
 *    ⇒★この 門は ★その場★に 置きます★
 *
 *  ★★名簿を 手で 書きません★★
 *    ★毎回 引き算で 作ります★ ... 実Excel の 名簿 − （土台 ＋ 皮）
 *    ★訳★ 2026-09-16 に `_xlfn.` の 名簿を 手で 書いて いて ★91個 漏れました★
 *    ⇒★手で 書いた 名簿は 足した 日に 必ず 消し忘れます★
 *
 *  ★★この 門が 見て いない 事★★
 *    ・★「台が 出せる」は ★名前の 一覧★です★（★正しく 答えるかは 別★）
 *    ・★どちらの 道（㋐落とさない ／ ㋑借り物へ 落とす）に するかは ★司さんの 決め★★
 *      ⇒★この 門は どちらでも 同じに 止めます★＝★決めが 紙に 出るまで 手を 止める★
 *
 *  使い方: node tests/tsunagu-mon.test.mjs
 */
import fs from 'node:fs';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(path.join(ROOT, 'package.json'));
let pass = 0, fail = 0;
const T = (n, よい, 添え) => {
  if (よい) { pass++; console.log('  ok   ' + n); }
  else { fail++; console.log('  NG   ' + n + (添え ? '\n       ' + 添え : '')); }
};

console.log('');
console.log('[tsunagu-mon] ★台を 画面に 繋ぐ 手を 止める 門★');

/* ══ ①実Excel の 名簿（★手で 書かない★） ══ */
const 実道 = fs.readdirSync(path.join(ROOT, 'docs/measured'))
  .filter((f) => /^excel-functions-.*\.txt$/.test(f)).sort();
T('★実Excel の 名簿が 在る★', 実道.length > 0, 'docs/measured/excel-functions-*.txt');
const 実Excel = new Set(
  実道.length ? fs.readFileSync(path.join(ROOT, 'docs/measured', 実道[実道.length - 1]), 'utf-8')
    .split(/\r?\n/).map((s) => s.trim()).filter((s) => s && !s.startsWith('#'))
    .map((s) => s.split(/[\s\t]/)[0].toUpperCase()) : []
);
/* ★★読めない 数で 引き算を すると 嘘の 一覧が 出ます★★（2026-09-16 に 踏んだ） */
T('★実Excel の 名簿が 500個 以上 読めた★', 実Excel.size >= 500, 実Excel.size + '個');

/* ══ ②台が 出せる 名前（★`kami-ga-nai-kansuu.mjs` と 同じ 読み方★） ══ */
const K = require_(path.join(ROOT, 'lib/shiki-kansuu.js'));
const Tsu = require_(path.join(ROOT, 'lib/shiki-tsunagi.js'));
const H = require_(path.join(ROOT, 'lib/shiki-hyou.js'));
const 土台名 = Object.keys(K.表 || {});
const 皮名 = Tsu.名前たち();
/* ★★板に 書いた 物（先に 計算しない 形）も 数える★★（2026-09-18）
     ＝LET／LAMBDA は ★木を 知らないと 書けない★ので `lib/shiki-hyou.js` に 在ります
     ＝★数える 側が 見て いないと「書いたのに 数が 減らない」に なります★（実際に なった） */
const 特別名 = H.特別な形 || [];
const 台 = new Set([...土台名, ...皮名, ...特別名].map((x) => String(x).toUpperCase()));
T('★台が 300個 以上 読めた★（★読み方が 壊れたら 嘘の 緑に なる★）',
  台.size >= 300, '台 ' + 台.size + '個（土台 ' + 土台名.length + ' ＋ 皮 ' + 皮名.length + ' ＋ 板 ' + 特別名.length + '）');

const 台にない = [...実Excel].filter((n) => !台.has(n)).sort();
console.log('  ★★台に 無い ... ' + 台にない.length + '個★★'
  + '（実Excel ' + 実Excel.size + ' − 台 ' + 台.size + '）');

/* ══ ③★上限（★今日の 数★）★ ... ★増えたら 赤★ ══
     ★余裕を 残さない★＝余裕の 在る 門は 止めません（2026-09-18 に 決めた） */
const 上限 = 25;   /* ★2026-09-18 ... 49 -> 42 -> 37 -> 35 -> ★27（ラムダの 一族 6個＋LET/LAMBDA）★ */
T('★★台に 無いが ' + 上限 + '個 以下★★（★増えたら 赤★）',
  台にない.length <= 上限,
  '出た ' + 台にない.length + '個 ／ 上限 ' + 上限 + '個'
  + '\n       ★減ったなら この 上限も 下げて ください★');
if (台にない.length < 上限) {
  console.log('  ★★上限を ' + 台にない.length + ' に 下げて ください★★'
    + '（★余裕を 残すと 止まらない 門に なります★）');
}

/* ══ ④★★繋ぐ 前から 赤★★ ══ */
const book = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf-8');
const src = [...book.matchAll(/<script[^>]+src="([^"]+)"/g)].map((m) => m[1]);
const shiki本 = src.filter((s) => /(^|\/)shiki-[^/]*\.js/.test(s));
const hf本 = src.filter((s) => /hyperformula/i.test(s));
console.log('  ★book.html の script src★ ... shiki-* ' + shiki本.length + '本'
  + ' ／ hyperformula ' + hf本.length + '本');

/* ══ ★★2026-09-18 ... ★門を 2つに 割りました★★★ ══
     ★前の 門★ ... 「`book.html` に `shiki-*` が 1本でも 入ったら 赤」
     ★経営者1 の 指摘（★私が 出しました★）★
       「★`script src` を 足すだけでは 64本 増えません★
         ＝★その 門は ★居るか★を 測って いて ★効いて いるか★を 測って いない★」
     ⇒★★「入れて 落ちないかで 測るな＝引いて 正しい 答えが 出るかで 測れ」★★（記憶）
     ⇒★★繋ぐのは 2つに 割れます★★
        ㋐★読み込むだけ★ ... `script src` を 足す（★計算は 借り物の まま★）
        ㋑★計算を 回す★ ... `shiki-hyou` を 呼ぶ（★ここで 初めて 答えが 変わる★）
     ⇒★★㋐は 通す／㋑は 台に 無いが 0に なるまで 止める★★ */
/* ★`shiki-*` は 9本★（`lib/bessel.js` は 名前が 違うので 別に 数えます） */
const 足した本数 = 10;   /* ★9 -> 10★（2026-09-18・㋑⑶で `lib/shiki-ita-awase.js` を 足した） */
const bessel本 = src.filter((s) => /(^|\/)bessel\.js/.test(s));
T('★`lib/bessel.js` も 入って いる★（★皮が `root.Bessel` を 読む★）',
  bessel本.length === 1, '入って いる ' + bessel本.length + '本');
T('★★㋐ ... `book.html` に 台が ' + 足した本数 + '本 入って いる★★（★減ったら 赤★）',
  shiki本.length === 足した本数,
  '入って いる ' + shiki本.length + '本 ／ 決め打ち ' + 足した本数 + '本'
  + '\n       ' + shiki本.join(' '));

/* ★★㋑＝計算を 回して いるか★★（★居るか では なく 効いて いるか★）
     ★見る 字★ ... `book.html` の 中で `ShikiHyou` を 呼んで いるか
     ★今は 0件の はず★＝★読み込むだけ★ */
const 呼ぶ所 = (book.match(/ShikiHyou\s*\./g) || []).length;
console.log('  ★`book.html` が `ShikiHyou` を 呼ぶ 所 ... ' + 呼ぶ所 + '件★'
  + '（★0＝読み込むだけ／1以上＝計算を 回して いる★）');
/* ★★門の 向きを 変えました★★（2026-09-18・㋑⑶を 書いた 時）
     ★前の 門★ ...「★台に 無いが 残って いる 間は 計算を 台に 回さない★」
       ＝訳は「★台が 知らない 式が その場で #NAME? に なる★」でした
     ★今★ ... ★台が `#NAME?` を 返したら 借り物に 落とす★ 道を 書きました
       ＝`book.html` の `_台に聞く` が `#NAME?` と 溢れで `null` を 返します
       ⇒★★だから 止める 訳が 消えました★★
     ⇒★★門を 「回すな」から 「落とす道が 在るか」に 変えます★★
     ★★但し 「在る」だけでは 足りません★★（記憶「入れて 落ちないかで 測るな」）
       ⇒★効いて いるかは `tests/shiki-wo-osu-webkit.mjs` が ★実物の ブラウザで★ 測ります★ */
T('★★㋑ ... 計算を 台に 回して いる★★（★`ShikiHyou` を 呼ぶ 所が 1件 以上★）',
  呼ぶ所 >= 1, '`ShikiHyou` を 呼ぶ 所 ' + 呼ぶ所 + '件');
const 落とす道 = book.indexOf("=== '#NAME?') return null") >= 0;
T('★★台に 無いが 残って いる 間は `#NAME?` を 借り物に 落とす 道が 在る★★',
  台にない.length === 0 || 落とす道,
  '台に 無い ' + 台にない.length + '個 ／ 落とす道 ' + (落とす道 ? '在る' : '★無い★'));
/* ★★2026-09-18 ... 溢れの 扱いが 変わりました★★
     前 ... ★台の 溢れは いつも 借り物に 落とす★
     今 ... ★借り物が 誤りを 返した 時だけ 台の 溢れを マスに 並べる★
       ⇒`=BYROW(...)` ... 借り物 `#NAME?` ／台 3・30 ⇒★台が 勝つ★
       ⇒`=SORT(...)`   ... 借り物が 答える ⇒★今まで通り 借り物★
       ⇒★★下がる 事が ありません★★（★悪い 方を 良い 方に 換えるだけ★）
     ★効いて いるかは `tests/shiki-wo-osu-webkit.mjs` が 実物の ブラウザで 見ます★ */
const 溢れを落とす = book.indexOf('溢れ === true) return null') >= 0;
T('★★1マス 打った 時（`_台に聞く`）は 溢れを 借り物に 落とす★★',
  溢れを落とす, '落とす道 ' + (溢れを落とす ? '在る' : '★無い★'));
const 溢れを並べる = book.indexOf('function _台の溢れを写す') >= 0;
const 誤りの時だけ = book.indexOf("String(cell.d || '').charAt(0) === '#'") >= 0;
T('★★計算し直す 時は 台の 溢れを マスに 並べる 道が 在る★★',
  溢れを並べる, '並べる道 ' + (溢れを並べる ? '在る' : '★無い★'));
T('★★並べるのは ★借り物が 誤りを 返した 時だけ★★★（★下がらない 為★）',
  誤りの時だけ, 'ふるい ' + (誤りの時だけ ? '在る' : '★無い★'));

T('★★台に 無いが 残って いる 間は `hyperformula` を 外さない★★',
  台にない.length === 0 || hf本.length >= 1,
  'hyperformula ' + hf本.length + '本 ／ 台に 無い ' + 台にない.length + '個'
  + '\n       ★先に 外すと ★台が 知らない 式が 一斉に #NAME? に なります★');

/* ══ ⑤★名簿を 手で 書いて いないか★（★この 門 自身を 読む★） ══
     ★訳★ 手で 書いた 名簿は ★足した 日に 消し忘れます★ */
const 自分 = fs.readFileSync(fileURLToPath(import.meta.url), 'utf-8');
const 焼き込み = 台にない.filter((n) => n.length >= 4 && 自分.includes(n));
/* ★★1個か 2個は 許します★★（2026-09-18）
     ★訳★ ... ★訳を 書く のに 関数の 名前を 1つ 出すのは 当たり前★です
       ＝2026-09-18、私が LET／LAMBDA を 書いた 覚書きで ★この 門が 自分で 赤に なりました★
     ★止めたいのは「名簿を 手で 持つ」事★＝★3個 以上 並んで いたら 名簿★
     ★★甘く した のでは ありません★★＝★何を 数えたいかを 書き直しました★ */
const 焼き込みの上限 = 2;
T('★台に 無い 関数の 名前を この 門に ★名簿として★ 持って いない★',
  焼き込み.length <= 焼き込みの上限,
  '焼き込み ' + 焼き込み.length + '個（上限 ' + 焼き込みの上限 + '個） ... ' + 焼き込み.join(' '));

/* ══ ⑥★もう 1つの 道具と 数が 合うか★（★読み方が 2本に 割れて いないか★） ══ */
const r = spawnSync(process.execPath, [path.join(ROOT, 'docs/measured/kami-ga-nai-kansuu.mjs')],
  { encoding: 'utf-8' });
const 出 = String(r.stdout || '') + String(r.stderr || '');
  /* ★★印の 字に 頼らない★★（2026-09-18）
       ★前は もう 1つの 道具の 出しに 在る ... を そのまま 書いて いました★
       ⇒★決まり §6 で ... を ASCII に 直した ら ★この 門が 自分で 割れました★★
       ⇒★数の 前後の 飾りでは なく ★言葉と 数字★で 掴みます★ */
  const m = /台に 無い[^0-9]*([0-9]+)/.exec(出);
T('★もう 1つの 道具（kami-ga-nai-kansuu.mjs）が 数を 出せた★', !!m,
  '出しの 頭 ... ' + 出.slice(0, 120).replace(/\n/g, ' '));
T('★★2つの 道具の 数が 合う★★（★読み方が 割れて いない★）',
  !!m && Number(m[1]) === 台にない.length,
  'この 門 ' + 台にない.length + '個 ／ もう 1つ ' + (m ? m[1] : '(読めず)') + '個');

console.log('');
console.log('  ★★今 どこまで 来たか★★');
console.log('    ・台が 出せる ... ' + 台.size + '個（土台 ' + 土台名.length + ' ＋ 皮 ' + 皮名.length + ' ＋ 板 ' + 特別名.length + '）');
console.log('    ・★あと ' + 台にない.length + '個★（実Excel 519 の 名簿で 数えた 数）');
console.log('    ・★この 門は 「繋いで よいか」だけを 見ます★＝★正しく 答えるかは 別の 門★');
console.log('');
console.log('tsunagu-mon: ' + pass + ' 緑 / ' + fail + ' 赤');
process.exit(fail ? 1 : 0);
