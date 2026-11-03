/* kansuu-kabaa.test.mjs — ★実Excel の 関数を どれだけ 動かせているか★（2026-09-06）
 *
 *  ★司さん★「細胞レベルに 細かく 全てに 適応しろ」「きっちり ファイルに するのと
 *            細胞レベルで 対応できるように しろや」
 *
 *  ★★なぜ 要るか★★
 *    2026-08-29 に「実Excel 507個／動く 394個＝70%」と 報告した。
 *    ⇒ ★その 一覧を scratchpad に 置いて 消した★
 *    ⇒ 09-05 に「507個の 中に YEN は 居るか」と 聞かれて ★答えられなかった★
 *    ⇒ ★数字だけ 残って 中身が 無い★＝もう 使えない
 *    ⇒★測った 物は repo に 置く★（docs/measured/）
 *
 *  ★★本番の 道は 3段★★（2026-09-06 に book.html:2290 を 読んで 分かった）
 *    ①`_jsComputeFormula`（JS層）… `_jsSet` に 在る 名前は ★ここが 答える★
 *    ②`convertFormula` → HyperFormula（プラグイン 2つ）
 *    ③②が `#…` を 返したら `evalFormula` で もう一度
 *    ⇒★①を 飛ばして 測ると DSUM・OFFSET・LINEST 等を「無い」と 誤判定する★
 *      （★実際に やった★＝1回目の 実測は ここが 抜けていた）
 *
 *  ★この 見張りが 見る 物★
 *    ①`docs/measured/` の 一覧が 在って 空でない
 *    ②★「動かない」と 書いた 物が 本当に 動かないか★を ★押して★ 確かめる
 *    ③★「動く」筈の 物が 本当に 動くか★も 押す（逆向き）
 *    ④★カバー率が 下がったら 赤★（後戻りを 止める）
 *
 *  使い方: node tests/kansuu-kabaa.test.mjs
 *          node tests/kansuu-kabaa.test.mjs --self-test
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(import.meta.url);
let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ✓ ' + n); } catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + (e && e.message)); } };

/** ★純関数★＝押した 結果と 台帳から 食い違いを 返す */
export function 食い違い(動かない台帳, 押す, 全部) {
  const 出 = { 動かないと書いたのに動く: [], 動くはずが動かない: [], 動いた数: 0 };
  const 無 = new Set(動かない台帳);
  for (const f of 全部) {
    const 動 = 押す(f);
    if (動) 出.動いた数++;
    if (動 && 無.has(f)) 出.動かないと書いたのに動く.push(f);
    if (!動 && !無.has(f)) 出.動くはずが動かない.push(f);
  }
  return 出;
}

if (process.argv.includes('--self-test')) {
  console.log('\n[kansuu-kabaa --self-test] わざと壊して赤になるか');
  T('★「動かない」と 書いたのに 動いたら 見つける', () => {
    const r = 食い違い(['A'], (f) => f === 'A', ['A', 'B']);
    if (r.動かないと書いたのに動く.length !== 1) throw new Error('見つけていない');
  });
  T('★「動く」筈が 動かなければ 見つける（後戻り）', () => {
    const r = 食い違い(['A'], () => false, ['A', 'B']);
    if (r.動くはずが動かない.join() !== 'B') throw new Error('見つけていない: ' + JSON.stringify(r));
  });
  T('★正しければ 何も 出さない', () => {
    const r = 食い違い(['A'], (f) => f !== 'A', ['A', 'B']);
    if (r.動かないと書いたのに動く.length || r.動くはずが動かない.length) throw new Error('誤検知');
    if (r.動いた数 !== 1) throw new Error('数え違い: ' + r.動いた数);
  });
  console.log('\n' + pass + ' passed, ' + fail + ' failed');
  process.exit(fail ? 1 : 0);
}

/* ══ 本番 ═══════════════════════════════════════════════════ */
const 部屋 = path.join(ROOT, 'docs/measured');
const 読む = (f) => fs.readFileSync(path.join(部屋, f), 'utf8').split(/\r?\n/).filter(Boolean);

console.log('\n[kansuu-kabaa] 実Excel の 関数を どれだけ 動かせているか');

T('★測った 物が repo に 在る（scratchpad に 置かない）', () => {
  for (const f of ['excel-functions-2026-09-06.txt', 'exally-missing-2026-09-07.txt', 'README.md']) {
    const p = path.join(部屋, f);
    if (!fs.existsSync(p)) throw new Error('無い: ' + f);
    if (!fs.readFileSync(p, 'utf8').trim()) throw new Error('空: ' + f);
  }
});

const 全部 = 読む('excel-functions-2026-09-06.txt');
/* ★09-07 に 22個 動くように なったので 台帳を 差し替えた★
   （★古い 09-06 の 紙は 消さない★＝いつ 何が 動くように なったかが 追える） */
const 動かない台帳 = 読む('exally-missing-2026-09-07.txt');

/* ★本番と 同じ物を 積む★（片方 忘れると 13個が 嘘に なる） */
const HFns = require_(path.join(ROOT, 'hyperformula.full.min.js'));
const EF = require_(path.join(ROOT, 'exally-formula.js'));
const 積1 = EF.registerExallyFunctions(HFns) === true;
const HF0 = HFns.HyperFormula;
const H = Object.assign(Object.create(HF0), HFns,
  { registerFunctionPlugin: HF0.registerFunctionPlugin.bind(HF0) });
const FXP = require_(path.join(ROOT, 'lib/formula-extra-plug.js'));
const 積2 = FXP.つなぐ(H, require_(path.join(ROOT, 'lib/formula-extra.js')));
/* ★2026-09-06 に 足した 3つ目★＝docs/measured/ の 一覧から 作った 物 */
const FNP = require_(path.join(ROOT, 'lib/formula-nokori-plug.js'));
const 積3 = FNP.つなぐ(H, require_(path.join(ROOT, 'lib/formula-nokori.js')));
/* ★2026-09-07 に 足した 4つ目★＝お金（債券・減価償却）22個
   ★本番の book.html が 積んでいる 物を ここでも 全部 積む★
   （1つ 忘れると その 分が「動かない」に 見えて ★嘘の 数★に なる） */
const FKP = require_(path.join(ROOT, 'lib/formula-kane-plug.js'));
const 積4 = FKP.つなぐ(H, require_(path.join(ROOT, 'lib/formula-kane.js')));
/* ★2026-09-07 に 足した 5つ目★＝予測・統計・単位 8個（CONVERT は ここへ 移した） */
const FYP = require_(path.join(ROOT, 'lib/formula-yosoku-plug.js'));
const 積5 = FYP.つなぐ(H, require_(path.join(ROOT, 'lib/formula-yosoku.js')));
/* ★2026-09-07 に 足した 6つ目★＝外へ 出る 4個
   ★ここでは 外へ 1回も 出さない★＝道具を「必ず 断る 作り物」に する
   ⇒ 見たいのは ★名前が 通るか★だけ（答えは 別の 試験 tests/formula-soto.test.mjs） */
const FSP = require_(path.join(ROOT, 'lib/formula-soto-plug.js'));
const 積6 = FSP.つなぐ(H, require_(path.join(ROOT, 'lib/formula-soto.js')), {
  取る: async () => { throw new Error('ここでは 外へ 出ません'); },
  聞く: async () => { throw new Error('ここでは AI に 聞きません'); },
  再計算: () => {},
});
/* ★2026-09-07 に 足した 7つ目★＝XML から 取り出す
   ★XPath は ブラウザの 物★＝ここでは jsdom の 同じ 物を 入れる
   （jsdom が 無ければ 入れない＝その時は #VALUE! に なるので ★偽の 緑に ならない★） */
let XML部品 = null;
try {
  const { JSDOM } = require_('jsdom');
  const w = new JSDOM('').window;
  XML部品 = { DOMParser: w.DOMParser, XPathResult: w.XPathResult };
} catch (e) { XML部品 = null; }
const 積7 = require_(path.join(ROOT, 'lib/formula-filterxml-plug.js'))
  .つなぐ(H, require_(path.join(ROOT, 'lib/formula-filterxml.js')), () => XML部品);
/* ★2026-09-07 に 足した 8つ目★＝CELL（そのマスが どう 見えているか）
   ★見た目は 画面が 入れる★ので ここでは ★何も 入れない★
   ⇒ 既定の 見た目（＝まっさらな Excel と 同じ）で 答える＝★空回りでは ない★ */
const 積8 = require_(path.join(ROOT, 'lib/formula-cell-plug.js'))
  .つなぐ(H, require_(path.join(ROOT, 'lib/formula-cell.js')), null);

T('★★本番と 同じ物を 積めた（片方 忘れると 13個が 嘘に なる）★★', () => {
  if (!積1) throw new Error('exally-formula の プラグインを 積めていない');
  if (!(積2 > 0)) throw new Error('formula-extra-plug を 積めていない（' + 積2 + '本）');
  if (!(積3 > 0)) throw new Error('formula-nokori-plug を 積めていない（' + 積3 + '本）');
  if (!(積4 > 0)) throw new Error('formula-kane-plug を 積めていない（' + 積4 + '本）');
  if (!(積5 > 0)) throw new Error('formula-yosoku-plug を 積めていない（' + 積5 + '本）');
  if (!(積6 > 0)) throw new Error('formula-soto-plug を 積めていない（' + 積6 + '本）');
  if (!(積7 > 0)) throw new Error('formula-filterxml-plug を 積めていない（' + 積7 + '本）');
  if (!(積8 > 0)) throw new Error('formula-cell-plug を 積めていない（' + 積8 + '本）');
});

const hf = HF0.buildEmpty({ licenseKey: 'gpl-v3' });
const SID = hf.getSheetId(hf.addSheet('S'));
/* ★JS層に 本番と 同じ engine を 渡す★（渡さないと ①が 丸ごと 素通りする） */
/* ★本番と 同じ★＝book.html は hf を 作った後 initExallyFormula(hf) を 1回 呼ぶ */
EF.initExallyFormula(hf);
const JS層 = EF._jsComputeFormula || (EF.__test && EF.__test._jsComputeFormula) || null;

T('★★本番の 1段目（JS層）に 手が 届く★★', () => {
  /* ★ここが 無いと DSUM・OFFSET・LINEST 等を「無い」と 誤判定する★
     （2026-09-06 の 1回目の 実測が それだった） */
  if (typeof JS層 !== 'function') {
    throw new Error('★_jsComputeFormula に 手が 届かない★＝1段目を 飛ばして 測る事に なります');
  }
});

/* ★引数は 決め打たない★＝足りないと ★動く物を 死んだと 誤判定する★ */
/* ★★形の 違う 関数を 入れ忘れると「動くのに 無い」と 出る★★（2026-09-06 実測）
   ⇒ LAMBDA を 取る 物（MAP/REDUCE/SCAN/BYROW/MAKEARRAY）は
     ★普通の 引数では 一度も 当たらない★＝★動いているのに「動かない」に 落ちた★
   ⇒★引数の 形も 手で 決め打たない★＝★当たるまで 増やす★ */
const 候補 = ['(1)', '()', '(A1:A2,1)', '(1,1)', '(A1:A2)', '(1,1,1)', '("a")', '(A1)', '(1,1,1,1)',
  '(A1:B2,1,A1:A2)', '(A1:A2,A1:A2)',
  '(A1:A2,LAMBDA(v,v*2))', '(0,A1:A2,LAMBDA(a,b,a+b))', '(2,2,LAMBDA(r,c,r*c))',
  '(x,2,x*3)', '(x,x+1)', '("<a><b>1</b></a>","//b")'];
/* ★台を 建てる★（★読めなければ null＝今まで通り 2段で 押す／黙って 緑に しない★） */
let 板 = null;
try {
  const SH = require_(path.join(ROOT, 'lib/shiki-hyou.js'));
  板 = SH.表();
  板.打つ('A1', 1); 板.打つ('A2', 2); 板.打つ('A3', 3);
} catch (e) { 板 = null; }

function 押す(f) {
  for (const a of 候補) {
    const 式 = '=' + f + a;
    /* ①JS層 */
    if (typeof JS層 === 'function') {
      /* ★★投げただけでは「動く」と 数えない★★（2026-09-07・指示役が 見つけた 穴A）
         ⇒ 下の ②engine では 同じ事を 2026-09-07 に 直していたのに
           ★こちら（①JS層）は 手つかずだった★＝★1か所 直して 終わりにするな★
         ⇒★実測★ この 直しで 減る 数は ★0個★（今は 誰も ここに 頼っていない）
           ⇒★でも 直す★＝★いつか 火が つく 罠★を 残さない */
      try { if (JS層(0, 式) !== null) return true; } catch (e) { /* 数えない */ }
    }
    /* ══ ★★2026-09-19 ... ★台の 段を 足しました★★★ ══
         ★訳★ ... 本番の 道は ★3段★です
             ①JS層（`_jsComputeFormula`） → ②★台★（`lib/shiki-hyou.js`） → ③借り物
           ＝`book.html:recalcSheet` の 実物
         ★この 門は ①と ③しか 押して いませんでした★
         ⇒`docs/measured/honban-no-michi.mjs` の 断り
           「★この 段を 飛ばすと 『うちが 負けて いる』向きの 偽の 答えが 出ます★」
         ⇒★実際に 起きて いました★ ... BYROW / BYCOL / LAMBDA / PHONETIC を
           ★「動かない」と 読んで いた★（★本番では 動いて います★）
         ⇒★★門を 甘く したのでは ありません＝★見る 範囲★を 本番に 揃えました★★ */
    if (板) {
      try {
        板.打つ('ZZ9990', 式);
        const _v = 板.値('ZZ9990');
        if (_v && _v.溢れ === true) return true;
        if (_v && !(_v.型 === '誤' && String(_v.値) === '#NAME?')) return true;
      } catch (e) { /* ★投げただけでは「動く」と 数えない★ */ }
    }
    /* ③engine（借り物） */
    let 後; try { 後 = EF.convertFormula(式); } catch (e) { continue; }
    try {
      hf.setSheetContent(SID, [[1, 3], [2, 4], [後]]);
      const v = hf.getCellValue({ sheet: SID, row: 2, col: 0 });
      if (!(v && v.type === 'NAME')) return true;
    } catch (e) { /* ★投げただけでは「動く」と 数えない★
      （2026-09-07 … FIELDVALUE が これで「動く」に 数えられていた。
        本当は 押すと ★#NAME?★＝動かない。
        ⇒★★『名前が エンジンに 在る』と『答えが 返る』は 別★★） */ }
  }
  return false;
}

const 食 = 食い違い(動かない台帳, 押す, 全部);

T('★★「動かない」と 書いた 物が 本当に 動かない（AIが 客に 嘘を 言わない）★★', () => {
  if (食.動かないと書いたのに動く.length) {
    throw new Error('★動くのに「動かない」と 書いてあります★: '
      + 食.動かないと書いたのに動く.join(' / ')
      + '\n   → docs/measured/exally-missing-2026-09-07.txt を 測り直してください');
  }
});

T('★★動いていた 物が 動かなく なっていない（後戻りを 止める）★★', () => {
  if (食.動くはずが動かない.length) {
    throw new Error('★前は 動いていた 物が 動きません★: ' + 食.動くはずが動かない.join(' / '));
  }
});

/* ★★『動く』と 呼ぶのを やめました★★（2026-09-07・指示役の 決定）
 *
 *  ★この 数の 本当の 意味★
 *    上の `押す(f)` は ★引数の 型紙 17通りを 519個 全部に 当てて★
 *    ★1つでも #NAME? で なければ 数える★ 作りです。
 *    ⇒★★答えが 合っているかは 1度も 見ていません★★
 *    ⇒★★他の 形も 1度も 押していません★★
 *    ⇒ だから これは『動く』では なく ★『名前が 通った』★です。
 *
 *  ★★数を 1つに 混ぜない★★（混ぜた 瞬間に 一部が 消える）
 *    ・答えを 実Excel と 突き合わせた … 別に 数える
 *    ・名前が 通っただけ ……………… ここ
 *    ・名前も 通らない ………………… ここ
 */
console.log('\n── 実測（★本番の 道★＝JS層 → convertFormula → engine） ──');
console.log('  実Excel が 知っている … ' + 全部.length + '個');
console.log('  ★★名前が 通った ……… ' + 食.動いた数 + '個（'
  + (食.動いた数 / 全部.length * 100).toFixed(1) + '%）★★');
console.log('      ⇒★これは「合っている」では ありません★');
console.log('        ＝★引数の 型紙 17通りの うち 1つが #NAME? を 返さなかった★だけ');
console.log('  ★名前も 通らない ……… ' + (全部.length - 食.動いた数) + '個');
console.log('');
console.log('  ★★『合っているか』は ★まだ 測っていません★★');
console.log('    ⇒ 今の 押し方は ★#NAME? だけ★を 赤に しています。');
console.log('      #VALUE! も #REF! も #NUM! も #N/A も ★全部 緑★です（穴B）。');
console.log('    ⇒★★でも「誤りしか 返らない＝壊れている」では ありません★★');
console.log('      実測 … `NA()` → `#N/A` は ★これが 正しい 動き★');
console.log('             `CELL("a")` → `#VALUE!` も ★実Excel と 同じ★');
console.log('             （CELL は 2026-09-07 に ★411本★ 実Excel と 突き合わせ済み）');
console.log('             お金の 関数に `(1)` を 渡して #NUM! も ★断るのが 正しい★');
console.log('    ⇒ 本当の 病気は ★17通りの 型紙を 519個 全部に 当てている★事です。');
console.log('    ⇒★いつ 全部 測るかは 司さんの 決め（お金と 順番）★');
console.log('    ⇒ 詳しく … docs/measured/ugoku-tana-no-uchiwake.md');
console.log('  積んだ プラグイン ……… exally-formula ' + (積1 ? '○' : '×')
  + ' ／ formula-extra-plug ' + 積2 + '本 ／ formula-nokori-plug ' + 積3 + '本'
  + ' ／ formula-kane-plug ' + 積4 + '本 ／ formula-yosoku-plug ' + 積5 + '本'
  + ' ／ formula-soto-plug ' + 積6 + '本 ／ formula-filterxml-plug ' + 積7 + '本'
  + ' ／ formula-cell-plug ' + 積8 + '本'
  + ' ／ JS層 ' + (typeof JS層 === 'function' ? '○' : '×'));
console.log('\n' + pass + ' passed, ' + fail + ' failed');
hf.destroy();
process.exit(fail ? 1 : 0);
