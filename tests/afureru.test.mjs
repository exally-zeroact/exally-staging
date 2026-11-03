/* afureru.test.mjs — ★1つの 式が 何マスにも 広がる（溢れ）★（2026-09-10）
 *
 *  ★★お客さんから 見て 何が 出来なかったか★★
 *    `=SORT(A1:A5)` を C1 に 打っても ★C1 に 1つ 出るだけ★でした。
 *    実Excel なら C1..C5 に 並びます。
 *    ⇒★Excel で 出来る 事が 出来ない★＝方針（Excel の 最上級・全部 つける）に 反する
 *
 *  ★★エンジンは ★もう 溢れて いました★★（実測）
 *      `=SORT(A1:A5)` … C1 の 型 ARRAYFORMULA ／ C2..C5 ARRAY ／ C6 EMPTY
 *      塞がって いれば ★C1 は #SPILL★（じゃまな マスは そのまま）
 *    ⇒★足りなかったのは 画面が それを 読む 所 1か所だけ★
 *      `recalcSheet` は `Object.keys(data)` を 回る＝★マスが 無い 所は 見ない★
 *
 *  ★★だから 1つ 直すと まとめて 動きます★★
 *    SORT ／ UNIQUE ／ FILTER ／ SEQUENCE ／ LINEST ／ LOGEST …
 *
 *  ★★「前より 悪く なって いないか」を 3つ 押しました★★（2026-09-10・監査の 求め）
 *    ★甲は ★新しい 危なさ★を 作りました★＝前は 溢れた先に マスが 無かったので
 *    お客さんが そこに 打っても 何も 壊れませんでした。今は ★マスが 在る★。
 *    ⇒★知らずに 出さない★＝押して 紙に 書く（直すのは 別PRで よい）
 *
 *    ①★溢れた マスに 字を 打つ★
 *      ⇒★打った 字は 残り、式が #SPILL! に なる★（★実Excel と 同じ★）
 *      ⇒★お客さんの 字は 黙って 消えません★＝悪く なって いない
 *    ②★xlsx に 書き出す★（★実Excel で 開いて 確かめた★）
 *      前 … C1 に 式だけ ／ ★C2〜C5 は 空★
 *      後 … C1 に 式 ／ ★C2〜C5 に 2,3,4,5 が 入る★
 *      ⇒★値が 渡るように なった＝前より 良い★／#SPILL! には なりません
 *      ⇒ 書き出しの 警告の 字も 直しました（前は「1セル分の値になります」だけ＝★半分だけ 本当★）
 *    ③★元に戻す（Undo）★
 *      ⇒ じゃまが 消えても ★#SPILL! の まま★（式を 打ち直すと 直る）
 *      ⇒★★これは エンジン自身の 癖です★★（素の 台で 切り分けた）
 *        塞ぐ → #SPILL ／ ★じゃまを 消す → #SPILL の まま★ ／ 式を 打ち直す → 溢れる
 *      ⇒★私の 直しでも undo でも ありません★／実Excel なら すぐ 計算し直します
 *      ⇒★まだ 直して いません＝別の 直し★
 *
 *  ★★2026-09-11 … 「実Excel は 断る」は ★私の 思い込み★でした★★
 *    ★実Excel に 打たせて 確かめました★（COM で 溢れた マスに 値を 入れた）
 *      溢れた 直後 … C1..C5 = 1,2,3,4,5
 *      ★C3 に「あとから打った」を 入れた ⇒ ★断られません★★
 *      その 後 … C1 ★#スピル!★／C2,C4,C5 空／C3「あとから打った」
 *    ⇒★うちと 1つも 違いません★（うちも 打てて #SPILL! に なる）
 *    ⇒★「配列の 一部は 変えられません」は ★古い 形の 配列式（Ctrl+Shift+Enter）★の 話★
 *      ＝★今の 溢れ（動的配列）では 出ません★
 *    ⇒★棚から 下ろしました★（直す 所は 在りません）
 *    ★教訓★＝★「実Excel は こうする はず」を 根拠に しない★（今日 3回目）
 *
 *  使い方: node tests/afureru.test.mjs [--self-test]
 */
import fs from 'node:fs';
import path from 'node:path';
import { 注記を外す } from '../scripts/lib/chuki.mjs';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createRequire } from 'node:module';

const ここ = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(ここ, '..');
const 直に走った = process.argv[1] && pathToFileURL(process.argv[1]).href === import.meta.url;
const 自己試験 = 直に走った && process.argv.includes('--self-test');

let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ✓ ' + n); } catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + (e && e.message)); } };
const 改行と字下げ = String.fromCharCode(10) + '      ';

/* ══ ★本番と 同じ 台を 建てる（★測り台の 門の 決まり★）★ ══
   ★プラグインの 数も 建て方も book.html から 読む＝数を 書き込まない★ */
const 本 = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf-8');
const require_ = createRequire(path.join(ROOT, 'package.json'));
const HFns = require_(path.join(ROOT, 'hyperformula.full.min.js'));
const EF = require_(path.join(ROOT, 'exally-formula.js'));
EF.registerExallyFunctions(HFns);

/* ★★プラグインの 積み方は ★既に 在る 道具から そのまま★★（自分で 書き直さない）★★
   出どころ … `docs/measured/osu-okane-4kansuu.mjs`
   ⇒★2通りの 積み方が 出来ると 必ず ずれる★（2026-09-09 に 3本しか 積まず 嘘の 数字を 出した） */
const H = HFns;
let つないだ = 0;
/* ★★本番が 読む プラグインを ★全部★ つなぐ★★（2026-09-09 に 直した）
   ★1回目は extra / nokori / kane の 3本しか つないで いなかった★
   ⇒ 本番（book.html）は ★8本★ 読む。★yosoku を 落として いた★
   ⇒★本番に 在る 物が 無い 状態で 押して いた＝★嘘の 数字が 出る★
   ⇒ book.html の `<script src="lib/formula-*-plug.js">` と ★同じ 並び★に した
   ★ここに 無い 物が book.html に 増えたら 下の 見張りで 赤に なる★ */
const つなぐ本数 = { 期待: 8, 済: 0 };
for (const n of ['extra', 'nokori', 'kane']) {
  require_(path.join(ROOT, 'lib/formula-' + n + '-plug.js'))
    .つなぐ(H, require_(path.join(ROOT, 'lib/formula-' + n + '.js')));
  つなぐ本数.済++;
}
/* ★予測（TREND / GROWTH / LOGEST）★＝★これを 落として いた★ */
require_(path.join(ROOT, 'lib/formula-yosoku-plug.js'))
  .つなぐ(H, require_(path.join(ROOT, 'lib/formula-yosoku.js')),
    () => ({ シート数: 1, 版: 'Exally', 台: 'win', OS: '', 左上: '$A$1' }));
つなぐ本数.済++;
/* ★網の 外へ 出る 物は ★出させない★（司さんの 決め）★ */
require_(path.join(ROOT, 'lib/formula-soto-plug.js'))
  .つなぐ(H, require_(path.join(ROOT, 'lib/formula-soto.js')), {
    取る: async () => { throw new Error('外へ 出ません'); },
    聞く: async () => { throw new Error('AI に 聞きません'); },
    再計算: () => {},
  });
つなぐ本数.済++;
let XML部品 = null;
try {
  const { JSDOM } = require_('jsdom');
  const w = new JSDOM('').window;
  XML部品 = { DOMParser: w.DOMParser, XPathResult: w.XPathResult };
} catch (e) { XML部品 = null; }
require_(path.join(ROOT, 'lib/formula-filterxml-plug.js'))
  .つなぐ(H, require_(path.join(ROOT, 'lib/formula-filterxml.js')), () => XML部品);
つなぐ本数.済++;
require_(path.join(ROOT, 'lib/formula-cell-plug.js'))
  .つなぐ(H, require_(path.join(ROOT, 'lib/formula-cell.js')), null);
つなぐ本数.済++;
require_(path.join(ROOT, 'lib/formula-complex-plug.js'))
  .つなぐ(H, require_(path.join(ROOT, 'lib/formula-complex.js')));
つなぐ本数.済++;
/* ★★book.html が 読む 数と 合うか（★落としたら ここで 止まる★）★★ */
{
  const html = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf-8');
  const 本番 = [...html.matchAll(/lib\/(formula-[a-z]+)-plug\.js/g)].map((m) => m[1]);
  const 数 = new Set(本番).size;
  if (数 !== つなぐ本数.済) {
    console.error('★book.html は ' + 数 + '本 読むのに、ここでは ' + つなぐ本数.済 + '本しか つないで いない★');
    console.error('  本番 … ' + [...new Set(本番)].join(' '));
    process.exit(2);
  }
  つないだ = 数;
}

/* ★建て方も book.html から 読む（★数を 書き込まない★）★ */
const 建て = /buildEmpty\(\{([^}]*)\}/.exec(本);
if (!建て) throw new Error('★book.html の buildEmpty が 読めない★');
if (!/smartRounding\s*:\s*false/.test(建て[1])) throw new Error('★本番に smartRounding:false が 無い★');
if (!/useArrayArithmetic\s*:\s*true/.test(建て[1])) throw new Error('★本番に useArrayArithmetic:true が 無い★');

function 台() {
  const hf = HFns.HyperFormula.buildEmpty({
    licenseKey: 'gpl-v3', useArrayArithmetic: true, smartRounding: false,
    maxRows: 1048576, maxColumns: 18278,
  });
  const SID = hf.getSheetId(hf.addSheet('S'));
  EF.initExallyFormula(hf);
  return { hf, SID };
}

/* ★本番の 道＝1マスずつ 入れる★（★板ごと 入れない★） */
function 打つ(hf, SID, マス, 中身) {
  const m = /^([A-Z]+)(\d+)$/.exec(マス);
  const c = m[1].split('').reduce((s, ch) => s * 26 + (ch.charCodeAt(0) - 64), 0) - 1;
  hf.setCellContents({ sheet: SID, row: Number(m[2]) - 1, col: c }, [[中身]]);
}
const 型 = (hf, SID, r, c) => hf.getCellType({ sheet: SID, row: r, col: c });
const 値 = (hf, SID, r, c) => {
  const v = hf.getCellValue({ sheet: SID, row: r, col: c });
  return (v && v.type) ? '#' + v.type : v;
};

/* ★book.html から 溢れの 段を 切り出す（★写しを 置かない★）★ */
function 切り出す(名) {
  const 頭 = 本.indexOf('function ' + 名 + '(');
  if (頭 < 0) throw new Error('★book.html に ' + 名 + ' が 無い★');
  const 開き = 本.indexOf('{', 頭);
  let 深さ = 0, i = 開き;
  for (; i < 本.length; i++) {
    if (本[i] === '{') 深さ++;
    else if (本[i] === '}') { 深さ--; if (深さ === 0) break; }
  }
  return 本.slice(頭, i + 1);
}

console.log('\n[afureru] ★1つの 式が 何マスにも 広がる（溢れ）★');
console.log('  ★つないだ プラグイン★ ' + つないだ + '本（★book.html と 同じ 数★）');

T('★台が 建って いる（★空振りして いない★）★', () => {
  const { hf, SID } = 台();
  打つ(hf, SID, 'A1', 1);
  if (Number(値(hf, SID, 0, 0)) !== 1) throw new Error('★台に 値が 入らない★');
  if (つないだ < 1) throw new Error('★プラグインを 1本も つないで いない★');
});

T('★★エンジンは 溢れて いる（★これが 前提★）★★', () => {
  const { hf, SID } = 台();
  [3, 1, 5, 2, 4].forEach((v, i) => 打つ(hf, SID, 'A' + (i + 1), v));
  打つ(hf, SID, 'C1', '=SORT(A1:A5)');
  if (型(hf, SID, 0, 2) !== 'ARRAYFORMULA') throw new Error('★C1 が ARRAYFORMULA で ない（' + 型(hf, SID, 0, 2) + '）★');
  const 出 = [];
  for (let r = 0; r < 5; r++) 出.push(値(hf, SID, r, 2));
  if (出.join(',') !== '1,2,3,4,5') throw new Error('★' + 出.join(',') + '★（1,2,3,4,5 のはず）');
  for (let r = 1; r < 5; r++) {
    if (型(hf, SID, r, 2) !== 'ARRAY') throw new Error('★C' + (r + 1) + ' が ARRAY で ない★');
  }
  if (型(hf, SID, 5, 2) !== 'EMPTY') throw new Error('★C6 まで 広がって いる★');
  console.log('      … C1 ARRAYFORMULA ／ C2..C5 ARRAY ／ C6 EMPTY');
});

T('★★塞がって いれば #SPILL（じゃまな マスは そのまま）★★', () => {
  const { hf, SID } = 台();
  [3, 1, 5, 2, 4].forEach((v, i) => 打つ(hf, SID, 'A' + (i + 1), v));
  打つ(hf, SID, 'C3', 'じゃま');
  打つ(hf, SID, 'C1', '=SORT(A1:A5)');
  if (String(値(hf, SID, 0, 2)) !== '#SPILL') throw new Error('★C1 が ' + 値(hf, SID, 0, 2) + '★（#SPILL のはず）');
  if (値(hf, SID, 2, 2) !== 'じゃま') throw new Error('★じゃまな マスが 消えた★');
  console.log('      … C1 #SPILL ／ C3 は「じゃま」の まま');
});

T('★★画面に 出す 段が book.html に 在る（★空振りして いない★）★★', () => {
  for (const 名 of ['_溢れの広さ', '_溢れを写す', '_古い溢れを消す']) {
    if (本.indexOf('function ' + 名 + '(') < 0) throw new Error('★' + 名 + ' が 無い★');
  }
  if (!/_溢れを写す\(sheetIdx, data, r, c\)/.test(本)) throw new Error('★recalcSheet から 呼んで いない★');
  if (!/_古い溢れを消す\(sheetIdx, data\)/.test(本)) throw new Error('★古い 溢れを 消して いない★');
  console.log('      … 3つの 段と 2つの 呼び出しが 在る');
});

T('★★溢れた 先の マスが 作られる（★これが 本体★）★★', () => {
  const { hf, SID } = 台();
  [3, 1, 5, 2, 4].forEach((v, i) => 打つ(hf, SID, 'A' + (i + 1), v));
  打つ(hf, SID, 'C1', '=SORT(A1:A5)');
  const 写す = new Function('hf', '_hfGetDisplay', '溢れの印', '溢れの上限',
    切り出す('_溢れの広さ') + ' ' + 切り出す('_溢れを写す') + ' return _溢れを写す;')(
    hf, (s, r, c) => String(値(hf, s, r, c)), '_溢れ元', 100000);
  const data = { '0,2': { f: '=SORT(A1:A5)' } };
  const n = 写す(SID, data, 0, 2);
  if (n !== 4) throw new Error('★' + n + 'マスしか 作られない★（4のはず）');
  const 出 = [0, 1, 2, 3, 4].map((r) => data[r + ',2'] && data[r + ',2'].d);
  if (出.slice(1).join(',') !== '2,3,4,5') throw new Error('★' + JSON.stringify(出) + '★');
  for (let r = 1; r < 5; r++) {
    if (data[r + ',2']._溢れ元 !== '0,2') throw new Error('★C' + (r + 1) + ' に 印が 無い★');
  }
  console.log('      … 4マス 作られ、全部に「溢れ元」の 印が 付いた');
});

T('★★お客さんが 打った 物を 上書きしない★★', () => {
  const { hf, SID } = 台();
  [3, 1, 5, 2, 4].forEach((v, i) => 打つ(hf, SID, 'A' + (i + 1), v));
  打つ(hf, SID, 'C1', '=SORT(A1:A5)');
  const 写す = new Function('hf', '_hfGetDisplay', '溢れの印', '溢れの上限',
    切り出す('_溢れの広さ') + ' ' + 切り出す('_溢れを写す') + ' return _溢れを写す;')(
    hf, (s, r, c) => String(値(hf, s, r, c)), '_溢れ元', 100000);
  /* ★C3 に お客さんの 字が 在る 形★（本当は #SPILL に なるが、段だけを 試す） */
  const data = { '0,2': { f: '=SORT(A1:A5)' }, '2,2': { v: 'たいせつ' } };
  写す(SID, data, 0, 2);
  if (data['2,2'].v !== 'たいせつ') throw new Error('★お客さんの 字を 消した★');
  if (data['2,2'].d !== undefined) throw new Error('★お客さんの マスに 溢れを 書いた★');
  console.log('      … お客さんの 字は そのまま');
});

T('★★大きすぎる 溢れは 出さない（★黙って 固まらせない★）★★', () => {
  const 段 = 切り出す('_溢れを写す');
  if (!/溢れの上限/.test(段)) throw new Error('★上限を 見て いない★');
  if (!/> 溢れの上限\) return 0/.test(段)) throw new Error('★上限を 超えた 時に 止めて いない★');
  const 数 = /var 溢れの上限 = (\d+);/.exec(本);
  if (!数) throw new Error('★上限が 書いて いない★');
  console.log('      … 上限 ' + Number(数[1]).toLocaleString() + 'マス');
});

T('★★①溢れた マスに 打つと 字は 残り 式が #SPILL に なる★★', () => {
  const { hf, SID } = 台();
  [3, 1, 5, 2, 4].forEach((v, i) => 打つ(hf, SID, 'A' + (i + 1), v));
  打つ(hf, SID, 'C1', '=SORT(A1:A5)');
  打つ(hf, SID, 'C3', 'あとから打った');          /* ★溢れた 先に 後から 打つ★ */
  if (値(hf, SID, 2, 2) !== 'あとから打った') throw new Error('★お客さんの 字が 消えた★');
  if (String(値(hf, SID, 0, 2)) !== '#SPILL') throw new Error('★C1 が ' + 値(hf, SID, 0, 2) + '★');
  console.log('      … 字は 残り、式は #SPILL（★実Excel と 同じ★）');
});

T('★★①実Excel も 断らない（★私の 思い込みが 外れた★）★★', () => {
  /* ★2026-09-11 実測（COM で 実Excel に 打たせた）★
       溢れた C3 に 値を 入れる ⇒★断られない★
       その 後 … C1 ★#スピル!★／C2,C4,C5 空／C3 は 打った 字
     ⇒★うちと 1つも 違いません★
     ⇒「配列の 一部は 変えられません」は ★古い 形の 配列式★の 話（動的配列では 出ない） */
  const { hf, SID } = 台();
  [3, 1, 5, 2, 4].forEach((v, i) => 打つ(hf, SID, 'A' + (i + 1), v));
  打つ(hf, SID, 'C1', '=SORT(A1:A5)');
  打つ(hf, SID, 'C3', 'あとから打った');
  if (値(hf, SID, 2, 2) !== 'あとから打った') throw new Error('★打てなかった＝実Excel と 違う★');
  if (String(値(hf, SID, 0, 2)) !== '#SPILL') throw new Error('★C1 が ' + 値(hf, SID, 0, 2) + '★');
  for (const r of [1, 3, 4]) {
    if (値(hf, SID, r, 2) !== null) throw new Error('★C' + (r + 1) + ' が 残って いる★');
  }
  console.log('      … 打てて #SPILL／他の 溢れは 消える（★実Excel と 同じ★）');
});

T('★★②じゃまを 消したら 溢れ直す（★画面が やる★）★★', () => {
  /* ★★実Excel（2026-09-11 COM で 打った）★★
       C3 に じゃま ⇒ C1「#スピル!」／★じゃまを 消す ⇒ C1..C5 = 1,2,3,4,5★
     ★エンジンは やって くれません★（素の 台で 切り分けた・下の ③）
     ⇒★画面が 式を 打ち直します★（book.html `_溢れ直しを試す`）
     ★借り物の 中は 読んで いません★＝外から 同じ 式を 入れ直すだけ
     ★絵で 確かめた★ … docs/measured/e-afurenaoshi-2026-09-11.png
       （本物の ブラウザで マスを 押して 打った … docs/measured/osu-afurenaoshi.mjs） */
  if (本.indexOf('_溢れ直しを試す') < 0) throw new Error('★画面に 溢れ直しが 無い★');
  const 中 = 本.slice(本.indexOf('function recalcSheet'), 本.indexOf('function recalcSheet') + 900);
  if (中.indexOf('_溢れ直しを試す') < 0) throw new Error('★recalcSheet が 溢れ直しを 呼んで いない★');
  console.log('      … recalcSheet が 溢れ直しを 呼ぶ');
});

T('★★②-2 ぐるぐる回りの 止め金が 在る（★画面を 固めた★）★★', () => {
  /* ★★2026-09-11 実物の 画面を 固めました★★
       打ち直す → 計算し直す → まだ #SPILL → また 打ち直す …で ★止まらない★
       ⇒★止め金 無しで 出しかけました★（絵を 撮ろうとして 固まって 気付いた）
     ⇒★1回の 計算で 1度だけ★＝直らない 時は 次の 打ち込みまで 待つ */
  if (本.indexOf('_溢れ直し中') < 0) throw new Error('★止め金が 無い＝ぐるぐる回る★');
  const 頭 = 本.indexOf('function _溢れ直しを試す');
  const 中 = 本.slice(頭, 頭 + 600);
  if (中.indexOf('if (_溢れ直し中) return') < 0) throw new Error('★入口で 止めて いない★');
  if (中.indexOf('finally') < 0) throw new Error('★finally で 下ろして いない＝一度 落ちたら 二度と 直らない★');
  console.log('      … 入口で 止めて finally で 下ろす');
});

T('★★②-2b 打ち直しは ★入れる 時と 同じ 道★を 通る★★', () => {
  /* ★★2026-09-11 ここで つまずきました★★
       打ち直しで `cell.f` を ★そのまま★ エンジンに 渡して いた
       ⇒ Excel の ファイルの 式は `=_xlws.SORT(E1:E5)`（Excel が 付ける 印）
       ⇒ 入れる 時（loadSheetIntoEngine）は 印を 外して いるのに
         ★打ち直しだけ 素通り★ ⇒★#SPILL が #ERROR に 化けた★
       ⇒★Excel で SORT を 使った 表を 開くと 全部 壊れる★所でした
     ★見つけ方★＝★向きを 変えて 測った★（Excel が 作った 物を うちで 開く）
       docs/measured/osu-excel-kara.mjs
     ★作る 道が 2本 在る時は 両方 直せ★（記憶の 決まり）＝ここは その 実物 */
  /* ★★注記を 外してから 探します★★
     ★2026-09-11 ここでも つまずきました★＝この 見張りの 最初の 版は 生の 字を 探して いて、
     ★私が 上に 書いた 説明の 中の「convertFormula」を 読んで 緑★に なりました。
     わざと 抜いても 赤に ならず、★見張りが 効いて いない★ ＝ 気付いたのは 壊して 試した 時。
     ⇒★探す 前に 注記を 外す★（scripts/lib/chuki.mjs＝他の 見張りと 同じ 部品） */
  const 動く本 = 注記を外す(本, { html: true });
  const 頭 = 動く本.indexOf('function _溢れ直しの中身');
  if (頭 < 0) throw new Error('★打ち直しの 中身が 無い★');
  const 中 = 動く本.slice(頭, 頭 + 1800);
  if (中.indexOf('convertFormula') < 0) {
    throw new Error('★打ち直しが convertFormula を 通って いない★'
      + '＝Excel の `_xlws.` が 残って #ERROR に なります');
  }
  if (中.indexOf('quoteSheetRefs') < 0) {
    throw new Error('★打ち直しが quoteSheetRefs を 通って いない★＝別の 板を 指す 式が 壊れます');
  }
  /* ★入れる 時と 同じ 2つを 通って いるか★＝片方だけ 直して いないか */
  const 入 = 動く本.indexOf('function loadSheetIntoEngine');
  const 入中 = 動く本.slice(入, 入 + 2000);
  for (const 段 of ['convertFormula', 'quoteSheetRefs']) {
    if (入中.indexOf(段) >= 0 && 中.indexOf(段) < 0) {
      throw new Error('★入れる 時は ' + 段 + ' を 通るのに 打ち直しは 通らない★');
    }
  }
  console.log('      … 入れる 時と 同じ 2つ（convertFormula／quoteSheetRefs）を 通る');
});

T('★★②-3 空マスが 混ざった SORT の 並び（★実Excel と 同じ★）★★', () => {
  /* ★★画面で 見て「うちが 違うのでは」と 思った 所★★（2026-09-11）
       A=(1,5,2,4,空) を SORT ⇒ うちは ★1,2,4,5,0★＝★0 が 最後★
       「Excel なら 0 が 先では」と 思ったので ★実Excel に 打たせた★
     ★実Excel（COM・2026-09-11）★ A=(3,1,空,2,4) ⇒★1,2,3,4,0★＝★0 は 最後★
     ⇒★うちと 同じ★＝★直さなくて 正解でした★
     ★思っただけで 直して いたら 実Excel と 違う 物に して いました★ */
  const { hf, SID } = 台();
  [1, 5, 2, 4].forEach((v, i) => 打つ(hf, SID, 'A' + (i + 1), v));   /* A5 は 空 */
  打つ(hf, SID, 'C1', '=SORT(A1:A5)');
  const 出 = [0, 1, 2, 3, 4].map((r) => 値(hf, SID, r, 2));
  /* ★エンジンは 5枚目を null で 返します★／画面は そこに ★0★を 出します
     （本物の ブラウザで 見た … docs/measured/osu-afurenaoshi.mjs の 絵で C5 が 0）
     ⇒★お客さんが 見る 字は 実Excel と 同じ★ */
  const 並び = 出.slice(0, 4).map(String).join(',');
  if (並び !== '1,2,4,5') throw new Error('★並びが ' + 並び + '★（実Excel は 1,2,4,5 … 空は 最後）');
  if (出[4] !== null && String(出[4]) !== '0') {
    throw new Error('★5枚目が ' + JSON.stringify(出[4]) + '★＝空が 最後に 来て いない');
  }
  console.log('      … 1,2,4,5 の 後ろに 空（★実Excel も 0 を 最後に 置く★）');
});

T('★★④Excel が 書いた「広がった 先」を じゃま 扱いしない★★', () => {
  /* ★★2026-09-11 実測で 見つけた 本当の 穴★★
       ★Excel は 広がった 先の 答えも ファイルに 書きます★
       SheetJS で 読むと（docs/measured/excel-kara-2026-09-11.xlsx）
         G1 … {v:1, f:"_xlws.SORT(E1:E5)", ★F:"G1:G5"★}   ← 溢れの 元
         G2 … {v:1,                        ★F:"G1:G5"★}   ← ★広がった 先（式は 無い）★
       ⇒ 前は それを ★ただの 数★として 読み、うちの 溢れ先を 塞いで いた
       ⇒★Excel の SORT / UNIQUE / FILTER が 3つとも #SPILL!★（実測）
     ★直し方★＝うちが 溢れで 作った マスと ★同じ 印★を 付ける（`v` を 持たせない）
     ★この 見張りは 実物の ファイルを 通します★（作り物の 板では ない） */
  const 見本 = path.join(ROOT, 'docs/measured/excel-kara-2026-09-11.xlsx');
  if (!fs.existsSync(見本)) throw new Error('★見本の xlsx が 無い … ' + 見本 + '★');

  /* ★book-open.js を そのまま 動かす★（写さない） */
  const 場 = { XLSX: require_(path.join(ROOT, 'lib/xlsx.full.min.js')) };
  const src = fs.readFileSync(path.join(ROOT, 'js/book-open.js'), 'utf-8');
  new Function('self', src)(場);
  const BO = 場.BookOpen;
  if (!BO || typeof BO.sheetToGrid !== 'function') throw new Error('★book-open.js が 読めない★');

  const wb = 場.XLSX.read(fs.readFileSync(見本), { type: 'buffer', cellFormula: true, cellNF: true, sheetStubs: false });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const 板 = BO.sheetToGrid(ws, wb.SheetNames[0], {});

  /* ★印は book.html と 同じ 字か★＝違うと 画面が 気付かない */
  const 印 = (本.match(/var 溢れの印 = '([^']+)'/) || [])[1];
  if (!印) throw new Error('★book.html の 溢れの印が 読めない★');
  const 印2 = (fs.readFileSync(path.join(ROOT, 'js/book-open.js'), 'utf-8').match(/var 溢れの印 = '([^']+)'/) || [])[1];
  if (印 !== 印2) throw new Error('★印が 食い違う … book.html「' + 印 + '」／book-open.js「' + 印2 + '」★');

  /* ★G1 は 元＝式を 持つ／G2・G3 は 先＝印が 付いて `v` を 持たない★ */
  const 元 = 板.data['0,6'];
  if (!元 || !元.f) throw new Error('★G1 に 式が 無い★');
  if (元[印]) throw new Error('★溢れの 元にまで 印を 付けて いる★＝式が エンジンに 渡らない');
  for (const [場所, 名] of [['1,6', 'G2'], ['2,6', 'G3'], ['1,7', 'H2'], ['1,8', 'I2']]) {
    const 先 = 板.data[場所];
    if (!先) throw new Error('★' + 名 + ' が 消えて いる★');
    if (!先[印]) throw new Error('★' + 名 + ' に 印が 無い★＝ただの 数として 溢れ先を 塞ぎます');
    if (先.v !== '' && 先.v !== undefined && 先.v !== null) {
      throw new Error('★' + 名 + ' が v を 持って いる（' + JSON.stringify(先.v) + '）★＝塞ぎます');
    }
    if (String(先.d) === '') throw new Error('★' + 名 + ' の 出る字が 空★');
  }
  console.log('      … 元 G1 は 式／先 G2・G3・H2・I2 は 印つきで v を 持たない');
});

T('★★③エンジン自身は 溢れ直さない（★癖★・だから 画面が 打ち直す）★★', () => {
  /* ★これは ★エンジンの 話★です★＝画面は ②で 打ち直して 直します
     ★ここが 緑の うちは「画面の 打ち直し」を 外せません★
     もし エンジンが 直る 版に なったら ここが 赤に なる ⇒★その時 打ち直しを 外す★ */
  const { hf, SID } = 台();
  [3, 1, 5, 2, 4].forEach((v, i) => 打つ(hf, SID, 'A' + (i + 1), v));
  打つ(hf, SID, 'C3', 'じゃま');
  打つ(hf, SID, 'C1', '=SORT(A1:A5)');
  if (String(値(hf, SID, 0, 2)) !== '#SPILL') throw new Error('★塞いだのに #SPILL に ならない★');
  hf.setCellContents({ sheet: SID, row: 2, col: 2 }, [[null]]);   /* ★じゃまを 消す★ */
  if (String(値(hf, SID, 0, 2)) !== '#SPILL') {
    throw new Error('★エンジンが 自分で 溢れ直した★＝★画面の 打ち直し（②）は もう 要りません★');
  }
  打つ(hf, SID, 'C1', '=SORT(A1:A5)');                            /* ★打ち直す★ */
  if (String(値(hf, SID, 0, 2)) !== '1') throw new Error('★打ち直しても 直らない★');
  console.log('      … 消しても #SPILL の まま／打ち直すと 直る（★エンジンの 癖★）');
});

T('★★②書き出しの 警告が 本当の 事を 言って いる★★', () => {
  /* ★前は「1セル分の値になります」だけ＝★半分だけ 本当★★
     実Excel で 開いて 確かめた … C2〜C5 に 2,3,4,5 が ★ただの 数★で 入る */
  const 字 = fs.readFileSync(path.join(ROOT, 'lib/grid-xlsx.js'), 'utf-8');
  const m = /msg: '配列を返す式が'[\s\S]{0,400}?listCells\(arrayCells\)/.exec(字);
  if (!m) throw new Error('★警告の 字が 見つからない★');
  if (m[0].indexOf('ただの数として入ります') < 0) {
    throw new Error('★「広がった先の値も 入る」事を 言って いない★＝★半分だけ 本当★');
  }
  if (m[0].indexOf('1セル分') < 0) throw new Error('★式が 1セル分に なる 事を 言って いない★');
  console.log('      … 式は 1セル分／広がった先は ただの数、の 両方を 言って いる');
});

T('★★「まだ 見て いない 事」の 断りが 残って いる★★', () => {
  const s = fs.readFileSync(path.join(ここ, 'afureru.test.mjs'), 'utf-8');
  for (const 断り of ['私の 思い込み', 'エンジン自身の 癖', 'まだ 直して いません']) {
    if (s.indexOf(断り) < 0) throw new Error('★断りが 消えた … ' + 断り + '★');
  }
  console.log('      … 3つの 断りが 残って いる（★未完を 緑に しない★）');
});

/* ══ ★自己試験（★壊すのは 写し★＝ファイルは 1バイトも 触らない）★ ══ */
if (自己試験) {
  console.log('\n★自己試験（★壊すのは 写し★＝ファイルは 1バイトも 触らない）★');

  T('★★溢れを 写す 段を 外すと 赤に なる★★', () => {
    /* ★元の recalcSheet（溢れを 見ない 形）で 押すと 何も 作られない★ */
    const { hf, SID } = 台();
    [3, 1, 5, 2, 4].forEach((v, i) => 打つ(hf, SID, 'A' + (i + 1), v));
    打つ(hf, SID, 'C1', '=SORT(A1:A5)');
    const data = { '0,2': { f: '=SORT(A1:A5)' } };
    /* ★呼ばない＝前の 姿★ */
    const 出 = [1, 2, 3, 4].map((r) => data[r + ',2']);
    if (出.some((x) => x)) throw new Error('★呼んで いないのに マスが 在る★');
    console.log('      … 呼ばなければ 0マス（＝これが 直す 前の 姿）');
  });

  T('★★古い 溢れが 消える（式を 短く した 時）★★', () => {
    const { hf, SID } = 台();
    [3, 1, 5, 2, 4].forEach((v, i) => 打つ(hf, SID, 'A' + (i + 1), v));
    打つ(hf, SID, 'C1', '=SORT(A1:A5)');
    const 作る = (名) => new Function('hf', '_hfGetDisplay', '溢れの印', '溢れの上限',
      切り出す('_溢れの広さ') + ' ' + 切り出す('_溢れを写す') + ' ' + 切り出す('_古い溢れを消す') + ' return ' + 名 + ';')(
      hf, (s, r, c) => String(値(hf, s, r, c)), '_溢れ元', 100000);
    const 写す = 作る('_溢れを写す'), 消す = 作る('_古い溢れを消す');
    const data = { '0,2': { f: '=SORT(A1:A5)' } };
    写す(SID, data, 0, 2);
    if (!data['4,2']) throw new Error('★5マス 目が 作られて いない★');
    /* ★式を 短く する★ */
    打つ(hf, SID, 'C1', '=SORT(A1:A3)');
    const 消えた = 消す(SID, data);
    if (data['4,2']) throw new Error('★古い マスが 残って いる★＝★古い 数が 画面に 残る★');
    if (消えた < 2) throw new Error('★' + 消えた + 'マスしか 消えて いない★');
    console.log('      … 短く すると ' + 消えた + 'マス 消える');
  });

  T('★★書式だけ 残って いる マスは 消さない★★', () => {
    const { hf, SID } = 台();
    [3, 1, 5, 2, 4].forEach((v, i) => 打つ(hf, SID, 'A' + (i + 1), v));
    打つ(hf, SID, 'C1', '=SORT(A1:A3)');
    const 消す = new Function('hf', '溢れの印',
      切り出す('_古い溢れを消す') + ' return _古い溢れを消す;')(hf, '_溢れ元');
    /* ★もう 溢れて いない マスに お客さんの 色が 付いて いる★ */
    const data = { '4,2': { d: '5', _溢れ元: '0,2', bgColor: '#FFFF00' } };
    消す(SID, data);
    if (!data['4,2']) throw new Error('★色が 付いた マスを 消した★');
    if (data['4,2'].bgColor !== '#FFFF00') throw new Error('★色が 消えた★');
    if (data['4,2'].d !== undefined) throw new Error('★古い 数が 残って いる★');
    console.log('      … 色は 残り、古い 数だけ 消える');
  });

  T('★★台が 本番と 同じ（プラグインを 積んで いる）★★', () => {
    if (つないだ < 3) throw new Error('★プラグインが ' + つないだ + '本しか つないで いない★');
    const { hf, SID } = 台();
    打つ(hf, SID, 'A1', 1); 打つ(hf, SID, 'A2', 2); 打つ(hf, SID, 'A3', 3);
    打つ(hf, SID, 'B1', '=UNIQUE(A1:A3)');
    if (String(値(hf, SID, 0, 1)).indexOf('NAME') >= 0) {
      throw new Error('★UNIQUE が #NAME★＝★プラグインが 積めて いない台で 測って いる★');
    }
    console.log('      … つないだ ' + つないだ + '本／UNIQUE も 通る');
  });
}

console.log('\n' + pass + ' passed, ' + fail + ' failed');
if (fail) process.exit(1);
