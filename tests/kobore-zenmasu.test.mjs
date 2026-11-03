/* kobore-zenmasu.test.mjs -- 溢れた先の ★全マス★が 実Excel と 同じか（2026-09-19）
 *
 *  ★★なぜ 在るか★★
 *    今まで の 分母は ★溢れの 左上しか 見て いませんでした★
 *    （`docs/measured/osu-dai-dake.mjs` 自身が そう 断って います）
 *    ⇒★溢れた 先の マスが 正しいかを 誰も 実Excel と 突き合わせて いなかった★
 *    ⇒ 経営者1 が 全マスを 聞いた（㊳）
 *      `docs/measured/golden-kobore-zenmasu-2026-09-19.tsv`
 *    ⇒★本番で 4本 外れて いる★のが 出た（MAP / SCAN / MAKEARRAY / FILTER）
 *
 *  ★★既に 在る 門との 違い（★作る前に 読みました★）★★
 *    `kobore.test.mjs` ......... 足した 関数が 2つ以上 返せるか（★仕組み★）
 *    `afureru.test.mjs` ........ 1つの 式が 何マスにも 広がるか（★仕組み★）
 *    `shiki-hyou-afure.test.mjs` 本体が 溢れを 置けるか（★仕組み★）
 *    ⇒★どれも 「★実Excel と 同じ 中身か★」は 見て いません★
 *    ⇒★この 門は ★紙と 字で 突き合わせます★★
 *
 *  ★★この 門が 言えない 事★★
 *    ・★台（lib/shiki-hyou.js）だけ★です＝★画面では ありません★
 *      ⇒画面は 経営者1 の `osu-kobore-zenmasu.mjs`（★配信に 当てる★）が 見ます
 *    ・★材料は 1組だけ★（A1:A3 = 3/1/2 ／ B1:B3 = 10/20/30）
 *      ⇒★横に 出る 形・四角に 出る 形・溢れ先が 塞がって いる 形は まだ★
 *      ⇒経営者1 が 物差しを 広げて います（2026-09-19）
 *
 *  使い方: node tests/kobore-zenmasu.test.mjs
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(import.meta.url);
let pass = 0, fail = 0;
const T = (n, よい, 添え) => {
  if (よい) { pass++; console.log('  ok   ' + n); }
  else { fail++; console.log('  NG   ' + n + (添え ? '\n       ' + 添え : '')); }
};

console.log('');
console.log('[kobore-zenmasu] 溢れた先の 全マスが 実Excel と 同じか');

const 紙の道 = path.join(ROOT, 'docs/measured/golden-kobore-zenmasu-2026-09-19.tsv');
/* ★★2026-09-19 ... ★経営者1 が 物差しを 広げました（14本 -> 34本）★★★
     `golden-kobore-hirogeru-2026-09-19.tsv` ... ★20本★
       横 1x4 ／ 縦 4x1 ／ 四角 3x3 ／ 天地 3x2 ／ ★字を 返す MAP★
       ★FILTER（当てはまる 物 なし ⇒ #CALC!）★
       ★★SEQUENCE(2,2) の 溢れ先を 塞いだ 時（#SPILL!）★★ ＋ 塞がない 時（対照）
     ⇒★「直った」と 言える 分母を 先に 作って もらいました★ */
const 広げた道 = path.join(ROOT, 'docs/measured/golden-kobore-hirogeru-2026-09-19.tsv');
T('★紙が 在る★', fs.existsSync(紙の道), 紙の道);
T('★広げた 紙も 在る★', fs.existsSync(広げた道), 広げた道);

const NL = String.fromCharCode(10), TAB = String.fromCharCode(9);
/* ★★実Excel の 誤りは 「数」で 返ります★★（★字と 比べると 合って いる 物を 外れに します★）
     ＝経営者1 も 私も 同じ日に ここで 偽の 外れを 出しました */
const 誤りの数 = {
  '-2146826238': '#CALC!', '-2146826243': '#SPILL!', '-2146826273': '#VALUE!',
  '-2146826281': '#DIV/0!', '-2146826246': '#N/A', '-2146826259': '#NAME?',
  '-2146826252': '#NUM!', '-2146826265': '#REF!', '-2146826288': '#NULL!',
};
const 正を直す = (t) => String(t).split('|')
  .map((x) => (誤りの数[x] !== undefined ? 誤りの数[x] : x)).join('|');

function 紙を読む(道, 邪魔列) {
  return fs.readFileSync(道, 'utf-8').split(NL)
    .filter((l) => l && l.charAt(0) !== '#')
    .map((l) => l.split(TAB))
    .filter((p) => p[1] && p[1].charAt(0) === '=')
    .map((p) => ({ 式: p[1], 行数: Number(p[2]), 列数: Number(p[3]), 並び: 正を直す(p[4]),
      邪魔: (邪魔列 !== null && p[邪魔列]) ? p[邪魔列].trim() : '' }));
}
const 行たち = 紙を読む(紙の道, null).concat(紙を読む(広げた道, 7));

T('★★紙の 分母が 34本★★（★形が 変わったら 赤★）', 行たち.length === 34, '今 ' + 行たち.length + '本');

const SH = require_(path.join(ROOT, 'lib/shiki-hyou.js'));
const 板 = SH.表();
/* ★材料は 紙の 頭に 書いて ある 物★ */
板.打つ('A1', 3); 板.打つ('A2', 1); 板.打つ('A3', 2);
板.打つ('B1', 10); 板.打つ('B2', 20); 板.打つ('B3', 30);

const 字に = (v) => (v && typeof v === 'object')
  ? String(v.値 !== undefined ? v.値 : (v.error !== undefined ? v.error : v))
  : String(v);

/* ══ ★★本番と 同じ 3段で 押します★★ ══（2026-09-19）
     ①JS層（`_jsComputeFormula`） → ②台（`lib/shiki-hyou.js`） → ③借り物
     ★★ここを 台だけに したら 門が 死にました★★
       ＝わざと `_jsSet` に SCAN / MAP / MAKEARRAY を 戻したのに ★緑の まま★
       ＝★「JS層が 先に 答えると 潰れます」と 注に 書きながら JS層を 通して いなかった★
       ⇒記憶「見張りは ★見る 範囲★を 先に 数えて 書く」
       ⇒★★今日 見張り 2本で 踏んだ 穴を 自分の 新しい 門で また 踏みました★★ */
const EF = require_(path.join(ROOT, 'exally-formula.js'));
/* ★★JS層は 借り物を 繋がないと ★何も 答えません★★★（実測・2026-09-19）
     `_jsComputeFormula` の 1行目 ... `if(!v||v[0]!=='='||!_hf) return null;`
     ⇒★`initExallyFormula(hf)` を 呼ぶまで ★`=SUM(1,2)` すら null★★
     ⇒★★繋がずに 測ると ★JS層が 居ない 世界★を 測る 事に なります★★
       ＝私は これで 一度 ★門が 何も 捕まえない★のを 見ました
       ⇒記憶「測る 道具が 返した 0 を 根拠に するな」 */
const HFns = require_(path.join(ROOT, 'hyperformula.full.min.js'));
EF.registerExallyFunctions(HFns);
const hf = HFns.HyperFormula.buildEmpty({ licenseKey: 'gpl-v3' });
/* ★★紙（シート）を 作って 材料を 置きます★★（2026-09-19・★これが 無くて 一度 門が 死にました★）
     ＝紙が 無いと JS層は ★"There's no sheet with id = 0" を 投げます★
     ＝私の 門は その 投げを `null` として 飲み込み ★何を 壊しても 緑★でした
     ⇒記憶「★口を 確かめずに 書いた コードは 静かに 死ぬ★」（try で 包むと 緑に なる）
     ⇒★だから 下に 「JS層が 本当に 動いて いる」門を 置いて います★ */
const SID = hf.getSheetId(hf.addSheet('Sheet1'));
hf.setSheetContent(SID, [[3, 10], [1, 20], [2, 30]]);
EF.initExallyFormula(hf);
const JS層 = EF._jsComputeFormula || (EF.__test && EF.__test._jsComputeFormula) || null;
T('★★本番の 1段目（JS層）に 手が 届く★★（★届かないと 1段目を 飛ばして 測ります★）',
  typeof JS層 === 'function');
/* ★★JS層が 本当に 動いて いるか★★（★投げも 数えます★）
     ★前は `!== undefined` で 見て いました★＝★投げても 緑に なる 書き方★
     ⇒★投げたら 赤★に します（★飲み込まない★） */
{
  let 生きて = false, 訳 = '';
  try {
    const r = JS層(0, '=DSUM(A1:B2,1,A1:A2)');
    生きて = (r !== null && r !== undefined);
    if (!生きて) 訳 = 'null を 返した＝繋がって いない か 対象外';
  } catch (e) { 訳 = '★投げた★ ... ' + (e && e.message); }
  T('★★JS層が 本当に 動いて いる★★（★繋がない／紙が 無いと 何を 押しても 偽の 緑★）',
    生きて, 訳);
}

function 押す(式, 邪魔, 列数) {
  /* ★材料を 毎回 置き直します★（★前の 溢れが 残ると 偽の 外れ★） */
  板.打つ('A1', 3); 板.打つ('A2', 1); 板.打つ('A3', 2);
  板.打つ('B1', 10); 板.打つ('B2', 20); 板.打つ('B3', 30);
  /* ★★邪魔が 在る 時は ★その 隣★に 置きます★★（離すと 塞がらない＝偽の 外れ） */
  if (邪魔) {
    板.打つ(邪魔, 'jama');
    板.打つ('D1', 式);
    const 横 = [];
    for (let k = 0; k < 列数; k++) 横.push(String(板.字(String.fromCharCode(68 + k) + '1')));
    return { 行数: 1, 列数: 列数, 並び: 横.concat(['', '', '', '']).slice(0, 5).join('|') };
  }
  /* ★①JS層★ ... 答えたら ★1つの 値★＝★溢れません★ */
  if (typeof JS層 === 'function') {
    let j = null;
    /* ★投げは 飲み込みません★＝★飲み込むと 門が 静かに 死にます★（2026-09-19 実測） */
    j = JS層(0, 式);
    if (j !== null && j !== undefined) return { 行数: 1, 列数: 1, 並び: String(j) + '||||' };
  }
  /* ★②台★ */
  板.打つ('ZZ9983', 式);
  const v = 板.値('ZZ9983');
  if (v && v.溢れ === true) {
    return {
      行数: v.並び.length,
      列数: (v.並び[0] || []).length,
      並び: v.並び.map((r) => r.map(字に).concat(['', '', '', '']).slice(0, 5).join('|')).join(' / '),
    };
  }
  return { 行数: 1, 列数: 1, 並び: 字に(v) + '||||' };
}

const 外れ = [];
for (const r of 行たち) {
  const 出 = 押す(r.式, r.邪魔, r.列数);
  if (出.行数 !== r.行数 || 出.列数 !== r.列数 || 出.並び !== r.並び) {
    外れ.push(r.式 + '\n         実Excel ' + r.行数 + 'x' + r.列数 + ' ... ' + r.並び
      + '\n         台      ' + 出.行数 + 'x' + 出.列数 + ' ... ' + 出.並び);
  }
}
T('★★34本 とも 溢れた先まで 実Excel と 同じ★★（★減っても 増えても 赤★）',
  外れ.length === 0, 外れ.join('\n       '));

/* ★★溢れる 形が 本当に 溢れて いるか★★（★1x1 に 潰れて いないか★） */
{
  const 溢れる頭 = ['=MAP(', '=SCAN(', '=MAKEARRAY(', '=BYROW(', '=BYCOL(',
    '=SORT(', '=UNIQUE(', '=FILTER(', '=SEQUENCE(', '=TRANSPOSE(', '=SORTBY(', '=TEXTSPLIT('];
  const 潰れた = [];
  for (const r of 行たち) {
    if (!溢れる頭.some((h) => r.式.indexOf(h) === 0)) continue;
    if (r.行数 === 1 && r.列数 === 1) continue;        /* ★紙の 方が 1x1 なら 潰れて いない★ */
    const 出 = 押す(r.式, r.邪魔, r.列数);
    if (出.行数 === 1 && 出.列数 === 1) 潰れた.push(r.式);
  }
  T('★★溢れる頭の 式が 1x1 に 潰れて いない★★（★JS層が 先に 答えると 潰れます★）',
    潰れた.length === 0, 潰れた.join(' / '));
}

/* ★★わざと 1つ 変えたら 赤に なるか★★（★門が 本当に 見て いる 証し★） */
{
  let 見つけた = 0;
  for (const r of 行たち) {
    const にせ = (r.式.indexOf('=MAP(') === 0) ? { ...r, 並び: 'X||||' } : r;
    const 出 = 押す(にせ.式, にせ.邪魔, にせ.列数);
    if (出.並び !== にせ.並び) 見つけた++;
  }
  const MAPの数 = 行たち.filter((r) => r.式.indexOf('=MAP(') === 0).length;
  T('★★紙の 字を 変えたら 赤に なる★★（★門が 見て いる 証し★）',
    見つけた === MAPの数 && MAPの数 > 0, '見つけた ' + 見つけた + ' ／ MAP の 行 ' + MAPの数);
}

console.log('      ... 押した ' + 行たち.length + '本 ／ 材料 A1:A3 = 3/1/2 ／ B1:B3 = 10/20/30');
console.log('');
console.log('kobore-zenmasu: ' + pass + ' 緑 / ' + fail + ' 赤');
process.exit(fail ? 1 : 0);
