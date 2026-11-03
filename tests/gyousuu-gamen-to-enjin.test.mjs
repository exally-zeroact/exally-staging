/* gyousuu-gamen-to-enjin.test.mjs — ★画面が 出す 大きさ と 計算エンジンが 受ける 大きさ を 揃える★
 *
 *  ★★何が 起きて いたか（2026-09-08・★本番で 実測★）★★
 *    book.html:9006 … 画面 ★ROWS = 1048576, COLS = 16384★（実Excel と 同じ）
 *    book.html:1726 … エンジン ★buildEmpty に maxRows を 渡して いなかった★＝既定 40,000
 *    ⇒★★画面が 出して いる 100万行の うち ★96%（1,008,576行）★が 計算されて いなかった★★
 *
 *    ★お客さんに 出て いた 字★（本番の book.html で お客さんと 同じ 関数を 呼んだ）
 *      A1 に 10 を 置いて 各行に `=A1*2`（答えは 20）
 *        行    101 … 打った直後 20      ／ 再計算の後 20
 *        行 40,000 … 打った直後 20      ／ 再計算の後 20
 *        行 40,002 … 打った直後 =A1*2   ／ ★再計算の後 0★
 *        行 50,001 … 打った直後 =A1*2   ／ ★再計算の後 0★
 *    ⇒★★誤りでは なく 0＝『静かに 違う 答え』★★
 *      （今日 REPLACEB でも 同じ 形を 見つけた。★断られる より 気づけない★）
 *
 *  ★★この 見張りが 守る 物（1つずつ）★★
 *    ①★画面の ROWS / COLS と エンジンの maxRows / maxColumns が ★同じ 数★★
 *      ⇒ 食い違いこそが 病気の 正体なので ★数を 突き合わせる★
 *    ②★その 数が ★実Excel の 大きさ★（1,048,576行 × 16,384列）★
 *      ⇒ 揃って いても 両方 小さかったら 意味が 無い
 *    ③★★お客さんの 道で 押して 4万行目より 下が 0 に ならない★★
 *      ⇒ 数だけ 見て 緑に しない（今日の 決まり＝『印が 付いたか』では なく『答えが 出たか』）
 *    ④★一番 下の 行・一番 右の 列でも 計算できる★
 *
 *  ★なぜ 数を grep で 取るか★
 *    book.html は 1MB で node から そのまま 動かせない。
 *    ⇒ ★書いて ある 数を 読んで 突き合わせる★
 *    ⇒ ただし ③④は ★本当に エンジンを 作って 押す★（字だけで 緑に しない）
 *
 *  使い方: node tests/gyousuu-gamen-to-enjin.test.mjs [--self-test]
 */
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath, pathToFileURL } from 'node:url';

const ここ = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(ここ, '..');
const require_ = createRequire(import.meta.url);
const 直に走った = process.argv[1] && pathToFileURL(process.argv[1]).href === import.meta.url;
const 自己試験 = 直に走った && process.argv.includes('--self-test');

let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ✓ ' + n); } catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + (e && e.message)); } };

/* ★実Excel の 大きさ★（2026-09-08 実測。golden-ookisa-2026-09-08.tsv に 在る）
   =SEQUENCE(1048576) … 通る ／ =SEQUENCE(1048577) … #VALUE!  */
const 実Excelの行 = 1048576;
const 実Excelの列 = 16384;

const 本文 = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf-8');

/* ★画面の 数★ … `var ROW_H = 22, …, ROWS = 1048576, COLS = 16384;` */
function 画面の大きさ(s) {
  const r = s.match(/\bROWS\s*=\s*(\d+)/);
  const c = s.match(/\bCOLS\s*=\s*(\d+)/);
  return { 行: r ? Number(r[1]) : null, 列: c ? Number(c[1]) : null };
}
/* ★エンジンに 渡して いる 数★ … `HyperFormula.buildEmpty({… maxRows:… maxColumns:…})` */
function エンジンの大きさ(s) {
  const m = s.match(/HyperFormula\.buildEmpty\(\{[\s\S]{0,400}?\}\)/);
  if (!m) return { 行: null, 列: null, 見つかった: false };
  const r = m[0].match(/maxRows\s*:\s*(\d+)/);
  const c = m[0].match(/maxColumns\s*:\s*(\d+)/);
  return { 行: r ? Number(r[1]) : null, 列: c ? Number(c[1]) : null, 見つかった: true };
}

console.log('\n★画面の 大きさ と 計算エンジンの 大きさ★');

const 画面 = 画面の大きさ(本文);
const エンジン = エンジンの大きさ(本文);

T('★書いて ある 数を 両方 読めた（読めなければ 止める）★', () => {
  if (画面.行 === null || 画面.列 === null) throw new Error('★画面の ROWS / COLS が 読めない★＝この 試験は 何も 見て いない');
  if (!エンジン.見つかった) throw new Error('★buildEmpty が 見つからない★＝この 試験は 何も 見て いない');
  console.log('      画面 … ' + 画面.行.toLocaleString() + '行 × ' + 画面.列.toLocaleString() + '列');
  console.log('      エンジン … ' + (エンジン.行 === null ? '★渡して いない★' : エンジン.行.toLocaleString() + '行')
    + ' × ' + (エンジン.列 === null ? '★渡して いない★' : エンジン.列.toLocaleString() + '列'));
});

T('★★エンジンに 大きさを 渡して いる（渡さないと 既定の 40,000行）★★', () => {
  if (エンジン.行 === null) {
    throw new Error('★maxRows を 渡して いない★＝既定の 40,000行に なり、'
      + '画面が 出す ' + 画面.行.toLocaleString() + '行の うち 96%が 計算されない');
  }
  if (エンジン.列 === null) throw new Error('★maxColumns を 渡して いない★');
});

T('★★行は 画面と エンジンで ★同じ 数★（食い違いが 今回の 病気）★★', () => {
  if (エンジン.行 !== 画面.行) {
    throw new Error('★行が 食い違う★ 画面 ' + 画面.行 + ' ／ エンジン ' + エンジン.行);
  }
});

/* ★★列は『同じ』では なく『画面以上』★★（★一度 揃えて 壊した★）
   私は 最初 maxColumns も 画面の 16,384 に した ⇒★1本 悪化した★
     `=SEQUENCE(1,16385)` … 実Excel ★通る★ ／ 揃えると ★#VALUE!★
   ★実Excel に 5通りの 物差しで 聞き直した（溢れさせずに）★
     列 16,385 … INDEX=1 ／ ★COLUMNS=16385★ ／ ISERROR=FALSE ／ SUM 通る
     行 1,048,577 … ★#VALUE!★
   ⇒★実Excel は ★行だけ 頭打ち★。列は 表の 中で 超えて よい（画面に こぼせないだけ）★
   ⇒★★『画面より エンジンが 狭い』が 病気。『広い』は 病気では ない★★ */
T('★★列は エンジンが 画面以上（狭いのが 病気・広いのは 病気で ない）★★', () => {
  if (エンジン.列 < 画面.列) {
    throw new Error('★エンジンの 列が 画面より 狭い★ 画面 ' + 画面.列 + ' ／ エンジン ' + エンジン.列
      + '＝画面が 出して いる 列が 計算されない');
  }
});

T('★その 数が 実Excel の 大きさ（揃って いても 小さければ 意味が 無い）★', () => {
  if (画面.行 !== 実Excelの行) throw new Error('★行が 実Excel と 違う★ ' + 画面.行 + '（実Excel ' + 実Excelの行 + '）');
  if (画面.列 !== 実Excelの列) throw new Error('★列が 実Excel と 違う★ ' + 画面.列 + '（実Excel ' + 実Excelの列 + '）');
});

/* ══ ★ここから 本当に 押す★（字だけで 緑に しない） ══════════════ */

globalThis.HF_ERR = {
  DIV_BY_ZERO: '#DIV/0!', NUM: '#NUM!', NA: '#N/A', VALUE: '#VALUE!',
  REF: '#REF!', NAME: '#NAME?', CYCLE: '#CYCLE!', NULL: '#NULL!',
  SPILL: '#SPILL!', GETTING_DATA: '#GETTING_DATA',
};
const HFns = require_(path.join(ROOT, 'hyperformula.full.min.js'));
const EF = require_(path.join(ROOT, 'exally-formula.js'));
EF.registerExallyFunctions(HFns);
const HF0 = HFns.HyperFormula;
const H = Object.assign(Object.create(HF0), HFns,
  { registerFunctionPlugin: HF0.registerFunctionPlugin.bind(HF0) });
for (const n of ['cell', 'complex', 'extra', 'filterxml', 'kane', 'nokori', 'soto', 'yosoku']) {
  require_(path.join(ROOT, 'lib/formula-' + n + '-plug.js'))
    .つなぐ(H, require_(path.join(ROOT, 'lib/formula-' + n + '.js')));
}

/* ★book.html が 渡して いる 物を そのまま 使う★（試験だけ 別の 設定にしない） */
function エンジンを作る(行, 列) {
  const hf = HF0.buildEmpty({
    licenseKey: 'gpl-v3', useArrayArithmetic: true, smartRounding: false,
    maxRows: 行, maxColumns: 列,
  });
  const SID = hf.getSheetId(hf.addSheet('S'));
  EF.initExallyFormula(hf);
  return { hf, SID };
}

/* ★★出口の 門★★＝誤りが 字に ならない 道具は 走らせない */
{
  const { hf, SID } = エンジンを作る(画面.行, 画面.列);
  hf.setCellContents({ sheet: SID, row: 0, col: 0 }, '=1/0');
  if (String(EF._hfGetDisplay(0, 0, 0, true)) !== '#DIV/0!') {
    console.error('★出口が 効いていない＝止める★'); process.exit(3);
  }
}

/* ★★お客さんの 道★★＝A1 に 10 を 置いて 各行に `=A1*2` を 打つ（答えは 20） */
function 打って読む(行数, 列数, 行) {
  const { hf, SID } = エンジンを作る(行数, 列数);
  hf.setCellContents({ sheet: SID, row: 0, col: 0 }, 10);
  try {
    hf.setCellContents({ sheet: SID, row: 行, col: 1 }, EF.convertFormula('=A1*2'));
  } catch (e) { /* ★本番も ここで 握りつぶして いた★＝同じ 形で 続ける */ }
  /* ★book.html:9533 recalcSheet と 同じ★ cell.d = _hfGetDisplay(…) */
  return String(EF._hfGetDisplay(0, 行, 1, true));
}

T('★★4万行目より 下でも 式が 計算される（0 に ならない）★★', () => {
  const 見る行 = [100, 39999, 40000, 40001, 50000, 500000, 実Excelの行 - 1];
  const 外れ = [];
  for (const r of 見る行) {
    const 出 = 打って読む(エンジン.行, エンジン.列, r);
    if (出 !== '20') 外れ.push((r + 1).toLocaleString() + '行目 → ' + 出);
  }
  console.log('      押した 行 … ' + 見る行.length + '通り（101 / 40,000 / 40,001 / 50,001 / 500,001 / 1,048,576 行目 ほか）');
  if (外れ.length) throw new Error('★' + 外れ.length + '通りで 20 に ならない★\n      ' + 外れ.join('\n      '));
});

T('★★16,385列の 表が 作れる（★私が 一度 壊した 所★）★★', () => {
  /* ★実Excel は 表の 中でだけなら 16,384列を 超えて よい★（画面に こぼせないだけ）
       =COLUMNS(SEQUENCE(1,16385)) … ★16385★（実Excel 実測・5通りの 物差しで 確かめた）
     ⇒ maxColumns を 画面の 16,384 に 揃えると ここが #VALUE! に なる
     ⇒★私は 一度 そう しました。この 試験は ★その 戻りを 止める★為に 在ります★ */
  const { hf, SID } = エンジンを作る(エンジン.行, エンジン.列);
  hf.setCellContents({ sheet: SID, row: 0, col: 0 }, EF.convertFormula('=COLUMNS(SEQUENCE(1,16385))'));
  const 出 = String(EF._hfGetDisplay(0, 0, 0, true));
  if (出 !== '16385') {
    throw new Error('★=COLUMNS(SEQUENCE(1,16385)) が ' + 出 + '★（実Excel は 16385）'
      + '＝エンジンの 列を 画面に 揃えて しまって いないか');
  }
});

T('★一番 右の 列でも 計算される★', () => {
  const { hf, SID } = エンジンを作る(エンジン.行, エンジン.列);
  hf.setCellContents({ sheet: SID, row: 0, col: 0 }, 10);
  const 右 = エンジン.列 - 1;
  try { hf.setCellContents({ sheet: SID, row: 0, col: 右 }, EF.convertFormula('=A1*2')); }
  catch (e) { throw new Error('★' + (右 + 1) + '列目に 打てない★ ' + e.constructor.name); }
  const 出 = String(EF._hfGetDisplay(0, 0, 右, true));
  if (出 !== '20') throw new Error('★' + (右 + 1) + '列目 → ' + 出 + '（20 のはず）');
});

/* ══ ★自己試験＝この 見張りが 本当に 見て いるか★ ═══════════════ */
if (自己試験) {
  console.log('\n★自己試験（★壊すのは 写し★＝book.html は 1バイトも 触らない）★');

  T('★壊し方① エンジンに 大きさを 渡さない（★直す前の 本番★）★', () => {
    const 写し = 本文.replace(/maxRows\s*:\s*\d+\s*,\s*maxColumns\s*:\s*\d+/, '');
    const e = エンジンの大きさ(写し);
    if (e.行 !== null) throw new Error('★消えて いない★＝この 壊し方が 効いて いない');
    console.log('      … 「大きさを 渡して いる」の 試験が 赤に なる');
  });

  T('★壊し方② 画面と エンジンを 食い違わせる★', () => {
    const 写し = 本文.replace(/maxRows\s*:\s*\d+/, 'maxRows:40000');
    const e = エンジンの大きさ(写し);
    const g = 画面の大きさ(写し);
    if (e.行 === g.行) throw new Error('★食い違って いない★＝この 壊し方が 効いて いない');
    console.log('      … 画面 ' + g.行.toLocaleString() + ' ／ エンジン ' + e.行.toLocaleString() + ' で 赤');
  });

  T('★壊し方③ 両方 揃って いるが 小さい（揃えば よい では ない）★', () => {
    let 写し = 本文.replace(/maxRows\s*:\s*\d+/, 'maxRows:40000');
    写し = 写し.replace(/\bROWS\s*=\s*\d+/, 'ROWS = 40000');
    const e = エンジンの大きさ(写し), g = 画面の大きさ(写し);
    if (e.行 !== g.行) throw new Error('★揃って いない★＝この 壊し方が 効いて いない');
    if (g.行 === 実Excelの行) throw new Error('★小さく なって いない★');
    console.log('      … 揃って いても 実Excel（' + 実Excelの行.toLocaleString() + '行）で ないので 赤');
  });

  T('★★壊し方④ 数は 直っているが ★実際に 押すと 0★（一番 大事）★★', () => {
    /* ★数だけ 見て 緑に する 見張りだったら、ここを 見逃す★
       ⇒ 40,000行の エンジンで 50,001行目を 押して ★0 が 返る★事を 確かめる
       ⇒ ＝直す前の 本番で 起きて いた 事 そのもの */
    const 出 = 打って読む(40000, 18278, 50000);
    if (出 === '20') throw new Error('★0 に ならない★＝この 壊し方が 効いて いない');
    console.log('      … 40,000行の エンジンで 50,001行目 → [' + 出 + ']（★これが 直す前の 本番★）');
  });
}

console.log('\n' + pass + ' passed, ' + fail + ' failed');
if (直に走った) process.exit(fail ? 1 : 0);
