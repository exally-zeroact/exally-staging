/* hikisuu-mihari.test.mjs — ★引数の 見張り＝実Excel が 断る 所で うちも 断る★
 *
 *  ★★何を 守るか（1つずつ 名指し）★★
 *    ①★TOCOL / TOROW の 2つ目は『大きさ』では なく『無視する 印』＝0〜3 だけ★
 *      ⇒ 直す前は ★何を 渡しても 受けて いた★
 *        `=TOCOL(A1:B5,DATE(2024,1,3))` … 実Excel #VALUE! ／ うち ★1★
 *    ②★FINDB / SEARCHB の『何文字目から』は 1 以上★（0 も −1 も #VALUE!）
 *      ⇒ 直す前は `Math.max(1, …)` で ★黙って 1 に 直して いた★
 *    ③★REPLACEB の『何文字目から』は 1〜32767・『何文字ぶん』は 0〜32767★
 *      ⇒ 32,767 ＝ 2の15乗−1 ＝ 実Excel の 1マスの 字数（★丸い 数＝仕様★）
 *    ④★★小数は 切り捨て★★（一番 危ない）
 *      ⇒ `=REPLACEB("abcde",1.5,1,"x")` … 実Excel ★xbcde★ ／ 直す前 ★axcde★
 *      ⇒★誤りでは なく ★静かに 違う 答え★＝画面には それらしい 字が 出る★
 *    ⑤★WRAPROWS / WRAPCOLS の 幅が 1未満は ★#NUM!★★（#VALUE! では ない）
 *    ⑥★下限を 見る 所は 関数ごとに 違う★
 *      ⇒ TOCOL は −0.5 を ★受ける★（切り捨てて 0）
 *      ⇒ REPLACEB の 長さは −0.5 を ★断る★（切り捨てる 前に 見る）
 *      ⇒★★1か所で まとめて 判定 できない＝まとめて 直すなの 実例★★
 *
 *  ★★どこから 断るかは ★私が 決めていません★★★
 *    紙 … docs/measured/kansuu46/golden-hikisuu-mihari-2026-09-08.tsv（★180本★）
 *    取り方 … docs/measured/kansuu46/toru-hikisuu-mihari.ps1
 *      ＝実Excel（16.0 build 20326）に ★12個の 引数 × 15通りの 数★ を 打たせた
 *      ＝★境目の 前後を 必ず 両方★（-2 -1 -0.5 0 0.5 1 1.5 2 2.5 3 3.5 4 5 6 45294）
 *    ⇒★私が 手で 選ぶと 自分が 通る 形しか 選ばない★
 *
 *  ★★物差しの 話（この 紙は 1度 作り直しました）★★
 *    表を 返す 式を そのまま COM で 打つと ★前の 溢れが 残って #SPILL! に 化ける★。
 *    ⇒ `=INDEX(式,1,1)` で ★溢れさせずに★ 左上だけ 取る 形に 直した。
 *    ⇒ 直す前の 紙は `=DROP(A1:A5,2)` を #VALUE! と 書いて いた（★嘘★）。
 *
 *  ★★まだ 出せない 6本（★隠さずに 名指しで 数える★）★★
 *    実Excel が ★#CALC!★ を 返す 6本は ★今は 出せません★。
 *      =TAKE(A1:A5,-0.5) =TAKE(A1:A5,0) =TAKE(A1:A5,0.5)
 *      =DROP(A1:A5,5)    =DROP(A1:A5,6)  =DROP(A1:A5,45294)
 *    訳 … HyperFormula の 誤りの 名前は 9個で 固定（CYCLE DIV_BY_ZERO ERROR NA
 *          NAME NUM REF SPILL VALUE）で ★CALC が 無い★。
 *          `new CellError('CALC')` は「errors.CALC の 訳が 無い」で 止まる。
 *          出すには ★言語ごと 登録し直す★ 要が 在り、★別件★に します。
 *    ⇒★★『まだ』を 緑に しない＝この 6本は ここで 数えて 残す★★
 *
 *  使い方: node tests/hikisuu-mihari.test.mjs [--self-test]
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

/* ★`_hfGetDisplay` は book.html の この 表を 見る★（1文字も 変えずに 写した）
   ⇒ 渡さないと `#NAME?` が `#NAME` に なる（2026-09-08 に 踏んだ） */
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
for (const n of ['extra', 'nokori', 'kane']) {
  require_(path.join(ROOT, 'lib/formula-' + n + '-plug.js'))
    .つなぐ(H, require_(path.join(ROOT, 'lib/formula-' + n + '.js')));
}
/* ★★エンジンは ★1本ごとに 作り直す★★★
   ⇒ 1つの エンジンで 180本 押したら ★4GB を 使い切って 道具ごと 死にました★
     （FATAL ERROR: Reached heap limit）
   ⇒ ★1本ずつ 隔離して 押し直したら 1本も 落ちません★（どれも 300ミリ秒 以下）
     ＝★1本が 重いのでは なく ★溜まる★のが 原因★
   ⇒ `setSheetContent` で 前の 表を 置き換えても HyperFormula は 手放しません。
     `=WRAPCOLS(A1:A5,45294)` の ような 4万マスが 180本ぶん 残る。
   ⇒★★これは ★大きさの 上限★の 話と 同じ 根っこ（別の PR）★★
   ⇒ ここでは ★試験を 落とさない★ 為に 作り直します（1本 約10ミリ秒） */
function エンジン() {
  const hf = HF0.buildEmpty({ licenseKey: 'gpl-v3', useArrayArithmetic: true, smartRounding: false });
  const SID = hf.getSheetId(hf.addSheet('Sheet1'));
  EF.initExallyFormula(hf);
  return { hf, SID };
}

/* ★★出口の 門★★＝この 道具が 誤りを 字に 出来ているか 先に 確かめる
   ⇒ 2026-09-08 に ★HF_ERR を 渡し忘れて 全部の 誤りが #ERR に なった 紙を commit した★ */
{
  const { hf, SID } = エンジン();
  hf.setSheetContent(SID, [['=1/0'], ['=NA()']]);
  const a = String(EF._hfGetDisplay(0, 0, 0, true));
  const b = String(EF._hfGetDisplay(0, 1, 0, true));
  if (a !== '#DIV/0!' || b !== '#N/A') {
    console.error('★出口が 効いていない（=1/0 → ' + a + ' ／ =NA() → ' + b + '）＝止める★');
    process.exit(3);
  }
}

/* ★実Excel を 測った 時と 同じ 土台★ A1:A5=1..5 ／ B1:B5=2,4,6,8,10 */
function 土台() {
  const 表 = [];
  for (let r = 0; r < 5; r++) 表.push([r + 1, (r + 1) * 2, null, null]);
  return 表;
}

/* ★★本番と 同じ 道で 押す★★（入口＝JS層／出口＝_hfGetDisplay）
   book.html … cell.d = jsResult!==null ? jsResult : _hfGetDisplay(sheetIdx, r, c, true) */
function 押す(式) {
  try {
    const js = EF._jsComputeFormula(0, 式);
    if (js !== null && js !== undefined) return String(js);
  } catch (e) { /* engine へ */ }
  let 後;
  try { 後 = EF.convertFormula(式); } catch (e) { return '★書き換えで 例外★'; }
  try {
    const { hf, SID } = エンジン();
    const 表 = 土台();
    表[0][3] = 後;
    hf.setSheetContent(SID, 表);
    return String(EF._hfGetDisplay(0, 0, 3, true));
  } catch (e) { return '★engine で 例外★'; }
}

/* ★PowerShell が 書いた 紙は 行末が CRLF＝最後の 列に CR が 残る★ */
const 紙 = fs.readFileSync(
  path.join(ROOT, 'docs/measured/kansuu46/golden-hikisuu-mihari-2026-09-08.tsv'), 'utf-8')
  .split(/\r?\n/).filter((l) => l && !l.startsWith('#'))
  .map((l) => l.split('\t').map((x) => x.replace(/\r$/, '')))
  .filter((p) => p.length >= 6)
  .map((p) => ({ 名: p[0], 式: p[1], 答: p[2], 型: p[3], 数: p[4], 引数: p[5] }));

/* ★★まだ 出せない 6本＝★式で 名指し★（「#CALC! なら 見逃す」に しない）★★
   ⇒ 曖昧に すると ★これから 増える #CALC! も 黙って 見逃す★ */
const まだ出せない = new Set([
  '=TAKE(A1:A5,-0.5)', '=TAKE(A1:A5,0)', '=TAKE(A1:A5,0.5)',
  '=DROP(A1:A5,5)', '=DROP(A1:A5,6)', '=DROP(A1:A5,45294)',
]);

const 同じか = (出, 正, 型) => {
  const s = String(出);
  if (型 === 'Double' || 型 === 'Int32') {
    const a = Number(出), b = Number(正);
    if (!isFinite(a) || !isFinite(b)) return s === 正;
    return a === b || Math.abs(a - b) <= Math.max(1, Math.abs(b)) * 1e-9;
  }
  return s === 正;
};

console.log('\n★引数の 見張り（実Excel に 180本 打って 決めた）★');

T('★★紙が 減っていない … 180本／12個の 引数／15通りの 数★★', () => {
  if (紙.length !== 180) throw new Error('★紙が ' + 紙.length + '本★（180本 在るはず）');
  const 引数 = new Set(紙.map((g) => g.引数));
  if (引数.size !== 12) throw new Error('★引数が ' + 引数.size + '個★（12個 在るはず）');
  const 数 = new Set(紙.map((g) => g.数));
  if (数.size !== 15) throw new Error('★数が ' + 数.size + '通り★（15通り 在るはず）');
  for (const n of ['-2', '-1', '-0.5', '0', '0.5', '1.5', '2.5', '3.5', '45294']) {
    if (!数.has(n)) throw new Error('★境目の 前後 ' + n + ' が 紙に 無い★');
  }
  console.log('      ' + 紙.length + '本 ／ 引数 ' + 引数.size + '個 ／ 数 ' + 数.size + '通り');
});

T('★★実Excel が 断る 所で うちも 断る（まだ 出せない 6本を 除く 174本）★★', () => {
  const 外れ = [];
  let 数えた = 0;
  for (const g of 紙) {
    if (まだ出せない.has(g.式)) continue;
    数えた++;
    const 出 = 押す(g.式);
    if (!同じか(出, g.答, g.型)) 外れ.push(g.式 + '   実Excel ' + g.答 + ' ／ 出 ' + 出);
  }
  if (数えた !== 174) throw new Error('★数えたのが ' + 数えた + '本★（174本 のはず）');
  console.log('      押した … ' + 数えた + '本');
  if (外れ.length) throw new Error('★' + 外れ.length + '本 違う★\n      ' + 外れ.join('\n      '));
});

T('★★まだ 出せない 6本は ★今も 出せない★（黙って 増えていない）★★', () => {
  /* ★『まだ』を 緑に しない＝出せる ように なったら ★この 試験が 赤に なって 教える★ */
  const 出せた = [];
  for (const 式 of まだ出せない) {
    const g = 紙.find((x) => x.式 === 式);
    if (!g) throw new Error('★紙に 無い 式を 見逃し表に 書いている★ … ' + 式);
    if (g.答 !== '#CALC!') throw new Error('★実Excel の 答えが #CALC! で ない★ … ' + 式 + ' → ' + g.答);
    if (同じか(押す(式), g.答, g.型)) 出せた.push(式);
  }
  console.log('      #CALC! の 6本 … まだ 出せない（HyperFormula に CALC が 無い）');
  if (出せた.length) {
    throw new Error('★' + 出せた.length + '本 出せる ように なっている★＝見逃し表から 外して ください\n      '
      + 出せた.join('\n      '));
  }
});

T('★★一番 危ないのは 誤りでは なく『静かに 違う 答え』★★', () => {
  /* ★直す前は 切り捨てず、2.5 を 3文字目に していた＝画面には それらしい 字が 出ていた★ */
  const 組 = [
    ['=REPLACEB("abcde",1.5,1,"x")', 'xbcde'],
    ['=REPLACEB("abcde",2.5,1,"x")', 'axcde'],
    ['=REPLACEB("abcde",1,0.5,"x")', 'xabcde'],
    ['=FINDB("b","abc",2.5)', '2'],
    ['=SEARCHB("b","abc",2.5)', '2'],
  ];
  for (const [式, 正] of 組) {
    const 出 = 押す(式);
    if (出 !== 正) throw new Error('★' + 式 + '★ 実Excel ' + 正 + ' ／ 出 ' + 出);
  }
  console.log('      小数の 切り捨て 5通り … 全部 実Excel と 同じ 字');
});

T('★★下限を 見る 所は 関数ごとに 違う（1か所で まとめて 判定 できない）★★', () => {
  /* ★TOCOL は −0.5 を 受ける（切り捨てて 0）／REPLACEB の 長さは −0.5 を 断る★
     ⇒ これを 1本に まとめると ★どちらかが 必ず 実Excel と 違う★ */
  const a = 押す('=TOCOL(A1:B5,-0.5)');
  if (a === '#VALUE!') throw new Error('★TOCOL(-0.5) を 断って いる★（実Excel は 受ける）');
  const b = 押す('=REPLACEB("abcde",1,-0.5,"x")');
  if (b !== '#VALUE!') throw new Error('★REPLACEB の 長さ −0.5 を 受けて いる★（実Excel は 断る） … ' + b);
  console.log('      TOCOL(-0.5) … ' + a + ' ／ REPLACEB(…,-0.5,…) … ' + b);
});

/* ══ ★自己試験＝この 見張りが 本当に 見ているか★ ═══════════════════
   ★『わざと 壊したら 赤が 1個 出た』で 信用しない＝★壊し方ごとに 何本 赤に なるか★を 数える★ */
if (自己試験) {
  console.log('\n★自己試験（★壊すのは 写し★＝repo の ファイルは 1バイトも 触らない）★');

  const 数える = (前の作り) => {
    let 赤 = 0;
    for (const g of 紙) {
      if (まだ出せない.has(g.式)) continue;
      if (!同じか(前の作り(g), g.答, g.型)) 赤++;
    }
    return 赤;
  };

  T('★壊し方① TOCOL/TOROW の 印を 見張らない（直す前の 作り）★', () => {
    /* 直す前 … 何を 渡しても 受けて いた ⇒ 0〜3 の 外は 全部 違う */
    const 赤 = 数える((g) => {
      if (!/^=(TOCOL|TOROW)\(/.test(g.式)) return g.答;   /* ここ以外は 触らない */
      const n = Math.trunc(Number(g.数));
      return (n >= 0 && n <= 3) ? g.答 : '1';              /* 断らずに 答えて しまう */
    });
    console.log('      … ★' + 赤 + '本★ 赤に なる');
    if (赤 !== 12) throw new Error('★' + 赤 + '本★＝実測は 12本（★この 試験は 何も 見ていない★）');
  });

  T('★壊し方② 小数を 切り捨てず 繰り上げる（静かに 違う 答え）★', () => {
    const 赤 = 数える((g) => {
      if (!/^=(FINDB|SEARCHB|REPLACEB)\(/.test(g.式)) return g.答;
      return Number(g.数) % 1 !== 0 ? '★ちがう字★' : g.答;
    });
    /* ★数え方を 書く★＝小数は −0.5 0.5 1.5 2.5 3.5 の ★5通り★
       × 引数は FINDB／SEARCHB／REPLACEB の 2番目／REPLACEB の 3番目 の ★4個★ ＝ 20本
       （最初 私は「4通り × 4引数 ＝ 16本」と 書いて 外しました＝★手で 並べたら その場で 数える★） */
    console.log('      … ★' + 赤 + '本★ 赤に なる');
    if (赤 !== 20) throw new Error('★' + 赤 + '本★＝実測は 20本（小数 5通り × 4引数）');
  });

  T('★壊し方③ WRAPROWS/WRAPCOLS を #VALUE! に 戻す（字だけ 違う）★', () => {
    const 赤 = 数える((g) => {
      if (!/^=WRAP(ROWS|COLS)\(/.test(g.式)) return g.答;
      return g.答 === '#NUM!' ? '#VALUE!' : g.答;
    });
    console.log('      … ★' + 赤 + '本★ 赤に なる');
    if (赤 !== 10) throw new Error('★' + 赤 + '本★＝実測は 10本（5通り × 2関数）');
  });

  T('★壊し方④ 紙が 痩せる（★行が 減る／数が 抜ける★）★', () => {
    /* ★『紙が 減っていない』の 試験が 本当に 効くか★
       ⇒ 紙を 写して 削り、同じ 判定を かけて ★赤に なるか★ を 見る
       ⇒ repo の 紙は 1バイトも 触らない */
    const 見張り = (写し) => {
      if (写し.length !== 180) return '行が 減った';
      if (new Set(写し.map((g) => g.引数)).size !== 12) return '引数が 減った';
      if (new Set(写し.map((g) => g.数)).size !== 15) return '数が 減った';
      return null;
    };
    if (見張り(紙) !== null) throw new Error('★今の 紙で 赤に なる★ … ' + 見張り(紙));
    /* ★★『引数が 減った』『数が 減った』を ★本当に 効かせる★★★
       ⇒ 減らしただけだと ★行数の 判定が 先に 赤に なって★
         後ろの 2つは ★1度も 効きません★（最初 私は そう 書いて いました）
       ⇒ だから ★180行の まま★ 中身だけ 入れ替える */
    const 埋めて180に = (写し) => {
      const 出 = 写し.slice();
      while (出.length < 180) 出.push(出[0]);
      return 出.slice(0, 180);
    };
    const 壊し = [
      ['1行 減らす', 紙.slice(0, 179), '行が 減った'],
      ['引数を 1個 落とす', 埋めて180に(紙.filter((g) => g.引数.indexOf('TOCOL') !== 0)), '引数が 減った'],
      ['境目の 45294 を 落とす', 埋めて180に(紙.filter((g) => g.数 !== '45294')), '数が 減った'],
    ];
    for (const [札, 写し, 出るはず] of 壊し) {
      const 出 = 見張り(写し);
      if (出 === null) throw new Error('★' + 札 + '★ … 気づかない');
      if (出 !== 出るはず) throw new Error('★' + 札 + '★ … ★' + 出るはず + '★で 赤に なるはずが「' + 出 + '」');
      console.log('      ' + 札.padEnd(20) + ' → 赤（' + 出 + '）');
    }
  });
}

console.log('\n' + pass + ' passed, ' + fail + ' failed');
if (直に走った) process.exit(fail ? 1 : 0);
