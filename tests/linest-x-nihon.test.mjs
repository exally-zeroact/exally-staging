/* linest-x-nihon.test.mjs — ★LINEST が 表として 正しく 返るか★（2026-09-09）
 *
 *  ★★何が 変わったか★★
 *    ★前★ `exally-formula.js` の JS層が ★裸の =LINEST(範囲,範囲) だけ★を 拾い ★傾き 1つ★を 返して いた
 *      ・`=INDEX(LINEST(…),1,1)` は ★#NAME?★（エンジンに 積んで いない）
 *      ・★x が 2本以上（重回帰）で 静かに 違う 答え★
 *        =LINEST(G1:G6,H1:I6) … 実Excel 0.7708333333333329 ／ うち ★7.390243902439025★
 *    ★後★ `lib/formula-yosoku.js` の `直線の係数()` が ★表として★ 返す
 *      ・LOGEST と ★同じ 土台（回帰）★＝重回帰も 統計の 5行も 出せる
 *
 *  ★★一度 「断る」形に しかけて 止めました★★
 *    x が 2本以上を ★エラーに する★案を 作りかけましたが、
 *    ★司さんの 方針は「全部 つける」★＝★出来ない から 断る のは 方針に 反する★。
 *    ⇒★断らずに 直しました★（道具は もう 在った＝LOGEST が 動いて いた）
 *
 *  ★★何を 守るか（1つずつ 名指し）★★
 *    ①★x が 2本（重回帰）で 実Excel と 同じ 答え★（★静かに 違う 答えを 出さない★）
 *    ②★入れ子（=INDEX(LINEST(…),行,列)）が 動く★（★#NAME? に ならない★）
 *    ③★表の 大きさ★＝列 ＝ x の 本数 ＋ 1 ／ 行 ＝ 既定 1・統計 TRUE で 5
 *    ④★余りの マスは #N/A★（★[object Object] に しない★）
 *    ⑤★自由度 0 なら F は #NUM!★（★割れない 物を 0 と 言わない★）
 *    ⑥★x が 1本（縦1列・横1行）も 今まで どおり 合う★
 *
 *  ★★見て いない 範囲（★書かない 見張りは「全部 守った」と 読まれる★）★★
 *    ・★INDEX の 境目は 見て いません★＝`=INDEX(…,0,1)` `(…,6,1)` など
 *      ⇒★これは LINEST では なく INDEX の 話★（普通の 範囲でも 同じ 答えに なる事を 確かめた）
 *        実Excel … 0 は「その 行/列 ぜんぶ」／範囲の 外は #REF!
 *        うち …… 0 は #VALUE ／ 範囲の 外は #NUM
 *      ⇒★別件として 棚に★
 *    ・★残差が ほぼ 0 の 組の 統計の 行は 見て いません★
 *      （ぴったり 乗る データでは 1.86e-15 対 1.64e-14 の ように ★0 に 近い 桁が ばらつく★）
 *    ・LOGEST / TREND / GROWTH は ★別の 紙で 測って 全部 合って います★（ここでは 見ない）
 *    ・★LOGEST の 余りの マスが '' か #N/A かは 未測定★（LINEST しか 測って いない）
 *
 *  使い方: node tests/linest-x-nihon.test.mjs [--self-test]
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createRequire } from 'node:module';

const ここ = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(ここ, '..');
const 直に走った = process.argv[1] && pathToFileURL(process.argv[1]).href === import.meta.url;
const 自己試験 = 直に走った && process.argv.includes('--self-test');

let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ✓ ' + n); } catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + (e && e.message)); } };

/* ══ ★本番と 同じ 建て方・同じ プラグイン★ ═══════════════════
   ★測り台が 本番と 違うと 嘘の 数字が 出ます★（2026-09-09 に 2つの 道具で 踏んだ）
   ⇒ book.html と 同じ 8本を つなぎ、★smartRounding:false★ で 建てる */
const require_ = createRequire(path.join(ROOT, 'package.json'));
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
require_(path.join(ROOT, 'lib/formula-yosoku-plug.js'))
  .つなぐ(H, require_(path.join(ROOT, 'lib/formula-yosoku.js')),
    () => ({ シート数: 1, 版: 'Exally', 台: 'win', OS: '', 左上: '$A$1' }));
const hf = HF0.buildEmpty({
  licenseKey: 'gpl-v3', useArrayArithmetic: true, smartRounding: false,
  maxRows: 1048576, maxColumns: 18278,
});
const SID = hf.getSheetId(hf.addSheet('S'));
EF.initExallyFormula(hf);

/* ══ ★材料（★紙と 同じ★・下の 試験で 突き合わせる）★ ══════════ */
const 材料 = {
  A1: 100, A2: 120, A3: 140, A4: 160, A5: 180,      /* y（きれいに 乗る） */
  B1: 1, B2: 2, B3: 3, B4: 4, B5: 5,                /* x 1本 */
  D1: 2.7, D2: 5.1, D3: 6.9, D4: 9.4, D5: 11.2,     /* y（端数） */
  E1: 1, E2: 2, E3: 3, E4: 4, E5: 5,
  G1: 10, G2: 14, G3: 21, G4: 25, G5: 33, G6: 38,   /* y（重回帰） */
  H1: 1, H2: 2, H3: 3, H4: 4, H5: 5, H6: 6,         /* x 1本目 */
  I1: 2, I2: 1, I3: 4, I4: 3, I5: 6, I6: 5,         /* x 2本目 */
  N1: 100, O1: 120, P1: 140, Q1: 160, R1: 180, S1: 200, /* ★横1行★の y */
  N2: 1, O2: 2, P2: 3, Q2: 4, R2: 5, S2: 6,             /* ★横1行★の x */
};
function マスを解く(s) {
  const m = /^([A-Z]+)(\d+)$/.exec(s);
  let c = 0;
  for (const ch of m[1]) c = c * 26 + (ch.charCodeAt(0) - 64);
  return { 行: Number(m[2]) - 1, 列: c - 1 };
}
const 式の行 = 40;
function 押す(式) {
  let 最大列 = 0;
  const 置く = [];
  for (const k of Object.keys(材料)) {
    const rc = マスを解く(k);
    置く.push([rc, 材料[k]]);
    if (rc.列 > 最大列) 最大列 = rc.列;
  }
  const 板 = [];
  for (let r = 0; r <= 式の行 + 1; r++) 板.push(new Array(最大列 + 3).fill(null));
  for (const [rc, v] of 置く) 板[rc.行][rc.列] = v;
  /* ★本番と 同じ 道★ … ①JS層 → ②convertFormula → エンジン */
  hf.setSheetContent(SID, 板.map((r) => r.slice()));
  const js = EF._jsComputeFormula(0, 式);
  if (js !== null && js !== undefined) return String(js);
  板[式の行][0] = EF.convertFormula(式);
  hf.setSheetContent(SID, 板);
  const v = hf.getCellValue({ sheet: SID, row: 式の行, col: 0 });
  if (v && v.type) return '#' + v.type;
  return v === null || v === undefined ? '(空)' : String(v);
}
/* ★合う 幅★ … ★実Excel 自身の 桁の 端まで 合わせる★（1e-9） */
const 近い = (う, 正) => {
  const a = Number(う), b = Number(正);
  if (!isFinite(a) || !isFinite(b)) return false;
  if (b === 0) return a === 0;
  return Math.abs((a - b) / b) <= 1e-9;
};

console.log('\n[linest-x-nihon] ★LINEST が 表として 正しく 返るか★');

T('★★x が 2本（重回帰）で 実Excel と 同じ 答え★★', () => {
  const r = 押す('=LINEST(G1:G6,H1:I6)');
  if (!近い(r, 0.7708333333333329)) {
    throw new Error('★うち ' + r + ' ／ 実Excel 0.7708333333333329★'
      + '\n      ⇒★直す 前は 7.390243902439025 を 返して いた（★静かに 違う 答え★）★');
  }
  console.log('      … =LINEST(G1:G6,H1:I6) → ' + r + '（実Excel 0.7708333333333329）');
});

T('★★入れ子（=INDEX(LINEST(…),行,列)）が 動く（#NAME? に ならない）★★', () => {
  const 組 = [
    ['=INDEX(LINEST(A1:A5,B1:B5),1,1)', 19.999999999999993, '傾き'],
    ['=INDEX(LINEST(A1:A5,B1:B5),1,2)', 80.00000000000003, '切片'],
    ['=INDEX(LINEST(G1:G6,H1:I6),1,1)', 0.7708333333333329, '重回帰 1つ目'],
    ['=INDEX(LINEST(G1:G6,H1:I6),1,2)', 5.104166666666667, '重回帰 2つ目'],
    ['=INDEX(LINEST(G1:G6,H1:I6),1,3)', 2.937499999999999, '重回帰 切片'],
  ];
  for (const [f, 正, 札] of 組) {
    const r = 押す(f);
    if (r === '#NAME') throw new Error('★' + f + ' が #NAME?★＝★まだ 積まれて いない★');
    if (!近い(r, 正)) throw new Error('★' + 札 + '（' + f + '）… うち ' + r + ' ／ 実Excel ' + 正 + '★');
  }
  console.log('      … ' + 組.length + '本 とも 実Excel と 合う（★前は 全部 #NAME?★）');
});

T('★★表の 大きさ … 列 ＝ x の 本数 ＋ 1 ／ 行 ＝ 既定 1・統計 TRUE で 5★★', () => {
  const 組 = [
    ['=ROWS(LINEST(A1:A5,B1:B5))', 1], ['=COLUMNS(LINEST(A1:A5,B1:B5))', 2],
    ['=ROWS(LINEST(A1:A5,B1:B5,TRUE,TRUE))', 5], ['=COLUMNS(LINEST(A1:A5,B1:B5,TRUE,TRUE))', 2],
    ['=COLUMNS(LINEST(G1:G6,H1:I6))', 3], ['=ROWS(LINEST(G1:G6,H1:I6,TRUE,TRUE))', 5],
    ['=COLUMNS(LINEST(G1:G6,H1:I6,TRUE,TRUE))', 3],
  ];
  for (const [f, 正] of 組) {
    const r = 押す(f);
    if (Number(r) !== 正) throw new Error('★' + f + ' … うち ' + r + ' ／ 実Excel ' + 正 + '★');
  }
  console.log('      … ' + 組.length + '本 とも 合う（x1本→2列 ／ x2本→3列）');
});

T('★★余りの マスは #N/A（★[object Object] に しない★）★★', () => {
  for (const f of ['=INDEX(LINEST(G1:G6,H1:I6,TRUE,TRUE),3,3)',
    '=INDEX(LINEST(G1:G6,H1:I6,TRUE,TRUE),4,3)',
    '=INDEX(LINEST(G1:G6,H1:I6,TRUE,TRUE),5,3)']) {
    const r = 押す(f);
    if (/object/i.test(r)) throw new Error('★' + f + ' が ' + r + '★＝★中の 誤りが 素通りして いる★');
    if (r !== '#NA') throw new Error('★' + f + ' … うち ' + r + ' ／ 実Excel #N/A★');
  }
  console.log('      … 3本 とも #N/A（★前は [object Object]★）');
});

T('★★自由度 0 なら F は #NUM!（割れない 物を 0 と 言わない）★★', () => {
  /* ★点が 2つだけ＝自由度 0★（K/L を その場で 敷く） */
  材料.K1 = 3; 材料.K2 = 7; 材料.L1 = 1; 材料.L2 = 2;
  const r = 押す('=INDEX(LINEST(K1:K2,L1:L2,TRUE,TRUE),4,1)');
  delete 材料.K1; delete 材料.K2; delete 材料.L1; delete 材料.L2;
  if (r !== '#NUM') throw new Error('★うち ' + r + ' ／ 実Excel #NUM!★（★前は 0 を 返して いた★）');
  console.log('      … 点が 2つだけ → #NUM!（実Excel と 同じ）');
});

T('★x が 1本（縦1列・横1行）も 今まで どおり 合う★', () => {
  const 組 = [
    ['=LINEST(A1:A5,B1:B5)', 19.999999999999993, '縦1列'],
    ['=LINEST(D1:D5,E1:E5)', 2.1299999999999994, '縦1列・端数'],
    ['=LINEST(N1:S1,N2:S2)', 20.000000000000004, '★横1行★'],
  ];
  for (const [f, 正, 札] of 組) {
    const r = 押す(f);
    if (!近い(r, 正)) throw new Error('★' + 札 + '（' + f + '）… うち ' + r + ' ／ 実Excel ' + 正 + '★');
  }
  console.log('      … ' + 組.length + '本 とも 合う（★出来て いた 物を 止めて いない★）');
});

T('★紙（実測）と 材料が 同じ（★試験の 中で 材料が ずれて いない★）★', () => {
  const p = path.join(ROOT, 'docs/measured/golden-linest-hyou-2026-09-09.tsv');
  if (!fs.existsSync(p)) throw new Error('★紙が 無い★ … ' + p);
  const 紙 = {};
  for (const l of fs.readFileSync(p, 'utf-8').split('\n')) {
    if (!l.startsWith('#材料')) continue;
    const c = l.split('\t');
    if (c.length >= 3) 紙[c[1].trim()] = Number(c[2]);
  }
  const 見る = ['A1', 'A5', 'B1', 'B5', 'D1', 'D5', 'E1', 'E5',
    'G1', 'G6', 'H1', 'H6', 'I1', 'I6', 'N1', 'S1', 'N2', 'S2'];
  const 違う = 見る.filter((k) => 紙[k] !== 材料[k]);
  if (違う.length) {
    throw new Error('★紙と 材料が 違う … '
      + 違う.map((k) => k + '（紙 ' + 紙[k] + ' ／ 試験 ' + 材料[k] + '）').join(' / ') + '★');
  }
  console.log('      … 紙と 同じ（' + 見る.length + 'マスを 名指しで 見た）');
});

/* ══ ★自己試験（★壊すのは 写し★＝ファイルは 1バイトも 触らない）★ ══ */
if (自己試験) {
  console.log('\n★自己試験（★壊すのは 写し★＝ファイルは 1バイトも 触らない）★');

  T('★★JS層に LINEST が 戻ったら 気づける★★', () => {
    const s = fs.readFileSync(path.join(ROOT, 'exally-formula.js'), 'utf-8');
    if (/if\s*\(\s*mLi\s*[&)]/.test(s)) {
      throw new Error('★JS層が また LINEST を 受けて いる★'
        + '\n      ⇒★JS層は エンジンより 先に 答える＝表を 返せず 1つの 数に なる★');
    }
    if (!/LINEST は ここで 受けません/.test(s)) {
      throw new Error('★「ここで 受けません」の 断りが 消えた★＝★なぜ 空けて 在るかが 分からなく なる★');
    }
    console.log('      … JS層は LINEST を 受けて いない（断りも 残って いる）');
  });

  T('★★エンジンに LINEST が 積んで 在る★★', () => {
    const s = fs.readFileSync(path.join(ROOT, 'lib/formula-yosoku-plug.js'), 'utf-8');
    if (!/'LINEST'\s*:/.test(s)) throw new Error('★プラグインの 登録に LINEST が 無い★');
    if (!/大きさLINEST/.test(s)) throw new Error('★大きさの 出し方が 無い＝表の 形が 決まらない★');
    const f = fs.readFileSync(path.join(ROOT, 'lib/formula-yosoku.js'), 'utf-8');
    if (!/function 直線の係数/.test(f)) throw new Error('★直線の係数() が 無い★');
    console.log('      … 登録・大きさ・中身 の 3つとも 在る');
  });

  T('★★表の 中の 誤りが 素通りしない（[object Object] を 出さない）★★', () => {
    const s = fs.readFileSync(path.join(ROOT, 'lib/formula-yosoku-plug.js'), 'utf-8');
    if (!/表の 中の マスの 誤りも 直す/.test(s)) {
      throw new Error('★中の 誤りを 直す 所が 無い★＝★また [object Object] が 出る★');
    }
    console.log('      … 範囲に() が 表の 中の 誤りも 直して いる');
  });
}

console.log('\n' + pass + ' passed, ' + fail + ' failed');
if (fail) process.exit(1);
