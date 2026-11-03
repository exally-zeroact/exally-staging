/* formula-yosoku.test.mjs — ★予測・統計・単位を 実Excel の 答えと 突き合わせる★（2026-09-07）
 *
 *  ★★本番の 道で 押す★★
 *    JS層(_jsComputeFormula) → convertFormula → HyperFormula（★プラグイン 4つ 全部★）
 *    ⇒ 1つでも 積み忘れると ★動く物が「動かない」に 見える★
 *
 *  ★★答えは 3枚の 紙から★★
 *    golden-yosoku-2026-09-07.tsv   … 予測・統計・表と情報（82本）
 *    golden-convert-2026-09-07.tsv  … 単位の 係数（243本）
 *    golden-convert3-2026-09-07.tsv … 接頭辞が 付く 単位（147本）
 *    ⇒ どれも Excel 16.0 build 20326（UI 1041）に 打たせた 物
 *
 *  ★★出さない 物も 見張る★★
 *    AREAS / CELL / FILTERXML / TRIMRANGE / FORECAST.ETS 4つ
 *    （★INFO は 2026-09-07「全部やって」で 出しました★）
 *    ⇒ 出す 名簿に 入っていない事／動かない 棚に 訳つきで 載っている事
 *
 *  使い方: node tests/formula-yosoku.test.mjs
 *          node tests/formula-yosoku.test.mjs --self-test
 */
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath, pathToFileURL } from 'node:url';

const ここ = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(ここ, '..');
const 紙 = path.join(ROOT, 'docs/measured/kansuu46');
const require_ = createRequire(import.meta.url);

const 直に走った = process.argv[1] && pathToFileURL(process.argv[1]).href === import.meta.url;
const 自己試験 = 直に走った && process.argv.includes('--self-test');

let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ✓ ' + n); } catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + (e && e.message)); } };

/* ── 本番と 同じ 物を 積む ── */
const HFns = require_(path.join(ROOT, 'hyperformula.full.min.js'));
const EF = require_(path.join(ROOT, 'exally-formula.js'));
const 積1 = EF.registerExallyFunctions(HFns) === true;
const HF0 = HFns.HyperFormula;
const H = Object.assign(Object.create(HF0), HFns,
  { registerFunctionPlugin: HF0.registerFunctionPlugin.bind(HF0) });
const 積2 = require_(path.join(ROOT, 'lib/formula-extra-plug.js'))
  .つなぐ(H, require_(path.join(ROOT, 'lib/formula-extra.js')));
const 積3 = require_(path.join(ROOT, 'lib/formula-nokori-plug.js'))
  .つなぐ(H, require_(path.join(ROOT, 'lib/formula-nokori.js')));
const 積4 = require_(path.join(ROOT, 'lib/formula-kane-plug.js'))
  .つなぐ(H, require_(path.join(ROOT, 'lib/formula-kane.js')));
const Y = require_(path.join(ROOT, 'lib/formula-yosoku.js'));
const 積5 = require_(path.join(ROOT, 'lib/formula-yosoku-plug.js')).つなぐ(H, Y, function () {
  /* ★場の 事★＝試験では 決めた 物を 入れる（画面は 本物を 入れる） */
  return { シート数: 3, 版: 'Exally test', 台: 'win', OS: 'Windows (64-bit) NT 10.00', 左上: '$A$1' };
});

const hf = HF0.buildEmpty({ licenseKey: 'gpl-v3', useArrayArithmetic: true, smartRounding: false });
const SID = hf.getSheetId(hf.addSheet('S'));
EF.initExallyFormula(hf);
const JS層 = EF._jsComputeFormula || null;

const 赤の名 = (t) => ({
  NA: '#N/A', DIV_BY_ZERO: '#DIV/0!', VALUE: '#VALUE!', NUM: '#NUM!',
  NAME: '#NAME?', REF: '#REF!', CYCLE: '#CYCLE!', ERROR: '#ERROR!', SPILL: '#SPILL!',
}[t] || ('#' + t));

const 下敷き = [];
for (let r = 0; r < 8; r++) 下敷き.push([r + 1, (r + 1) * 2, (r + 1) * 3, (r + 1) * 4]);

/** ★1本 押す★（本番と 同じ 3段） */
export function 押す(式) {
  if (typeof JS層 === 'function') {
    try {
      const v = JS層(0, 式);
      if (v !== null && v !== undefined) return v;
    } catch (e) { return '★JS層で 例外★'; }
  }
  let 後;
  try { 後 = EF.convertFormula(式); } catch (e) { return '★書き換えで 例外★'; }
  try {
    const 表 = 下敷き.map((r) => r.slice());
    表.push([後]);
    /* ★こぼれる 場所を 空ける★（空文字を 置くと ★#SPILL!★ に なる＝null） */
    for (let i = 0; i < 40; i++) 表.push([null]);
    hf.setSheetContent(SID, 表);
    const v = hf.getCellValue({ sheet: SID, row: 下敷き.length, col: 0 });
    if (v && v.type) return 赤の名(v.type);
    return v;
  } catch (e) { return '★engine で 例外★'; }
}

/** ★合っているか★（純粋＝自己試験からも 呼べる） */
export function 合うか(行, 出) {
  if (行.型 === 'Double' && typeof 出 === 'number') {
    const 正 = Number(行.答);
    return Math.abs(出 - 正) <= Math.max(1e-9, Math.abs(正) * 1e-9);
  }
  if (行.型 === 'Boolean') return String(出).toUpperCase() === String(行.答).toUpperCase();
  return String(出) === String(行.答);
}

const 読む = (名) => fs.readFileSync(path.join(紙, 名), 'utf-8').split('\n')
  .filter((l) => l && !l.startsWith('#'))
  .map((l) => { const p = l.split('\t'); return { 名: p[0], 式: p[1], 答: p[2], 型: p[3] }; });

const 紙たち = ['golden-yosoku-2026-09-07.tsv', 'golden-convert-2026-09-07.tsv', 'golden-convert3-2026-09-07.tsv'];
const 行たち = [].concat(...紙たち.map(読む));
const 出す = Y.足した名前();
const 保留 = Y.保留の名前();

console.log('\n[formula-yosoku] 予測・統計・単位 — 実Excel の 答えと 突き合わせ');
console.log('  実Excel の 答え … ' + 行たち.length + '本（Excel 16.0 build 20326 / UI 1041）');

T('★★本番と 同じ物を 積めた★★（1つ 忘れると 嘘の 数に なる）', () => {
  if (!積1) throw new Error('exally-formula を 積めていない');
  if (!(積2 > 0 && 積3 > 0 && 積4 > 0 && 積5 > 0)) {
    throw new Error('プラグインを 積めていない（' + [積2, 積3, 積4, 積5].join('/') + '）');
  }
  if (typeof JS層 !== 'function') throw new Error('JS層に 手が 届かない');
});

T('★答えの 紙が 空でない／どの Excel かが 書いてある★', () => {
  if (行たち.length < 300) throw new Error('答えが ' + 行たち.length + '本しか ない');
  for (const n of 紙たち) {
    if (!/build/.test(fs.readFileSync(path.join(紙, n), 'utf-8').slice(0, 400))) {
      throw new Error(n + ' に どの Excel で 打ったかが 書かれていない');
    }
  }
});

/* ★★実Excel と 同じに ならない と 決めた 物★★（1つずつ 訳を 書く）
   ★2026-09-14 に `tests/betsu-atsukai.mjs` へ 出しました★＝
     突き合わせの 道具が 同じ 名簿を ★もう一度 手で 持つ★のを 止める為。
     （物差し `tests/monosashi.mjs`・材料（紙の `#材料`）と 同じ 考え＝★1か所★） */
const { 別扱い } = await import('./betsu-atsukai.mjs');

T('★実Excel と 同じに しない 物には 訳が 書いてある★', () => {
  for (const n of Object.keys(別扱い)) {
    if (出す.indexOf(n) < 0) throw new Error(n + ' は 出していないのに 別扱いに 書いてある');
    if (String(別扱い[n]).length < 20) throw new Error(n + ' の 訳が 短すぎる');
  }
});

T('★出す 物は 1本 残らず 実Excel と 同じ★（別扱いを 除く）', () => {
  const 外れ = [];
  let 見た = 0;
  for (const 行 of 行たち) {
    if (出す.indexOf(行.名) < 0) continue;
    if (別扱い[行.名]) continue;
    見た++;
    const 出 = 押す(行.式);
    if (!合うか(行, 出)) 外れ.push(行.式 + '\n        正 ' + 行.答 + ' ／ 出 ' + 出);
  }
  if (見た < 300) throw new Error('押した 式が ' + 見た + '本しか ない（材料 不足）');
  console.log('      押した 式 … ' + 見た + '本');
  if (外れ.length) throw new Error('★' + 外れ.length + '本 違う★\n      ' + 外れ.slice(0, 5).join('\n      '));
});

T('★★そのまま セルに 入れても 動く（こぼれる 関数）★★', () => {
  /* 実測（2026-09-07）… 大きさを ★ArraySize の 形で★ 返していなかった 間、
     `=MUNIT(3)` は ★#ERROR!★ だった。`=INDEX(MUNIT(3),1,1)` は 動くので
     ★試験では 見つからなかった★ ⇒ ★そのまま 入れる 形も 押す★ */
  const 見る = ['=MUNIT(3)', '=TREND(A1:A6)', '=GROWTH(A1:A6)', '=RANDARRAY(2,2)',
    '=LOGEST(A1:A6)', '=TEXTSPLIT("a,b,c",",")', '=MINVERSE({1,2;3,4})'];
  const 死 = [];
  for (const f of 見る) {
    const v = 押す(f);
    if (typeof v === 'string' && v.charAt(0) === '#') 死.push(f + ' → ' + v);
  }
  if (死.length) throw new Error('★そのまま 入れると 死ぬ★: ' + 死.join(' / '));
});

T('★でたらめの 表は 頼まれた 形と 範囲を 守る★（計算だけを 直に 押す）', () => {
  /* ★エンジンを 通さずに 押す★＝
     `SUMPRODUCT(--(MOD(RANDARRAY(…),1)=0))` の ような 式で 見ようとして
     ★組み合わせの 側の 違い★を 掴んで しまった（実Excel 5／こちら 2）。
     ⇒★見たい 物（整数か・範囲の 中か）だけを 直に 見る★ */
  for (let 回 = 0; 回 < 200; 回++) {
    const a = Y.でたらめの表(3, 2, 5, 10, true);
    if (a.length !== 3 || a[0].length !== 2) throw new Error('形が 違う');
    for (const 行 of a) for (const v of 行) {
      if (v < 5 || v > 10) throw new Error('範囲の 外 … ' + v);
      if (Math.floor(v) !== v) throw new Error('整数に なっていない … ' + v);
    }
    const b = Y.でたらめの表(1, 1);
    if (!(b[0][0] >= 0 && b[0][0] < 1)) throw new Error('0以上1未満に なっていない');
  }
});

T('★★INFO は 分かる 物だけ 答え、無い 物は 作らない★★（2026-09-07 司さん「全部やって」）', () => {
  const 見 = [['=INFO("numfile")', 3], ['=INFO("recalc")', '自動'],
    ['=INFO("system")', 'pcdos'], ['=INFO("release")', 'Exally test']];
  for (const [f, 正] of 見) {
    const v = 押す(f);
    if (String(v) !== String(正)) throw new Error(f + ' … 正 ' + 正 + ' ／ 出 ' + v);
  }
  /* ★パソコンの フォルダは ブラウザに 無い★＝それらしい 物を 作らない */
  for (const f of ['=INFO("directory")', '=INFO("memavail")']) {
    if (押す(f) !== '#N/A') throw new Error(f + ' が #N/A で ない … ' + 押す(f));
  }
  if (押す('=INFO("なんとか")') !== '#VALUE!') throw new Error('知らない 合言葉が #VALUE! で ない');
});

T('★単位の 表は 手で 書いていない★（答えの 紙から 機械で 作った ままか）', () => {
  const 出 = require_('node:child_process').spawnSync(process.execPath,
    [path.join(ROOT, 'scripts/make-tanni-hyou.mjs'), '--check'], { encoding: 'utf-8' });
  if (出.status !== 0) throw new Error((出.stdout || '') + (出.stderr || ''));
});

T('★保留の 物は 出していない★（半分 合う 計算を 客に 見せない）', () => {
  for (const n of 保留) if (出す.indexOf(n) >= 0) throw new Error(n + ' が 出す 名簿に 入っている');
  if (!保留.length) throw new Error('保留が 0個＝棚が 空＝数え忘れ');
});

T('★保留の 物は「動かない 棚」に 訳つきで 載っている★', () => {
  const EX = require_(path.join(ROOT, 'lib/formula-extra.js'));
  const 棚 = (EX.数える && EX.数える().足さない) || {};
  for (const n of 保留) {
    const 訳 = 棚[n];
    if (!訳) throw new Error(n + ' が 動かない 棚に 載っていない（客は #NAME? を 食らうだけ）');
    if (String(訳).length < 8) throw new Error(n + ' の 訳が 短すぎる（' + 訳 + '）');
  }
});

T('★名簿は 1つだけ★（繋ぐ 側が 別の 名簿を 持っていない）', () => {
  const src = fs.readFileSync(path.join(ROOT, 'lib/formula-yosoku-plug.js'), 'utf-8');
  const 並 = [...src.matchAll(/^\s*'([A-Z][A-Z0-9.]*)':\s*\{\s*method:/gm)].map((m) => m[1]);
  const a = [...出す].sort().join(','), b = [...並].sort().join(',');
  if (a !== b) throw new Error('名簿が 違う\n      lib  … ' + a + '\n      plug … ' + b);
});

T('★CONVERT は もう JS層に 居ない★（1つの 関数は 1か所でだけ）', () => {
  const src = fs.readFileSync(path.join(ROOT, 'exally-formula.js'), 'utf-8');
  if (/^\s*CONVERT\s*:\s*1\s*,/m.test(src)) throw new Error('_jsSet に CONVERT が 残っている');
  if (/function\s+_jsConvert\s*\(/.test(src)) throw new Error('_jsConvert が 残っている（手で 書いた 概数）');
});

if (自己試験) {
  console.log('\n[self-test] ★物差しが 効いているか★');
  T('わざと 1つ ずらした 答えは 赤に なる', () => {
    const 元 = 行たち.find((r) => r.名 === 'CONVERT' && r.型 === 'Double' && Number(r.答) > 1);
    if (合うか({ 答: String(Number(元.答) * 1.001), 型: 'Double' }, Number(元.答))) {
      throw new Error('違う 答えでも 緑に なった');
    }
  });
  T('赤い 答えを 数字に すり替えたら 赤に なる', () => {
    if (合うか({ 答: '#N/A', 型: 'String' }, 123)) throw new Error('赤なのに 数字で 緑に なった');
  });
  T('★1ポンドが 4桁に 丸まっていたら 赤に なる★（前の 中身）', () => {
    if (合うか({ 答: '0.45359237', 型: 'Double' }, 0.4536)) {
      throw new Error('★丸めた 数でも 緑に なった＝前の 間違いを 見つけられない★');
    }
  });
}

console.log('\n' + pass + ' passed, ' + fail + ' failed');
if (直に走った) process.exit(fail ? 1 : 0);
