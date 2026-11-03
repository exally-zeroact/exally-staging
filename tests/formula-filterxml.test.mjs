/* formula-filterxml.test.mjs — ★FILTERXML を 実Excel の 答えと 突き合わせる★（2026-09-07）
 *
 *  ★★本番の 道で 押す★★
 *    JS層 → convertFormula → HyperFormula（★プラグインは 本番と 同じ 全部★）
 *  ★★XPath は ブラウザの 物★★
 *    画面は `window.DOMParser`／試験は ★jsdom の 同じ 物★を 入れる
 *    ⇒★自分で 少しだけ 作ると 通る 式と 通らない 式が 混ざる★ので 作っていない
 *
 *  ★答え★ … docs/measured/kansuu46/golden-filterxml-2026-09-07.tsv（37本）
 *
 *  使い方: node tests/formula-filterxml.test.mjs
 *          node tests/formula-filterxml.test.mjs --self-test
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

/* ── XPath の 道具（本番は 窓／ここは jsdom） ── */
let 部品 = null;
try {
  const { JSDOM } = require_('jsdom');
  const w = new JSDOM('').window;
  部品 = { DOMParser: w.DOMParser, XPathResult: w.XPathResult };
} catch (e) { 部品 = null; }

const F = require_(path.join(ROOT, 'lib/formula-filterxml.js'));

/* ── 本番と 同じ 物を 積む ── */
const HFns = require_(path.join(ROOT, 'hyperformula.full.min.js'));
const EF = require_(path.join(ROOT, 'exally-formula.js'));
const 積1 = EF.registerExallyFunctions(HFns) === true;
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
require_(path.join(ROOT, 'lib/formula-soto-plug.js'))
  .つなぐ(H, require_(path.join(ROOT, 'lib/formula-soto.js')), {
    取る: async () => { throw new Error('ここでは 外へ 出ません'); },
    聞く: async () => { throw new Error('ここでは AI に 聞きません'); },
    再計算: () => {},
  });
const 積X = require_(path.join(ROOT, 'lib/formula-filterxml-plug.js')).つなぐ(H, F, () => 部品);

const hf = HF0.buildEmpty({ licenseKey: 'gpl-v3', useArrayArithmetic: true, smartRounding: false });
const SID = hf.getSheetId(hf.addSheet('S'));
EF.initExallyFormula(hf);

const 赤の名 = (t) => ({
  NA: '#N/A', DIV_BY_ZERO: '#DIV/0!', VALUE: '#VALUE!', NUM: '#NUM!',
  NAME: '#NAME?', REF: '#REF!', CYCLE: '#CYCLE!', ERROR: '#ERROR!', SPILL: '#SPILL!',
}[t] || ('#' + t));

function 押す(式) {
  let 後;
  try { 後 = EF.convertFormula(式); } catch (e) { return '★書き換えで 例外★'; }
  try {
    const 表 = [[後]];
    for (let i = 0; i < 40; i++) 表.push([null]);
    hf.setSheetContent(SID, 表);
    const v = hf.getCellValue({ sheet: SID, row: 0, col: 0 });
    if (v && v.type) return 赤の名(v.type);
    return v;
  } catch (e) { return '★engine で 例外★'; }
}

const 金 = fs.readFileSync(path.join(ROOT, 'docs/measured/kansuu46/golden-filterxml-2026-09-07.tsv'), 'utf-8')
  .split('\n').filter((l) => l && !l.startsWith('#'))
  .map((l) => { const p = l.split('\t'); return { 名: p[0], 式: p[1], 答: p[2], 型: p[3] }; });

console.log('\n[formula-filterxml] XML から 取り出す — 実Excel の 答えと 突き合わせ');
console.log('  実Excel の 答え … ' + 金.length + '本（Excel 16.0 build 20326 / UI 1041）');

T('★★XPath の 道具が 在る★★（無いと 全部 #VALUE! で ★偽の 緑★に なる）', () => {
  if (!部品) throw new Error('jsdom が 無い＝本番と 同じ 道具で 押せていない');
  if (!(積X > 0)) throw new Error('繋げていない');
  if (!積1) throw new Error('exally-formula を 積めていない');
  /* ★物差しが 生きている★＝ちゃんと 取れる 事を 1本 押す */
  const v = 押す('=FILTERXML("<a><b>7</b></a>","//b")');
  if (String(v) !== '7') throw new Error('取れるはずの 物が 取れない … ' + v);
});

/* ★★jsdom では 通らないが 本物の ブラウザでは 通る 式★★（1つずつ 訳と 証拠を 書く）
   ⇒★『試験が 赤い＝お客さんの 所でも 動かない』とは 限らない★
   ⇒★でも ★確かめずに「ブラウザなら 動く」と 書くのは 嘘★★
   ⇒ だから ★本物の Chrome で 押した 記録★を 置いて、それを 名指しで 見る */
const jsdomが弱い = {
  "name()": {
    印: "name()",
    訳: 'jsdom の XPath は `name()` を 知らない（投げる）。'
      + ' ★本物の Chrome では 3件 取れる★＝実Excel と 同じ。',
    証拠: 'docs/measured/kansuu46/xpath-honmono-de-osu.txt',
  },
};

T('★★「ブラウザなら 動く」は ★本物で 押した 記録★が 在る★★', () => {
  for (const k of Object.keys(jsdomが弱い)) {
    const 道 = path.join(ROOT, jsdomが弱い[k].証拠);
    if (!fs.existsSync(道)) throw new Error('証拠の 紙が 無い … ' + jsdomが弱い[k].証拠);
    const 紙 = fs.readFileSync(道, 'utf-8');
    if (!/本物の Chrome/.test(紙)) throw new Error('本物で 押した 紙では ない');
    if (紙.indexOf(k) < 0) throw new Error('その 式が 紙に 無い … ' + k);
    if (!/本物の Chrome では 取れる/.test(紙)) throw new Error('本物で 取れた事が 書かれていない');
  }
});

T('★★実Excel と 1本 残らず 同じ★★（jsdom が 弱い 分は 別に 数える）', () => {
  const 外れ = [];
  let 逃 = 0;
  for (const 行 of 金) {
    const 弱 = Object.keys(jsdomが弱い).find((k) => 行.式.indexOf(k) >= 0);
    if (弱) { 逃++; continue; }
    const 出 = 押す(行.式);
    let よい;
    if (行.型 === 'Double' && typeof 出 === 'number') {
      よい = Math.abs(出 - Number(行.答)) <= 1e-9;
    } else よい = String(出) === String(行.答);
    if (!よい) 外れ.push(行.式.slice(0, 70) + '\n        正 ' + 行.答 + ' ／ 出 ' + 出);
  }
  console.log('      押した 式 … ' + (金.length - 逃) + '本'
    + (逃 ? '／★jsdom が 弱くて 押せない ' + 逃 + '本（本物の Chrome では 取れる事を 別に 数えた）★' : ''));
  if (逃 > 4) throw new Error('★逃がした 数が 多すぎる★（' + 逃 + '本）＝物差しが 空洞に なる');
  if (外れ.length) throw new Error('★' + 外れ.length + '本 違う★\n      ' + 外れ.slice(0, 5).join('\n      '));
});

T('★そのまま セルに 入れても こぼれる★（中に 入れた 時だけ 動く を 止める）', () => {
  const v = 押す('=FILTERXML("<a><b>1</b><b>2</b></a>","//b")');
  if (typeof v === 'string' && v.charAt(0) === '#') throw new Error('裸で 死んだ … ' + v);
});

T('★出す 名簿は lib が 正本★', () => {
  const src = fs.readFileSync(path.join(ROOT, 'lib/formula-filterxml-plug.js'), 'utf-8');
  const 並 = [...src.matchAll(/^\s*'([A-Z][A-Z0-9.]*)':\s*\{\s*method:/gm)].map((m) => m[1]);
  if ([...F.足した名前()].sort().join(',') !== [...並].sort().join(',')) throw new Error('名簿が 違う');
});

if (自己試験) {
  console.log('\n[self-test] ★物差しが 効いているか★');
  T('★道具を 外したら 全部 #VALUE!★（＝偽の 緑を 見分けられる）', () => {
    const r = F.取り出す('<a><b>1</b></a>', '//b', null);
    if (!(r && r.誤り === 'VALUE')) throw new Error('道具が 無いのに 答えを 返した');
  });
  T('★1つも 見つからない時は 空では なく #VALUE!★（実Excel と 同じ）', () => {
    const r = F.取り出す('<a><b>1</b></a>', '//zzz', 部品);
    if (!(r && r.誤り === 'VALUE')) throw new Error('空を 返した');
  });
  T('★text() は 取らない★（実Excel と 同じ）', () => {
    const r = F.取り出す('<a><b>1</b></a>', '//b/text()', 部品);
    if (!(r && r.誤り === 'VALUE')) throw new Error('字の かたまりを 返した');
  });
}

console.log('\n' + pass + ' passed, ' + fail + ' failed');
if (直に走った) process.exit(fail ? 1 : 0);
