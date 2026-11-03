/* formula-cell.test.mjs — ★CELL を 実Excel の 答えと 突き合わせる★（2026-09-07）
 *
 *  ★★本番の 道で 押す★★
 *    JS層 → convertFormula → HyperFormula（★プラグインは 本番と 同じ 全部★）
 *
 *  ★答え★ … 実Excel に 打たせた 7枚（合わせて 411本）
 *    golden-cell  … 91本（種類 × 場所）          ← ★engine で 押す★
 *    golden-cell2 … 111本（表示形式・そろえ・変わり種）
 *    golden-cell3 … 102本（色・かっこ・円・うちの 表示形式）
 *    golden-cell4 … 10本（別シート・隠し列・鍵・保存後）
 *    golden-cell5 … 81本（日付/時刻の 合図）
 *    golden-cell6 … 16本（残った あいまいな 形）
 *    golden-cell7 … 36本（通貨の 見分け方）
 *
 *  ★★実Excel と 同じ 場面を 作ってから 押す★★
 *    CELL は「そのマスが どう 見えているか」を 返すので、
 *    ★空の シートに 式だけ 置いても 物差しに ならない★。
 *    ⇒ 測った 時と ★同じ 中身・同じ 幅・同じ そろえ★を 作って 押す。
 *
 *  使い方: node tests/formula-cell.test.mjs
 *          node tests/formula-cell.test.mjs --self-test
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

const F = require_(path.join(ROOT, 'lib/formula-cell.js'));

/* ── 実Excel と 同じ 場面を 作る ──────────────────────────────
   toru-cell.ps1 が 置いた 中身と ★1つ 残らず 同じ★に する
     A1=123.456 / A2='あいう'(右そろえ) / A3='=1+2' / A4=空
     B1=-5 / B2='x' / C1=DATE(2024,1,2) 表示 yyyy/m/d / D1=1000 表示 #,##0
     B列の 幅=12文字
   ★幅は 点（px）で 持つ★ので 文字数 × (既定80点 ÷ 8.43文字) に 直して 入れる */
const 既定の点 = 80;
const 一文字の点 = 既定の点 / 8.43;
const 見た目の表 = {
  '0,0': {}, '1,0': { そろえ: '右' }, '2,0': {}, '3,0': {},
  '0,1': {}, '1,1': {},
  '0,2': { 表示形式: 'yyyy/m/d' },
  '0,3': { 表示形式: '#,##0' },
};
const 列の幅の点 = { 1: 12 * 一文字の点 };            /* B列（0 起点で 1） */
const 見た目 = (sheet, row, col) => Object.assign(
  { 幅の点: 列の幅の点[col] !== undefined ? 列の幅の点[col] : 既定の点, 既定の点: 既定の点,
    シート名: 'Sheet1', ブック名: '' },
  見た目の表[row + ',' + col] || {}
);

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
let 見た目を返す = 見た目;
const 積C = require_(path.join(ROOT, 'lib/formula-cell-plug.js'))
  .つなぐ(H, F, (s, r, c) => 見た目を返す(s, r, c));

const hf = HF0.buildEmpty({ licenseKey: 'gpl-v3', useArrayArithmetic: true, smartRounding: false });
const SID = hf.getSheetId(hf.addSheet('Sheet1'));
EF.initExallyFormula(hf);

const 赤の名 = (t) => ({
  NA: '#N/A', DIV_BY_ZERO: '#DIV/0!', VALUE: '#VALUE!', NUM: '#NUM!',
  NAME: '#NAME?', REF: '#REF!', CYCLE: '#CYCLE!', ERROR: '#ERROR!', SPILL: '#SPILL!',
}[t] || ('#' + t));

/* ★中身は 実Excel と 同じ★（F1 に 式を 置く＝測った 時と 同じ 場所） */
const 土台 = () => {
  const 表 = [];
  for (let r = 0; r < 12; r++) 表.push([null, null, null, null, null, null, null, null]);
  表[0][0] = 123.456;  表[1][0] = 'あいう'; 表[2][0] = '=1+2'; 表[3][0] = null;
  表[0][1] = -5;       表[1][1] = 'x';
  表[0][2] = '=DATE(2024,1,2)';
  表[0][3] = 1000;
  return 表;
};

function 押す(式, 表 = 土台(), 行 = 0, 列 = 5) {
  let 後;
  try { 後 = EF.convertFormula(式); } catch (e) { return '★書き換えで 例外★'; }
  try {
    const t = 表.map((r) => r.slice());
    t[行][列] = 後;
    hf.setSheetContent(SID, t);
    const v = hf.getCellValue({ sheet: SID, row: 行, col: 列 });
    if (v && v.type) return 赤の名(v.type);
    return v;
  } catch (e) { return '★engine で 例外★ ' + (e && e.message); }
}

/* ── 金（実Excel の 答え）を 読む ── */
const 読む = (名) => fs.readFileSync(path.join(ROOT, 'docs/measured/kansuu46/' + 名), 'utf-8')
  .split('\n').filter((l) => l && !l.startsWith('#')).map((l) => l.split('\t'));

const 金1 = 読む('golden-cell-2026-09-07.tsv').map((p) => ({ 式: p[1], 答: p[2], 型: p[3] }));
const 金2 = 読む('golden-cell2-2026-09-07.tsv').map((p) => ({ 式: p[1], 答: p[2], 型: p[3], 場面: p[4] || '' }));
const 金4 = 読む('golden-cell4-2026-09-07.tsv').map((p) => ({ 式: p[1], 答: p[2], 場面: p[3] || '' }));
const 形の金 = ['golden-cell3-2026-09-07.tsv', 'golden-cell5-2026-09-07.tsv',
  'golden-cell6-2026-09-07.tsv', 'golden-cell7-2026-09-07.tsv']
  .flatMap((n) => 読む(n).map((p) => ({ 元: n, 式: p[1], 答: p[2], 場面: p[3] || '' })));

console.log('\n[formula-cell] CELL — 実Excel の 答えと 突き合わせ');
console.log('  実Excel の 答え … ' + (金1.length + 金2.length + 金4.length + 形の金.length)
  + '本（Excel 16.0 build 20326 / UI 1041）');

T('★★物差しが 生きている★★（繋げていない／空回りを 見分ける）', () => {
  if (!(積C > 0)) throw new Error('CELL を 繋げていない');
  if (!積1) throw new Error('exally-formula を 積めていない');
  const v = 押す('=CELL("address",B7)');
  if (String(v) !== '$B$7') throw new Error('番地が 出ない … ' + v);
});

/* ── ①種類 × 場所（91本）を ★engine で★ 押す ───────────────── */
T('★★1枚目 91本＝実Excel と 1本 残らず 同じ★★', () => {
  const 外れ = [], 逃 = [];
  for (const 行 of 金1) {
    /* ★filename は うちが 返せない 物★＝別に 数える（下の 試験で 名指しで 見る） */
    if (/CELL\("filename"/.test(行.式)) { 逃.push(行.式); continue; }
    const 出 = 押す(行.式);
    let よい;
    if (行.答 === '(空)') よい = (出 === '' || 出 === null);
    else if (行.型 === 'Double') よい = Math.abs(Number(出) - Number(行.答)) <= 1e-9;
    else よい = String(出) === String(行.答);
    if (!よい) 外れ.push(行.式 + '\n        正 ' + JSON.stringify(行.答) + ' ／ 出 ' + JSON.stringify(出));
  }
  console.log('      押した 式 … ' + (金1.length - 逃.length) + '本');
  if (逃.length !== 7) throw new Error('逃がす 数が 変わった（filename の 7本のはず）… ' + 逃.length);
  if (外れ.length) throw new Error('★' + 外れ.length + '本 違う★\n      ' + 外れ.slice(0, 6).join('\n      '));
});

/* ── ②表示形式ごとの format / color / parentheses ─────────────
   ★ここは engine を 通さない★＝答えは ★表示形式の 字だけ★で 決まるので
   lib を 直に 押す（画面が どんな 表示形式を 持っているかは ①で 押している） */
function 場面から表示形式(場面) {
  const m = /（実際に 付いた 形 ([\s\S]*)）\s*$/.exec(場面);
  if (m) return m[1];
  const m2 = /^表示形式 ([\s\S]*)$/.exec(場面);
  return m2 ? m2[1] : null;
}
T('★★表示形式の 合図（format / color / parentheses）＝実Excel と 同じ★★', () => {
  const 全 = [
    ...金2.filter((r) => /^表示形式 /.test(r.場面)).map((r) => ({ 式: r.式, 答: r.答, 場面: r.場面 })),
    ...形の金.filter((r) => /表示形式|入れた 形/.test(r.場面)),
  ];
  const 外れ = [];
  let 数 = 0;
  for (const 行 of 全) {
    if (!行.式 || 行.式.indexOf('CELL(') !== 0 && 行.式.indexOf('=CELL(') !== 0) continue;
    const k = /CELL\("([a-z]+)"/.exec(行.式);
    if (!k) continue;
    const fmt = 場面から表示形式(行.場面);
    if (fmt === null) continue;
    let 出;
    if (k[1] === 'format') 出 = F.書式の合図(fmt);
    else if (k[1] === 'color') 出 = F.色がつくか(fmt) ? 1 : 0;
    else if (k[1] === 'parentheses') 出 = F.かっこがつくか(fmt) ? 1 : 0;
    else continue;
    数++;
    if (String(出) !== String(行.答)) {
      外れ.push('表示形式 ' + JSON.stringify(fmt) + ' の ' + k[1]
        + '\n        正 ' + 行.答 + ' ／ 出 ' + 出);
    }
  }
  console.log('      押した 表示形式 … ' + 数 + '本');
  if (数 < 200) throw new Error('★押した 数が 少なすぎる★（' + 数 + '本）＝読み取りが 壊れている');
  if (外れ.length) throw new Error('★' + 外れ.length + '本 違う★\n      ' + 外れ.slice(0, 8).join('\n      '));
});

/* ── ③そろえ方ごとの prefix ─────────────────────────────── */
T('★そろえ方ごとの prefix＝実Excel と 同じ（10本）★', () => {
  const 表 = { 既定: '', 左: '左', 中央: '中央', 右: '右', 繰り返し: '繰り返し' };
  const 外れ = [];
  let 数 = 0;
  for (const 行 of 金2) {
    const m = /^そろえ (.+)／中身 (.+)$/.exec(行.場面 || '');
    if (!m) continue;
    数++;
    const 出 = F.そろえの印(表[m[1]], m[2] === '字');
    const 正 = 行.答 === '(空)' ? '' : 行.答;
    if (出 !== 正) 外れ.push(行.場面 + '  正 ' + JSON.stringify(正) + ' ／ 出 ' + JSON.stringify(出));
  }
  if (数 !== 10) throw new Error('そろえの 本数が 違う … ' + 数);
  if (外れ.length) throw new Error('★' + 外れ.length + '本 違う★\n      ' + 外れ.join('\n      '));
});

/* ── ④type / contents の 変わり種を ★engine で★ 押す ────────── */
T('★★中身の 変わり種（空の 字・誤り・真偽・空）＝実Excel と 同じ★★', () => {
  /* 測った 時と 同じ 中身を C1..C6 に 置く（式は G1 に 置く）
     ★2枚目を 測った 時の A1 は 1234.5678★（1枚目の 123.456 では ない）
     ⇒`=CELL("contents",C6)` は C6 の `=A1` を 見るので ★そこを 合わせないと 嘘の 赤★ */
  const 表 = 土台();
  表[0][0] = 1234.5678;
  表[0][2] = '=""'; 表[1][2] = '=1/0'; 表[2][2] = '=TRUE()';
  表[3][2] = null;  表[4][2] = null;   表[5][2] = '=A1';
  const 場面の場 = { 'C1': [0, 2], 'C2': [1, 2], 'C3': [2, 2], 'C4': [3, 2], 'C5': [4, 2], 'C6': [5, 2] };
  const 外れ = [];
  let 数 = 0;
  for (const 行 of 金2) {
    const m = /^=CELL\("(type|contents)",(C[1-6])\)$/.exec(行.式 || '');
    if (!m) continue;
    数++;
    const [r, c] = 場面の場[m[2]];
    const 式 = '=CELL("' + m[1] + '",' + String.fromCharCode(65 + c) + (r + 1) + ')';
    const 出 = 押す(式, 表, 0, 6);
    let 正 = 行.答;
    /* ★C4 は 実Excel でも 打てなかった マス★＝空と 同じ 扱い */
    if (正 === 'True') 正 = 'TRUE';
    let よい;
    if (正 === '(空)') よい = (出 === '' || 出 === null);
    else if (/^-?[0-9.]+$/.test(正)) よい = Math.abs(Number(出) - Number(正)) <= 1e-9;
    else よい = String(出).toUpperCase() === String(正).toUpperCase();
    if (!よい) 外れ.push(式 + '  正 ' + JSON.stringify(正) + ' ／ 出 ' + JSON.stringify(出));
  }
  if (数 !== 12) throw new Error('変わり種の 本数が 違う … ' + 数);
  if (外れ.length) throw new Error('★' + 外れ.length + '本 違う★\n      ' + 外れ.join('\n      '));
});

/* ── ⑤形の 変わり種（2つ目が 無い・範囲・大文字） ────────────── */
T('★★2つ目が 無い／範囲／大文字＝実Excel と 同じ★★', () => {
  const 期待 = {
    '=CELL("address",A1:C3)': '$A$1',
    '=CELL("row",A1:C3)': 1,
    '=CELL("contents",A1:C3)': 123.456,
    '=CELL("ADDRESS",A1)': '$A$1',
    '=CELL("Address",A1)': '$A$1',
  };
  const 外れ = [];
  for (const 式 of Object.keys(期待)) {
    const 出 = 押す(式);
    if (String(出) !== String(期待[式])) 外れ.push(式 + '  正 ' + 期待[式] + ' ／ 出 ' + 出);
  }
  /* ★2つ目が 無い＝式が 居る マス★（測った 時は H1＝1行1列）
     ここでは F1（1行6列）に 置くので 行1・列6・$F$1 に なる */
  const a = 押す('=CELL("row")'), b = 押す('=CELL("col")'), c = 押す('=CELL("address")');
  if (Number(a) !== 1) 外れ.push('=CELL("row") 正 1 ／ 出 ' + a);
  if (Number(b) !== 6) 外れ.push('=CELL("col") 正 6（F列）／ 出 ' + b);
  if (String(c) !== '$F$1') 外れ.push('=CELL("address") 正 $F$1 ／ 出 ' + c);
  if (外れ.length) throw new Error('★' + 外れ.length + '本 違う★\n      ' + 外れ.join('\n      '));
});

/* ── ⑥4枚目（別シート・隠し列・鍵・保存後） ────────────────── */
T('★★隠した 列の 幅は 0／鍵を 外した マスは 0★★（実測）', () => {
  if (F.幅を文字にする(80, 80, true) !== 0) throw new Error('隠した 列で 0 に ならない');
  if (F.幅を文字にする(80, 80, false) !== 8) throw new Error('ふつうの 列で 8 に ならない');
  if (F.セルの事('protect', { ロックなし: true }) !== 0) throw new Error('鍵を 外して 0 に ならない');
  if (F.セルの事('protect', {}) !== 1) throw new Error('ふつうの マスで 1 に ならない');
});

T('★★別の シートを 指した 番地の 形が 実Excel と 同じ★★', () => {
  const 正 = 金4.find((r) => r.式 === '=CELL("address",二枚目!B3)');
  if (!正) throw new Error('金の 行が 無い');
  const 出 = F.番地(2, 1, true, '二枚目', 'Book3');
  if (出 !== 正.答) throw new Error('正 ' + 正.答 + ' ／ 出 ' + 出);
  /* ★うちに ブック名が 無い 時は 空の かっこを 出さない★ */
  if (F.番地(2, 1, true, '二枚目', '') !== '二枚目!$B$3') throw new Error('空の かっこを 出した');
});

T('★★filename＝うちは ""（保存していない Excel と 同じ）★★', () => {
  const 前 = 金4.find((r) => r.場面 === '保存する 前');
  if (!前) throw new Error('金の 行が 無い');
  if (!(前.答 === '' || 前.答 === undefined)) throw new Error('金が 空では ない … ' + JSON.stringify(前.答));
  const 出 = 押す('=CELL("filename",A1)');
  if (!(出 === '' || 出 === null)) throw new Error('空を 返していない … ' + JSON.stringify(出));
  /* ★保存した 後の 形は 出さない★＝うちに ファイルの 置き場が 無い（作らない） */
  const 後 = 金4.find((r) => r.場面 === '保存した 後');
  if (!後 || 後.答.indexOf('[') < 0) throw new Error('金の 形が 変わった');
});

T('★知らない 種類は #VALUE!★（実測）', () => {
  const 出 = 押す('=CELL("こんなの",A1)');
  if (String(出) !== '#VALUE!') throw new Error('#VALUE! で ない … ' + 出);
});

/* ★★ここまでは 部品と エンジン★★＝★画面の 繋ぎは 別に 押した★
   book.html が 見た目を 渡す 所（`マス.numFmt` など）は ここでは 通っていない。
   ⇒★書き間違えても この 試験は 全部 緑に なる★
   ⇒ だから ★本物の Chrome で お客さんと 同じ 押し方★で 押した 記録を 名指しで 見る */
T('★★「画面でも 動く」は ★本物で 押した 記録★が 在る★★', () => {
  const 道 = path.join(ROOT, 'docs/measured/kansuu46/cell-honmono-de-osu.txt');
  if (!fs.existsSync(道)) throw new Error('証拠の 紙が 無い … docs/measured/kansuu46/cell-honmono-de-osu.txt');
  const 紙 = fs.readFileSync(道, 'utf-8');
  if (!/本物の Chrome/.test(紙)) throw new Error('本物で 押した 紙では ない');
  if (!/0 赤/.test(紙)) throw new Error('赤が 残っている 紙');
  /* ★何を 押したかを 名指しで 見る★（紙だけ 置いて 中身が 空を 止める） */
  for (const 印 of ['リボンの「桁区切り」', 'リボンの「通貨」', 'リボンの「右揃え」',
    '=CELL("format",A1) が ,0', '=CELL("format",A2) が C0', '=CELL("prefix",A3) が "',
    'book.html の 繋ぎが 効いている']) {
    if (紙.indexOf(印) < 0) throw new Error('紙に 無い … ' + 印);
  }
  const 絵 = path.join(ROOT, 'docs/measured/kansuu46/cell-honmono-de-osu.png');
  if (!fs.existsSync(絵)) throw new Error('絵が 無い（数字が 緑でも 絵を 開くまで OK に しない）');
  if (fs.statSync(絵).size < 10000) throw new Error('絵が 小さすぎる … ' + fs.statSync(絵).size + 'バイト');
});

T('★出す 名簿は lib が 正本★', () => {
  const src = fs.readFileSync(path.join(ROOT, 'lib/formula-cell-plug.js'), 'utf-8');
  const 並 = [...src.matchAll(/^\s*'([A-Z][A-Z0-9.]*)':\s*\{\s*method:/gm)].map((m) => m[1]);
  if ([...F.足した名前()].sort().join(',') !== [...並].sort().join(',')) throw new Error('名簿が 違う');
});

if (自己試験) {
  console.log('\n[self-test] ★物差しが 効いているか★');
  T('★" " の 中の ( を 数えたら 赤に なる★（実測 "("#,##0")" は 0）', () => {
    if (F.かっこがつくか('"("#,##0")"')) throw new Error('字の かたまりの 中の ( を 数えた');
    if (!F.かっこがつくか('(#,##0);(#,##0)')) throw new Error('本物の かっこを 数えていない');
  });
  T('★$ を 通貨に したら 赤に なる★（日本語の Excel では ,0）', () => {
    if (F.書式の合図('$#,##0') !== ',0') throw new Error('$ を 通貨に した');
    if (F.書式の合図('¥#,##0') !== 'C0') throw new Error('¥ を 通貨に していない');
  });
  T('★年+日（月が 無い）を D1 に したら 赤に なる★（実測 G）', () => {
    if (F.書式の合図('yyyy/d') !== 'G') throw new Error('yyyy/d を D系に した');
    if (F.書式の合図('yyyy/m/d') !== 'D1') throw new Error('yyyy/m/d が D1 で ない');
  });
  T('★時だけ／秒だけを 時刻に したら 赤に なる★（実測 G）', () => {
    for (const f of ['h', 'ss', 'mm:ss', 'h:ss', '[h]:mm']) {
      if (F.書式の合図(f) !== 'G') throw new Error(f + ' を 時刻に した');
    }
    if (F.書式の合図('h:mm') !== 'D9') throw new Error('h:mm が D9 で ない');
  });
  T('★見た目を 入れ忘れても 嘘を 言わない★（既定＝まっさらな Excel と 同じ）', () => {
    const 前 = 見た目を返す;
    見た目を返す = () => null;
    try {
      if (String(押す('=CELL("format",D1)')) !== 'G') throw new Error('表示形式が 無いのに G 以外');
      if (Number(押す('=CELL("width",B1)')) !== 8) throw new Error('幅が 無いのに 8 以外');
    } finally { 見た目を返す = 前; }
  });
}

console.log('\n' + pass + ' passed, ' + fail + ' failed');
if (直に走った) process.exit(fail ? 1 : 0);
