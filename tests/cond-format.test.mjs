/* cond-format.test.mjs — ★条件付き書式の当たり判定を 実Excelの真値と突き合わせる★
 *
 *  真値 = tests/fixtures/cond-format-golden.json
 *        （実Excel 16.0.20228 を COM で動かし、Range.DisplayFormat.Interior.Color で
 *          「そのセルが実際に塗られたか」を読んだ物。tools/measure-cond-format.ps1 で測り直せる）
 *
 *  ★真値の字をそのまま読む★＝ここに数字や答えを 手で書き写さない（写すと 真値が2つになる）。
 *
 *  使い方: node tests/cond-format.test.mjs [--self-test]
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createRequire } from 'node:module';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(__dirname, '..');
const SELF = process.argv.includes('--self-test');
const OVERRIDE = process.env.EXALLY_CF_OVERRIDE ? JSON.parse(process.env.EXALLY_CF_OVERRIDE) : {};
const require_ = createRequire(pathToFileURL(path.join(ROOT, 'package.json')));
const CF = require_(OVERRIDE['lib/cond-format.js'] || path.join(ROOT, 'lib/cond-format.js'));
const GOLD = JSON.parse(fs.readFileSync(path.join(ROOT, 'tests/fixtures/cond-format-golden.json'), 'utf8'));

let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ok   ' + n); } catch (e) { fail++; console.log('  NG   ' + n + '\n       ' + (e && e.message)); } };
const ok = (c, m) => { if (!c) throw new Error(m || 'expected truthy'); };
const eq = (a, b, m) => { if (a !== b) throw new Error((m ? m + ': ' : '') + '期待=' + JSON.stringify(b) + ' 実際=' + JSON.stringify(a)); };

/* ★実Excelで下地に置いた物と 同じ形のセル★ */
const セル = {
  '10': { v: 10 }, '20': { v: 20 }, '30': { v: 30 }, '40': { v: 40 }, '50': { v: 50 },
  '空セル': null,
  '文字『あいう』': { v: 'あいう' },
  'エラー #DIV/0!': { f: '=1/0', d: '#DIV/0!' },
};
セル['空文字を返す式 =IF(1=1,"","x")'] = { f: '=IF(1=1,"","x")', d: '' };
セル['空文字を返す式'] = セル['空文字を返す式 =IF(1=1,"","x")'];

const 真偽 = (s) => {
  const t = String(s).replace(/★/g, '');
  if (t.startsWith('当たらない')) return false;
  if (t.startsWith('当たる')) return true;
  throw new Error('真値の書き方が変わった：' + s);
};

console.log('');
console.log('[cond-format] ★実Excelの真値と 同じ答えを出すか★');
console.log('  真値 = ' + GOLD._measured_with);

/* ── ① 25より大きい ── */
{
  const 表 = GOLD['①「25 より大きい」を A1:A9 に当てた'];
  const ルール = { 種類: 'セルの値', 演算: 'より大きい', 値1: 25 };
  for (const [名, 答え] of Object.entries(表)) {
    T('① ' + 名 + ' → ' + String(答え).replace(/★/g, ''), () => {
      ok(セル[名] !== undefined, '見本のセルが無い：' + 名);
      eq(CF.当たるか(セル[名], ルール, {}), 真偽(答え));
    });
  }
}

/* ── ② 25より小さい ── */
{
  const 表 = GOLD['②「25 より小さい」を A1:A9 に当てた'];
  const ルール = { 種類: 'セルの値', 演算: 'より小さい', 値1: 25 };
  for (const [名, 答え] of Object.entries(表)) {
    T('② ' + 名 + ' → ' + String(答え).replace(/★/g, ''), () => {
      eq(CF.当たるか(セル[名], ルール, {}), 真偽(答え));
    });
  }
}

/* ── ③ "" に等しい ── */
{
  const 表 = GOLD['③「\"\" に等しい」'];
  const ルール = { 種類: 'セルの値', 演算: '等しい', 値1: '' };
  for (const [名, 答え] of Object.entries(表)) {
    T('③ 「""に等しい」 ' + 名 + ' → ' + 答え, () => {
      eq(CF.当たるか(セル[名], ルール, {}), 真偽(答え));
    });
  }
}

/* ── ④ 2本が重なった時 ── */
T('④ ★番号の小さいルールが勝つ★（>15 が先／20・40・50）', () => {
  const 並び = [
    { 種類: 'セルの値', 演算: 'より大きい', 値1: 15, 書式: { 塗り: '赤' } },
    { 種類: 'セルの値', 演算: 'より大きい', 値1: 35, 書式: { 塗り: '緑' } },
  ];
  for (const v of [20, 40, 50]) eq(CF.効く書式({ v }, 並び, {}).塗り, '赤', v + ' の勝ち色');
});
T('④ ★順番を入れ替えると 40・50 は 緑になる（20 は赤のまま）★', () => {
  const 並び = [
    { 種類: 'セルの値', 演算: 'より大きい', 値1: 35, 書式: { 塗り: '緑' } },
    { 種類: 'セルの値', 演算: 'より大きい', 値1: 15, 書式: { 塗り: '赤' } },
  ];
  eq(CF.効く書式({ v: 20 }, 並び, {}).塗り, '赤');
  eq(CF.効く書式({ v: 40 }, 並び, {}).塗り, '緑');
  eq(CF.効く書式({ v: 50 }, 並び, {}).塗り, '緑');
});
T('④ ★項目が違えば 混ざる（塗りだけの物と 文字色だけの物）★', () => {
  const 並び = [
    { 種類: 'セルの値', 演算: 'より大きい', 値1: 35, 書式: { 文字色: '青' } },
    { 種類: 'セルの値', 演算: 'より大きい', 値1: 15, 書式: { 塗り: '赤' } },
  ];
  const r = CF.効く書式({ v: 40 }, 並び, {});
  eq(r.文字色, '青'); eq(r.塗り, '赤');
});
T('④ ★どれにも当たらなければ 何も返さない（手で塗った色を消さない）★', () => {
  eq(CF.効く書式({ v: 1 }, [{ 種類: 'セルの値', 演算: 'より大きい', 値1: 35, 書式: { 塗り: '緑' } }], {}), null);
});

/* ── ⑤ 式のルール ── */
T('⑤ ★式のルールは 渡された計算エンジンで測る（決め打ちしない）★', () => {
  const ルール = { 種類: '式', 値1: '=$A1>25' };
  eq(CF.当たるか({ v: 1 }, ルール, {}), false, '★エンジンが無いのに 当てている★');
  eq(CF.当たるか({ v: 1 }, ルール, { 式で判定: () => true }), true);
  eq(CF.当たるか({ v: 1 }, ルール, { 式で判定: () => false }), false);
});
T('⑤ ★行がずれる形（=$A1>25 を A1:C3 に当てると 3行目だけ）を 呼ぶ側の判定で再現できる★', () => {
  const A = { 1: 10, 2: 20, 3: 30 };            /* 実Excelの下地と同じ */
  const ルール = { 種類: '式', 値1: '=$A1>25' };
  const 当たり = [];
  for (const 行 of [1, 2, 3]) {
    for (const 列 of ['A', 'B', 'C']) {
      const r = CF.当たるか({ v: 0 }, ルール, { 式で判定: () => A[行] > 25 });
      if (r) 当たり.push(列 + 行);
    }
  }
  eq(当たり.join(','), 'A3,B3,C3', '真値＝' + GOLD['⑤式のルール =$A1>25 を A1:C3 に当てた（下地 A1=10 A2=20 A3=30）']['3行目 A3/B3/C3']);
});

/* ── ⑥ 上位下位・重複・文字を含む ── */
T('⑥ ★上位2項目＝40と50／下位2項目＝10と20（TopBottom 0は下位）★', () => {
  const 値 = [10, 20, 30, 40, 50].map((v) => ({ v }));
  const 上 = 値.filter((c) => CF.当たるか(c, { 種類: '上位下位', 演算: '上位', 値1: 2 }, { 範囲の値たち: 値 })).map((c) => c.v);
  const 下 = 値.filter((c) => CF.当たるか(c, { 種類: '上位下位', 演算: '下位', 値1: 2 }, { 範囲の値たち: 値 })).map((c) => c.v);
  eq(上.join(','), '40,50');
  eq(下.join(','), '10,20', '真値＝' + GOLD['⑥そのほかのルール']['上位2項目（10/20/30/40/50）']);
});
T('⑥ ★重複する値＝あ と あ／一意の値＝い と う★', () => {
  const 値 = ['あ', 'い', 'あ', 'う'].map((v) => ({ v }));
  const 重 = 値.filter((c) => CF.当たるか(c, { 種類: '重複する値' }, { 範囲の値たち: 値 })).map((c) => c.v);
  const 一 = 値.filter((c) => CF.当たるか(c, { 種類: '一意の値' }, { 範囲の値たち: 値 })).map((c) => c.v);
  eq(重.join(','), 'あ,あ');
  eq(一.join(','), 'い,う');
});
T('⑥ ★「abc を含む」は 大文字小文字を区別しない・途中でもよい★', () => {
  const 値 = ['ABC', 'abc', 'xyz', 'xABCx'].map((v) => ({ v }));
  const r = 値.filter((c) => CF.当たるか(c, { 種類: '文字を含む', 値1: 'abc' }, {})).map((c) => c.v);
  eq(r.join(','), 'ABC,abc,xABCx', '真値＝' + GOLD['⑥そのほかのルール']['「abc を含む」（ABC / abc / xyz / xABCx）']);
});
T('⑥ ★空のマスは「文字を含む」に当たらない★', () => {
  eq(CF.当たるか(null, { 種類: '文字を含む', 値1: 'あ' }, {}), false);
  eq(CF.当たるか({ v: '' }, { 種類: '文字を含む', 値1: 'あ' }, {}), false);
});

/* ── ★素通りした2つを 塞ぐ（2026-08-22 自己確認で見つけた）★ ── */
T('★中身が空のセル（v:""）も 空セルと同じ扱い★', () => {
  /* ★null だけで測っていたので、v:"" の道が 検査の外に居た（壊しても赤にならなかった）★ */
  const 空2 = { v: '' };
  eq(CF.当たるか(空2, { 種類: 'セルの値', 演算: 'より大きい', 値1: 25 }, {}), false, '空が 文字扱いになっている');
  eq(CF.当たるか(空2, { 種類: 'セルの値', 演算: 'より小さい', 値1: 25 }, {}), true, '★空は 0 として比べる（実測②）★');
  eq(CF.当たるか(空2, { 種類: 'セルの値', 演算: '等しい', 値1: '' }, {}), true);
  eq(CF.読む(空2).型, '空');
});
T('★エラーは どの種類のルールにも 当たらない★', () => {
  const err = { f: '=1/0', d: '#DIV/0!' };
  const 並び = [err, err, { v: 1 }];
  eq(CF.当たるか(err, { 種類: 'セルの値', 演算: 'より大きい', 値1: 25 }, {}), false, 'セルの値');
  eq(CF.当たるか(err, { 種類: '文字を含む', 値1: 'DIV' }, {}), false, '★文字を含む＝エラーの字を 探させない★');
  eq(CF.当たるか(err, { 種類: '重複する値' }, { 範囲の値たち: 並び }), false, '★重複＝エラー2つを 重複と数えない★');
  eq(CF.当たるか(err, { 種類: '一意の値' }, { 範囲の値たち: 並び }), false, '一意');
  eq(CF.当たるか(err, { 種類: '上位下位', 演算: '上位', 値1: 2 }, { 範囲の値たち: 並び }), false, '上位');
  eq(CF.当たるか(err, { 種類: '式', 値1: '=1' }, { 式で判定: () => true }), false, '★式でも 当てない★');
  eq(CF.効く書式(err, [{ 種類: 'セルの値', 演算: 'より小さい', 値1: 25, 書式: { 塗り: '赤' } }], {}), null);
});

T('★守りは2重＝比べる 自体も エラーには「比べられない」を返す★', () => {
  /* ★片方を外しても もう片方で止まる★。だから ★どちらの道も 別々に押しておく★
     （2026-08-22 自己確認：ここを押していなかったので 外しても赤にならなかった） */
  const err = CF.読む({ f: '=1/0', d: '#DIV/0!' });
  eq(CF.比べる(err, 25), null, '★エラーを 比べてしまっている★');
  eq(CF.比べる(err, ''), null);
  eq(CF.比べる(CF.読む({ v: 10 }), 25), -1, '普通の数まで 比べられなくなっている');
});

/* ── 押す前に見せる数 ── */
T('★今 何個 当たるかを 数えられる（押す前に見せる数）★', () => {
  const 値 = [10, 20, 30, 40, 50].map((v) => ({ v }));
  eq(CF.当たる数({ 種類: 'セルの値', 演算: 'より大きい', 値1: 25 }, 値), 3);
  eq(CF.当たる数({ 種類: 'セルの値', 演算: 'より小さい', 値1: 25 }, 値), 2);
});

/* ── 見た目だけ・値に触らない ── */
T('★値も式も 1バイトも変えない（見た目だけ）★', () => {
  const c = { v: 30, f: '=A1+20', d: 30 };
  const 前 = JSON.stringify(c);
  CF.当たるか(c, { 種類: 'セルの値', 演算: 'より大きい', 値1: 25 }, {});
  CF.効く書式(c, [{ 種類: 'セルの値', 演算: 'より大きい', 値1: 25, 書式: { 塗り: '#FFC7CE' } }], {});
  eq(JSON.stringify(c), 前, '★セルを書き換えている★');
});
T('★Excelの既定の見た目（実測値）を持っている★', () => {
  eq(CF.既定の書式.塗り, '#FFC7CE');
  eq(CF.既定の書式.文字色, '#9C0006');
});

console.log('');
console.log('  ' + pass + ' 緑 / ' + fail + ' 赤');

if (SELF) {
  const { spawnSync } = await import('node:child_process');
  const os = await import('node:os');
  const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'exally-cf-'));
  console.log('');
  console.log('[self-test] わざと壊して 赤くなるかを数える（★repo は読むだけ★）');
  const 元 = fs.readFileSync(path.join(ROOT, 'lib/cond-format.js'), 'utf8');
  const BREAKS = [
    ['★文字を 数より小さい事にする★', (s) => s.replace('if (左が文字 && !右が文字) return 1;', 'if (左が文字 && !右が文字) return -1;')],
    ['★空セル(null)を 0 として扱わない★', (s) => s.replace("if (!cell) return { 型: '空', 数: 0, 文字: '' };", "if (!cell) return { 型: '文字', 数: null, 文字: '' };")],
    ['★式のエラーにも 当てる（比べる側の守りを外す）★', (s) => s.replace("    if (中身.型 === 'エラー') return null;", '')],
    ['★空セル(v は空文字)を 0 として扱わない★', (s) => s.replace("if (生 === undefined || 生 === null || 生 === '') return { 型: '空', 数: 0, 文字: '' };", "if (生 === undefined || 生 === null || 生 === '') return { 型: '文字', 数: null, 文字: '' };")],
    ['★空文字を返す式を 空セル扱いにする★', (s) => s.replace("      return { 型: '文字', 数: null, 文字: ds };", "      return { 型: '空', 数: 0, 文字: '' };")],
    ['★エラーにも 当てる★', (s) => s.replace("    if (中身.型 === 'エラー') return false;", '')],
    ['★番号の大きいルールが勝つ（後から上書き）★', (s) => s.replace('      if (出.塗り === undefined && f.塗り) 出.塗り = f.塗り;', '      if (f.塗り) 出.塗り = f.塗り;')],
    ['★当たらなくても 書式を返す★', (s) => s.replace('      if (!当たるか(cell, 並び[i], 文脈)) continue;', '')],
    ['★上位と下位を 取り違える★', (s) => s.replace("        if (ルール.演算 === '下位') {", '        if (false) {')],
    ['★重複を 1個でも当てる★', (s) => s.replace("        return ルール.種類 === '重複する値' ? 数え >= 2 : 数え === 1;", "        return ルール.種類 === '重複する値' ? 数え >= 1 : 数え === 1;")],
    ['★大文字小文字を 区別する★', (s) => s.replace('        return 中身.文字.toLowerCase().indexOf(探す.toLowerCase()) >= 0;', '        return 中身.文字.indexOf(探す) >= 0;')],
    ['★式のルールを エンジン無しで 当てる★', (s) => s.replace("        if (typeof 文脈.式で判定 !== 'function') return false;", "        if (typeof 文脈.式で判定 !== 'function') return true;")],
    ['★セルを書き換える★', (s) => s.replace('  function 当たるか(cell, ルール, 文脈) {', '  function 当たるか(cell, ルール, 文脈) {\n    if (cell) cell.v = 0;')],
    ['★既定の色を 変える★', (s) => s.replace("塗り: '#FFC7CE'", "塗り: '#FF0000'")],
  ];
  let red = 0;
  for (const [名, 壊す] of BREAKS) {
    const bad = 壊す(元);
    if (bad === 元) { console.log('  ★置換できず★  ' + 名); continue; }
    const f = path.join(TMP, 'cond-format.js');
    fs.writeFileSync(f, bad, 'utf8');
    const env = Object.assign({}, process.env, { EXALLY_CF_OVERRIDE: JSON.stringify({ 'lib/cond-format.js': f }) });
    const r = spawnSync(process.execPath, [path.join(__dirname, 'cond-format.test.mjs')], { encoding: 'utf8', env });
    if (r.status !== 0) { red++; console.log('  赤くなった  ' + 名); }
    else console.log('  ★素通り★  ' + 名);
  }
  const 今 = fs.readFileSync(path.join(ROOT, 'lib/cond-format.js'), 'utf8');
  if (今 !== 元) { console.log('  ★NG★ repo の lib/cond-format.js を書き換えている'); process.exit(1); }
  try { fs.rmSync(TMP, { recursive: true, force: true }); } catch (e) { /* 消せなくても検査は済んでいる */ }
  console.log('');
  console.log('  ' + red + '/' + BREAKS.length + ' 通りで赤くなった');
  process.exit(red === BREAKS.length ? 0 : 1);
}

process.exit(fail ? 1 : 0);
