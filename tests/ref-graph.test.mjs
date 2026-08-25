/* ref-graph.test.mjs — ★ブック全体の参照の網を、実物で 1本も落とさず拾えるか★
 *
 *  ★指示役の検証要件（2026-08-25）★
 *    ①開いてから ◯秒以内に 全シートの参照の網が出来る（秒数を報告に書く）
 *    ②その間 画面が固まらない（★1回の止まり時間★を数える）
 *    ③★別シート参照 377本（08-22 指示役の実測）を 1本も落とさず拾える★
 *    ④別ファイル参照も拾える
 *
 *  ★数え方は2通り在る。混ぜない★
 *    ①のべ本数 … 式1本が別シートを3か所 見ていたら 3本
 *    ②式の本数 … 上を1本と数える  ←★指示役の「377本」は こちら★
 *
 *  真値 = tests/fixtures/ref-graph-golden.json（★司さんの実物2本から取った★）
 *  使い方: node tests/ref-graph.test.mjs [--self-test]
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createRequire } from 'node:module';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(__dirname, '..');
const SELF = process.argv.includes('--self-test');
const OVERRIDE = process.env.EXALLY_RG_OVERRIDE ? JSON.parse(process.env.EXALLY_RG_OVERRIDE) : {};
const require_ = createRequire(pathToFileURL(path.join(ROOT, 'package.json')));
const RG = require_(OVERRIDE['lib/ref-graph.js'] || path.join(ROOT, 'lib/ref-graph.js'));
const GOLD = JSON.parse(fs.readFileSync(path.join(ROOT, 'tests/fixtures/ref-graph-golden.json'), 'utf8'));

let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ok   ' + n); } catch (e) { fail++; console.log('  NG   ' + n + '\n       ' + (e && e.message)); } };
const ok = (c, m) => { if (!c) throw new Error(m || 'expected truthy'); };
const eq = (a, b, m) => { if (a !== b) throw new Error((m ? m + ': ' : '') + '期待=' + JSON.stringify(b) + ' 実際=' + JSON.stringify(a)); };

console.log('');
console.log('[ref-graph] ★ブック全体の参照の網★');

/* ── 作り物で 形を固める（実物が無い機械でも これは必ず走る）── */

function ブック(sheets) { return sheets.map((s) => ({ name: s[0], data: s[1] })); }
const F = (f) => ({ f, v: 0 });

T('★別シートを見ている式を 拾う（同じシートの中は数えない）★', () => {
  const g = RG.作る(ブック([
    ['計算', { '0,0': F('=SUM(A1:A9)'), '1,0': F('=売上表!B2'), '2,0': F('=売上表!B3+計算!A1') }],
    ['売上表', { '0,1': F('=1'), '0,2': F('=月別!C1') }],
    ['月別', { '0,2': { v: 5 } }],
  ]));
  eq(g.別シート参照.length, 3, 'のべ本数');
  eq(new Set(g.別シート参照.map((r) => r.from + ':' + r.fromCell)).size, 3, '式の本数');
  eq(g.シートの網[0].join(','), '1', '計算→売上表');
  eq(g.シートの網[1].join(','), '2', '売上表→月別');
});
T('★空白や記号の入った名前も 拾う（\'4月 実績\'! ／ \'It\'\'s\'!）★', () => {
  const g = RG.作る(ブック([
    ['まとめ', { '0,0': F("='4月 実績'!E14"), '1,0': F("='It''s'!A1") }],
    ['4月 実績', { '0,0': { v: 1 } }],
    ["It's", { '0,0': { v: 2 } }],
  ]));
  eq(g.別シート参照.length, 2, '2本とも拾う');
  eq(g.分からない参照.length, 0, '名前を取り違えていない');
});
T('★長い名前を先に当てる（「4月」と「4月 実績」を取り違えない）★', () => {
  const g = RG.作る(ブック([
    ['まとめ', { '0,0': F("='4月 実績'!E14") }],
    ['4月', { '0,0': { v: 1 } }],
    ['4月 実績', { '0,0': { v: 2 } }],
  ]));
  eq(g.別シート参照.length, 1);
  eq(g.別シート参照[0].to, 2, '★短い方(4月)に当ててしまっている★');
});
T('★「長い名前を先に」は 念のため（本当に効いているのは 後ろの ! ）★', () => {
  /* ★自己確認で分かった事（2026-08-25）★
     sort を外しても 答えは変わらなかった。理由＝★名前の後ろに ! が要る★ので、
     「4月」で当てようとしても 次が「実績」で ! ではなく、機械が自分で戻って「4月実績」を当てる。
     ⇒ ★sort は「念のため」であって 唯一の守りではない★。
     ★だから これを壊す形の自己確認は 置かない（壊しても赤にならない＝嘘の守りになる）★ */
  const 当て = RG.当て方 ? null : null;   /* 当て方は外に出していない（中で使うだけ） */
  const g = RG.作る(ブック([
    ['まとめ', { '0,0': F('=4月実績!E14'), '1,0': F('=4月!A1') }],
    ['4月', { '0,0': { v: 1 } }],
    ['4月実績', { '0,0': { v: 2 } }],
  ]));
  eq(g.別シート参照.length, 2, '2本とも拾う');
  eq(g.別シート参照.map((r) => r.to).sort().join(','), '1,2', '★どちらかに寄せている★');
});
T('★囲っていない長い名前も 取り違えない（4月 と 4月実績）★', () => {
  /* ★'…' で囲むと ①の枝で当たるので sort が効かない★＝
     囲っていない形で試さないと「長い方を先に」の検査にならない（自己確認で素通りした） */
  const g = RG.作る(ブック([
    ['まとめ', { '0,0': F('=4月実績!E14') }],
    ['4月', { '0,0': { v: 1 } }],
    ['4月実績', { '0,0': { v: 2 } }],
  ]));
  eq(g.別シート参照.length, 1, '1本');
  eq(g.別シート参照[0].to, 2, '★短い方(4月)に当ててしまっている★');
  eq(g.分からない参照.length, 0, '拾い損ねている');
});
T('★字の中は 参照ではない（="売上表!A1" を数えない）★', () => {
  const g = RG.作る(ブック([
    ['計算', { '0,0': F('="売上表!A1"'), '1,0': F('=CONCAT("月別!",A1)') }],
    ['売上表', {}], ['月別', {}],
  ]));
  eq(g.別シート参照.length, 0, '★字の中を拾っている★');
});
T('★別ファイル参照も 拾う（[Book.xlsx]Sheet!A1 ／ フルパス付き）★', () => {
  const g = RG.作る(ブック([
    ['計算', { '0,0': F('=[2025.xlsb]売上!A1'), '1,0': F("='C:\\\\x\\\\[去年.xlsx]まとめ'!B2") }],
  ]));
  eq(g.別ファイル参照.length, 2, '2本とも拾う');
  eq(g.別ファイル参照[0].ファイル, '2025.xlsb');
  eq(g.別ファイル参照[1].ファイル, '去年.xlsx');
});
T('★知らない名前は 捨てずに「分からない」に置く★', () => {
  const g = RG.作る(ブック([['計算', { '0,0': F('=消えたシート!A1') }]]));
  eq(g.別シート参照.length, 0);
  eq(g.分からない参照.length, 1, '★黙って捨てている★');
});
T('★間接の参照も たどれる（A→B→C）★', () => {
  const g = RG.作る(ブック([
    ['A', { '0,0': F('=B!A1') }], ['B', { '0,0': F('=C!A1') }], ['C', { '0,0': { v: 1 } }],
  ]));
  eq(RG.たどる(g, 0).sort().join(','), '1,2', '★C まで辿れていない＝合計が古いまま残る★');
});
T('★輪になっていても 止まる（A→B→A）★', () => {
  const g = RG.作る(ブック([['A', { '0,0': F('=B!A1') }], ['B', { '0,0': F('=A!A1') }]]));
  eq(RG.たどる(g, 0).join(','), '1');
});
T('★小分けにしても 答えが同じ★', () => {
  const b = ブック([
    ['計算', { '0,0': F('=売上表!A1'), '1,0': F('=売上表!A2'), '2,0': F('=月別!A1') }],
    ['売上表', {}], ['月別', {}],
  ]);
  const 一気 = RG.作る(b);
  const it = RG.作る途中(b, { 一度に: 1 });
  let r = it.next(), 回 = 0;
  while (!r.done) { 回++; r = it.next(); }
  eq(r.value.別シート参照.length, 一気.別シート参照.length, '本数が違う');
  ok(回 >= 1, '★1回も分けていない＝画面が固まる★');
});

/* ── ★形の一覧を 1つずつ（指示役 2026-08-25 の指定）★ ── */
{
  const b = ブック([
    ['計算', {}], ['売上表', {}], ['名前に スペース', {}], ["名前に'引用符", {}], ['Sheet2', {}],
  ]);
  const 表 = [
    ['A1（同じシート）', '=A1+1', { 別: 0 }],
    ['$A$1（同じシート）', '=$A$1', { 別: 0 }],
    ['A1:B9（同じシート）', '=SUM(A1:B9)', { 別: 0 }],
    ['A:A（列まるごと）', '=SUM(A:A)', { 別: 0 }],
    ['1:1（行まるごと）', '=SUM(1:1)', { 別: 0 }],
    ['Sheet2!A1', '=Sheet2!A1', { 別: 1 }],
    ["'名前に スペース'!A1", "='名前に スペース'!A1", { 別: 1 }],
    ["'名前に''引用符'!A1", "='名前に''引用符'!A1", { 別: 1 }],
    ['[2025.xlsb]売上!A1（囲っていない別ファイル）', '=[2025.xlsb]売上!A1', { 別: 0, ファイル: 1 }],
    ['Sheet2:Sheet3!A1（3-D参照・端が一覧に無い→取りこぼし）', '=SUM(Sheet2:Sheet3!A1)', { 別: 0, 取りこぼし: 1 }],
    [String.raw`'C:\x\[去年.xlsx]まとめ'!B2（囲った別ファイル）`, String.raw`='C:\x\[去年.xlsx]まとめ'!B2`, { 別: 0, ファイル: 1 }],
    ['INDIRECT("売上表!A1")', '=INDIRECT("売上表!A1")', { 別: 0, 解けない: 1 }],
    ['OFFSET(A1,1,1)', '=OFFSET(A1,1,1)', { 別: 0, 解けない: 1 }],
    ['INDEX(売上表!A1:A9,2)', '=INDEX(売上表!A1:A9,2)', { 別: 1 }],
    ['#REF!（消えた参照）', '=INDEX(#REF!, MATCH(B1, 計算!A1:A9, 0))', { 別: 0, 取りこぼし: 1 }],
    ['表の参照が残っている（直せなかった）', '=SUM(Table1[#Data])', { 別: 0, 取りこぼし: 1 }],
    ['1つの式に何本も混ざる', "=Sheet2!A1+'名前に スペース'!B2+売上表!C3", { 別: 3 }],
    ['字の中（参照ではない）', '="Sheet2!A1"', { 別: 0 }],
  ];
  for (const [名, f, 期待] of 表) {
    T('形 … ' + 名, () => {
      const bb = ブック([['計算', { '0,0': F(f) }], ['売上表', {}], ['名前に スペース', {}], ["名前に'引用符", {}], ['Sheet2', {}]]);
      const g = RG.作る(bb);
      eq(g.別シート参照.length, 期待.別 || 0, '別シート参照');
      eq(g.別ファイル参照.length, 期待.ファイル || 0, '別ファイル参照');
      eq(g.解けない.length, 期待.解けない || 0, '★解けない（印を付ける）★');
      eq(g.取りこぼし.length, 期待.取りこぼし || 0, '★取りこぼし（隠さない）★');
    });
  }
}
T('★3-D参照（Sheet1:Sheet3!A1）は 間のシートも全部 数える★', () => {
  /* ★2026-08-25 実測で見つけた穴★＝前は 後ろの Sheet3 だけ拾って ★Sheet1 を黙って落としていた★ */
  const g = RG.作る(ブック([
    ['計算', { '0,0': F('=SUM(Sheet1:Sheet3!A1)') }], ['Sheet1', {}], ['Sheet2', {}], ['Sheet3', {}],
  ]));
  eq(g.別シート参照.map((r) => r.to).sort().join(','), '1,2,3', '★間のシート(Sheet2)を落としている★');
  eq(g.取りこぼし.length, 0);
});
T('★3-D参照の端が分からない時は 取りこぼしに置く（黙って捨てない）★', () => {
  const g = RG.作る(ブック([['計算', { '0,0': F('=SUM(消えた1:消えた2!A1)') }]]));
  eq(g.取りこぼし.length, 1, '★黙って捨てている★');
});

/* ── 実物（★在る時だけ・無ければ 未測定と書く★）── */
const 実物 = GOLD.本.map((b) => ({ ...b, 在る: fs.existsSync(b.場所) }));
let XLSX = null;
try { XLSX = require_(path.join(ROOT, 'lib/xlsx.full.min.js')); } catch (e) { /* 無ければ下で未測定 */ }

for (const b of 実物) {
  if (!b.在る || !XLSX) {
    console.log('  ★未測定★ ' + b.名 + ' … ' + (!b.在る ? '実物がこの機械に無い' : '読み込み部品が無い') + '（0件・異常なしにしない）');
    continue;
  }
  T('★実物 ' + b.名 + '：別シートを見ている式 ' + b.式の本数 + '本を 1本も落とさない★', () => {
    const wb = XLSX.read(fs.readFileSync(b.場所), { type: 'buffer', cellFormula: true, sheetStubs: false });
    const sheets = wb.SheetNames.map((name) => {
      const ws = wb.Sheets[name] || {};
      const data = {};
      for (const k of Object.keys(ws)) {
        if (k[0] === '!') continue;
        const c = ws[k];
        const a = XLSX.utils.decode_cell(k);
        data[a.r + ',' + a.c] = { v: c.v, f: c.f ? '=' + c.f : undefined };
      }
      return { name, data };
    });
    const t0 = Date.now();
    const g = RG.作る(sheets);
    const 秒 = (Date.now() - t0) / 1000;
    eq(g.シート数, b.シート数, 'シート数');
    eq(g.式の数, b.式の数, '式の数');
    eq(g.別シート参照.length, b.のべ本数, '★のべ本数★');
    eq(new Set(g.別シート参照.map((r) => r.from + ':' + r.fromCell)).size, b.式の本数, '★式の本数★');
    eq(g.分からない参照.length, 0, '★名前を取り違えている★');
    ok(秒 <= b.何秒以内, '★' + 秒 + '秒 かかった（' + b.何秒以内 + '秒以内のはず）★');
    console.log('       … ' + g.シート数 + 'シート／式' + g.式の数 + '本／別シート のべ'
      + g.別シート参照.length + '本・式' + b.式の本数 + '本／★' + 秒 + '秒★');
  });
  T('★実物 ' + b.名 + '：小分けの1回が ' + b.止まり上限ミリ秒 + 'ミリ秒 以内（画面が固まらない）★', () => {
    const wb = XLSX.read(fs.readFileSync(b.場所), { type: 'buffer', cellFormula: true, sheetStubs: false });
    const sheets = wb.SheetNames.map((name) => {
      const ws = wb.Sheets[name] || {};
      const data = {};
      for (const k of Object.keys(ws)) {
        if (k[0] === '!') continue;
        const c = ws[k];
        const a = XLSX.utils.decode_cell(k);
        data[a.r + ',' + a.c] = { v: c.v, f: c.f ? '=' + c.f : undefined };
      }
      return { name, data };
    });
    const it = RG.作る途中(sheets, { 一度に: 3000 });
    let 最長 = 0, 前 = Date.now(), r = it.next(), 回 = 0;
    while (!r.done) { const d = Date.now() - 前; if (d > 最長) 最長 = d; 前 = Date.now(); 回++; r = it.next(); }
    ok(回 >= 2, '★1回も分けていない（' + 回 + '回）＝画面が固まる★');
    ok(最長 <= b.止まり上限ミリ秒, '★1回で ' + 最長 + 'ミリ秒 止まった★');
    console.log('       … ' + 回 + '回に分けた／1回の最長 ' + 最長 + 'ミリ秒');
  });
}

console.log('');
console.log('  ' + pass + ' 緑 / ' + fail + ' 赤');

if (SELF) {
  const { spawnSync } = await import('node:child_process');
  const os = await import('node:os');
  const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'exally-rg-'));
  console.log('');
  console.log('[self-test] わざと壊して 赤くなるかを数える（★repo は読むだけ★）');
  const 元 = fs.readFileSync(path.join(ROOT, 'lib/ref-graph.js'), 'utf8');
  const BREAKS = [
    ['★同じシートの中も 数えてしまう★', (s) => s.replace('if (j === s) continue;', '')],
    ['★字の中も 参照として拾う★', (s) => s.replace('var 素 = 字を外す(f);', 'var 素 = f;')],
    ['★別ファイル参照を 捨てる★', (s) => s.replace('網.別ファイル参照.push', 'void')],
    ['★知らない名前を 黙って捨てる★', (s) => s.replace('網.分からない参照.push({ from: s, fromCell: keys[k], 名: 名 }); continue;', 'continue;')],
    ['★間接の参照を たどらない★', (s) => s.replace('見た[j] = 1; 出.push(j); 積.push(j);', '見た[j] = 1; 出.push(j);')],
    ['★小分けにしない（画面が固まる）★', (s) => s.split('if (網.読んだセル % 一度に === 0) yield { 進み: 網.読んだセル };').join('')],
    ["★'' を ' に直さない（名前を取り違える）★", (s) => s.replace(".replace(/''/g, \"'\")", '')],
  ];
  let red = 0;
  for (const [name, brk] of BREAKS) {
    const bad = brk(元);
    if (bad === 元) { console.log('  ★置換できず★  ' + name); continue; }
    const f = path.join(TMP, 'ref-graph.js');
    fs.writeFileSync(f, bad, 'utf8');
    const env = Object.assign({}, process.env, { EXALLY_RG_OVERRIDE: JSON.stringify({ 'lib/ref-graph.js': f }) });
    const r = spawnSync(process.execPath, ['--max-old-space-size=4096', path.join(__dirname, 'ref-graph.test.mjs')], { encoding: 'utf8', env });
    if (r.status !== 0) { red++; console.log('  赤くなった  ' + name); }
    else console.log('  ★素通り★  ' + name);
  }
  if (fs.readFileSync(path.join(ROOT, 'lib/ref-graph.js'), 'utf8') !== 元) {
    console.log('  ★NG★ repo の lib/ref-graph.js を書き換えている'); process.exit(1);
  }
  try { fs.rmSync(TMP, { recursive: true, force: true }); } catch (e) { /* 消せなくても検査は済んでいる */ }
  console.log('');
  console.log('  ' + red + '/' + BREAKS.length + ' 通りで赤くなった');
  process.exit(red === BREAKS.length ? 0 : 1);
}

process.exit(fail ? 1 : 0);
