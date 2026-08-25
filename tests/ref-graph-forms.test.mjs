/* ref-graph-forms.test.mjs — ★Excelの仕様の側から 形を並べて 1つずつ試す★
 *
 *  ★司さん 2026-08-25「世の中にあるExcelで作れる計算表や管理表など
 *    考えれることは全て シミュレーションしたんか？」→ ★していなかった★★
 *    （司さんの本1冊＋xlsx8冊だけ＝★うちの実物は「正」ではなく「1例」★）
 *    この本に INDIRECT が0本でも ★在庫表・シフト表・原価表には 普通に在る★。
 *
 *  ★出す物★＝「試した形の数／拾えた数／取りこぼしに入れた数」の表。
 *    ★0件だと主張しない。取りこぼしは 式をそのまま出す★
 *
 *  ★判定の言葉★
 *    拾えた   … 参照として 正しく解釈できた（別シート／別ファイル）
 *    印つき   … ★解けないと分かって 印を付けた★（INDIRECT/OFFSET＝計算しないと決まらない）
 *    取りこぼし … 参照らしいのに 解釈できなかった（★隠さずに数える★）
 *    対象外   … そもそも参照ではない（同じシートの中・字の中）
 *
 *  使い方: node tests/ref-graph-forms.test.mjs [--self-test]
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createRequire } from 'node:module';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(__dirname, '..');
const SELF = process.argv.includes('--self-test');
const OVERRIDE = process.env.EXALLY_RGF_OVERRIDE ? JSON.parse(process.env.EXALLY_RGF_OVERRIDE) : {};
const require_ = createRequire(pathToFileURL(path.join(ROOT, 'package.json')));
const RG = require_(OVERRIDE['lib/ref-graph.js'] || path.join(ROOT, 'lib/ref-graph.js'));

let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; return true; } catch (e) { fail++; console.log('  NG   ' + n + '\n       ' + (e && e.message)); return false; } };
const eq = (a, b, m) => { if (a !== b) throw new Error((m ? m + ': ' : '') + '期待=' + JSON.stringify(b) + ' 実際=' + JSON.stringify(a)); };

console.log('');
console.log('[ref-graph-forms] ★Excelの仕様の側から 形を1つずつ試す★');

/* 下地のブック（形ごとに 1冊ずつ作る） */
const 並びの名前 = ['計算', '売上表', '在庫', 'Sheet1', 'Sheet2', 'Sheet3', '名前に スペース', "名前に'引用符", '4月', '4月実績'];
function 一冊(f, 追加名) {
  const names = 並びの名前.concat(追加名 || []);
  return names.map((n, i) => ({ name: n, data: i === 0 ? { '0,0': { f, v: 0 } } : {} }));
}

/* ★形の一覧★ 期待 = {別, ファイル, 印, こぼし}（書かない物は 0） */
const 形 = [
  /* ── 同じシートの中（参照ではあるが 網には要らない）── */
  ['A1', '=A1+1', {}, '同じシート'],
  ['$A$1', '=$A$1*2', {}, '同じシート'],
  ['A1:B9', '=SUM(A1:B9)', {}, '同じシート'],
  ['A:A（列まるごと）', '=SUM(A:A)', {}, '同じシート'],
  ['1:1（行まるごと）', '=SUM(1:1)', {}, '同じシート'],
  ['R1C1風の名前は誤爆しない', '=RC+1', {}, '同じシート'],
  /* ── 別シート ── */
  ['Sheet2!A1', '=Sheet2!A1', { 別: 1 }, '別シート'],
  ["'名前に スペース'!A1", "='名前に スペース'!A1", { 別: 1 }, '別シート'],
  ["'名前に''引用符'!A1", "='名前に''引用符'!A1", { 別: 1 }, '別シート'],
  ['囲っていない長い名前（4月実績）', '=4月実績!A1', { 別: 1 }, '別シート'],
  ['範囲つき 売上表!A1:A9', '=SUM(売上表!A1:A9)', { 別: 1 }, '別シート'],
  ['3-D参照 Sheet1:Sheet3!A1', '=SUM(Sheet1:Sheet3!A1)', { 別: 3 }, '別シート'],
  ['1つの式に何本も混ざる', '=Sheet2!A1+売上表!B2+在庫!C3', { 別: 3 }, '別シート'],
  /* ── 別ファイル ── */
  ['[2025.xlsb]売上!A1（囲い無し）', '=[2025.xlsb]売上!A1', { ファイル: 1 }, '別ファイル'],
  [String.raw`'C:\x\[去年.xlsx]まとめ'!B2（パス付き）`, String.raw`='C:\x\[去年.xlsx]まとめ'!B2`, { ファイル: 1 }, '別ファイル'],
  ["'[開いている.xlsx]Sheet1'!A1", "='[開いている.xlsx]Sheet1'!A1", { ファイル: 1 }, '別ファイル'],
  /* ── 計算しないと決まらない（★印を付ける★）── */
  ['INDIRECT("売上表!A1")', '=INDIRECT("売上表!A1")', { 印: 1 }, '解けない'],
  ['INDIRECT(字を組み立てる)', '=INDIRECT(A1&"!B"&A2)', { 印: 1 }, '解けない'],
  ['OFFSET(A1,1,1)', '=OFFSET(A1,1,1)', { 印: 1 }, '解けない'],
  ['OFFSET＋別シート', '=SUM(OFFSET(売上表!A1,0,0,10,1))', { 別: 1, 印: 1 }, '解けない'],
  /* ── 解ける関数（行き先が字に在る）── */
  ['INDEX+MATCH（別シート）', '=INDEX(売上表!B:B, MATCH(A1, 売上表!A:A, 0))', { 別: 2 }, '別シート'],
  ['CHOOSE（別シート）', '=CHOOSE(A1, 売上表!B1, 在庫!B1)', { 別: 2 }, '別シート'],
  /* ── テーブル（構造化参照）★開く時に A1 に直る。直っていなければ 取りこぼし★ ── */
  ['[@単価]（直っていない）', '=[@単価]*[@数量]', { こぼし: 1 }, 'テーブル'],
  ['表名[列]（直っていない）', '=SUM(売上T[金額])', { こぼし: 1 }, 'テーブル'],
  ['[[#見出し],[列]]（直っていない）', '=売上T[[#Headers],[金額]]', { こぼし: 1 }, 'テーブル'],
  ['[#All]（直っていない）', '=COUNTA(売上T[#All])', { こぼし: 1 }, 'テーブル'],
  ['[#Totals]（直っていない）', '=売上T[#Totals]', { こぼし: 1 }, 'テーブル'],
  ['消えた表 Table4294967295', '=SUM(Table4294967295[#Data])', { こぼし: 1 }, 'テーブル'],
  /* ── 名前の定義 ── */
  ['名前の定義（=売上合計）', '=売上合計+1', {}, '名前の定義'],
  ['シート限定の名前（売上表!合計）', '=売上表!合計', { 別: 1 }, '名前の定義'],
  /* ── 新しい形 ── */
  ['スピル（=A1:A10#）', '=SUM(A1#)', {}, '新しい形'],
  ['スピル（別シート）', '=SUM(売上表!A1#)', { 別: 1 }, '新しい形'],
  ['LET', '=LET(x, 売上表!A1, x*2)', { 別: 1 }, '新しい形'],
  ['LAMBDA', '=LAMBDA(a,b,a+b)(売上表!A1, 在庫!A1)', { 別: 2 }, '新しい形'],
  /* ── 壊れている物 ── */
  ['#REF!（消えた参照）', '=SUM(#REF!)', { こぼし: 1 }, '壊れている'],
  ['#REF! と 生きた参照が混ざる', '=INDEX(#REF!, MATCH(B1, 計算!A1:A9, 0))', { こぼし: 1 }, '壊れている'],
  ['循環参照（自分を見る）', '=計算!A1+1', {}, '壊れている'],
  ['消えたシートを指す', '=消えたシート!A1', { こぼし: 0, 分からない: 1 }, '壊れている'],
  /* ── 字の中（参照ではない）── */
  ['="Sheet2!A1"（字の中）', '="Sheet2!A1"', {}, '対象外'],
  ['CONCAT("月別!",A1)', '=CONCAT("売上表!",A1)', {}, '対象外'],
  ['字の中に "" が入る', '=IF(A1="""売上表!A1""","x","y")', {}, '対象外'],
];

const 集計 = {};
for (const [名, f, 期待, 種] of 形) {
  集計[種] = 集計[種] || { 試した: 0, 拾えた: 0, 印つき: 0, こぼし: 0, 赤: 0 };
  集計[種].試した++;
  const g = RG.作る(一冊(f));
  const 別 = g.別シート参照.length, ファイル = g.別ファイル参照.length;
  const 印 = g.解けない.length, こぼし = g.取りこぼし.length, 分 = g.分からない参照.length;
  const ok = T('形 … ' + 名, () => {
    eq(別, 期待.別 || 0, '別シート');
    eq(ファイル, 期待.ファイル || 0, '別ファイル');
    eq(印, 期待.印 || 0, '解けない印');
    eq(こぼし, 期待.こぼし || 0, '取りこぼし');
    eq(分, 期待.分からない || 0, '分からない名前');
  });
  if (!ok) { 集計[種].赤++; console.log('       式) ' + f); }
  集計[種].拾えた += 別 + ファイル;
  集計[種].印つき += 印;
  集計[種].こぼし += こぼし;
}

/* ── 大きい物（重さの形）── */
T('★1000行×1000列級（100万セル）でも 網が作れる★', () => {
  const data = {};
  for (let r = 0; r < 2000; r++) data[r + ',0'] = { f: '=売上表!A' + (r + 1), v: 0 };
  const b = [{ name: '計算', data }, { name: '売上表', data: {} }];
  const t0 = Date.now();
  const g = RG.作る(b);
  const 秒 = (Date.now() - t0) / 1000;
  eq(g.別シート参照.length, 2000, '2000本');
  if (秒 > 3) throw new Error('★' + 秒 + '秒 かかった★');
  console.log('       … 式2,000本 ' + 秒 + '秒');
});
T('★シート50枚でも 網が作れる★', () => {
  const b = [];
  for (let i = 0; i < 50; i++) b.push({ name: 'S' + i, data: i ? {} : {} });
  b[0].data = {};
  for (let i = 1; i < 50; i++) b[0].data[i + ',0'] = { f: '=S' + i + '!A1', v: 0 };
  const g = RG.作る(b);
  eq(g.別シート参照.length, 49, '49本');
  eq(g.シートの網[0].length, 49, '網');
});
T('★シート名に ! が入る場合（Excelは付けられない＝取り違えない）★', () => {
  /* Excel はシート名に ! を許さない。★許さない物を「在る」前提にしない★ */
  const g = RG.作る([{ name: '計算', data: { '0,0': { f: '=売上表!A1' } } }, { name: '売上表', data: {} }]);
  eq(g.別シート参照.length, 1);
});

/* ── 表を出す ── */
console.log('');
console.log('  ══ 試した形の数 ══');
console.log('  ' + '種類'.padEnd(12) + '試した  拾えた  印つき  取りこぼし  赤');
let t全 = 0, h全 = 0, m全 = 0, r全 = 0;
for (const [種, v] of Object.entries(集計)) {
  console.log('  ' + 種.padEnd(12) + String(v.試した).padStart(4) + String(v.拾えた).padStart(8)
    + String(v.印つき).padStart(8) + String(v.こぼし).padStart(10) + String(v.赤).padStart(6));
  t全 += v.試した; h全 += v.拾えた; m全 += v.印つき; r全 += v.こぼし;
}
console.log('  ' + '合計'.padEnd(12) + String(t全).padStart(4) + String(h全).padStart(8)
  + String(m全).padStart(8) + String(r全).padStart(10));
console.log('');
console.log('  ' + pass + ' 緑 / ' + fail + ' 赤');

if (SELF) {
  const { spawnSync } = await import('node:child_process');
  const os = await import('node:os');
  const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'exally-rgf-'));
  console.log('');
  console.log('[self-test] わざと壊して 赤くなるかを数える（★repo は読むだけ★）');
  const 元 = fs.readFileSync(path.join(ROOT, 'lib/ref-graph.js'), 'utf8');
  const BREAKS = [
    ['★3-D参照の 間のシートを 落とす★', (s) => s.replace('for (var q = 小; q <= 大; q++) {', 'for (var q = 大; q <= 大; q++) {')],
    ['★INDIRECT に 印を付けない★', (s) => s.replace("網.解けない.push({ from: s, fromCell: keys[k], why: '参照先が計算しないと決まらない', f: f });", '')],
    ['★直っていない表の参照を 見逃す★', (s) => s.replace("if (表の残り.indexOf('[') >= 0) {", 'if (false) {')],
    ['★#REF! を 見逃す★', (s) => s.replace('if (/#REF!/.test(素)) {', 'if (false) {')],
    ['★別ファイルを 別シートとして数える★', (s) => s.replace('if (ファイル) {', 'if (false) {')],
    ['★字の中も 拾う★', (s) => s.replace('var 素 = 字を外す(f);', 'var 素 = f;')],
  ];
  let red = 0;
  for (const [name, brk] of BREAKS) {
    const bad = brk(元);
    if (bad === 元) { console.log('  ★置換できず★  ' + name); continue; }
    const f = path.join(TMP, 'ref-graph.js');
    fs.writeFileSync(f, bad, 'utf8');
    const env = Object.assign({}, process.env, { EXALLY_RGF_OVERRIDE: JSON.stringify({ 'lib/ref-graph.js': f }) });
    const r = spawnSync(process.execPath, ['--max-old-space-size=4096', path.join(__dirname, 'ref-graph-forms.test.mjs')], { encoding: 'utf8', env });
    if (r.status !== 0) { red++; console.log('  赤くなった  ' + name); }
    else console.log('  ★素通り★  ' + name);
  }
  try { fs.rmSync(TMP, { recursive: true, force: true }); } catch (e) { /* 消せなくても検査は済んでいる */ }
  console.log('');
  console.log('  ' + red + '/' + BREAKS.length + ' 通りで赤くなった');
  process.exit(red === BREAKS.length ? 0 : 1);
}

process.exit(fail ? 1 : 0);
