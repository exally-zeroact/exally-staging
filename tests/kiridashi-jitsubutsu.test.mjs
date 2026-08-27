/* kiridashi-jitsubutsu.test.mjs — ★司さんの実物で 1人分×何月分を 切り出す★
 * =============================================================================
 * ★司さんの言葉（2026-08-27）★
 *   「今 同じシートに複数人 あるやろ？ ワンクリックで 1人分と 何月分
 *     （1日からの分／10日からの分）って 分けて表示もさせれるん？」
 *
 * ★指示役の検証要件（2026-08-27）★
 *   ・★司さんの実物（給料表 468行×131列）で 実際に切り出す★
 *   ・★「白石正人の 1月分」★が 1列×その期間だけ 出る
 *   ・★「10日からの分」★も 出る（★締め期間の決まりを 使う★）
 *   ・★切り出した数が 元と合う★（★紙・画面に描かれた文字を足す★／★中の値で閉じない★）
 *
 * ★数え方（2つの道で 数えて 突き合わせる）★
 *   ①切り出した表の ★描かれた字★ を 1行ずつ 足す（客が見る物）
 *   ②元の表から ★別の道で★ 数え直す（うちの切り出しを 使わない）
 *   ⇒ ★同じにならなければ 赤★
 *
 * 使い方: node tests/kiridashi-jitsubutsu.test.mjs [--self-test]
 */
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createRequire } from 'node:module';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(__dirname, '..');
const SELF = process.argv.includes('--self-test');
const OVERRIDE = process.env.EXALLY_KIRIDASHI_OVERRIDE ? JSON.parse(process.env.EXALLY_KIRIDASHI_OVERRIDE) : {};
const require_ = createRequire(pathToFileURL(path.join(ROOT, 'package.json')));
const Recipe = require_(OVERRIDE['lib/recipe.js'] || path.join(ROOT, 'lib/recipe.js'));
const Kikan = require_(path.join(ROOT, 'lib/kikan.js'));
const GOLDEN = JSON.parse(fs.readFileSync(path.join(ROOT, 'tests/fixtures/shindan-golden.json'), 'utf8'));

let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ok   ' + n); } catch (e) { fail++; console.log('  NG   ' + n + '\n       ' + (e && e.message)); } };
const ok = (c, m) => { if (!c) throw new Error(m || 'expected truthy'); };
const eq = (a, b, m) => { if (a !== b) throw new Error((m ? m + ': ' : '') + '期待=' + JSON.stringify(b) + ' 実際=' + JSON.stringify(a)); };

console.log('');
console.log('[切り出し・実物] ★司さんの給料表から 1人分×何月分★');

const 実物 = GOLDEN.本.場所;
if (!fs.existsSync(実物)) {
  console.log('  ★未測定★ 実物が無い機械です（0件・異常なしにしない）');
  console.log('');
  console.log('  ' + pass + ' 緑 / ' + fail + ' 赤');
  process.exit(fail ? 1 : 0);
}

const XLSX = require_(path.join(ROOT, 'lib/xlsx.full.min.js'));
const ZipSurgeon = require_(path.join(ROOT, 'lib/zip-surgeon.js'));
const TableRefs = require_(path.join(ROOT, 'lib/table-refs.js'));

const bytes = new Uint8Array(fs.readFileSync(実物));
const wb = XLSX.read(bytes, { type: 'array', cellFormula: true });
const rr = await TableRefs.resolve(bytes, 'xlsb', wb, ZipSurgeon);
const fixes = (rr && rr.fixes) || {};
const sheets = wb.SheetNames.map((name) => {
  const ws = wb.Sheets[name] || {}; const data = {};
  for (const k of Object.keys(ws)) {
    if (k[0] === '!') continue;
    const a = XLSX.utils.decode_cell(k); const c = ws[k];
    const fx = fixes[name + '|' + a.r + ',' + a.c];
    data[a.r + ',' + a.c] = { v: c.v, f: fx !== undefined ? fx : (c.f ? '=' + c.f : undefined), d: c.w, numFmt: c.z };
  }
  let 枠 = null;
  try { if (ws['!ref']) { const g = XLSX.utils.decode_range(ws['!ref']); 枠 = { 行数: g.e.r + 1, 列数: g.e.c + 1 }; } } catch (e) { /* 無ければ 無いまま */ }
  return { name, data, 枠 };
});
const 給料表 = sheets.find((s) => s.name === '給料表');
ok(給料表, '★給料表が 無い★');

/* ★別の道で 数え直す（うちの切り出しを 使わない）★ */
function 別の道で数える(人, 期間) {
  const data = 給料表.data;
  /* 見出しの行を 自分で探す（2行目のはず）＝★同じ道具を 使わない★ */
  let 見出しの行 = -1, 人の列 = -1;
  for (let r = 0; r < 5 && 見出しの行 < 0; r++) {
    for (let c = 0; c < 200; c++) {
      const cell = data[r + ',' + c];
      if (cell && String(cell.v).trim() === 人) { 見出しの行 = r; 人の列 = c; break; }
    }
  }
  let 数 = 0;
  const 日付たち = [];
  for (let r = 見出しの行 + 1; r < 1000; r++) {
    const a = data[r + ',0'];
    if (!a || typeof a.v !== 'number') continue;
    /* ★通し番号を 自分で 日付に直す★（うちの kikan を 使わない） */
    const d = new Date(Date.UTC(1899, 11, 30) + Math.floor(a.v) * 86400000);
    const ymd = d.toISOString().slice(0, 10);
    if (ymd >= 期間.from && ymd <= 期間.to) { 数++; 日付たち.push(ymd); }
  }
  return { 見出しの行, 人の列, 数, 日付たち };
}

/* ★描かれた字を 1行ずつ 足す★（客が見る物・中の値で閉じない） */
function 描かれた字を足す(出) {
  const 行 = [];
  for (let i = 1; i <= 出.何行; i++) {
    const a = 出.data[i + ',0'], b = 出.data[i + ',1'];
    行.push([
      String(a && a.v !== undefined ? a.v : ''),
      String(b && b.v !== undefined ? b.v : ''),
    ].join('\t'));
  }
  const 字 = 行.join('\n');
  return {
    行数: 行.length,
    字: 字,
    /* 描かれた字から 金額だけを 拾って 足す（★中の値を そのまま足さない★） */
    合計: 行.reduce((a, ln) => {
      const t = ln.split('\t')[1].replace(/[,¥￥\s]/g, '');
      const n = Number(t);
      return a + (isFinite(n) ? n : 0);
    }, 0),
  };
}

T('★実物の 給料表は 人が 横に並んでいる（見出しは2行目）★', () => {
  const 見 = 別の道で数える('白石正人', { from: '2026-01-01', to: '2026-01-31' });
  eq(見.見出しの行, 1, '★見出しが 2行目でない★');
  eq(見.人の列, 1, '★白石正人が B列でない★');
  console.log('       … 表の枠 ' + 給料表.枠.行数 + '行×' + 給料表.枠.列数 + '列');
});

T('★「白石正人の 1月分（1日から）」が 出る★', () => {
  const 出 = Recipe.切り出す(給料表, { 人: '白石正人', 月: '2026-01', 始まりの日: 1 });
  ok(!出.なぜ, '★出せなかった：' + 出.なぜ);
  eq(出.期間.from, '2026-01-01'); eq(出.期間.to, '2026-01-31');
  eq(出.何列, 2, '★1人分（日付＋その人）に なっていない★');
  eq(出.人の列, 'B');
  /* ★2つの道で 数えて 突き合わせる★ */
  const 描 = 描かれた字を足す(出);
  const 別 = 別の道で数える('白石正人', 出.期間);
  eq(描.行数, 別.数, '★描かれた行数と 別の道の数が 合わない★');
  eq(描.行数, 31, '★1月は31日★');
  console.log('       … 描かれた字 ' + 描.行数 + '行／別の道 ' + 別.数 + '行／合計 ' + Math.round(描.合計).toLocaleString() + '円');
});

T('★「10日からの分」も 出る（締め期間の決まりを 使う）★', () => {
  const 出 = Recipe.切り出す(給料表, { 人: '白石正人', 月: '2026-01', 始まりの日: 10 });
  ok(!出.なぜ, '★出せなかった：' + 出.なぜ);
  eq(出.期間.from, '2025-12-10', '★前の月から に なっていない★');
  eq(出.期間.to, '2026-01-09');
  const 描 = 描かれた字を足す(出);
  const 別 = 別の道で数える('白石正人', 出.期間);
  eq(描.行数, 別.数, '★描かれた行数と 別の道の数が 合わない★');
  /* ★この本は 2026年しか 持っていない★＝2025-12-10〜12-31 は 元に無い。
     ⇒ ★9行★（1/1〜1/9）が 正しい。★無い日を 作らない★ */
  eq(描.行数, 9, '★元に無い日を 作っている（か、落としている）★');
  console.log('       … ' + 出.期間.言い方 + ' → ' + 描.行数 + '行（元に在る日だけ）');
});

T('★実物の締め（21日から）でも 出る★', () => {
  const 出 = Recipe.切り出す(給料表, { 人: '白石正人', 月: '2026-03', 始まりの日: 21 });
  ok(!出.なぜ, '★出せなかった：' + 出.なぜ);
  eq(出.期間.from, '2026-02-21'); eq(出.期間.to, '2026-03-20');
  const 描 = 描かれた字を足す(出);
  eq(描.行数, 別の道で数える('白石正人', 出.期間).数);
  eq(描.行数, 28, '★2/21〜3/20 は 28日★');
});

T('★8人 全員 出せる（1人だけ たまたま 通ったのではない）★', () => {
  const 人たち = ['白石正人', '長野孝', '長野真道', '竹内真一郎', '八木俊幸', '結田航平', '正岡卓', '向垣内'];
  const 出た = [];
  for (const 人 of 人たち) {
    const 出 = Recipe.切り出す(給料表, { 人, 月: '2026-01', 始まりの日: 1 });
    ok(!出.なぜ, '★' + 人 + ' が 出せない：' + 出.なぜ);
    const 描 = 描かれた字を足す(出);
    eq(描.行数, 31, '★' + 人 + ' の行数が 違う★');
    eq(描.行数, 別の道で数える(人, 出.期間).数, '★' + 人 + ' が 別の道と 合わない★');
    出た.push(人 + '=' + Math.round(描.合計).toLocaleString() + '円');
  }
  eq(出た.length, 8);
  console.log('       … 1月分（1/1〜1/31）: ' + 出た.join(' / '));
});

T('★12月分まで 出せる（月をまたぐ所で 落ちない）★', () => {
  let 合計 = 0;
  for (let m = 1; m <= 12; m++) {
    const ym = '2026-' + String(m).padStart(2, '0');
    const 出 = Recipe.切り出す(給料表, { 人: '白石正人', 月: ym, 始まりの日: 1 });
    ok(!出.なぜ, '★' + ym + ' が 出せない：' + 出.なぜ);
    const 描 = 描かれた字を足す(出);
    eq(描.行数, 別の道で数える('白石正人', 出.期間).数, '★' + ym + ' が 合わない★');
    合計 += 描.行数;
  }
  eq(合計, 365, '★1年ぶんの日数に ならない（' + 合計 + '日）★');
  console.log('       … 1月分〜12月分を 足すと ' + 合計 + '日（2026年は 365日）');
});

T('★元の表は 1セルも 触っていない★', () => {
  const 前 = Object.keys(給料表.data).length;
  const 印 = JSON.stringify(給料表.data['2,1']);
  Recipe.切り出す(給料表, { 人: '白石正人', 月: '2026-01', 始まりの日: 1 });
  eq(Object.keys(給料表.data).length, 前, '★セルが 増えた／減った★');
  eq(JSON.stringify(給料表.data['2,1']), 印, '★中身が 書き換わった★');
});

T('★居ない人・1日も無い月は はっきり断る（黙って 空を出さない）★', () => {
  const a = Recipe.切り出す(給料表, { 人: '居ない人', 月: '2026-01' });
  ok(a.なぜ && !a.data, '★空のシートを 作った★');
  const b = Recipe.切り出す(給料表, { 人: '白石正人', 月: '2027-05', 始まりの日: 1 });
  ok(b.なぜ && b.なぜ.indexOf('1日も') > 0, '★2027年に 出せてしまった：' + JSON.stringify(b));
});

console.log('');
console.log('  ' + pass + ' 緑 / ' + fail + ' 赤');

if (SELF) {
  const { spawnSync } = await import('node:child_process');
  const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'exally-kiri-'));
  console.log('');
  console.log('[self-test] わざと壊して 赤くなるかを数える（★repo は読むだけ★）');
  const BREAKS = [
    ['★期間で 絞らない（何月分でも 全部 出す）★',
      (s) => s.replace('if (期間 && !KI.期間の中か(ymd, 期間)) continue;', '')],
    ['★別の人の列を 出す★',
      (s) => s.replace('if (名 && hv === 名) 当たり.push(c2);', 'if (名) 当たり.push(c2);')],
    /* ★lib/kikan.js を 差し替えても recipe.js は 自分の隣の物を 読む★＝壊れない。
       ⇒ ★recipe.js の側で 端をずらす★（実際に 素通りしたので 直した 2026-08-27）。 */
    ['★期間の端を 1日 ずらす（始まりの日が 落ちる）★',
      (s) => s.replace('if (期間 && !KI.期間の中か(ymd, 期間)) continue;',
                       'if (期間 && !(ymd > 期間.from && ymd <= 期間.to)) continue;')],
    ['★元の表を 書き換える★',
      (s) => s.replace("出['0,0'] = { v: 元の見出し ? セルの値(元の見出し) : '日付' };",
                       "data['0,0'] = { v: 'こわした' };\n    出['0,0'] = { v: 元の見出し ? セルの値(元の見出し) : '日付' };")],
    ['★居ない人でも 空で 出す★',
      (s) => s.replace("if (!当たり.length) return { なぜ: '「' + 名 + '」という見出しの列が 見つかりません' };", '')],
    ['★1日も無い月でも 出す★',
      (s) => s.replace('if (!出す行.length) {', 'if (false) {')],
  ];
  let red = 0;
  for (const [name, brk] of BREAKS) {
    const rel = 'lib/recipe.js';
    const orig = fs.readFileSync(path.join(ROOT, rel), 'utf8');
    const bad = brk(orig);
    if (bad === orig) { console.log('  ★置換できず★  ' + name); continue; }
    const f = path.join(TMP, rel.replace(/[\\/]/g, '_'));
    fs.writeFileSync(f, bad, 'utf8');
    const env = Object.assign({}, process.env, { EXALLY_KIRIDASHI_OVERRIDE: JSON.stringify({ [rel]: f }) });
    const r = spawnSync(process.execPath, ['--max-old-space-size=4096', path.join(__dirname, 'kiridashi-jitsubutsu.test.mjs')], { encoding: 'utf8', env });
    if (r.status !== 0) { red++; console.log('  赤くなった  ' + name); }
    else console.log('  ★素通り★  ' + name);
  }
  try { fs.rmSync(TMP, { recursive: true, force: true }); } catch (e) { /* 消せなくても検査は済んでいる */ }
  console.log('');
  console.log('  ' + red + '/' + BREAKS.length + ' 通りで赤くなった');
  process.exit(red === BREAKS.length ? 0 : 1);
}

process.exit(fail ? 1 : 0);
