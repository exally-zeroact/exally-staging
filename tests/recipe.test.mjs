/* recipe.test.mjs — ★6 レシピ＝2回目からは AIを呼ばない★
 *
 *  ★指示役の検証要件（2026-08-26）★
 *    ★レシピを保存して 2回目に AIが0回である事（記録の行数で）★
 *
 *  使い方: node tests/recipe.test.mjs [--self-test]
 */
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createRequire } from 'node:module';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(__dirname, '..');
const SELF = process.argv.includes('--self-test');
const OVERRIDE = process.env.EXALLY_RECIPE_OVERRIDE ? JSON.parse(process.env.EXALLY_RECIPE_OVERRIDE) : {};
const require_ = createRequire(pathToFileURL(path.join(ROOT, 'package.json')));
const Recipe = require_(OVERRIDE['lib/recipe.js'] || path.join(ROOT, 'lib/recipe.js'));

let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ok   ' + n); } catch (e) { fail++; console.log('  NG   ' + n + '\n       ' + (e && e.message)); } };
const ok = (c, m) => { if (!c) throw new Error(m || 'expected truthy'); };
const eq = (a, b, m) => { if (a !== b) throw new Error((m ? m + ': ' : '') + '期待=' + JSON.stringify(b) + ' 実際=' + JSON.stringify(a)); };

/* 見本の表（毎月 来る形。★中身は毎月 変わる／形は同じ★） */
const 表 = (行数, ずらす) => {
  const data = { '0,0': { v: '日付' }, '0,1': { v: '売上' }, '0,2': { v: '数量' } };
  for (let r = 1; r <= 行数; r++) {
    data[r + ',0'] = { v: '2026-08-' + ('0' + r).slice(-2) };
    data[r + ',1'] = { v: 1000 * r + (ずらす || 0) };
    data[r + ',2'] = { v: r };
  }
  return { name: '売上表', data };
};

console.log('');
console.log('[recipe] ★手順を覚えて 2回目からは AIを呼ばない★');

/* ══ ①要約（AIへ渡す物は 小さい）══ */
T('★列名と型と 見本3行だけ（客の全データを送らない）★', () => {
  const y = Recipe.要約を作る(表(500));
  eq(y.列.map((c) => c.名).join(','), '日付,売上,数量');
  eq(y.列.map((c) => c.型).join(','), '日付,数,数');
  eq(y.見本.length, 3, '★見本が 3行でない★');
  eq(y.行数, 501);
});
T('★AIへ渡す字が 小さい（500行でも 1,500字を超えない）★', () => {
  const y = Recipe.要約を作る(表(500));
  const 字 = Recipe.AIへ渡す物(y, '売上に消費税を足した列がほしい');
  ok(字.length < 1500, '★' + 字.length + '字（大きすぎる＝毎回 お金がかかる）★');
  ok(字.indexOf('9000') < 0 || 字.indexOf('見本') > 0, '');
  /* ★4行目より後ろの中身は 入っていない★ */
  ok(字.indexOf('2026-08-09') < 0, '★見本を超えて 客の中身を送っている★');
  console.log('       … AIへ渡す字 ' + 字.length + '字（500行の表）');
});
T('★見出しが無くても 落ちない★', () => {
  const y = Recipe.要約を作る({ name: 'あ', data: { '0,0': { v: '' }, '1,0': { v: 1 } } });
  eq(y.列[0].名, '（見出し無し）');
});
T('★空の表でも 落ちない★', () => {
  const y = Recipe.要約を作る({ name: 'あ', data: {} });
  eq(y.列数, 0);
  eq(y.見本.length, 0);
});

/* ══ ②「前と同じ事」を AIを呼ぶ前に 当てる ══ */
T('★列の名前と並びが同じなら 同じ形（中身と行数は 変わってよい）★', () => {
  ok(Recipe.同じ形か(Recipe.要約を作る(表(10)), Recipe.要約を作る(表(999, 7))));
});
T('★列が1つ増えたら 別の形（黙って当てない）★', () => {
  const b = 表(10);
  b.data['0,3'] = { v: '備考' };
  ok(!Recipe.同じ形か(Recipe.要約を作る(表(10)), Recipe.要約を作る(b)));
});
T('★当てられない時は 理由を返す（黙って外さない）★', () => {
  const r0 = Recipe.当ててみる(Recipe.要約を作る(表(10)), []);
  eq(r0.見つかった, false);
  ok(r0.なぜ.length > 3, '★なぜ当たらなかったかを 言っていない★');
  const b = 表(10); b.data['0,1'] = { v: '金額' };
  const r1 = Recipe.当ててみる(Recipe.要約を作る(b), [{ 要約: Recipe.要約を作る(表(10)) }]);
  eq(r1.見つかった, false);
  ok(r1.なぜ.indexOf('列') >= 0, '★理由が 具体的でない★：' + r1.なぜ);
});
T('★覚えていれば 当たる★', () => {
  const 覚え = [{ 名: '消費税', 要約: Recipe.要約を作る(表(10)), 手順: [] }];
  const r = Recipe.当ててみる(Recipe.要約を作る(表(300, 5)), 覚え);
  eq(r.見つかった, true);
  eq(r.レシピ.名, '消費税');
});

/* ══ ③AIの返事を 手順として読む（決めた形以外は 断る）══ */
T('★決めた4つだけ 受け取る★', () => {
  const r = Recipe.手順を読む('こうします\n{"手順":[{"種類":"式の列を足す","見出し":"税込","式":"=B{行}*1.1"}]}');
  eq(r.ok, true);
  eq(r.手順.length, 1);
});
T('★知らない種類は 断る（AIの言う通りに動かない）★', () => {
  const r = Recipe.手順を読む('{"手順":[{"種類":"ファイルを消す","場所":"C:/"}]}');
  eq(r.ok, false);
  ok(r.なぜ.indexOf('ファイルを消す') >= 0, '★何を断ったかを 言っていない★');
});
T('★中身が足りない手順は 断る★', () => {
  eq(Recipe.手順を読む('{"手順":[{"種類":"式の列を足す","見出し":"税込"}]}').ok, false);
  eq(Recipe.手順を読む('{"手順":[{"種類":"式の列を足す","見出し":"税込","式":"B*1.1"}]}').ok, false, '★= で始まらない式★');
});
T('★手順が返ってこなければ 断る（黙って何もしない ではない）★', () => {
  const r = Recipe.手順を読む('すみません、分かりません');
  eq(r.ok, false);
  ok(r.なぜ.length > 3);
});
T('★良い手順と 悪い手順が混ざったら 良い方だけ通し、断った物を出す★', () => {
  const r = Recipe.手順を読む('{"手順":[{"種類":"式の列を足す","見出し":"税込","式":"=B{行}*1.1"},{"種類":"社外へ送る"}]}');
  eq(r.ok, true);
  eq(r.手順.length, 1);
  eq(r.断った.join(','), '社外へ送る');
});

/* ══ ④全行に 決定論で 当てる（★実行に上限を置かない★）══ */
T('★1行も飛ばさず 全行に当てる（5,000行）★', () => {
  const sh = 表(5000);
  const r = Recipe.手順を当てる(sh, [{ 種類: '式の列を足す', 見出し: '税込', 式: '=B{行}*1.1' }]);
  eq(r.変える['0,3'].v, '税込');
  eq(r.変える['1,3'].f, '=B2*1.1');
  eq(r.変える['5000,3'].f, '=B5001*1.1');
  eq(Object.keys(r.変える).length, 5001, '★行を飛ばしている（上限を置いている）★');
});
T('★同じ手順を 2回 当てたら 同じ結果（決定論）★', () => {
  const a = Recipe.手順を当てる(表(50), [{ 種類: '式の列を足す', 見出し: 'x', 式: '=B{行}+1' }]);
  const b = Recipe.手順を当てる(表(50), [{ 種類: '式の列を足す', 見出し: 'x', 式: '=B{行}+1' }]);
  eq(JSON.stringify(a.変える), JSON.stringify(b.変える));
});
T('★列の名前を変える／列を空にする★', () => {
  const r = Recipe.手順を当てる(表(3), [{ 種類: '列の名前を変える', 元: 'B', 新: '金額' }, { 種類: '列を消す', 列: 'C' }]);
  eq(r.変える['0,1'].v, '金額');
  eq(r.変える['3,2'].v, '');
});
T('★何をしたかを 客の言葉で出す★', () => {
  const r = Recipe.手順を当てる(表(10), [{ 種類: '式の列を足す', 見出し: '税込', 式: '=B{行}*1.1' }]);
  ok(r.なに[0].indexOf('税込') > 0 && r.なに[0].indexOf('10行') > 0, '★何行に何をしたか 出していない★：' + r.なに[0]);
});
T('★客に見せる字に ★ を書かない★', () => {
  const r = Recipe.手順を当てる(表(3), [{ 種類: '式の列を足す', 見出し: 'x', 式: '=B{行}+1' }]);
  for (const s of r.なに) ok(s.indexOf('★') < 0, '★客の字に ★ が出ている★：' + s);
  const y = Recipe.要約を作る(表(3));
  ok(Recipe.AIへ渡す物(y, 'あ').indexOf('★') < 0, '★AIへ渡す字に ★ が出ている★');
  ok(Recipe.当ててみる(y, []).なぜ.indexOf('★') < 0);
});

/* ══ ⑤覚える形（★客の中身は 覚えない★）══ */
T('★覚えるのは 形と手順だけ（見本＝客の中身は 覚えない）★', () => {
  const y = Recipe.要約を作る(表(10));
  const rec = Recipe.レシピを作る('消費税', '税込を足して', y, [{ 種類: '式の列を足す', 見出し: '税込', 式: '=B{行}*1.1' }]);
  ok(!rec.要約.見本, '★見本（客の中身）を 覚えている★');
  eq(rec.指紋, Recipe.指紋(y));
  ok(JSON.stringify(rec).indexOf('2026-08-01') < 0, '★客の中身が 混ざっている★');
});

/* ══ ★⑥2回目は AIを呼ばない（記録の行数で数える）★ ══ */
T('★1回目は AI 1回／2回目は AI 0回（同じ形の表）★', () => {
  /* ★偽のAI★＝呼ばれた回数を数える（お金は1円も使わない） */
  let 呼んだ = 0;
  const 記録 = [];
  const 覚え = [];
  function 頼む(sh, 頼み) {
    const y = Recipe.要約を作る(sh);
    const 当て = Recipe.当ててみる(y, 覚え);
    if (当て.見つかった) {
      const r = Recipe.手順を当てる(sh, 当て.レシピ.手順);
      記録.push({ 種類: '自動化', AI: 0, なに: r.なに });
      return { AIを呼んだ: false, 変える: r.変える, なぜ: 当て.なぜ };
    }
    呼んだ++;
    const 返事 = '{"手順":[{"種類":"式の列を足す","見出し":"税込","式":"=B{行}*1.1"}]}';
    const 読み = Recipe.手順を読む(返事);
    覚え.push(Recipe.レシピを作る('消費税', 頼み, y, 読み.手順));
    const r2 = Recipe.手順を当てる(sh, 読み.手順);
    記録.push({ 種類: '自動化', AI: 1, なに: r2.なに });
    return { AIを呼んだ: true, 変える: r2.変える, なぜ: 当て.なぜ };
  }
  const 一 = 頼む(表(10), '税込を足して');
  eq(一.AIを呼んだ, true, '1回目');
  const 二 = 頼む(表(31, 100), '税込を足して');   /* ★来月の表（行数も中身も違う・形は同じ）★ */
  eq(二.AIを呼んだ, false, '★2回目も AIを呼んでいる★');
  eq(呼んだ, 1, '★AIを呼んだ回数★');
  /* ★記録の行数で 数える（指示役の検証要件）★ */
  eq(記録.length, 2, '★記録が 2行ない★');
  eq(記録[0].AI, 1);
  eq(記録[1].AI, 0, '★2回目の記録に AI 1回 と書いてある★');
  eq(記録.reduce((a, x) => a + x.AI, 0), 1, '★AIの合計回数★');
  /* ★2回目も ちゃんと全行 直っている（手抜きしていない）★ */
  eq(Object.keys(二.変える).length, 32);
  eq(二.変える['31,3'].f, '=B32*1.1');
  console.log('       … 1回目 AI 1回 ／ 2回目 AI ★0回★ ／ 記録 2行（合計 AI 1回）');
});

console.log('');
console.log('  ' + pass + ' 緑 / ' + fail + ' 赤');

if (SELF) {
  const { spawnSync } = await import('node:child_process');
  const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'exally-recipe-'));
  console.log('');
  console.log('[self-test] わざと壊して 赤くなるかを数える（★repo は読むだけ★）');
  const BREAKS = [
    ['★見本を 3行より多く送る（客の中身が 出て行く）★', (s) => s.replace('var 見本の行数 = opt.見本の行数 || 3;', 'var 見本の行数 = 999;')],
    ['★形が違っても 同じだと言う（黙って当てる）★', (s) => s.replace('return !!A && !!B && 指紋(A) === 指紋(B) && 指紋(A) !== \'\';', 'return true;')],
    ['★当てられない理由を 言わない★', (s) => s.replace("return { 見つかった: false, レシピ: null, なぜ: '列の名前か 並びが 前と違います' };", 'return { 見つかった: false, レシピ: null, なぜ: \'\' };')],
    ['★知らない種類も 受け取る（AIの言う通りに動く）★', (s) => s.replace('if (使える種類.indexOf(t.種類) < 0) { 断った.push(String(t.種類)); continue; }', '')],
    ['★= で始まらない式も 通す★', (s) => s.replace("if (!t.見出し || !t.式 || String(t.式).charAt(0) !== '=') { 断った.push('式の列を足す（中身が足りない）'); continue; }", '')],
    ['★全行に当てず 上限を置く★', (s) => s.replace('for (var r = 1; r <= 最大行; r++) {', 'for (var r = 1; r <= Math.min(最大行, 100); r++) {')],
    ['★客の中身（見本）まで 覚える★', (s) => s.replace('要約: { シート: 要約.シート, 列: 要約.列, 列数: 要約.列数 },', '要約: 要約,')],
    ['★何をしたかを 出さない★', (s) => s.replace("なに.push('列 ' + 列の字(c) + ' に「' + t.見出し + '」を足す（' + 最大行 + '行）');", '')],
    ['★客の字に ★ を書く★', (s) => s.replace("なに.push('列 ' + 列の字(c) + ' に「'", "なに.push('★列 ' + 列の字(c) + ' に「'")],
    ['★行の番号を ずらす（1行 ずれる）★', (s) => s.replace('String(t.式).replace(/\\{行\\}/g, String(r + 1))', 'String(t.式).replace(/\\{行\\}/g, String(r))')],
  ];
  let red = 0;
  for (const [name, brk] of BREAKS) {
    const orig = fs.readFileSync(path.join(ROOT, 'lib/recipe.js'), 'utf8');
    const bad = brk(orig);
    if (bad === orig) { console.log('  ★置換できず★  ' + name); continue; }
    const f = path.join(TMP, 'recipe.js');
    fs.writeFileSync(f, bad, 'utf8');
    const env = Object.assign({}, process.env, { EXALLY_RECIPE_OVERRIDE: JSON.stringify({ 'lib/recipe.js': f }) });
    const r = spawnSync(process.execPath, ['--max-old-space-size=4096', path.join(__dirname, 'recipe.test.mjs')], { encoding: 'utf8', env });
    if (r.status !== 0) { red++; console.log('  赤くなった  ' + name); }
    else console.log('  ★素通り★  ' + name);
  }
  try { fs.rmSync(TMP, { recursive: true, force: true }); } catch (e) { /* 消せなくても検査は済んでいる */ }
  console.log('');
  console.log('  ' + red + '/' + BREAKS.length + ' 通りで赤くなった');
  process.exit(red === BREAKS.length ? 0 : 1);
}

process.exit(fail ? 1 : 0);
