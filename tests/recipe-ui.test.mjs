/* recipe-ui.test.mjs — ★8-③ 覚えた手順（レシピ）を 画面で 実際に押す★
 *
 *  ★指示役の指示（2026-08-27）★
 *    ・押す前「AIを使います」／押した後「AIを ◯回 使いました」
 *    ・★「前と同じ事」を AIを呼ぶ前に 機械が当てる★（当てられなければ 普通に呼ぶ・★黙って外さない★）
 *    ・★実行に上限を置かない★
 *  ★部品が緑＝画面で使える ではない★ので、ここでは ★本物の book.html の字を そのまま動かす★。
 *
 *  使い方: node tests/recipe-ui.test.mjs [--self-test]
 */
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createRequire } from 'node:module';
import { 注記を外す } from '../scripts/lib/chuki.mjs';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(__dirname, '..');
const SELF = process.argv.includes('--self-test');
const OVERRIDE = process.env.EXALLY_RECIPEUI_OVERRIDE ? JSON.parse(process.env.EXALLY_RECIPEUI_OVERRIDE) : {};
const 読む = (rel) => fs.readFileSync(OVERRIDE[rel] || path.join(ROOT, rel), 'utf8');
const require_ = createRequire(pathToFileURL(path.join(ROOT, 'package.json')));
const Recipe = require_(OVERRIDE['lib/recipe.js'] || path.join(ROOT, 'lib/recipe.js'));
const DiffPreview = require_(path.join(ROOT, 'lib/diff-preview.js'));
const Kikan = require_(path.join(ROOT, 'lib/kikan.js'));

let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ok   ' + n); } catch (e) { fail++; console.log('  NG   ' + n + '\n       ' + (e && e.message)); } };
const ok = (c, m) => { if (!c) throw new Error(m || 'expected truthy'); };
const eq = (a, b, m) => { if (a !== b) throw new Error((m ? m + ': ' : '') + '期待=' + JSON.stringify(b) + ' 実際=' + JSON.stringify(a)); };

const book = 読む('book.html');

console.log('');
console.log('[recipe-ui] ★覚えた手順を 画面で 押す（AIは 1回も 呼ばない）★');

/* ── 本物の画面から そのまま切り出す（写し取らない＝写した瞬間に 古くなる）── */
const 切る = (頭, 尻, なに) => {
  const i = book.indexOf(頭), j = book.indexOf(尻, i + 1);
  if (i < 0 || j < 0) throw new Error('★' + なに + ' が 画面に 見つかりません★');
  return book.slice(i, j);
};
const レシピの所 = () => 切る('var _覚えた手順 = [];', 'function 診断を始める(){', 'レシピの所');
const 窓の字 = () => 切る('<div id="rcOverlay"', '<!-- ★6 履歴', '見せる窓');

let JSDOM;
try { ({ JSDOM } = await import('jsdom')); }
catch { console.log('★jsdomが入っていません。この検証は飛ばせません（SKIPを緑と呼ばない）'); process.exit(1); }

/** ★本物の画面の字を そのまま動かす★（周りだけ こちらで用意する） */
function 台(opt) {
  opt = opt || {};
  const dom = new JSDOM('<!doctype html><html><body>' + 窓の字()
    + '<div id="ai-response-area"></div><div id="popup-response-area"></div></body></html>',
    { pretendToBeVisual: true });
  const w = dom.window;
  const 記録 = { 書いた: [], 知らせ: [], 履歴: [], 言った: [], 計算に足した: [] };
  const 台本 = {
    document: w.document,
    sheets: opt.sheets || [],
    activeSheet: 0,
    Recipe: Recipe,
    DiffPreview: DiffPreview,
    RecipeStore: { 覚える: () => Promise.resolve(true), 覚えた物: () => Promise.resolve({ レシピたち: [], 倉庫: true }) },
    setCell: (r, c, v) => 記録.書いた.push({ r, c, v }),
    render: () => {},
    switchSheet: (i) => { 記録.移った = i; },
    rebuildSheetTabs: () => { 記録.タブを作り直した = true; },
    addSheetToEngine: (n) => 記録.計算に足した.push(n),
    loadSheetIntoEngine: () => {},
    Kikan: Kikan,
    showToast: (h) => 記録.知らせ.push(String(h)),
    履歴に残す: (種類, 見出し, 中身, f, credit) => 記録.履歴.push({ 種類, 見出し, credit }),
    addAIChatMsg: (t, x) => 記録.言った.push({ どこ: 'panel', 字: String(x) }),
    _addPopupMsg: (t, x) => 記録.言った.push({ どこ: 'popup', 字: String(x) }),
    _aiFetch: () => Promise.resolve({ ok: false, 言葉: 'この検査では AIを 呼びません' }),
  };
  const 名 = Object.keys(台本);
  const f = new w.Function(...名, レシピの所()
    + ';return {先に手順で当てる:先に手順で当てる,手順を書く:手順を書く,手順をやめる:手順をやめる,'
    + '覚えた:function(a){_覚えた手順=a;},いまの表の要約:いまの表の要約,'
    + '使えるシート名:使えるシート名,きれいな値:_きれいな値};');
  return { api: f(...名.map((k) => 台本[k])), 記録, w };
}

const 表 = (行数) => {
  const data = { '0,0': { v: '日付' }, '0,1': { v: '金額' } };
  for (let r = 1; r <= 行数; r++) { data[r + ',0'] = { v: '8/' + r }; data[r + ',1'] = { v: 100 * r }; }
  return [{ name: '売上', data: data }];
};
const レシピ = (頼み, 表たち) => Recipe.レシピを作る(頼み, 頼み, Recipe.要約を作る(表たち[0]),
  [{ 種類: '式の列を足す', 見出し: '税込', 式: '=B{行}*1.1' }]);

/* ══ ①AIを呼ぶ前に 機械が当てる ══ */
T('★覚えた物が 1つも無ければ 当てない（黙って何も言わない）★', () => {
  const { api, 記録 } = 台({ sheets: 表(5) });
  eq(api.先に手順で当てる('税込を足して'), false);
  eq(記録.言った.length, 0, '★覚えていないのに 客に 言い訳している★');
});
T('★前と同じ事なら AIを1回も呼ばずに 当てる★', () => {
  const t = 表(5);
  const { api, 記録, w } = 台({ sheets: t });
  api.覚えた([レシピ('税込を足して', t)]);
  eq(api.先に手順で当てる('税込を足して'), true, '★当てていない（AIに行ってしまう）★');
  eq(記録.書いた.length, 0, '★見せる前に 書き換えている（勝手に直している）★');
  eq(w.document.getElementById('rcOverlay').style.display, 'flex', '★見せていない★');
  ok(記録.言った.some((x) => x.字.indexOf('AIは使いません') > 0), '★0円だと 言っていない★');
});
T('★実行に上限を置かない（500行でも 全部・見せる数と 直す数が 同じ）★', () => {
  const t = 表(500);
  const { api, 記録, w } = 台({ sheets: t });
  api.覚えた([レシピ('税込を足して', t)]);
  api.先に手順で当てる('税込を足して');
  const 字 = w.document.getElementById('rcGo').textContent;
  eq(字, 'この 501か所を 直す', '★上限で 切っている（見出し1＋500行）★');
  /* ★見せた数と 直した数が 同じか★＝見せる所だけ直しても、書く所で 切っていたら 嘘になる。
     ★実際に 切ってみたら 素通りした（2026-08-27）★＝小さい表でしか 押していなかった。 */
  api.手順を書く();
  eq(記録.書いた.length, 501, '★見せた数だけ 直していない（書く所で 切っている）★');
  eq(記録.書いた[500].v, '=B501*1.1', '★最後の行に 当たっていない★');
});
T('★押して はじめて 直る／AIを 0回 と 言う★', () => {
  const t = 表(5);
  const { api, 記録 } = 台({ sheets: t });
  api.覚えた([レシピ('税込を足して', t)]);
  api.先に手順で当てる('税込を足して');
  api.手順を書く();
  eq(記録.書いた.length, 6, '★全行に 当たっていない★');
  eq(記録.書いた[1].v, '=B2*1.1', '★行の番号に なっていない★');
  ok(記録.知らせ.some((x) => x.indexOf('AIを 0回 使いました') > 0), '★0回だと 言っていない★');
  eq(記録.履歴[0].credit, 0, '★履歴に AIを使った事にしている★');
});
T('★やめたら 1セルも 書かない★', () => {
  const t = 表(5);
  const { api, 記録, w } = 台({ sheets: t });
  api.覚えた([レシピ('税込を足して', t)]);
  api.先に手順で当てる('税込を足して');
  api.手順をやめる();
  api.手順を書く();
  eq(記録.書いた.length, 0, '★やめたのに 書いている★');
  eq(w.document.getElementById('rcOverlay').style.display, 'none', '★窓が 開いたまま★');
});

/* ══ ②当てられない時は 黙って外さない ══ */
T('★同じ形でも 違う頼みなら 当てない（理由を言ってから AIへ）★', () => {
  const t = 表(5);
  const { api, 記録 } = 台({ sheets: t });
  api.覚えた([レシピ('税込を足して', t)]);
  eq(api.先に手順で当てる('平均を出して'), false);
  ok(記録.言った.some((x) => x.字.indexOf('前と違う頼みです') > 0), '★なぜ使わなかったかを 言っていない★');
  ok(記録.言った.some((x) => x.字.indexOf('AIに聞きます') > 0), '★このあと どうなるかを 言っていない★');
});
T('★列の名前が変わったら 当てない（違う表に 当てない）★', () => {
  const t = 表(5);
  const 覚 = レシピ('税込を足して', t);
  const 別 = 表(5); 別[0].data['0,1'] = { v: '売上高' };
  const { api, 記録 } = 台({ sheets: 別 });
  api.覚えた([覚]);
  eq(api.先に手順で当てる('税込を足して'), false);
  ok(記録.言った.some((x) => x.字.indexOf('覚えた表と違います') > 0), '★理由を 言っていない★');
});
T('★スマホの窓でも 同じ言葉が 出る（口が違っても 中身は1本）★', () => {
  const t = 表(5);
  const { api, 記録 } = 台({ sheets: t });
  api.覚えた([レシピ('税込を足して', t)]);
  const 口 = { 足す: (ty, x) => 記録.言った.push({ どこ: 'popup', 字: String(x) }) };
  api.先に手順で当てる('平均を出して', 口);
  ok(記録.言った.length > 0, '★何も 言っていない★');
  ok(記録.言った.every((x) => x.どこ === 'popup'), '★スマホで押したのに パネルへ出している★');
});

/* ══ ★切り出す（1人分×何月分）を 画面で 押す★ ══════════════
   ★司さんの言葉（2026-08-27）★「1人分と 何月分って 分けて表示もさせれるん？」 */
const 給料表 = (人たち, 始まりの通し番号, 日数) => {
  const d = { '0,0': { v: 'ZERO代行　給料表' }, '1,0': { v: '日付' } };
  人たち.forEach((n, i) => { d['1,' + (i + 1)] = { v: n }; });
  for (let r = 0; r < 日数; r++) {
    /* ★画面と同じ形★＝式のセルは v が空で 計算した値は d（字） */
    d[(r + 2) + ',0'] = { v: '', f: '=A' + r, d: String(始まりの通し番号 + r), numFmt: 'm/d;@' };
    人たち.forEach((n, i) => { d[(r + 2) + ',' + (i + 1)] = { v: '', f: '=X', d: String((i + 1) * 100 + r) }; });
  }
  return [{ name: '給料表', data: d }];
};
const 切り出しレシピ = (t, 頼み) => Recipe.レシピを作る(頼み, 頼み, Recipe.要約を作る(t[0]),
  [{ 種類: '切り出す', 人: '白石正人', 月: '2026-01', 始まりの日: 1 }]);

T('★1人分×1月分を 出す（AIを 1回も 呼ばない）★', () => {
  const t = 給料表(['白石正人', '長野孝'], 46023, 60);
  const { api, 記録, w } = 台({ sheets: t });
  api.覚えた([切り出しレシピ(t, '白石正人の1月分を出して')]);
  eq(api.先に手順で当てる('白石正人の1月分を出して'), true, '★当てていない（AIに行ってしまう）★');
  eq(記録.書いた.length, 0, '★見せる前に 書いている★');
  eq(w.document.getElementById('rcOverlay').style.display, 'flex', '★見せていない★');
  eq(w.document.getElementById('rcGo').textContent, '新しいシートに 31行 出す', '★出す数が 違う★');
  ok(w.document.getElementById('rcList').textContent.indexOf('白石正人 1月分') > 0, '★シートの名前を 見せていない★');
});
T('★押すと 新しいシートが 出来る（元の表は 触らない）★', () => {
  const t = 給料表(['白石正人'], 46023, 60);
  const 元 = JSON.stringify(t[0].data);
  const { api, 記録 } = 台({ sheets: t });
  api.覚えた([切り出しレシピ(t, '白石正人の1月分を出して')]);
  api.先に手順で当てる('白石正人の1月分を出して');
  api.手順を書く();
  eq(t.length, 2, '★シートが 増えていない★');
  eq(t[1].name, '白石正人 1月分');
  eq(Object.keys(t[1].data).length, 64, '★31行×2列＋見出し で 64セル★');
  eq(JSON.stringify(t[0].data), 元, '★元の表が 書き換わった★');
  eq(記録.書いた.length, 0, '★元のシートに 書き込んでいる★');
  ok(記録.タブを作り直した, '★タブを 作り直していない（見えない）★');
  eq(記録.計算に足した[0], '白石正人 1月分', '★計算する側に 入れていない（式が 動かない）★');
  ok(記録.知らせ.some((x) => x.indexOf('AIを 0回 使いました') > 0), '★0回だと 言っていない★');
  eq(記録.履歴[0].credit, 0, '★履歴に AIを使った事にしている★');
});
T('★2回目も AIを 0回（同じ頼み・同じ形）★', () => {
  const t = 給料表(['白石正人'], 46023, 60);
  const { api, 記録 } = 台({ sheets: t });
  api.覚えた([切り出しレシピ(t, '白石正人の1月分を出して')]);
  for (let i = 0; i < 2; i++) {
    eq(api.先に手順で当てる('白石正人の1月分を出して'), true, '★' + (i + 1) + '回目に 外れた★');
    api.手順を書く();
  }
  eq(t.length, 3, '★2回目が 出ていない★');
  eq(t[2].name, '白石正人 1月分(2)', '★同じ名前の シートを 2枚 作った★');
  eq(記録.履歴.filter((x) => x.credit > 0).length, 0, '★AIを使った事に なっている★');
});
T('★居ない人の時は 理由を言って AIへ（黙って 外さない）★', () => {
  const t = 給料表(['白石正人'], 46023, 60);
  const { api, 記録 } = 台({ sheets: t });
  const レ = Recipe.レシピを作る('居ない人の1月分', '居ない人の1月分', Recipe.要約を作る(t[0]),
    [{ 種類: '切り出す', 人: '居ない人', 月: '2026-01', 始まりの日: 1 }]);
  api.覚えた([レ]);
  eq(api.先に手順で当てる('居ない人の1月分'), false);
  ok(記録.言った.some((x) => x.字.indexOf('見つかりません') > 0), '★理由を 言っていない★');
  ok(記録.言った.some((x) => x.字.indexOf('AIに聞きます') > 0), '★このあと どうなるかを 言っていない★');
});
T('★スマホの窓でも 同じ道を通る（口が違っても 中身は1本）★', () => {
  const t = 給料表(['白石正人'], 46023, 60);
  const { api, 記録, w } = 台({ sheets: t });
  api.覚えた([切り出しレシピ(t, '白石正人の1月分を出して')]);
  const 口 = { 足す: (ty, x) => 記録.言った.push({ どこ: 'popup', 字: String(x) }) };
  eq(api.先に手順で当てる('白石正人の1月分を出して', 口), true, '★スマホで 当たらない★');
  ok(記録.言った.length > 0 && 記録.言った.every((x) => x.どこ === 'popup'), '★パネルへ出している★');
  eq(w.document.getElementById('rcOverlay').style.display, 'flex', '★スマホで 窓が 出ない★');
  api.手順を書く();
  eq(t.length, 2, '★スマホで 出せていない★');
});
T('★同じ名前のシートを 2枚 作らない★', () => {
  const t = 給料表(['白石正人'], 46023, 60);
  const { api } = 台({ sheets: t });
  eq(api.使えるシート名('給料表'), '給料表(2)');
  eq(api.使えるシート名('新しい名前'), '新しい名前');
});
T('★日付は 通し番号のまま 見せない（1/1 と 見せる）★', () => {
  const { api } = 台({ sheets: 給料表(['白石正人'], 46023, 60) });
  eq(api.きれいな値({ v: 46023 }), '1/1', '★通し番号のまま 見せている★');
  eq(api.きれいな値({ v: 1234 }), '1234', '★金額を 日付に化けさせている★');
  eq(api.きれいな値({ v: '', d: '9,775' }), '9,775');
});

/* ══ ③画面の作り（本物の字を読む） ══ */
T('★聞く口は 1つだけ＝スマホも 同じ1本を通る★', () => {
  const b = 注記を外す(book, { html: true });
  ok(/function _sendToAIMobile\(text\)\{\s*return AIに聞く\(text, _AIの出し口\('popup'\)\);/.test(b),
    '★スマホが 別の道を 通っている（地図・掘る・回数・履歴・レシピが 抜ける）★');
  ok(b.indexOf('async function AIに聞く(text, 口)') > 0, '★中身が 1本に なっていない★');
  eq((b.match(/await _aiFetch\(\{ message: 送る/g) || []).length, 1, '★AIを呼ぶ所が 2つ在る★');
});
T('★AIを呼ぶ前に 機械が当てる（呼んでから ではない）★', () => {
  const b = 注記を外す(book, { html: true });
  const i = b.indexOf('async function AIに聞く(text, 口)');
  const 当 = b.indexOf('先に手順で当てる(text, 口)', i);
  const 呼 = b.indexOf('await _aiFetch(', i);
  ok(当 > 0 && 呼 > 0 && 当 < 呼, '★AIを呼んでから 当てている（お金がかかってから 当てても 遅い）★');
});
T('★押す前は「AIを使います」＝数を書かない★', () => {
  const b = 注記を外す(book, { html: true });
  ok(b.indexOf('この やり方を 覚える（AIを使います）') > 0, '★押す前に 言っていない★');
  ok(b.indexOf('AIを 1回 使います') < 0, '★押す前に 数を 書いている（掘ると 何度も動く）★');
});
T('★開いた時に 覚えた物を 読む（押す時に 待たせない）★', () => {
  const b = 注記を外す(book, { html: true });
  const i = b.indexOf('提案を始める();');
  ok(i > 0 && b.slice(i, i + 300).indexOf('覚えた手順を読む();') > 0, '★開いた時に 読んでいない★');
});
T('★客に見せる字に ★ を書かない（レシピの所）★', () => {
  const 所 = 注記を外す(レシピの所(), { html: true });
  const 字 = (所.match(/'[^']*'/g) || []).join('');
  ok(字.indexOf('★') < 0, '★客の字に ★ が出ている★：' + 字.slice(0, 120));
});
T('★中の言葉（指紋・レシピ・JSON）を 客に見せない★', () => {
  const 所 = 注記を外す(レシピの所(), { html: true });
  const 字 = (所.match(/'[^']*'/g) || []).join(' ');
  for (const 語 of ['指紋', 'レシピ', 'JSON', '手順を読む']) {
    ok(字.indexOf(語) < 0, '★中の言葉「' + 語 + '」が 客に出ている★');
  }
});

console.log('');
console.log('  ' + pass + ' 緑 / ' + fail + ' 赤');

if (SELF) {
  const { spawnSync } = await import('node:child_process');
  const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'exally-recipeui-'));
  console.log('');
  console.log('[self-test] わざと壊して 赤くなるかを数える（★repo は読むだけ★）');
  const BREAKS = [
    ['book.html', '★AIを呼んでから 当てる★', (s) => s.replace('  if(先に手順で当てる(text, 口)) return;\n', '')],
    ['book.html', '★見せずに 直す★', (s) => s.replace("  document.getElementById('rcOverlay').style.display = 'flex';", '  手順を書く();')],
    ['book.html', '★やめても 書く★', (s) => s.replace('  if(!待) return;', '')],
    ['book.html', '★上限を置く（先頭100行だけ）★',
      (s) => s.replace('  var 番地 = Object.keys(待.変える);', '  var 番地 = Object.keys(待.変える).slice(0, 100);')],
    ['book.html', '★0回だと 言わない★', (s) => s.replace('<span class="toast-n">AIを 0回 使いました</span>', '')],
    ['book.html', '★履歴に AIを使った事にする★',
      (s) => s.replace("{ やった: '覚えた手順を 当てた（AIは呼んでいない）', 何か所: 番地.length + 出した行 }, null, 0);",
                       "{ やった: '覚えた手順を 当てた', 何か所: 番地.length + 出した行 }, null, 1);")],
    ['book.html', '★黙って外す（理由を言わない）★',
      (s) => s.replace("    if(_覚えた手順.length) 言う('ai', '覚えた手順は 使いませんでした（' + 答.なぜ + '）。AIに聞きます。');", '')],
    ['book.html', '★スマホを 別の道に 戻す★',
      (s) => s.replace("function _sendToAIMobile(text){\n  return AIに聞く(text, _AIの出し口('popup'));\n}",
                       "async function _sendToAIMobile(text){\n  var r = await _aiFetch({ message: text, history: [] });\n  _addPopupMsg('ai', r.言葉);\n}")],
    ['book.html', '★スマホで押した言葉を パネルへ出す★',
      (s) => s.replace("  var 言う = (口 && 口.足す) ? 口.足す : function(t, x){ addAIChatMsg(t, x); };",
                       "  var 言う = function(t, x){ addAIChatMsg(t, x); };")],
    ['book.html', '★押す前に 数を書く★', (s) => s.replace('この やり方を 覚える（AIを使います）', 'この やり方を 覚える（AIを 1回 使います）')],
    ['book.html', '★開いた時に 覚えた物を 読まない★', (s) => s.replace('      覚えた手順を読む();\n', '')],
    ['book.html', '★客の字に ★ を書く★',
      (s) => s.replace("'前と同じ事なので、覚えた手順で やります。AIは使いません。'", "'★前と同じ事なので、覚えた手順で やります。AIは使いません。★'")],
    ['lib/recipe.js', '★行の番号に 置き換えない★',
      (s) => s.replace('String(t.式).replace(/\\{行\\}/g, String(r + 1))', 'String(t.式)')],
    ['lib/recipe.js', '★頼みが違っても 当てる★',
      (s) => s.replace('if (同形 && 頼みの形(R.頼み) === t) {', 'if (同形) {')],
    /* ★切り出す（1人分×何月分）★ */
    ['book.html', '★新しいシートを 作らない（切り出しが 一生 出ない）★',
      (s2) => s2.replace('  for(var m=0;m<待.出すシート.length;m++){', '  for(var m=0;m<0;m++){')],
    ['book.html', '★出すシートを 見せずに 出す（直す前に 見せない）★',
      (s2) => s2.replace('  var 新 = _手順の待ち.出すシート;', '  var 新 = [];')],
    ['book.html', '★同じ名前のシートを 2枚 作る★',
      (s2) => s2.replace("  while(在る(今)){ 今 = 元 + '(' + n + ')'; n++; }", '')],
    ['book.html', '★タブを 作り直さない（出したのに 見えない）★',
      (s2) => s2.replace('    rebuildSheetTabs();\n    switchSheet(sheets.length - 1);', '    switchSheet(sheets.length - 1);')],
    ['book.html', '★計算する側に 入れない（式が 動かない）★',
      (s2) => s2.replace("    if(typeof addSheetToEngine === 'function'){ try{ addSheetToEngine(名); }catch(e){} }", '')],
    ['book.html', '★日付を 通し番号のまま 見せる★',
      (s2) => s2.replace('    if(ymd && cell.v > 40000 && cell.v < 60000) return Kikan.日の字(ymd, false);', '')],
    ['book.html', '★出せない理由を 言わない（黙って 外す）★',
      (s2) => s2.replace("      ? ('（' + 出.出すシート[0].なぜ + '）') : '（当てても 変わる所が ありませんでした）';",
                         "      ? '' : '';")],
  ];
  let red = 0;
  for (const [rel, name, brk] of BREAKS) {
    const orig = fs.readFileSync(path.join(ROOT, rel), 'utf8');
    const bad = brk(orig);
    if (bad === orig) { console.log('  ★置換できず★  ' + name); continue; }
    const f = path.join(TMP, rel.replace(/[\\/]/g, '_'));
    fs.writeFileSync(f, bad, 'utf8');
    const env = Object.assign({}, process.env, { EXALLY_RECIPEUI_OVERRIDE: JSON.stringify({ [rel]: f }) });
    const r = spawnSync(process.execPath, ['--max-old-space-size=4096', path.join(__dirname, 'recipe-ui.test.mjs')], { encoding: 'utf8', env });
    if (r.status !== 0) { red++; console.log('  赤くなった  ' + name); }
    else console.log('  ★素通り★  ' + name);
  }
  try { fs.rmSync(TMP, { recursive: true, force: true }); } catch (e) { /* 消せなくても検査は済んでいる */ }
  console.log('');
  console.log('  ' + red + '/' + BREAKS.length + ' 通りで赤くなった');
  process.exit(red === BREAKS.length ? 0 : 1);
}

process.exit(fail ? 1 : 0);
