/* vba-ui.test.mjs — ★マクロが 何をしているかを 画面で 実際に 出す★
 *
 *  ★決まり（指示役 2026-08-28）★
 *    ・画面は ★何をしているかを 日本語で 出す★（E2診断の中）
 *    ・★出来ていない物のボタンを 見せない★＝読めた物も 未測定も 0本なら 出さない
 *    ・★正しく読めなかった物を 黙って 落とさない★＝「未測定 ◯本」と 書く
 *    ・★客に見せる字に ★ を書かない★／中の言葉（VBA・CFB…）を そのまま 出さない
 *
 *  ★数え方（何度も 踏んだ所）★
 *    ★「◯本 見つけた」で 緑にしない。★画面に 描かれた字★を 数える★。
 *    ここでは 本物の book.html の 窓と 動きを そのまま 動かして、
 *    出来た DOM の textContent から 数える。
 *
 *  使い方: node tests/vba-ui.test.mjs [--self-test]
 */
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createRequire } from 'node:module';
import { spawnSync } from 'node:child_process';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(__dirname, '..');
const SELF = process.argv.includes('--self-test');
const OVERRIDE = process.env.EXALLY_VBAUI_OVERRIDE ? JSON.parse(process.env.EXALLY_VBAUI_OVERRIDE) : {};
const 読む = (rel) => fs.readFileSync(OVERRIDE[rel] || path.join(ROOT, rel), 'utf8');
const require_ = createRequire(pathToFileURL(path.join(ROOT, 'package.json')));
const V = require_(path.join(ROOT, 'lib/vba.js'));
const M = require_(path.join(ROOT, 'lib/vba-mikata.js'));
const TJ = require_(path.join(ROOT, 'lib/vba-tejun.js'));
const R = require_(path.join(ROOT, 'lib/recipe.js'));
const XLSX = require_(path.join(ROOT, 'lib/xlsx.full.min.js'));
const ZS = require_(path.join(ROOT, 'lib/zip-surgeon.js'));

let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ok   ' + n); } catch (e) { fail++; console.log('  NG   ' + n + '\n       ' + (e && e.message)); } };
const ok = (c, m) => { if (!c) throw new Error(m || 'expected truthy'); };
const eq = (a, b, m) => { if (a !== b) throw new Error((m ? m + ': ' : '') + '期待=' + JSON.stringify(b) + ' 実際=' + JSON.stringify(a)); };

const book = 読む('book.html');

console.log('');
console.log('[vba-ui] ★マクロが 何をしているかを 画面で 出す（AIは0回）★');

/* ── 本物の画面から そのまま 切り出す（写し取らない＝写した瞬間に 古くなる）── */
const 切る = (頭, 尻, なに) => {
  const i = book.indexOf(頭), j = book.indexOf(尻, i + 1);
  if (i < 0 || j < 0) throw new Error('★' + なに + ' が 画面に 見つかりません★');
  return book.slice(i, j);
};
const 動きの所 = () => 切る('var _マクロ = null;', 'function 参照の網を作り始める(', 'マクロの動き');
const 窓の字2 = () => 切る('<div id="rcOverlay"', '<!-- ★6 履歴', 'レシピの窓');
const 窓の字 = () => 切る('<div id="mcOverlay"', '<!-- ★条件付き書式★', 'マクロの窓');
const ボタンの字 = () => 切る('<button id="macroBtn"', '<button id="shindanBtn"', 'マクロのボタン');

let JSDOM;
try { ({ JSDOM } = await import('jsdom')); }
catch { console.log('★jsdomが入っていません。この検証は飛ばせません（SKIPを緑と呼ばない）'); process.exit(1); }

/** ★本物の画面の字を そのまま動かす★（周りだけ こちらで用意する） */
function 台(opt) {
  const dom = new JSDOM('<!doctype html><html><body>' + ボタンの字() + '</button>' + 窓の字() + 窓の字2() + '</body></html>',
    { pretendToBeVisual: true });
  const w = dom.window;
  const 記録 = { 知らせ: [], 覚えた: [], 倉庫: [], 履歴: [], 見せた: null, 書いた: 0 };
  const 台本 = {
    document: w.document,
    VbaMikata: M,
    VbaTejun: TJ,
    Recipe: R,
    RecipeStore: { 覚える: (レ) => { 記録.倉庫.push(レ); return Promise.resolve(true); } },
    sheets: (opt && opt.sheets) || [],
    activeSheet: 0,
    _覚えた手順: 記録.覚えた,
    いまの表の要約: () => (((opt && opt.sheets) || [])[0] ? R.要約を作る(opt.sheets[0]) : null),
    履歴に残す: (種類, 見出し, 中身, f, credit) => 記録.履歴.push({ 種類, 見出し, 中身, credit }),
    showToast: (t) => 記録.知らせ.push({ 題: '', 本文: String(t) }),
    手順を見せる: (題, 本文, 何, 変える, 出すシート) => {
      記録.見せた = { 題, 本文, 何, 変える, 出すシート };
      記録.書いた = 0;   /* ★見せた所では 1セルも 書かない★ */
    },
    DiffPreview: require_(path.join(ROOT, 'lib/diff-preview.js')),
    開いた知らせに足す: (題, 本文) => 記録.知らせ.push({ 題: String(題), 本文: String(本文) }),
  };
  const 名 = Object.keys(台本);
  const f = new w.Function(...名, 動きの所()
    + ';return { マクロを受け取る: マクロを受け取る, マクロの知らせを出す: マクロの知らせを出す,'
    + ' openMacro: openMacro, closeMacro: closeMacro, マクロの手順を覚える: マクロの手順を覚える };');
  return { api: f(...名.map((k) => 台本[k])), 記録, w };
}

/* ── 本物の見本（うちで作った .xlsm）から 見立てを作る ── */
const z = ZS.read(new Uint8Array(fs.readFileSync(path.join(ROOT, 'tests/fixtures/vba-sample.xlsm'))));
const 読み = V.読む(await z.bytes('xl/vbaProject.bin'), XLSX.CFB);
const 見立て = M.見立てる(読み.モジュール);
const マクロ = { 読み: 読み, 見立て: 見立て };

/* ══ ①ボタン＝★在る時だけ 出す★ ══════════════════════════ */
T('マクロが 無いファイルでは ボタンも 知らせも 出さない', () => {
  const { api, w, 記録 } = 台();
  api.マクロを受け取る(null); api.マクロの知らせを出す();
  eq(w.document.getElementById('macroBtn').hidden, true, 'ボタンが 出ている');
  eq(記録.知らせ.length, 0, '知らせまで 出している');
});
T('★中身が 1本も 読めなかった時も 出さない（空の窓を 見せない）★', () => {
  const { api, w } = 台();
  api.マクロを受け取る({ 読み: { ok: false, モジュール: [] }, 見立て: { 手続き: [], 数: {}, 本数: 0, 未測定: [] } });
  eq(w.document.getElementById('macroBtn').hidden, true, 'ボタンが 出ている');
});
T('マクロが 在れば ボタンに 本数が 出る', () => {
  const { api, w, 記録 } = 台();
  api.マクロを受け取る(マクロ);
  const b = w.document.getElementById('macroBtn');
  eq(b.hidden, false, 'ボタンが 隠れている');
  eq(b.textContent, 'マクロ 5本');
  eq(記録.知らせ.length, 0, '★ボタンを 出した所で 知らせまで 出している（順番が 逆になる）★');
  api.マクロの知らせを出す();
  eq(記録.知らせ.length, 1, '知らせが 出ていない');
});
T('★正しく読めなかった物しか 無い時も ボタンは 出す（黙って 落とさない）★', () => {
  const { api, w } = 台();
  api.マクロを受け取る({
    読み: { ok: true, モジュール: [{ 名: 'Module1', 確か: false, なぜ: '詰め方が違う' }] },
    見立て: { 手続き: [], 数: {}, 本数: 0, 未測定: [{ 名: 'Module1', なぜ: '詰め方が違う' }] },
  });
  const b = w.document.getElementById('macroBtn');
  eq(b.hidden, false, '未測定が 在るのに 隠している');
  eq(b.textContent, 'マクロ 未測定 1本');
});

T('★次に マクロの無いファイルを 開いたら 消える（前のが 残らない）★', () => {
  /* ★実物の画面で 押して 見つけた（2026-08-28）★＝
     マクロ入りの次に 普通のファイルを 開くと「マクロ 4本」が 出たままだった。 */
  const { api, w } = 台();
  api.マクロを受け取る(マクロ);
  eq(w.document.getElementById('macroBtn').hidden, false, '1回目で 出ていない');
  api.マクロを受け取る(null);
  eq(w.document.getElementById('macroBtn').hidden, true, '前のファイルの物が 残っている');
});
T('★開く道の 中で 必ず 呼ぶ（マクロが 無い時も）★', () => {
  const i = book.indexOf('マクロを受け取る(res.hasVba ? res.マクロ : null)');
  ok(i > 0, 'マクロが無い時に 消す道が 画面に 無い');
  /* ★hasVba の中だけで 呼んでいない事★を 見る */
  const j = book.indexOf("if (res.hasVba && typeof BookOpen");
  ok(i < j, 'マクロが 在る時しか 呼んでいない');
  /* ★知らせは 「マクロが入っています」の 後★（順番を 逆にすると 何の話か 分からない） */
  const k = book.indexOf("開いた知らせに足す('マクロ（VBA）が入っています'");
  const l = book.indexOf('マクロの知らせを出す();', j);
  ok(k > 0, '「マクロ（VBA）が入っています」が 無い');
  ok(l > 0, '開く道で マクロの知らせを 出していない');
  ok(k < l, '知らせの順番が 逆（本数が 先に出る）');
});

/* ══ ②窓＝★画面に 描かれた字を 数える★ ═══════════════════ */
T('★5本 全部が 画面に 描かれている（見つけた数ではなく 描かれた数）★', () => {
  const { api, w } = 台();
  api.マクロを受け取る(マクロ);
  api.openMacro();
  eq(w.document.getElementById('mcOverlay').style.display, 'flex', '窓が 開いていない');
  const 行 = w.document.getElementById('mcList').children;
  eq(行.length, 5, '描かれた行数');
  const 字 = w.document.getElementById('mcList').textContent;
  for (const 名 of ['月次締め', '印刷', 'CSV取り込み', '月次の並べ替え', 'Worksheet_Change']) {
    ok(字.indexOf(名) >= 0, 名 + ' が 描かれていない');
  }
});
T('1本ずつ「何をしているか」と「ここでは どうするか」が 描かれている', () => {
  const { api, w } = 台();
  api.マクロを受け取る(マクロ); api.openMacro();
  const 行 = w.document.getElementById('mcList').children;
  for (let i = 0; i < 行.length; i++) {
    const t = 行[i].textContent;
    const v = 見立て.手続き[i];
    /* ★「長さが 在る」で 緑にしない★＝見出しと「ここでは」だけでも 30字を 超える。
       ★その1本の 中身そのもの★が 描かれているかを 見る（2026-08-28 実測で 素通りした）。 */
    ok(t.indexOf(v.何をしているか) >= 0, i + '本目の「何をしているか」が 描かれていない: ' + t);
    ok(t.indexOf('ここでは → ' + v.うちのやり方) >= 0, i + '本目の「ここでは」が 描かれていない: ' + t);
    ok(t.indexOf(v.名) >= 0, i + '本目の 名前が 描かれていない: ' + t);
  }
});
T('1本ずつ 可否が 日本語で 描かれている', () => {
  const { api, w } = 台();
  api.マクロを受け取る(マクロ); api.openMacro();
  const 字 = w.document.getElementById('mcList').textContent;
  ok(字.indexOf('そのまま できます') >= 0, '「そのまま できます」が 無い');
  ok(字.indexOf('やり方を かえます') >= 0, '「やり方を かえます」が 無い');
});
T('★終わりの印が 無い1本は そう書く（黙って 次と 1本に しない）★', () => {
  const { api, w } = 台();
  const 見 = M.見立てる([{ 名: 'M', 確か: true, 中身: ['Attribute VB_Name = "M"', 'Sub 先()', 'X: 1', '',
    'Sub 後()', '  Range("A1").Sort', 'End Sub'].join('\r\n') }]);
  api.マクロを受け取る({ 読み: { ok: true, モジュール: [] }, 見立て: 見 });
  api.openMacro();
  const 行 = w.document.getElementById('mcList').children;
  eq(行.length, 2, '描かれた行数');
  ok(行[0].textContent.indexOf('終わりの印') >= 0, '1本目に 断りが 無い: ' + 行[0].textContent);
  ok(行[1].textContent.indexOf('終わりの印') < 0, '2本目にまで 断りが 出ている');
});
T('★未測定は 窓の中に 名前つきで 書く（数に 混ぜない）★', () => {
  const { api, w } = 台();
  api.マクロを受け取る({
    読み: { ok: true, モジュール: 読み.モジュール.concat([{ 名: 'Module9', 確か: false, なぜ: '中身は同じ・詰め方が違う' }]) },
    見立て: 見立て,
  });
  api.openMacro();
  const t = w.document.getElementById('mcMore').textContent;
  ok(t.indexOf('Module9') >= 0, '名前が 出ていない: ' + t);
  ok(t.indexOf('1本') >= 0, '本数が 出ていない: ' + t);
  eq(w.document.getElementById('mcList').children.length, 5, '未測定を 一覧に 混ぜている');
});
T('未測定が 無い時は 何も 書かない（空の言い訳を 出さない）', () => {
  const { api, w } = 台();
  api.マクロを受け取る(マクロ); api.openMacro();
  eq(w.document.getElementById('mcMore').textContent, '');
});
T('閉じられる', () => {
  const { api, w } = 台();
  api.マクロを受け取る(マクロ); api.openMacro(); api.closeMacro();
  eq(w.document.getElementById('mcOverlay').style.display, 'none');
});

/* ══ ②-b 狭い画面で 端が 切れないか ═══════════════════════
   ★jsdom には 幅が 無い★ので ここでは ★作りだけ★を 見る。
   ★実際の幅は 実機で 測った★（2026-08-28・幅390）＝
     直す前 … 「マクロ」と「Excel版」が 画面の外（x=389 / x=478）・横に 動かす道も 無し
     直した後 … 帯を 横に 動かせる（336px の枠に 567px）・押せた／幅1280は 前と同じ */
T('★狭い画面で 帯を 横に 動かせる作りに なっている★', () => {
  const i = book.indexOf('id="headerTools"');
  ok(i > 0, '帯に 名前が ついていない（headerTools）');
  const 帯 = book.slice(i, book.indexOf('>', i));
  ok(/overflow-x:\s*auto/.test(帯), '横に 動かせない: ' + 帯);
  ok(/min-width:\s*0/.test(帯), '縮まない作りに なっている: ' + 帯);
});

/* ══ ②-c ★③レシピにして 会社に残す★（AIは 0回） ══════════════ */
console.log('  -- 手順を 覚える（AIは 0回）--');
const 表 = () => {
  const data = { '0,0': { v: '日付' }, '0,1': { v: '名前' }, '0,2': { v: '金額' } };
  const 名 = ['え', 'あ', 'う', 'い', 'お'];
  for (let r = 1; r <= 5; r++) {
    data[r + ',0'] = { v: '8/' + r }; data[r + ',1'] = { v: 名[r - 1] }; data[r + ',2'] = { v: r * 100 };
  }
  return [{ name: '売上', data }];
};
const 手順つき = () => M.見立てる([{ 名: 'Module1', 確か: true, 中身: ['Attribute VB_Name = "Module1"',
  'Sub 月次()', '  Range("A1:C6").Sort Key1:=Range("B2"), Order1:=xlDescending',
  '  Range("D1").Value = "税込"', '  Range("D2:D6").Formula = "=C2*1.1"', 'End Sub'].join('\r\n') }]);

T('★手順が 取り出せた本だけ「覚える」ボタンを 出す★', () => {
  const { api, w } = 台({ sheets: 表() });
  api.マクロを受け取る({ 読み: { ok: true, モジュール: [] }, 見立て: 手順つき() });
  api.openMacro();
  const 行 = w.document.getElementById('mcList').children;
  eq(行.length, 1, '描かれた行数');
  const ボタン = [...行[0].querySelectorAll('button')].map((b) => b.textContent);
  eq(ボタン.join(','), 'この手順を 覚える', 'ボタン');
  const t = 行[0].textContent;
  ok(t.indexOf('手順は 2つです') >= 0, '手順の数が 出ていない: ' + t);
  ok(t.indexOf('列 B で 降順に 並べ替える') >= 0, '手順の中身が 出ていない: ' + t);
});
T('★2本目以降の 手順も 取り出す（先頭だけで 止めない）★', () => {
  /* ★同じ形のループが 画面に 2つ ある★（手順を取り出す所と 描く所）。
     片方だけ 止まっても 気づけるように ★2本 入れて 両方に ボタンが 出る事★を 数える。 */
  const 見 = M.見立てる([{ 名: 'Module1', 確か: true, 中身: ['Attribute VB_Name = "Module1"',
    'Sub 一つ目()', '  Columns("C").Delete', 'End Sub', '',
    'Sub 二つ目()', '  Range("A1:C6").Sort Key1:=Range("B2")', 'End Sub'].join('\r\n') }]);
  const { api, w } = 台({ sheets: 表() });
  api.マクロを受け取る({ 読み: { ok: true, モジュール: [] }, 見立て: 見 });
  api.openMacro();
  const 行 = w.document.getElementById('mcList').children;
  eq(行.length, 2, '描かれた行数');
  eq(行[0].querySelectorAll('button').length, 1, '1本目に ボタンが 無い');
  eq(行[1].querySelectorAll('button').length, 1, '★2本目に ボタンが 無い（先頭だけで 止まっている）★');
});
T('★手順が 取り出せない本には ボタンを 出さない★（出来ていない物のボタンを 見せない）', () => {
  const { api, w } = 台({ sheets: 表() });
  api.マクロを受け取る({ 読み: { ok: true, モジュール: [] }, 見立て: M.見立てる([{ 名: 'M', 確か: true,
    中身: ['Attribute VB_Name = "M"', 'Sub a()', '  MsgBox "hi"', 'End Sub'].join('\r\n') }]) });
  api.openMacro();
  const 行 = w.document.getElementById('mcList').children;
  eq(行.length, 1);
  eq(行[0].querySelectorAll('button').length, 0, 'ボタンが 出ている');
});
T('★覚えると 倉庫にも 入る（会社に残る）／AIは 0回★', async () => {
  const { api, 記録 } = 台({ sheets: 表() });
  api.マクロを受け取る({ 読み: { ok: true, モジュール: [] }, 見立て: 手順つき() });
  api.openMacro();
  await api.マクロの手順を覚える('Module1.月次');
  eq(記録.覚えた.length, 1, 'この画面に 覚えていない');
  eq(記録.倉庫.length, 1, '倉庫に 入れていない');
  eq(記録.倉庫[0].手順.length, 2, '覚えた手順の数');
  eq(記録.履歴.length, 1, '履歴に 残していない');
  eq(記録.履歴[0].credit, 0, '★AIを 使った事に している★');
});
T('★覚えた所では 1セルも 書かず、先に 見せる★', async () => {
  const { api, 記録 } = 台({ sheets: 表() });
  api.マクロを受け取る({ 読み: { ok: true, モジュール: [] }, 見立て: 手順つき() });
  await api.マクロの手順を覚える('Module1.月次');
  ok(記録.見せた, '見せていない');
  eq(記録.書いた, 0, '書いている');
  const 番地 = Object.keys(記録.見せた.変える);
  ok(番地.length > 0, '変える所が 空');
  /* ★見せる中身が 本当に 並べ替えと 式か★（名前だけ 出して 何もしない を 防ぐ） */
  eq(記録.見せた.変える['0,3'].v, '税込', '新しい列の 見出し');
  eq(記録.見せた.変える['1,3'].f, '=C2*1.1', '式');
  ok(記録.見せた.何.some((x) => x.indexOf('並べ替え') >= 0), '並べ替えが 無い: ' + JSON.stringify(記録.見せた.何));
});
T('★読み取れなかった所は 覚えた時にも 言う（黙って 減らさない）★', async () => {
  const { api, 記録 } = 台({ sheets: 表() });
  const 見 = M.見立てる([{ 名: 'M2', 確か: true, 中身: ['Attribute VB_Name = "M2"', 'Sub b()',
    '  Columns("C").Delete', '  ActiveSheet.Range("A1").CurrentRegion.RemoveDuplicates Columns:=1',
    'End Sub'].join('\r\n') }]);
  api.マクロを受け取る({ 読み: { ok: true, モジュール: [] }, 見立て: 見 });
  await api.マクロの手順を覚える('M2.b');
  const 字 = 記録.知らせ.map((x) => x.本文).join(' ');
  ok(字.indexOf('1つ 覚えました') >= 0, 字);
  ok(字.indexOf('読み取れなかった所が 1か所') >= 0, 字);
});
T('★倉庫に 入らなかった時は そう言う（入ったふりを しない）★', async () => {
  const { api, 記録, w } = 台({ sheets: 表() });
  api.マクロを受け取る({ 読み: { ok: true, モジュール: [] }, 見立て: 手順つき() });
  /* 倉庫が 断る形に する */
  const 台2 = 台({ sheets: 表() });
  台2.記録.倉庫 = null;
  await api.マクロの手順を覚える('Module1.月次');
  ok(記録.知らせ.length > 0, '何も 言っていない');
});

T('★知らせの 見出しと 補足は 行を 分ける（字が くっつかない）★', () => {
  /* ★実物の画面で 見つけた（2026-08-28）★＝この形は #toast.warn にしか 無く、
     普通の知らせでは「18か所 直しました覚えた手順で やりました。」と 1行に くっついていた。
     ★幅は jsdom では 測れない★ので ここでは 形が 在る事だけを 見る
     （実機では 3行に 分かれる事を 押して 確かめた）。 */
  const i = book.indexOf('#toast .toast-h');
  ok(i > 0, '普通の知らせの 見出しの形が 無い');
  const 行 = book.slice(i, book.indexOf('}', i));
  ok(/display:\s*block/.test(行), '行を 分けていない: ' + 行);
  const j = book.indexOf('#toast .toast-n');
  ok(j > 0, '普通の知らせの 補足の形が 無い');
  ok(/display:\s*block/.test(book.slice(j, book.indexOf('}', j))), '補足が 行を 分けていない');
});

/* ══ ③客に見せる字 ════════════════════════════════════ */
T('★客に見せる字に ★ を書かない★', () => {
  const { api, w, 記録 } = 台();
  api.マクロを受け取る(マクロ); api.マクロの知らせを出す(); api.openMacro();
  const 見る = [w.document.getElementById('mcList').textContent,
    w.document.getElementById('mcBody').textContent,
    w.document.getElementById('macroBtn').textContent]
    .concat(記録.知らせ.map((x) => x.題 + x.本文));
  for (const s of 見る) ok(s.indexOf('★') < 0, '★が 出ている: ' + s.slice(0, 80));
});
T('★中の言葉を そのまま 客に 出さない★（CFB・dir・stream・undefined）', () => {
  const { api, w, 記録 } = 台();
  api.マクロを受け取る(マクロ); api.マクロの知らせを出す(); api.openMacro();
  const 全部 = w.document.body.textContent + 記録.知らせ.map((x) => x.題 + x.本文).join('');
  for (const 語 of ['CFB', 'undefined', 'NaN', 'MS-OVBA', 'vbaProject', '[object']) {
    ok(全部.indexOf(語) < 0, 語 + ' が 客に 出ている');
  }
});
T('★AIを 使っていないと 言う★（押す前に 回数を 書かない）', () => {
  const { api, 記録 } = 台();
  api.マクロを受け取る(マクロ); api.マクロの知らせを出す();
  const s = 記録.知らせ[0].題 + 記録.知らせ[0].本文;
  ok(s.indexOf('AIは 使っていません') >= 0, 'AIを使っていないと 言っていない: ' + s);
});

console.log('');
console.log('  ' + pass + '緑 / ' + fail + '赤');

/* ══ ④壊して 赤くなるか ══════════════════════════════ */
if (SELF) {
  console.log('');
  console.log('[vba-ui --self-test] ★壊したら 赤くなるか★');
  const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'exally-vbaui-'));
  const BREAKS = [
    ['★知らせを「マクロが入っています」より 先に 出す★',
      (s) => s.replace("      if (typeof マクロの知らせを出す === 'function') マクロの知らせを出す();", '')],
    ['★次のファイルで 前の物を 消さない★',
      (s) => s.replace('if (typeof マクロを受け取る === \'function\') マクロを受け取る(res.hasVba ? res.マクロ : null);',
        'if (res.hasVba && typeof マクロを受け取る === \'function\') マクロを受け取る(res.マクロ);')],
    ['★マクロが 無くても ボタンを 出す★',
      (s) => s.replace('if(!見 || (!見.本数 && !未)){ b.hidden = true; return; }', '')],
    ['★一覧を 先頭1本で 打ち切る（描く所）★',
      (s) => s.replace('  for(var i=0;i<見.手続き.length;i++){\n    var v = 見.手続き[i];\n    var row = document.createElement',
        '  for(var i=0;i<1;i++){\n    var v = 見.手続き[i];\n    var row = document.createElement')],
    ['★手順を 先頭1本しか 取り出さない★',
      (s) => s.replace('    for(var i=0;i<見.手続き.length;i++){\n      var v = 見.手続き[i];\n      try{ _マクロの手順',
        '    for(var i=0;i<1;i++){\n      var v = 見.手続き[i];\n      try{ _マクロの手順')],
    ['★「ここでは どうするか」を 描かない★',
      (s) => s.replace('if(v.うちのやり方){', 'if(false){')],
    ['★未測定を 黙って 落とす★',
      (s) => s.replace("document.getElementById('mcMore').textContent = 未.length", "document.getElementById('mcMore').textContent = false")],
    ['★知らせの 見出しを 行で 分けない★',
      (s) => s.replace('#toast .toast-h { display:block;', '#toast .toast-h { display:inline;')],
    ['★狭い画面で 帯が 動かせなくなる★',
      (s) => s.replace('overflow-x:auto;overflow-y:hidden;', '')],
    ['★終わりの印が 無い事を 言わない★',
      (s) => s.replace('if(v.閉じていない){', 'if(false){')],
    ['★何をしているかを 描かない★',
      (s) => s.replace('何.textContent = v.何をしているか;', "何.textContent = '';")],
    ['★客に見せる字に ★ を書く★',
      (s) => s.replace("頭.textContent = v.モジュール + ' の '", "頭.textContent = '★' + v.モジュール + ' の '")],
    ['★中の言葉を そのまま 出す★',
      (s) => s.replace("何.textContent = v.何をしているか;", "何.textContent = v.何をしているか + ' CFB';")],
    ['★押す前に AIの回数を 書かない（言わない）★',
      (s) => s.replace('（AIは 使っていません）', '')],
  ];
  let red = 0;
  for (const [name, brk] of BREAKS) {
    const 元 = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf8');
    const 壊 = brk(元);
    if (壊 === 元) { console.log('  ★置換できず★  ' + name); continue; }
    const f = path.join(TMP, 'book.html');
    fs.writeFileSync(f, 壊, 'utf8');
    const env = Object.assign({}, process.env, { EXALLY_VBAUI_OVERRIDE: JSON.stringify({ 'book.html': f }) });
    const r = spawnSync(process.execPath, [path.join(__dirname, 'vba-ui.test.mjs')], { encoding: 'utf8', env });
    if (r.status !== 0) { red++; console.log('  赤くなった  ' + name); }
    else console.log('  ★素通り★  ' + name);
  }
  try { fs.rmSync(TMP, { recursive: true, force: true }); } catch (e) { /* 消せなくても 検査は済んでいる */ }
  console.log('');
  console.log('  ' + red + '/' + BREAKS.length + ' 通りで赤くなった');
  process.exit(red === BREAKS.length ? 0 : 1);
}

process.exit(fail ? 1 : 0);
