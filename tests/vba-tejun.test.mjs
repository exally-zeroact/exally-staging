/* vba-tejun.test.mjs — ★マクロから 手順を 取り出す（③レシピにして 会社に残す）★
 *
 *  ★決まり★
 *    ・言葉は レシピ（lib/recipe.js）の 物を そのまま使う＝2つ作らない
 *    ・★取り出せた分だけ★ 手順にする。取り出せない所は ★数と理由で 言う★
 *      （黙って 少ない手順を 返すのは「合計が 静かに 小さくなる」型）
 *    ・★当てずっぽうで 埋めない★（列の名前が 無い式は 手順にしない）
 *
 *  ★部品が緑＝使える ではない★ので、ここでは
 *    ★取り出した手順を レシピの検品に通し、実際に 表へ 当てて セルが 変わる所まで★ 測る。
 *
 *  使い方: node tests/vba-tejun.test.mjs [--self-test]
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
const OVERRIDE = process.env.EXALLY_VBATEJUN_OVERRIDE ? JSON.parse(process.env.EXALLY_VBATEJUN_OVERRIDE) : {};
const require_ = createRequire(pathToFileURL(path.join(ROOT, 'package.json')));
const T = require_(OVERRIDE['lib/vba-tejun.js'] || path.join(ROOT, 'lib/vba-tejun.js'));
const M = require_(path.join(ROOT, 'lib/vba-mikata.js'));
const R = require_(OVERRIDE['lib/recipe.js'] || path.join(ROOT, 'lib/recipe.js'));

let pass = 0, fail = 0;
const T_ = (n, fn) => { try { fn(); pass++; console.log('  ok   ' + n); } catch (e) { fail++; console.log('  NG   ' + n + '\n       ' + (e && e.message)); } };
const ok = (c, m) => { if (!c) throw new Error(m || 'expected truthy'); };
const eq = (a, b, m) => { if (a !== b) throw new Error((m ? m + ': ' : '') + '期待=' + JSON.stringify(b) + ' 実際=' + JSON.stringify(a)); };

/** VBAの字 → 手続き1本 → 取り出す */
const 読む = (行たち) => T.取り出す(M.手続きに切る(行たち.join('\r\n'))[0]);

console.log('');
console.log('[vba-tejun] ★マクロから 手順を 取り出す（AIは 0回）★');

/* ══ ①1つずつ 読めるか ═══════════════════════════════════ */
console.log('  -- 1つずつ --');
T_('並べ替え（Key1 が Range）', () => {
  const r = 読む(['Sub a()', '  Range("A1:D9").Sort Key1:=Range("B2"), Order1:=xlDescending', 'End Sub']);
  eq(r.手順.length, 1);
  eq(r.手順[0].種類, '並べ替え');
  eq(r.手順[0].列, 'B');
  eq(r.手順[0].向き, '降順');
});
T_('並べ替え（向きを書いていなければ 昇順）', () => {
  const r = 読む(['Sub a()', '  Range("A1:D9").Sort Key1:=Range("C2")', 'End Sub']);
  eq(r.手順[0].向き, '昇順');
});
T_('並べ替え（Key が Cells の 番号）', () => {
  const r = 読む(['Sub a()', '  ws.Sort.SortFields.Add Key:=Cells(2, 3), Order:=xlAscending', 'End Sub']);
  eq(r.手順[0].種類, '並べ替え');
  eq(r.手順[0].列, 'C', '3列目は C');
});
T_('★並べ替える列が 書いていなければ 手順にしない★（理由を残す）', () => {
  const r = 読む(['Sub a()', '  ActiveSheet.Range("A1").Sort Header:=xlYes', 'End Sub']);
  eq(r.手順.length, 0);
  eq(r.取り出せなかった.length, 1);
  ok(r.取り出せなかった[0].なぜ.indexOf('列') >= 0, r.取り出せなかった[0].なぜ);
});
T_('列を消す（字・番号・A:A の 3通り）', () => {
  const 例 = [['  Columns("C").Delete', 'C'], ['  Columns(4).Delete', 'D'], ['  Range("B:B").Delete', 'B']];
  for (const [行, 期待] of 例) {
    const r = 読む(['Sub a()', 行, 'End Sub']);
    eq(r.手順.length, 1, 行);
    eq(r.手順[0].種類, '列を消す', 行);
    eq(r.手順[0].列, 期待, 行);
  }
});
T_('★行を消すのは まだ 出来ないと 言う（黙って 列と間違えない）★', () => {
  const r = 読む(['Sub a()', '  Rows(5).Delete', 'End Sub']);
  eq(r.手順.length, 0);
  ok(r.取り出せなかった[0].なぜ.indexOf('行を消す') >= 0, r.取り出せなかった[0].なぜ);
});
T_('見出し＋式 → 「式の列を足す」1つに まとまる', () => {
  const r = 読む(['Sub a()', '  Range("E1").Value = "税込"', '  Range("E2:E9").Formula = "=D2*1.1"', 'End Sub']);
  eq(r.手順.length, 1, '2つに 割れている');
  eq(r.手順[0].種類, '式の列を足す');
  eq(r.手順[0].見出し, '税込');
  eq(r.手順[0].式, '=D{行}*1.1');
});
T_('見出しが Cells(1, n) でも 読める', () => {
  const r = 読む(['Sub a()', '  Cells(1, 5).Value = "税込"', '  Range("E2:E9").Formula = "=D2*1.1"', 'End Sub']);
  eq(r.手順[0].見出し, '税込');
});
T_('★式の中の 他の数字を 壊さない★（置き換えるのは 始まりの行だけ）', () => {
  const r = 読む(['Sub a()', '  Range("E1").Value = "合計"', '  Range("E2:E99").Formula = "=SUM(D2:D99)+100"', 'End Sub']);
  eq(r.手順[0].式, '=SUM(D{行}:D99)+100');
});
T_('★列の名前が どこにも 無い式は 手順にしない★（当てずっぽうで 名づけない）', () => {
  const r = 読む(['Sub a()', '  Range("G2:G9").Formula = "=A2+B2"', 'End Sub']);
  eq(r.手順.length, 0);
  eq(r.取り出せなかった.length, 1);
  ok(r.取り出せなかった[0].なぜ.indexOf('名前') >= 0, r.取り出せなかった[0].なぜ);
});
T_('見出しだけなら「列の名前を変える」', () => {
  const r = 読む(['Sub a()', '  Range("C1").Value = "担当"', 'End Sub']);
  eq(r.手順.length, 1);
  eq(r.手順[0].種類, '列の名前を変える');
  eq(r.手順[0].元, 'C');
  eq(r.手順[0].新, '担当');
});
T_('知らない書き方は 理由つきで 落とす', () => {
  const r = 読む(['Sub a()', '  ActiveSheet.Range("A1").CurrentRegion.RemoveDuplicates Columns:=1', 'End Sub']);
  eq(r.手順.length, 0);
  eq(r.取り出せなかった.length, 1);
});
T_('写して貼る（丸ごと／値だけ）', () => {
  const a = 読む(['Sub a()', '  Range("A1:D9").Copy Destination:=Range("F1")', 'End Sub']);
  eq(a.手順.length, 1);
  eq(a.手順[0].種類, '写す'); eq(a.手順[0].元, 'A1:D9'); eq(a.手順[0].先, 'F1');
  eq(a.手順[0].値だけ, false, '丸ごとの はずが 値だけに なっている');
  const b = 読む(['Sub b()', '  Range("F1:I9").Value = Range("A1:D9").Value', 'End Sub']);
  eq(b.手順[0].元, 'A1:D9', '★元と 先が 逆★（VBAは 左が 貼り先）');
  eq(b.手順[0].先, 'F1');
  eq(b.手順[0].値だけ, true);
});
T_('★別のシートへ 写す物は 断る（同じシートに 当てない）★', () => {
  const 例 = ['  Sheets("集計").Range("A1:D9").Value = Sheets("明細").Range("A1:D9").Value',
    '  Range("A1:D9").Copy Destination:=Sheets("集計").Range("A1")'];
  for (const 行 of 例) {
    const r = 読む(['Sub c()', 行, 'End Sub']);
    eq(r.手順.length, 0, 行);
    ok(r.取り出せなかった[0].なぜ.indexOf('別のシート') >= 0, 行 + ' → ' + r.取り出せなかった[0].なぜ);
  }
});
T_('写す先が 変数なら 手順にしない（理由を残す）', () => {
  const r = 読む(['Sub d()', '  Range("A1:D9").Copy Destination:=Cells(行, 1)', 'End Sub']);
  eq(r.手順.length, 0);
  ok(r.取り出せなかった[0].なぜ.indexOf('写す先') >= 0, r.取り出せなかった[0].なぜ);
});
T_('★入れ物を置く行・条件の行・With は 母数に 入れない★（読めない行に 数えない）', () => {
  /* ★これを 数えると「読み取れない所が 5か所」と 嘘の数が 出る★＝
     Set / If / With / End With は ★表を いじっている行ではない★。 */
  const r = 読む(['Sub a()', '  Set ws = Sheets("集計")', '  With ActiveSheet',
    '    If Cells(2, 1).Value = "" Then Exit Sub', '    Columns("C").Delete', '  End With', 'End Sub']);
  eq(r.数.変える行, 1, '効く行（Columns("C").Delete の 1つだけ）');
  eq(r.数.読めない行, 0, '読めない行');
  eq(r.数.手順, 1);
});
T_('★読んでいるだけの行は 母数に 入れない★（実物で 42行の 半分ちかくが これだった）', () => {
  /* Set（日本語の変数名）／For の 頭／最終行を 取る行 … ★1セルも 変えていない★ */
  const r = 読む(['Sub a()', '  Set 集計シート = Sheets("集計")',
    '  For i = 2 To 集計シート.Cells(Rows.Count, 1).End(xlUp).Row',
    '    最終行 = 集計シート.Cells(Rows.Count, 4).End(xlUp).Row',
    '    担当 = 集計シート.Cells(1, 3).Value',
    '  Next i', '  Columns("C").Delete', 'End Sub']);
  eq(r.数.変える行, 1, '変える行（Columns("C").Delete の 1つだけ）');
  eq(r.数.読めない行, 0, '読めない行');
});
T_('★書く行は 数える（右側で 読んでいても 左側が セルなら 変える行）★', () => {
  const r = 読む(['Sub a()', '  ws.Cells(行, 3).Value = ws.Cells(行, 2).Value', 'End Sub']);
  eq(r.数.変える行, 1);
  eq(r.数.読めない行, 1, '読めない行');
  eq(r.数.手順, 0);
});
T_('段取りの行（ちらつき止め・Dim・コメント）は 母数に 入れない', () => {
  const r = 読む(['Sub a()', '  Dim i As Long', "  ' メモ", '  Application.ScreenUpdating = False',
    '  Columns("C").Delete', '  Application.ScreenUpdating = True', 'End Sub']);
  eq(r.数.変える行, 1, '変える行');
  eq(r.数.手順, 1);
});

/* ══ ②数が 合っているか（★減った事を 隠さない★） ═════════════ */
console.log('  -- 数 --');
T_('★効く行 ＝ 読めた行 ＋ 読めない行★', () => {
  const r = 読む(['Sub a()', '  Range("A1:D9").Sort Key1:=Range("B2")', '  Columns("C").Delete',
    '  ActiveSheet.Range("A1").CurrentRegion.RemoveDuplicates Columns:=1', 'End Sub']);
  eq(r.数.変える行, 3);
  eq(r.数.読めた行 + r.数.読めない行, r.数.変える行, '足して 合わない');
  eq(r.数.読めない行, 1);
});
T_('★行の数と 手順の数を 混ぜない★（2行で 1つの手順になる）', () => {
  const r = 読む(['Sub a()', '  Range("E1").Value = "税込"', '  Range("E2:E9").Formula = "=D2*1.1"', 'End Sub']);
  eq(r.数.変える行, 2, '変える行');
  eq(r.数.手順, 1, '手順');
});
T_('★読み取れない所が 在れば 1行で そう言う★', () => {
  const r = 読む(['Sub a()', '  Columns("C").Delete', '  Range("A1").CurrentRegion.RemoveDuplicates Columns:=1', 'End Sub']);
  const s = T.知らせの字(r);
  ok(s.indexOf('2か所のうち 1か所') >= 0, s);
  ok(s.indexOf('残り 1か所') >= 0, s);
  ok(s.indexOf('この 1つだけ') >= 0, s);
});
T_('1つも 取り出せない時は そう言う（0本と 言い切らない）', () => {
  const r = 読む(['Sub a()', '  ActiveSheet.Range("A1").CurrentRegion.RemoveDuplicates Columns:=1', 'End Sub']);
  const s = T.知らせの字(r);
  ok(s.indexOf('1か所 あります') >= 0, s);
  ok(s.indexOf('覚えられる書き方では ありません') >= 0, s);
});
T_('表をいじる所が 無い時', () => {
  const r = 読む(['Sub a()', '  MsgBox "こんにちは"', 'End Sub']);
  eq(r.数.変える行, 0);
  ok(T.知らせの字(r).indexOf('見つかりません') >= 0, T.知らせの字(r));
});

/* ══ ③★通し★＝レシピの検品を 通り、実際に 表が 変わるか ═══════ */
console.log('  -- 通し（取り出す → 検品 → 表に当てる）--');
const 表 = () => {
  const data = { '0,0': { v: '日付' }, '0,1': { v: '名前' }, '0,2': { v: '金額' }, '0,3': { v: 'メモ' } };
  const 名 = ['え', 'あ', 'う', 'い', 'お', 'か', 'き', 'く'];
  for (let r = 1; r <= 8; r++) {
    data[r + ',0'] = { v: '8/' + r };
    data[r + ',1'] = { v: 名[r - 1] };
    data[r + ',2'] = { v: r * 100 };
    data[r + ',3'] = { v: 'x' };
  }
  return { name: '売上', data };
};
const マクロ = ['Sub 月次()', '  Application.ScreenUpdating = False',
  '  Range("A1:D9").Sort Key1:=Range("B2"), Order1:=xlDescending',
  '  Columns("D").Delete', '  Range("E1").Value = "税込"',
  '  Range("E2:E9").Formula = "=C2*1.1"', '  MsgBox "できました"', 'End Sub'];

T_('★取り出した手順が レシピの検品を 1つも 落ちずに 通る★', () => {
  const 出 = 読む(マクロ);
  eq(出.手順.length, 3, '取り出した数');
  const 読 = R.手順を読む(JSON.stringify({ 手順: 出.手順 }));
  eq(読.ok, true, 読.なぜ);
  eq(読.手順.length, 3, '検品を 通った数');
  eq((読.断った || []).length, 0, '断られた: ' + JSON.stringify(読.断った));
});
T_('★実際に 表が 変わる（並べ替えは 行が 動く・数を 数える）★', () => {
  const sh = 表();
  const 出 = 読む(マクロ);
  const 当 = R.手順を当てる(sh, R.手順を読む(JSON.stringify({ 手順: 出.手順 })).手順, { 見出しの行: 0 });
  /* 並べ替え … 名前の列が 降順に なっているか（★動いた数ではなく 並びを 見る★） */
  const 並び = [];
  for (let r = 1; r <= 8; r++) {
    const c = 当.変える[r + ',1'];
    並び.push(c ? c.v : sh.data[r + ',1'].v);
  }
  eq(並び.join(''), 'くきかおえういあ', '降順に なっていない');
  /* 列を消す … D列（3）が 空になるか */
  eq(当.変える['1,3'].v, '', '列Dが 空になっていない');
  /* 式の列 … 見出しと 式が 入るか */
  eq(当.変える['0,4'].v, '税込', '見出し');
  eq(当.変える['1,4'].f, '=C2*1.1', '2行目の式');
  eq(当.変える['8,4'].f, '=C9*1.1', '9行目の式');
});
T_('★手順が 0本の時は 表を 1セルも 触らない★', () => {
  const sh = 表();
  const 出 = 読む(['Sub a()', '  MsgBox "なにもしない"', 'End Sub']);
  eq(出.手順.length, 0);
  const 当 = R.手順を当てる(sh, 出.手順, { 見出しの行: 0 });
  eq(Object.keys(当.変える).length, 0, '触っている');
});

T_('★写す＝実際に セルが 写る（丸ごとは 式も 運ぶ・値だけは 値に する）★', () => {
  const sh = { name: '売上', data: { '0,0': { v: 'あ' }, '0,1': { v: 1 }, '1,0': { f: '=A1' }, '1,1': { v: 2 } } };
  const 出 = 読む(['Sub a()', '  Range("A1:B2").Copy Destination:=Range("D1")', 'End Sub']);
  const 読 = R.手順を読む(JSON.stringify({ 手順: 出.手順 }));
  eq(読.ok, true, 読.なぜ);
  const 当 = R.手順を当てる(sh, 読.手順, { 見出しの行: 0 });
  eq(Object.keys(当.変える).length, 4, '写した数');
  eq(当.変える['0,3'].v, 'あ', 'D1');
  eq(当.変える['1,3'].f, '=A1', '★式を 値に 潰している★');
  const 出2 = 読む(['Sub b()', '  Range("D1:E2").Value = Range("A1:B2").Value', 'End Sub']);
  const 当2 = R.手順を当てる(sh, R.手順を読む(JSON.stringify({ 手順: 出2.手順 })).手順, { 見出しの行: 0 });
  eq(当2.変える['1,3'].f, undefined, '★値だけの はずが 式を 運んでいる★');
});
T_('★手順は 上から 順に 起きる（並べ替えた後の 物を 写す）★', () => {
  /* ★実機で 押して 見つけた（2026-08-28）★＝写す時に 元の表だけを 見ていたので、
     並べ替えの 後に 写しても ★並べ替える前の 並び★を 写していた。 */
  const data = { '0,0': { v: '名前' }, '0,1': { v: '金額' } };
  const 名 = ['え', 'あ', 'う', 'い', 'お'];
  for (let r = 1; r <= 5; r++) { data[r + ',0'] = { v: 名[r - 1] }; data[r + ',1'] = { v: r * 100 }; }
  const 当 = R.手順を当てる({ name: 'a', data },
    [{ 種類: '並べ替え', 列: 'A', 向き: '降順' }, { 種類: '写す', 元: 'A1:B6', 先: 'D1' }], { 見出しの行: 0 });
  const 並べた = [], 写した = [];
  for (let r = 1; r <= 5; r++) {
    並べた.push(当.変える[r + ',0'] ? 当.変える[r + ',0'].v : data[r + ',0'].v);
    写した.push(当.変える[r + ',3'] ? 当.変える[r + ',3'].v : '(無)');
  }
  eq(並べた.join(','), 'お,え,う,い,あ', '並べ替えが 効いていない');
  eq(写した.join(','), 並べた.join(','), '★写す時に 並べ替える前の 物を 見ている★');
});
T_('★写す先が 読めない手順は レシピの検品で 断られる★', () => {
  const 読 = R.手順を読む(JSON.stringify({ 手順: [{ 種類: '写す', 元: 'A:A', 先: 'D1' }] }));
  eq(読.ok, false, '通してしまっている');
  ok((読.断った || []).join('').indexOf('写す') >= 0, JSON.stringify(読.断った));
});

/* ══ ④客に見せる字 ═══════════════════════════════════════ */
console.log('  -- 客に見せる字 --');
T_('★客に見せる字に ★ を書かない★', () => {
  const r = 読む(マクロ);
  const 見る = [T.知らせの字(r)].concat(r.手順.map((t) => T.手順の字(t)))
    .concat(r.取り出せなかった.map((x) => x.なぜ));
  for (const s of 見る) ok(s.indexOf('★') < 0, '★が 入っている: ' + s);
});
T_('★「壊れています」と言わない★', () => {
  const r = 読む(['Sub a()', '  Rows(5).Delete', 'End Sub']);
  const 見る = [T.知らせの字(r)].concat(r.取り出せなかった.map((x) => x.なぜ));
  for (const s of 見る) ok(s.indexOf('壊れ') < 0, s);
});
T_('手順の字が 1つずつ 日本語で 出る', () => {
  const r = 読む(マクロ);
  const 字 = r.手順.map((t) => T.手順の字(t));
  ok(字[0].indexOf('列 B で 降順') >= 0, 字[0]);
  ok(字[1].indexOf('列 D を 空') >= 0, 字[1]);
  ok(字[2].indexOf('税込') >= 0, 字[2]);
  for (const s of 字) ok(s.length > 4, '短すぎる: ' + s);
});

console.log('');
console.log('  ' + pass + '緑 / ' + fail + '赤');

/* ══ ⑤壊して 赤くなるか ══════════════════════════════════ */
if (SELF) {
  console.log('');
  console.log('[vba-tejun --self-test] ★壊したら 赤くなるか★');
  const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'exally-vbatejun-'));
  const BREAKS = [
    ['vba-tejun.js', '★読み取れない行を 黙って 落とす★',
      (s) => s.replace("      取り出せなかった.push({ 行: i + 1, なぜ: 'この書き方は まだ 手順に 出来ません' });", '')],
    ['vba-tejun.js', '★列の名前が 無い式に 勝手に 名前を つける★',
      (s) => s.replace("      if (!名) { 取り出せなかった.push({ 行: w.行, なぜ: 'この式の 列の名前が どこにも 書かれていません' }); continue; }",
        "      if (!名) 名 = '列' + w.列;")],
    ['vba-tejun.js', '★式の中の 数字を 全部 置き換える★',
      (s) => s.replace("var 直した = 式.replace(new RegExp('(\\\\$?[A-Z]{1,3}\\\\$?)' + 始まり + '(?![0-9])', 'g'), '$1{行}');",
        "var 直した = 式.replace(/(\\$?[A-Z]{1,3}\\$?)\\d+/g, '$1{行}');")],
    ['vba-tejun.js', '★行を消すのを 列を消すと 間違える★',
      (s) => s.replace("    if (/Rows\\(|EntireRow/i.test(行)) return { だめ: '行を消す作業は まだ 手順に 出来ません' };", '')],
    ['recipe.js', '★写す時に 1つ前の手順の 結果を 見ない（元の表を 写す）★',
      (s) => s.replace('var から = Object.prototype.hasOwnProperty.call(変える, 元番地) ? 変える[元番地] : data[元番地];',
        'var から = data[元番地];')],
    ['vba-tejun.js', '★写す 元と 先を 逆にする★',
      (s) => s.replace("if (v) return { 手順: { 種類: '写す', 元: v[2].toUpperCase(), 先: String(v[1]).toUpperCase().split(':')[0], 値だけ: true } };",
        "if (v) return { 手順: { 種類: '写す', 元: v[1].toUpperCase().split(':')[0], 先: v[2].toUpperCase(), 値だけ: true } };")],
    ['vba-tejun.js', '★別のシートへ 写す物を 同じシートに 当てる★',
      (s) => s.replace("    if (よそのシート.test(l)) return { だめ: '別のシートへ 写す物は まだ 手順に 出来ません' };", '')],
    ['vba-tejun.js', '★値だけ の 印を 落とす（式を 運んでしまう）★',
      (s) => s.replace("値だけ: true } };", "値だけ: false } };")],
    ['vba-tejun.js', '★並べ替えの 向きを 見ない（いつも 昇順）★',
      (s) => s.replace('var 降順 = /Order1?\\s*:=\\s*xlDescending/i.test(行);', 'var 降順 = false;')],
    ['vba-tejun.js', '★列が 書いていない並べ替えを 手順にしてしまう★',
      (s) => s.replace("      if (!c) return { だめ: '並べ替える列が 書かれていません' };",
        "      if (!c) return { 手順: { 種類: '並べ替え', 列: 'A', 向き: '昇順' } };")],
    ['vba-tejun.js', '★読むだけの行まで 母数に 入れる★',
      (s) => s.replace('    if (読み飛ばす.test(l)) return false;', '')],
    ['vba-tejun.js', '★代入の 右側だけで 変える行と 決める★（読むだけの行を 数える）',
      (s) => s.replace('    return 場所.test(l.slice(0, i));', '    return 場所.test(l);')],
    ['vba-tejun.js', '★行の数と 手順の数を 同じにする★',
      (s) => s.replace('        手順: 手順.length,', '        手順: Math.max(0, 効いた - 取り出せなかった.length),')],
    ['vba-tejun.js', '★読み取れなかった事を 客に 言わない★',
      (s) => s.replace('    if (出.数.読めない行) {', '    if (false) {')],
    ['vba-tejun.js', '★1つも 取り出せない時に 出来たように 言う★',
      (s) => s.replace("      return '表を 変える所が ' + 出.数.変える行 + 'か所 ありますが、'", "      return '手順に 出来ました。'; //")],
  ];
  let red = 0;
  for (const [対象, name, brk] of BREAKS) {
    const 元 = fs.readFileSync(path.join(ROOT, 'lib', 対象), 'utf8');
    const 壊 = brk(元);
    if (壊 === 元) { console.log('  ★置換できず★  ' + name); continue; }
    const f = path.join(TMP, 対象);
    fs.writeFileSync(f, 壊, 'utf8');
    const env = Object.assign({}, process.env, { EXALLY_VBATEJUN_OVERRIDE: JSON.stringify({ ['lib/' + 対象]: f }) });
    const r = spawnSync(process.execPath, [path.join(__dirname, 'vba-tejun.test.mjs')], { encoding: 'utf8', env });
    if (r.status !== 0) { red++; console.log('  赤くなった  ' + name); }
    else console.log('  ★素通り★  ' + name);
  }
  try { fs.rmSync(TMP, { recursive: true, force: true }); } catch (e) { /* 消せなくても 検査は済んでいる */ }
  console.log('');
  console.log('  ' + red + '/' + BREAKS.length + ' 通りで赤くなった');
  process.exit(red === BREAKS.length ? 0 : 1);
}

process.exit(fail ? 1 : 0);
