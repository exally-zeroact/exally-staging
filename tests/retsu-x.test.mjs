/* retsu-x.test.mjs — ★列まるごと・行まるごと（A:A・1:1）を 台が 自分で 読む★（2026-11-04・案X）
 *
 *  ★★本番に 在った 壊れ（17fcf11・WebKit で 客の 画面を 実測）★★
 *    A1:A5=1..5／B1:B5=10..50 で `=SUM(A:A)` → #REF!（実Excel 15）・`=COUNTA(B:B)` → 1（5）・`=COUNT(A:A)` → 0（5）
 *    ＝台は `A:A` の 端を 番地に 出来ず、誤りか ★黙った 数★ を 返して いた。
 *
 *  ★★直し★★（lib/shiki-hyou.js の 帯の形・帯を読む・帯の置き場・見る者たち）
 *    ・帯は 板の 使って いる 端（＋空の 1行・1列）まで 切って 読む＝★端は 板で 1つ・増えるだけ★
 *    ・★切っても 答えが 変わらない 所★（SUM 等・完全一致の 探し）だけ 読む／大きさで 変わる 物は 本当の 大きさ（ROWS・COLUMNS・COUNTBLANK・INDEX）
 *    ・それ 以外（演算子の 中・UNIQUE・SORT・OFFSET・空に 当たる 条件 等）は ★前と 同じ★
 *    ・網＝帯の 見張り（列・行・板）＋直す群 の 待ち数を 同じ 辺で 数える
 *
 *  ★赤から 始めた★（2026-11-04）直す前の 台（fce13a0）で 回すと 物差し 43本の 合うは 0・網の 試験は 全部 赤。
 *  物差し: tests/fixtures/retsu-marugoto-excel.json（43本）・tests/fixtures/retsu-x-tsuika-excel.json（28本）＝実Excel 16.0
 *  使い方: node tests/retsu-x.test.mjs
 */
import path from 'node:path';
import fs from 'node:fs';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';

const ここ = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(ここ, '..');
const require_ = createRequire(import.meta.url);
const H = require_(path.join(ROOT, 'lib/shiki-hyou.js'));

let pass = 0, fail = 0;
const T = (n, よい, 添え) => {
  if (よい) { pass++; console.log('  ✓ ' + n); }
  else { fail++; console.log('  ✗ ' + n + (添え !== undefined ? ' — ' + 添え : '')); }
};

/* ★まだ（実Excel と 違う・訳つき）★＝字は ★今の 本番と 同じ か 誤りの 印★（黙った 数を 新しく 作って いない 事を 見る） */
const まだ = {
  '=COUNTIF(A:A,"")': ['#REF!', '空に 当たる 条件＝台の 条件が 空の マスで ずれる（ふつうの 四角でも・別の 1件）'],
  '=COUNTIF(A:A,"<>3")': ['#REF!', '同じ'],
  '=SUMPRODUCT(--(A:A=""))': ['#REF!', '演算子の 中の 帯＝本当の 大きさの 配列が 要る（押し2）'],
  '=ROWS(UNIQUE(A:A))': ['1', '★今の 本番と 同じ 黙った 数★（UNIQUE が 帯を 受けない・押し2）'],
  '=INDEX(UNIQUE(A:A),6)': ['#REF!', 'UNIQUE（押し2）'],
  '=ROWS(SORT(A:A))': ['1', '★今の 本番と 同じ 黙った 数★（SORT・押し2）'],
  '=INDEX(SORT(A:A),1)': ['#REF!', 'SORT（押し2）'],
  '=ROWS(FILTER(A:A,A:A>2))': ['#CALC!', 'FILTER と 演算子（押し2）'],
  '=ROWS(OFFSET(A:A,0,1))': ['#REF!', 'OFFSET（押し2）'],
  '=SUM(OFFSET(A:A,0,1))': ['#REF!', 'OFFSET（押し2）'],
  '=ROW(A:A)': ['#VALUE!', '溢れが 板の 端を 越える（押し2）'],
  '=A:A': ['#REF!', '同じ'],
  '=XLOOKUP(D9,A:A,B:B)': ['#N/A', '空の 鍵の XLOOKUP＝ふつうの 四角でも 台 #N/A／Excel 0（別の 1件）'],
  '=XLOOKUP(D9,Sheet1!A1:A6,Sheet1!B1:B6)': ['#N/A', '帯で ない（ふつうの 四角の ずれ・別の 1件）'],
  '=COUNTIF(Sheet1!A1:A6,"")': ['0', '帯で ない（同じ）'],
  '=COUNTIF(Sheet1!A1:A6,"<>3")': ['4', '帯で ない（同じ）'],
  '=INDEX(A:A,0)': ['#REF!', 'INDEX の 0＝まるごとの 参照（押し2）'],
  '=SUM(INDEX(A:A,0))': ['#REF!', '同じ'],
};

/* ★実Excel の 置き方と 同じ★＝式は 別の 板 CALC・材料は SHEET1（板の 名の 無い 帯は Sheet1! に 書き換える＝tools/retsu-marugoto-excel.ps1） */
/* ★後読みは 使わない★（古い Safari が 読めない＝no-lookbehind の 決まり）＝前の 字を 捕まえて そのまま 戻す */
const 書換 = f => f.replace(/(^|[^!A-Za-z$])(\$?[A-Z]{1,3}:\$?[A-Z]{1,3}|\$?\d+:\$?\d+)(?![0-9(])/g, '$1Sheet1!$2').replace(/Sheet1!Sheet1!/g, 'Sheet1!');
function 板を作る() {
  const h = H.表();
  h.板を足す('SHEET1'); h.板を足す('CALC');
  for (let i = 1; i <= 5; i++) { h.打つ('SHEET1!A' + i, String(i)); h.打つ('SHEET1!B' + i, String(i * 10)); }
  h.板('CALC');
  return h;
}

console.log('');
console.log('[retsu-x] ★列まるごと・行まるごとを 台が 読む（実Excel の 物差し）★');
for (const 紙 of ['retsu-marugoto-excel.json', 'retsu-x-tsuika-excel.json']) {
  const j = JSON.parse(fs.readFileSync(path.join(ROOT, 'tests/fixtures', 紙), 'utf8'));
  for (const r of j.行) {
    const h = 板を作る();
    h.打つ('CALC!A3', 書換(r.式));
    const v = h.字('CALC!A3');
    const e = r.Excel === '#スピル!' ? '#SPILL!' : r.Excel;
    if (まだ[r.式]) T('まだ（' + まだ[r.式][1] + '）... ' + r.式, v === まだ[r.式][0], '字=' + v + ' Excel=' + r.Excel);
    else T('実Excel と 同じ ... ' + r.式 + ' = ' + e, v === e, '字=' + v);
  }
}

console.log('');
console.log('[retsu-x] ★網＝帯を 見る 式が 古い 答えで 残らない★');
function 材料() { const h = H.表(); for (let i = 1; i <= 5; i++) { h.打つ('A' + i, String(i)); h.打つ('B' + i, String(i * 10)); } return h; }
{
  const h = 材料(); h.打つ('D1', '=SUM(A:A)'); h.打つ('E1', '=D1*2');
  h.打つ('A6', '10');
  T('A6 を 後から 打つ → SUM(A:A) 25', h.字('D1') === '25', h.字('D1'));
  T('鎖 =D1*2 → 50', h.字('E1') === '50', h.字('E1'));
  h.打つ('A3', '');
  T('A3 を 消す → 22', h.字('D1') === '22', h.字('D1'));
}
{
  const h = 材料(); h.打つ('A3', '=C3*2'); h.打つ('D1', '=SUM(A:A)'); h.打つ('C3', '100');
  T('A3 が 式・C3 を 変える → 212（頼りの 順）', h.字('D1') === '212', h.字('D1'));
}
{
  const h = 材料(); h.打つ('D1', '=SUM(A:A)'); h.打つ('C1', '={1;2;3}'); h.打つ('A7', '=C1:C3');
  T('溢れが A 列を 埋める → 21', h.字('D1') === '21', h.字('D1'));
  h.打つ('A7', '');
  T('溢れを 消す → 15', h.字('D1') === '15', h.字('D1'));
}
{
  const h = 材料(); h.まとめて打つ始め(); h.打つ('D1', '=SUM(A:A)'); h.打つ('A6', '5'); h.まとめて打つ終わり();
  T('まとめて 打つ → 20', h.字('D1') === '20', h.字('D1'));
}
{
  const h = H.表(); h.板を足す('S2'); h.打つ('D1', '=SUM(S2!A:A)'); h.打つ('S2!A9', '7');
  T('板またぎ S2!A:A → 7', h.字('D1') === '7', h.字('D1'));
}
{
  const h = 材料(); h.打つ('D9', '=SUM(1:1)'); h.打つ('Z1', '100');
  T('行まるごと 1:1 に 後から Z1 → 111', h.字('D9') === '111', h.字('D9'));
}
{
  const h = 材料(); h.打つ('D1', '=VLOOKUP(7,A:B,2,FALSE)'); h.打つ('A7', '7'); h.打つ('B7', '77');
  T('VLOOKUP の 先を 後から 足す → 77', h.字('D1') === '77', h.字('D1'));
}
{
  const h = 材料(); h.打つ('D1', '=INDEX(A:A,9)');
  T('INDEX(A:A,9) 端の 外 → 0', h.字('D1') === '0', h.字('D1'));
  h.打つ('A9', '99');
  T('A9 を 後から → 99', h.字('D1') === '99', h.字('D1'));
}
{
  const h = 材料(); h.打つ('D1', '=SUM(A:A)'); h.置く('A6', 5);
  T('置く でも 直る → 20', h.字('D1') === '20', h.字('D1'));
}
{
  const h = 材料(); h.打つ('D1', '=COUNTBLANK(A:A)'); h.打つ('Z9000', '1');
  T('端が 伸びても COUNTBLANK(A:A) は 1048571', h.字('D1') === '1048571', h.字('D1'));
}
{
  const h = 材料(); h.打つ('D1', '=SUM(A:A)'); h.打つ('D1', '=SUM(B:B)'); h.打つ('A6', '100');
  T('見張りを 外す（打ち直した 式は A 列を もう 見ない）→ 150', h.字('D1') === '150', h.字('D1'));
}

console.log('');
console.log('[retsu-x] ★設計前 taiketsu（2026-11-04）の 反例★');
{
  const h = 材料();
  h.打つ('D1', '=SUMIF(A:A,">2",B1:B3)'); h.打つ('D2', '=SUMIF(A:A,">2",B1)'); h.打つ('D3', '=SUMIF(A:A,">2",B:B)');
  T('止め2 SUMIF 帯と 書いた 範囲の 混ざり → 前と 同じ #REF!（黙った 30 に しない）', h.字('D1') === '#REF!', h.字('D1'));
  T('止め2 SUMIF 帯と 1マス → #REF!（黙った 0 に しない）', h.字('D2') === '#REF!', h.字('D2'));
  T('止め2 SUMIF 帯どうし → 120', h.字('D3') === '120', h.字('D3'));
}
{
  /* テスト線前 taiketsu（2回目）＝帯が 2列 以上 */
  const h = 材料(); for (let i = 1; i <= 5; i++) h.打つ('C' + i, String(i * 100));
  const 止め = ['=SUMIF(A:A,">2",B:C)', '=SUMIF(A:B,">2",C:C)', '=XLOOKUP(3,A:A,B:C)'];
  止め.forEach((f, i) => { h.打つ('F' + (i + 1), f); T('2列の 帯 → #REF!（黙った 数に しない）... ' + f, h.字('F' + (i + 1)) === '#REF!', h.字('F' + (i + 1))); });
  /* 2列の 帯を 受ける 物は ★書いた 範囲（A1:C5 等＝尻尾の 空を 多めに 足した 形）と 同じ 答え★ */
  const 対 = [['=SUMIFS(B:C,A:B,">2")', '=SUMIFS(B1:C9,A1:B9,">2")'], ['=COUNTIFS(A:B,">2")', '=COUNTIFS(A1:B9,">2")'],
    ['=SUMPRODUCT(A:B,B:C)', '=SUMPRODUCT(A1:B9,B1:C9)'], ['=VLOOKUP(3,A:C,3,FALSE)', '=VLOOKUP(3,A1:C9,3,FALSE)'],
    ['=INDEX(A:C,4,3)', '=INDEX(A1:C9,4,3)'], ['=SUM(A:C)', '=SUM(A1:C9)'], ['=COUNTBLANK(A:B)-2097142', '=COUNTBLANK(A1:B9)-8']];
  対.forEach(([帯f, 書f], i) => {
    h.打つ('G' + (i + 1), 帯f); h.打つ('H' + (i + 1), 書f);
    T('2列の 帯 ＝ 書いた 範囲 ... ' + 帯f + ' = ' + h.字('H' + (i + 1)), h.字('G' + (i + 1)) === h.字('H' + (i + 1)), h.字('G' + (i + 1)));
  });
}
{
  /* 止め3＝打ってから 直した 答えと 作り直した 答えが 同じ（端は 帯の 列の 中身だけで 決まる） */
  const 式ら = ['=MATCH(12,A:A,1)', '=MATCH(4,A:A,1)', '=SUM(A:A)', '=COUNTBLANK(A:A)', '=INDEX(A:A,12)', '=MATCH(10,A:A,0)'];
  const h = 材料(); h.打つ('A10', '10');
  式ら.forEach((f, i) => h.打つ('D' + (i + 1), f));
  h.打つ('Y18', 'x'); h.打つ('Q20000', '1'); h.打つ('Q20000', '');
  const g = 材料(); g.打つ('A10', '10'); g.打つ('Y18', 'x');
  式ら.forEach((f, i) => g.打つ('D' + (i + 1), f));
  式ら.forEach((f, i) => T('止め3 打ってから＝作り直し ... ' + f + ' = ' + g.字('D' + (i + 1)), h.字('D' + (i + 1)) === g.字('D' + (i + 1)), h.字('D' + (i + 1))));
  T('止め3 空を 挟んだ MATCH の 近い 方 → 前と 同じ #REF!（実Excel 未測定）', g.字('D1') === '#REF!', g.字('D1'));
  T('止め3 空を 挟んでも 完全一致は 読む → 10', g.字('D6') === '10', g.字('D6'));
}
{
  /* 止め4＝重さ：関係ない 長い 列・打って 消した 遠い マスで 1打ちが 重く ならない（書いた 範囲の 時と 比べる） */
  function 一打ち(範囲) {
    const h = H.表();
    for (let i = 1; i <= 100; i++) { h.打つ('A' + i, String(i)); h.打つ('B' + i, String(i * 2)); }
    h.まとめて打つ始め();
    for (let i = 1; i <= 20000; i++) h.打つ('Z' + i, '1');
    for (let i = 1; i <= 1000; i++) { h.打つ('C' + i, String(i % 100 + 1)); h.打つ('D' + i, '=VLOOKUP(C' + i + ',' + 範囲 + ',2,0)'); }
    h.まとめて打つ終わり();
    h.打つ('Q200000', '1'); h.打つ('Q200000', '');
    const t = Date.now(); h.打つ('C5', '7'); const ms = Date.now() - t;
    return { ms, 字: h.字('D5') };
  }
  const 帯 = 一打ち('A:B'), 書 = 一打ち('$A$1:$B$100');
  console.log('    1打ち 帯 ' + 帯.ms + 'ms ／ 書いた 範囲 ' + 書.ms + 'ms');
  T('止め4 答えは 同じ（14）', 帯.字 === '14' && 書.字 === '14', 帯.字 + '/' + 書.字);
  T('止め4 1打ちが 書いた 範囲の 5倍＋200ms 以内', 帯.ms <= 書.ms * 5 + 200, 帯.ms + 'ms / ' + 書.ms + 'ms');
}

console.log('');
console.log('[retsu-x] pass ' + pass + ' / fail ' + fail);
process.exit(fail ? 1 : 0);
