/* moji-no-kotae.test.mjs — ★文字を 返す 式が 数に 化けない★（2026-09-11）
 *
 *  ★★何が 起きて いたか★★
 *    `=TEXT(A1,"0.00")` の 答えは 実Excel では ★文字の "3.00"★
 *      実測 … LEN=4 ／ ISTEXT=TRUE（docs/measured/golden-moji-no-kotae-2026-09-11.tsv）
 *    ★エンジンも 文字で 返して いました★（typeof が 'string'）
 *    ★でも `_hfGetDisplay` が `String(val)` に した 時点で 型が 消える★
 *    ⇒ 画面が 数として 読み直し ★"3"★ と 出して いた
 *
 *  ★★道は 3本 在りました★★（2026-09-11 … また 片方だけ 直した）
 *    ①`recalcSheet`   … ファイルを 開いた 時   ⇒ 直した
 *    ②`_溢れを写す`    … 溢れた 先             ⇒ 直した
 *    ③★`setCellFormula` の 呼び手★＝お客さんが 打った 時 ⇒★忘れて いた★
 *    ★本物の ブラウザで 押して 初めて 見つけました★（台の 数では 出ない）
 *
 *  ★この 台は エンジンを 建てません★＝字を 読んで 数えるだけ。
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';
import { 注記を外す } from '../scripts/lib/chuki.mjs';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const 本 = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf-8');
const 式元 = fs.readFileSync(path.join(ROOT, 'exally-formula.js'), 'utf-8');
const 紙 = path.join(ROOT, 'docs/measured/golden-moji-no-kotae-2026-09-11.tsv');

let pass = 0, fail = 0;
const T = (n, f) => {
  try { f(); pass++; console.log('  ✓ ' + n); }
  catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + e.message); }
};

console.log('');
console.log('[moji-no-kotae] ★文字を 返す 式が 数に 化けない★');

const 実 = [];
for (const l of fs.readFileSync(紙, 'utf-8').split(/\r?\n/)) {
  if (!l || l.startsWith('#')) continue;
  const c = l.split('\t');
  if (c.length < 5) continue;
  実.push({ 式: c[0], 字: c[1], 型: c[2] });
}

T('★紙が 読めて いる（空振りして いない）★', () => {
  if (実.length < 10) throw new Error('★実Excel の 紙が ' + 実.length + '行★');
  const 字 = 実.filter((x) => x.型 === 'String').length;
  if (字 < 8) throw new Error('★文字を 返す 式が ' + 字 + '本しか 無い★＝紙が 薄い');
});

T('★★①実Excel は 文字の まま 出す（紙の 中身）★★', () => {
  const 見る = { '=TEXT(A1,"0.00")': '3.00', '=LEFT("3.00",4)': '3.00', '=TRIM("  4.50  ")': '4.50' };
  for (const [式, 期待] of Object.entries(見る)) {
    const r = 実.find((x) => x.式 === 式);
    if (!r) throw new Error('★紙に 無い … ' + 式 + '★');
    if (r.字 !== 期待) throw new Error('★' + 式 + ' が "' + r.字 + '"★（実Excel は "' + 期待 + '"）');
    if (r.型 !== 'String') throw new Error('★' + 式 + ' の 型が ' + r.型 + '★');
  }
});

T('★★②エンジンの 答えの 型を 聞ける★★', () => {
  const 動く = 注記を外す(式元);
  if (動く.indexOf('function _hf答えの型') < 0) {
    throw new Error('★型を 聞く 道が 無い★＝`String(val)` で 消えた まま に なります');
  }
  if (動く.indexOf("return 'string'") < 0) throw new Error("★'string' を 返して いない★");
});

T('★★③道が 全部 通って いる（★1本でも 抜けると 化ける★）★★', () => {
  /* ★注記の 外だけ 探します★＝説明文の 字で 緑に なるのを 防ぐ
     （2026-09-11 に 同じ日 2回 踏んだ … 注記の 中の convertFormula ／ 引用の 中の 言い切り） */
  const 動く = 注記を外す(本, { html: true });
  /* ★★探す 所を 近くで 切ります★★（2026-09-11）
     ★最初の 版は 赤に なりませんでした★＝
       `_打った答えは字か =` を repo 全体から 探して いたので、
       ★入れ物の 宣言（`var _打った答えは字か = false;`）だけで 緑★に なって いた。
     ⇒★立てて いる 所の 近くで 探します★（宣言では 通らない） */
  const 近くで = (印, 字, 長さ) => {
    const i = 動く.indexOf(印);
    return i >= 0 && 動く.slice(i, i + 長さ).indexOf(字) >= 0;
  };
  const 道 = [
    ['ファイルを 開いた 時（recalcSheet）',
      /* ★注記を 外した 後も 桁は ずれません★（空白で 埋める）＝説明が 長い分 だけ 離れます
         ⇒ 探す 幅は ★その 段が 収まる 長さ★に する（近すぎると 中身が 在るのに 赤） */
      () => 近くで('cell.d = (jsResult!==null)'   /* ★(...) を 足しました★（2026-09-18・㍸″″で 台に 回す 時に 3つ巴に なった） */, 'cell.d字 = true', 1600)],
    ['溢れた 先（_溢れを写す）',
      () => 近くで('t.d = _hfGetDisplay(sheetIdx, r + i, c + j, true)', 't.d字 = true', 400)],
    ['お客さんが 打った 時（setCellFormula の 呼び手）',
      () => 近くで('var hfResult = setCellFormula(activeSheet, r, c, v)', '_打った答えは字か =', 900)],
    ['打った 答えを マスに 入れる 所',
      () => 近くで('var next = Object.assign({}, existing,', 'next.d字 = true', 400)],
    ['★false を 置かず 消して いる★（空かを 見る 所が 狂う）',
      () => (動く.match(/delete (cell|t|next)\.d字/g) || []).length >= 3],
    ['★古い 溢れを 消す 時に d字 も 消す★',
      () => 近くで('delete t.d; delete t.d字', 'delete t[溢れの印]', 120)],
  ];
  const 抜け = [];
  for (const [名, 見る] of 道) if (!見る()) 抜け.push(名);
  if (抜け.length) {
    throw new Error('★' + 抜け.length + '本 抜けて いる★  ' + 抜け.join(' ／ ')
      + '  ⇒★1本でも 抜けると その 道だけ 数に 化けます★');
  }
  console.log('      … 開いた時／溢れた先／打った時／入れる所／false を 置かない／掃除 の 6つ');
});

T('★★④文字なら そのまま 出す（描く 所 2つ とも）★★', () => {
  const 動く = 注記を外す(本, { html: true });
  const 数 = (動く.match(/_答えは字か\(cell\) \? String\(raw\)/g) || []).length;
  if (数 < 2) {
    throw new Error('★' + 数 + 'か所しか 無い★＝ふつうの マスと 結合した マスの ★両方★ が 要ります');
  }
});

T('★★⑤文字は #### に しない（実Excel は はみ出す）★★', () => {
  const 動く = 注記を外す(本, { html: true });
  const 数 = (動く.match(/!_答えは字か\(cell\) && _数が入らないか/g) || []).length;
  if (数 < 2) throw new Error('★' + 数 + 'か所しか 無い★＝描く 所 2つ とも 要ります');
});

T('★★⑥JS層の 答えは「まだ 見て いない」と 書いて 在る★★', () => {
  /* ★当て推量で 決めない★＝JS層は 数も 字で 返すので 見分けが 付かない
     ⇒★今まで通り「字では ない」に して、断りを 残す★ */
  if (本.indexOf('まだ 見て いません') < 0 && 本.indexOf('まだ 測って いません') < 0) {
    throw new Error('★JS層を どう 扱ったかの 断りが 無い★＝未測定を 黙って 緑に しない');
  }
});

/* ★`async` に しない★＝`T` は 待たないので ★中の 失敗が 届かず 緑に なります★
   （2026-09-11 実際 そうなり、壊しても 赤に なりませんでした） */
T('★★⑦TEXT に 空の マスを 渡すと 0（★空の 字とは 別物★）★★', () => {
  /* ★★実Excel に 打たせて 測った★★（2026-09-11）
       =TEXT(空のマス,"aaa")   → ★土★     （0日目＝1900/1/0 は 土曜）
       =TEXT(空のマス,"0.00")  → ★0.00★
       =TEXT("","aaa")        → ★""★      （★空の 字は そのまま★）
       =TEXT("abc","0.00")    → "abc"／=TEXT(" ","0.00") → " "
       =TEXT(TRUE,"0.00")     → ★TRUE★    （うちは 1.00 に して いた）
       =TEXT("12","0.00")     → "12.00"
     ★見つけ方★＝司さんの 実物（計算の 板）を 1マスずつ 突き合わせた
       B299 `=TEXT(A299,"aaa")`（A299 は 空）… 実Excel「土」／うち「」
     ★★空の マスは ★記号★で 来ます★★＝`unwrap` が `''` に するので
       ★空の 字と 見分けが 付かなく なります★ ⇒★記号の うちに 見ます★
     ★まだ 合って いない 1本★（断り）
       =TEXT(0,"yyyy/m/d") … 実Excel「1900/1/0」／うち「1899/12/30」
       ＝Excel 独特の「0日目」＝★直して いません★ */
  /* ★★字を 探すのでは なく 本物を 動かします★★（2026-09-11）
     ★最初の 版は 字で 探して いて、抜いても 赤に なりませんでした★
     ⇒★本番と 同じ 台を 建てて 実際に 打ちます★ */
  const require2 = createRequire(path.join(ROOT, 'package.json'));
  const HFns = require2(path.join(ROOT, 'hyperformula.full.min.js'));
  const EF = require2(path.join(ROOT, 'exally-formula.js'));
  EF.registerExallyFunctions(HFns);
  const hf = HFns.HyperFormula.buildEmpty({ licenseKey: 'gpl-v3', useArrayArithmetic: true,
    smartRounding: false, maxRows: 1048576, maxColumns: 18278 });
  const SID = hf.getSheetId(hf.addSheet('S'));
  EF.initExallyFormula(hf);
  hf.setCellContents({ sheet: SID, row: 4, col: 0 }, [['abc']]);
  const 打つ = (f) => {
    hf.setCellContents({ sheet: SID, row: 2, col: 1 }, [[EF.convertFormula(f, 'S')]]);
    const v = hf.getCellValue({ sheet: SID, row: 2, col: 1 });
    return (v && v.type) ? ('#' + v.type) : String(v == null ? '' : v);
  };
  /* ★実Excel に 打たせて 測った 正解★（2026-09-11） */
  const 正 = [
    ['=TEXT(A1,"aaa")', '土'], ['=TEXT(A1,"aaaa")', '土曜日'], ['=TEXT(A1,"0.00")', '0.00'],
    ['=TEXT(0,"aaa")', '土'], ['=TEXT(1,"aaa")', '日'],
    ['=TEXT("","aaa")', ''], ['=TEXT("abc","0.00")', 'abc'], ['=TEXT(A5,"0.00")', 'abc'],
    ['=TEXT(" ","0.00")', ' '], ['=TEXT(TRUE,"0.00")', 'TRUE'], ['=TEXT("12","0.00")', '12.00'],
  ];
  const 違い = [];
  for (const [f, 期待] of 正) { const 出 = 打つ(f); if (出 !== 期待) 違い.push(f + ' 実Excel「' + 期待 + '」／うち「' + 出 + '」'); }
  if (違い.length) throw new Error('★' + 違い.length + '本 違う★  ' + 違い.join(' ／ '));
  console.log('      … 空のマス＝0／空の字＝そのまま／真偽＝字');
});

console.log('');
console.log('  ' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
