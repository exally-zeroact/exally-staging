/* text-format.test.mjs — ★TEXT() の書式コードを Excel と同じに読めているか★
 *
 * なぜ必要か（2026-08-10・司さんの実物で実測）:
 *   実物（式15,126本）のうち ★730本★ が =TEXT(A5,"aaa")＝曜日 だった。
 *   この書式には y/m/d/h/s が1文字も入っていないので「数の書式」として扱われ、
 *   どの形にも当てはまらず ★シリアル値 46023 がそのまま画面に出ていた★。
 *     売上表 B5  =TEXT(A5,"aaa")   Excel="木"  ／ うち="46023"
 *   本番と同じ経路で数え直した結果:
 *     直す前  Excelと同じ答え 2151 ／ 答えが違う 730
 *     直した後 Excelと同じ答え 2881 ／ 答えが違う   0   （#ERROR は 12245 のまま増減なし）
 *
 * ★ここで固定する事★
 *   ① 実Excelの真値と1つずつ一致する（想像で書かない。全部 COM で聞いた実測値）
 *   ② 1年ぶん365日の曜日が、独立に出した曜日と全部一致する
 *   ③ ★計算側(TEXT)と表示側(BookOpen.withWeekday+SSF)が同じ字を出す★
 *      曜日を出す道が2本あるので、片方だけ直すと静かにズレる
 *   ④ ★読めない書式でもシリアル値を出さない★（46023 が答えとして出てはいけない）
 *   ⑤ ハーネスに曜日のケースが残っている（真値ごと消されない）
 *
 * 使い方: node tests/text-format.test.mjs
 *         node tests/text-format.test.mjs --self-test   … わざと壊して赤になるか
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createRequire } from 'node:module';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require = createRequire(pathToFileURL(path.join(ROOT, 'package.json')));
const HF = require(path.join(ROOT, 'hyperformula.full.min.js'));
const F = require(path.join(ROOT, 'exally-formula.js'));
const XLSX = require(path.join(ROOT, 'lib/xlsx.full.min.js'));

/* 表示側は book.html の fmtForDisplay が使っている物と同じ関数を、そのファイルから読む
   （書き写すと「直っていない物が直って見える」ため） */
const box = { XLSX };
new Function('self', 'window', fs.readFileSync(path.join(ROOT, 'js/book-open.js'), 'utf8')
  + '\n;self.__BookOpen = self.BookOpen;')(box, box);
const BookOpen = box.__BookOpen;

/* ── 本番と同じ計算の入口（book.html の initFormulaEngine と同じ作り） ── */
F.registerExallyFunctions(HF);
const hf = HF.HyperFormula.buildEmpty({ licenseKey: 'gpl-v3', useArrayArithmetic: true, smartRounding: false });
hf.addSheet('S');
const SID = hf.getSheetId('S');
function TEXT(v, fmt) {
  hf.setSheetContent(SID, [[v, '=TEXT(A1,"' + String(fmt).replace(/"/g, '""') + '")']]);
  const r = hf.getCellValue({ sheet: SID, row: 0, col: 1 });
  return r === null || r === undefined ? '' : String(r);
}

/* ★実Excel 365 / 16.0.20228.20158（日本語UI 1041）に COM で1つずつ聞いた実測値★
   46053 = 2026-01-31(土) / 46023 = 2026-01-01(木) / 46053.5 = 同日12:00 / 0.75 = 18:00 */
const EXCEL = [
  [46053, 'aaa', '土'], [46053, 'aaaa', '土曜日'],
  [46023, 'aaa', '木'], [46023, 'aaaa', '木曜日'],
  [46053, '[$-411]aaa', '土'], [46053, '[$-411]aaaa', '土曜日'],
  [46053, 'm/d(aaa)', '1/31(土)'], [46053, 'yyyy/m/d(aaa)', '2026/1/31(土)'],
  [46053, 'aaa;@', '土'], [46053, 'm/d aaa', '1/31 土'],
  [46053, 'ddd', 'Sat'], [46053, 'dddd', 'Saturday'],
  [46023, 'ddd', 'Thu'], [46023, 'dddd', 'Thursday'],
  [46053, '[$-409]ddd', 'Sat'], [46053, '[$-409]dddd', 'Saturday'],
  [46053, 'mmm', 'Jan'], [46053, 'mmmm', 'January'], [46053, 'mmmmm', 'J'],
  [1234.5, 'mmm', 'May'],
  [46053, 'yyyy', '2026'], [46053, 'yy', '26'], [46053, 'dd', '31'], [46053, 'd', '31'],
  [46053, 'm/d;@', '1/31'], [1234.5, 'm/d/yyyy', '5/18/1903'],
  /* ★m は月か分か（時の直後なら分・秒の直前なら分） */
  [46053.5, 'hh:mm', '12:00'], [46053.5, 'h:mm:ss', '12:00:00'],
  [46053.5, 'm/d h:mm', '1/31 12:00'], [46053.5, 'yyyy/mm/dd hh:mm:ss', '2026/01/31 12:00:00'],
  [46053.5, 'mm:ss', '00:00'], [46053.5, 'm', '1'], [46053.5, 'mm', '01'],
  [46053.5, 'h', '12'], [46053.5, 's', '0'],
  [0.75, 'h:mm AM/PM', '6:00 PM'], [0.75, 'hh:mm:ss AM/PM', '06:00:00 PM'],
  /* ★元から合っていた物＝壊していないかの確認（goldenにもある形） */
  [1234.5, '#,##0', '1,235'], [1234.567, '#,##0.00', '1,234.57'], [0.1235, '0.0%', '12.4%'],
  [7, '0000', '0007'], [2.675, '0.00', '2.68'], [-1234, '#,##0;(#,##0)', '(1,234)'],
  [46234, 'yyyy/mm/dd', '2026/07/31'], [46234, 'yy/m/d', '26/7/31'],
];

/* 独立に出した曜日（Excelの1900日付系。うちの実装は使わない） */
const WD = ['日', '月', '火', '水', '木', '金', '土'];
const weekdayOf = (serial) => WD[new Date(Date.UTC(1899, 11, 30) + Math.floor(serial) * 86400000).getUTCDay()];

let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ✓ ' + n); } catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + (e && e.message)); } };

/* ═══ 自己テスト：わざと壊して赤になるか ═══ */
if (process.argv.includes('--self-test')) {
  console.log('\n[text-format --self-test] ★わざと壊して赤になるか');
  const src = fs.readFileSync(path.join(ROOT, 'exally-formula.js'), 'utf8');
  const breaks = [
    ['① 曜日を日付の書式と見なす所を外す（aaa が数の書式に落ちる）',
      /\|\| \/a\{3,4\}\/i\.test\(f\)/],
    ['② 曜日の字(aaa)を書式の読み取り表から外す',
      /case 'aaa':\s+out \+= _WD_JA\[W\];/],
    ['③ 言語の指定 [$-411] を外す所を消す',
      /function _fmtStripLocale/],
    ['④ 区画(;)の1つ目だけを使う所を消す',
      /_fmtDate\(num, _fmtSections\(f\)\[0\]\)/],
    ['⑤ m が「分」か「月」かを見分ける所を消す',
      /function _fmtIsMinute/],
  ];
  for (const [name, re] of breaks) {
    T(name + ' → 見つからなくなったら赤にできる', () => {
      if (!re.test(src)) throw new Error('この目印が exally-formula.js に無い＝壊れても気付けない: ' + re);
    });
  }
  T('★壊した時に本当に答えが変わる（空振りしていない）', () => {
    //  曜日を見ない実装＝旧 _fmtIsDate に戻すと "aaa" は数の書式になり、どの形にも当たらず
    //  最後に String(n) が返る＝シリアル値がそのまま出る。その状態を再現して確かめる。
    const old = (fmt) => /[ymdhs]/i.test(String(fmt).replace(/"[^"]*"/g, '')) && !/[#0]/.test(fmt);
    if (old('aaa')) throw new Error('旧判定でも aaa を日付と見なしてしまう＝この比較が空振り');
    if (TEXT(46023, 'aaa') === '46023') throw new Error('直っていない');
  });
  console.log('\n' + pass + ' passed, ' + fail + ' failed');
  hf.destroy();
  process.exit(fail ? 1 : 0);
}

console.log('\n[text-format] TEXT() の書式コードが Excel と同じに読めているか');

T('★実Excelの真値と1つずつ一致する（' + EXCEL.length + '件）', () => {
  const ng = [];
  for (const [v, fmt, want] of EXCEL) {
    const got = TEXT(v, fmt);
    if (got !== want) ng.push(`TEXT(${v},"${fmt}") Excel=${JSON.stringify(want)} うち=${JSON.stringify(got)}`);
  }
  if (ng.length) throw new Error('\n     ' + ng.join('\n     '));
});

T('★1年ぶん365日の曜日が全部合う（実物と同じ =TEXT(日付,"aaa") の形）', () => {
  const ng = [];
  for (let s = 46023; s < 46023 + 365; s++) {
    const got = TEXT(s, 'aaa');
    if (got !== weekdayOf(s)) ng.push(s + ': ' + got + ' ≠ ' + weekdayOf(s));
  }
  if (ng.length) throw new Error(ng.length + '日ぶん違う（例 ' + ng.slice(0, 3).join(' / ') + '）');
});

T('★計算側と表示側が同じ曜日を出す（道が2本ある＝片方だけ直すと静かにズレる）', () => {
  if (typeof BookOpen.withWeekday !== 'function') throw new Error('BookOpen.withWeekday が無い＝表示側を確かめられない');
  const ng = [];
  for (let s = 46023; s < 46023 + 365; s++) {
    //  表示側：ファイルのセル書式 m/d([$-411]aaa) を、本番の描画と同じ道具で文字にする
    const shown = XLSX.SSF.format(BookOpen.withWeekday('m/d\\([$-411]aaa\\)', s), s);
    const calc = TEXT(s, 'm/d(aaa)');                 // 計算側：同じ日付・同じ書式
    if (shown !== calc) ng.push(s + ': 表示=' + shown + ' 計算=' + calc);
  }
  if (ng.length) throw new Error(ng.length + '日ぶん食い違う（例 ' + ng.slice(0, 3).join(' / ') + '）');
});

T('★読めない書式でもシリアル値を答えにしない（46023 が出てはいけない）', () => {
  //  和暦(ge.m.d)と経過時間([h]:mm)は まだ読めない＝台帳(known-diffs)に載せてある。
  //  読めない時に「数をそのまま返す」と、日付が 46023 という別物になって静かに紙に載る。
  const ng = [];
  for (const fmt of ['aaa', 'aaaa', 'ddd', 'dddd', 'mmm', 'mmmm', 'mmmmm',
    'm/d(aaa)', '[$-411]aaa', 'ge.m.d', 'ggge', '[h]:mm', 'hh:mm']) {
    const got = TEXT(46023, fmt);
    if (got === '46023' || got === String(46023)) ng.push(fmt);
  }
  if (ng.length) throw new Error('シリアル値がそのまま出た書式: ' + ng.join(' / '));
});

T('★ハーネスに曜日の真値ケースが残っている（真値ごと消されない）', () => {
  const g = JSON.parse(fs.readFileSync(path.join(ROOT, 'tests/xlsx-harness/golden/excel-365-16.0.20228.json'), 'utf8'));
  const need = ['TEXT_weekday_ja', 'TEXT_weekday_ja_long', 'TEXT_weekday_ja_locale',
    'TEXT_weekday_in_date', 'TEXT_weekday_section', 'TEXT_minute_vs_month'];
  const missing = need.filter((k) => !g.cases[k]);
  if (missing.length) throw new Error('goldenから消えている: ' + missing.join(', '));
  if (g.cases.TEXT_weekday_ja.v !== '土') throw new Error('goldenの真値が書き換わっている: ' + g.cases.TEXT_weekday_ja.v);
});

T('★まだ読めない書式は台帳に載っている（黙って放置しない）', () => {
  const k = JSON.parse(fs.readFileSync(path.join(ROOT, 'tests/xlsx-harness/known-diffs.json'), 'utf8'));
  const ids = (k.diffs || []).map((d) => d.id);
  for (const id of ['TEXT_era_wareki', 'TEXT_elapsed_hours']) {
    if (ids.indexOf(id) < 0) throw new Error('known-diffs.json に ' + id + ' が無い');
  }
});

console.log('\n── 実測（司さんの実物と同じ形） ──');
[46023, 46024, 46025, 46026, 46027, 46028, 46029].forEach((s) => {
  console.log(`  =TEXT(${s},"aaa") → ${TEXT(s, 'aaa')}   （Excelの答え ${weekdayOf(s)}）`);
});
console.log('\n' + pass + ' passed, ' + fail + ' failed');
hf.destroy();
process.exit(fail ? 1 : 0);
