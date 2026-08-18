/* smart-rounding.test.mjs — ★計算の結果を「きれいに」丸めさせない★
 *
 * なぜ必要か（2026-08-09〜10・実測）:
 *   HyperFormula の既定 smartRounding:true は、答えを14桁で丸めて
 *   「きれいな数字」に直す。★Excel は丸めない。★
 *   司さんの実物（式15,126本）で測ったら、切るだけで ★41件が Excel と一致し、
 *   新しい食い違いは 0件★ だった。しかも直った中に:
 *       売上表 D37  =D36/1.1        188000  → 187999.99999999997（税抜き）
 *       売上表 AN38 =AN36*0.1/1.1   200     → 199.99999999999997（★消費税★）
 *       距離   D1   =C36+G36+…      26718.5 → 26718.499999999996（合計距離）
 *   ★切り捨てが1つ入るだけで 1円 ズレる★＝記憶の決まり
 *   「グリッドの数式では作らない（黙って1円ズレる）」の実例だった。
 *
 *   ★書き出す時のキャッシュ値もうちの計算結果なので、切らないと★渡した紙まで違う★。
 *
 * ★次の人が「桁が出て見にくい」で戻せないように、ここで赤にする。★
 *   見にくさは ★表示側(forDisplay)★ で直す。計算は触らない。
 *
 * 使い方: node tests/smart-rounding.test.mjs
 *         node tests/smart-rounding.test.mjs --self-test
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const require = createRequire(import.meta.url);
const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const HF = require(path.join(ROOT, 'hyperformula.full.min.js'));

/* ★Excel が実際に出した答え（司さんの実物から読み取った実測値）★ */
export const EXCEL = [
  { name: '売上表 D37  =D36/1.1（税抜き）', a: 206800, f: '=A1/1.1', want: 187999.99999999997, floorSmart: 188000, floorReal: 187999 },
  { name: '売上表 AN37 =AN36/1.1（税抜き）', a: 2200, f: '=A1/1.1', want: 1999.9999999999998, floorSmart: 2000, floorReal: 1999 },
  { name: '売上表 AN38 =AN36*0.1/1.1（消費税）', a: 2200, f: '=A1*0.1/1.1', want: 199.99999999999997, floorSmart: 200, floorReal: 199 },
];

/** 表示だけ整える（book.html の forDisplay と同じ物。★計算は触らない★）
 *  ★文字も受ける★: recalcSheet は答えを _hfGetDisplay の戻り＝文字で置くので、
 *  数だけ見ていると 187999.99999999997 が素通りする（2026-08-10 実測）。 */
export function forDisplay(v) {
  var n = (typeof v === 'number') ? v
        : ((typeof v === 'string' && /^-?\d+\.\d+$/.test(v)) ? Number(v) : null);
  if (n === null || !isFinite(n)) return v;
  if (Number.isInteger(n)) return v;
  var s = n.toPrecision(15);
  if (s.indexOf('e') >= 0 || s.indexOf('E') >= 0) return v;
  return String(Number(s));
}

function calc(smart, a, formula) {
  const hf = HF.buildEmpty({ licenseKey: 'gpl-v3', useArrayArithmetic: true, smartRounding: smart });
  hf.addSheet('S');
  const id = hf.getSheetId('S');
  hf.setSheetContent(id, [[a, formula]]);
  const v = hf.getCellValue({ sheet: id, row: 0, col: 1 });
  hf.destroy();
  return v;
}

let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ✓ ' + n); } catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + (e && e.message)); } };

if (process.argv.includes('--self-test')) {
  console.log('\n[smart-rounding] ★わざと戻して赤になるか★');
  T('★既定(true)に戻すと Excel と答えが違う（＝赤になる材料）', () => {
    const bad = EXCEL.filter((c) => calc(true, c.a, c.f) === c.want);
    if (bad.length) throw new Error('true でも一致してしまった: ' + bad.map((x) => x.name).join(','));
  });
  T('★false なら Excel と同じ', () => {
    EXCEL.forEach((c) => { if (calc(false, c.a, c.f) !== c.want) throw new Error(c.name); });
  });
  T('★1円ズレる事を数字で示せる（切り捨て）', () => {
    EXCEL.forEach((c) => {
      if (Math.floor(calc(true, c.a, c.f)) !== c.floorSmart) throw new Error(c.name + ': 丸めありの切り捨て');
      if (Math.floor(calc(false, c.a, c.f)) !== c.floorReal) throw new Error(c.name + ': 丸めなしの切り捨て');
    });
  });
  T('表示だけ整える方は 見やすい', () => {
    if (String(forDisplay(0.1 + 0.2)) !== '0.3') throw new Error('0.1+0.2');
    if (String(forDisplay(187999.99999999997)) !== '188000') throw new Error('187999…');
  });
  console.log('\n' + pass + ' passed, ' + fail + ' failed');
  process.exit(fail ? 1 : 0);
}

console.log('\n[smart-rounding] 計算の結果を「きれいに」丸めていないか');
const html = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf8');

T('★book.html が smartRounding:false で計算する側を作っている', () => {
  const m = html.match(/HyperFormula\.buildEmpty\(\{([^}]*)\}/);
  if (!m) throw new Error('buildEmpty が見つからない');
  if (!/smartRounding\s*:\s*false/.test(m[1])) {
    throw new Error('★smartRounding:false が指定されていません★\n'
      + '   既定(true)は答えを14桁で丸めます。Excelは丸めません。\n'
      + '   ★切り捨てが1つ入るだけで 1円 ズレます（消費税の式で実測済み）★\n'
      + '   見にくさは forDisplay（表示側）で直してください。計算は触らないこと。\n'
      + '   今の中身: ' + m[1].trim());
  }
});
T('★Excel が出した答えと 1つずつ一致する（実測値）', () => {
  EXCEL.forEach((c) => {
    const got = calc(false, c.a, c.f);
    if (got !== c.want) throw new Error(c.name + ': ' + got + ' ≠ ' + c.want);
  });
});
T('★丸めありだと 切り捨てで 1円 変わる（戻した時に何が起きるか）', () => {
  EXCEL.forEach((c) => {
    if (Math.floor(calc(true, c.a, c.f)) === Math.floor(calc(false, c.a, c.f))) {
      throw new Error(c.name + ': 差が出ない＝この検査が空振りしている');
    }
  });
});
T('★見にくさは表示側で直している（計算は触っていない）', () => {
  if (String(forDisplay(0.1 + 0.2)) !== '0.3') throw new Error('0.1+0.2 が見やすくならない');
  if (String(forDisplay(26718.499999999996)) !== '26718.5') throw new Error('26718.5');
  if (forDisplay(0.1 + 0.2) === 0.1 + 0.2) throw new Error('数のまま返している（文字にして見せる）');
  if (!/function forDisplay/.test(html)) throw new Error('book.html に forDisplay が無い');
});
T('★答えが「文字」で来ても丸めて見せる（recalcSheet は文字を置く）', () => {
  //  ここを数だけにすると、実物で 41件が 187999.99999999997 のまま画面に出る（実測）
  if (forDisplay('187999.99999999997') !== '188000') throw new Error('文字の 187999… が素通り');
  if (forDisplay('26718.499999999996') !== '26718.5') throw new Error('文字の 26718.5 が素通り');
  if (forDisplay('1/21') !== '1/21') throw new Error('日付の文字を触ってはいけない');
  if (forDisplay('7,000') !== '7,000') throw new Error('桁区切りの文字を触ってはいけない');
  if (!/typeof v === 'string'/.test(html.slice(html.indexOf('function forDisplay'), html.indexOf('function forDisplay') + 500))) {
    throw new Error('★book.html の forDisplay が数しか見ていない★（文字で来る答えが素通りする）');
  }
});
T('★描く所は fmtForDisplay を通している（書式つきのセルも救われる）', () => {
  if (!/function fmtForDisplay/.test(html)) throw new Error('book.html に fmtForDisplay が無い');
  if (!/var display = fmtForDisplay\(raw, cell\.numFmt\);/.test(html)) {
    throw new Error('★描く所で fmtForDisplay を通していない★\n'
      + '   applyNumFmt は Excel の書式をほとんど知らない。司さんの実物では\n'
      + '   ★書式つき 6,568セルが全部 素通り★していた（"#,##0_ " も "General" も知らない）。\n'
      + '   素通りすると 187999.99999999997 がそのまま画面に出る。');
  }
  //  順番が命: ①ファイルの書式器 → ②applyNumFmt → ③15桁に丸める
  const body = html.slice(html.indexOf('function fmtForDisplay'));
  const i1 = body.indexOf('XLSX.SSF.format'), i2 = body.indexOf('applyNumFmt('), i3 = body.indexOf('forDisplay(raw)');
  if (!(i1 > 0 && i2 > i1 && i3 > i2)) throw new Error('順番が違う（書式器→applyNumFmt→15桁 の順で落とすこと）');
  if (body.slice(0, 900).indexOf('/^-?\\d+(\\.\\d+)?$/') < 0) {
    throw new Error('★裸の数字だけに使う歯止めが無い★（"1/21" を数に戻すと 1/1 になる）');
  }
});
T('★書式器を呼ぶための withWeekday が公開されている', () => {
  const bo = fs.readFileSync(path.join(ROOT, 'js', 'book-open.js'), 'utf8');
  if (!/withWeekday:\s*withWeekday/.test(bo)) {
    throw new Error('BookOpen.withWeekday が公開されていない＝fmtForDisplay が例外で落ちる');
  }
});

console.log('\n── 実測 ──');
EXCEL.forEach((c) => {
  console.log(`  ${c.name}`);
  console.log(`     丸めあり(既定)=${calc(true, c.a, c.f)}  → 切り捨て ${c.floorSmart}`);
  console.log(`     ★丸めなし   =${calc(false, c.a, c.f)}  → 切り捨て ${c.floorReal}（Excelと同じ）★`);
});
console.log('\n' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
