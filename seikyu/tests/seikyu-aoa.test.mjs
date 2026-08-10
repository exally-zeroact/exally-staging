/* seikyu-aoa.test.mjs — ★Excelで渡した相手の画面で、読めて・足せるか★
 *
 * ここで止めたい事故:
 *   ① ★列幅を付けずに出して ######## になる★（前科あり。渡した相手の画面で初めて分かる）
 *   ② ★金額を "1,234円" のような文字で出す★（相手が足し算できない＝Excelで渡す意味が消える）
 *   ③ 税率ごとの区分が Excel に出ない（紙と食い違う）
 *   ④ 明細0行で落ちる／空の表を出す
 *
 * ★実際に .xlsx を組んで読み戻して測る★（作った物を見るだけにしない）。
 *
 * 使い方: node seikyu/tests/seikyu-aoa.test.mjs
 *         node seikyu/tests/seikyu-aoa.test.mjs --self-test
 */
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..', '..');
const require_ = createRequire(path.join(ROOT, 'package.json'));
const AOA = require_(path.join(ROOT, 'seikyu/lib/seikyu-aoa.js'));
const TAX = require_(path.join(ROOT, 'seikyu/lib/seikyu-tax.js'));
const SR = require_(path.join(ROOT, 'kyuyo/lib/shouhizei-ritsu.js'));
const XLSX = require_(path.join(ROOT, 'lib/xlsx.full.min.js'));

const STD = Math.round(SR.hyojun * 10000) / 100;
const RED = Math.round(SR.keigen * 10000) / 100;

let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ✓ ' + n); } catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + (e && e.message)); } };
const eq = (a, b, m) => { if (a !== b) throw new Error((m ? m + ': ' : '') + 'expected ' + JSON.stringify(b) + ' got ' + JSON.stringify(a)); };
const ok = (v, m) => { if (!v) throw new Error(m || 'false'); };

function sample() {
  const lines = [
    { name: '運転代行 9月分', qty: 42, unit: '件', price: 3200, rate: STD },
    { name: 'お弁当代', amount: 1000, rate: RED },
    { name: '立替金（対象外）', amount: 500, rate: 0 },
  ];
  const tax = TAX.compute({ lines, taxMode: 'exclusive', rounding: 'floor' });
  return {
    inv: { doc_type: 'invoice', no: '202609-001', issue_ymd: '2026-09-30', due_ymd: '2026-10-31', data: { subject: '9月分', memo: '備考です' } },
    tax,
    partner: { name: '藤原建設株式会社', keisho: '御中', addr: '愛媛県今治市1-2-3' },
    org: { yago: '株式会社ゼロアクト', addr: '今治市4-5-6', tel: '0898-00-0000', invoiceNo: 'T1234567890123', bank: '伊予銀行' },
  };
}

/* 実際に .xlsx を組んで、読み戻す（＝相手のExcelが受け取る形で測る） */
function roundTrip(sheet) {
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(sheet.aoa);
  if (sheet.cols) ws['!cols'] = sheet.cols;
  (sheet.numFmt || []).forEach((f) => {
    const ref = XLSX.utils.encode_cell({ r: f.r, c: f.c });
    if (ws[ref] && typeof ws[ref].v === 'number') ws[ref].z = f.z;
  });
  XLSX.utils.book_append_sheet(wb, ws, sheet.name || 'Sheet1');
  const buf = XLSX.write(wb, { bookType: 'xlsx', type: 'buffer', cellStyles: true });
  const back = XLSX.read(buf, { type: 'buffer', cellStyles: true });
  return back.Sheets[back.SheetNames[0]];
}

/* ★純関数（self-test で作り物を通せる）★ やってはいけない方の表 */
export function sheetBad(kind, sheet) {
  const s = JSON.parse(JSON.stringify(sheet));
  if (kind === 'noCols') { delete s.cols; return s; }
  if (kind === 'strMoney') {
    s.aoa = s.aoa.map((row) => row.map((v) => (typeof v === 'number' ? v.toLocaleString('ja-JP') + '円' : v)));
    s.numFmt = [];
    return s;
  }
  return s;
}

/* ── self-test ────────────────────────────────────────────────── */
if (process.argv.includes('--self-test')) {
  console.log('\n[seikyu-aoa --self-test] わざと壊して赤になるか');
  let sp = 0, sf = 0;
  const S = (n, fn) => { try { fn(); sp++; console.log('  ✓ ' + n); } catch (e) { sf++; console.log('  ✗ ' + n + ' — ' + (e && e.message)); } };

  S('① 列幅を外した表は「######## にならない」検査に落ちる', () => {
    const bad = roundTrip(sheetBad('noCols', AOA.build(sample())));
    ok(!bad['!cols'] || !bad['!cols'].length, '作り物なのに列幅が残っている＝この検査が空振り');
    const good = roundTrip(AOA.build(sample()));
    ok(good['!cols'] && good['!cols'].length, '本物に列幅が無い');
  });

  S('② 金額を文字にした表は「数のまま」検査に落ちる', () => {
    const bad = roundTrip(sheetBad('strMoney', AOA.build(sample())));
    const badNums = Object.keys(bad).filter((k) => k[0] !== '!' && bad[k].t === 'n').length;
    eq(badNums, 0, '作り物なのに数が残っている＝この検査が空振り');
    const good = roundTrip(AOA.build(sample()));
    const goodNums = Object.keys(good).filter((k) => k[0] !== '!' && good[k].t === 'n').length;
    ok(goodNums > 5, '本物に数のセルが少なすぎる: ' + goodNums);
  });

  console.log('\n[self-test] ' + sp + ' passed, ' + sf + ' failed');
  if (sf) process.exit(1);
}

/* ── 本体 ──────────────────────────────────────────────────────── */
console.log('\n[請求書 Excel の中身]');
const S1 = sample();
const SH = AOA.build(S1);
const WS = roundTrip(SH);
const cellAt = (r, c) => WS[XLSX.utils.encode_cell({ r, c })];
const asRows = () => XLSX.utils.sheet_to_json(WS, { header: 1, raw: true });

T('★列幅が全部の列に付いている（相手の画面で ######## にしない）', () => {
  ok(WS['!cols'] && WS['!cols'].length === 7, '列幅の数=' + ((WS['!cols'] || []).length));
  WS['!cols'].forEach((c, i) => ok(c && Number(c.wch) > 0, i + '列目に幅が無い'));
  // 品名は他より広い（長い文が入る列）
  ok(WS['!cols'][1].wch >= 30, '品名の列が狭い: ' + WS['!cols'][1].wch);
});

T('★金額は数のまま出る（相手が足し算できる）', () => {
  const rows = asRows();
  const head = rows.findIndex((r) => r[0] === '#' && r[1] === '品名・内容');
  ok(head > 0, '明細の見出し行が無い');
  for (let i = 0; i < S1.tax.lines.length; i++) {
    const cell = cellAt(head + 1 + i, 5);
    ok(cell, (i + 1) + '行目の金額が無い');
    eq(cell.t, 'n', (i + 1) + '行目の金額が数でない');
    eq(cell.v, S1.tax.lines[i].amount, (i + 1) + '行目の金額');
  }
});

T('★金額のセルに桁区切りの書式が付いている', () => {
  const rows = asRows();
  const head = rows.findIndex((r) => r[0] === '#' && r[1] === '品名・内容');
  const cell = cellAt(head + 1, 5);
  ok(cell.z && /#,##0/.test(String(cell.z)), '書式=' + cell.z);
});

T('★小計・消費税・合計 が数で出て、税抜+消費税=税込 が一致する', () => {
  const rows = asRows();
  const find = (label) => rows.find((r) => r[4] === label);
  const sub = find('小計'), tx = find('消費税'), gr = find('合計');
  ok(sub && tx && gr, '合計欄が出ていない');
  eq(typeof sub[5], 'number', '小計が数でない');
  eq(sub[5] + tx[5], gr[5], '税抜+消費税≠税込');
  eq(gr[5], S1.tax.grandTotal, '合計が計算と違う');
});

T('★税率ごとの区分が Excel にも出る（紙と食い違わせない）', () => {
  const rows = asRows();
  const head = rows.findIndex((r) => r[0] === '区分' && r[1] === '対象額');
  ok(head > 0, '区分の見出しが無い');
  S1.tax.byRate.forEach((b, i) => {
    const r = rows[head + 1 + i];
    ok(r, i + 1 + 'つ目の区分が無い');
    eq(r[0], b.pct + '% 対象');
    eq(r[1], b.base, '対象額');
    eq(r[2], b.tax, '消費税');
  });
  ok(rows.some((r) => r[0] === '消費税の対象外'), '対象外の行が無い');
});

T('★取れなかったを空欄にしない（相手・自社・番号）', () => {
  const s = AOA.build({ inv: { doc_type: 'invoice', no: '', issue_ymd: '', data: {} }, tax: TAX.compute({ lines: [], taxMode: 'exclusive', rounding: 'floor' }), partner: {}, org: {} });
  const flat = JSON.stringify(s.aoa);
  ok(/（取引先が未選択）/.test(flat), '宛先が空欄');
  ok(/（自社情報が未入力）/.test(flat), '自社が空欄');
  ok(/（未採番）/.test(flat), '番号が空欄');
  ok(/明細がまだ1行もありません/.test(flat), '空の表が出ている');
  ok(/区分はまだありません/.test(flat), '空の区分が出ている');
});

T('★undefined / NaN を1つも出さない', () => {
  const flat = JSON.stringify(SH.aoa);
  ok(!/null,null,null,null,null,null,null/.test(flat) || true, '（空行そのものは可）');
  SH.aoa.forEach((row, i) => row.forEach((v, j) => {
    ok(v !== undefined, i + '行' + j + '列が undefined');
    ok(!(typeof v === 'number' && !Number.isFinite(v)), i + '行' + j + '列が NaN');
  }));
});

T('見積書は呼び方が変わる（シート名も）', () => {
  const s = AOA.build(Object.assign({}, S1, { inv: Object.assign({}, S1.inv, { doc_type: 'quote' }) }));
  eq(s.name, '見積書');
  ok(JSON.stringify(s.aoa).includes('見積番号'), '番号の呼び方が請求書のまま');
  ok(JSON.stringify(s.aoa).includes('お見積金額'), '金額の呼び方が請求書のまま');
});

T('★網羅：税率の組み合わせ×内外×丸め を全部書き出して、合計が Excel の中でも一致', () => {
  let n = 0;
  const sets = [
    [{ name: 'a', amount: 105, rate: STD }],
    [{ name: 'a', amount: 105, rate: RED }],
    [{ name: 'a', amount: 105, rate: 0 }],
    [{ name: 'a', amount: 105, rate: STD }, { name: 'b', amount: 1000, rate: RED }, { name: 'c', amount: 500, rate: 0 }],
    [{ name: 'a', amount: -1100, rate: STD }],
  ];
  for (const lines of sets) for (const mode of ['exclusive', 'inclusive']) for (const rd of ['floor', 'ceil', 'round']) {
    const t = TAX.compute({ lines, taxMode: mode, rounding: rd });
    const s = AOA.build({ inv: { doc_type: 'invoice', no: 'X', issue_ymd: '2026-09-30', data: {} }, tax: t, partner: { name: 'A' }, org: {} });
    const ws = roundTrip(s);
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true });
    const gr = rows.find((r) => r[4] === '合計');
    if (!gr) throw new Error('合計が無い');
    if (gr[5] !== t.grandTotal) throw new Error('合計が違う: ' + gr[5] + ' / ' + t.grandTotal);
    if (!ws['!cols'] || ws['!cols'].length !== 7) throw new Error('列幅が落ちた');
    n++;
  }
  if (n < 25) throw new Error('組合せが少なすぎる（検査が空振り）: ' + n);
  console.log('     実測: ' + n + '通りを書き出して読み戻し、矛盾0件');
});

console.log('\n' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
