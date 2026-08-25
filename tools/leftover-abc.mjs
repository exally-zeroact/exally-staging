/* leftover-abc.mjs — ★取りこぼしを A/B/C/D に仕分ける★（★読むだけ・1バイトも書かない★）
 *
 *  ★指示役 2026-08-25★「189 ≠ 69。差を説明しろ。★Bを0にしろ★」
 *    A … ★Excel自身も エラーを保存している★（突き合わせの相手＝cell.t==='e'）
 *    B … ★Excelは答えを出しているのに うちが落とした★ ←★1本でも在れば 穴★
 *    C … ★もらい事故★（式の字に #REF! は在るが、IFERROR で包まれていて Excelは空を出す）
 *    D … ★Excelはエラーだが 式の字には出ない（伝染）★＝SUM の範囲に #REF! のセルが混ざる形
 *
 *  ★突き合わせの相手は「Excelが保存した計算結果」★＝自分の変換結果どうしで閉じない。
 *
 *  使い方: node tools/leftover-abc.mjs <ブックの場所>
 */
import fs from 'node:fs';
import { createRequire } from 'node:module';
const require_ = createRequire(import.meta.url);
const R = 'C:/Users/zeroa/exally-prod/';
const XLSX = require_(R + 'lib/xlsx.full.min.js');
const ZipSurgeon = require_(R + 'lib/zip-surgeon.js');
const TableRefs = require_(R + 'lib/table-refs.js');
const RG = require_(R + 'lib/ref-graph.js');

const 本 = process.argv[2];
const bytes = new Uint8Array(fs.readFileSync(本));
const wb = XLSX.read(bytes, { type: 'array', cellFormula: true, sheetStubs: false });
const r = await TableRefs.resolve(bytes, 'xlsb', wb, ZipSurgeon);
const fixes = (r && r.fixes) || {};

/* ★突き合わせの相手＝Excelが保存した計算結果（cell.t==='e'）★ 自分の変換結果で閉じない */
const Excelのエラー = new Set();
const エラーの中身 = {};
for (const name of wb.SheetNames) {
  const ws = wb.Sheets[name] || {};
  for (const k of Object.keys(ws)) {
    if (k[0] === '!') continue;
    if (ws[k].t === 'e') {
      const a = XLSX.utils.decode_cell(k);
      Excelのエラー.add(name + '|' + a.r + ',' + a.c);
      const w = ws[k].w || String(ws[k].v);
      エラーの中身[w] = (エラーの中身[w] || 0) + 1;
    }
  }
}
console.log('★Excel自身が保存したエラーのセル … ' + Excelのエラー.size + '個★  内訳: ' +
  Object.entries(エラーの中身).map(([k,v])=>k+'×'+v).join(' / '));

const sheets = wb.SheetNames.map((name) => {
  const ws = wb.Sheets[name] || {};
  const data = {};
  for (const k of Object.keys(ws)) {
    if (k[0] === '!') continue;
    const c = ws[k];
    const a = XLSX.utils.decode_cell(k);
    const key = name + '|' + a.r + ',' + a.c;
    data[a.r + ',' + a.c] = { v: c.v, f: fixes[key] !== undefined ? fixes[key] : (c.f ? '=' + c.f : undefined), t: c.t };
  }
  return { name, data };
});
const g = RG.作る(sheets);

/* ★数え方は2通り。混ぜない★ */
const のべ = g.取りこぼし.length;
const 式本数 = new Set(g.取りこぼし.map(x => x.from + ':' + x.fromCell)).size;
console.log('★取りこぼし … のべ ' + のべ + '本 ／ 式の本数 ' + 式本数 + '本★');

const A = [], B = [], C = [];
for (const x of g.取りこぼし) {
  const key = sheets[x.from].name + '|' + x.fromCell;
  const cell = sheets[x.from].data[x.fromCell];
  if (Excelのエラー.has(key)) A.push(x);                    /* A: Excel自身もエラー */
  else if (cell && cell.t !== 'e' && cell.v !== undefined && cell.v !== null && String(cell.v) !== '')
    B.push(x);                                              /* B: Excelは答えを出している */
  else C.push(x);                                           /* C: それ以外（空・もらい事故） */
}
const 一意 = (arr) => new Set(arr.map(x => x.from + ':' + x.fromCell)).size;
console.log('  ★A Excel自身も エラー … のべ ' + A.length + ' ／ 式 ' + 一意(A) + '★');
console.log('  ★B Excelは答えを出しているのに うちが落とした … のべ ' + B.length + ' ／ 式 ' + 一意(B) + '★');
console.log('  ★C もらい事故（空・値なし） … のべ ' + C.length + ' ／ 式 ' + 一意(C) + '★');
for (const x of B.slice(0, 10)) {
  const cell = sheets[x.from].data[x.fromCell];
  console.log('    B) ' + sheets[x.from].name + '!' + x.fromCell + '  Excelの答え=' + JSON.stringify(cell.v) + '  式) ' + String(x.f).slice(0, 90));
}
for (const x of C.slice(0, 5)) console.log('    C) ' + sheets[x.from].name + '!' + x.fromCell + '  式) ' + String(x.f).slice(0, 90));

/* ★Excelがエラーを保存しているのに 取りこぼしに入っていないセル（69-67=2）★ */
const 取りこぼしkey = new Set(g.取りこぼし.map(x => sheets[x.from].name + '|' + x.fromCell));
const 残り = [...Excelのエラー].filter(k => !取りこぼしkey.has(k));
console.log('★Excelはエラーなのに 取りこぼしに入っていない … ' + 残り.length + '個★');
for (const k of 残り) {
  const [nm, rc] = k.split('|');
  const i = sheets.findIndex(s => s.name === nm);
  const c = sheets[i].data[rc];
  console.log('    ' + nm + '!' + rc + '  Excelの答え=' + JSON.stringify(c.v) + '  式) ' + String(c.f).slice(0, 90));
}
/* ★IFERROR に隠れている物を数える（E2診断の材料）★ */
const 隠れ = g.取りこぼし.filter(x => /IFERROR|IFNA/i.test(String(x.f)));
console.log('★#REF! が IFERROR で隠れている式 … ' + 隠れ.length + '本★（画面には空が出るだけ＝気づけない）');
