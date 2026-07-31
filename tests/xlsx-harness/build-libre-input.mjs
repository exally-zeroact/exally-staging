// build-libre-input.mjs — LibreOffice に食わせる .xlsx を作る。
//   入力セルは Excel の golden 生成器と同じ配置(A1:H8)、式は T列に同じ順番で1本ずつ。
//   ★キャッシュ値は入れない。LibreOffice に必ず計算させるため。
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createRequire } from 'node:module';
import { loadCases } from './run-exally.mjs';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const require = createRequire(pathToFileURL(path.join(__dirname, '..', '..', 'package.json')));
const XlsxIO = require(path.join(__dirname, 'xlsx-io.js'));

const outPath = process.argv[2];
if (!outPath) { console.error('使い方: node build-libre-input.mjs <出力.xlsx>'); process.exit(1); }

const { inputs, cases } = loadCases();
const cells = {};
for (const [addr, spec] of Object.entries(inputs.cells)) {
  if (spec.t === 'n') cells[addr] = { v: spec.v, t: 'n' };
  else if (spec.t === 'b') cells[addr] = { v: spec.v, t: 'b' };
  else cells[addr] = { v: String(spec.v), t: 's' };     // 文字列は文字列のまま
}
//  ★スピルする式は1本ずつ別の列(AD以降)に置く。T列に並べると隣の行へ溢れて
//    次の式を潰し、Excelで開いた時に #SPILL! だらけになる(golden生成器と同じ決まり)。
function colName(c){ var s=''; c=c+1; while(c>0){ var m=(c-1)%26; s=String.fromCharCode(65+m)+s; c=Math.floor((c-1)/26);} return s; }
const order = [];
let row = 1, spillCol = 29;                              // 0始まり: 29 = AD列
for (const c of cases) {
  if (c.volatile) continue;                              // 揮発性は対象外
  const addr = c.spill ? (colName(spillCol++) + '1') : ('T' + row++);
  cells[addr] = { f: c.f };                              // ★キャッシュ値は入れない(必ず計算させる)
  order.push({ id: c.id, addr });
}
fs.writeFileSync(outPath, XlsxIO.writeBook({ sheets: [{ name: 'Sheet1', cells }] }));
fs.writeFileSync(outPath + '.order.json', JSON.stringify(order, null, 1));
console.log(`${outPath} に ${order.length} 本の式を書き出した`);
