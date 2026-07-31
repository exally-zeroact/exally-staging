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
const order = [];
let row = 1;
for (const c of cases) {
  if (c.volatile) continue;                              // 揮発性は対象外
  cells['T' + row] = { f: c.f };   // ★キャッシュ値は入れない(必ず計算させる)
  order.push(c.id);
  row++;
}
fs.writeFileSync(outPath, XlsxIO.writeBook({ sheets: [{ name: 'Sheet1', cells }] }));
fs.writeFileSync(outPath + '.order.json', JSON.stringify(order, null, 1));
console.log(`${outPath} に ${order.length} 本の式を書き出した`);
