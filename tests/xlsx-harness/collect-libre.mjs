// collect-libre.mjs — LibreOffice が吐いた csv を読んで golden/libreoffice-<版>.json にする。
//   ★これは Excel の真値ではない。「LibreOffice ではこうなる」という別の版の記録。
//   csv は表示テキストなので、型は数値に見えるかどうかで推定する(Excelのgoldenほど厳密ではない)。
import fs from 'node:fs';
import path from 'node:path';

const [, , csvPath, version, outDir] = process.argv;
if (!csvPath || !version || !outDir) { console.error('使い方: node collect-libre.mjs <csv> <版> <出力先>'); process.exit(1); }

function parseCsv(text) {
  const rows = [];
  let row = [], field = '', q = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (q) {
      if (ch === '"') { if (text[i + 1] === '"') { field += '"'; i++; } else q = false; }
      else field += ch;
    } else if (ch === '"') q = true;
    else if (ch === ',') { row.push(field); field = ''; }
    else if (ch === '\n') { row.push(field); rows.push(row); row = []; field = ''; }
    else if (ch !== '\r') field += ch;
  }
  if (field !== '' || row.length) { row.push(field); rows.push(row); }
  return rows;
}

const ERRS = ['#DIV/0!', '#N/A', '#NAME?', '#NULL!', '#NUM!', '#REF!', '#VALUE!', '#SPILL!', '#CALC!', 'Err:501', 'Err:502'];
const order = JSON.parse(fs.readFileSync(csvPath.replace(/\.csv$/, '.xlsx.order.json'), 'utf8'));
const rows = parseCsv(fs.readFileSync(csvPath, 'utf8'));

const cases = {};
order.forEach((o, i) => {
  const id = typeof o === 'string' ? o : o.id;
  const raw = (rows[i] && rows[i][19] !== undefined) ? rows[i][19] : '';   // T列 = 20列目
  let rec;
  if (ERRS.includes(raw)) rec = { t: 'e', v: raw };
  else if (raw === 'TRUE' || raw === 'FALSE') rec = { t: 'b', v: raw === 'TRUE' };
  else if (raw !== '' && !Number.isNaN(Number(raw))) rec = { t: 'n', v: Number(raw) };
  else rec = { t: 's', v: raw };
  rec.text = raw;
  cases[id] = rec;
});

const meta = {
  source: 'libreoffice-headless',
  version,
  caution: '★Excelの真値ではない。LibreOfficeという別の版の答え。ここが緑でも「Excel一致」とは呼ばない。',
  valueSource: 'csv(表示テキスト)なので型は推定。Excel golden ほど厳密ではない',
  caseCount: order.length
};
fs.mkdirSync(outDir, { recursive: true });
const out = path.join(outDir, `libreoffice-${version}.json`);
fs.writeFileSync(out, JSON.stringify({ meta, cases }, null, 1) + '\n');
console.log(`${out} に ${order.length} 件`);
