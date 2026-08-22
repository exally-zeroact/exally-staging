/* count-books.mjs — ★同じ名前のブックが何本あるかを数え、1本ずつ 中身を数える★（★読むだけ・1バイトも書かない★）
 *
 *  なぜ（2026-08-22）
 *    指示役「7シート・式19,323本」／私「14シート・式15,126本」で ★食い違った★。
 *    ★読み込みライブラリが .xlsb で落としている疑い★が出たので 数え直した。
 *    ⇒ ★原因は ライブラリではなく「同じ名前のファイルが5本 在る」事★だった。
 *       同じ道具で1本ずつ数えたら ★どちらの数字も そのまま再現できた★。
 *    ★教訓＝ファイル名で同じ物だと思うな。バイト数と sha を先に出せ★
 *
 *  使い方: node tools/count-books.mjs
 */
import fs from 'node:fs';
import crypto from 'node:crypto';
import path from 'node:path';
import { createRequire } from 'node:module';
const require_ = createRequire(import.meta.url);
const XLSX = require_('C:/Users/zeroa/exally-prod/lib/xlsx.full.min.js');

const 本 = [
  'C:/Users/zeroa/OneDrive/デスクトップ/代行計算表2026.xlsb',
  'C:/Users/zeroa/OneDrive/ドキュメント/代行計算表2026.xlsb',
  'C:/Users/zeroa/OneDrive/ZEROact税理士/2026/代行計算表2026.xlsb',
  'C:/Users/zeroa/OneDrive/ZEROact税理士/2025/代行計算表2025.xlsb',
  'C:/Users/zeroa/OneDrive/代行計算表.xlsb',
];
for (const f of 本) {
  if (!fs.existsSync(f)) { console.log('無い', f); continue; }
  const buf = fs.readFileSync(f);
  const sha = crypto.createHash('sha256').update(buf).digest('hex').slice(0, 16).toUpperCase();
  const st = fs.statSync(f);
  let シート = 0, 値 = 0, 式 = 0, 字 = 0, エラー = 0;
  try {
    const wb = XLSX.read(buf, { type: 'buffer', cellFormula: true, cellNF: true, sheetStubs: false });
    シート = wb.SheetNames.length;
    for (const name of wb.SheetNames) {
      const ws = wb.Sheets[name];
      if (!ws) continue;
      for (const k of Object.keys(ws)) {
        if (k[0] === '!') continue;
        const c = ws[k];
        値++;
        if (c.f) 式++;
        if (c.t === 'e') エラー++;
        字 += String(c.w !== undefined ? c.w : (c.v !== undefined ? c.v : '')).length;
      }
    }
  } catch (e) { console.log('読めない', f, e.message); continue; }
  console.log([path.basename(path.dirname(f)) + '/' + path.basename(f),
    'バイト=' + buf.length, 'sha=' + sha, '更新=' + st.mtime.toISOString().slice(0, 16),
    'シート=' + シート, '値の在るセル=' + 値, '★式=' + 式 + '★', 'エラー=' + エラー,
    '字数=' + 字].join(' / '));
}
