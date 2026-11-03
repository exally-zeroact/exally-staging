/* ketsugou-yomu.test.mjs — ★Excel の 結合した マスを 読む★（2026-09-11）
 *
 *  ★★前は 1組も 読んで いませんでした★★
 *    `js/book-open.js` は `!merges` を ★一度も 見て いません★でした。
 *    ⇒ 実Excel で ★G1:H1★ と 2マス分に 広げて ある 所を 1マス分と して 扱い、
 *      字が 入らず ★`######`★ に なって いました。
 *    実測 … 司さんの 実物「給料表」G1「640,098 円」／ 幅59点（本当は 2マス分＝118点）
 *
 *  ★画面の 持ち方★（book.html と 同じ）
 *    元の マス … `mergeEnd = {r, c}`（右下の 場所）
 *    中の マス … `merged = {r, c}`（元の 場所）
 *
 *  ★見本★ tests/fixtures/ketsugou-2026-09-11.xlsx（★実Excel に 作らせました★）
 *
 *  ★この 台は エンジンを 建てません★＝`book-open.js` を そのまま 動かすだけ。
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(path.join(ROOT, 'package.json'));
const 開く元 = fs.readFileSync(path.join(ROOT, 'js/book-open.js'), 'utf-8');
const 見本 = path.join(ROOT, 'tests/fixtures/ketsugou-2026-09-11.xlsx');

let pass = 0, fail = 0;
const T = (n, f) => {
  try { f(); pass++; console.log('  ✓ ' + n); }
  catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + e.message); }
};

console.log('');
console.log('[ketsugou-yomu] ★Excel の 結合した マスを 読む★');

let 板 = null, XLSX = null, ws = null;
T('★見本が 在る／読める（空振りして いない）★', () => {
  if (!fs.existsSync(見本)) throw new Error('★見本が 無い … ' + 見本 + '★');
  XLSX = require_(path.join(ROOT, 'lib/xlsx.full.min.js'));
  const 場 = { XLSX: XLSX };
  new Function('self', 開く元)(場);
  const wb = XLSX.read(fs.readFileSync(見本), { type: 'buffer', cellFormula: true, cellNF: true, sheetStubs: false, cellStyles: true });
  ws = wb.Sheets[wb.SheetNames[0]];
  板 = 場.BookOpen.sheetToGrid(ws, wb.SheetNames[0], {});
  if (!板) throw new Error('★板が 作れない★');
});

T('★★借り物は 結合を くれる（★前提の 確かめ★）★★', () => {
  const m = ws['!merges'];
  if (!m || !m.length) {
    throw new Error('★`!merges` が 空★＝借り物が 変わったか `cellStyles` が 外れた'
      + '（この 見張りの 前提が 崩れました）');
  }
});

T('★★元の マスに mergeEnd が 付く★★', () => {
  const 数 = ws['!merges'].filter((m) => !(m.s.r === m.e.r && m.s.c === m.e.c)).length;
  let 付いた = 0;
  for (const k in 板.data) if (板.data[k] && 板.data[k].mergeEnd) 付いた++;
  if (付いた !== 数) {
    throw new Error('★' + 付いた + '組しか 付いて いない（ファイルには ' + 数 + '組）★');
  }
  console.log('      … ' + 付いた + '組（ファイルと 同じ 数）');
});

T('★★中の マスに merged が 付く（元を 指す）★★', () => {
  const 悪い = [];
  for (const m of ws['!merges']) {
    if (m.s.r === m.e.r && m.s.c === m.e.c) continue;
    for (let r = m.s.r; r <= m.e.r; r++) {
      for (let c = m.s.c; c <= m.e.c; c++) {
        if (r === m.s.r && c === m.s.c) continue;
        const 子 = 板.data[r + ',' + c];
        if (!子 || !子.merged) { 悪い.push(r + ',' + c + ' に merged が 無い'); continue; }
        if (子.merged.r !== m.s.r || 子.merged.c !== m.s.c) 悪い.push(r + ',' + c + ' が 別の 元を 指す');
      }
    }
  }
  if (悪い.length) throw new Error('★' + 悪い.length + '個 おかしい★  ' + 悪い.slice(0, 3).join(' ／ '));
});

T('★★1マスだけの「結合」は 結合に しない★★', () => {
  /* ★ファイルに 1マスだけの 記録が 入って いる 事が 在ります★＝
     それを 結合に すると ★元も 中も 同じ マス★に なり 画面が 壊れます。
     ★見本を 書き換えるのでは なく ここで 足します★
     （書き出しは 借り物が 落ちる事が 在る＝実測 2026-09-11） */
  const 場 = { XLSX: XLSX };
  new Function('self', 開く元)(場);
  const wb2 = XLSX.read(fs.readFileSync(見本), { type: 'buffer', cellFormula: true, cellNF: true, sheetStubs: false, cellStyles: true });
  const ws2 = wb2.Sheets[wb2.SheetNames[0]];
  ws2['!merges'] = (ws2['!merges'] || []).concat([{ s: { r: 6, c: 0 }, e: { r: 6, c: 0 } }]);
  const 板2 = 場.BookOpen.sheetToGrid(ws2, wb2.SheetNames[0], {});
  const 悪い = [];
  for (const k in 板2.data) {
    const c = 板2.data[k];
    if (c && c.mergeEnd) {
      const p = k.split(',');
      if (c.mergeEnd.r === +p[0] && c.mergeEnd.c === +p[1]) 悪い.push(k);
    }
  }
  if (悪い.length) {
    throw new Error('★1マスだけの 結合が ' + 悪い.length + '個 入った … ' + 悪い.join(' ') + '★'
      + '  ⇒ 元も 中も 同じ マスに なり 画面が 壊れます');
  }
  console.log('      … 1マスだけの 記録を 足しても 結合に ならない');
});

console.log('');
console.log('  ' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
