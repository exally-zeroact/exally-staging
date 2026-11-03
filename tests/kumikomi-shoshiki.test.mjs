/* kumikomi-shoshiki.test.mjs — ★国で 中身が 変わる 書式番号／15桁に 丸めてから 出す★（2026-09-11）
 *
 *  ★★どうやって 見つけたか★★
 *    ★司さんの 実物（代行計算表2026.xlsb）の 出る字を 実Excel と 1マスずつ 突き合わせた★
 *    （今まで 実物で 合わせたのは ★答え★だけ＝08-29 の 19,323/19,323。
 *      ★出る字は 一度も 突き合わせて いませんでした★）
 *    「計算」の 板 5,437マス ⇒★437マス 違った★ ⇒ 直して ★5,436/5,437★
 *
 *  ★★① 書式番号 38番（国で 中身が 変わる）★★
 *    ★5〜8番と 37〜40番は 国で 中身が 変わります★（Microsoft の 決め）
 *    借り物（SheetJS）は ★世界共通（英語）の 表★を 持って いる
 *      38番 … `#,##0 ;[Red](#,##0)`  ⇒ 後ろに 空白／マイナスは 括弧
 *    ★日本語の Excel の 38番★（実Excel に 打たせて 測った）
 *      `#,##0;[赤]-#,##0`            ⇒ 空白なし／マイナスは -
 *      2220→「2,220」／-1234.5→「-1,235」／0→「0」
 *    ⇒ 実物で ★785マス★ 使って いて ★323マス 違って 見えて★ いた
 *    ★`SSF.load` では 届きません★＝.xlsb を 読む 所が 自前の 写しを 持つ（実測）
 *      ⇒★読んだ 後に 直します★
 *
 *  ★★② 15桁に 丸めてから 出す★★
 *    実Excel は ★15桁に 丸めてから★ 出します
 *      34779.49999999999（16桁）⇒ 15桁で 34779.5 ⇒★34,780★
 *    うちは 丸めずに 出して ★34,779★＝★お金が 1円 ずれて いました★
 *    `forDisplay` は 前から 丸めて いた ⇒★書式を 持つ マスだけ 素通り★（また 道が 2本）
 *
 *  ★見本★ tests/fixtures/kumikomi-38-2026-09-11.xlsx
 *    ★実Excel に 作らせました★（司さんの 実物は repo に 入れません＝取引先の 名前と 金額）
 *
 *  ★この 台は エンジンを 建てません★＝`book-open.js` を そのまま 動かすだけ。
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';
import { 注記を外す } from '../scripts/lib/chuki.mjs';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(path.join(ROOT, 'package.json'));
const 本 = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf-8');
const 開く元 = fs.readFileSync(path.join(ROOT, 'js/book-open.js'), 'utf-8');
const 見本 = path.join(ROOT, 'tests/fixtures/kumikomi-38-2026-09-11.xlsx');

let pass = 0, fail = 0;
const T = (n, f) => {
  try { f(); pass++; console.log('  ✓ ' + n); }
  catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + e.message); }
};

console.log('');
console.log('[kumikomi-shoshiki] ★国で 変わる 書式／15桁に 丸めてから 出す★');

/* ★実Excel に 打たせて 測った 正解★（このファイルを 作った時に 一緒に 測った） */
const 正 = {
  A1: { 値: 2220, 字: '2,220' },
  A2: { 値: -1234.5, 字: '-1,235' },
  A3: { 値: 0, 字: '0' },
  B1: { 値: 34779.49999999999, 字: '34,780 ' },
  B2: { 値: 12169.499999999998, 字: '12,170 ' },
  B3: { 値: 11903.499999999998, 字: '11,904 ' },
};

let 板 = null, XLSX = null;
T('★見本が 在る／読める（空振りして いない）★', () => {
  if (!fs.existsSync(見本)) throw new Error('★見本が 無い … ' + 見本 + '★');
  XLSX = require_(path.join(ROOT, 'lib/xlsx.full.min.js'));
  const 場 = { XLSX: XLSX };
  new Function('self', 開く元)(場);
  const wb = XLSX.read(fs.readFileSync(見本), { type: 'buffer', cellFormula: true, cellNF: true, sheetStubs: false, cellStyles: true });
  板 = 場.BookOpen.sheetToGrid(wb.Sheets[wb.SheetNames[0]], wb.SheetNames[0], {});
  if (!板 || !板.data['0,0']) throw new Error('★見本から マスが 作れない★');
});

T('★★①借り物は まだ 世界共通の 字を 返す（★直す 前の 姿★）★★', () => {
  const wb = XLSX.read(fs.readFileSync(見本), { type: 'buffer', cellFormula: true, cellNF: true, sheetStubs: false, cellStyles: true });
  const z = String(wb.Sheets[wb.SheetNames[0]].A1.z);
  if (z !== '#,##0 ;[Red](#,##0)') {
    throw new Error('★借り物の 字が 変わった … ' + z + '★'
      + '＝この 見張りの 前提が 崩れました（借り物の 版を 上げた？）');
  }
});

T('★★①うちは 日本語の Excel の 38番に 直す★★', () => {
  const z = String(板.data['0,0'].numFmt);
  if (z !== '#,##0;[Red]-#,##0') {
    throw new Error('★' + z + '★（実Excel（日本語）は `#,##0;[赤]-#,##0`＝空白なし／マイナスは -）'
      + '  ⇒ 直さないと 実物で ★785マス★ が「2,220 」の ように 空白つきに なります');
  }
});

T('★★②15桁に 丸めてから 出す（book.html）★★', () => {
  const 動く = 注記を外す(本, { html: true });
  if (動く.indexOf('function _十五桁') < 0) throw new Error('★`_十五桁` が 無い★');
  /* ★書式を 掛ける 所で 通って いるか★＝ここを 外すと お金が 1円 ずれる */
  if (動く.indexOf('XLSX.SSF.format(BookOpen.withWeekday(fmt, +raw), _十五桁(+raw))') < 0) {
    throw new Error('★書式を 掛ける 前に 15桁に 丸めて いない★'
      + '＝34779.49999999999 が 実Excel「34,780」／うち「34,779」に なります');
  }
});

T('★★②実Excel の 出る字と 合う（6本）★★', () => {
  /* ★画面と 同じ 道で 字を 作ります★＝book.html の `fmtForDisplay` を 切り出して 使う */
  const 切る = (名) => {
    const i = 本.indexOf('function ' + 名 + '(');
    if (i < 0) throw new Error('★' + 名 + ' が 無い★');
    let 深 = 0, j = 本.indexOf('{', i);
    for (let k = j; k < 本.length; k++) {
      if (本[k] === '{') 深++;
      else if (本[k] === '}') { 深--; if (!深) return 本.slice(i, k + 1); }
    }
    throw new Error('★' + 名 + ' を 切り出せない★');
  };
  const 作る = new Function('XLSX', 'BookOpen',
    切る('_十五桁') + ' ' + 切る('fmtForDisplay') + ' ' + 切る('applyNumFmt') + ' ' + 切る('forDisplay')
    + ' ' + 切る('excelGeneral') + ' var GEN_枠 = 11; return fmtForDisplay;');
  let 出す;
  try {
    出す = 作る(XLSX, { isOpened: () => true, withWeekday: (f) => f });
  } catch (e) {
    throw new Error('★切り出せません … ' + e.message + '★（形が 変わったら ここを 直す）');
  }
  const 違い = [];
  for (const [名, x] of Object.entries(正)) {
    const p = 名 === 'A1' || 名 === 'A2' || 名 === 'A3' ? '0,0' : '0,1';
    const fmt = 板.data[p].numFmt;
    const 出 = String(出す(x.値, fmt, 11));
    if (出 !== x.字) 違い.push(名 + ' 実Excel「' + x.字 + '」／うち「' + 出 + '」');
  }
  if (違い.length) throw new Error('★' + 違い.length + '本 違う★  ' + 違い.join(' ／ '));
  console.log('      … 2,220／-1,235／0／34,780 ／12,170 ／11,904 （実Excel と 同じ）');
});

console.log('');
console.log('  ' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
