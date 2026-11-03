/* hozon-de-kieru.test.mjs — ★うちで 作った 物は 保存すると 消える★（2026-09-06）
 *
 *  ★★何を 見るか★★
 *    ★預かった ファイルは 1バイトも 壊していない★（2026-08-30 実測ずみ）。
 *    ★でも うちで 足した 物は 書き出すと 消える★。
 *    ⇒ 画面は「値だけ 書き換えました」としか 言わないので
 *      ★お客さんは 消えた事に 気づけない★＝★黙って 消える★
 *
 *  ★★一番 確かで 一番 起こりやすい 所（指示役の 裁定）★★
 *    `js/book-open.js` の `saveXlsxLike()`
 *      `if (opened.sheetNames.indexOf(sh.name) < 0) return;  // 元に無いシートは触らない`
 *    ⇒★★うちで 足した シートは 丸ごと 飛ばされる★★
 *    ⇒「集計用の シートを 1枚 足す」は ★お客さんが 普通に やる事★
 *
 *  ★この 試験が する事★
 *    ①本物の .xlsx を 開く
 *    ②★新しいシートを 1枚 足して 字を 書く★
 *    ③保存して ★出た zip の 中に そのシートが 在るか★を 見る
 *    ④★元から 在った シートの 値は ちゃんと 書き戻っているか★も 見る
 *      （＝「全部 落ちる」では なく「新しい物だけ 落ちる」事を はっきりさせる）
 *
 *  ★★この 試験は 今 赤です（それが 正しい）★★
 *    直す 順番は 指示役の 裁定どおり
 *      ①測る（ここ）→ ②★画面で「入りません」と 言う★ → ③zip へ 足す（別の 回）
 *    ⇒ ③が 出来たら ここが 緑に なる。★それまで 赤のまま★。
 *    ⇒ だから ★run.js には まだ 登録しない★（CI を 赤に しない）
 *      ＝★「出来ていない事を 隠さない」代わりに「赤を 常態に しない」★
 *
 *  使い方: node tests/hozon-de-kieru.test.mjs
 */
import { createRequire } from 'node:module';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const require_ = createRequire(import.meta.url);
const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const XLSX = require_(path.join(ROOT, 'lib/xlsx.full.min.js'));
const ZipSurgeon = require_(path.join(ROOT, 'lib/zip-surgeon.js'));
const XlsxEdit = require_(path.join(ROOT, 'lib/xlsx-edit.js'));
const XlsbEdit = require_(path.join(ROOT, 'lib/xlsb-edit.js'));

const box = { XLSX, ZipSurgeon, XlsxEdit, XlsbEdit };
new Function('self', 'window', fs.readFileSync(path.join(ROOT, 'js/book-open.js'), 'utf8')
  + '\n;self.__BookOpen = self.BookOpen;')(box, box);
const BookOpen = box.__BookOpen;

const FIX = path.join(ROOT, 'tests/fixtures/book-open-sample.xlsx');
const 元 = new Uint8Array(fs.readFileSync(FIX));
const 偽ファイル = (name, b) => ({ name, arrayBuffer: () => Promise.resolve(b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength)) });

let pass = 0, fail = 0;
const T = async (n, fn) => { try { await fn(); pass++; console.log('  ✓ ' + n); } catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + (e && e.message)); } };

console.log('\n[hozon-de-kieru] うちで 足した 物は 書き出すと 消えるか');

const 開いた = await BookOpen.openFile(偽ファイル('sample.xlsx', 元));
const 元のシート名 = (BookOpen.current() || {}).sheetNames || [];
console.log('  元の シート … ' + 元のシート名.join(' / '));

/* ★画面と 同じ 形の sheets を 作る★＝元のシート ＋ ★うちで 足した 1枚★ */
const sheets = 元のシート名.map((n) => ({ name: n, data: {} }));
/* 元のシートの 1マスを 直す（★書き戻る 事の 確かめ★） */
sheets[0].data['0,0'] = { v: 'なおした', f: '', d: 'なおした' };
/* ★うちで 足した シート★ */
const 足したシート名 = '★うちで 足した★';
sheets.push({ name: 足したシート名, data: { '0,0': { v: 'ここに 集計を 書いた', f: '', d: 'ここに 集計を 書いた' } } });

await T('★材料が 正しい（元に その シートは 無い）', () => {
  if (元のシート名.indexOf(足したシート名) >= 0) throw new Error('元から 在る＝検査が 空振り');
  if (!元のシート名.length) throw new Error('元の シートが 読めていない');
});

const 出 = await BookOpen.saveOpened(sheets);
const バイト = 出 && 出.bytes ? 出.bytes : 出;
console.log('  書き出した … ' + バイト.length.toLocaleString() + 'バイト（元 ' + 元.length.toLocaleString() + 'バイト）');

/* ★出た zip を 開いて 中を 見る★＝「在るはず」ではなく ★在るか★ */
const 中 = XLSX.read(バイト, { type: 'array', bookSheets: true });
console.log('  出た シート … ' + (中.SheetNames || []).join(' / '));

await T('★元から 在った シートの 直しは 書き戻っている（全部 落ちる わけでは ない）', () => {
  const b = XLSX.read(バイト, { type: 'array' });
  const ws = b.Sheets[元のシート名[0]];
  const c = ws && ws.A1;
  const v = c ? String(c.v) : '';
  if (v !== 'なおした') throw new Error('A1 が 書き戻っていない（今 「' + v + '」）');
});

await T('★★うちで 足した シートが 書き出した ファイルに 在る★★', () => {
  if ((中.SheetNames || []).indexOf(足したシート名) < 0) {
    throw new Error('★★消えました★★＝うちで 足した シート「' + 足したシート名 + '」が'
      + ' 書き出した ファイルに ありません'
      + '\n     出た シート … ' + (中.SheetNames || []).join(' / ')
      + '\n     場所 … js/book-open.js saveXlsxLike()'
      + '\n            if (opened.sheetNames.indexOf(sh.name) < 0) return;'
      + '\n     ⇒ ★お客さんから 見れば「自分の 仕事が 消えた」★'
      + '\n     ⇒ しかも 画面は ★成功したように 見える★');
  }
});

console.log('\n' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
