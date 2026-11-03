/* hashira-haba.test.mjs — ★列の 幅を 実Excel と 同じに する★（2026-09-11）
 *
 *  ★★何を 守るか★★
 *    ①★列の 幅を 読む★（`cellStyles` が 無いと SheetJS は 1つも くれない）
 *    ②★点(px)への 直し方★＝`ファイルの width × 一字の幅`（実Excel と 差 0.00）
 *    ③★一字の幅★＝その本の 既定の 字体で「0」を 測って ★切り捨て★
 *    ④★開いた ブックは その 字体で 描く★（うちの 字は 細くて 桁が 多く 入って いた）
 *    ⑤★マスの 余白は 3点★（実Excel の 境目から 割り出した）
 *    ⑥★`General` は 書式では ない★
 *    ⑦★書式を 持つ マスは 中の 数から 字を 作る★（出来た 字に また 書式を 掛けない）
 *
 *  ★★実Excel に 聞いた 事（docs/measured/toru-hashira-haba.ps1）★★
 *    字数     2     3     5     8.44   10    12     15     20     30
 *    実Excelの点 20.5 28.5  44.5  72     84.5  100.5  124.5  164.5  244.5
 *    ⇒ 一直線＝`点 = 字数 × 8 + 4.5`（游ゴシック 11pt）
 *    ファイルの width で 突き合わせると ★8列とも 差 0.00★
 *
 *  ★この 台は エンジンを 建てません★＝`book-open.js` を そのまま 動かして 板を 作るだけ。
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
const 見本 = path.join(ROOT, 'docs/measured/hashira-haba-2026-09-11.xlsx');
const 紙 = path.join(ROOT, 'docs/measured/golden-hashira-haba-2026-09-11.tsv');

let pass = 0, fail = 0;
const T = (n, f) => {
  try { f(); pass++; console.log('  ✓ ' + n); }
  catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + e.message); }
};

console.log('');
console.log('[hashira-haba] ★列の 幅を 実Excel と 同じに する★');

/* ★実Excel の 紙を 読む★（数を 書き込まない） */
const 実 = [];
for (const l of fs.readFileSync(紙, 'utf-8').split(/\r?\n/)) {
  if (!l || l.startsWith('#')) continue;
  const c = l.split('\t');
  if (c.length < 5) continue;
  実.push({ 列: Number(c[0]), 字数: Number(c[1]), 点: Number(c[4]) });
}

T('★紙が 読めて いる（空振りして いない）★', () => {
  if (実.length < 5) throw new Error('★実Excel の 紙が ' + 実.length + '行★');
});

T('★★①`cellStyles` が 在る（無いと 幅が 1つも 来ない）★★', () => {
  const 動く = 注記を外す(開く元);
  if (動く.indexOf('cellStyles: true') < 0) {
    throw new Error('★`cellStyles: true` が 無い★＝SheetJS は `!cols` を 作りません'
      + '（幅を 読む 所は 在るのに ★材料が 空★に なります）');
  }
});

T('★★②実Excel の 点と 合う（★1点も ずれない★）★★', () => {
  const 場 = { XLSX: require_(path.join(ROOT, 'lib/xlsx.full.min.js')) };
  new Function('self', 開く元)(場);
  const BO = 場.BookOpen;
  const wb = 場.XLSX.read(fs.readFileSync(見本), { type: 'buffer', cellFormula: true, cellNF: true, sheetStubs: false, cellStyles: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  /* ★字体は node に 無い＝逃げの 8（游ゴシック 11pt の 実測値）が 使われる★
     ⇒ ここで 見るのは ★直し方★（width × 8）／画面の 絵は 別に 取って 在る */
  const 板 = BO.sheetToGrid(ws, wb.SheetNames[0], {});
  const 違い = [];
  for (const r of 実) {
    const 出た = 板.colW[r.列 - 1];
    if (出た === undefined) continue;               /* ★標準の 幅の 列は ファイルに 書かれない★ */
    if (Math.abs(出た - r.点) > 0.5) 違い.push('列' + r.列 + ' 実Excel ' + r.点 + '点 ／ うち ' + 出た + '点');
  }
  if (違い.length) throw new Error('★' + 違い.length + '本 ずれた★  ' + 違い.join(' ／ '));
  console.log('      … ファイルに 幅が 在る 列は 実Excel と ★ぴったり★');
});

T('★★②-2 機械の 字体が 違っても（「0」が 10点＝Linux の WebKit の 形）列の 点は 実Excel と 同じ★★', () => {
  /* ★★2026-10-04 夜★★ 一字の幅を ★お客さんの 機械で 測って★ いたので、
       Linux の WebKit では 標準の 列が 74→93点 ⇒ #### が 出ず 桁が 多く 出た（kakareta-ji の 7組）。
     ⇒ 表に 在る 字体（游ゴシック 11＝8）は 機械で 測らない。ここでは ★機械が「0」を 10.2点と 測る★ 場を 作って 開く
     ★見本の 既定の 字体は 游ゴシック 11★（下で 確かめる＝本番は 開く 時に これを 渡す） */
  const 測り = { 回: 0 };
  const 場 = { XLSX: require_(path.join(ROOT, 'lib/xlsx.full.min.js')),
    document: { createElement: () => ({ getContext: () => ({ font: '', measureText: () => { 測り.回++; return { width: 10.2 }; } }) }) } };
  new Function('self', 開く元)(場);
  const wb = 場.XLSX.read(fs.readFileSync(見本), { type: 'buffer', cellFormula: true, cellNF: true, sheetStubs: false, cellStyles: true });
  const 既定 = wb.Styles && wb.Styles.Fonts && wb.Styles.Fonts[0];
  if (!既定 || 既定.name !== '游ゴシック' || Number(既定.sz) !== 11) throw new Error('★見本の 既定の 字体が 游ゴシック 11 で ない★ ' + JSON.stringify(既定));
  const 板 = 場.BookOpen.sheetToGrid(wb.Sheets[wb.SheetNames[0]], wb.SheetNames[0], {});
  const 違い = [];
  let 見た = 0;
  for (const r of 実) {
    const 出た = 板.colW[r.列 - 1];
    if (出た === undefined) continue;
    見た++;
    if (Math.abs(出た - r.点) > 0.5) 違い.push('列' + r.列 + ' 実Excel ' + r.点 + '点 ／ うち ' + 出た + '点');
  }
  if (見た < 5) throw new Error('★比べた 列が ' + 見た + '本★（空振り）');
  if (違い.length) throw new Error('★機械の 字体で 列が 動いた★ ' + 違い.length + '本  ' + 違い.join(' ／ '));
  console.log('      ... 機械の「0」が 10.2点でも ' + 見た + '列 とも 実Excel と 同じ（機械で 測った 回数 ' + 測り.回 + '）');
});

T('★★③SheetJS の `wpx` を 使って いない（★1字 14点と 思い込んで いる★）★★', () => {
  const 動く = 注記を外す(開く元);
  if (/col\.wpx/.test(動く)) {
    throw new Error('★`col.wpx` を 使って いる★＝5字の 列で 実Excel 44.5点 ／ SheetJS 78点'
      + '（★幅が 足りて しまい #### に ならない★）');
  }
});

T('★★④一字の幅は 切り捨て（游ゴシック 8 ／ Calibri 7）★★', () => {
  const 動く = 注記を外す(開く元);
  if (動く.indexOf('Math.floor') < 0 || 動く.indexOf("measureText('0')") < 0) {
    throw new Error('★「0」を 測って 切り捨てて いない★'
      + '（游ゴシック11pt 8.153→8 ／ Calibri11pt 7.430→7 で 実Excel と 合う）');
  }
});

T('★★⑤開いた ブックは その 字体で 描く★★', () => {
  const 動く = 注記を外す(本, { html: true });
  if (動く.indexOf('既定の字体名') < 0 || 動く.indexOf('既定の字大') < 0) {
    throw new Error('★ブックの 字体を 使って いない★'
      + '＝うちの 字は 細く、同じ 幅に ★桁が 多く 入ります★'
      + '（実測 … 実Excel「1.001193」8桁 ／ うち「1.001193321」11桁）');
  }
});

T('★★⑤-2 字体を 作る 道が ★1本★（結合した マスも 同じ）★★', () => {
  /* ★★2026-09-11 ここで つまずきました★★
       ふつうの マスだけ 直して、★結合した マスは うちの 字の まま★でした
       ⇒ 同じ 表の 中で ★字が 2種類★に なり、桁数も 揃いません
     ⇒ `_マスの字大` `_マスの書体` の 1本に まとめた
     ★探すのは 注記の 外だけ★（説明文の 字で 緑に なるのを 防ぐ … 同じ日に 2回 踏んだ） */
  const 動く = 注記を外す(本, { html: true });
  for (const 名 of ['function _マスの字大', 'function _マスの書体']) {
    if (動く.indexOf(名) < 0) throw new Error('★' + 名 + ' が 無い★');
  }
  /* ★古い 書き方が 残って いないか★＝残ると そこだけ 別の 字に なる */
  const 残り = (動く.match(/cell\.fontSize\s*\|\|\s*12/g) || []).length;
  if (残り) {
    throw new Error('★まだ 12点を 直に 使って いる 所が ' + 残り + '本★'
      + '＝そこだけ ブックの 字体を 見ません');
  }
  /* ★書体を 直に 書いて いる 描き所が 無いか★（部品・目盛りなどは 別＝マスの 字だけ 見る） */
  const 直書き = (動く.match(/\(cell\.fontSize[^)]*\)\s*\*\s*scale/g) || []).length;
  if (直書き) throw new Error('★マスの 字を 直に 作って いる 所が ' + 直書き + '本★');
  console.log('      … 字の 大きさも 書体も ★1本★');
});

T('★★⑥マスの 余白は 3点（実Excel の 境目から 割り出した）★★', () => {
  const 動く = 注記を外す(本, { html: true });
  const m = /var マスの余白 = (\d+)/.exec(動く);
  if (!m) throw new Error('★`マスの余白` が 無い★');
  if (Number(m[1]) !== 3) {
    throw new Error('★余白が ' + m[1] + '★＝実Excel の 境目は ★0.6〜3.5★'
      + '（12.3% は 幅40.5で 出ず 44.5で 出る／字の 点 39.99）');
  }
  /* ★引き算の 所が 全部 この 数を 使って いるか★＝1つ 取り残すと そこだけ 別の 幅に なる */
  if (/w - 8[^0-9]/.test(動く.slice(動く.indexOf('var マスの余白'), 動く.indexOf('var マスの余白') + 2500))) {
    throw new Error('★まだ 8 を 引いて いる 所が 在る★');
  }
});

T('★★⑦`General` は 書式では ない★★', () => {
  const 動く = 注記を外す(本, { html: true });
  if (動く.indexOf("=== 'General'") < 0) {
    throw new Error("★`General` を 本物の 書式として 扱って いる★"
      + '＝幅で 桁を 決める 道を 通らず、実Excel「1.001193」が「1.001193321」に なります');
  }
});

T('★★⑧書式を 持つ マスは 中の 数から 字を 作る★★', () => {
  const 動く = 注記を外す(本, { html: true });
  if (!/cell\.numFmt && cell\.numFmt !== 'General' && typeof cell\.v === 'number'/.test(動く)) {
    throw new Error('★出来上がった 字に また 書式を 掛けて いる★'
      + '＝"1,234,567.89" を 数と 読むと ★1★（コンマで 止まる）⇒ 画面に "1.00"');
  }
});

T('★★⑨書式を 持つ マスも #### に なる★★', () => {
  const 動く = 注記を外す(本, { html: true });
  if (/!cell\.numFmt && _数が入らないか/.test(動く)) {
    throw new Error('★書式付きは #### に ならない ままに なって いる★'
      + '（実Excel は #,##0.00 の 1234567.891 を 幅5の 列で ★####★）');
  }
  /* ★2026-10-04 夜★ ①b で マスを 4つ目に 渡す 形に なった（書式を 持つ マスは 実Excel の 余白で 判じる） */
  if (!/_数が入らないか\(raw, display, w(, cell)?\)/.test(動く)) throw new Error('★#### の 判じが 無い★');
});

T('★★⑩新しく 作る ブックも 実Excel と 同じ 既定★★', () => {
  /* ★★司さん（2026-09-11）「フォントと揃えとけや」★★
     ★実Excel に 聞いた 既定★（新しい ブックを 作って COM で 測った）
       標準の 列幅 8.44字 ＝★72点★ ／ 標準の 行高 17.7pt ＝★23.5点★
       字体 ★游ゴシック 11pt★
     ★前★ … 列80点・行22点・Noto Sans JP 12点
       ⇒ 広くて 細いので ★同じ マスに 桁が 多く 入って★ いた
       実測 … =1/3 … 実Excel「0.333333」6桁 ／ うち「0.333333333」9桁
     ★合わせた 後★ … 本物の ブラウザで 6本 打って ★実Excel と 1文字も 違わない★
       =1/3→0.333333 ／ =206800/1.1→188000 ／ =1.64E-14→1.64E-14
       ／ =100000→100000 ／ =1234567.891→1234568 ／ =XIRR(…)→1.001193 */
  const 動く = 注記を外す(本, { html: true });
  const m = /var ROW_H = (\d+), COL_W = (\d+)/.exec(動く);
  if (!m) throw new Error('★ROW_H / COL_W が 読めない★');
  if (Number(m[2]) !== 72) {
    throw new Error('★列の 既定が ' + m[2] + '点★（実Excel は ★72点★＝8.44字）');
  }
  if (Number(m[1]) < 23 || Number(m[1]) > 25) {
    throw new Error('★行の 既定が ' + m[1] + '点★（実Excel は ★23.5点★＝17.7pt）');
  }
  /* ★立てて いる 所の 近くで 見ます★＝repo 全体から 探すと
     ★別の 所に 在る 同じ 字★（測りの 説明など）で 緑に なります（実際 なりました） */
  const 書体頭 = 動く.indexOf('function _マスの書体');
  const 書体中 = 書体頭 >= 0 ? 動く.slice(書体頭, 書体頭 + 500) : '';
  if (書体中.indexOf("'游ゴシック'") < 0) {
    throw new Error('★`_マスの書体` の 既定が 游ゴシック で ない★＝うちの 字は 細く 桁が 多く 入ります');
  }
  if (!/\|\| 11;/.test(動く)) {
    throw new Error('★既定の 大きさが 11pt で ない★（実Excel は 11pt）');
  }
  console.log('      … 列72点／行24点／游ゴシック 11pt（実Excel と 同じ）');
});

console.log('');
console.log('  ' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
