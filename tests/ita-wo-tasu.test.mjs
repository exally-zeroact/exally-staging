/* ita-wo-tasu.test.mjs — ★足した 板が 書き出す 先に 入るか★（2026-09-21）
 *
 *  ★★なぜ 要るか★★
 *    司さん「★全部 保存しろや、断る 理由が なんか あるんか★」（ア）
 *    `lib/hairanai.js` の 頭（2026-09-06 実測）
 *      「新しいシートを 1枚 足して 保存 ⇒ ★出た ファイルに 無い★」
 *
 *  ★★穴は 2つ でした★★（★1つ 直しただけでは 出ませんでした★）
 *    ⑴`js/book-open.js` `saveXlsxLike()`
 *        `if (opened.sheetNames.indexOf(sh.name) < 0) return;`
 *    ⑵`js/book-open.js` `anyChanged()`
 *        `if (opened.sheetNames.indexOf(sheets[i].name) < 0) continue;`
 *      ⇒`saveOpened()` は 1つも 変わって いなければ ★元の バイト列を そのまま 返す★
 *      ⇒★板を 足しただけ では 「変わって いない」★に なり
 *        ★出た ファイルが 元と 1バイトも 同じ★でした（sha256 が 一致・実測）
 *    ⇒★記憶「作る道が 2本 在る時は 両方 直せ」★
 *
 *  ★★この 門が 守る 事★★
 *    ①足した 板の 部品が ★包みに 入る★
 *    ②★元の 部品が 1つも 減らない★
 *    ③★★元の 部品の 中身が 1バイトも 変わらない★★
 *       ＝★判子・罫線・塗り・グラフは 元の 部品の 中★
 *       ＝★直すのは 仕組みの 3本だけ★（`[Content_Types].xml` / rels / `workbook.xml`）
 *    ④`<sheets>` と rels と 型 が ★3つとも★ 直る（★1つでも 抜けると 実Excel は 修復を 言う★）
 *    ⑤★番号が ぶつからない★（`sheetN` / `rIdN` / `sheetId`）
 *    ⑥★もう 在る 名前は 足せない★
 *
 *  ★★見て いない 事★★
 *    ・★実Excel が 開くかは 測れません★（COM が 要る＝経営者1 の 持ち場）
 *      ⇒`docs/measured/toru-jitsu-excel-ga-shuufuku-shita-ka.ps1`（93）
 *    ・★お客さんの 道（画面）では ここでは 押して いません★
 *      ⇒`docs/measured/hakaru-tashita-ita-ga-hozon-sareru-ka.mjs`
 *    ・★足した 板に 飾りは 書きません★＝値と 式だけ
 *    ・★`xl/calcChain.xml` は 触りません★＝★消して 良いか 測って いません★
 *
 *  使い方: node tests/ita-wo-tasu.test.mjs
 *          node tests/ita-wo-tasu.test.mjs --self-test
 */
import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(path.join(ROOT, 'package.json'));
const Z = require_(path.join(ROOT, 'lib/zip-surgeon.js'));
const E = require_(path.join(ROOT, 'lib/xlsx-edit.js'));

let pass = 0, fail = 0;
const NL = String.fromCharCode(10);
const 待 = async (n, f) => {
  try { await f(); pass++; console.log('  ok   ' + n); }
  catch (e) { fail++; console.log('  NG   ' + n + NL + '       ' + e.message); }
};

const 材料道 = path.join(ROOT, 'tests/fixtures/kazari-hiraku3.xlsx');
const 中 = fs.existsSync(材料道) ? fs.readFileSync(材料道) : null;

console.log('');
console.log('[ita-wo-tasu] ★足した 板が 書き出す 先に 入るか★');

await 待('★材料が 在る（空振りして いない）★', async () => {
  if (!中) throw new Error('★材料が 無い★ ' + 材料道);
  const h = crypto.createHash('sha256').update(中).digest('hex');
  if (h !== 'affd42e9f5c6de9d5565c8ffd623ffd2e37a7ba5ca8820fdd5c971b15ff93053') {
    throw new Error('★材料が 入れ替わって います★ ' + h);
  }
});

/* ★元の 包みを 全部 控えます★ */
const 元z = Z.read(new Uint8Array(中));
const 元の名 = 元z.names().slice();
const 元の中身 = {};
for (const n of 元の名) 元の中身[n] = await 元z.bytes(n);

/* ★★板を 1枚 足して 書き出す★★ */
const 出 = await (async () => {
  const book = await E.open(new Uint8Array(中));
  const 足 = E.板を足す(book, 'Tashita', {
    A1: { v: 123, t: 'n' }, B1: { v: 'tashita', t: 's' }, A2: { v: 2.5, t: 'n' },
  });
  const r = await E.save(book);
  return { 足: 足, z: Z.read(r.bytes), 大: r.bytes.length };
})();
console.log('      ＝ 足した 部品 ' + 出.足.部品 + ' ／ ' + 出.足.rId
  + ' ／ sheetId ' + 出.足.sheetId + ' ／ 出た 包み ' + 出.大 + 'B');

await 待('★★足した 板の 部品が 包みに 入って いる★★', async () => {
  if (!出.z.has(出.足.部品)) throw new Error('★' + 出.足.部品 + ' が 無い★');
  const t = await 出.z.text(出.足.部品);
  if (t.indexOf('<v>123</v>') < 0) throw new Error('★A1 の 123 が 無い★ ' + t.slice(0, 120));
  if (t.indexOf('tashita') < 0) throw new Error('★B1 の 字が 無い★');
  if (t.indexOf('t="inlineStr"') < 0) {
    throw new Error('★字が `inlineStr` で 入って いません★'
      + '／★共有文字列を 触ると 元の 板の 番号が ずれます★');
  }
});

await 待('★★元の 部品が 1つも 減って いない★★', async () => {
  const 無 = 元の名.filter((n) => !出.z.has(n));
  if (無.length) throw new Error('★' + 無.length + '本 消えました★ ' + 無.join(' '));
});

await 待('★★★元の 部品の 中身が 1バイトも 変わって いない★★★'
  + '（★直すのは 仕組みの 3本だけ★）', async () => {
  const 直して良い = ['[Content_Types].xml', 'xl/_rels/workbook.xml.rels', 'xl/workbook.xml'];
  const 違 = [];
  for (const n of 元の名) {
    if (直して良い.indexOf(n) >= 0) continue;
    const a = 元の中身[n];
    const b = await 出.z.bytes(n);
    if (a.length !== b.length) { 違.push(n + ' ' + a.length + 'B ⇒ ' + b.length + 'B'); continue; }
    for (let i = 0; i < a.length; i++) {
      if (a[i] !== b[i]) { 違.push(n + ' ' + i + 'バイト目'); break; }
    }
  }
  if (違.length) {
    throw new Error('★' + 違.length + '本 変わりました★' + NL + '       ' + 違.join(NL + '       ')
      + NL + '       ★判子・罫線・塗り・グラフは この 中に 在ります★');
  }
});

await 待('★★仕組みの 3本が ★3つとも★ 直って いる★★'
  + '（★1つでも 抜けると 実Excel は 修復を 言います★）', async () => {
  const wb = await 出.z.text('xl/workbook.xml');
  const rels = await 出.z.text('xl/_rels/workbook.xml.rels');
  const ct = await 出.z.text('[Content_Types].xml');
  if (wb.indexOf('name="Tashita"') < 0) throw new Error('★workbook.xml に 板の 名前が 無い★');
  if (wb.indexOf('r:id="' + 出.足.rId + '"') < 0) throw new Error('★workbook.xml に rId が 無い★');
  if (rels.indexOf('Id="' + 出.足.rId + '"') < 0) throw new Error('★rels に rId が 無い★');
  if (rels.indexOf(出.足.部品.replace('xl/', '')) < 0) throw new Error('★rels に 行き先が 無い★');
  if (ct.indexOf('/' + 出.足.部品) < 0) throw new Error('★[Content_Types].xml に 型が 無い★');
});

await 待('★★番号が ぶつからない（元の 板と 同じ 番号を 使って いない）★★', async () => {
  const wb = await 出.z.text('xl/workbook.xml');
  if (出.足.部品 === 'xl/worksheets/sheet1.xml') throw new Error('★元の 板に 被って います★');
  if (出.足.sheetId === 1) throw new Error('★sheetId が 1★（元の 板と 同じ）');
  /* ★`sheetId` も `rId` も 二度 出て こない★ */
  const 数 = (s, k) => s.split(k).length - 1;
  if (数(wb, 'sheetId="' + 出.足.sheetId + '"') !== 1) {
    throw new Error('★sheetId ' + 出.足.sheetId + ' が 2つ 在ります★');
  }
  if (数(wb, 'r:id="' + 出.足.rId + '"') !== 1) {
    throw new Error('★' + 出.足.rId + ' が 2つ 在ります★');
  }
});

await 待('★★もう 在る 名前は 足せない（黙って 2枚に しない）★★', async () => {
  const book = await E.open(new Uint8Array(中));
  let 投げた = false;
  try { E.板を足す(book, 'Sheet1', {}); } catch (e) { 投げた = true; }
  if (!投げた) throw new Error('★同じ 名前の 板を 足せて しまいました★');
});

await 待('★2枚 続けて 足しても 番号が ぶつからない★', async () => {
  const book = await E.open(new Uint8Array(中));
  const a = E.板を足す(book, 'A', {});
  const b = E.板を足す(book, 'B', {});
  if (a.部品 === b.部品) throw new Error('★同じ 部品名★ ' + a.部品);
  if (a.rId === b.rId) throw new Error('★同じ rId★ ' + a.rId);
  if (a.sheetId === b.sheetId) throw new Error('★同じ sheetId★ ' + a.sheetId);
  const r = await E.save(book);
  const z = Z.read(r.bytes);
  if (!z.has(a.部品) || !z.has(b.部品)) throw new Error('★2枚とも 入って いません★');
});

/* ══ ★★`.xlsb` も 板を 足せる★★ ══（2026-09-22）
     ★司さんの 実物は この 形★。`.xlsx` とは 別の 道（2進）。
     ★★数しか 書けません★★＝字の マスは `sharedStrings.bin` を 指す
     ★実Excel で 数えて もらいました★（経営者1・★赤 0件★）
       `Open` が 投げない ／ 判子 そのまま ／ ★溢れが 3組 とも 生きて いる★ */
const XB = require_(path.join(ROOT, 'lib/xlsb-edit.js'));
const b材料道 = path.join(ROOT, 'tests/fixtures/kazari-hiraku3.xlsb');
const b中 = fs.existsSync(b材料道) ? fs.readFileSync(b材料道) : null;

await 待('★.xlsb の 材料が 在る★', async () => {
  if (!b中) throw new Error('★材料が 無い★ ' + b材料道);
  const h = crypto.createHash('sha256').update(b中).digest('hex');
  if (h !== 'e4154779716440540e7e402d6f2435218b0beea6b477afcfee4e5d6d7828c6e4') {
    throw new Error('★材料が 入れ替わって います★ ' + h);
  }
});

await 待('★★.xlsb に 板を 足すと 元の 部品が 1つも 減らない★★', async () => {
  const z = Z.read(new Uint8Array(b中));
  const 前の名 = z.names().slice();
  const 見本 = XB.parse(await z.bytes('xl/worksheets/sheet1.bin'), XB.SHAPE.sheet).recs;
  const 足 = await XB.xlsb板を足す(z, 見本, 'Tashita', { A1: { v: 123, t: 'n' } });
  const r = await z.build();
  const z2 = Z.read(r.bytes);
  const 無 = 前の名.filter((n) => !z2.has(n));
  if (無.length) throw new Error('★' + 無.length + '本 消えました★ ' + 無.join(' '));
  if (!z2.has(足.部品)) throw new Error('★足した 板が 無い★ ' + 足.部品);
  /* ★★`xl/metadata.bin` は 1バイトも 触らない★★＝★触ると 溢れが 壊れます★ */
  if (z2.has('xl/metadata.bin')) {
    const a2 = await z.bytes('xl/metadata.bin');
    const b2 = await z2.bytes('xl/metadata.bin');
    if (a2.length !== b2.length) throw new Error('★metadata.bin の 大きさが 変わりました★');
  }
});

await 待('★★.xlsb の 板を 名前で 引ける（番号当てでは ない）★★', async () => {
  const z = Z.read(new Uint8Array(b中));
  const 見本 = XB.parse(await z.bytes('xl/worksheets/sheet1.bin'), XB.SHAPE.sheet).recs;
  await XB.xlsb板を足す(z, 見本, 'Tashita', { A1: { v: 123, t: 'n' } });
  const r = await z.build();
  const z2 = Z.read(r.bytes);
  const 板 = XB.板たち(await z2.bytes('xl/workbook.bin'),
    await z2.text('xl/_rels/workbook.bin.rels'));
  if (!板 || 板.length !== 2) throw new Error('★' + (板 ? 板.length : 0) + '枚★');
  if (板[1].名 !== 'Tashita') throw new Error('★2枚目が ' + 板[1].名 + '★');
});

await 待('★★RK＝縮めた 数が 読み戻せる★★（3 / 1 / 0.25 / 123 / 2.5）', async () => {
  const 組 = [3, 1, 0.25, 123, 2.5, -7, 0];
  const RK = (u) => {
    const 百 = (u & 1) !== 0, 整 = (u & 2) !== 0;
    if (整) return (百 ? ((u | 0) >> 2) / 100 : ((u | 0) >> 2));
    const bb = new ArrayBuffer(8), dv = new DataView(bb);
    dv.setUint32(4, (u & 0xFFFFFFFC) >>> 0, true);
    const v = dv.getFloat64(0, true);
    return 百 ? v / 100 : v;
  };
  for (const n of 組) {
    const 出 = RK(XB.RKにする(n));
    if (Math.abs(出 - n) > 1e-9) throw new Error('★' + n + ' ⇒ ' + 出 + '★');
  }
});

console.log('');
console.log('  ★見て いない 事★');
console.log('    ・★実Excel が 開くかは 測れません★（COM が 要る＝経営者1 の 持ち場）');
console.log('    ・★お客さんの 道（画面）は 別の 道具★');
console.log('      ＝`docs/measured/hakaru-tashita-ita-ga-hozon-sareru-ka.mjs`');
console.log('    ・★足した 板に 飾りは 書きません★＝値と 式だけ');
console.log('    ・★`xl/calcChain.xml` は 触りません★＝★消して 良いか 測って いません★');
console.log('ita-wo-tasu: ' + pass + ' 緑 / ' + fail + ' 赤');

if (process.argv.includes('--self-test')) {
  console.log('');
  console.log('--self-test: ★わざと壊して 赤に なるか★（★lib は 1字も 触りません★）');
  let 悪 = 0;
  const 見る = (な, ok) => { if (ok) console.log('  ok   ' + な); else { console.log('  NG   ' + な); 悪++; } };

  /* ㋐ ★足さずに 保存したら 部品は 増えない★＝門①が 本当に 数えて いるか */
  const b1 = await E.open(new Uint8Array(中));
  const r1 = await E.save(b1);
  見る('★足さなければ 部品は 増えない★', Z.read(r1.bytes).names().length === 元の名.length);

  /* ㋑ ★型を 入れずに 保存したら 門④が 赤に なる形か★ */
  const b2 = await E.open(new Uint8Array(中));
  const 足2 = E.板を足す(b2, 'X', {});
  b2.addedTypes = [];                       /* ★型を わざと 捨てる★ */
  const r2 = await E.save(b2);
  const ct2 = await Z.read(r2.bytes).text('[Content_Types].xml');
  見る('★型を 捨てると [Content_Types].xml に 入らない（門が 気づける 形）★',
    ct2.indexOf('/' + 足2.部品) < 0);

  /* ㋒ ★板の 中身を 変えたら 部品の 中身も 変わる★＝門①が 中を 見て いるか */
  const b3 = await E.open(new Uint8Array(中));
  const 足3 = E.板を足す(b3, 'Y', { A1: { v: 999, t: 'n' } });
  const r3 = await E.save(b3);
  const t3 = await Z.read(r3.bytes).text(足3.部品);
  見る('★中身を 変えると 部品の 中も 変わる（門が 中を 見て います）★',
    t3.indexOf('<v>999</v>') >= 0);

  process.exit(悪 ? 1 : 0);
}
process.exit(fail ? 1 : 0);
