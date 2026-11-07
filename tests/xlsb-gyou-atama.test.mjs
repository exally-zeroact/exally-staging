/* xlsb-gyou-atama.test.mjs — ★.xlsb の 新しい 行の 頭を 0 から 作るか★（2026-11-07）
 *
 *  ★★何が 起きて いたか★★（実Excel で 測りました）
 *    `マスを足す`（空いた 行に 打つ）と `xlsb板を作る` が ★最初の 行の 頭を 丸写し★ して いた
 *      ⇒最初の 行が 隠れて いる・手で 高さを 決めた・段の組 に 入って いる 本では
 *        ★打った 行も 隠れた★（行5 h=0 hid=True lv=3）／xlsb板を作る は ★全部の 行が 隠れた★
 *      ⇒最初の 行が 33バイト（1024列を 越える）の 本では ★「見本が 在りません」で 断った★
 *
 *  ★★この 見張りが 見る 物★★
 *    ⑴ 新しい 行の 頭＝書式 0・印 0（byte10/11/12）・高さ＝板の 既定（記録485 u16@6＝353）・範囲＝マスの 列
 *    ⑵ 狭い 本（25バイト）と 広い 本（33バイト）の 両方で 足せる
 *    ⑶ xlsb板を作る が 記録485 を 134 の 後・145 の 前 に 写す／全部の 行の 頭が ⑴ の 形
 *    ⑷ 借り物（SheetJS）で 読み直して 打った 数が 在る
 *    ⑸ ★壊し★＝丸写し・300 固定・印の 無い 見本の 行・範囲を 見本のまま を この 見張りが ★赤★ に する
 *    ★実Excel での 見え方（h=17.7 hid=False lv=1）は tools/atama-mihon-excel.ps1 -Mode hakaru で 測る★（CI では 測れない）
 *
 *  ★見本★ tests/fixtures/gyou-atama-2026-11-07.xlsb（★実Excel に 作らせた★＝1行目 隠す/手/段3・25バイト）
 *          tests/fixtures/gyou-atama-hiroi-2026-11-07.xlsb（3行目の 1501列目にも 数＝行の頭 33バイト）
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(path.join(ROOT, 'package.json'));
const XE = require_(path.join(ROOT, 'lib/xlsb-edit.js'));
const X = require_(path.join(ROOT, 'lib/xlsx.full.min.js'));
const 狭い = path.join(ROOT, 'tests/fixtures/gyou-atama-2026-11-07.xlsb');
const 広い = path.join(ROOT, 'tests/fixtures/gyou-atama-hiroi-2026-11-07.xlsb');

let pass = 0, fail = 0;
const T = (n, f) => {
  try { f(); pass++; console.log('  ✓ ' + n); }
  catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + e.message); }
};
const ok = (c, m) => { if (!c) throw new Error(m); };

console.log('');
console.log('[xlsb-gyou-atama] ★.xlsb の 新しい 行の 頭を 0 から 作るか★');

const 板を読む = (file) => {
  const cfb = X.CFB.read(fs.readFileSync(file), { type: 'buffer' });
  const i = cfb.FullPaths.findIndex((p) => /sheet1\.bin$/.test(p));
  return new Uint8Array(cfb.FileIndex[i].content);
};
const 本に戻す = (file, bytes) => {
  const cfb = X.CFB.read(fs.readFileSync(file), { type: 'buffer' });
  const i = cfb.FullPaths.findIndex((p) => /sheet1\.bin$/.test(p));
  cfb.FileIndex[i].content = Buffer.from(bytes);
  cfb.FileIndex[i].size = bytes.length;
  return X.CFB.write(cfb, { fileType: 'zip', type: 'buffer' });
};
const u32 = (d, at) => (d[at] | (d[at + 1] << 8) | (d[at + 2] << 16) | (d[at + 3] << 24)) >>> 0;
const 行の頭 = (recs, r) => (recs.find((x) => x.id === 0 && u32(x.data, 0) === r) || {}).data;

/* ★新しい 行の 頭の 形を 見る★＝★壊しにも 同じ 物を 当てる★（⑸） */
const 形を見る = (d, r, 列ら, 高さ) => {
  ok(d, '行 ' + r + ' の 頭が 無い');
  const 区 = {};
  列ら.forEach((c) => { const k = Math.floor(c / 1024); const z = 区[k] || (区[k] = [c, c]); z[0] = Math.min(z[0], c); z[1] = Math.max(z[1], c); });
  const ks = Object.keys(区).map(Number).sort((a, b) => a - b);
  ok(d.length === 17 + 8 * ks.length, '長さ ' + d.length + '（期待 ' + (17 + 8 * ks.length) + '）');
  ok(u32(d, 0) === r, '行番号 ' + u32(d, 0));
  ok(u32(d, 4) === 0, '書式 ' + u32(d, 4) + '（期待 0）');
  ok((d[8] | (d[9] << 8)) === 高さ, '高さ ' + (d[8] | (d[9] << 8)) + '（期待 ' + 高さ + '）');
  ok(d[10] === 0 && d[11] === 0 && d[12] === 0, '印 ' + [d[10], d[11], d[12]].map((b) => '0x' + b.toString(16)).join(' ') + '（期待 全部 0）');
  ok(u32(d, 13) === ks.length, '範囲の数 ' + u32(d, 13));
  ks.forEach((k, n) => {
    ok(u32(d, 17 + 8 * n) === 区[k][0] && u32(d, 21 + 8 * n) === 区[k][1],
      '範囲' + n + ' ' + u32(d, 17 + 8 * n) + '-' + u32(d, 21 + 8 * n) + '（期待 ' + 区[k].join('-') + '）');
  });
};

const 狭 = 板を読む(狭い), 広 = 板を読む(広い);
const 狭recs = XE.parse(狭, XE.SHAPE.sheet).recs;

T('⑴ 板の 既定の 高さ＝記録485 u16@6＝353／485 が 無ければ 300', () => {
  ok(XE.既定の高さ(狭recs) === 353, '狭い 本 ' + XE.既定の高さ(狭recs));
  ok(XE.既定の高さ(狭recs.filter((x) => x.id !== 485)) === 300, '485 抜き ' + XE.既定の高さ(狭recs.filter((x) => x.id !== 485)));
});

T('⑴ 行の頭を作る＝1列 25バイト／1024列を 越えると 33バイト／17区切りは 投げる', () => {
  形を見る(XE.行の頭を作る(4, [2], 353), 4, [2], 353);
  形を見る(XE.行の頭を作る(6, [5, 1499, 0], 353), 6, [0, 5, 1499], 353);
  let 投げた = false;
  try { XE.行の頭を作る(0, Array.from({ length: 17 }, (_, i) => i * 1024), 300); } catch (e) { 投げた = true; }
  ok(投げた, '17区切りで 投げない');
});

const 打つ = { '4,2': 5, '6,1499': 7 };
for (const [名, bin, file] of [['狭い 本（25バイト）', 狭, 狭い], ['広い 本（33バイト）', 広, 広い]]) {
  T('⑵ マスを足す：' + 名 + ' の 空いた 行5・行7 に 打つ＝新しい 行の 頭が 0 から', () => {
    const a = XE.editSheet(bin, 打つ);
    ok(a.ok, '断った：' + a.why);
    const recs = XE.parse(a.bytes, XE.SHAPE.sheet).recs;
    形を見る(行の頭(recs, 4), 4, [2], 353);
    形を見る(行の頭(recs, 6), 6, [1499], 353);
    /* ★元から 在る 1行目の 頭は 触らない★ */
    const 元 = 行の頭(XE.parse(bin, XE.SHAPE.sheet).recs, 0), 今 = 行の頭(recs, 0);
    ok(Buffer.from(元).equals(Buffer.from(今)), '1行目の 頭が 変わった');
    /* ⑷ 借り物で 読み直す */
    const wb = X.read(本に戻す(file, a.bytes), { type: 'buffer' });
    const ws = wb.Sheets[wb.SheetNames[0]];
    ok(ws.C5 && ws.C5.v === 5, 'C5 ' + JSON.stringify(ws.C5));
    ok(ws[X.utils.encode_cell({ r: 6, c: 1499 })] && ws[X.utils.encode_cell({ r: 6, c: 1499 })].v === 7, '行7 列1500 が 無い');
  });
}

T('⑶ xlsb板を作る：記録485 を 134 の 後・145 の 前 に 写す（段の 深さは 0）／全部の 行の 頭が 0 から', () => {
  const b = XE.xlsb板を作る(狭recs, { A1: { v: 1 }, B3: { v: 2 }, C5: { v: 5 }, F7: { v: 7 }, A7: { v: 8 } });
  const recs = XE.parse(b, XE.SHAPE.sheet).recs;
  const ids = recs.map((x) => x.id);
  const i485 = ids.indexOf(485);
  ok(i485 > ids.indexOf(134) && i485 < ids.indexOf(145), '485 の 所 ' + i485 + '（134=' + ids.indexOf(134) + ' 145=' + ids.indexOf(145) + '）');
  const 元485 = 狭recs.find((x) => x.id === 485).data, 今485 = recs[i485].data;
  ok(元485[10] === 2, '見本の 段の 深さが 2 で ない（' + 元485[10] + '）＝壊しが 空振り');
  ok(今485[10] === 0 && 今485[11] === 0, '段の 深さを 写した ' + 今485[10] + '/' + 今485[11] + '（期待 0/0）');
  ok(Buffer.from(今485.subarray(0, 10)).equals(Buffer.from(元485.subarray(0, 10))), '485 の 高さ 等が 写って いない');
  形を見る(行の頭(recs, 0), 0, [0], 353);
  形を見る(行の頭(recs, 2), 2, [1], 353);
  形を見る(行の頭(recs, 4), 4, [2], 353);
  形を見る(行の頭(recs, 6), 6, [0, 5], 353);
  const wb = X.read(本に戻す(狭い, b), { type: 'buffer' });
  const ws = wb.Sheets[wb.SheetNames[0]];
  ok(ws.F7 && ws.F7.v === 7 && ws.A7 && ws.A7.v === 8, 'F7/A7 ' + JSON.stringify([ws.F7, ws.A7]));
});

/* ⑸ ★壊し★＝前の 形・取り違えの 形を 作って ★形を見る が 赤に する★ か */
const 赤か = (名, d, r, 列ら) => T('⑸ 壊し「' + 名 + '」は 赤', () => {
  let 赤 = false;
  try { 形を見る(d, r, 列ら, 353); } catch (e) { 赤 = true; }
  ok(赤, '★壊しを 見逃した★');
});
const 丸写し = (src, r) => { const d = Uint8Array.from(src); d[0] = r & 0xff; d[1] = 0; d[2] = 0; d[3] = 0; return d; };
赤か('最初の 行の 頭を 丸写し（前の 形）', 丸写し(行の頭(狭recs, 0), 4), 4, [2]);
赤か('高さ 300 固定', XE.行の頭を作る(4, [2], 300), 4, [2]);
赤か('印の 無い 見本の 行（2行目）を 写す', 丸写し(行の頭(狭recs, 1), 4), 4, [2]);
const 範囲そのまま = XE.行の頭を作る(4, [2], 353); 範囲そのまま.set([0, 0, 0, 0, 0, 0, 0, 0], 17);
赤か('範囲を 見本のまま（0-0）', 範囲そのまま, 4, [2]);

console.log('');
console.log('  [xlsb-gyou-atama] ' + pass + ' 通過 / ' + fail + ' 失敗');
if (fail) process.exit(1);
