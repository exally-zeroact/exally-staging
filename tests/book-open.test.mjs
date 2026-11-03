/* book-open.test.mjs — ★受け取ったブックを「何も変えずに保存」しても壊れないこと★
 *
 * なぜ必要か（2026-08-09・指示役が実機で発見）:
 *   実物(.xlsb)を開いて ★何も編集せずに保存を押しただけで★
 *     「保存しませんでした：この数式セルの答えの形は まだ直せません（記録 8）」
 *   になった。原因は ★見た目の違いを「変わった」と数えていた★こと:
 *     開いた時の控え = ファイルの表示文字 "1,000" ／ 計算し直した後 = 生の数 1000
 *   → 1つも触っていないのに数式セルが全部「変わった」ことになり、
 *     文字を返す数式セル(BrtFmlaString)に当たって保存そのものが断られた。
 *   ★司さんの実物には BrtFmlaString が 2,866個＝必ず当たる。★
 *
 *   ★「いちばん普通の操作（開いて、何も変えずに保存）」がテストに無かった。★
 *   だから今まで1度も出なかった。ここで固定する。
 *
 * 使い方: node tests/book-open.test.mjs
 */
import { createRequire } from 'node:module';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const require = createRequire(import.meta.url);
const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const XLSX = require(path.join(ROOT, 'lib/xlsx.full.min.js'));
const ZipSurgeon = require(path.join(ROOT, 'lib/zip-surgeon.js'));
const XlsxEdit = require(path.join(ROOT, 'lib/xlsx-edit.js'));
const XlsbEdit = require(path.join(ROOT, 'lib/xlsb-edit.js'));
const XlsbJitai = require(path.join(ROOT, 'lib/xlsb-jitai.js'));

/* book-open.js はブラウザの物なので、必要な物だけ載せた入れ物で読み込む */
const box = { XLSX, ZipSurgeon, XlsxEdit, XlsbEdit };
const src = fs.readFileSync(path.join(ROOT, 'js/book-open.js'), 'utf8');
new Function('self', 'window', src + '\n;self.__BookOpen = self.BookOpen;')(box, box);
const BookOpen = box.__BookOpen;

const FIX = path.join(ROOT, 'tests/fixtures/book-open-sample.xlsb');
const bytes = new Uint8Array(fs.readFileSync(FIX));
const fakeFile = (name, b) => ({ name: name, arrayBuffer: () => Promise.resolve(b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength)) });

let pass = 0, fail = 0;
const T = async (n, fn) => { try { await fn(); pass++; console.log('  ✓ ' + n); } catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + (e && e.message)); } };
const eq = (a, b, m) => { if (a !== b) throw new Error(`${m}: ${a} ≠ ${b}`); };

console.log('\n[book-open] 受け取ったブックを開いて保存する道');

/* 材料が本当に「文字を返す式」を持っているか（★空振りしていないことの確認★） */
await T('★見本に「文字を返す式」が入っている（無いとこの検査は意味が無い）', async () => {
  const z = ZipSurgeon.read(bytes);
  let str = 0, num = 0;
  for (const n of z.names().filter((x) => /worksheets\/sheet\d+\.bin$/.test(x))) {
    const p = XlsbEdit.parse(await z.bytes(n), XlsbEdit.SHAPE.sheet);
    p.recs.forEach((r) => { if (r.id === XlsbEdit.R.FMLA_STRING) str++; if (r.id === XlsbEdit.R.FMLA_NUM) num++; });
  }
  if (str < 1) throw new Error('BrtFmlaString が0個');
  if (num < 1) throw new Error('BrtFmlaNum が0個');
  console.log(`      （BrtFmlaString ${str}個 / BrtFmlaNum ${num}個）`);
});

/* ★本命：開いて、何も変えずに保存★ */
await T('★開いて 何も変えずに保存 → 1バイトも変わらない（作り直さない）', async () => {
  const r = await BookOpen.openFile(fakeFile('sample.xlsb', bytes));
  eq(r.kind, 'xlsb', '形式');
  const res = await BookOpen.saveOpened(r.sheets);
  const out = res.bytes || res;
  eq(res.log && res.log.noChange, true, '変更なしの印');
  eq(Buffer.compare(Buffer.from(out), Buffer.from(bytes)), 0, '中身');
});

/* ★見た目だけ違う（"1,000" と 1000）を「変わった」と数えない★ */
await T('★"1,000" と 1000 を同じと見る（見た目の差で全セルが変わった事にしない）', () => {
  eq(BookOpen.normForCompare('1,000'), BookOpen.normForCompare(1000), '桁区切り');
  eq(BookOpen.normForCompare('¥1,234'), BookOpen.normForCompare(1234), '通貨記号');
  eq(BookOpen.normForCompare(' 5 '), BookOpen.normForCompare(5), '前後の空白');
  if (BookOpen.normForCompare('商品1') === BookOpen.normForCompare('商品2')) throw new Error('文字まで同じにした');
});

/* ★1つだけ変えたら、書き換わる記録も1つだけ★ */
await T('★1セルだけ変えたら、変わる記録も1つだけ', async () => {
  const r = await BookOpen.openFile(fakeFile('sample.xlsb', bytes));
  const sh = r.sheets[0];
  sh.data['1,1'].v = 99; sh.data['1,1'].d = 99;          // B2（数量）を99に
  eq(Object.keys(BookOpen.changedCells(sh)).length, 1, '変わったセルの数');
  const res = await BookOpen.saveOpened(r.sheets);
  const out = res.bytes || res;
  const a = ZipSurgeon.read(bytes), b = ZipSurgeon.read(out);
  const changed = a.entries.filter((e) => !b.index[e.name] || b.index[e.name].crc !== e.crc).map((e) => e.name);
  // シート1本 ＋ binaryIndex を外したぶん（消えた物は changed に出る）
  const sheetsChanged = changed.filter((n) => /worksheets\/sheet\d+\.bin$/.test(n));
  eq(sheetsChanged.length, 1, '中身が変わったシートの数');
  // 記録の中身も1つだけか
  const before = XlsbEdit.parse(await a.bytes(sheetsChanged[0]), XlsbEdit.SHAPE.sheet).recs;
  const after = XlsbEdit.parse(await b.bytes(sheetsChanged[0]), XlsbEdit.SHAPE.sheet).recs;
  eq(before.length, after.length, '記録の本数');
  let diff = 0;
  for (let i = 0; i < before.length; i++) {
    if (before[i].id !== after[i].id || Buffer.compare(Buffer.from(before[i].data), Buffer.from(after[i].data)) !== 0) diff++;
  }
  eq(diff, 1, '★中身が違う記録の数★');
});

/* ══ ★★決めを 変えました（2026-09-25）★★ ══
     ★★前の 決め★★
       「答えが 字の 式（記録8）を 書き換えようと したら ★本 1冊 まるごと 断る★」
       ＝★壊すより 断る★ という 考えで 入れて いました。
     ★★何が 起きて いたか（経営者1 が お客さんの 道で 押しました・司さんの 実物）★★
       ★入力値を 1つ 直して 「書き出す」を 押すと 1冊も 出ません★
       帯 ･･･「書き出しませんでした／この数式セルの答えの形は まだ直せません（記録 8）」
       ⇒★本番と 同じ 木（15cc377）でも 1字も 違わず 出ません＝元から です★
       ⇒★★＝お客さんは 「直して 保存する」が 出来ません★★
       ⇒司さんの 決め ア「★全部 保存しろや、断る 理由が なんか あるんか★」に 直に 当たります
     ★★新しい 決め★★
       ⑴★答えが 字の 式は 触らない（書き換えない）★＝★壊しません★
       ⑵★でも 本は 出す★＝★お客さんが 打った マスは ちゃんと 書きます★
       ⑶★触れなかった 物が 在れば `workbook.bin` に 「開いたら 全部 計算しろ」の 印を 立てる★
          ＝★実Excel が 開いた 時に 計算し直す ので 古い 答えは 直ります★
     ★★だから この 見張りが 守る 物も 変わります★★
       前 ･･･「★断るか★」／今 ･･･「★出すか／触らないか／印を 立てるか★」 */
await T('★★答えが 字の 式が 在っても 本は 出す★★（★前は 1冊 まるごと 断って いた★）', async () => {
  const r = await BookOpen.openFile(fakeFile('sample.xlsb', bytes));
  const sh = r.sheets[0];
  sh.data['1,4'] = sh.data['1,4'] || { v: '', f: '', d: '' };
  sh.data['1,4'].d = 123456;                              // E2（=TEXT(...)）の答えを数に変える
  const res = await BookOpen.saveOpened(r.sheets);
  const out = res && res.bytes ? res.bytes : res;
  if (!out || !out.length) throw new Error('本が 出て いません');
  /* ★触らなかった 事を 見ます★＝★その 記録の 中身が 1バイトも 変わって いない★ */
  const a2 = ZipSurgeon.read(bytes), b2 = ZipSurgeon.read(out);
  const 板名 = a2.entries.map((e) => e.name).filter((n) => /worksheets[/]sheet\d+[.]bin$/.test(n))[0];
  const 前 = XlsbEdit.parse(await a2.bytes(板名), XlsbEdit.SHAPE.sheet).recs;
  const 後 = XlsbEdit.parse(await b2.bytes(板名), XlsbEdit.SHAPE.sheet).recs;
  let 字の式が変わった = 0;
  for (let i = 0; i < 前.length; i++) {
    if (前[i].id !== XlsbEdit.R.FMLA_STRING) continue;
    if (後[i].id !== 前[i].id
      || Buffer.compare(Buffer.from(前[i].data), Buffer.from(後[i].data)) !== 0) 字の式が変わった++;
  }
  eq(字の式が変わった, 0, '★答えが 字の 式を 触った 数★');
  /* ★触れなかった 物が 在る ので 印が 立って いる はず★ */
  const wb = a2.entries.map((e) => e.name).filter((n) => /^xl[/]workbook[.]bin$/.test(n))[0];
  if (!wb) throw new Error('workbook.bin が 在りません＝この 検査は 空振り');
  const w後 = XlsbEdit.parse(await b2.bytes(wb), XlsbEdit.SHAPE.workbook);
  if (!w後.ok) throw new Error('workbook.bin を 歩けません＝この 検査は 空振り');
  eq(XlsbJitai.開いたら全部計算するか(w後.recs), true, '★開いたら 全部 計算しろ の 印★');
});

/* ★.xlsx でも同じ穴が無いか（形式ごとに道が違うので、形式ごとに確かめる）★ */
const FIX_X = path.join(ROOT, 'tests/fixtures/book-open-sample.xlsx');
const bytesX = new Uint8Array(fs.readFileSync(FIX_X));
await T('★.xlsx も 何も変えずに保存 → 1バイトも変わらない', async () => {
  const r = await BookOpen.openFile(fakeFile('sample.xlsx', bytesX));
  eq(r.kind, 'xlsx', '形式');
  const res = await BookOpen.saveOpened(r.sheets);
  const out = res.bytes || res;
  eq(res.log && res.log.noChange, true, '変更なしの印');
  eq(Buffer.compare(Buffer.from(out), Buffer.from(bytesX)), 0, '中身');
});
await T('★.xlsx で1セル変えたら、変わる部品は狙った物だけ', async () => {
  const r = await BookOpen.openFile(fakeFile('sample.xlsx', bytesX));
  const sh = r.sheets[0];
  sh.data['1,1'].v = 77; sh.data['1,1'].d = 77;
  eq(Object.keys(BookOpen.changedCells(sh)).length, 1, '変わったセルの数');
  const res = await BookOpen.saveOpened(r.sheets);
  const a = ZipSurgeon.read(bytesX), b = ZipSurgeon.read(res.bytes || res);
  const changed = a.entries.filter((e) => !b.index[e.name] || b.index[e.name].crc !== e.crc).map((e) => e.name);
  const unexpected = changed.filter((n) => !/worksheets\/sheet\d+\.xml$|workbook\.xml$|sharedStrings\.xml$/.test(n));
  if (unexpected.length) throw new Error('狙っていない部品が変わった: ' + unexpected.join(', '));
  if (!changed.length) throw new Error('1つも変わっていない');
});

console.log('\n' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
