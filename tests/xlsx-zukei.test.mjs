/* xlsx-zukei.test.mjs — ★図形（判子）を 自前で 読めて いるか★（2026-09-21）
 *
 *  ★★なぜ 要るか★★
 *    実Excel が 作った ファイルを お客さんの 道で 開いて 数えたら
 *    ★図形が 画面に 1つも 出て いません★でした
 *    （`docs/measured/golden-kazari-gamen-made-2026-09-21.tsv`）。
 *    因 ＝ ★`xl/drawings/drawing1.xml` を 読む 所が 1つも 在りません★。
 *
 *  ★★物差し★★
 *    ①`tests/fixtures/kazari-hiraku3.xlsx`（11639B・実Excel 16.0 build 20326 が COM で 作った 物。10-10 に 作者の 欄・10-11 に 保存した PC の 道を 空に した＝元 11725B）
 *    ②実Excel が その ファイルを 開いて 言った 数（経営者1 の 実測・2026-09-21）
 *       `docs/measured/golden-moto-no-katachi-excel-2026-09-21.tsv`
 *         図形 Left ★449.375★ ／ Top ★20★ ／ Width ★60★ ／ Height ★60★（ポイント）
 *         図形の 名 ★hanko★ ／ AutoShapeType ★1（四角）★
 *
 *  ★★点（px）と ポイントの 直し方★★
 *    実Excel は ★ポイント★で 言います。画面は ★96dpi の 点★です。
 *    ⇒ 1pt ＝ 4/3 点 ⇒ 60pt ＝ ★80点★ ／ 20pt ＝ ★26.67点★ ／ 449.375pt ＝ ★599.17点★
 *
 *  ★★★位置は `xfrm` を 使いません★★★
 *    この ファイルの `xfrm` は `x=4064000 EMU`＝★320pt★ ですが、
 *    実Excel は ★449.375pt★ と 言います。
 *    ⇒★`xfrm` は 置いた 時の 古い 数★（`Placement 1` ＋ A列を 30 に 広げた）
 *    ⇒★`twoCellAnchor` の マスと ずれ から 出します★
 *    ⇒★往復の 穴では ありません★（元の ファイルでも 449.375・経営者1 が 09-21 に 実測）
 *
 *  使い方: node tests/xlsx-zukei.test.mjs
 *          node tests/xlsx-zukei.test.mjs --self-test
 */
import fs from 'node:fs';
import path from 'node:path';
import zlib from 'node:zlib';
import crypto from 'node:crypto';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(path.join(ROOT, 'package.json'));
const Z = require_(path.join(ROOT, 'lib/xlsx-zukei.js'));

let pass = 0, fail = 0;
const NL = String.fromCharCode(10);
const T = (n, f) => {
  try { f(); pass++; console.log('  ok   ' + n); }
  catch (e) { fail++; console.log('  NG   ' + n + NL + '       ' + e.message); }
};

/* ★★2進の まま ほどく★★（`.bin` は 字に すると 壊れます）
     ＝2026-09-21 に ここで 1回 踏みました＝`workbook.bin` を 字で 渡して
       `板たち()` が `null` を 返し ★門が 4本 赤★に なりました。
     ＝★本番の 穴では なく 門の 穴★でした。 */
function ほどく生(buf) {
  const 出 = {};
  let p = 0;
  while (p + 30 <= buf.length) {
    if (buf.readUInt32LE(p) !== 0x04034b50) break;
    const 方 = buf.readUInt16LE(p + 8);
    const 圧 = buf.readUInt32LE(p + 18);
    const n = buf.readUInt16LE(p + 26);
    const x = buf.readUInt16LE(p + 28);
    const 名 = buf.toString('utf8', p + 30, p + 30 + n);
    const 頭 = p + 30 + n + x;
    const 体 = buf.slice(頭, 頭 + 圧);
    出[名] = 方 === 8 ? zlib.inflateRawSync(体) : 体;
    p = 頭 + 圧;
  }
  return 出;
}

function ほどく(buf) {
  const 出 = {};
  let p = 0;
  while (p + 30 <= buf.length) {
    if (buf.readUInt32LE(p) !== 0x04034b50) break;
    const 方 = buf.readUInt16LE(p + 8);
    const 圧 = buf.readUInt32LE(p + 18);
    const n = buf.readUInt16LE(p + 26);
    const x = buf.readUInt16LE(p + 28);
    const 名 = buf.toString('utf8', p + 30, p + 30 + n);
    const 頭 = p + 30 + n + x;
    const 体 = buf.slice(頭, 頭 + 圧);
    出[名] = 方 === 8 ? zlib.inflateRawSync(体).toString('utf8') : 体.toString('utf8');
    p = 頭 + 圧;
  }
  return 出;
}

const 材料道 = path.join(ROOT, 'tests/fixtures/kazari-hiraku3.xlsx');
const 中 = fs.existsSync(材料道) ? fs.readFileSync(材料道) : null;

console.log('');
console.log('[xlsx-zukei] ★図形（判子）を 自前で 読めて いるか★');

T('★材料が 在る（空振りして いない）★', () => {
  if (!中) throw new Error('★材料が 無い★ ' + 材料道);
  const h = crypto.createHash('sha256').update(中).digest('hex');
  if (h !== 'affd42e9f5c6de9d5565c8ffd623ffd2e37a7ba5ca8820fdd5c971b15ff93053') {
    throw new Error('★材料が 入れ替わって います★ sha256=' + h);
  }
});

const 部品 = 中 ? ほどく(中) : {};

/* ★板 → 図形の 部品名★（★並び順では 当てません★＝rels で 解きます） */
const 部品名 = 中 ? Z.図形の部品名(
  部品['xl/worksheets/sheet1.xml'],
  部品['xl/worksheets/_rels/sheet1.xml.rels'],
  'xl/worksheets/sheet1.xml') : null;

T('★★rels から 図形の 部品名を 解けて いる★★（★並び順で 当てない★）', () => {
  if (部品名 !== 'xl/drawings/drawing1.xml') {
    throw new Error('★部品名が ' + 部品名 + '★（xl/drawings/drawing1.xml の はず）');
  }
});

const 図たち = 部品名 && 部品[部品名] ? Z.図形を読む(部品[部品名]) : [];
console.log('      ＝ 読んだ 図形 ' + 図たち.length + '個 ／ ' + JSON.stringify(図たち));

T('★図形が 1つ 読めた（名前は hanko）★', () => {
  if (図たち.length !== 1) throw new Error('★' + 図たち.length + '個★（1個 の はず）');
  if (図たち[0].名 !== 'hanko') throw new Error('★名が ' + 図たち[0].名 + '★（hanko の はず）');
});

T('★形が 四角（prst="rect"）★', () => {
  if (図たち[0].種類 !== '四角') throw new Error('★種類が ' + 図たち[0].種類 + '★（四角 の はず）');
});

/* ★★実Excel の 数と 突き合わせます★★（★1点まで 見ます★） */
const 点 = (pt) => pt * 4 / 3;
T('★★大きさが 実Excel と 合う（60pt × 60pt ＝ 80点 × 80点）★★', () => {
  if (図たち[0].w !== 80) throw new Error('★幅が ' + 図たち[0].w + '★（80 の はず）');
  if (図たち[0].h !== 80) throw new Error('★高さが ' + 図たち[0].h + '★（80 の はず）');
});

/* ★★場所★★
     ★うちの 列幅で 解きます★＝`colW` は ★画面が 実際に 使う 数★を 渡します。
     ★2026-09-21 に 本番の 道で 測った 数★ ＝ A列 ＝ 245点／他 ＝ 72点（標準）
     ⇒ この 数で 解くと 実Excel の 449.375pt（＝599.17点）に なる はず。 */
const 場 = Z.場所を決める(図たち[0] || {}, { 0: 245 }, 72, {}, 24);
console.log('      ＝ 場所 x=' + 場.x + ' y=' + 場.y
  + '（実Excel ＝ x ' + 点(449.375).toFixed(2) + ' / y ' + 点(20).toFixed(2) + '）');

T('★★横の 場所が 実Excel と 1点 以内★★（★`xfrm` の 320pt では ありません★）', () => {
  const 待 = 点(449.375);
  const 差 = Math.abs(場.x - 待);
  if (差 > 1) {
    throw new Error('★x が ' + 場.x + '★（' + 待.toFixed(2) + ' の はず・差 ' + 差.toFixed(2) + '）'
      + '／★`xfrm` を 使うと ' + 点(320).toFixed(2) + ' に なります★');
  }
});

T('★★縦の 場所が 実Excel と 1点 以内★★', () => {
  const 待 = 点(20);
  const 差 = Math.abs(場.y - 待);
  if (差 > 1) throw new Error('★y が ' + 場.y + '★（' + 待.toFixed(2) + ' の はず）');
});

T('★★`xfrm` の 古い 数（320pt）に なって いない★★（★これが 一番 起きやすい 間違い★）', () => {
  if (Math.abs(場.x - 点(320)) < 1) {
    throw new Error('★x が ' + 場.x + ' ＝ `xfrm` の 数です★'
      + '／★マスと ずれ（twoCellAnchor）から 出して ください★');
  }
});

T('★台に 載せる 形（`sheets[i].objects` と 同じ）が 出る★', () => {
  const o = Z.台に載せる形(図たち, { 0: 245 }, 72, {}, 24);
  if (o.length !== 1) throw new Error('★' + o.length + '個★');
  for (const k of ['x', 'y', 'w', 'h', 'z', '名']) {
    if (o[0][k] === undefined) throw new Error('★' + k + ' が 無い★ ' + JSON.stringify(o[0]));
  }
  if (o[0].種類 !== '四角') throw new Error('★種類が ' + o[0].種類 + '★');
});

/* ★★色は 付けません★★（★当て推量で 色を 作らない★） */
/* ══ ★★テーマの 色★★ ══（2026-09-21）
     この 判子は `<a:fillRef idx="1"><a:schemeClr val="accent1"/>` ＝ ★テーマの 色★
     ＝実Excel に 聞くと 塗り ＝ COM 8544277 ＝ ★#156082★（経営者1・09-21）
     ＝`xl/theme/theme1.xml` の accent1 と 同じ
     ★★テーマを 渡さない 時は 付けません★★＝★当て推量で 色を 作らない★ */
const KZ = require_(path.join(ROOT, 'lib/xlsx-kazari.js'));
const テーマ = KZ.テーマを読む(部品['xl/theme/theme1.xml']);

T('★テーマを 渡さない 時は 色を 付けない★', () => {
  const 無 = Z.図形を読む(部品[部品名]);
  if (無[0] && 無[0].塗り) throw new Error('★塗りが ' + 無[0].塗り + ' と 付いて います★');
});

T('★★濃さ（shade）＝光の 量で 0.15倍 が ぴったり★★', () => {
  const 出 = Z.暗くする('#156082', 0.15);
  if (出 !== '#042433') {
    throw new Error('★' + 出 + '★（#042433 の はず）'
      + '／★実Excel の COM 3351556 と 同じ 色です★'
      + '／★字の 値を そのまま 0.15倍 すると #030E14 に なります（外れた 見立て）★');
  }
});

T('★★判子の 線の 色が 実Excel と 合う★★（#042433）', () => {
  const 有 = Z.図形を読む(部品[部品名], テーマ);
  if (!有.length) throw new Error('★図形が 0個★');
  if (有[0].線 !== '#042433') {
    throw new Error('★線が ' + 有[0].線 + '★（#042433 の はず）');
  }
});

T('★★テーマを 渡すと 判子の 塗りが 実Excel と 合う★★（#156082）', () => {
  const 有 = Z.図形を読む(部品[部品名], テーマ);
  if (!有.length) throw new Error('★図形が 0個★');
  if (有[0].塗り !== '#156082') {
    throw new Error('★塗りが ' + 有[0].塗り + '★（#156082 の はず）'
      + '／★実Excel の COM 8544277 と 同じ 色です★');
  }
});

T('★★`lib/xlsx-kazari.js` と テーマの 並びが 同じ★★（★2か所に 書いて います★）', () => {
  const a = (KZ.テーマの並び || []).join(',');
  const b = (Z.テーマの並び || []).join(',');
  if (!a || a !== b) {
    throw new Error('★並びが 違います★'
      + '／kazari ' + a + '／zukei ' + b
      + '／★片方を 直したら もう片方も 直して ください★');
  }
});
T('★字で 書いた 色（srgbClr）なら 使う★', () => {
  const 作 = 部品[部品名].split('<a:prstGeom').join(
    '<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:prstGeom');
  const 出 = Z.図形を読む(作);
  if (!出.length || 出[0].塗り !== '#FF0000') {
    throw new Error('★塗りが ' + (出[0] || {}).塗り + '★（#FF0000 の はず）');
  }
});

/* ══ ★★`.xlsb` でも 読めるか★★ ══（2026-09-21）
     ★★なぜ 要るか★★
       ★司さんの 実物は `.xlsb`★ です。
       `.xlsb` の 包みの 中は ほとんど `.bin` ですが、
       ★`xl/drawings/drawing1.xml` だけは XML の まま 残ります★（経営者1 の 実測・09-21）。
     ★★但し 板は 2進です★★（`xl/worksheets/sheet1.bin`）
       ⇒`<drawing r:id=>` を ★字として 探せません★
       ⇒★rels の `Type` で 解きます★（`lib/xlsx-zukei.js`）
     ★材料★ `tests/fixtures/kazari-hiraku3.xlsb`（経営者1 が 実Excel で 作った 作り物） */
const b材料道 = path.join(ROOT, 'tests/fixtures/kazari-hiraku3.xlsb');
const b中 = fs.existsSync(b材料道) ? fs.readFileSync(b材料道) : null;

T('★.xlsb の 材料が 在る（空振りして いない）★', () => {
  if (!b中) throw new Error('★材料が 無い★ ' + b材料道);
  const h = crypto.createHash('sha256').update(b中).digest('hex');
  if (h !== 'e4154779716440540e7e402d6f2435218b0beea6b477afcfee4e5d6d7828c6e4') {
    throw new Error('★材料が 入れ替わって います★ sha256=' + h);
  }
});

const b部品 = b中 ? ほどく(b中) : {};
T('★★.xlsb の 板は 2進・図形は XML の まま★★（★ここが 肝★）', () => {
  if (!b部品['xl/worksheets/sheet1.bin']) throw new Error('★sheet1.bin が 無い★');
  if (!b部品['xl/drawings/drawing1.xml']) throw new Error('★drawing1.xml が 無い★');
  if (b部品['xl/worksheets/sheet1.xml']) throw new Error('★sheet1.xml が 在る★（.xlsb の はず）');
  if (b部品['xl/styles.xml']) throw new Error('★styles.xml が 在る★（.xlsb は styles.bin の はず）');
});

/* ★★板の 字を 渡しません★★＝★2進なので 渡せません★（`Type` で 解けるか を 見ます） */
const b部品名 = b中 ? Z.図形の部品名('',
  b部品['xl/worksheets/_rels/sheet1.bin.rels'], 'xl/worksheets/sheet1.bin') : null;

T('★★板が 2進でも `Type` で 図形に 辿り着ける★★', () => {
  if (b部品名 !== 'xl/drawings/drawing1.xml') {
    throw new Error('★部品名が ' + b部品名 + '★（xl/drawings/drawing1.xml の はず）');
  }
});

const b図 = b部品名 && b部品[b部品名] ? Z.図形を読む(b部品[b部品名]) : [];
const b場 = b図.length ? Z.場所を決める(b図[0], { 0: 245 }, 72, {}, 24) : {};
console.log('      ＝ .xlsb の 図形 ' + b図.length + '個 ／ x=' + b場.x + ' y=' + b場.y);

T('★★.xlsb の 図形も 実Excel と 1点 以内★★（★xlsx と 同じ 数★）', () => {
  if (b図.length !== 1) throw new Error('★' + b図.length + '個★（1個 の はず）');
  if (b図[0].名 !== 'hanko') throw new Error('★名が ' + b図[0].名 + '★');
  if (b図[0].w !== 80 || b図[0].h !== 80) {
    throw new Error('★大きさが ' + b図[0].w + 'x' + b図[0].h + '★（80x80 の はず）');
  }
  if (Math.abs(b場.x - 点(449.375)) > 1) {
    throw new Error('★x が ' + b場.x + '★（' + 点(449.375).toFixed(2) + ' の はず）');
  }
  if (Math.abs(b場.y - 点(20)) > 1) {
    throw new Error('★y が ' + b場.y + '★（' + 点(20).toFixed(2) + ' の はず）');
  }
});

T('★★`r:id` の 道でも 型の 道でも 同じ 答え★★（★.xlsx で 割ります★）', () => {
  const あ = Z.図形の部品名(部品['xl/worksheets/sheet1.xml'],
    部品['xl/worksheets/_rels/sheet1.xml.rels'], 'xl/worksheets/sheet1.xml');
  const い = Z.図形の部品名('',
    部品['xl/worksheets/_rels/sheet1.xml.rels'], 'xl/worksheets/sheet1.xml');
  if (あ !== い) throw new Error('★r:id ' + あ + ' ／ 型 ' + い + '★');
});

/* ══ ★★板が 2枚 以上の `.xlsb` で 判子が 正しい 板に 行くか★★ ══（2026-09-21）
     ★★なぜ 要るか★★
       `.xlsb` の 板の 名前 ⇒ 部品名 を 引き違えると
       ★判子が 別の 板に 出ます★（★消えるのでは なく ずれる★）。
       ＝★1枚の 材料では 絶対に 割れません★
     ★★経営者1 の 断り（大事）★★
       「★Excel が 作り直した 物では 番号が 並びに 付いて 来る★」
       ⇒★★この 2本では 「番号当て」でも 同じ 答えに なります★★
       ⇒★だから 答えが 合う だけでは 「引いて いる」証しに なりません★
       ⇒★★`板たち()` が 返す `rId` を 字で 見ます★★＝★引いた 証し★
     ★材料★ ... 経営者1 が 実Excel で 作った 作り物 2本
       `ita2mai.xlsb`        1枚目 Ita1 ／ 2枚目 Ita2
       `ita2mai-irekae.xlsb` ★1枚目 Ita2 ／ 2枚目 Ita1★（並びを 入れ替えた 物） */
const XE = require_(path.join(ROOT, 'lib/xlsb-edit.js'));
const 二枚 = [
  { 名: 'ita2mai.xlsb', 大: 10792,
    sha: 'd18f873285928461272ee00408f9ad68f4e44082caa773f6aded604f356901e8',
    待つ: [{ 名: 'Ita1', rId: 'rId1', 部品: 'xl/worksheets/sheet1.bin', 判子: 'hanko1', 形: '四角' },
           { 名: 'Ita2', rId: 'rId2', 部品: 'xl/worksheets/sheet2.bin', 判子: 'hanko2', 形: '丸' }] },
  { 名: 'ita2mai-irekae.xlsb', 大: 10796,
    sha: 'acdca1da6d442cb92ef08a4cb284e92b138ab3fb18edf7bcc19683b296980afc',
    待つ: [{ 名: 'Ita2', rId: 'rId1', 部品: 'xl/worksheets/sheet1.bin', 判子: 'hanko2', 形: '丸' },
           { 名: 'Ita1', rId: 'rId2', 部品: 'xl/worksheets/sheet2.bin', 判子: 'hanko1', 形: '四角' }] },
];

for (const 本 of 二枚) {
  const 道 = path.join(ROOT, 'tests/fixtures/' + 本.名);
  const 中身 = fs.existsSync(道) ? fs.readFileSync(道) : null;
  T('★' + 本.名 + ' が 在る（空振りして いない）★', () => {
    if (!中身) throw new Error('★材料が 無い★ ' + 道);
    if (中身.length !== 本.大) throw new Error('★大きさが ' + 中身.length + '★');
    const h = crypto.createHash('sha256').update(中身).digest('hex');
    if (h !== 本.sha) throw new Error('★入れ替わって います★ ' + h);
  });
  if (!中身) continue;
  const 生 = ほどく生(中身);
  const 部 = ほどく(中身);
  /* ★`workbook.bin` は 2進の まま 渡します★（字に すると 読めません） */
  const 板 = XE.板たち(生['xl/workbook.bin'], 部['xl/_rels/workbook.bin.rels']);
  T('★★' + 本.名 + ' ＝ 板の 名前 ⇒ rId ⇒ 部品名 を 引けて いる★★', () => {
    if (!板) throw new Error('★引けません（null）★');
    if (板.length !== 本.待つ.length) throw new Error('★' + 板.length + '枚★');
    for (let i = 0; i < 板.length; i++) {
      const あ = 板[i], い = 本.待つ[i];
      if (あ.名 !== い.名) throw new Error('★板' + (i + 1) + ' の 名が ' + あ.名 + '★（' + い.名 + ' の はず）');
      if (あ.rId !== い.rId) throw new Error('★板' + (i + 1) + ' の rId が ' + あ.rId + '★（' + い.rId + ' の はず）');
      if (あ.部品 !== い.部品) throw new Error('★板' + (i + 1) + ' の 部品が ' + あ.部品 + '★');
    }
  });
  T('★★' + 本.名 + ' ＝ 判子が 正しい 板に 行く★★', () => {
    if (!板) throw new Error('★板を 引けません★');
    for (let i = 0; i < 板.length; i++) {
      const 名 = 板[i].部品.split('/').pop();
      const r = 部['xl/worksheets/_rels/' + 名 + '.rels'];
      if (!r) throw new Error('★' + 名 + ' の rels が 無い★');
      const 絵 = Z.図形の部品名('', r, 板[i].部品);
      if (!絵 || !部[絵]) throw new Error('★図形の 部品名が ' + 絵 + '★');
      const 図 = Z.図形を読む(部[絵]);
      if (図.length !== 1) throw new Error('★板' + (i + 1) + ' の 図形が ' + 図.length + '個★');
      if (図[0].名 !== 本.待つ[i].判子) {
        throw new Error('★★板' + (i + 1) + '（' + 板[i].名 + '）に ' + 図[0].名
          + ' が 出ました★★（' + 本.待つ[i].判子 + ' の はず）★判子が 別の 板に ずれて います★');
      }
      if (図[0].種類 !== 本.待つ[i].形) {
        throw new Error('★板' + (i + 1) + ' の 形が ' + 図[0].種類 + '★（' + 本.待つ[i].形 + ' の はず）');
      }
    }
  });
}

/* ★★入れ替えた 本で 「並び当て」なら どう なるかを 出します★★
     ＝★経営者1 の 断り（どちらでも 同じ 答えに なる）を 実物で 確かめる★
     ＝★同じなら 「割れない」と 書く／違えば 「割れる」と 書く★ */
{
  const 道 = path.join(ROOT, 'tests/fixtures/ita2mai-irekae.xlsb');
  if (fs.existsSync(道)) {
    const 生 = ほどく生(fs.readFileSync(道));
    const 部 = ほどく(fs.readFileSync(道));
    const 板 = XE.板たち(生['xl/workbook.bin'], 部['xl/_rels/workbook.bin.rels']) || [];
    const 並び = ['xl/worksheets/sheet1.bin', 'xl/worksheets/sheet2.bin'];
    const 同じ = 板.length === 2 && 板[0].部品 === 並び[0] && 板[1].部品 === 並び[1];
    console.log('      ＝ 入れ替えた 本で 「引く」と 「並び」の 答えは '
      + (同じ ? '★同じ★（この 材料では 割れません）' : '★違う★（割れます）'));
  }
}

console.log('');
console.log('  ★見て いない 事★');
console.log('    ・★色は 付けて いません★＝テーマの 色が 決め打ちに できるかは ★未測定★');
console.log('    ・`oneCellAnchor` / `absoluteAnchor` は ★まだ★（この 材料に 在りません）');
console.log('    ・図形の 中の 字／線の 太さ／回転の 見た目は ★未測定★');
console.log('    ・★★`.xlsb` の マスの 飾りは まだ★★＝`xl/styles.bin`（2進）');
console.log('      ＝お客さんの 道で 測ると ★.xlsb は 8 / 13★（図形は 出ます）');
console.log('    ・★★.xlsb の 板の 名前は 「並び」で 当てて います★★');
console.log('      ＝★板が 2枚 以上 在る `.xlsb` で 合うかは 未測定★');
console.log('    ・★画面に 描いて いるかは ここでは 測って いません★');
console.log('      ＝`docs/measured/hakaru-kazari-ga-gamen-made-todoku-ka.mjs`');
console.log('xlsx-zukei: ' + pass + ' 緑 / ' + fail + ' 赤');

if (process.argv.includes('--self-test')) {
  console.log('');
  console.log('--self-test: ★わざと壊して 赤に なるか★（★lib は 1字も 触りません★）');
  const 絵 = 部品[部品名];
  const 板 = 部品['xl/worksheets/sheet1.xml'];
  const rels = 部品['xl/worksheets/_rels/sheet1.xml.rels'];
  /* ★★壊し方の 表★★＝★元の 字を 書き換えて 押します★ */
  const 壊し方 = [
    { 名: 'マスの 番号を 変える（場所が 動く はず）',
      絵: (x) => x.replace(new RegExp('<xdr:col>5</xdr:col>'), '<xdr:col>1</xdr:col>'),
      見る: (図) => Math.abs(Z.場所を決める(図[0], { 0: 245 }, 72, {}, 24).x - 点(449.375)) > 1 },
    { 名: '大きさを 半分に する（80 → 40 に なる はず）',
      絵: (x) => x.replace(new RegExp('cx="762000" cy="762000"'), 'cx="381000" cy="381000"'),
      見る: (図) => 図[0] && 図[0].w === 40 && 図[0].h === 40 },
    { 名: '名前を 変える（hanko で なくなる はず）',
      絵: (x) => x.replace(new RegExp('name="hanko"'), 'name="betsu"'),
      見る: (図) => 図[0] && 図[0].名 === 'betsu' },
    { 名: '形を 丸に する（種類が 丸に なる はず）',
      絵: (x) => x.replace(new RegExp('prst="rect"'), 'prst="ellipse"'),
      見る: (図) => 図[0] && 図[0].種類 === '丸' },
    { 名: '`<xdr:sp>` を 消す（図形が 0個に なる はず）',
      絵: (x) => x.replace(new RegExp('<xdr:sp ', 'g'), '<xdr:zzz '),
      見る: (図) => 図.length === 0 },
  ];
  let 悪 = 0;
  for (const こ of 壊し方) {
    const 絵2 = こ.絵(絵);
    if (絵2 === 絵) { console.log('  NG   ★1字も 変わって いません★ ' + こ.名); 悪++; continue; }
    if (こ.見る(Z.図形を読む(絵2))) console.log('  ok   ' + こ.名);
    else { console.log('  NG   ★壊したのに 同じ 答え★ ' + こ.名); 悪++; }
  }
  /* ★rels を 壊すと 部品名が 出ない はず★ */
  const r2 = rels.replace(new RegExp('drawings/drawing1.xml'), 'drawings/nai.xml');
  const 名2 = Z.図形の部品名(板, r2, 'xl/worksheets/sheet1.xml');
  if (名2 === 'xl/drawings/nai.xml') console.log('  ok   rels の 行き先を 変えると 部品名も 変わる');
  else { console.log('  NG   ★rels を 変えたのに ' + 名2 + '★'); 悪++; }
  process.exit(悪 ? 1 : 0);
}
process.exit(fail ? 1 : 0);
