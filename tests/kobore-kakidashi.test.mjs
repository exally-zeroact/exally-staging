/* kobore-kakidashi.test.mjs — ★溢れ(スピル)を xlsx へ 書き出す時に 運べているか★
 *
 *  ★なぜ 要るか★（2026-09-20 生の xml で 実測）
 *    溢れた 式は 頭の マスに だけ 式が 在り、先の マスは 値です。
 *    今まで 書き出しは 頭を ★普通の 1マスの 式★として 書いて いました:
 *        <c r="A1"><f>_xlfn.SEQUENCE(3)</f><v>1</v></c>
 *        <c r="A2"><v>2</v></c>
 *    ⇒ 実Excel で 開くと 隣に 値が 在る ので ★#SPILL!★ です。
 *
 *  ★正しい 形★（実Excel が 作った xlsx を ほどいて 数えた 形と 同じ）
 *        <c r="A1" cm="1"><f t="array" ref="A1:A3">_xlfn.SEQUENCE(3)</f><v>1</v></c>
 *        <c r="A2"><v>2</v></c>
 *    ・`ref=`  ＝ 溢れの 範囲（SheetJS には `F` で 渡す）
 *    ・`cm="1"`＝ ★動く並び(dynamic array)の 印★（SheetJS には `D` で 渡す）
 *                 `xl/metadata.xml` は SheetJS が いつも 書いて います（904B・
 *                 `dynamicArrayProperties fDynamic="1"` 在り）。
 *
 *  ★ここが 守る 事★
 *    ・溢れの 印(`_溢れ元`)から ★範囲を 数え直せる★（縦・横・2次元・頭が 途中・2つ在る）
 *    ・★溢れない 式には 付けない★（普通の 式が 配列式に なると 実Excel の 見え方が 変わる）
 *    ・書き出した 物を ★読み戻して 式が 生きている★（引いて 確かめる）
 *    ・`xl/metadata.xml` が 包みに 入って いる
 *
 *  使い方: node tests/kobore-kakidashi.test.mjs
 *          node tests/kobore-kakidashi.test.mjs --self-test  ＝ わざと壊して赤になるか
 */
import path from 'node:path';
import zlib from 'node:zlib';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createRequire } from 'node:module';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(__dirname, '..');
const require = createRequire(pathToFileURL(path.join(ROOT, 'package.json')));
const G = require(path.join(ROOT, 'lib', 'grid-xlsx.js'));
const IO = require(path.join(ROOT, 'lib', 'xlsx-io.js'));

let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ✓ ' + n); } catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + e.message); } };

/* ★包みを ほどく★（借り物に 読ませず 局所ヘッダを 順に 舐めて ★生の字★を 見る
     ─ SheetJS に 読ませると「SheetJS が どう 読んだか」しか 分からない） */
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

function 書く(data) {
  return Buffer.from(IO.writeBook(G.gridToBook([{ name: 'S', data: data, colW: {} }])));
}
function 紙(data) { return ほどく(書く(data))['xl/worksheets/sheet1.xml'] || ''; }

/* 頭の マスの `ref=` を 取る（無ければ null） */
function 範囲(s) {
  const m = s.match(/<c r="[A-Z]+\d+" cm="1"><f t="array" ref="([^"]+)">/);
  return m ? m[1] : null;
}

/* ══ 形ごとの 材料 ═══════════════════════════════════════════════════
     ★`_溢れ元` は 先の マスだけが 持ちます★（頭には 付きません＝
       `_溢れを写す` が i===0&&j===0 を 飛ばす 為）                     */
const 縦3 = { '0,0': { f: '=SEQUENCE(3)', d: '1' }, '1,0': { d: '2', _溢れ元: '0,0' }, '2,0': { d: '3', _溢れ元: '0,0' } };
const 横3 = { '0,0': { f: '=SEQUENCE(1,3)', d: '1' }, '0,1': { d: '2', _溢れ元: '0,0' }, '0,2': { d: '3', _溢れ元: '0,0' } };
const 二次元 = {
  '0,0': { f: '=SEQUENCE(2,3)', d: '1' }, '0,1': { d: '2', _溢れ元: '0,0' }, '0,2': { d: '3', _溢れ元: '0,0' },
  '1,0': { d: '4', _溢れ元: '0,0' }, '1,1': { d: '5', _溢れ元: '0,0' }, '1,2': { d: '6', _溢れ元: '0,0' },
};
const 途中から = { '1,1': { f: '=SEQUENCE(2)', d: '1' }, '2,1': { d: '2', _溢れ元: '1,1' } };
const 溢れない = { '0,0': { f: '=SUM(1,2)', d: '3' } };
const 二つ = {
  '0,0': { f: '=SEQUENCE(2)', d: '1' }, '1,0': { d: '2', _溢れ元: '0,0' },
  '0,2': { f: '=SEQUENCE(3)', d: '1' }, '1,2': { d: '2', _溢れ元: '0,2' }, '2,2': { d: '3', _溢れ元: '0,2' },
};

export function 走らせる() {
  console.log('kobore-kakidashi — 溢れを xlsx へ 運べて いるか');

  T('① 縦 3 は ref="A1:A3"', () => {
    const r = 範囲(紙(縦3));
    if (r !== 'A1:A3') throw new Error('ref=' + r);
  });
  T('② 横 3 は ref="A1:C1"', () => {
    const r = 範囲(紙(横3));
    if (r !== 'A1:C1') throw new Error('ref=' + r);
  });
  T('③ 2次元 2x3 は ref="A1:C2"', () => {
    const r = 範囲(紙(二次元));
    if (r !== 'A1:C2') throw new Error('ref=' + r);
  });
  T('④ 頭が B2 からでも ref="B2:B3"', () => {
    const r = 範囲(紙(途中から));
    if (r !== 'B2:B3') throw new Error('ref=' + r);
  });
  T('⑤ 溢れない 式には 付けない', () => {
    const s = 紙(溢れない);
    if (/t="array"/.test(s)) throw new Error('t="array" が 付いた');
    if (/cm="1"/.test(s)) throw new Error('cm="1" が 付いた');
  });
  T('⑥ 溢れが 2つ なら cm も t="array" も 2つ', () => {
    const s = 紙(二つ);
    const cm = (s.match(/cm="1"/g) || []).length;
    const ar = (s.match(/t="array"/g) || []).length;
    if (cm !== 2 || ar !== 2) throw new Error('cm=' + cm + ' array=' + ar);
  });

  /* ★入れて 落ちないかで 測らない＝★引いて★ 正しい 答えが 出るかで 測る★ */
  T('⑦ 読み戻しても 頭の 式が 生きて いる', () => {
    const cells = IO.readBook(書く(縦3)).sheets[0].cells;
    const f = cells['A1'] && cells['A1'].f;
    if (!f || f.charAt(0) !== '=') throw new Error('A1 の f=' + JSON.stringify(f));
    if (f.indexOf('SEQUENCE') < 0) throw new Error('式が 変わった: ' + f);
  });
  T('⑧ 読み戻しても 先の 値が 残る', () => {
    const cells = IO.readBook(書く(縦3)).sheets[0].cells;
    if (!cells['A2'] || cells['A2'].v !== 2) throw new Error('A2=' + JSON.stringify(cells['A2']));
    if (!cells['A3'] || cells['A3'].v !== 3) throw new Error('A3=' + JSON.stringify(cells['A3']));
  });

  /* ★`cm="1"` は `xl/metadata.xml` を 指す＝紙が 無いと 実Excel が 開けない★ */
  T('⑨ xl/metadata.xml が 包みに 入って いる', () => {
    const 中 = ほどく(書く(縦3));
    const meta = 中['xl/metadata.xml'];
    if (!meta) throw new Error('metadata.xml が 無い');
    if (meta.indexOf('dynamicArrayProperties') < 0) throw new Error('dynamicArrayProperties が 無い');
    const ct = 中['[Content_Types].xml'] || '';
    if (ct.indexOf('metadata') < 0) throw new Error('Content_Types に 無い');
  });

  /* ★`_xlfn.` が 外れると 実Excel は SEQUENCE を 知らない 名前に する★ */
  T('⑩ 頭の 式に _xlfn. が 付いて いる', () => {
    const s = 紙(縦3);
    if (s.indexOf('_xlfn.SEQUENCE(3)') < 0) throw new Error('_xlfn. が 無い');
  });

  console.log('  → ' + pass + ' 通り / ' + fail + ' 失敗');
  return fail;
}

export const 本数 = 10;

if (process.argv[1] && path.resolve(process.argv[1]) === path.resolve(fileURLToPath(import.meta.url))) {
  if (process.argv.includes('--self-test')) {
    /* ★わざと 壊して 赤に なるか★＝溢れの 印を 消した 物を 押す */
    console.log('--self-test: 溢れの 印を 消して 押す');
    const 壊 = JSON.parse(JSON.stringify(縦3));
    delete 壊['1,0']._溢れ元; delete 壊['2,0']._溢れ元;
    const r = 範囲(紙(壊));
    if (r !== null) { console.log('  ✗ 印を 消したのに ref=' + r); process.exit(1); }
    console.log('  ✓ 印を 消すと ref が 消える（門が 見て いる）');
    process.exit(0);
  }
  process.exit(走らせる() ? 1 : 0);
}
