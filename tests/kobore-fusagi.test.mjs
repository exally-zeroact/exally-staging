/* kobore-fusagi.test.mjs — ★塞がれた 溢れ(#SPILL!)を xlsx へ 書き出す時★
 *
 *  ★★なぜ 要るか★★（2026-09-20 実Excel で 実測・経営者1 の ㊺）
 *    溢れる 式の 行き先が 塞がれて いる 時、画面は `#SPILL!` を 出します。
 *    その まま 書き出すと 頭に `cm` も `ref` も 付かない ので
 *        <c r="D1" t="str"><f>_xlfn.SEQUENCE(3)</f><v>#SPILL!</v></c>
 *    ⇒ 実Excel は ★昔の（動かない）式★として 読み、★暗黙の 交差で 先頭だけ★ 返します。
 *    ⇒★★D1 が 黙って `1`（Double・ISNUMBER True）に なりました★★
 *    ⇒★誤りを 出すべき 所で 数を 出す＝お客さんは それを 正しい 答えだと 思います★
 *    ⇒ 塞ぎ（D2）を 消しても ★溢れ直しません★（昔の式として 固定される）
 *
 *  ★★直した 後（実Excel で 実測）★★
 *        <c r="D1" t="str" cm="1"><f t="array" ref="D1:D1">_xlfn.SEQUENCE(3)</f><v>#SPILL!</v></c>
 *    ⑴ 開いた 瞬間 ... ★-2146826243＝#SPILL!★（ISERROR True・ISNUMBER False）
 *    ⑵ HasArray ..... False（＝動く並び）
 *    ⑶ 塞ぎを 消す .. ★D1 1 / D2 2 / D3 3 に 溢れ直す★／式は D1 だけ
 *    ⇒★★㊵の 実Excel 本物（台本3・台本4）と 1字 違わず 同じ★★
 *
 *  ★なぜ `ref` が 自分 1マス（D1:D1）か★
 *    ㋐ グリッドに ★本来の 広さが 残って いません★
 *       （塞がれた 時は `_溢れ元` が 1つも 付かない＝溢れて いないので）
 *    ㋑ `ref="D1:D3"` は ★中身が 食い違います★
 *       D2 は ★値を 持つ マス★なのに ★配列の 一部★にも なる＝両方は 成り立たない
 *    ⇒ `D1:D1`＝「1マスに 収まる 動く並び」なら 筋が 通り、
 *       実Excel が 計算して ★自分で #SPILL! を 出します★
 *
 *  ★★見て いない 事★★（★書かない 見張りは「全部 守った」と 読まれる★）
 *    ・★`#SPILL!` の ★字★を 見て 判じて います★＝★飾りの 字に 頼った 門に 近い★
 *      ⇒台が 日本語の `#スピル!` を 出す 道が 出来たら ★黙って 割れます★
 *      ⇒2026-09-20 に 数えた 時点では ★`#スピル!` を 作る 所は 0か所★
 *        （「スピル」の 字 9件は ★全部 注（コメント）★）
 *      ⇒★次は 台に 広さを 聞く 形に 直す★（`lib/shiki-hyou.js` は
 *        `#SPILL!` を 返す 直前に `v.行数` `v.列数` を 持って います）
 *    ・★横・2次元が 塞がれた 時は 実Excel で 測って いません★（縦だけ）
 *    ・★`#SPILL!` 以外の 誤りは ここでは 見ません★（誤りの 型は 別件）
 *
 *  使い方: node tests/kobore-fusagi.test.mjs
 *          node tests/kobore-fusagi.test.mjs --self-test
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

/* ★包みを ほどく★（借り物に 読ませず 生の字を 見る） */
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
function 紙(data) {
  const buf = Buffer.from(IO.writeBook(G.gridToBook([{ name: 'S', data: data, colW: {} }])));
  return ほどく(buf)['xl/worksheets/sheet1.xml'] || '';
}
/* ★正規表現を 使いません★＝逆斜線の 逃がしが 落ちる（2026-09-20 に 2回 踏んだ） */
function マス(s, a) {
  const 頭 = s.indexOf('<c r="' + a + '"');
  if (頭 < 0) return null;
  const 尾 = s.indexOf('</c>', 頭);
  return 尾 < 0 ? null : s.slice(頭, 尾 + 4);
}

/* ══ 材料 ═══════════════════════════════════════════════════════════
     ★塞がれた★ ... D1(0,3) に 溢れる 式・D2(1,3) に 値 ⇒ 画面は #SPILL!
     ★対照★     ... F1(0,5) に 同じ 式・F1:F3 は 空 ⇒ ちゃんと 溢れる     */
const 塞がれた = {
  '0,3': { f: '=SEQUENCE(3)', d: '#SPILL!' },
  '1,3': { v: '9', d: '9' },
  '0,5': { f: '=SEQUENCE(3)', d: '1' },
  '1,5': { d: '2', _溢れ元: '0,5' },
  '2,5': { d: '3', _溢れ元: '0,5' },
};
const 誤りだが溢れない = { '0,0': { f: '=1/0', d: '#DIV/0!' } };
const 塞がれた横 = { '0,0': { f: '=SEQUENCE(1,3)', d: '#SPILL!' }, '0,1': { v: 'x', d: 'x' } };

export function 走らせる() {
  console.log('kobore-fusagi — 塞がれた 溢れを どう 書き出すか');
  const s = 紙(塞がれた);

  T('① 塞がれた 頭に cm="1" が 付く', () => {
    const c = マス(s, 'D1');
    if (!c) throw new Error('D1 が 無い');
    if (c.indexOf('cm="1"') < 0) throw new Error(c);
  });
  T('② 塞がれた 頭の ref は 自分 1マス（D1:D1）', () => {
    const c = マス(s, 'D1') || '';
    if (c.indexOf('ref="D1:D1"') < 0) throw new Error(c);
  });
  T('③ 塞ぎ(D2)は ただの 値の まま', () => {
    const c = マス(s, 'D2') || '';
    if (c.indexOf('cm=') >= 0 || c.indexOf('t="array"') >= 0) throw new Error(c);
    if (c.indexOf('<v>9</v>') < 0) throw new Error(c);
  });
  T('④ 対照(F1)は 本来の 広さ（F1:F3）の まま＝下がって いない', () => {
    const c = マス(s, 'F1') || '';
    if (c.indexOf('ref="F1:F3"') < 0) throw new Error(c);
    if (c.indexOf('cm="1"') < 0) throw new Error(c);
  });
  T('⑤ 溢れの 数は 2つ（塞がれた 1 ＋ 対照 1）', () => {
    const cm = (s.match(/cm="1"/g) || []).length;
    const ar = (s.match(/t="array"/g) || []).length;
    if (cm !== 2 || ar !== 2) throw new Error('cm=' + cm + ' array=' + ar);
  });
  T('⑥ 溢れない 誤り(#DIV/0!)には 付けない', () => {
    const s2 = 紙(誤りだが溢れない);
    if (/cm="1"/.test(s2)) throw new Error('cm が 付いた: ' + (マス(s2, 'A1') || ''));
    if (/t="array"/.test(s2)) throw new Error('t="array" が 付いた');
  });
  T('⑦ 横が 塞がれた 時も 自分 1マス（A1:A1）', () => {
    const s3 = 紙(塞がれた横);
    const c = マス(s3, 'A1') || '';
    if (c.indexOf('ref="A1:A1"') < 0) throw new Error(c);
    if (c.indexOf('cm="1"') < 0) throw new Error(c);
  });
  T('⑧ 読み戻しても 式が 生きて いる', () => {
    const buf = Buffer.from(IO.writeBook(G.gridToBook([{ name: 'S', data: 塞がれた, colW: {} }])));
    const cells = IO.readBook(buf).sheets[0].cells;
    const f = cells['D1'] && cells['D1'].f;
    if (!f || f.indexOf('SEQUENCE') < 0) throw new Error('D1 の f=' + JSON.stringify(f));
  });

  console.log('  → ' + pass + ' 通り / ' + fail + ' 失敗');
  return fail;
}

export const 本数 = 8;

if (process.argv[1] && path.resolve(process.argv[1]) === path.resolve(fileURLToPath(import.meta.url))) {
  if (process.argv.includes('--self-test')) {
    /* ★★わざと 壊して 赤に なるか★★＝★ 3通り 壊します★
         （★「赤が 1個 出た」で 見張りを 信用するな★—— 記憶の 決まり）
         ★ここで 壊すのは 材料です★（lib は 1字も 触りません）
         ★lib を 壊した 測りは 別に 取って あります★
           ＝★範囲(F)を 付けない／印(D)を 立てない／字を 間違える／
             範囲を 広く する／塞がれて いない 方を 消す の 5か所★
           ⇒★ 5か所とも 赤★（壊す 所ごとに ★違う 試験★が 赤） */
    let 赤 = 0;
    const T2 = (n, ok, m) => {
      if (ok) console.log('  ✓ ' + n);
      else { 赤++; console.log('  ✗ ' + n + (m ? ' — ' + m : '')); }
    };

    /* ㋐★表示を 表から 抜く★＝`#SPILL!` が 無ければ 付けて は いけない */
    {
      const 壊 = JSON.parse(JSON.stringify(塞がれた));
      delete 壊['0,3'].d;
      const c = マス(紙(壊), 'D1') || '';
      T2('㋐ 表示を 表から 抜く ⇒ cm が 付かない', c.indexOf('cm="1"') < 0, c);
    }
    /* ㋑★別の 字に 書き換える★＝数なら 付けて は いけない */
    {
      const 壊 = JSON.parse(JSON.stringify(塞がれた));
      壊['0,3'].d = '1';
      const c = マス(紙(壊), 'D1') || '';
      T2('㋑ 表示を 1 に 書き換える ⇒ cm が 付かない', c.indexOf('cm="1"') < 0, c);
    }
    /* ㋒★塞ぎを 表から 抜く★＝溢れる ので ★本来の 広さ★に なる はず */
    {
      const 壊 = JSON.parse(JSON.stringify(塞がれた));
      delete 壊['1,3'];
      壊['0,3'].d = '1';
      壊['1,3'] = { d: '2', _溢れ元: '0,3' };
      壊['2,3'] = { d: '3', _溢れ元: '0,3' };
      const c = マス(紙(壊), 'D1') || '';
      T2('㋒ 塞ぎを 抜いて 溢れさせる ⇒ ref が D1:D3 に なる',
         c.indexOf('ref="D1:D3"') >= 0, c);
    }
    console.log('  → 壊した 3通りの うち 拾えなかった のは ' + 赤 + '個');
    process.exit(赤 ? 1 : 0);
  }
  process.exit(走らせる() ? 1 : 0);
}
