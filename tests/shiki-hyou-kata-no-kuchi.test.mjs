/* shiki-hyou-kata-no-kuchi.test.mjs — ★型つきの 口（`置く`）と「直に 書き換えて いないか」の 見張り★（2026-09-15）
 *
 *  ★★なぜ 要るか★★
 *    台の `打つ` は ★字しか 受けません★（`=` で 始まれば 式）。
 *    ⇒ ★字の "2"★ や ★真偽★ を 置く 口が 無かったので、呼ぶ側が
 *      ★`表.中身[マス].値 = {…}` を 直に 書き換えて★ いました（`tests/zairyou.mjs`／測り道具）。
 *    ⇒ ★★直に 書き換えても そのマスを 見て いる 式に 報せが 行きません★★
 *      ＝★先に 式を 打つと その式は 古い 答えの まま★。
 *    ★今の 測り道具は「材料 先・式 後」なので 当たって いません★が、
 *    ★順を 変えた 途端に 黙って 嘘に なります★（＝★誰も 止めない 型★）。
 *
 *  ★★この 見張りが 守る 事★★
 *    ①★`置く` が 5つの 型を 受ける★（数／字／真偽／誤／空）
 *    ②★★`置く` の 後に 見て いる 式が 直る★★（★★⑲の 芯★★）
 *    ③★分からない 形は 黙って 通さない★（投げる）
 *    ④★★`中身[…].値 =` の 直書きが repo に 何個 在るか を 数えて 0 で なければ 赤★★
 *       ＝★「全部 移した」を 人の 記憶で 保たない★（指示役1 の 注文・2026-09-15）
 *
 *  ★★★この 見張りが 見て いない 形（＝守れて いない 物）★★★
 *    ★書かないと「全部 守った」と 読まれます★ので 名指しで 書きます。
 *    ★字で 探して いる 限り 必ず 抜けます★＝★下の 形は 赤に なりません★：
 *      ㋐★2つ先の 受け渡し★ … `const m = 表.中身[a]; なにか(m);`
 *          ⇒★呼んだ 先で `m.値 =` された 分は 見えません★
 *      ㋑★回して 受ける★ … `for (const m of 何か) { m.値 = … }`
 *          ⇒★`中身[` が その行に 無いので 見えません★
 *      ㋒★名前を 組み立てる★ … `表['中' + '身'][a].値 = …`
 *      ㋓★6行 以上 離れる★ … 変数に 受けてから ★7行 後★に 書き換える（今は 5行まで）
 *      ㋔★`中身` を 丸ごと 渡す★ … `こわす(表.中身)`
 *    ⇒★★だから この 見張りは「直書きが 0個」を 言えますが
 *         「追従しない 書き方が 0個」は 言えません★★
 *    ⇒★本当の 守りは ★試験 ⑥⑦⑧⑨★（★置いた 事が 式に 伝わるか★）です。
 *      ★この 数え上げは その 念押し★＝★数だけで 緑に しない★。
 *
 *  ★★見て いない 場所★★
 *    ㋕★`book.html` は 入れて いません★
 *        ＝★別の `中身`★を 持って います（★描く物の 中身★＝`g.中身[j].物`）。
 *        ★同じ 名前でも 別の 物★なので 混ぜると ★意味の 無い 赤★が 出ます。
 *        ⇒★その 代わり `book.html` の 中の 直書きは この 見張りでは 見つかりません★。
 *          （`book.html` は 台の `表` を まだ 使って いません＝今は 当たりませんが、
 *           ★㋐画面を 台へ 寄せる 時に ここを 見直す 事★）
 *    ㋖★`tools/` `scripts/` も 見て いません★（★台を 使って いないから★）。
 *        ⇒★台を 使い 始めたら 見る 場所に 足す★。
 *    ㋗★免除 2本の 中は 見て いません★（台の 実装／この 見張り自身）。
 *
 *  使い方: node tests/shiki-hyou-kata-no-kuchi.test.mjs
 *          node tests/shiki-hyou-kata-no-kuchi.test.mjs --self-test
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';
import { 注記を外す } from '../scripts/lib/chuki.mjs';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(path.join(ROOT, 'package.json'));
const H = require_(path.join(ROOT, 'lib/shiki-hyou.js'));

let pass = 0, fail = 0;
const T = (n, f) => {
  try { f(); pass++; console.log('  ✓ ' + n); }
  catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + e.message); }
};

console.log('\n[shiki-hyou-kata-no-kuchi] ★型つきの 口＋直書きの 見張り★');

/* ══════════ ㋐ 口そのもの ══════════ */

T('★① 字の "2" は 数に ならない★（`打つ` との 違い）', () => {
  const h = H.表();
  h.打つ('A1', '2');
  h.置く('B1', '2');
  const a = h.値('A1'), b = h.値('B1');
  if (a.型 !== '数') throw new Error('打つ … ' + a.型 + '（数 のはず）');
  if (b.型 !== '字') throw new Error('★置く … ' + b.型 + '（字 のはず）★');
  if (b.値 !== '2') throw new Error('置く … ' + JSON.stringify(b.値));
});

T('★② 5つの 型を 受ける★（数／字／真偽／誤／空）', () => {
  const h = H.表();
  h.置く('A1', 5);
  h.置く('A2', 'あ');
  h.置く('A3', true);
  h.置く('A4', { 型: '誤', 値: '#N/A' });
  h.置く('A5', null);
  const 見 = (n) => h.値(n).型 + ':' + String(h.値(n).値);
  if (見('A1') !== '数:5') throw new Error('A1 … ' + 見('A1'));
  if (見('A2') !== '字:あ') throw new Error('A2 … ' + 見('A2'));
  if (見('A3') !== '真偽:true') throw new Error('A3 … ' + 見('A3'));
  if (見('A4') !== '誤:#N/A') throw new Error('A4 … ' + 見('A4'));
  if (見('A5') !== '空:undefined' && h.値('A5').型 !== '空') throw new Error('A5 … ' + 見('A5'));
});

T('★③ 台の 値の 形（{型,値}）も そのまま 受ける★', () => {
  const h = H.表();
  h.置く('A1', { 型: '字', 値: '007' });
  if (h.字('A1') !== '007') throw new Error('A1 … ' + h.字('A1'));
  h.置く('A2', { 型: '数', 値: 7 });
  if (h.字('A2') !== '7') throw new Error('A2 … ' + h.字('A2'));
});

T('★④ 分からない 形は 黙って 通さない★（投げる）', () => {
  const h = H.表();
  let 投げた = false;
  try { h.置く('A1', { なんとか: 1 }); } catch (e) { 投げた = true; }
  if (!投げた) throw new Error('★形が 違うのに 通りました★＝黙って 嘘に なります');
});

T('★⑤ `打った字` は 省ける／渡せる★', () => {
  const h = H.表();
  /* ★`中身` を 直に 引かない★＝★鍵は `板名!A1`★（2026-09-15・㍸） */
  h.置く('A1', true);
  if (h.打った字('A1') !== 'TRUE') throw new Error('省いた時 … ' + h.打った字('A1'));
  h.置く('A2', { 型: '字', 値: '2' }, '2');
  if (h.打った字('A2') !== '2') throw new Error('渡した時 … ' + h.打った字('A2'));
});

/* ══════════ ㋑ ★⑲の 芯＝報せが 行くか★ ══════════ */

T('★★⑥ 先に 式を 打ってから 置いても 式が 直る★★（⑲の 芯）', () => {
  const h = H.表();
  h.打つ('B1', '=A1&"x"');      /* ★式が 先★ */
  h.置く('A1', 'あ');            /* ★材料が 後★ */
  if (h.字('B1') !== 'あx') throw new Error('★B1 が ' + h.字('B1') + '（あx のはず）★'
    + '＝★置いた 事が 式に 伝わって いません★');
  h.置く('A1', 'い');            /* ★置き直し★ */
  if (h.字('B1') !== 'いx') throw new Error('★置き直し後 B1 が ' + h.字('B1') + '（いx のはず）★');
});

T('★⑦ 鎖の 先まで 届く★（A→B→C）', () => {
  const h = H.表();
  h.打つ('B1', '=A1*2');
  h.打つ('C1', '=B1+1');
  h.置く('A1', 10);
  if (h.字('C1') !== '21') throw new Error('C1 … ' + h.字('C1') + '（21 のはず）');
  h.置く('A1', 100);
  if (h.字('C1') !== '201') throw new Error('★置き直し後 C1 が ' + h.字('C1') + '（201 のはず）★');
});

T('★⑧ 菱形でも 古い 答えが 残らない★（⑱と 同じ 盤を `置く` で）', () => {
  const h = H.表();
  h.打つ('Q1', '=K1*2');
  h.打つ('M1', '=K1-Q1');
  h.置く('K1', 10);
  if (h.字('M1') !== '-10') throw new Error('M1 … ' + h.字('M1') + '（-10 のはず）');
  h.置く('K1', 100);
  if (h.字('M1') !== '-100') throw new Error('★置き直し後 M1 が ' + h.字('M1') + '（-100 のはず）★');
});

T('★⑨ 式だった マスに 置くと 式が 消える★（前の 頼りも 外れる）', () => {
  const h = H.表();
  h.打つ('A1', '5');
  h.打つ('B1', '=A1*2');
  h.打つ('C1', '=B1+1');
  if (h.字('C1') !== '11') throw new Error('先に C1 … ' + h.字('C1'));
  h.置く('B1', 100);                 /* ★B1 の 式を 数で 潰す★ */
  if (h.字('C1') !== '101') throw new Error('C1 … ' + h.字('C1') + '（101 のはず）');
  if (h.見ている('B1').length) throw new Error('★B1 の 頼りが 残って います★ … '
    + h.見ている('B1').join(','));
  h.打つ('A1', '999');               /* ★もう B1 は A1 を 見て いない★ */
  if (h.字('B1') !== '100') throw new Error('★B1 が ' + h.字('B1') + '（100 のまま のはず）★');
});

/* ══════════ ㋒ ★直書きが 残って いないか を 数える★ ══════════ */

/* ★見る 範囲を 先に 決める★（★決めずに 数えると 今日の 形を 通します★）
     ㋐ lib/         … 台そのもの ＋ 台を 使う 本番の 部品
     ㋑ js/          … 本番の 画面まわり
     ㋒ tests/       … 試験
     ㋓ docs/measured/ … 測り道具
   ★book.html は 別の `中身`（描く物の 中身）を 持って いるので 入れない★
     ＝★同じ 名前でも 別の 物★。混ぜると ★意味の 無い 赤★が 出ます。 */
const 見る場所 = ['lib', 'js', 'tests', 'docs/measured'];
/* ★免除★（★免除は 増やす ほど 見張りが 弱く なる★ので 訳を 書く）
     ㋓ lib/shiki-hyou.js … ★台の 中の 実装★（ここだけは 直に 書いて よい）
     ㋔ この 見張り自身 … ★探す 形を 試す 為の 文字列★を 持って います */
const 免除 = new Set([
  'lib/shiki-hyou.js',
  'tests/shiki-hyou-kata-no-kuchi.test.mjs'
]);

/* ★覚書きは 数えない★（覚書きは 動きません。
     ★覚書きを 数えると「この 形は ダメ」と 書いた 人が 赤に なります★）
   ★★外す 道具は 自前で 書かない★★＝`scripts/lib/chuki.mjs`（2026-08-26 の 決まり）
     ★訳★…★字の 中の `//` や 正規表現の 中の `/` を 注記と 読む★のを 避ける為。
     ★私は 2026-09-15 に これを 自前で 書いて `chuki.test.mjs` に 赤く されました★
     （見張りは 正しかった＝★作る前に 探せ★）
   ★`残す` の 既定が true★＝★行番号を ずらさない★ */
const 素にする = (s) => 注記を外す(s);

function 集める(d) {
  const p = path.join(ROOT, d);
  if (!fs.existsSync(p)) return [];
  const 出 = [];
  for (const e of fs.readdirSync(p, { withFileTypes: true })) {
    const 中 = path.join(d, e.name).replace(/\\/g, '/');
    if (e.isDirectory()) { 出.push(...集める(中)); continue; }
    if (!/\.(js|mjs|cjs)$/.test(e.name)) continue;
    if (/min\.js$/.test(e.name)) continue;      /* ★借り物は 見ない★ */
    出.push(中);
  }
  return 出;
}

/* ★2通りで 探す★（★1通りだと 変数に 受けて 逃げられる★）
     ㋐ `…中身[ … ].値 =` ／ `.打った字 =`   … 直に 書く 形
     ㋑ `… = …中身[ … ]` で 受けた 変数に、★その後 5行 以内★で `.値 =` ／ `.打った字 =` */
const 直書き = /中身\s*\[[^\]]*\]\s*\.\s*(値|打った字)\s*=[^=]/;
const 受け = /(?:var|let|const)\s+([A-Za-z0-9_$぀-ヿ一-鿿]+)\s*=\s*[^;]*中身\s*\[/;

function 直に書いている(全文) {
  const 行 = 全文.split('\n');
  const 当 = [];
  for (let i = 0; i < 行.length; i++) {
    if (直書き.test(行[i])) { 当.push({ 行: i + 1, 形: '㋐直に', 字: 行[i].trim().slice(0, 78) }); continue; }
    const m = 受け.exec(行[i]);
    if (!m) continue;
    const 名 = m[1];
    const 後 = new RegExp('(?:^|[^A-Za-z0-9_$])' + 名.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
      + '\\s*\\.\\s*(値|打った字)\\s*=[^=]');
    for (let k = i + 1; k <= Math.min(i + 5, 行.length - 1); k++) {
      if (後.test(行[k])) { 当.push({ 行: k + 1, 形: '㋑受けて', 字: 行[k].trim().slice(0, 78) }); break; }
    }
  }
  return 当;
}

T('★★⑩ `中身[…].値` の 直書きが 0個★★（★数えて 出す★）', () => {
  const 本 = [];
  for (const d of 見る場所) 本.push(...集める(d));
  if (本.length < 40) throw new Error('★見た 本が ' + 本.length + '本＝少なすぎ＝拾い方が 壊れて います★');
  const 見つけた = [];
  for (const f of 本) {
    if (免除.has(f)) continue;
    const 当 = 直に書いている(素にする(fs.readFileSync(path.join(ROOT, f), 'utf8')));
    for (const c of 当) 見つけた.push(f + ':' + c.行 + '  ' + c.形 + '  ' + c.字);
  }
  console.log('      … ★見た 本 ' + 本.length + '本（' + 見る場所.join('／') + '）'
    + '／免除 ' + 免除.size + '本／★直書き ' + 見つけた.length + '個★');
  if (見つけた.length) {
    throw new Error('★直書きが 残って います★＝★`表.置く(マス, 値)` へ 移して ください★\n        '
      + 見つけた.join('\n        '));
  }
});

T('★⑪ 免除に 挙げた 本が 本当に 在る★（免除が 腐って いないか）', () => {
  for (const f of 免除) {
    if (!fs.existsSync(path.join(ROOT, f))) throw new Error('★免除 ' + f + ' が 無い★');
  }
});

if (process.argv.includes('--self-test')) {
  console.log('\n[shiki-hyou-kata-no-kuchi --self-test] ★わざと 壊したら 赤に なるか★');

  T('★⑫ 探す 形が ㋐（直に）を 捕まえる★', () => {
    const 当 = 直に書いている('表.中身[マス].値 = { 型: "字", 値: "1" };');
    if (当.length !== 1 || 当[0].形 !== '㋐直に') throw new Error('捕まえて いない … ' + JSON.stringify(当));
  });

  T('★★⑬ 探す 形が ㋑（変数に 受けて 逃げる）も 捕まえる★★', () => {
    const 当 = 直に書いている('const m = 表.中身[マス];\nm.値 = { 型: "字", 値: "1" };');
    if (当.length !== 1 || 当[0].形 !== '㋑受けて') throw new Error('★変数に 受けたら 逃げられます★ … '
      + JSON.stringify(当));
  });

  T('★⑭ 読むだけ／比べるだけは 赤に しない★（要らない 赤を 出さない）', () => {
    const 読 = '表.中身[a].打った字 || ""';
    const 比 = 'if (表.中身[a].値 === undefined) {}';
    const 別 = 'var o = g.中身[j].物;';
    for (const s of [読, 比, 別]) {
      const 当 = 直に書いている(s);
      if (当.length) throw new Error('★要らない 赤★ … ' + s + ' → ' + JSON.stringify(当));
    }
  });

  T('★⑮ `置く` が 無ければ 赤に なる（口が 消えたら 気づける）', () => {
    if (typeof H.表().置く !== 'function') throw new Error('★`置く` が 無い★');
  });

  T('★⑯ 棚に ⑲が 書いて 在る★', () => {
    const 棚 = fs.readFileSync(path.join(ROOT, 'docs/measured/karimono-hazushi-no-tana.md'), 'utf8');
    if (!/表\.置く/.test(棚)) throw new Error('★棚に 口の 名前（表.置く）が 無い★');
    if (!/直書き 0個/.test(棚)) throw new Error('★棚に 数えた 結果（直書き 0個）が 無い★');
  });
}

console.log('\n' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
