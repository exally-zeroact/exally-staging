/* shiki-hyou-hishigata.test.mjs — ★菱形の 頼りで 古い 答えが 残らないか★（2026-09-15）
 *
 *  ★★なぜ 要るか（実物で 見つけた）★★
 *    司さんの 実物（代行計算表2026.xlsb）を 自前の 土台だけで 押したら
 *    ★押した 8,019本の うち 37本が 実Excel と 合いません★でした。
 *    ★しかも 押す順を 逆に すると 176本★（★共通は 6本だけ★）
 *    ＝★合わない マスが 押す順で 変わる★＝★追いかけが 1周で 届いて いない★
 *
 *  ★★正体＝菱形★★
 *      K1 = 10
 *      Q1 = `=K1*2`      … Q は K を 見る
 *      M1 = `=K1-Q1`     … ★M は K も Q も 見る★
 *    ⇒ K が 変わると ★M と Q の 両方★に 報せが 行く
 *    ⇒ ★M を 先に 直すと 古い Q を 使う★
 *    ⇒ その後 Q を 直しても ★M は「済」印が 付いて いて もう 直らない★
 *    ★前の `直す` は「先に 着いた 順」で 直して いました★（幅優先＋済み印）
 *
 *  ★★直し★★ … ★自分が 見て いる 物を 全部 直してから 自分を 直す★（頼りの 順）
 *    ・★数え上げで 回す★（O(マス＋線)）＝★2乗に しない★
 *      （`tests/shiki-hyou-omosa.test.mjs` が 2乗を 赤に します）
 *    ・★輪は 順番では 解けない★ ⇒ 残った 物は そのまま 直す（輪の 見つけ方は `計算する` が 持つ）
 *
 *  ★★直した 後（実物で 測った）★★
 *    ★順に 押して 0本／逆に 押して 0本★（前は 37本／176本）
 *    ⇒★★押す順に 依らなく なりました★★＝実Excel と 同じ 振る舞い
 *
 *  ★★この 見張りが 守る 事★★
 *    ①★菱形で 古い 答えが 残らない★（3通り）
 *    ②★押す順を 変えても 同じ 答え★
 *    ③★輪は 今までどおり 見つける★（★直しで 輪の 見分けを 壊して いない★）
 *
 *  使い方: node tests/shiki-hyou-hishigata.test.mjs
 *          node tests/shiki-hyou-hishigata.test.mjs --self-test
 */
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(path.join(ROOT, 'package.json'));
const H = require_(path.join(ROOT, 'lib/shiki-hyou.js'));

let pass = 0, fail = 0;
const T = (n, f) => {
  try { f(); pass++; console.log('  ✓ ' + n); }
  catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + e.message); }
};

console.log('\n[shiki-hyou-hishigata] ★菱形の 頼りで 古い 答えが 残らないか★');

/* ★菱形★ K → Q ／ K → M ／ Q → M */
function 菱形(順) {
  const h = H.表();
  const 手 = {
    K1: '10',
    Q1: '=K1*2',
    M1: '=K1-Q1',
  };
  for (const n of 順) h.打つ(n, 手[n]);
  return h;
}

T('★★① 建てる 順が どれでも 同じ 答え★★（菱形 3通り）', () => {
  for (const 順 of [['K1', 'Q1', 'M1'], ['M1', 'Q1', 'K1'], ['Q1', 'M1', 'K1']]) {
    const h = 菱形(順);
    const m = String(h.字('M1')), q = String(h.字('Q1'));
    if (q !== '20') throw new Error(順.join('→') + ' … Q1 が ' + q + '（20 のはず）');
    if (m !== '-10') throw new Error(順.join('→') + ' … ★M1 が ' + m + '（-10 のはず）★'
      + '＝★古い Q1 を 使って います★');
  }
});

T('★★② 材料を 打ち直しても 古い 答えが 残らない★★', () => {
  for (const 順 of [['K1', 'Q1', 'M1'], ['M1', 'Q1', 'K1']]) {
    const h = 菱形(順);
    h.打つ('K1', '100');
    const m = String(h.字('M1')), q = String(h.字('Q1'));
    if (q !== '200') throw new Error(順.join('→') + ' … Q1 が ' + q + '（200 のはず）');
    if (m !== '-100') throw new Error(順.join('→') + ' … ★M1 が ' + m + '（-100 のはず）★'
      + '＝★Q1 を 直した 後に M1 を 直して いません★');
  }
});

T('★③ 深い 菱形（2段 重ねても 残らない）★', () => {
  const h = H.表();
  h.打つ('A1', '=B1-C1');      /* A は B と C を 見る */
  h.打つ('B1', '=D1*2');       /* B は D */
  h.打つ('C1', '=B1+D1');      /* ★C は B も D も 見る★ */
  h.打つ('D1', '5');
  /* D=5 → B=10 ／ C=10+5=15 ／ A=10-15=-5 */
  if (String(h.字('B1')) !== '10') throw new Error('B1 … ' + h.字('B1'));
  if (String(h.字('C1')) !== '15') throw new Error('C1 … ' + h.字('C1'));
  if (String(h.字('A1')) !== '-5') throw new Error('★A1 が ' + h.字('A1') + '（-5 のはず）★');
  h.打つ('D1', '50');
  /* D=50 → B=100 ／ C=150 ／ A=-50 */
  if (String(h.字('A1')) !== '-50') throw new Error('★打ち直し後 A1 が ' + h.字('A1') + '（-50 のはず）★');
});

T('★④ 輪は 今までどおり 見つける★（直しで 壊して いない）', () => {
  const h = H.表();
  h.打つ('A1', '=B1+1');
  h.打つ('B1', '=A1+1');
  const a = String(h.字('A1'));
  if (a !== '#REF!') throw new Error('★輪を 見つけて いない★ … A1 = ' + a);
});

T('★⑤ 菱形と 輪が 混ざっても 落ちない★', () => {
  const h = H.表();
  h.打つ('K1', '10');
  h.打つ('Q1', '=K1*2');
  h.打つ('M1', '=K1-Q1');
  h.打つ('Z1', '=Y1+1');
  h.打つ('Y1', '=Z1+1');       /* ★輪★ */
  h.打つ('K1', '100');          /* ★輪が 在る 盤で 直す★ */
  if (String(h.字('M1')) !== '-100') throw new Error('M1 … ' + h.字('M1'));
  if (String(h.字('Z1')) !== '#REF!') throw new Error('Z1 … ' + h.字('Z1'));
});

if (process.argv.includes('--self-test')) {
  console.log('\n[shiki-hyou-hishigata --self-test] ★わざと 壊したら 赤に なるか★');

  T('★★「先に 着いた 順」に 戻したら 赤に なる★★', () => {
    /* ★壊すのは 写し★＝ファイルは 1バイトも 触りません
       ★前の 作り（幅優先＋済み印）を ここで 作り直して 同じ 答えに なるか 見ます★ */
    const 見ている = { M1: ['K1', 'Q1'], Q1: ['K1'], K1: [] };
    const 見られている = { K1: { M1: true, Q1: true }, Q1: { M1: true } };
    /* ★前の やり方★＝着いた 順に 1回だけ */
    const 順 = [];
    const 待ち = Object.keys(見られている.K1), 済 = {};
    while (待ち.length) {
      const n = 待ち.shift();
      if (済[n]) continue;
      済[n] = true;
      順.push(n);
      for (const x of Object.keys(見られている[n] || {})) if (!済[x]) 待ち.push(x);
    }
    /* ★M が Q より 先に 来たら 古い 答えに なります★ */
    const mi = 順.indexOf('M1'), qi = 順.indexOf('Q1');
    if (mi < 0 || qi < 0) throw new Error('組み方が おかしい');
    if (mi > qi) throw new Error('★この 並びでは 壊れません＝試しに なって いない★');
    console.log('      … 前の やり方だと ' + 順.join('→') + '＝★M が Q より 先★＝古い 答えに なる');
  });

  T('★実物で 直った 数を 覚えて いる★（37／176 → 0／0）', () => {
    const 棚 = require_('node:fs').readFileSync(path.join(ROOT, 'docs/measured/karimono-hazushi-no-tana.md'), 'utf8');
    /* ★探す 字は 棚の 書き方に 合わせる★（★字を 決めた 時点で 答えが 決まる★ので
       ★数そのもの★を 探します＝言い回しが 変わっても 当たる） */
    if (!/37本/.test(棚)) throw new Error('★棚に 直す前の 数（37本）が 書いて いない★');
    if (!/176本/.test(棚)) throw new Error('★棚に 直す前の 数（176本）が 書いて いない★');
    if (!/菱形/.test(棚)) throw new Error('★棚に 正体（菱形）が 書いて いない★');
  });
}

console.log('\n' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
