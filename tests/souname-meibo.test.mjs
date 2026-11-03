/* souname-meibo.test.mjs — ★tests/ の 試験が 1本 残らず 走って いるか★（2026-09-16）
 *
 *  ★★なぜ 要るか（★私が 嘘の 報告を しました★）★★
 *    2026-09-16 の commit `586d2eb` で
 *    「★試験の 登録も 自動で 通りました（224本・登録漏れ 0件）★」と 報告しました。
 *    ★間違いです★。`tests/xlfn-morenashi.test.mjs` は
 *      ・`tests/run.js` の 名簿（FILES）に ★無い★
 *      ・`.github/workflows/*.yml` からも ★呼ばれて いない★
 *    ⇒★★試験の 顔を して 1回も 走って いませんでした★★
 *
 *  ★★なぜ 気づけなかったか★★
 *    `scripts/tests-registered.mjs` は ★playwright の 一覧★を 見て います。
 *    ★`tests/run.js` の FILES は 見て いません★。
 *    ⇒★「登録漏れ 0件」は ★別の 物★を 数えた 緑でした★
 *    ⇒記憶「★何を 何で 数えたか を 数の 隣に 書く★」を 踏みました。
 *
 *  ★★この 見張りが 見る 物★★
 *    ①名簿が 読めて いる（★空振りして いない★）
 *    ②★`tests/` の 試験が 1本 残らず ★名簿 か yml の どちらか★で 走る★
 *    ③★名簿に 在るのに ファイルが 無い 物が 無い★（★名簿だけ 残った 幽霊★）
 *
 *  ★同じ 型★（★手で 並べた 名簿は 必ず 漏れる★）
 *    ・PERMUTATIONA（08-02）／RANK.AVG（08-01）… `_xlfn.` の 一覧
 *    ・`_xlfn.` 91個（09-16）
 *    ・`toru-oufuku-shinki.ps1` の 見る マス（09-16・26本 押して 23本しか 見て いなかった）
 *    ・★この 4本（09-16）★
 *
 *  使い方: node tests/souname-meibo.test.mjs
 */
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(path.join(ROOT, 'package.json'));

let pass = 0, fail = 0;
const T = (n, f) => {
  try { f(); pass++; console.log('  ✓ ' + n); }
  catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + e.message); }
};

console.log('\n[souname-meibo] ★tests/ の 試験が 1本 残らず 走って いるか★');

/* ★名簿は 写さない＝`tests/run.js` 本人から 読む★ */
const { FILES } = require_(path.join(ROOT, 'tests/run.js'));
/* ★★名簿の 形を 先に 見る★★（2026-09-18）
     ★訳★ ... 名簿の 行末の カンマが ★覚書きの 中に 入る★と その 行は undefined に なります
     ⇒★★本数は 減らないのに 1本 走りません★★（★黙って 見逃す★）
     ⇒★2026-09-18 に ★同じ 日に 2回★ 踏みました★
     ⇒★下の map は undefined で 転びます＝★転ぶ 前に 名指しで 止めます★★ */
const 形が変 = FILES.map((f, i) => [i, f])
  .filter(([, f]) => !(Array.isArray(f) ? typeof f[0] === 'string' && f[0] : typeof f === 'string' && f));
if (形が変.length) {
  console.log('');
  console.log('  NG   ★★名簿に 形の 変な 行が ' + 形が変.length + '本★★（★1本 走りません★）');
  for (const [i, f] of 形が変) {
    console.log('       ・' + i + '番目 ... ' + JSON.stringify(f)
      + '（前 ' + JSON.stringify(FILES[i - 1]) + '）');
  }
  console.log('       ★行末の カンマが 覚書きの 中に 入って いませんか★');
  process.exit(1);
}
const 名簿 = new Set(FILES.map((f) => (Array.isArray(f) ? f[0] : f).split(/[\\/]/).pop()));

/* ★yml が 直に 呼ぶ 物も 「走って いる」に 数える★ */
const yml道 = path.join(ROOT, '.github/workflows');
const yml字 = fs.readdirSync(yml道).filter((f) => /\.ya?ml$/.test(f))
  .map((f) => fs.readFileSync(path.join(yml道, f), 'utf-8')).join('\n');

/* ★★「走らせない」と 決めて ある 物は 免除します★★（`tests-no-ci.json`）
     ★なぜ★ 2026-09-16 に 私は `hozon-de-kieru.test.mjs` を 名簿に 足して
       ★総なめを 赤に しました★。これは ★わざと 赤のまま★と
       2026-09-07（commit 20ac3f6）に 決めて あった 物です。
     ⇒★★見張りが 「決めて ある 免除」を 知らないと
         ★直す 必要の 無い 物を 直させようと します★★
     ★免除は 全部 ★訳（why）と 戻す 条件（back）★ 付き★＝★黙って 見逃しません★ */
const 免除紙 = JSON.parse(fs.readFileSync(path.join(ROOT, 'tests-no-ci.json'), 'utf-8'));
const 免除 = new Map();
for (const k of Object.keys(免除紙)) 免除.set(k.split('/').pop(), 免除紙[k]);

const 歩く = (d, 前) => {
  let 出 = [];
  for (const e of fs.readdirSync(path.join(ROOT, d), { withFileTypes: true })) {
    if (e.name === 'node_modules') continue;
    const p = 前 ? 前 + '/' + e.name : e.name;
    if (e.isDirectory()) 出 = 出.concat(歩く(d + '/' + e.name, p));
    else if (/\.(test|spec)\.(mjs|cjs|js)$/.test(e.name)) 出.push(p);
  }
  return 出;
};
const 在る = 歩く('tests', '');

console.log('      … 名簿 ' + FILES.length + '本 ／ tests/ に 在る 試験 ' + 在る.length + '本');

T('★名簿が 読めて いる（★空振りして いない★）★', () => {
  if (FILES.length < 100) throw new Error('★名簿が ' + FILES.length + '本しか 読めない★');
  if (在る.length < 100) throw new Error('★tests/ から ' + 在る.length + '本しか 拾えない★');
  /* ★必ず 在る はずの 物を 名指しで★＝拾い方が 痩せたら ここで 止まる */
  for (const n of ['shiki-kansuu.test.mjs', 'stamp.test.mjs']) {
    if (!名簿.has(n)) throw new Error('★名簿に ' + n + ' が 無い★（読み方が 壊れて いる）');
  }
});

T('★★tests/ の 試験が 1本 残らず 走って いる（名簿 か yml）★★', () => {
  const 走らない = [];
  for (const f of 在る) {
    const 名 = f.split('/').pop();
    if (名簿.has(名)) continue;
    if (yml字.indexOf(名) >= 0) continue;   /* ★yml が 直に 呼ぶ★ */
    if (免除.has(名)) continue;            /* ★走らせないと 決めて ある★ */
    走らない.push(f);
  }
  if (走らない.length) {
    throw new Error('★' + 走らない.length + '本 どこからも 呼ばれて いない★ … ' + 走らない.join(' ')
      + '／★試験の 顔を して 1回も 走りません★（tests/run.js の FILES に 足す）');
  }
});

T('★★免除は 全部 訳つき（★黙って 見逃さない★）★★', () => {
  /* ★訳が 無い 免除は ★ただ 止めて いるだけ★に なります★ */
  const 訳なし = [];
  for (const [名, v] of 免除) {
    if (!v || !v.why || !v.back) 訳なし.push(名);
  }
  if (訳なし.length) {
    throw new Error('★' + 訳なし.length + '本 訳（why）か 戻す 条件（back）が 無い★ … ' + 訳なし.join(' '));
  }
  console.log('      … 免除 ' + 免除.size + '本（全部 訳つき）');
});

T('★免除した 物は 名簿に 入って いない（★両方に 在ると 走って しまう★）★', () => {
  /* ★これで 2026-09-16 の 間違い（hozon-de-kieru を 名簿に 足した）が 赤に なります★ */
  const 両方 = [...免除.keys()].filter((n) => 名簿.has(n));
  if (両方.length) {
    throw new Error('★' + 両方.length + '本 ★走らせない★と 決めた 物が 名簿に 在る★ … ' + 両方.join(' ')
      + '／★tests-no-ci.json の 訳を 読む★');
  }
});

T('★名簿に 在るのに ファイルが 無い 物が 無い（★幽霊★）★', () => {
  const 幽霊 = [];
  for (const f of FILES) {
    const p = Array.isArray(f) ? f[0] : f;
    if (!fs.existsSync(path.join(ROOT, 'tests', p))) 幽霊.push(p);
  }
  if (幽霊.length) {
    throw new Error('★' + 幽霊.length + '本 名簿に 在るのに ファイルが 無い★ … ' + 幽霊.join(' '));
  }
});

console.log('\n' + pass + ' passed, ' + fail + ' failed');
if (fail) process.exit(1);
