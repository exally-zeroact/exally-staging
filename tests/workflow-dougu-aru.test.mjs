/* workflow-dougu-aru.test.mjs — ★見張りが 呼ぶ 道具が repo に 在るか★（2026-09-16）
 *
 *  ★★なぜ 要るか★★
 *    `.github/workflows/source-urls.yml` が ★2026-09-14 から ずっと 赤★でした。
 *    訳は `Cannot find module '…/kyuyo/scripts/check-source-urls.mjs'`
 *    ＝★`kyuyo/` は この repo から 出て 行って いた★（`git ls-files 'kyuyo/**'` … 0本）
 *    ⇒★★見張りの 顔を して 1行も 見て いなかった★★
 *
 *  ★★「赤だから 気づく」は 嘘でした★★
 *    ・週1の 見張りは ★次に 走るまで 色が 変わりません★
 *    ・★ずっと 赤★＝★本物の 欠陥が 入っても 色は 動かない★
 *    ⇒記憶「材料が 無いのに 定時で 走り 緑を 返す物は 止めろ」の ★赤 版★
 *      ＝★赤でも 同じ 害★
 *    ⇒★★週1では なく ★毎回 走る 総なめ★の 中で 見ます★★
 *
 *  ★★この 見張りが 見る 物★★
 *    ①yml が 読めて いる（★空振りして いない★＝分母を 出さない 緑は 嘘）
 *    ②★yml が 呼ぶ 道具が 1つ 残らず repo に 在る★
 *    ③★yml の 数が 黙って 減って いない★
 *
 *  ★外した 見張りの 記録★ … `docs/hazushita-mihari.md`
 *
 *  使い方: node tests/workflow-dougu-aru.test.mjs
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
let pass = 0, fail = 0;
const T = (n, f) => {
  try { f(); pass++; console.log('  ✓ ' + n); }
  catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + e.message); }
};

console.log('\n[workflow-dougu-aru] ★見張りが 呼ぶ 道具が repo に 在るか★');

const 道 = path.join(ROOT, '.github/workflows');
const ymlたち = fs.readdirSync(道).filter((f) => /\.ya?ml$/.test(f));

/* ★★呼び出しを 拾う★★
     `run:` の 中の `node xxx.mjs` / `python xxx.py` / `bash xxx.sh`
     ★`|` の 複数行も 拾えます★（yml 全体を 字として 見る） */
const 拾う = (s) => {
  const 出 = [];
  const re = /(?:^|[\s;&|])(?:node|python3?|bash|sh)\s+(--[\w-]+\s+)*([A-Za-z0-9_][A-Za-z0-9_./-]*\.(?:mjs|cjs|js|py|sh))/g;
  let m;
  while ((m = re.exec(s))) 出.push(m[2]);
  return 出;
};

const 呼び出し = [];   /* {yml, 道具} */
for (const y of ymlたち) {
  const s = fs.readFileSync(path.join(道, y), 'utf-8');
  for (const d of 拾う(s)) 呼び出し.push({ yml: y, 道具: d });
}

const 道具ら = [...new Set(呼び出し.map((r) => r.道具))];
console.log('      … yml ' + ymlたち.length + '本 ／ 呼び出し ' + 呼び出し.length + '件 ／ 道具 ' + 道具ら.length + '個');

T('★yml が 読めて いる（★空振りして いない★）★', () => {
  if (!ymlたち.length) throw new Error('★yml が 1本も 無い★');
  if (呼び出し.length < 10) {
    throw new Error('★呼び出しが ' + 呼び出し.length + '件しか 拾えない★（拾い方が 壊れて いる）');
  }
  /* ★必ず 在る はずの 物を 名指しで 確かめる★＝拾い方が 痩せたら ここで 止まる */
  for (const n of ['tests/run.js', 'scripts/stamp-build.mjs']) {
    if (道具ら.indexOf(n) < 0) throw new Error('★' + n + ' を 拾えて いない★');
  }
});

T('★★yml が 呼ぶ 道具が 1つ 残らず repo に 在る★★', () => {
  const 無 = [];
  for (const r of 呼び出し) {
    if (!fs.existsSync(path.join(ROOT, r.道具))) 無.push(r.yml + ' → ' + r.道具);
  }
  if (無.length) {
    throw new Error('★' + 無.length + '件 無い★ … ' + 無.join(' ／ ')
      + '／★その 見張りは 1行も 見て いません★（外すなら docs/hazushita-mihari.md に 書く）');
  }
});

T('★yml の 数が 黙って 減って いない★', () => {
  /* ★数は 決め打ち★＝★減らしたら ここも 直す＝直さないと 赤で 止まる★
       2026-09-16 … source-urls.yml を 外して 5 → ★4★
         （訳は `docs/hazushita-mihari.md`。向こう（rakually）で 週1 緑で 走って います） */
  const 期待 = 4;
  if (ymlたち.length !== 期待) {
    throw new Error('★yml が ' + ymlたち.length + '本★（' + 期待 + '本の はず）'
      + ' … ' + ymlたち.join(' ') + '／足したなら ここも 直す');
  }
});

console.log('\n' + pass + ' passed, ' + fail + ' failed');
if (fail) process.exit(1);
