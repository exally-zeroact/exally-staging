/* phonetic-ramuda.test.mjs -- PHONETIC と 裸の LAMBDA が 実Excel と 同じか（2026-09-19）
 *
 *  ★なぜ 在るか★
 *    ★分母が 溢れの 左上しか 見て いなかった★ので、この 2つは ずっと 出ませんでした。
 *    ⇒ 経営者1 が 実Excel に 聞いた（㊴）⇒ `docs/measured/golden-phonetic-lambda-2026-09-19.tsv`
 *    ⇒ `=PHONETIC(A1:A2)` が うちは ★#VALUE!★／実Excel は ★空の 字★ だと 分かった
 *
 *  ★この 門が 見る 物★
 *    ・紙の 6本を ★台（lib/shiki-hyou.js）★に 押して ★出る 字★で 比べる
 *    ・★分母を 必ず 出す★（6本）／★増えても 減っても 赤★
 *
 *  ★この 門が 言えない 事★
 *    ・★ふりがなの 入った ブックを 読んだ 時★は 見て いません
 *      ＝`=PHONETIC(B1)` が ヤマダ を 返すには xlsx の `<rPh>` を 読む 要が 在る
 *      ＝★画面に ふりがなを 付ける 口が 無い★ので まだ 測れません
 *      ⇒ だから ★B1 の 行は この 門から 外して あります（名指し）★
 *    ・画面では ありません（台だけ）
 *
 *  使い方: node tests/phonetic-ramuda.test.mjs
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(import.meta.url);
let pass = 0, fail = 0;
const T = (n, よい, 添え) => {
  if (よい) { pass++; console.log('  ok   ' + n); }
  else { fail++; console.log('  NG   ' + n + (添え ? '\n       ' + 添え : '')); }
};

console.log('');
console.log('[phonetic-ramuda] PHONETIC と 裸の LAMBDA');

const 紙の道 = path.join(ROOT, 'docs/measured/golden-phonetic-lambda-2026-09-19.tsv');
T('★紙が 在る★', fs.existsSync(紙の道), 紙の道);

const NL = String.fromCharCode(10), TAB = String.fromCharCode(9);
const 行たち = fs.readFileSync(紙の道, 'utf-8').split(NL)
  .filter((l) => l && l.charAt(0) !== '#')
  .map((l) => l.split(TAB))
  .filter((p) => p[1] && p[1].charAt(0) === '=')
  .map((p) => ({ 式: p[1], 出る字: p[3] === undefined ? '' : p[3] }));

/* ★ふりがなの 入った マスは まだ 測れない＝名指しで 外す★ */
const 外す = { '=PHONETIC(B1)': 'xlsx の <rPh> を まだ 読んで いない／画面に ふりがなを 付ける 口が 無い' };
const 押す行 = 行たち.filter((r) => !外す[r.式]);

T('★★紙の 分母が 6本★★（★形が 変わったら 赤★）', 行たち.length === 6, '今 ' + 行たち.length + '本');
T('★★押す 分母が 5本★★（★1本だけ 名指しで 外して いる★）', 押す行.length === 5, '今 ' + 押す行.length + '本');

const SH = require_(path.join(ROOT, 'lib/shiki-hyou.js'));
const 板 = SH.表();
板.打つ('A1', 1); 板.打つ('A2', 2); 板.打つ('B1', '山田');

const 外れ = [];
for (const r of 押す行) {
  let 出;
  try { 板.打つ('ZZ9986', r.式); 出 = String(板.字('ZZ9986')); }
  catch (e) { 出 = '（投げた: ' + (e && e.message) + '）'; }
  if (出 !== r.出る字) 外れ.push(r.式 + ' ... うち ' + JSON.stringify(出) + ' ／実Excel ' + JSON.stringify(r.出る字));
}
T('★★5本 とも 実Excel と 同じ 字★★（★減っても 増えても 赤★）', 外れ.length === 0,
  外れ.join('\n       '));

/* ★★わざと 壊した 時に 赤に なるか★★（★門が 本当に 見て いるか★） */
{
  const にせ = 押す行.map((r) => (r.式 === '=PHONETIC(A1:A2)' ? { 式: r.式, 出る字: '#VALUE!' } : r));
  let 見つけた = 0;
  for (const r of にせ) {
    板.打つ('ZZ9985', r.式);
    if (String(板.字('ZZ9985')) !== r.出る字) 見つけた++;
  }
  T('★★紙の 字を 1つ 変えたら 赤に なる★★（★門が 見て いる 証し★）', 見つけた === 1,
    '見つけた ' + 見つけた + '本（1本の はず）');
}

console.log('      ... 押した ' + 押す行.length + '本 ／ 名指しで 外した ' + (行たち.length - 押す行.length) + '本');
console.log('');
console.log('phonetic-ramuda: ' + pass + ' 緑 / ' + fail + ' 赤');
process.exit(fail ? 1 : 0);
