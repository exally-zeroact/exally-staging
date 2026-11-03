/* js-sou-sekisho.test.mjs — ★JS層に ★関所を 通らない 段★を 書かせない★
 *
 *  ★★何を 守るか★★
 *    `_jsComputeFormula` の 入口に 関所が 在る:
 *      if(!_jsSet[_fnBase]) return null;   // JS非対象 → HFへ
 *    ⇒★`_jsSet` に 載っていない 関数の 段は ★1度も 届きません★＝★死にコード★
 *    ⇒ 2026-09-08 に ★12か所★在り、★12個 とも 押して「届かない」を 確かめて★ 消した
 *      （訳と 名簿 … docs/measured/shini-code-2026-09-08.md）
 *
 *  ★★なぜ コメントでは なく 見張りか（指示役 2026-09-08）★★
 *    ⇒★コメントは「また 書き戻される」を ★止められません★★
 *    ⇒★★見張りは 止めます★★
 *    ⇒ これは ISOMITTED で 作った「★名簿に 無ければ 積まない★」と 同じ 形
 *
 *  ★★死にコードが なぜ 悪いか（実際に 起きた）★★
 *    ⇒ 私も 指示役も ★「JS層は 38個」★と 数えて 話していました
 *    ⇒ 本当は ★26個★（12個は 死にコード）
 *    ⇒★★数え間違いの 元に なります★★
 *
 *  使い方: node tests/js-sou-sekisho.test.mjs [--self-test]
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const ここ = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(ここ, '..');
const 直に走った = process.argv[1] && pathToFileURL(process.argv[1]).href === import.meta.url;
const 自己試験 = 直に走った && process.argv.includes('--self-test');

const { 注記を外す } = await import(
  pathToFileURL(path.join(ROOT, 'scripts/lib/chuki.mjs')).href);

let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ✓ ' + n); } catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + (e && e.message)); } };

/** ★段を 数える★（★字を 読む 所を 先に 名指しする★） */
function 数える(生) {
  /* ★★注記を 外してから 関所を 読む（2026-09-09 に 直した）★★
     ★前は 生の まま★読んで いました。
     ⇒ 関所の 注記に `=MIRR(A1:A3,-1,0.12)` と 書いたら
       ★`A1` を 関数名として 拾い「関所に 在るのに 段が 無い」と 赤★に なりました
     ⇒★注記の 字を 数えて いた＝見張りの 側の 欠陥★
     ⇒★注記外しは 自前で 書かない★＝共通の 部品を 使う
       （自前だと URL の // や 字の中の /* を 間違える。tests/chuki.test.mjs が 自前を 赤に する） */
  生 = 注記を外す(生);
  const m = 生.match(/var _jsSet\s*=\s*\{[\s\S]*?\};/);
  if (!m) throw new Error('★_jsSet（関所）が 見つからない★＝この 見張りは 何も 見ていない');
  const 通す = new Set([...m[0].matchAll(/([A-Z][A-Z0-9]*)\s*:/g)].map((x) => x[1]));
  if (!通す.size) throw new Error('★関所が 空★');

  const 頭 = 生.indexOf('function _jsComputeFormula');
  if (頭 < 0) throw new Error('★_jsComputeFormula が 見つからない★');
  const 行 = 生.slice(頭).split('\n');

  const 生きて = [], 死んで = [];
  for (const l of 行) {
    const mm = l.match(/fOrig\.match\(\/\^\\?\(?([A-Z][A-Z0-9._|]*)/);
    if (!mm) continue;
    const 名たち = mm[1].split('|').filter((n) => /^[A-Z][A-Z0-9._]*$/.test(n));
    if (!名たち.length) continue;
    /* ★関所は `.` の 前で 見る★（PERCENTILE.INC → PERCENTILE） */
    (名たち.some((n) => 通す.has(n.split('.')[0])) ? 生きて : 死んで).push(名たち.join('|'));
  }
  return { 通す, 生きて, 死んで, 行数: 行.length };
}

const 生 = fs.readFileSync(path.join(ROOT, 'exally-formula.js'), 'utf-8');

console.log('\n★JS層の 関所★');

T('★★関所を 通らない 段が 1つも 無い★★', () => {
  const r = 数える(生);
  console.log('      見た 範囲 … _jsComputeFormula の 中 ' + r.行数 + '行');
  console.log('      関所が 通す 関数 ' + r.通す.size + '個 ／ 生きている 段 ' + r.生きて.length + 'か所');
  if (r.死んで.length) {
    throw new Error('★関所を 通らない 段が ' + r.死んで.length + 'か所★ … ' + r.死んで.join(' ')
      + '\n      ⇒★1度も 届きません（死にコード）／数え間違いの 元に なります★'
      + '\n      ⇒ 関所（_jsSet）に 足すか、段を 消して ください');
  }
});

T('★関所に 居るのに 段が 無い 関数が 無い★', () => {
  /* ★逆向きも 見る★＝名簿に 在るのに 横取りしない＝どちらも ずれの 元 */
  const r = 数える(生);
  const 段の名 = new Set(r.生きて.flatMap((s) => s.split('|').map((n) => n.split('.')[0])));
  const 無し = [...r.通す].filter((n) => !段の名.has(n)).sort();
  if (無し.length) {
    throw new Error('★関所に 在るのに 段が 無い★ … ' + 無し.join(' ')
      + '\n      ⇒★名簿と 中身が ずれています★');
  }
});

T('★消した 物が 本当に 消えている★', () => {
  /* ★2026-09-08 に 消した 下請け★＝コードに 1個も 残っていない
     （★他の 11個は 消していません★＝プラグイン登録／AGGREGATE／他の 下請けから 呼ばれる） */
  const 消した = ['_jsXlookup'];
  const 残り = 消した.filter((n) => new RegExp('\\b' + n + '\\b').test(生));
  if (残り.length) throw new Error('★消したはずの 物が 残っている★ … ' + 残り.join(' '));
});

if (自己試験) {
  console.log('\n[self-test] ★物差しが 効いているか★');
  T('★★関所を 通らない 段を 1つ 足すと 赤に なる★★', () => {
    /* ★壊すのは ★写し★★＝repo の ファイルは 1バイトも 触らない */
    const 悪い = 生.replace('function _jsComputeFormula',
      'function _jsComputeFormula_DUMMY(){ var mZZ=fOrig.match(/^ZZTOP\\s*\\(/i); }\nfunction _jsComputeFormula');
    let 赤 = false;
    try { 数える(悪い); } catch (e) { 赤 = true; }
    if (!赤) {
      const r = 数える(悪い);
      if (r.死んで.length) 赤 = true;
    }
    console.log('      … 関所に 無い 段（ZZTOP）を 足した ⇒ ' + (赤 ? '★赤★' : '素通り'));
    if (!赤) throw new Error('★素通りした★＝この 見張りは 何も 見ていない');
  });
  T('★★関所が 読めない 時に 黙って 緑に しない★★', () => {
    const 悪い = 生.replace(/var _jsSet\s*=\s*\{/, 'var _jsSet_KESHITA = {');
    let 赤 = false;
    try { 数える(悪い); } catch (e) { 赤 = true; }
    console.log('      … 関所を 消した ⇒ ' + (赤 ? '★止まる★' : '素通り'));
    if (!赤) throw new Error('★関所が 読めないのに 緑に した★＝『無い』を『済んだ』と 読んでいる');
  });
}

console.log('\n' + pass + ' passed, ' + fail + ' failed');
if (直に走った) process.exit(fail ? 1 : 0);
