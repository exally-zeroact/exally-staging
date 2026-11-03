/* machi-no-kime.test.mjs — ★待ちの 決め(timeout)が 本当に 効いて いるか★ 2026-09-24
 *
 *  ★★何が 在ったか★★
 *    経営者1 が 私の 測り道具で 気づきました（2026-09-24）。
 *      ★字には `{ timeout: 300000 }` と 書いて ある のに★
 *      ★出しは `Timeout 30000ms exceeded`★
 *
 *  ★★実物で 割りました★★（★人の 言を 根拠に しない★）
 *    絶対に 真に ならない 関数に `{ timeout: 1000 }` を 渡して 何秒で 落ちるかを 測った:
 *      2つ目の 席に 決めを 置く ......... ★30,006 ms★（`Timeout 30000ms exceeded`）
 *      `null` を 挟んで 3つ目に 置く ... ★ 1,009 ms★（`Timeout 1000ms exceeded`）
 *      （★呼び名を ここに 字で 書きません★＝★この 門が 自分の 覚書きを 拾って しまう★）
 *    ⇒★2つ目の 席は 「関数に 渡す 引数」★。決めは ★3つ目★。
 *    ⇒2つ目に 決めを 置くと ★黙って 既定の 30秒★ に なる。
 *
 *  ★★なぜ 門に するか★★
 *    ★落ちない★＝★誰も 気づかない★。
 *      ・書いた 人は 「5分 待つ」と 思って いる
 *      ・実際は 30秒で 落ちる
 *      ・★遅い 本を 「開けない 本」と 言い間違える★（2026-09-24 に 実際 そう 見えた）
 *    ＝★紙に 書いた 数が 効いて いない★ 型。★字を 見るしか 捕まえ方が ない★。
 *
 *  ★★字の 形を 見る 門で よい 訳★★
 *    普段は ★字の 形を 見る 門は 相手が 言葉を 直した 日に 割れる★ ので 避ける。
 *    ここは ★守る 相手が 「言葉」では なく 「呼び方(引数の 席)」★ なので、
 *    ★字の 形こそが 中身★。
 *
 *  走らせ方: node tests/machi-no-kime.test.mjs [--self-test]
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const 壊す = process.argv.includes('--self-test');
const 置き場 = ['tests', 'docs/measured', 'scripts'];
/* ★決めを 3つ目に 置く 呼び方★（2つ目は 引数） */
const 決めが3つ目 = ['waitForFunction'];

let pass = 0, fail = 0;
const T = (n, ok, m) => {
  if (ok) { pass++; console.log('  ok   ' + n); }
  else { fail++; console.log('  NG   ' + n + (m ? '\n       ' + m : '')); }
};

/** ★呼び出しの 引数を 一番 外の カンマで 割る★
 *  ＝`{ timeout: 1000 }` の 中の カンマで 割らない（★括弧の 深さを 数える★） */
function 引数に割る(字, 頭) {
  const k = 字.indexOf('(', 頭);
  if (k < 0) return null;
  let 深 = 0, p = k;
  for (; p < 字.length; p++) {
    const c = 字[p];
    if (c === '(') 深++;
    else if (c === ')') { 深--; if (深 === 0) break; }
  }
  if (深 !== 0) return null;
  const 中 = 字.slice(k + 1, p);
  const 区 = [];
  let 深2 = 0, 始 = 0;
  for (let q = 0; q < 中.length; q++) {
    const c = 中[q];
    if (c === '(' || c === '[' || c === '{') 深2++;
    else if (c === ')' || c === ']' || c === '}') 深2--;
    else if (c === ',' && 深2 === 0) { 区.push(中.slice(始, q)); 始 = q + 1; }
  }
  区.push(中.slice(始));
  return { 区: 区, 終: p };
}

/** ★2つ目の 席に 決めが 座って いる 所を 拾う★ */
export function 座り間違いを拾う(字) {
  const 出 = [];
  for (const 名 of 決めが3つ目) {
    let i = 0;
    for (;;) {
      const j = 字.indexOf(名 + '(', i);
      if (j < 0) break;
      const w = 引数に割る(字, j);
      if (!w) break;
      if (w.区.length === 2 && /\btimeout\s*:/.test(w.区[1])) {
        出.push({ 名: 名, 行: 字.slice(0, j).split('\n').length, 字: w.区[1].trim().slice(0, 40) });
      }
      i = w.終;
    }
  }
  return 出;
}

console.log('[machi-no-kime] ★待ちの 決めが 本当に 効いて いるか★');

/* ══ ★数える 前に 「どこを 数えたか」を 出す★ ══ */
const 見た = [];
for (const d of 置き場) {
  const 所 = path.join(ROOT, d);
  if (!fs.existsSync(所)) continue;
  for (const 名 of fs.readdirSync(所)) {
    if (!/\.(mjs|js)$/.test(名)) continue;
    見た.push(path.join(d, 名));
  }
}
const 呼ぶ本 = 見た.filter((f) => fs.readFileSync(path.join(ROOT, f), 'utf8').indexOf('waitForFunction(') >= 0);
console.log('      ── 実測 ── 見た ' + 見た.length + '本（' + 置き場.join(' / ') + '）'
  + ' ／ `waitForFunction` を 呼ぶ ' + 呼ぶ本.length + '本');

const 悪い = [];
for (const f of 呼ぶ本) {
  for (const x of 座り間違いを拾う(fs.readFileSync(path.join(ROOT, f), 'utf8'))) {
    悪い.push(f + ':' + x.行 + '  ' + x.名 + '(関数, ' + x.字 + ')');
  }
}
T('★決めが 2つ目に 座って いる 所が 0件★（2つ目は 引数・決めは 3つ目）',
  悪い.length === 0,
  悪い.length ? ('★' + 悪い.length + 'か所★\n       ' + 悪い.join('\n       ')
    + '\n       ⇒ `}, { timeout: N })` を `}, null, { timeout: N })` に する') : '');

T('★見る 所が 空に なって いない★（門が 何も 見なく なるのが 一番 こわい）',
  呼ぶ本.length > 0, '呼ぶ本 ' + 呼ぶ本.length + '本');

/* ══ ★わざと 壊して 赤に なるか★ ══ */
if (壊す) {
  console.log('\n★わざと 壊して 赤に なるか★');
  /* ★★見本の 字は その場で 組み立てます★★（2026-09-24 ここで 1回 踏みました）
       最初は 見本を そのまま 字で 書いたら ★この 門が 自分の 見本を 拾って 赤★に なりました。
       ⇒★自分を 免除しません★（免除の 訳は 外して 測るまで 見立て）。
       ⇒★★代わりに 見本の 中に 呼び名を 字で 書かない★★＝その場で 繋ぎます。
       ⇒★門は 1本も 見逃さず、自分の 見本にも 引っかかりません★ */
  const W = 'waitFor' + 'Function';
  const 組 = [
    ['2つ目に 決め（★悪い★）', 'await page.' + W + '(() => window.x, { timeout: 60000 });', 1],
    ['null を 挟む（良い）', 'await page.' + W + '(() => window.x, null, { timeout: 60000 });', 0],
    ['決めを 渡さない（良い）', 'await page.' + W + '(() => window.x);', 0],
    ['1行で 2つ目に 決め（★悪い★）', 'await p.' + W + "(() => typeof window.f === 'function', { timeout: 30000 });", 1],
    ['★中に カンマが 在る 関数★（良い）',
      'await page.' + W + '((a, b) => a + b > 0, null, { timeout: 5000 });', 0],
    ['★決めの 中に カンマ★（★悪い★・割り方を 間違えると 見逃す）',
      'await page.' + W + '(() => window.x, { timeout: 60000, polling: 100 });', 1],
    ['2つ目が ただの 引数（良い）', 'await page.' + W + '((n) => window.x > n, 5);', 0],
  ];
  for (const [名, 字, 期待] of 組) {
    const 出 = 座り間違いを拾う(字);
    T('  ' + 名 + ' ⇒ ' + 期待 + '件', 出.length === 期待, '出た ' + 出.length + '件');
  }

  /* ══ ★★見本では なく 本物の 本を 壊して 確かめる★★ ══（2026-09-24）
       ★★なぜ 見本だけでは 足りないか★★
         見本は ★私が 書いた 字★ です。
         ⇒★私が 思って いる 形しか 出て きません★。
         ⇒★★2026-09-24 に 実際に 割れたのは 本物の 本の 中でした★★（10本・18か所）。
       ★★だから 本物を 読んで、その場で 直しを 巻き戻します★★
         ＝★手元の ファイルは 1バイトも 触りません★（読んだ 字の 上だけ）
         ＝★触って いない 事も 最後に 数えます★ */
  {
    const 本 = path.join(ROOT, 'tests/doko-wo-sawatta-bun-webkit.mjs');
    const 前 = fs.readFileSync(本, 'utf8');
    const 今 = 座り間違いを拾う(前);
    T('  ★本物の 本は 今 0件★（直して あるので）', 今.length === 0, '出た ' + 今.length + '件');

    /* ★直しを 巻き戻す★＝`null` を 抜いて 09-24 より 前の 形に 戻す */
    const 壊した = 前.replace(/\}, null, \{ timeout:/g, '}, { timeout:');
    const 出 = 座り間違いを拾う(壊した);
    T('  ★★巻き戻したら 赤に なる★★（★壊れて いるのに 緑★ では ない）',
      出.length > 0, '出た ' + 出.length + '件');

    /* ★★「壊したのに 赤に ならない」の 裏★★＝★本当に 壊れたか★ を 先に 見る */
    T('  ★巻き戻しが 本当に 効いて いる★（字が 変わって いる）',
      壊した !== 前, '字が 1バイトも 変わって いない＝★巻き戻せて いない★');

    T('  ★★手元の ファイルを 触って いない★★（読んだ 字の 上だけ）',
      fs.readFileSync(本, 'utf8') === 前, '★書き換わって います★');
  }
}

console.log('\n' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
