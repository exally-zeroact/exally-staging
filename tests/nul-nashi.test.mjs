/* nul-nashi.test.mjs — ★生の NUL を 配る物に 入れない★（2026-09-07）
 *
 *  ★★何が 起きていたか（実測）★★
 *    ファイルの ★先頭8000バイト★に ★生の NUL★が 1つ 在ると、
 *    git は その ファイルを ★絵（binary）★と 決めます。
 *    ⇒★★`git diff` が ★空★に なる＝中で 何を しても 見えない★★
 *    ⇒ CI は 緑・PR の 差分は 空白 ⇒★変えた人も 見る人も 気づけない★
 *    ★実測（2026-09-07・550本 中）★
 *      lib/flash-fill.js（2個）／lib/ref-graph.js（1個）／lib/ribbon.js（1個）
 *      tests/pivot.test.mjs（1個） …★この 4本が 絵 扱い★
 *      scripts/check-parity.mjs（1個・★8889バイト目★）…★今は 字 扱い★
 *        ⇒★★でも ファイルが 縮んだ 日から ★黙って 差分が 消えます★＝時限式★★
 *
 *  ★★もっと 悪い 事★★
 *    ★見張りは 既に 在りました★（`tests/pivot.test.mjs`）。
 *    ⇒★★でも その 見張り自身が 生の NUL を 書いて ★自分を 絵に して★ いました★★
 *    ⇒★★しかも 見ていたのは `lib/pivot.js` ★1本だけ★★（550本 中／そこには 無い＝ずっと 緑）
 *    ⇒★『見張りが 在る』は『守られている』では ない★
 *    ⇒★★見張りを 見つけたら 3つ 見る★★
 *        ①見る範囲を 数える ②守りたい物が 入っているか ★名指し★
 *        ③★その 見張り自身が 同じ 病気に かかっていないか★
 *
 *  ★★この 紙は 生の NUL を 1つも 持ちません★★
 *    ⇒ 書く 時は ★`'\u0000'`★（逆斜線・u・0000）と ★字で★ 書く。
 *      中身は 同じ 1バイト／git からは ★字に 見える★。
 *
 *  ★★空白に 直しては いけません★★（2026-09-07 指示役が 止めた）
 *    `lib/ribbon.js` の NUL は ★区切り★です（人の 字に 絶対 出てこないから 選ばれた）。
 *    空白に すると 親='A B'／元名='C' と 親='A'／元名='B C' が ★同じ 鍵★に なります。
 *    ⇒★『見た目が 同じ』と『役目が 同じ』は 別★
 *
 *  使い方: node tests/nul-nashi.test.mjs [--self-test]
 */
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { fileURLToPath, pathToFileURL } from 'node:url';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const 直に走った = process.argv[1] && pathToFileURL(process.argv[1]).href === import.meta.url;
const 自己試験 = 直に走った && process.argv.includes('--self-test');

let pass = 0, fail = 0;
const T = (n, よい, 添え) => {
  if (よい) { pass++; console.log('  ok   ' + n); }
  else { fail++; console.log('  NG   ' + n + (添え ? '\n       ' + 添え : '')); }
};

/* ★★本当の 絵・音・ブックは 外す★★＝NUL が 在って 当たり前
   ★外した 数も 出す★（黙って 減らさない） */
const 本物の絵 = ['.png', '.jpg', '.jpeg', '.gif', '.ico', '.bmp', '.webp',
  '.woff', '.woff2', '.ttf', '.otf', '.eot',
  '.zip', '.gz', '.pdf', '.mp3', '.mp4', '.webm', '.wasm',
  '.xlsx', '.xlsm', '.xlsb', '.xls', '.docx', '.pptx'];
const 絵か = (f) => 本物の絵.some((e) => f.toLowerCase().endsWith(e));

/* ★git が 絵と 決める 境目★＝先頭 8000バイトに NUL が 在るか */
const 境目 = 8000;
const NUL = '\u0000'.charCodeAt(0);   /* ★字で 書く★＝この ファイルに 生の NUL を 入れない */

/** ★1本 読んで NUL を 数える★ */
function 数える(道) {
  let b;
  try { b = fs.readFileSync(道); } catch (e) { return null; }
  let 全 = 0, 頭 = 0;
  for (let i = 0; i < b.length; i++) {
    if (b[i] === NUL) { 全++; if (i < 境目) 頭++; }
  }
  return { 全, 頭, 大きさ: b.length };
}

/** ★見る 範囲を 先に 決めて 数える★（見張りの 決まり①） */
function 見る(files, 元) {
  const 見た = [], 外した = [], 見つけた = [];
  for (const f of files) {
    if (絵か(f)) { 外した.push(f); continue; }
    const 道 = path.join(元, f.replace(/\//g, path.sep));
    const r = 数える(道);
    if (!r) continue;
    見た.push(f);
    if (r.全) 見つけた.push({ 名: f, 全: r.全, 頭: r.頭 });
  }
  return { 見た, 外した, 見つけた };
}

console.log('');
console.log('[nul-nashi] ★生の NUL を 配る物に 入れない★');

/* ★git が 追いかけている 物を 数える★（node_modules 等は 元から 入らない） */
let 一覧 = [];
try {
  一覧 = execFileSync('git', ['ls-files'], { cwd: ROOT, encoding: 'utf-8' })
    .split('\n').map((s) => s.trim()).filter(Boolean);
} catch (e) { 一覧 = []; }

T('★★見る 範囲が 取れた（取れないのに 緑に しない）★★', 一覧.length > 100,
  'git が 追いかけている ファイル ' + 一覧.length + '本＝★少なすぎる★');

const 結 = 見る(一覧, ROOT);
console.log('       … ★見た ' + 結.見た.length + '本★ ／ 外した（本当の 絵・ブック）' + 結.外した.length + '本'
  + ' ／ 合わせて ' + 一覧.length + '本');

/* ★★守りたい 物が 見る 範囲に 入っているか 名指しで 確かめる★★（見張りの 決まり②）
   ⇒ 2026-09-07 … 前の 見張りは `lib/pivot.js` ★1本だけ★を 見ていた */
const 名指し = ['book.html', 'hub.html', 'lib/ribbon.js', 'lib/flash-fill.js', 'lib/ref-graph.js',
  'exally-formula.js', 'scripts/check-parity.mjs', 'tests/pivot.test.mjs'];
T('★★守りたい 物が 名指しで 範囲に 入っている★★', 名指し.every((f) => 結.見た.includes(f)),
  '入っていない … ' + 名指し.filter((f) => !結.見た.includes(f)).join(' '));

T('★★生の NUL を 持つ ファイルが 1本も 無い★★', 結.見つけた.length === 0,
  結.見つけた.map((x) => '  ' + x.名 + ' … NUL ' + x.全 + '個'
    + (x.頭 ? '（★先頭' + 境目 + 'の 中＝git が 絵 扱い＝差分が 見えない★）'
            : '（先頭' + 境目 + 'の 外＝今は 字 扱い／★縮んだら 消える 時限式★）')).join('\n'));

/* ★★位置で 見逃さない★★＝先頭8000の 外でも 赤に する（上の 1本で 両方 見ている）
   ⇒ ただし ★出す 字には 位置を 書く★＝直す人が
     「今 見えていないのか／いつか 消えるのか」を 分けられる 為 */

if (自己試験) {
  console.log('\n  [self-test] ★物差しが 効いているか★');
  /* ★★壊すのは ★写し★★＝repo の ファイルは 1バイトも 触らない */
  const 仮 = fs.mkdtempSync(path.join(os.tmpdir(), 'nul-'));
  try {
    const 埋めた = 5;
    const 名 = [];
    for (let i = 0; i < 埋めた; i++) {
      const f = 'utsushi-' + i + '.js';
      /* ★i 番目は 頭に／後ろの 物は 8000の 外に★＝位置の 見分けも 試す */
      const 前 = i < 3 ? 'var a = 1;' : 'x'.repeat(9000);
      fs.writeFileSync(path.join(仮, f), 前 + '\u0000' + 'var b = 2;', 'utf-8');
      名.push(f);
    }
    fs.writeFileSync(path.join(仮, 'kirei.js'), 'var c = 3;', 'utf-8');
    fs.writeFileSync(path.join(仮, 'hontou-no-e.png'), Buffer.from([0x89, 0x50, 0, 0, 0]));

    const r = 見る([...名, 'kirei.js', 'hontou-no-e.png'], 仮);
    T('★★わざと ' + 埋めた + '個 埋めたら ' + 埋めた + '個 見つかる★★'
      + '（★埋めた ' + 埋めた + '個 → 見つけた ' + r.見つけた.length + '個★）',
      r.見つけた.length === 埋めた,
      '★1個や 2個 見つかって 緑に なるなら この 見張りは 何も 見ていない★');
    T('★NUL の 無い ファイルは 拾わない★', !r.見つけた.some((x) => x.名 === 'kirei.js'));
    T('★本当の 絵は 外す（外した 数も 出す）★', r.外した.length === 1 && r.見た.length === 6,
      '外した ' + r.外した.length + '本 ／ 見た ' + r.見た.length + '本');
    const 頭にある = r.見つけた.filter((x) => x.頭 > 0).length;
    const 外にある = r.見つけた.filter((x) => x.頭 === 0).length;
    T('★位置を 見分けられる（先頭の 中 3本／外 2本）★', 頭にある === 3 && 外にある === 2,
      '中 ' + 頭にある + '本 ／ 外 ' + 外にある + '本');
    console.log('       … ★埋めた 5個 → 見つけた ' + r.見つけた.length + '個★'
      + '（先頭の 中 ' + 頭にある + '／外 ' + 外にある + '）');
  } finally {
    try { fs.rmSync(仮, { recursive: true, force: true }); } catch (e) { /* 消せなくても 手元の 仮 */ }
  }
}

console.log('');
console.log(pass + ' passed, ' + fail + ' failed');
if (直に走った) process.exit(fail ? 1 : 0);
