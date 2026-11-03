/* ci-same.mjs — ★押す前に「CI と 同じ 物」を 手元で 走らせる★ 2026-09-05
 *
 *  ★なぜ 作ったか（同じ日に 2回 踏んだ）★
 *    ①手元の ★未commit★ を 見て「緑です」と 言った（CI は 押した 物しか 見ない）
 *    ②手元で ★tests/run.js の 202本★だけ 走らせて「全部 緑」と 言った
 *       ⇒ CI は ★それ以外に 直に 走らせる 段★が 在り、そこで 赤に なった
 *    ⇒★どちらも「CI と 同じ 物を 見ていない」★
 *    ⇒★決意では 直らない★＝★1本の 命令に する★
 *
 *  ★やる事★
 *    ・.github/workflows/ci.yml の `run:` を ★全部 拾う★
 *    ・★その通りに 走らせる★（省かない・並べ替えない）
 *    ・★走らせた 段の 数を 出す★＝★0段 走って 緑★を 見破る
 *    ・★赤は 全部 出す★（1本目で 止めない＝1回で 全部 分かる）
 *
 *  ★なぜ ci.yml だけを 読むか★
 *    ワークフローは 全部で 5本 在るが、★他の 4本は 外を 叩く／週１の 重い 回★
 *    （実配信を 押す／ブラウザを 落とす／夕方の 見張り）
 *    ⇒★押す前に 走らせては いけない 物★＝ここでは 見ない
 *    （見たい 時は --all）
 *
 *  ★これは CI の 代わりでは ありません★
 *    ・OS も node の 版も ちがう（CI は ubuntu / node 20）
 *    ・★押した 物★では なく ★手元の 物★を 見ている
 *    ⇒★押す前の 下見★です。★最後は CI が 正★。
 *
 *  使い方: node scripts/ci-same.mjs [--list]
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { spawnSync } from 'node:child_process';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const WF = path.join(ROOT, '.github/workflows');

/** ワークフローから `run:` の 中身を 拾う（★1行の 形と 複数行の 形の 両方★） */
function 段を拾う(file) {
  const src = fs.readFileSync(file, 'utf8').split('\n');
  const 出 = [];
  for (let i = 0; i < src.length; i++) {
    const m = /^(\s*)run:\s*(\|)?\s*(.*)$/.exec(src[i]);
    if (!m) continue;
    const 字下げ = m[1].length;
    if (m[2]) {                                  /* run: | の 形＝下の 行が 中身 */
      const 中 = [];
      for (let j = i + 1; j < src.length; j++) {
        if (src[j].trim() === '') { 中.push(''); continue; }
        const 今 = src[j].length - src[j].replace(/^\s*/, '').length;
        if (今 <= 字下げ) break;
        中.push(src[j].trim());
      }
      const s = 中.join('\n').trim();
      if (s) 出.push(s);
    } else if (m[3].trim()) {
      出.push(m[3].trim());
    }
  }
  return 出;
}

/** ★手元で 走らせられない 段★（理由つき・★黙って 飛ばさない★） */
const 飛ばす = [
  { 印: 'actions/', 訳: 'GitHub の 部品（手元には 無い）' },
  { 印: 'npm install', 訳: '入れるだけ（手元は 既に 入っている）' },
  { 印: 'npm ci', 訳: '同上' },
  { 印: 'npx playwright install', 訳: 'ブラウザを 落とすだけ（週1の 回）' },
];

const files = fs.existsSync(WF)
  ? fs.readdirSync(WF).filter((f) => /\.ya?ml$/.test(f)).sort() : [];
if (!files.length) {
  console.log('★ワークフローが 1本も 無い＝この 道具が 空振り★');
  process.exit(1);
}

/* ★どのワークフローを 見るか★＝既定は CI だけ（週1の 回は ブラウザが 要る） */
const 対象 = process.argv.includes('--all') ? files : files.filter((f) => f === 'ci.yml');
const 段 = [];
for (const f of 対象) for (const c of 段を拾う(path.join(WF, f))) 段.push({ wf: f, cmd: c });

console.log('');
console.log('[ci-same] ★CI と 同じ 物を 手元で 走らせる★');
console.log('  見た ワークフロー … ' + 対象.join('・') + '（全部で ' + files.length + '本 在る）');
console.log('  拾った 段 ……… ' + 段.length + '段');

if (process.argv.includes('--list')) {
  段.forEach((s, i) => console.log('  ' + String(i + 1).padStart(2) + ' [' + s.wf + '] ' + s.cmd.replace(/\n/g, ' ; ')));
  process.exit(0);
}

/** ★走らせる 貝殻を 先に 決める★（★2026-09-05 ここで 転んだ★）
 *  PowerShell から 離して 走らせたら `bash` が 道に 無く、
 *  ★19段 全部 赤（中身は 空）★に なった。
 *  ⇒★中身が 空の 赤は「試験が 落ちた」では なく「命令が 始まっていない」★
 *  ⇒★見分けの つかない 赤を 出さない★＝先に 貝殻を 探し、無ければ その場で 止まる。 */
function 貝殻を探す() {
  if (process.platform !== 'win32') return 'sh';
  const 候補 = ['bash',
    'C:/Program Files/Git/bin/bash.exe',
    'C:/Program Files/Git/usr/bin/bash.exe',
    'C:/Program Files (x86)/Git/bin/bash.exe'];
  for (const c of 候補) {
    const t = spawnSync(c, ['-lc', 'echo ok'], { encoding: 'utf8' });
    if (!t.error && t.status === 0) return c;
  }
  return null;
}
const 貝殻 = 貝殻を探す();
if (!貝殻) {
  console.log('');
  console.log('★bash が 見つからない＝1段も 走らせられません（道具の 不調・試験の 赤では ない）★');
  console.log('  Git for Windows の bash.exe に 道を 通してから もう一度。');
  process.exit(2);
}
console.log('  使う 貝殻 ……… ' + 貝殻);

let 走った = 0, 赤 = [], 飛 = [];
for (const s of 段) {
  const 理由 = 飛ばす.find((x) => s.cmd.indexOf(x.印) >= 0);
  if (理由) { 飛.push({ cmd: s.cmd, 訳: 理由.訳 }); continue; }
  走った++;
  const r = spawnSync(貝殻, ['-lc', s.cmd],
    { cwd: ROOT, encoding: 'utf8', maxBuffer: 64 * 1024 * 1024 });
  /* ★命令が 始まらなかった 時は 赤に しない★＝止まる（黙って 赤に 混ぜない） */
  if (r.error) {
    console.log('  ★道具の 不調★ ' + s.cmd.slice(0, 70) + '  （' + r.error.code + '）');
    console.log('★命令を 始められませんでした＝試験の 赤では ありません★');
    process.exit(2);
  }
  const ok = r.status === 0;
  console.log('  ' + (ok ? 'ok  ' : '★赤★') + ' ' + s.cmd.replace(/\n/g, ' ; ').slice(0, 92));
  if (!ok) {
    var 全 = (r.stdout || '') + (r.stderr || '');
    /* ★一言も 言わずに 落ちた 赤★は 中身が 分からない＝そう 書く（推し量らない） */
    /* ★ただ 末尾を 切るな★＝2026-09-05 実測。tests/run.js が 出す
       『deprecated』の 注意書きが 14行を 埋め、★本当の 赤の 理由が 押し出された★。
       ⇒★赤らしい 行を 先に 拾い★、それでも 無ければ 末尾を 見せる。 */
    var 行 = 全.split('\n');
    /* ★ゆるく Error や failed で 拾うと 逆に 見えなくなる★＝2026-09-05 実測。
       ★わざと 失敗させる 試験★が『Claude API error: …』を 正しく 出しており、
       それが 20行を 埋めて ★本当に 落ちた 試験の 名前が 押し出された★。
       ⇒★試験の 出し手が『これが 赤だ』と 言っている 行だけ★を 拾う。 */
    var 目印 = /(★落ちた★|ファイルで失敗|^\s*(✗|✘|NG |not ok))/;
    var 拾い = 行.filter(function (x) { return 目印.test(x); });
    var 中 = 全.trim()
      ? (拾い.length
          ? 拾い.slice(-20).join('\n') + '\n' + '  …（★赤だと 言っている 行★だけ 抜き出しました）'
          : 行.slice(-20).join('\n'))
      : '★何も 言わずに 落ちました（終了の 印 ' + r.status + '）＝道具の 側かも しれません★';
    赤.push({ cmd: s.cmd, out: 中 });
  }
}

console.log('');
console.log('  ★走らせた ' + 走った + '段 ／ 赤 ' + 赤.length + '段 ／ 飛ばした ' + 飛.length + '段★');
for (const x of 飛) console.log('    飛ばした … ' + x.cmd.slice(0, 60) + '  （' + x.訳 + '）');
if (!走った) { console.log('★1段も 走っていない＝この 道具が 空振り★'); process.exit(1); }
for (const x of 赤) {
  console.log('');
  console.log('★赤★ ' + x.cmd);
  console.log(x.out);
}
console.log('');
console.log('ci-same: ' + (走った - 赤.length) + ' 緑 / ' + 赤.length + ' 赤'
  + '（★これは 下見です。押した 物を 見るのは CI★）');
process.exit(赤.length ? 1 : 0);
