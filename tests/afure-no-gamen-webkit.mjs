/* afure-no-gamen-webkit.mjs — ★新しい 本に 溢れる 式を 打った 時、画面の 字が 実Excel と 同じか★（2026-10-05）
 *
 *  ★★物差し★★ 経営者の 道具137 の 紙 golden-jitsu-excel-no-afure-2026-10-05.tsv（日本語の Excel 16.0.20430・新しい 本・57行・24組）
 *  ★★測る 道具★★ docs/measured/hakaru-afure-no-gamen.mjs（お客さんの 道 setCell・打つ 順は 道具137 の 組の 定義・待ちは 条件）
 *  ★★見る 物★★
 *    ① 既知の 違い（下の 名簿）★以外の 違いが 0★（画面の 字・描いた 字の 両方）
 *    ② 待ちの 条件が 全部の 組で 立った（計算し直しの 前を 数えて いない）
 *  ★★既知の 違い★★（まだ 直して いない＝消えたら 名簿から 外す。★足すのは 測って 訳を 書いた 時だけ★）
 *    G・O の F1 ･･･ `=SUM(E1#)`（溢れを 丸ごと 見る # の 参照）を 台が 知らない
 *    P #NULL! ･･･ `=SUM(A1 B1)`（交わらない）が #VALUE!
 *  ★--self-test★ ･･･ 道具の --壊す（#SPILL! を #スピル! に 換える 所を 素通し）で ★実Excel が #スピル! の 行の 数だけ★ 赤に なるか
 *  使い方: node tests/afure-no-gamen-webkit.mjs [--self-test]
 */
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { borrow } from '../scripts/_borrow-playwright.mjs';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const 自己 = process.argv.includes('--self-test');
/* ★既知の 違い★＝★0★（2026-10-05 に G・O の E1# と P・Q の 交わりを 直して 67/67）。足すのは 測って 訳を 書いた 時だけ */
const 既知 = [];

console.log('[afure-no-gamen] ★溢れる 式の 画面の 字が 実Excel と 同じか★' + (自己 ? '（--self-test＝#スピル! に 換える 所を 壊す）' : ''));
/* ★先に 借りる★＝毎回の CI（webkit 無し）では 借り方が 未測定の 声を 出して 緑で 終わる（ほかの webkit の 見張りと 同じ） */
const wk = await borrow('afure-no-gamen', 'webkit');
if (!wk) process.exit(0);
const r = spawnSync(process.execPath, [path.join(ROOT, 'docs/measured/hakaru-afure-no-gamen.mjs'), '--台', 'webkit'].concat(自己 ? ['--壊す'] : []), { cwd: ROOT, encoding: 'utf8' });
const 出 = (r.stdout || '') + (r.stderr || '');
console.log(出.split('\n').map((l) => '    ' + l).join('\n'));
const 数 = /★画面の 字 (\d+)\/(\d+) ／ 描いた 字 (\d+)\/(\d+)★/.exec(出);
const 違い = 出.split('\n').filter((l) => /^ {5}\S/.test(l)).map((l) => l.trim());
const 名簿外 = 違い.filter((l) => !既知.some((k) => l.indexOf(k + ' ') === 0));
const 消えた既知 = 既知.filter((k) => !違い.some((l) => l.indexOf(k + ' ') === 0));
const 条件 = 出.indexOf('条件が 立たない') < 0 && 出.indexOf('条件 立たず') < 0;
let 赤 = 0;
const T = (n, ok, m) => { if (ok) console.log('  ok   ' + n); else { 赤++; console.log('  NG   ' + n + (m ? '\n       ' + m : '')); } };
T('★道具の 出しが 読める（行数が 出た）★', !!数, '出しが 読めない');
T('★待ちの 条件が 全部の 組で 立った★', 条件);
T('★既知 以外の 違いが 0★（既知 ' + 既知.length + '）', 名簿外.length === 0, 名簿外.join('\n       '));
T('★描いた 字の 違いの 数＝画面の 字の 違いの 数★', !!数 && 数[1] === 数[3], 数 ? '画面 ' + 数[1] + ' ／ 描いた ' + 数[3] : '');
if (消えた既知.length) console.log('  ★既知の 違いが 消えた（名簿から 外して ください）★ ' + 消えた既知.join(' ／ '));
T('★既知の 違いが 全部 まだ 出て いる（直ったら 名簿を 減らす）★', 消えた既知.length === 0, 消えた既知.join(' ／ '));
console.log('\nafure-no-gamen' + (自己 ? ' --self-test' : '') + ': ' + (赤 ? '赤 ' + 赤 : '緑'));
/* ★自己試験は 赤に なる のが 正しい★＝違いの 数が ★紙で 実Excel が #スピル! の 行の 数★と 同じか（壊した 数と 赤の 数を 並べる） */
if (自己) {
  const fs = await import('node:fs');
  const 紙 = fs.readFileSync(path.join(ROOT, 'docs/measured/golden-jitsu-excel-no-afure-2026-10-05.tsv'), 'utf8').split(/\r?\n/).filter((l) => l && l.charAt(0) !== '#');
  const 頭 = 紙.shift().split('\t'), i字 = 頭.indexOf('画面の字(.Text)');
  const 欲しい = 紙.filter((l) => l.split('\t')[i字] === '#スピル!').length;
  console.log('  壊した 時の 違い ' + 名簿外.length + ' ／ 紙で #スピル! の 行 ' + 欲しい);
  process.exit(欲しい > 0 && 名簿外.length === 欲しい ? 0 : 1);
}
process.exit(赤 ? 1 : 0);
