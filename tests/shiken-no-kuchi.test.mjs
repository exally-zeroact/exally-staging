/* shiken-no-kuchi.test.mjs — ★実の倉庫に触る道具の門「試験用の口か」が 正しく 止めるか★（2026-10-10）
 *  門＝メールの +タグ が 白名簿（e0test）に 在る 時だけ 通す（tests/shiken-no-kuchi.mjs）。
 *  作り物の 箱だけで 当てる（本物の メールは ここにも 書かない）。
 */
import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import crypto from 'node:crypto';
const blobのid = (p) => { const b = fs.readFileSync(p); return crypto.createHash('sha1').update('blob ' + b.length + '\0').update(b).digest('hex'); };
import { 試験の口か, 許すタグ } from './shiken-no-kuchi.mjs';

let 緑 = 0, 赤 = 0;
const T = (n, ok) => { if (ok) { 緑++; console.log('  ok   ' + n); } else { 赤++; console.log('  NG   ' + n); } };
console.log('[shiken-no-kuchi] ★試験用の 口か の 門★');
T('★白名簿は e0test だけ★', 許すタグ.length === 1 && 許すタグ[0] === 'e0test');
T('通す：mihon+e0test@example.com', 試験の口か('mihon+e0test@example.com'));
const 止める = [
  ['タグ 無し', 'mihon@example.com'],
  ['別の タグ', 'mihon+honban@example.com'],
  ['タグの 一部だけ 同じ', 'mihon+e0test2@example.com'],
  ['タグが 大文字', 'mihon+E0TEST@example.com'],
  ['+ が 2つ', 'mihon+x+e0test@example.com'],
  ['空白 入り', 'mihon +e0test@example.com'],
  ['@ が 無い', 'mihon+e0test'],
  ['ドメインに 点が 無い', 'mihon+e0test@example'],
  ['空', ''],
  ['undefined', undefined],
];
for (const [n, e] of 止める) T('止める：' + n, !試験の口か(e));

/* ★メールが repo に 戻って いないか★＝git ls-files の 字の ファイルの メールの ドメインは 作り物の 白名簿 だけ
 *  （10-10 本番前の 対立役＝戻っても 赤に なる 試験が 無かった。sha256 は 置かない＝白名簿で 見る） */
const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
/* users.noreply.github.com＝GitHub の 表に 出さない 用（受信箱で ない）・noreply@github.com＝GitHub 自身・noreply@anthropic.com＝Claude の 共作の 印（10-11） */
const 許すドメイン = new Set(['example.com', 'test.com', 'users.noreply.github.com']);
const 許す宛先 = new Set(['noreply@github.com', 'noreply@anthropic.com']);
/* 全席 共通の 物を 逐語で 写した 物＝自己確認の 作り物の メール（架空の ドメイン）を 持つ。
   ★道と 中身（blob の id）が 写した 版と 同じ 時だけ★ 外す（中身が 変われば 見る・10-11 本番前の 対立役）
   禁止の 字の 見張り（10-11）・commit の メールの 門（10-20 の 6回目＝.githooks/ を 追跡に した 時に 赤に なる と 対立役が 測った） */
const 逐語の写し = {
  'tests/kinshi-ji.test.mjs': 'f4aa3cebe77f8783d9cc991f905f02409310d65e',
  '.githooks/pre-push-author.mjs': '669910aae9202ea128ea356ef7bc8fd34085091b',
};
const 字の本 = execFileSync('git', ['-C', ROOT, 'ls-files'], { encoding: 'utf8' }).split('\n')
  .filter((f) => f && !/\.(png|jpe?g|gif|ico|xlsx|xlsb|xlsm|pdf|woff2?|ttf)$/i.test(f) && !/\.min\.js$/.test(f)
    && !(Object.hasOwn(逐語の写し, f) && blobのid(path.join(ROOT, f)) === 逐語の写し[f]));
const 外 = [];
for (const f of 字の本) {
  let s; try { s = fs.readFileSync(path.join(ROOT, f), 'utf8'); } catch (e) { continue; }
  for (const m of s.matchAll(/[A-Za-z0-9._%+-]+@([A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)*\.[A-Za-z]{2,})\b/g)) {
    if (!許すドメイン.has(m[1].toLowerCase()) && !許す宛先.has(m[0].toLowerCase())) 外.push(f + '（ドメイン ' + m[1].length + '字）');
  }
}
T('★メールの ドメインは 作り物（example.com・test.com）と GitHub の noreply だけ（' + 字の本.length + '本を 見た）★', 外.length === 0 && 字の本.length >= 500, 外.slice(0, 5).join(' / '));

/* ★実の 倉庫に 触る 道具が 門を 呼んで いるか★（呼ぶ 行を 消しても 赤に なる 試験が 無かった） */
for (const f of ['tests/live-seed.mjs', 'tests/live-roundtrip.mjs']) {
  const s = fs.readFileSync(path.join(ROOT, f), 'utf8');
  T('★' + f + ' が 試験の口を読む と 試験の口か を 呼ぶ★', /from '\.\/shiken-no-kuchi\.mjs'/.test(s) && /試験の口を読む\(\)/.test(s) && /試験の口か\(/.test(s));
}
console.log('\nshiken-no-kuchi: ' + 緑 + ' 緑 / ' + 赤 + ' 赤');
process.exit(赤 ? 1 : 0);
