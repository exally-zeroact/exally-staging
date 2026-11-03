/* pre-push-noreply.test.mjs — ★押す 前の 作者の 門（scripts/pre-push-noreply.sh）が 止める 物を 止めるか★（2026-10-11）
 *  一時の git の 中で、作り物の メール（example.com）と noreply の commit を 作り、門に 当てる。
 *  ★止めた 時に メールの 字を 出さない★ことも 見る。
 */
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { spawnSync, execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const 門 = path.join(ROOT, 'scripts', 'pre-push-noreply.sh');
let 緑 = 0, 赤 = 0;
const T = (n, ok, m) => { if (ok) { 緑++; console.log('  ok   ' + n); } else { 赤++; console.log('  NG   ' + n + (m ? '\n       ' + m : '')); } };
console.log('[pre-push-noreply] ★押す 前の 作者の 門★');

const 置き場 = fs.mkdtempSync(path.join(os.tmpdir(), 'pre-push-noreply-'));
const 遠く = path.join(置き場, 'remote.git'), 別 = path.join(置き場, 'other.git'), 手元 = path.join(置き場, 'w');
const NR = '1+mihon@users.noreply.github.com', 偽 = 'mihon@example.com';
const 基 = Object.assign({}, process.env, { GIT_CONFIG_NOSYSTEM: '1' });
const git = (args, env) => execFileSync('git', args, { cwd: 手元, encoding: 'utf8', env: Object.assign({}, 基, env || {}) }).trim();
const 人 = (a, c) => ({ GIT_AUTHOR_NAME: 'm', GIT_AUTHOR_EMAIL: a, GIT_COMMITTER_NAME: 'm', GIT_COMMITTER_EMAIL: c });
const 作る = (名, 作者, 記録者, 本文) => {
  fs.writeFileSync(path.join(手元, 名), 名);
  git(['add', 名]);
  git(['commit', '-q', '-m', 名 + (本文 ? '\n\n' + 本文 : '')], 人(作者, 記録者));
  return git(['rev-parse', 'HEAD']);
};
const 当てる = (sha, remote) => spawnSync('sh', [門, remote || 'origin'], { cwd: 手元, encoding: 'utf8', input: 'refs/heads/main ' + sha + ' refs/heads/main 0000000000000000000000000000000000000000\n' });
const 出し = (r) => (r.stdout || '') + (r.stderr || '');
try {
  execFileSync('git', ['init', '-q', '--bare', 遠く]);
  execFileSync('git', ['init', '-q', '--bare', 別]);
  execFileSync('git', ['init', '-q', '-b', 'main', 手元]);
  git(['remote', 'add', 'origin', 遠く]);
  git(['remote', 'add', 'other', 別]);
  /* 遠くに 在る commit は わざと noreply で ない＝見たら 赤に なる（空振り 止め） */
  const a = 作る('a', 偽, 偽);
  git(['push', '-q', '--no-verify', 'origin', 'main']);
  const b = 作る('b', NR, NR);
  let r = 当てる(b);
  T('★遠く（origin）に 在る commit は 見ない＝noreply の b だけ 見て 通す★', r.status === 0 && /見た commit 1 本/.test(r.stdout), 出し(r));
  git(['reset', '-q', '--hard', a]);
  const c = 作る('c', 偽, NR);
  r = 当てる(c);
  T('★作者が noreply で ない ⇒ 止める★', r.status === 1 && /の 作者/.test(r.stderr), 出し(r));
  T('★止めた 時に メールの 字を 出さない★', !出し(r).includes(偽) && !出し(r).includes('example.com'));
  /* 別の remote に だけ 在る commit を origin へ 押す ⇒ 見る（止める） */
  git(['push', '-q', '--no-verify', 'other', 'HEAD:refs/heads/x']);
  git(['fetch', '-q', 'other']);
  r = 当てる(c, 'origin');
  T('★別の remote に 在っても 押す 先に 無ければ 見る★', r.status === 1, 出し(r));
  git(['reset', '-q', '--hard', a]);
  const d = 作る('d', NR, 偽);
  r = 当てる(d);
  T('★commit した 人が noreply で ない ⇒ 止める★', r.status === 1 && /の commit した 人/.test(r.stderr), 出し(r));
  git(['reset', '-q', '--hard', a]);
  const e = 作る('e', NR, 'noreply@github.com', 'Co-authored-by: Claude <noreply@anthropic.com>');
  r = 当てる(e);
  T('★GitHub 自身と Claude の 印は 通す★', r.status === 0, 出し(r));
  git(['reset', '-q', '--hard', a]);
  const f = 作る('f', NR, NR, 'Co-authored-by: m <' + 偽 + '>');
  r = 当てる(f);
  T('★本文の Co-authored-by に noreply で ない メール ⇒ 止める★', r.status === 1 && /Co-authored-by/.test(r.stderr), 出し(r));
  git(['reset', '-q', '--hard', a]);
  const g = 作る('g', 'mihon@users.noreply.github.com', NR);
  r = 当てる(g);
  T('★noreply の 形を 装った 字（数字+口 で ない）⇒ 止める★', r.status === 1, 出し(r));
  git(['reset', '-q', '--hard', a]);
  const h = 作る('h', '1+MIHON@USERS.NOREPLY.GITHUB.COM', '2+mihon[bot]@users.noreply.github.com');
  r = 当てる(h);
  T('★大文字の noreply・bot の noreply は 通す★', r.status === 0, 出し(r));
  git(['tag', '-a', 't1', '-m', 't', h], { GIT_COMMITTER_NAME: 'm', GIT_COMMITTER_EMAIL: 偽 });
  const タグ = git(['rev-parse', 't1']);
  r = spawnSync('sh', [門, 'origin'], { cwd: 手元, encoding: 'utf8', input: 'refs/tags/t1 ' + タグ + ' refs/tags/t1 0000000000000000000000000000000000000000\n' });
  T('★注釈付きの タグの tagger が noreply で ない ⇒ 止める★', r.status === 1 && /タグの tagger/.test(r.stderr), 出し(r));
  r = spawnSync('sh', [門, 'origin'], { cwd: 手元, encoding: 'utf8', input: '(delete) 0000000000000000000000000000000000000000 refs/heads/x ' + a + '\n' });
  T('★消す 押しは 通す（見る commit 無し）★', r.status === 0, 出し(r));
  r = spawnSync('sh', [門, 'origin'], { cwd: 手元, encoding: 'utf8', input: 'refs/heads/main deadbeefdeadbeefdeadbeefdeadbeefdeadbeef refs/heads/main 0000000000000000000000000000000000000000\n' });
  T('★並べられない sha は 止める（黙って 通さない）★', r.status === 1, 出し(r));
} finally { fs.rmSync(置き場, { recursive: true, force: true }); }
console.log('\npre-push-noreply: ' + 緑 + ' 緑 / ' + 赤 + ' 赤');
process.exit(赤 ? 1 : 0);
