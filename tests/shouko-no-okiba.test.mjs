/* shouko-no-okiba.test.mjs — ★消える場所に repo の 物を 置いていないか★（2026-09-05）
 *
 *  ★なぜ 要るか（★同じ日に 2回 落ちた★）★
 *    ①2026-08-29 実Excel 507個の 一覧 ⇒ `scratchpad/exally_missing_functions.txt`
 *       ⇒★消えた★／09-05 に 聞かれて ★答えられなかった★
 *       ⇒★「507個」は ★中身の 無い 数字★に なった★
 *    ②2026-09-05 実Excel を 叩き直した YEN の 直し ⇒ ★また `scratchpad/hold2/` に 置いた★
 *       ⇒★「消えた」と 報告した ★その 30分後★★
 *    ⇒★★決意では 直らない★★＝★手を 変えるか 機械を 置くか★
 *
 *  ★形★ … 置き場の 中に ★repo が 追っている ファイルと 同じ 名前★の
 *          `.js` `.mjs` `.cjs` `.json` が 在ったら ★赤★
 *  ★控えを 取るなら★ … `git switch -c <枝>` ＋ commit（★push しなくてよい★）
 *
 *  ★★2026-09-05 指示役の 差し戻しで 直した 3つ★★
 *    ①★セッションの 番号を 焼き込んでいた★
 *       ⇒ 同じ 部屋に 置き場は ★73個★ 在るのに ★1個しか 見ていなかった★
 *       ⇒ 次の セッションでは 番号が 変わる＝★いつも 緑★
 *       ⇒★★親の 部屋を 読んで 全部 見る（番号を 書かない）★★
 *    ②★材料が 無い時に 緑を 出していた★
 *       ⇒ CI(ubuntu) に `C:/Users/…` は 無い ⇒ 消えるの=[] ⇒ ★いつも 緑★
 *       ⇒★★手元(Windows)で 置き場が 無ければ 赤／CI では「見ていない」と 毎回 名乗る★★
 *    ③★`f.split('/')[0]` は 深さ1の git複製しか 除けなかった★
 *       ⇒★どの 段の `.git` でも 部屋ごと 落とす★（掘る 前に 見る）
 *
 *  ★★どこまでを 赤に するか（2026-09-05 実測の 上で 決めた）★★
 *    親の 部屋の 置き場 ★67個★を 全部 見たら ★危ない 55本★
 *      ⇒ 55本は ★全部 他の セッション（b1c00d76 / b43c1163）★
 *      ⇒ 中身は ★別アプリの 丸ごと 複製(before-tree 318本)★＝`package.json` `sw.js`
 *        `tests/run.js` の ような ★どこにでも 在る 名前★が 当たっていた
 *    ★最後に 書かれた 時刻（実測）★
 *      私の 置き場 …………… ★0.4時間前★（生きている）
 *      b1c00d76 …………… ★80.9時間前★
 *      b43c1163 …………… ★79.9時間前★
 *    ⇒★線は 48時間★（0.4 と 80 の 間・どちらからも 遠い）
 *      ★48時間 動いていない 置き場＝そのセッションは もう 居ない★＝私には 直せない
 *      ⇒★赤に すると 一生 緑に ならない＝見張りが 切られる★
 *        （指示役 09-05「雑音 123件だったので 入れなかった」と 同じ 家）
 *      ⇒★★赤に するのは 生きている 置き場だけ／古い 置き忘れは 数を 必ず 出す★★
 *
 *  ★★この 機械は 半分しか 守りません（指示役 2026-09-05）★★
 *    ★見るのは「★控えた コード★」だけ★＝repo に ★同じ 名前★が 在る `.js/.mjs/.cjs/.json`
 *    ★捕まえられない 物★＝★測って 出した 結果★（新しい 名前の 一覧・記録）
 *      例）08-29 の 実Excel 507個 ⇒ `scratchpad/exally_missing_functions.txt`
 *          ⇒★repo に 同じ 名前の 相手が 居ない★＝★原理的に 捕まらない★
 *    ★広げては いけない★＝`.txt` まで 見ると 掃引の 記録・覚え書きが 全部 赤に なる
 *      ⇒★正しい 物を 赤に する 見張りは 人が 切る★
 *    ⇒★★測って 出した 結果は「その場で repo に commit」＝手を 変えるしか ない★★
 *
 *  使い方: node tests/shouko-no-okiba.test.mjs
 *          node tests/shouko-no-okiba.test.mjs --self-test
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ✓ ' + n); } catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + (e && e.message)); } };

/** ★純関数★＝repo の 名前の 集まりと 消える場所の 名前から「危ない物」を 返す
 *  ★git の 作業用 複製(worktree)は 除く★（2026-09-05 実測で 199本 出た）
 *    ＝★中は git が 守っている＝消えても 元が 在る★／★控えでは ない★ */
export function 危ないか(repoの名前, 消える場所の名前, git複製) {
  const 見る = /\.(js|mjs|cjs|json)$/i;
  const 集 = new Set(repoの名前.map((f) => f.split('/').pop()));
  /* ★git の 作業用 複製は ★どの 段に 在っても★ 除く★（2026-09-05 指示役の 差し戻し③）
     ★前は `f.split('/')[0]` ＝ ★深さ1しか 除けなかった★
     ⇒ 深い 所に 在ると 除けず ★199本の 雑音が 戻る★
     ★部屋の 境目まで 見る★＝`wt-a` で `wt-ab` を 巻き添えに しない */
  const 除く = (git複製 || []).map((d) => String(d).replace(/\\/g, '/').replace(/^\/+|\/+$/g, ''));
  const 除くか = (f) => 除く.some((d) => f === d || f.indexOf(d + '/') === 0);
  return 消える場所の名前.filter((f) => {
    if (!見る.test(f)) return false;
    if (除くか(f)) return false;
    return 集.has(f.split('/').pop());
  });
}

/** ★この 回を どう 名乗るか★（差し戻し②の 判定を ★純関数★に 出す）
 *  ⇒ ubuntu を 手元に 持っていないので、★枝の 中身を 直に 掴めない★
 *  ⇒★判定だけ 取り出して 自分で 壊す★（機械が 実際に 走る のは 下の 本番） */
export function この回の名乗り(親が在る, 手元か) {
  if (親が在る) return '見る';
  if (手元か) return '赤';   /* ★手元なのに 道が 無い＝道が 変わった＝赤★ */
  return '見ていない';       /* ★CI＝毎回 そう 名乗る（黙って 緑に しない）★ */
}

/* ══ 自分で 壊して 赤に なるか ═══════════════════════════════ */
if (process.argv.includes('--self-test')) {
  console.log('\n[shouko-no-okiba --self-test] わざと壊して赤になるか');
  const repo = ['lib/formula-extra.js', 'api/claude.js', 'prompt/kansuu.md'];
  T('★repo と 同じ 名前の 控えが 在れば 赤', () => {
    const r = 危ないか(repo, ['hold2/formula-extra.js']);
    if (r.length !== 1) throw new Error('見つけていない: ' + JSON.stringify(r));
  });
  T('★★掃引の 記録・絵の 元は 素通り（正しい 物を 赤に しない）★★', () => {
    const r = 危ないか(repo, ['cisame.log', 'shot/E1_mae.html', 'shot/toru.mjs', 'ai.json', 'st.json']);
    if (r.length) throw new Error('★誤検知★: ' + JSON.stringify(r));
  });
  T('★.md や .txt は 見ない（絵の 説明や 覚え書き）', () => {
    const r = 危ないか(['docs/EXCEL_NAYAMI.md'], ['EXCEL_NAYAMI.md', 'memo.txt']);
    if (r.length) throw new Error('誤検知: ' + JSON.stringify(r));
  });
  T('★深い 所に 置いても 見つける', () => {
    const r = 危ないか(repo, ['a/b/c/claude.js']);
    if (r.length !== 1) throw new Error('見つけていない');
  });
  /* ★差し戻し③の 実測★＝深さ1でしか 除けなかった 事を ここで 掴む */
  T('★★git複製は ★深さ1★でも 除く★★', () => {
    const r = 危ないか(repo, ['wt-a/lib/formula-extra.js'], ['wt-a']);
    if (r.length) throw new Error('除けていない: ' + JSON.stringify(r));
  });
  T('★★git複製は ★深い 段★でも 除く（前の 形は ここで 落ちた）★★', () => {
    const r = 危ないか(repo, ['tmp/wt-a/lib/formula-extra.js'], ['tmp/wt-a']);
    if (r.length) throw new Error('★深さ1しか 除けていない★: ' + JSON.stringify(r));
  });
  T('★除くのは 部屋の 境目まで（wt-a で wt-ab を 巻き添えに しない）', () => {
    const r = 危ないか(repo, ['tmp/wt-ab/lib/formula-extra.js'], ['tmp/wt-a']);
    if (r.length !== 1) throw new Error('★除きすぎ★: ' + JSON.stringify(r));
  });
  /* ★差し戻し②＝材料が 無い時に 緑を 出していた★ */
  T('★★材料が 無い回は「緑」に ならない（CIは 見ていないと 名乗る／手元は 赤）★★', () => {
    const a = この回の名乗り(true, true);
    const b = この回の名乗り(false, true);
    const c = この回の名乗り(false, false);
    if (a !== '見る') throw new Error('道が 在るのに 見ない: ' + a);
    if (b !== '赤') throw new Error('★手元で 道が 無いのに 赤に ならない★: ' + b);
    if (c !== '見ていない') throw new Error('★CIで 黙って 緑に している★: ' + c);
    /* ★★「いつも 緑」に なっていないか★★＝3つとも 同じ 答えなら 何も 分けていない */
    if (a === b && b === c) throw new Error('★3つとも 同じ＝何も 分けていない★');
  });
  console.log('\n' + pass + ' passed, ' + fail + ' failed');
  process.exit(fail ? 1 : 0);
}

/* ══ 本番 ═══════════════════════════════════════════════════ */
const SKIP = new Set(['node_modules', '.git', 'tmp', '.vercel', 'dist']);

/* ★`.git` は SKIP に 入っている＝集める は 一生 返さない★
   ⇒ 前の 形（集めた 名前から git複製を 探す）は ★除外が 一度も 効かなかった★
   ⇒ ★掘る 前に「その 部屋は git の 複製か」を 見て、部屋ごと 落とす★
   `wt` に 落とした 部屋を 積む（何を 除いたか 報告に 出す為） */
function 集める(base, rel, out, wt) {
  let 中;
  try { 中 = fs.readdirSync(path.join(base, rel || '.')); } catch (e) { return out; }
  if (rel && 中.includes('.git')) { if (wt) wt.push(rel); return out; }
  for (const n of 中) {
    if (SKIP.has(n)) continue;
    const r = rel ? rel + '/' + n : n;
    let st;
    try { st = fs.statSync(path.join(base, r)); } catch (e) { continue; }
    if (st.isDirectory()) 集める(base, r, out, wt); else out.push(r);
  }
  return out;
}

/* ★★置き場は「親の 部屋を 読んで 全部 見る」★★（差し戻し①）
   ★セッションの 番号を 書かない★＝次の 回でも 効く */
const 親の部屋 = process.env.EXALLY_SCRATCHPAD_OYA
  || 'C:/Users/zeroa/AppData/Local/Temp/claude/C--WINDOWS-System32-WindowsPowerShell-v1-0';
const 生きている線 = 48 * 3600 * 1000; /* ★実測 0.4時間 と 80時間 の 間★（上の 見出しに 根拠） */

function 置き場を調べる() {
  const 出 = [];
  const 見る = [];
  const 直 = process.env.EXALLY_SCRATCHPAD;
  if (直 && fs.existsSync(直)) 見る.push(直);
  try {
    for (const n of fs.readdirSync(親の部屋)) {
      const p = path.join(親の部屋, n, 'scratchpad');
      try { if (fs.statSync(p).isDirectory()) 見る.push(p); } catch (e) { /* 無い */ }
    }
  } catch (e) { /* 親の部屋が 無い */ }
  for (const 場 of [...new Set(見る)]) {
    const 名 = 場.replace(/\\/g, '/');
    const wt = [];
    const ファイル = 集める(場, '', [], wt);
    let 新 = 0;
    for (const f of ファイル) {
      try { const s = fs.statSync(path.join(場, f)); if (s.mtimeMs > 新) 新 = s.mtimeMs; } catch (e) { /* 消えた */ }
    }
    出.push({
      場: 名,
      名前: ファイル.map((f) => 名 + '/' + f),
      git複製: wt.map((d) => 名 + '/' + d),
      新,
      生きている: 新 > 0 && (Date.now() - 新) < 生きている線,
    });
  }
  return 出;
}

const repoの = 集める(ROOT, '', [], null);
const 置き場たち = 置き場を調べる();
const 生きている = 置き場たち.filter((x) => x.生きている);
const 古い = 置き場たち.filter((x) => !x.生きている);

const 名前 = (群) => 群.flatMap((x) => x.名前);
const 複製 = (群) => 群.flatMap((x) => x.git複製);
const 危ない = 危ないか(repoの, 名前(生きている), 複製(生きている));
const 古い置き忘れ = 危ないか(repoの, 名前(古い), 複製(古い));

/* ★★材料が 無い時に 緑を 出さない★★（差し戻し②） */
const 親が在る = fs.existsSync(親の部屋);
const 手元 = process.platform === 'win32';
if (この回の名乗り(親が在る, 手元) === '見ていない') {
  /* ★CI(ubuntu) には この 道が 無い★
     ⇒★「見ていない」と ★毎回 記録に 残す★★＝黙って 緑に しない
     ⇒ この 見張りが 本当に 働くのは ★手元(Windows)★。そこでは 下の 検査が 走る */
  console.log('\n[shouko-no-okiba] ★★この 回は 見ていません★★');
  console.log('  ★' + 親の部屋 + ' が 在りません（platform=' + process.platform + '）★');
  console.log('  ⇒ ★この 見張りは 手元(Windows)専用★＝CI では ★消える場所を 見ていない★');
  console.log('  ⇒ ★緑＝「置き場が きれい」では ありません★');
  process.exit(0);
}

console.log('\n[shouko-no-okiba] 消える場所に repo の 物を 置いていないか');

T('★★材料が 在る（親の 部屋が 読めて 置き場が 1つ以上）★★', () => {
  if (!親が在る) {
    throw new Error('★未測定（親の 部屋が 在りません）★'
      + ' … ' + 親の部屋
      + ' / 手元(Windows)なのに 道が 変わっています。EXALLY_SCRATCHPAD_OYA で 渡してください');
  }
  if (!置き場たち.length) {
    throw new Error('★未測定（置き場が 1つも ありません）★ … ' + 親の部屋
      + ' / ★この 回は 何も 見ていません＝緑を 出しては いけない★');
  }
  if (!生きている.length) {
    throw new Error('★未測定（生きている 置き場が ありません）★'
      + ' … 置き場 ' + 置き場たち.length + '個 は 全部 48時間 以上 動いていません'
      + ' / ★自分の 置き場が 数えられていない＝道が 変わった合図★');
  }
});

T('★★消える場所に repo と 同じ 名前の コードが 無い★★', () => {
  if (危ない.length) {
    throw new Error('★消える場所に repo の 物が 在ります★:\n'
      + 危ない.map((f) => '   - ' + f).join('\n')
      + '\n   → ★控えは `git switch -c <枝>` ＋ commit★（push しなくてよい）'
      + '\n     ★scratchpad は 消えます★＝2026-08-29 の 実Excel 507個は これで 消え、'
      + '\n     ★「507個」は 中身の 無い 数字に なりました★');
  }
});

T('★この 機械が「半分」だと 書いてある（守ったと 読ませない）', () => {
  /* ★捕まえられない 物が 在る★のに、それが 書いていないと
     ★「見張りが 在る＝安全」と 読まれる★＝一番 危ない */
  const 本文 = fs.readFileSync(fileURLToPath(import.meta.url), 'utf8');
  if (本文.indexOf('半分しか 守りません') < 0) throw new Error('★半分だと 書いていない★');
});

T('検査が空振りしていない（repo を 実際に 数えている）', () => {
  if (repoの.length < 100) throw new Error('repo の ファイルが 少なすぎます: ' + repoの.length);
});

console.log('\n── 実測 ──');
console.log('  repo … ' + repoの.length + '本');
console.log('  置き場 … ' + 置き場たち.length + '個'
  + '（★生きている ' + 生きている.length + '個★ / 48時間 動いていない ' + 古い.length + '個）');
console.log('  見た ファイル … ' + 名前(生きている).length + '本'
  + '（除いた git の 作業用 複製 ' + 複製(生きている).length + '個）');
console.log('  ★危ない（生きている 置き場）… ' + 危ない.length + '本★');
/* ★古い 置き忘れは 赤に しないが ★黙らない★★＝数と 場所を 必ず 出す */
if (古い置き忘れ.length) {
  console.log('  ─ 参考：48時間 動いていない 置き場の 置き忘れ … ' + 古い置き忘れ.length + '本');
  const 部屋 = {};
  for (const f of 古い置き忘れ) { const k = f.split('/scratchpad/')[0].split('/').pop(); 部屋[k] = (部屋[k] || 0) + 1; }
  for (const k of Object.keys(部屋)) console.log('      ' + k + ' … ' + 部屋[k] + '本');
  console.log('    ⇒★もう 居ない セッションの 物＝私には 直せない★ので 赤に していません');
  console.log('    ⇒★中身を 見て 要る物なら 枝に して ください★');
}
console.log('\n' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
