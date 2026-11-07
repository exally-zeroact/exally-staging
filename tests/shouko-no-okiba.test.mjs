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
 *  ★★どこまでを 赤に するか（2026-09-05 実測の 上で 決めた）★★（★経緯★＝48時間の 線は 2026-11-07 の 作り直しで 消した・下を 見る）
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
 *
 *  ★★2026-11-07 作り直し（親の 置き場が 古い道のまま＝手元で 毎回 赤）★★
 *    実測：既定の 親（Temp/claude/C--WINDOWS-…）＝置き場106 全部 48h超＝「生きている 0」で 毎回 赤
 *          親を 今の 部屋に 替えると 置き場138・★危ない 29401本★（他の 席の 木の 写し・他アプリ）
 *          ★中身の blob 指紋が git に 在る物を 除くと 1565本（48h内）★
 *    ⇒ 設計（設計前 taiketsu の 止め2を 入れた形）：
 *      ①見る所＝和集合［EXALLY_SCRATCHPAD・tmpdir/claude/＊/<CLAUDE_CODE_SESSION_ID>/scratchpad・ROOT を 含む scratchpad］
 *        ★親の 名を 決め打ちしない★／★他の 席の 置き場は 見ない★（「見ていない N個」の 1行だけ）
 *      ②赤＝repo と 同じ 名前の .js/.mjs/.cjs/.json のうち ★中身の blob 指紋が `git rev-list --all --objects` に 無い物★
 *        ＝git に 元が 在る 写しは 消えても 困らない／★rev-list が 転んだら 空の 集まり＝全部 赤★
 *      ③名乗り：CI＝見ていない／EXALLY_SCRATCHPAD が 無い道＝赤／番号が 在るのに 置き場が 無い＝赤／
 *               win32 で CLAUDECODE=1 なのに 番号 無し＝赤／人が 手で 回す（どちらも 無し）＝見ていない
 *      ④48時間の 線は 消した（自分の 席だけ 見る＝古い 他席の 物で 赤に ならない）
 *      ⑤rev-list は 名乗りの 後で 呼ぶ（CI は 浅い 複製）
 *
 *  使い方: node tests/shouko-no-okiba.test.mjs
 *          node tests/shouko-no-okiba.test.mjs --self-test
 */
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import crypto from 'node:crypto';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ✓ ' + n); } catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + (e && e.message)); } };

/** ★純関数★＝repo の 名前の 集まりと 消える場所の 名前から「危ない物」を 返す
 *  ★git の 作業用 複製(worktree)は 除く★（2026-09-05 実測で 199本 出た）
 *  ★元が在るか(f)★＝渡されたら ★中身が git に 在る 写しは 除く★（2026-11-07） */
export function 危ないか(repoの名前, 消える場所の名前, git複製, 元が在るか) {
  const 見る = /\.(js|mjs|cjs|json)$/i;
  const 集 = new Set(repoの名前.map((f) => f.split('/').pop()));
  /* ★部屋の 境目まで 見る★＝`wt-a` で `wt-ab` を 巻き添えに しない */
  const 除く = (git複製 || []).map((d) => String(d).replace(/\\/g, '/').replace(/^\/+|\/+$/g, ''));
  const 除くか = (f) => 除く.some((d) => f === d || f.indexOf(d + '/') === 0);
  return 消える場所の名前.filter((f) => {
    if (!見る.test(f)) return false;
    if (除くか(f)) return false;
    if (!集.has(f.split('/').pop())) return false;
    if (元が在るか && 元が在るか(f)) return false;
    return true;
  });
}

/** ★この 回を どう 名乗るか★（純関数）
 *  返す＝{ 答: '見る'|'赤'|'見ていない', 訳 } */
export function この回の名乗り(x) {
  if (x.CI) return { 答: '見ていない', 訳: 'CI（消える場所は 手元にしか 無い）' };
  if (x.直が渡された && !x.直が在る) return { 答: '赤', 訳: 'EXALLY_SCRATCHPAD が 無い道を 指している' };
  if (x.番号) {
    if (!x.番号の置き場が在る) return { 答: '赤', 訳: '番号 ' + x.番号 + ' の 置き場が 見つからない（道が 変わった合図）' };
    return { 答: '見る', 訳: '' };
  }
  if (x.直が渡された) return { 答: '見る', 訳: '' };
  if (x.win32 && x.CLAUDECODE) return { 答: '赤', 訳: 'Claude の 席なのに CLAUDE_CODE_SESSION_ID が 無い＝自分の 置き場を 探せない' };
  return { 答: '見ていない', 訳: '人が 手で 回した（番号も EXALLY_SCRATCHPAD も 無い）' };
}

/** ★番号の 置き場が 見る所に 入っているか★（Windows の 道・大小文字を 揃える・末尾で 比べる） */
export function 番号の置き場が在るか(場たち, 番号) {
  if (!番号) return false;
  const 鍵 = '/' + 番号.toLowerCase() + '/scratchpad';
  return 場たち.some((p) => p.replace(/\\/g, '/').toLowerCase().replace(/\/+$/, '').endsWith(鍵));
}

/** CI か（CI=false・CI=0・空は CI でない） */
export function CIか(v) { return !!v && v !== 'false' && v !== '0'; }

/** ★見る 置き場の 和集合★（親の 名を 決め打ちしない） */
export function 置き場の候補({ tmp, 番号, 直, root }) {
  const 出 = [];
  if (直 && fs.existsSync(直)) 出.push(path.resolve(直));
  const 上 = path.join(tmp, 'claude');
  if (番号) {
    let 親たち = [];
    try { 親たち = fs.readdirSync(上); } catch (e) { /* 無い */ }
    for (const p of 親たち) {
      const s = path.join(上, p, 番号, 'scratchpad');
      try { if (fs.statSync(s).isDirectory()) 出.push(path.resolve(s)); } catch (e) { /* 無い */ }
    }
  }
  /* ★ROOT を 含む scratchpad★＝作業木を 置き場に 作った時、その 置き場の 控えも 見る */
  let d = path.resolve(root);
  for (;;) {
    if (path.basename(d).toLowerCase() === 'scratchpad') { 出.push(d); break; }
    const u = path.dirname(d);
    if (u === d) break;
    d = u;
  }
  const 済 = new Set();
  return 出.filter((p) => { const k = p.replace(/\\/g, '/').toLowerCase(); if (済.has(k)) return false; 済.add(k); return true; });
}

/** ★見ない 置き場の 数★（他の 席・数だけ 出す） */
function 全部の置き場の数(tmp) {
  let n = 0;
  const 上 = path.join(tmp, 'claude');
  let 親たち = [];
  try { 親たち = fs.readdirSync(上); } catch (e) { return 0; }
  for (const p of 親たち) {
    let ids = [];
    try { ids = fs.readdirSync(path.join(上, p)); } catch (e) { continue; }
    for (const id of ids) { try { if (fs.statSync(path.join(上, p, id, 'scratchpad')).isDirectory()) n++; } catch (e) { /* 無い */ } }
  }
  return n;
}

const SKIP = new Set(['node_modules', '.git', 'tmp', '.vercel', 'dist']);
/* ★掘る 前に「その 部屋は git の 複製か」を 見て、部屋ごと 落とす★（.git は ファイルでも 部屋でも） */
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

/** git の blob 指紋（`git hash-object` と 同じ） */
export function blob指紋(buf) {
  return crypto.createHash('sha1').update('blob ' + buf.length + '\0').update(buf).digest('hex');
}

/** ★rev-list が 転んだら 空の 集まり★＝全部 赤（黙って 緑に しない） */
function git指紋(root) {
  const 集 = new Set();
  let 訳 = '';
  try {
    const out = execFileSync('git', ['-C', root, 'rev-list', '--all', '--objects'], { maxBuffer: 1 << 30, stdio: ['ignore', 'pipe', 'pipe'] }).toString();
    for (const line of out.split('\n')) { const h = line.slice(0, 40); if (/^[0-9a-f]{40}$/.test(h)) 集.add(h); }
  } catch (e) { 訳 = String((e && e.message) || e).split('\n')[0]; }
  return { 集, 訳 };
}

/** 置き場ごとに 危ない物を 数える（道を 渡せる＝self-test でも 使う） */
export function 置き場を調べる(場たち, repoの, 指紋の集) {
  const 危 = [];
  let 見た = 0, 複 = 0;
  for (const 場 of 場たち) {
    const 名 = 場.replace(/\\/g, '/');
    const wt = [];
    const ファイル = 集める(場, '', [], wt);
    見た += ファイル.length; 複 += wt.length;
    const 元が在るか = (f) => { try { return 指紋の集.has(blob指紋(fs.readFileSync(path.join(場, f)))); } catch (e) { return false; } };
    for (const f of 危ないか(repoの, ファイル, wt, 元が在るか)) 危.push(名 + '/' + f);
  }
  return { 危, 見た, 複 };
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
  T('★★git複製は ★深さ1★でも 除く★★', () => {
    const r = 危ないか(repo, ['wt-a/lib/formula-extra.js'], ['wt-a']);
    if (r.length) throw new Error('除けていない: ' + JSON.stringify(r));
  });
  T('★★git複製は ★深い 段★でも 除く★★', () => {
    const r = 危ないか(repo, ['tmp/wt-a/lib/formula-extra.js'], ['tmp/wt-a']);
    if (r.length) throw new Error('★深さ1しか 除けていない★: ' + JSON.stringify(r));
  });
  T('★除くのは 部屋の 境目まで（wt-a で wt-ab を 巻き添えに しない）', () => {
    const r = 危ないか(repo, ['tmp/wt-ab/lib/formula-extra.js'], ['tmp/wt-a']);
    if (r.length !== 1) throw new Error('★除きすぎ★: ' + JSON.stringify(r));
  });
  T('★★指紋が git に 在る 写しは 緑／無い 控えは 赤★★', () => {
    const r = 危ないか(repo, ['a/claude.js', 'b/claude.js'], [], (f) => f.startsWith('a/'));
    if (r.length !== 1 || r[0] !== 'b/claude.js') throw new Error('分けていない: ' + JSON.stringify(r));
  });
  T('★blob 指紋は git hash-object と 同じ（空と "a\\n"）', () => {
    if (blob指紋(Buffer.from('')) !== 'e69de29bb2d1d6434b8b29ae775ad8c2e48c5391') throw new Error('空が 違う');
    if (blob指紋(Buffer.from('a\n')) !== '78981922613b2afb6025042ff6bd878ac1994e85') throw new Error('a が 違う');
  });
  T('★★名乗り：材料が 無い回は「緑」に ならない★★', () => {
    const 基 = { CI: false, win32: true, CLAUDECODE: true, 番号: 'S1', 番号の置き場が在る: true, 直が渡された: false, 直が在る: false };
    const 例 = [
      [{}, '見る', '全部 揃っている'],
      [{ CI: true }, '見ていない', 'CI'],
      [{ 番号: '' }, '赤', '★番号を 消す（win32・CLAUDECODE=1）★'],
      [{ 番号の置き場が在る: false }, '赤', '番号の 置き場が 無い'],
      [{ 直が渡された: true, 直が在る: false }, '赤', '★無い EXALLY_SCRATCHPAD★'],
      [{ 番号: '', 直が渡された: true, 直が在る: true }, '見る', 'EXALLY_SCRATCHPAD だけ'],
      [{ 番号: '', CLAUDECODE: false }, '見ていない', '人が 手で 回す'],
    ];
    for (const [差, 答, 名] of 例) {
      const r = この回の名乗り({ ...基, ...差 });
      if (r.答 !== 答) throw new Error(名 + ' が ' + r.答 + '（' + 答 + ' のはず）');
    }
  });
  /* ★実物の 道で 壊す★＝tmp/claude/<親>/<番号>/scratchpad を 作って 見る */
  const 試 = fs.mkdtempSync(path.join(os.tmpdir(), 'okiba-st-'));
  try {
    const 自 = path.join(試, 'claude', 'oya1', 'S1', 'scratchpad');
    const 他 = path.join(試, 'claude', 'oya2', 'S2', 'scratchpad');
    fs.mkdirSync(path.join(自, 'hold'), { recursive: true });
    fs.mkdirSync(path.join(他, 'hold'), { recursive: true });
    fs.writeFileSync(path.join(自, 'hold', 'claude.js'), 'ひかえ\n');
    fs.writeFileSync(path.join(他, 'hold', 'claude.js'), 'ひかえ\n');
    const 木 = path.join(他, 'wt');
    fs.mkdirSync(path.join(木, 'tests'), { recursive: true });
    fs.writeFileSync(path.join(木, '.git'), 'gitdir: x\n');
    fs.writeFileSync(path.join(木, 'tests', 'claude.js'), 'き\n');
    T('★★番号で 自分の 置き場を 見つけ・他の 席は 見ない★★', () => {
      const c = 置き場の候補({ tmp: 試, 番号: 'S1', 直: '', root: path.join(試, 'repo') });
      if (c.length !== 1 || path.resolve(c[0]) !== path.resolve(自)) throw new Error(JSON.stringify(c));
    });
    T('★★番号の 置き場が 在るかの 判定（Windows の 道・大小文字・他の 番号）★★', () => {
      if (!番号の置き場が在るか(['C:\\T\\claude\\p\\AbC\\scratchpad'], 'abc')) throw new Error('★Windows の 道で 当たらない★');
      if (番号の置き場が在るか(['C:/T/claude/p/xyz/scratchpad'], 'abc')) throw new Error('★他の 番号で 在ると 言った★');
      if (番号の置き場が在るか(['C:/T/claude/p/abc/scratchpad/wt/scratchpad'], 'abc')) throw new Error('★中の 置き場で 在ると 言った★');
      if (番号の置き場が在るか(['C:/T/claude/p/abc/scratchpad'], '')) throw new Error('★番号 無しで 在ると 言った★');
      const c = 置き場の候補({ tmp: 試, 番号: 'S1', 直: '', root: path.join(試, 'repo') });
      if (!番号の置き場が在るか(c, 'S1')) throw new Error('★実物の 道で 当たらない★');
      if (番号の置き場が在るか(置き場の候補({ tmp: 試, 番号: 'S9', 直: '', root: path.join(試, 'repo') }), 'S9')) throw new Error('★無い 番号で 在ると 言った★');
    });
    T('★CI=true は CI／CI=false・0・空は CI でない', () => {
      if (!CIか('true') || CIか('false') || CIか('0') || CIか('') || CIか(undefined)) throw new Error('分けていない');
    });
    T('★★ROOT が 置き場の 中なら その 置き場も 見る（木の 中は 除く）★★', () => {
      const c = 置き場の候補({ tmp: 試, 番号: 'S1', 直: '', root: 木 });
      if (c.length !== 2) throw new Error('2つの はず: ' + JSON.stringify(c));
      const r = 置き場を調べる(c, repo, new Set());
      if (r.危.length !== 2) throw new Error('★控え 2本の はず★: ' + JSON.stringify(r.危));
      if (r.危.some((f) => f.includes('/wt/'))) throw new Error('★木の 中を 赤に した★');
    });
    T('★★指紋の 集まりに 在れば 緑（git に 元が 在る）★★', () => {
      const h = blob指紋(Buffer.from('ひかえ\n'));
      const r = 置き場を調べる([自], repo, new Set([h]));
      if (r.危.length) throw new Error('★元が 在るのに 赤★: ' + JSON.stringify(r.危));
    });
    T('★★rev-list が 転んだ（空の 集まり）＝全部 赤★★', () => {
      const g = git指紋(path.join(試, 'gitでない'));
      if (g.集.size) throw new Error('git でない 所で 指紋が 出た');
      const r = 置き場を調べる([自], repo, g.集);
      if (r.危.length !== 1) throw new Error('★黙って 緑★: ' + JSON.stringify(r.危));
    });
  } finally { fs.rmSync(試, { recursive: true, force: true }); }
  T('★この 機械が「半分」だと 書いてある', () => {
    if (fs.readFileSync(fileURLToPath(import.meta.url), 'utf8').indexOf('半分しか 守りません') < 0) throw new Error('★半分だと 書いていない★');
  });
  console.log('\n' + pass + ' passed, ' + fail + ' failed');
  process.exit(fail ? 1 : 0);
}

/* ══ 本番 ═══════════════════════════════════════════════════ */
const tmp = os.tmpdir();
const 番号 = process.env.CLAUDE_CODE_SESSION_ID || '';
const 直 = process.env.EXALLY_SCRATCHPAD || '';
const 場たち = 置き場の候補({ tmp, 番号, 直, root: ROOT });
const 番号の置き場が在る = 番号の置き場が在るか(場たち, 番号);
const 名乗り = この回の名乗り({
  CI: CIか(process.env.CI),
  win32: process.platform === 'win32',
  CLAUDECODE: process.env.CLAUDECODE === '1',
  番号, 番号の置き場が在る,
  直が渡された: !!直, 直が在る: !!直 && fs.existsSync(直),
});

if (名乗り.答 === '見ていない') {
  console.log('\n[shouko-no-okiba] ★★この 回は 見ていません★★ … ' + 名乗り.訳);
  console.log('  ⇒ ★緑＝「置き場が きれい」では ありません★');
  process.exit(0);
}

console.log('\n[shouko-no-okiba] 消える場所に repo の 物を 置いていないか');
T('★★この 回の 名乗り（材料が 在る）★★', () => {
  if (名乗り.答 === '赤') throw new Error('★未測定★ … ' + 名乗り.訳);
});
if (名乗り.答 === '赤') { console.log('\n' + pass + ' passed, ' + fail + ' failed'); process.exit(1); }

/* ★rev-list は 名乗りの 後★（CI は 浅い 複製＝ここまで 来ない） */
const t0 = Date.now();
const 指紋 = git指紋(ROOT);
const repoの = 集める(ROOT, '', [], null);
const 調 = 置き場を調べる(場たち, repoの, 指紋.集);

T('★★git の 指紋が 読めた★★', () => {
  if (!指紋.集.size) throw new Error('★rev-list が 転んだ★（' + 指紋.訳 + '）＝名前の 当たった 物は 全部 赤に しています');
});
T('★★消える場所に git に 無い repo の コードが 無い★★', () => {
  if (調.危.length) {
    throw new Error('★消える場所に git に 元の 無い 控えが 在ります★:\n'
      + 調.危.map((f) => '   - ' + f).join('\n')
      + '\n   → ★控えは `git switch -c <枝>` ＋ commit★（push しなくてよい）'
      + '\n     ★scratchpad は 消えます★＝2026-08-29 の 実Excel 507個は これで 消えました');
  }
});
T('★この 機械が「半分」だと 書いてある（守ったと 読ませない）', () => {
  if (fs.readFileSync(fileURLToPath(import.meta.url), 'utf8').indexOf('半分しか 守りません') < 0) throw new Error('★半分だと 書いていない★');
});
T('検査が空振りしていない（repo を 実際に 数えている）', () => {
  if (repoの.length < 100) throw new Error('repo の ファイルが 少なすぎます: ' + repoの.length);
});

console.log('\n── 実測 ──');
console.log('  repo … ' + repoの.length + '本 ／ git の 指紋 … ' + 指紋.集.size + '個（' + (Date.now() - t0) + 'ms）');
console.log('  見た 置き場 … ' + 場たち.length + '個 ／ 見た ファイル … ' + 調.見た + '本（除いた git の 作業用 複製 ' + 調.複 + '個）');
console.log('  ★危ない … ' + 調.危.length + '本★');
console.log('  見ていない 置き場（他の 席）… ' + Math.max(0, 全部の置き場の数(tmp) - 場たち.length) + '個');
console.log('\n' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
