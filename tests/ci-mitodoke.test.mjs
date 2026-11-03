/* ci-mitodoke.test.mjs -- ★CI の 見届けの 道具が「偽の 緑」を 返さないか★（2026-10-10）
 *
 *  道具＝scripts/ci-mitodoke.mjs（判じは scripts/lib/ci-mitodoke.mjs の 見届ける）。
 *  作り物の ログは 本物の gh ログ（49d56dc の CI 37905233170・WebKit 37905233158）の 行の 形を 写した：
 *    「job\t段の名前\t(BOM)時刻 本文」・段の 頭は「##[group]Run <命令>」。
 *  緑の 形 5つ・赤の 形 38（経営者の 受け入れ 5形と 対立役の 8形を 含む）を 当てる。
 *  --self-test ... 判じを 19通り 壊した 写しに 同じ 試験を 当て、★どれも 赤が 出る★のを 見る
 */
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createRequire } from 'node:module';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const LIB = path.join(ROOT, 'scripts', 'lib', 'ci-mitodoke.mjs');

/* ══ 作り物 ══ */
const SHA = 'a'.repeat(40), 違うSHA = 'b'.repeat(40);
const 名簿 = ['stamp.test.mjs', ['sql-guard.test.mjs', '--self-test'], ['shiki-kiru.test.mjs'], 'webkit-a.mjs'];
const 名前 = ['stamp.test.mjs', 'sql-guard.test.mjs --self-test', 'shiki-kiru.test.mjs', 'webkit-a.mjs'];
const 時 = (n) => '2026-10-09T08:29:' + String(25 + n).padStart(2, '0') + '.7295892Z ';
const 行 = (job, 段, 本文, bom) => job + '\t' + 段 + '\t' + (bom ? '\uFEFF' : '') + 時(0) + 本文;
const CI段 = 'Exally tests (共有データ/集計)';
function CIログ(o) {
  o = o || {};
  const 出 = names => names.map((n) => 行('test', CI段, '=== ' + n + ' ==='));
  const L = [行('test', 'Install test deps', '##[group]Run npm install', true), 行('test', 'Install test deps', 'added 10 packages')];
  if (o.別段) L.push(...名前.map((n) => 行('test', 'Install test deps', '=== ' + n + ' ===')));
  L.push(行('test', o.段名 || CI段, o.頭 || '##[group]Run node tests/run.js', true));
  L.push(...(o.前 || []).map((s) => 行('test', CI段, s)));
  L.push(...出(o.名前 || 名前).map((s) => o.段名 ? s.replace(CI段, o.段名) : s));
  if (o.落ちた) L.push(行('test', CI段, '  ★落ちた★ shiki-kiru.test.mjs ･･･ 自分で 1 を返した'));
  if (!o.終わり無し) L.push(行('test', o.段名 || CI段, '全テストファイル 緑'));
  L.push(...(o.後 || []).map((s) => 行('test', CI段, s)));
  L.push(行('test', '★重複ガード', '##[group]Run node tests/no-duplicate-libs.test.mjs', true));
  const t = L.join('\n') + '\n' + 'x'.repeat(1000);
  return o.crlf ? t.replace(/\n/g, '\r\n') : t;
}
const WK段 = '★名簿の webkit の 見張り 全部（tests/run.js から 拾う）';
function WKログ(頼, 走, 緑, o) {
  o = o || {};
  const 見出し = (o.名前 || ['webkit-a.mjs']).map((n) => 行('webkit', WK段, '=== ' + n + ' ==='));
  const L = [];
  if (o.echo) L.push(行('webkit', 'echo の 段', '##[group]Run echo node scripts/run-webkit-tests.mjs', true), 行('webkit', 'echo の 段', 'node scripts/run-webkit-tests.mjs'));
  L.push(行('webkit', '★借り物', '##[group]Run node tests/karimono.test.mjs', true),
    行('webkit', '★借り物', '★webkit の 見張り ... 頼んだ 9本 ／ 走らせた 9本 ／ 緑 9本★'));
  if (o.別段見出し) L.push(行('webkit', '★借り物', '=== webkit-a.mjs ==='));
  L.push(
    行('webkit', WK段, o.頭 || '##[group]Run node scripts/run-webkit-tests.mjs', true),
    ...見出し);
  if (o.未測定) L.push(行('webkit', WK段, '  ★未測定★ 借りられません（作り物）'));
  /* 崩れた 数の 行を ★先に★ 置く（最後の 行だけ 信じる 判じでは 緑に なる 形） */
  if (o.二本目) L.push(行('webkit', WK段, '★webkit の 見張り ... 頼んだ 1本 ／ 走らせた 1本 ／ 緑 0本★'));
  if (!o.切れ) L.push(行('webkit', WK段, '★webkit の 見張り ... 頼んだ ' + 頼 + '本 ／ 走らせた ' + 走 + '本 ／ 緑 ' + 緑 + '本★'));
  return L.join('\n');
}
const 回 = (id, 足し) => Object.assign({ databaseId: id, status: 'completed', conclusion: 'success', headSha: SHA, attempt: 1, 前の回: [] }, 足し || {});
function 組(o) {
  o = o || {};
  const ciRuns = o.ciRuns || [回(1)], wkRuns = o.wkRuns || [回(2)];
  const ci = { runs: ciRuns, ログ: {} }, wk = { runs: wkRuns, ログ: {} };
  for (const r of ciRuns) ci.ログ[r.databaseId] = 'ciLog' in o ? o.ciLog : CIログ(o.ci);
  for (const r of wkRuns) wk.ログ[r.databaseId] = 'wkLog' in o ? o.wkLog : WKログ(...(o.wk || [1, 1, 1]));
  return { sha: o.sha || SHA, ci, wk, 名簿: o.名簿 || 名簿, webkit期待: 'webkit期待' in o ? o.webkit期待 : ['webkit-a.mjs'] };
}

/* ══ 試験の 中身（lib を 差し替えて 自己試験でも 使う） ══ */
async function 試す(libの道, 黙る) {
  const lib = await import(pathToFileURL(libの道).href + '?t=' + Date.now() + Math.random());
  let 緑 = 0, 赤 = 0;
  const T = (n, ok, m) => {
    if (ok) { 緑++; if (!黙る) console.log('  ok   ' + n); }
    else { 赤++; if (!黙る) console.log('  NG   ' + n + (m ? '\n       ' + m : '')); }
  };
  const 緑に = (n, o) => { const r = lib.見届ける(組(o)); T('緑に なる: ' + n, r.赤.length === 0, JSON.stringify(r)); };
  const 赤に = (n, o) => { const r = lib.見届ける(組(o)); T('赤に なる: ' + n, r.赤.length > 0, '★緑に なって しまった★ ' + JSON.stringify(r)); };

  緑に('★本物の 形の 緑ログ★（名簿 4／走った 4／WebKit 1/1/1・期待 1）');
  緑に('★CRLF の ログ★', { ci: { crlf: true } });
  緑に('★段の 名前が 変わっても 中身（Run node tests/run.js）で 拾う★', { ci: { 段名: 'テスト全部' } });
  赤に('★別の 回が cancelled（success が 1回 在っても）★', { ciRuns: [回(1), 回(3, { conclusion: 'cancelled' })] });

  /* 経営者の 受け入れ 5形 */
  赤に('① CI が in_progress', { ciRuns: [回(1, { status: 'in_progress', conclusion: null })] });
  赤に('① WebKit が queued', { wkRuns: [回(2, { status: 'queued', conclusion: null })] });
  赤に('② CI が cancelled だけ', { ciRuns: [回(1, { conclusion: 'cancelled' })] });
  赤に('② CI が skipped', { ciRuns: [回(1, { conclusion: 'skipped' })] });
  赤に('② CI の conclusion が null（completed なのに）', { ciRuns: [回(1, { conclusion: null })] });
  赤に('③ attempt 2 で 1回目が failure', { ciRuns: [回(1, { attempt: 2, 前の回: ['failure'] })] });
  赤に('③ attempt 2 なのに 前の 回が 取れない', { ciRuns: [回(1, { attempt: 2, 前の回: [] })] });
  赤に('③ 別の 回が success でも、取り消した 回の 前の 回が failure', { ciRuns: [回(1), 回(3, { conclusion: 'cancelled', attempt: 2, 前の回: ['failure'] })] });
  赤に('③ attempt が 無い（読めない）', { ciRuns: [回(1, { attempt: undefined })] });
  赤に('④ headSha が 頭と 違う', { wkRuns: [回(2, { headSha: 違うSHA })] });
  赤に('⑤ 名簿 4 ／ ログに 3 しか 出ない', { ci: { 名前: 名前.slice(0, 3) } });

  /* 対立役の 8形 ほか */
  赤に('(1) WebKit が 0/0/0（期待 0＝名簿を 読み損じ）', { wk: [0, 0, 0], webkit期待: [], wkLog: WKログ(0, 0, 0, { 名前: [] }) });
  赤に('(2) WebKit 3つは 揃うが 期待と 違う', { wk: [1, 1, 1], webkit期待: ['webkit-a.mjs', 'webkit-b.mjs'], wkLog: WKログ(1, 1, 1, { 名前: ['webkit-a.mjs', 'webkit-b.mjs'] }) });
  赤に('(3) 同じ sha に CI 2回・新しい方 success／古い方 failure', { ciRuns: [回(1), 回(3, { conclusion: 'failure' })] });
  赤に('(3) 同じ sha に CI 2回・並びが 逆', { ciRuns: [回(3, { conclusion: 'failure' }), 回(1)] });
  赤に('(5) run.js の 段が 無い（命令が 変わった）', { ci: { 頭: '##[group]Run node tests/zenbu.js' } });
  赤に('(6) 外の 見出しが 欠け、中の 試験が 同じ 名前を 前に 出した', { ci: { 名前: ['stamp.test.mjs', 'sql-guard.test.mjs --self-test', 'webkit-a.mjs'], 前: ['=== shiki-kiru.test.mjs ==='] } });
  赤に('★落ちた★ が 在る', { ci: { 落ちた: true } });
  赤に('終わりの 緑が 無い', { ci: { 終わり無し: true } });
  赤に('終わりの 緑が 最後の 見出しより 前にしか 無い', { ci: { 終わり無し: true, 前: ['全テストファイル 緑'] } });
  赤に('WebKit の 最後の 行が 無い（途中で 切れた）', { wkLog: WKログ(1, 1, 1, { 切れ: true }) });
  赤に('WebKit の 3つが 揃わない', { wk: [1, 1, 0] });
  赤に('WebKit の 見出しが 期待と 違う 名前（数は 揃う）', { wkLog: WKログ(1, 1, 1, { 名前: ['other.mjs'] }) });
  赤に('WebKit の 見出しが 0本（数だけ 1/1/1）', { wkLog: WKログ(1, 1, 1, { 名前: [] }) });
  赤に('WebKit の ログが 取れない（null）', { wkLog: null });
  赤に('WebKit の 見出しが 別の 段（karimono）にだけ 在る', { wkLog: WKログ(1, 1, 1, { 名前: [], 別段見出し: true }) });
  赤に('CI の 見出しが 別の 段にだけ 在る', { ci: { 名前: [], 別段: true } });
  緑に('★命令の 字を 含む 別の 段（echo）が 先に 在っても 本物の 段を 拾う★', { wkLog: WKログ(1, 1, 1, { echo: true }) });
  赤に('WebKit の 数の 行が 2本（1本目 崩れ・2本目 1/1/1）', { wkLog: WKログ(1, 1, 1, { 二本目: true }) });
  赤に('run-webkit-tests の 段が 無い（命令が 変わった）', { wkLog: WKログ(1, 1, 1, { 頭: '##[group]Run node scripts/zenbu-webkit.mjs' }) });
  緑に('★WebKit 2本・期待の 順に 見出し★', { webkit期待: ['webkit-a.mjs', 'webkit-b.mjs --self-test'], wkLog: WKログ(2, 2, 2, { 名前: ['webkit-a.mjs', 'webkit-b.mjs --self-test'] }) });
  赤に('WebKit の 見出しの 順が 期待と 逆（数は 揃う）', { webkit期待: ['webkit-a.mjs', 'webkit-b.mjs --self-test'], wkLog: WKログ(2, 2, 2, { 名前: ['webkit-b.mjs --self-test', 'webkit-a.mjs'] }) });
  赤に('WebKit の 段に ★未測定 の 行（数は 揃う）', { wkLog: WKログ(1, 1, 1, { 未測定: true }) });
  赤に('status だけが まだ（conclusion は success）', { ciRuns: [回(1, { status: 'in_progress' })] });
  赤に('sha だけが 全桁で ない（回の headSha も 同じ 短い sha）', { sha: 'abc1234', ciRuns: [回(1, { headSha: 'abc1234' })], wkRuns: [回(2, { headSha: 'abc1234' })] });
  赤に('CI の ログが 取れない（null）', { ciLog: null });
  赤に('WebKit の 回が 無い', { wkRuns: [] });
  赤に('名簿が 0本', { 名簿: [], ci: { 名前: [] } });
  赤に('sha が 全桁で ない', { sha: 'abc1234' });

  /* 道具と 相手の 字が 揃って いるか（相手の 出しを 変えた 日に 気づく） */
  const 相手 = (f) => fs.readFileSync(path.join(ROOT, f), 'utf8');
  T('★名前の 組み立てが run.js の 見出しと 同じ★', 相手('tests/run.js').includes("'\\n=== ' + file + (args.length ? ' ' + args.join(' ') : '') + ' ==='")
    && JSON.stringify(lib.名簿の名前(名簿)) === JSON.stringify(名前));
  T('★run.js の 終わりの 字が 同じ★', 相手('tests/run.js').includes("'全テストファイル 緑'") && 相手('tests/run.js').includes("'  ★落ちた★ '"));
  T('★WebKit の 最後の 行の 字が 同じ★', 相手('scripts/run-webkit-tests.mjs').includes("'\\n★webkit の 見張り ... 頼んだ ' + 名簿.length + '本 ／ 走らせた ' + 走った + '本 ／ 緑 ' + 緑 + '本★'"));
  T('★ci.yml に run.js の 段・webkit.yml に run-webkit-tests の 段★', /run:\s*node tests\/run\.js\s*$/m.test(相手('.github/workflows/ci.yml'))
    && /run:\s*node scripts\/run-webkit-tests\.mjs\s*$/m.test(相手('.github/workflows/webkit.yml')));

  /* WebKit の 期待を 独りで 数える＝走らせる 側（run-webkit-tests の 名簿）と 今の 木で 同じ 数か
     ★走らせる 側が 黙って 試験を 落とすと ここが 赤★（10-10 本番前の 対立役＝.slice(0, 5) で 13本 消えても 全部の 門が 緑だった） */
  {
    const { FILES } = createRequire(import.meta.url)(path.join(ROOT, 'tests', 'run.js'));
    const 期待名 = lib.webkitの期待(FILES, (n) => fs.readFileSync(path.join(ROOT, 'tests', n), 'utf8'), 相手('.github/workflows/webkit.yml'));
    const r = spawnSync(process.execPath, [path.join(ROOT, 'scripts', 'run-webkit-tests.mjs'), '--名簿だけ'], { encoding: 'utf8', cwd: ROOT });
    const 出し = String(r.stdout || '').split('\n').map((s) => s.replace(/\r$/, '')).filter((s) => s);
    const m = /★webkit の 見張り (\d+)本★/.exec(出し[出し.length - 1] || '');
    const 走名 = 出し.slice(0, -1);
    /* ★数で なく 名前の 並び そのもの★（引数を 落とす・個別の 段を 除かない 等の 壊しは 数では 見えない・10-10 3回目の 対立役） */
    T('★WebKit の 期待（独りで 数えた ' + 期待名.length + '本）＝走らせる 側の 名簿（' + (m && m[1]) + '本）の 名前の 並び そのもの・1本 以上★',
      !!m && +m[1] === 走名.length && JSON.stringify(走名) === JSON.stringify(期待名) && 期待名.length >= 1 && 期待名.some((n) => / /.test(n)),
      JSON.stringify(期待名) + ' / ' + JSON.stringify(走名));
  }

  /* 直に 走らせる 口：既定は exit 1 */
  if (!黙る) {
    const 口 = (a) => spawnSync(process.execPath, [path.join(ROOT, 'scripts', 'ci-mitodoke.mjs')].concat(a), { encoding: 'utf8', cwd: ROOT });
    const r0 = 口([]), r1 = 口(['zzzz-not-a-sha']);
    T('★口：sha 無しは exit 1★', r0.status === 1, 'exit ' + r0.status);
    T('★口：引けない sha は exit 1★', r1.status === 1 && /赤/.test(r1.stdout), 'exit ' + r1.status + ' ' + r1.stdout);
  }
  return { 緑, 赤 };
}

/* ══ 自己試験：判じを 壊した 写しで 赤が 出るか ══ */
const 壊し方 = [
  ['いつも 緑を 返す', 'return { 赤, 行 };', 'return { 赤: [], 行 };'],
  ['期待の 本数を 見ない', "if (頼 !== 期待) 赤", "if (false) 赤"],
  ['headSha を 見ない', 'if (r.headSha !== sha)', 'if (false)'],
  ['名簿の 順を 見ない', 'while (j < 本文.length && 本文[j] !== 見出し) j++;', 'j = 本文.indexOf(見出し); if (j < 0) j = 本文.length;'],
  ['前の 回を 見ない', 'if (前の赤.length)', 'if (false)'],
  ['attempt を 見ない', '!(Number.isInteger(r.attempt) && r.attempt >= 1)', 'false'],
  ['期待を 借りる 試験で 絞らない', 'if (借りる(s) && !個別.has(名)) 出.push', 'if (true) 出.push'],
  ['status を 見ない', "if (r.status !== 'completed')", 'if (false)'],
  ['sha の 桁を 見ない', "if (!/^[0-9a-f]{40}$/.test(String(sha)))", 'if (false)'],
  ['cancelled を 除く', "if (r.conclusion === 'cancelled') { 赤.push", "if (r.conclusion === 'cancelled') { continue; 赤.push"],
  ['WebKit の 見出しを 見ない', 'if (欠け.length) 赤.push(\'WebKit', 'if (false) 赤.push(\'WebKit'],
  ['未測定を 見ない', 'if (未測定) 赤.push', 'if (false) 赤.push'],
  ['WebKit の ログ無しを 見ない', "if (!生) { 赤.push('WebKit ' + r.databaseId", "if (!生) { continue; 赤.push('WebKit ' + r.databaseId"],
  ['WebKit の 段無しを 見ない', "if (!段) { 赤.push('WebKit ' + id", "if (!段) { return; 赤.push('WebKit ' + id"],
  ['WebKit の 順を 見ない', "while (j < 本文.length && 本文[j] !== '=== ' + n + ' ===') j++;", "j = 本文.indexOf('=== ' + n + ' ==='); if (j < 0) j = 本文.length;"],
  ['段で 絞らない', 'return 頭 ? 行.filter((x) => x.段 === 頭.段) : null;', 'return 頭 ? 行 : null;'],
  ['段の 頭を 字の 一部で 照らす', "x.本文 === '##[group]Run ' + 命令", 'x.本文.includes(命令)'],
  ['数の 行の 本数を 見ない', 'if (数の行 !== 1) 赤.push', 'if (false) 赤.push'],
  ['期待の 名前から 引数を 落とす', "出.push(a.join(' '))", '出.push(名)'],
];

if (process.argv.includes('--self-test')) {
  console.log('[ci-mitodoke --self-test] ★判じを 壊すと 赤が 出るか★');
  const 元 = fs.readFileSync(LIB, 'utf8');
  const 置き場 = fs.mkdtempSync(path.join(os.tmpdir(), 'ci-mitodoke-self-'));
  let 赤 = 0;
  try {
    for (const [名, 前, 後] of 壊し方) {
      if (!元.includes(前)) { 赤++; console.log('  NG   壊す 所が 見つからない: ' + 名); continue; }
      const 写し = path.join(置き場, 'lib-' + 壊し方.findIndex((x) => x[0] === 名) + '.mjs');
      fs.writeFileSync(写し, 元.replace(前, 後));
      const r = await 試す(写し, true);
      if (r.赤 > 0) console.log('  ok   壊すと 赤が 出る: ' + 名 + '（赤 ' + r.赤 + '）');
      else { 赤++; console.log('  NG   ★壊しても 全部 緑★: ' + 名); }
    }
  } finally { fs.rmSync(置き場, { recursive: true, force: true }); }
  console.log('\nci-mitodoke --self-test: ' + (壊し方.length - 赤) + ' 緑 / ' + 赤 + ' 赤');
  process.exit(赤 ? 1 : 0);
} else {
  console.log('[ci-mitodoke] ★見届けの 道具が 偽の 緑を 返さないか★');
  const r = await 試す(LIB, false);
  console.log('\nci-mitodoke: ' + r.緑 + ' 緑 / ' + r.赤 + ' 赤');
  process.exit(r.赤 ? 1 : 0);
}
