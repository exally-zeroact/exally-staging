/* ci-mitodoke.mjs (lib) -- ★1つの sha の CI と WebKit を「数まで」見届ける★（2026-10-10）
 *
 *  ★★なぜ 在るか★★
 *    「CI 緑」を gh の success だけで 言うと、名簿の 何本が 走ったかは 見て いない。
 *    10-09 までは scratchpad の 写し（消える 所）で 4回 見届けて いた ⇒ repo に 入れる。
 *    使う 口は scripts/ci-mitodoke.mjs（薄い 口）。判じは ここ（試験 tests/ci-mitodoke.test.mjs）。
 *
 *  ★★緑と 言う 時（全部 そろった 時だけ）★★
 *    ・CI（ci.yml）と WebKit（webkit.yml）の 回が その sha に 在り、全部 completed
 *    ・回の headSha が 頼んだ sha と 同じ
 *    ・回が 全部 success（cancelled も 赤）。再実行した 回は ★前の 回も 全部 success★
 *    ・CI の ログ：`Run node tests/run.js` の 段の 行だけを 見て、★その sha の★ 名簿の 全部が
 *      名簿の 順に「=== 名前 ===」で 出て、「★落ちた★」0、最後の 見出しの 後に「全テストファイル 緑」
 *    ・WebKit の ログ：その sha の 名簿から 出した 期待（1本以上）の 全部が 順に「=== 名前 ===」で 出て、
 *      ★未測定 の 行が 0、最後の 行の 頼んだ＝走らせた＝緑＝期待の 本数
 *
 *  ★★言えない 事（書いて おく）★★
 *    ・PR の CI は refs/pull/N/merge（土台と 混ぜた 物）を 走らせる ⇒ 混ぜた 側の 名簿は sha の 名簿 以上。
 *      見るのは「sha の 名簿の 全部が 走ったか」（部分集合）。★名簿から 黙って 外された 試験★は 捕まえない
 *      （それは tests-registered の 門の 仕事）。
 *    ・中の 試験が 名簿と 同じ 名前の「=== 名前 ===」を 出すと、名簿の 順に 当たる 限り 走ったと 数える。
 *    ・ci.yml の run.js 以外の 段の 中身は 見ない（その 段の 色は 回の conclusion に 入る）。
 *    ・子の 試験が 何も 言わずに exit 0 で 終われば、見出しと 最後の 行だけで 緑に なる＝子の 終わり値と
 *      webkit.yml の MEASURE_REQUIRED（未測定を 赤に する）を 信じて いる。★未測定 の 字は その 2枚目の 守り。
 *    ・出しは 数と run id だけ。ログの 全文・gh の 鍵は 出さない／書かない。
 *    ・WebKit の 期待は「借りる 試験か」の 決まりを ここに 写して 数える。決まり その物が 間違って いる
 *      （借りる 試験を 借りないと 読む）時は 両方 揃って 間違える＝それは tests/webkit-hashiru.test.mjs の 仕事。
 *
 *  ★★わざと 赤に 倒して いる 形（手で 見届ける 事に なる）★★
 *    ・★cancelled の 回が 1つでも 在れば 赤★（走らせ直しの 前の 回でも、別の 回でも）。GitHub は 時間切れも cancelled と 付けるので、
 *      固まった 試験を 走らせ直しで 隠さない。前の 本番 dc78b2e の WebKit（Install WebKit で 固まって
 *      走らせ直した）が この 形。次に 同じ 固まり方を したら、道具は 赤 ⇒ 固まった 段を ログで 見て 手で 見届ける。
 *    ・同じ sha の 回が 100 以上、tests/run.js に 名簿だけを 読む 守りが 無い 古い sha も 赤。
 */
import { execFileSync } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath, pathToFileURL } from 'node:url';

export const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..', '..');
const 時刻 = /^\uFEFF?\d{4}-\d\d-\d\dT[\d:.]+Z ?/;

/** ★WebKit の 期待の 名簿（見出しの 名前）★＝ブラウザを 借りる 試験から webkit.yml に 個別の 段が 在る 物を 除いた 数。
 *  ★わざと scripts/run-webkit-tests.mjs を import しない★（10-10 本番前の 対立役）：同じ 関数から
 *  期待と 頼んだを 出すと、その 関数が 黙って 試験を 落とした 時に 両方 減って 緑に なる。
 *  決まりは 同じ（import して いる 物だけ・注記の 字では 拾わない）。字が ずれれば 赤（閉じる 側）。 */
export function webkitの期待(FILES, 読む, yml) {
  const 借りる = (s) => /^\s*import\s[^;\n]*['"][^'"]*_borrow-playwright\.mjs['"]/m.test(String(s))
    || /import\(\s*[^)]*_borrow-playwright\.mjs/.test(String(s));
  const 個別 = new Set();
  for (const m of String(yml || '').matchAll(/run:\s*node\s+tests\/([A-Za-z0-9_.\/-]+\.mjs)/g)) 個別.add(m[1]);
  const 出 = [];
  for (const f of FILES) {
    const a = (Array.isArray(f) ? f : [f]).filter((x) => x !== null && x !== undefined);
    const 名 = String(a[0]);
    let s = '';
    try { s = 読む(名); } catch (e) { continue; }
    if (借りる(s) && !個別.has(名)) 出.push(a.join(' '));
  }
  return 出;
}

/** run.js と 同じ 組み立て（file + ' ' + args.join(' ')） */
export function 名簿の名前(FILES) {
  return FILES.map((f) => {
    const [file, ...args] = Array.isArray(f) ? f : [f];
    return file + (args.length ? ' ' + args.join(' ') : '');
  });
}

/** ログを 段ごとに 分ける：「job\t段の名前\t時刻 本文」 ⇒ { 段, 本文 } */
export function ログを読む(生) {
  const 行 = [];
  for (const l of String(生 || '').split('\n')) {
    const s = l.replace(/\r$/, '');
    const a = s.indexOf('\t'), b = a < 0 ? -1 : s.indexOf('\t', a + 1);
    if (b < 0) continue;
    行.push({ 段: s.slice(a + 1, b), 本文: s.slice(b + 1).replace(時刻, '') });
  }
  return 行;
}

/** 中身で 段を 拾う：`##[group]Run <命令>` の 行の 段の 名前 */
function 段の行(行, 命令) {
  const 頭 = 行.find((x) => x.本文 === '##[group]Run ' + 命令);
  return 頭 ? 行.filter((x) => x.段 === 頭.段) : null;
}

/** 回の 並び（1つの ワークフロー）を 見る */
function 回を見る(名, runs, sha, 赤, 行) {
  if (!runs || !runs.length) { 赤.push(名 + ' の 回が 無い（paths に 当たらない か PR が 無い か まだ 作られて いない）'); return []; }
  const 取消 = runs.filter((r) => r.status === 'completed' && r.conclusion === 'cancelled').length;
  const 良い = [];
  for (const r of runs) {
    const id = r.databaseId + (r.attempt > 1 ? ' attempt ' + r.attempt : '');
    if (r.headSha !== sha) { 赤.push(名 + ' ' + id + ' の headSha が 違う（' + String(r.headSha).slice(0, 7) + '）'); continue; }
    if (r.status !== 'completed') { 赤.push(名 + ' ' + id + ' が まだ ' + r.status); continue; }
    if (!(Number.isInteger(r.attempt) && r.attempt >= 1)) { 赤.push(名 + ' ' + r.databaseId + ' の attempt が 読めない'); continue; }
    /* ★前の 回は cancelled の 回でも 見る★（取り消した 回の 中に 前の failure が 隠れる・10-10 本番前の 対立役） */
    const 前 = r.前の回 || [];
    if (r.attempt > 1 && 前.length !== r.attempt - 1) { 赤.push(名 + ' ' + id + ' の 前の 回が 取れない'); continue; }
    const 前の赤 = 前.filter((c) => c !== 'success');
    if (前の赤.length) { 赤.push(名 + ' ' + id + ' は 再実行の 回（前の 回 ' + 前の赤.join('/') + '）'); continue; }
    /* ★別の 回の cancelled も 赤★（時間切れも cancelled＝定時の 回が 固まり push の 回が 緑、を 隠さない・10-10 叩き直し） */
    if (r.conclusion === 'cancelled') { 赤.push(名 + ' ' + id + ' が cancelled（時間切れも cancelled と 付く）'); continue; }
    if (r.conclusion !== 'success') { 赤.push(名 + ' ' + id + ' が ' + r.conclusion); continue; }
    良い.push(r);
  }
  if (!良い.length && !赤.some((s) => s.startsWith(名))) 赤.push(名 + ' に success の 回が 無い（cancelled ' + 取消 + '）');
  行.push(名 + '：回 ' + runs.length + ' ／ success ' + 良い.length + ' ／ cancelled ' + 取消 + '（' + 良い.map((r) => r.databaseId + ' attempt ' + (r.attempt || 1)).join(', ') + '）');
  return 良い;
}

/** CI の ログの 中身 */
export function CIログを見る(生, 名前, 赤, 行, id) {
  const 段 = 段の行(ログを読む(生), 'node tests/run.js');
  if (!段) { 赤.push('CI ' + id + '：run.js の 段が ログに 無い'); return; }
  const 本文 = 段.map((x) => x.本文);
  let i = 0, 最後の見出し = -1;
  const 欠け = [];
  for (const n of 名前) {
    const 見出し = '=== ' + n + ' ===';
    let j = i;
    while (j < 本文.length && 本文[j] !== 見出し) j++;
    if (j >= 本文.length) { 欠け.push(n); continue; }
    最後の見出し = j; i = j + 1;
  }
  const 落 = 本文.filter((s) => /^\s*★落ちた★/.test(s)).length;
  const 終 = 本文.some((s, k) => k > 最後の見出し && s === '全テストファイル 緑');
  行.push('CI ' + id + '：名簿 ' + 名前.length + ' ／ 走った ' + (名前.length - 欠け.length) + ' ／ 落ちた ' + 落 + ' ／ 終わりの緑 ' + 終);
  if (!名前.length) 赤.push('CI ' + id + '：名簿が 0本');
  if (欠け.length) 赤.push('CI ' + id + '：名簿の ' + 欠け.length + '本が 走って いない（' + 欠け.slice(0, 3).join(', ') + '）');
  if (落) 赤.push('CI ' + id + '：落ちた ' + 落);
  if (!終) 赤.push('CI ' + id + '：最後の 見出しの 後に 終わりの 緑が 無い');
}

/** WebKit の ログの 中身 */
export function WebKitログを見る(生, 期待名, 赤, 行, id) {
  const 段 = 段の行(ログを読む(生), 'node scripts/run-webkit-tests.mjs');
  if (!段) { 赤.push('WebKit ' + id + '：run-webkit-tests の 段が ログに 無い'); return; }
  const 期待 = (期待名 || []).length;
  /* ★数だけで なく 見出しの 名前と 順★・★未測定★ の 行（10-10 叩き直し＝全部 未測定でも 18/18/18 で 緑だった） */
  const 本文 = 段.map((x) => x.本文);
  let i = 0;
  const 欠け = [];
  for (const n of 期待名 || []) {
    let j = i;
    while (j < 本文.length && 本文[j] !== '=== ' + n + ' ===') j++;
    if (j >= 本文.length) { 欠け.push(n); continue; }
    i = j + 1;
  }
  const 未測定 = 本文.filter((s) => s.includes('★未測定')).length;
  if (欠け.length) 赤.push('WebKit ' + id + '：期待の ' + 欠け.length + '本が 走って いない（' + 欠け.slice(0, 3).join(', ') + '）');
  if (未測定) 赤.push('WebKit ' + id + '：★未測定 の 行が ' + 未測定);
  let m = null;
  let 数の行 = 0;
  for (const x of 段) { const t = /^★webkit の 見張り \.\.\. 頼んだ (\d+)本 ／ 走らせた (\d+)本 ／ 緑 (\d+)本★$/.exec(x.本文); if (t) { m = t; 数の行++; } }
  if (!m) { 赤.push('WebKit ' + id + '：最後の 行が 無い（途中で 切れた）'); return; }
  /* ★数の 行は ちょうど 1本★（2本 在ると どちらを 信じるか 決まらない・10-10 4回目の 対立役） */
  if (数の行 !== 1) 赤.push('WebKit ' + id + '：数の 行が ' + 数の行 + '本（1本で ない）');
  const [頼, 走, 緑] = [+m[1], +m[2], +m[3]];
  行.push('WebKit ' + id + '：期待 ' + 期待 + ' ／ 頼んだ ' + 頼 + ' ／ 走らせた ' + 走 + ' ／ 緑 ' + 緑 + ' ／ 見出しの 欠け ' + 欠け.length + ' ／ 未測定 ' + 未測定);
  if (!(頼 === 走 && 走 === 緑)) 赤.push('WebKit ' + id + '：3つの 数が 揃わない');
  if (!(期待 >= 1)) 赤.push('WebKit ' + id + '：期待の 本数が 0（名簿を 読み損じた）');
  if (頼 !== 期待) 赤.push('WebKit ' + id + '：頼んだ ' + 頼 + ' が 期待 ' + 期待 + ' と 違う');
}

/** ★判じ★（網も git も 叩かない）。ci/wk = { runs, ログ: { [databaseId]: 生 | null } } */
export function 見届ける({ sha, ci, wk, 名簿, webkit期待 }) {
  const 赤 = [], 行 = [];
  if (!/^[0-9a-f]{40}$/.test(String(sha))) 赤.push('sha が 全桁で ない');
  for (const r of 回を見る('CI', ci && ci.runs, sha, 赤, 行)) {
    const 生 = ci.ログ && ci.ログ[r.databaseId];
    if (!生) { 赤.push('CI ' + r.databaseId + '：ログが 取れない'); continue; }
    CIログを見る(生, 名簿の名前(名簿 || []), 赤, 行, r.databaseId);
  }
  for (const r of 回を見る('WebKit', wk && wk.runs, sha, 赤, 行)) {
    const 生 = wk.ログ && wk.ログ[r.databaseId];
    if (!生) { 赤.push('WebKit ' + r.databaseId + '：ログが 取れない'); continue; }
    WebKitログを見る(生, webkit期待, 赤, 行, r.databaseId);
  }
  return { 赤, 行 };
}

/* ---------- ここから下は 網と git を 叩く（scripts/ci-mitodoke.mjs から 呼ぶ） ---------- */

const REPO = 'exally-zeroact/exally';
const 叩く = (cmd, args) => execFileSync(cmd, args, { encoding: 'utf8', maxBuffer: 1 << 30, stdio: ['ignore', 'pipe', 'pipe'], cwd: ROOT });

/** その sha の tests/run.js から FILES を 読む（手元の 作業木では ない） */
export function shaの名簿(sha) {
  const 置き場 = fs.mkdtempSync(path.join(os.tmpdir(), 'ci-mitodoke-'));
  try {
    const f = path.join(置き場, 'run.js');
    const 生 = 叩く('git', ['show', sha + ':tests/run.js']);
    /* 守りの 無い 古い run.js は require した 途端に 全部を 走らせる（e6698b4 で 実測）＝読まずに 赤 */
    if (!生.includes('require.main !== module')) throw new Error('この sha の tests/run.js は 名簿だけを 読む 守りが 無い（古い）');
    fs.writeFileSync(f, 生);
    const { FILES } = createRequire(pathToFileURL(f))(f);
    const 読む = (n) => 叩く('git', ['show', sha + ':tests/' + n]);
    let yml = '';
    try { yml = 叩く('git', ['show', sha + ':.github/workflows/webkit.yml']); } catch (e) { yml = ''; }
    return { FILES, webkit期待: webkitの期待(FILES, 読む, yml) };
  } finally { fs.rmSync(置き場, { recursive: true, force: true }); }
}

function 回を引く(ファイル, sha) {
  const 上限 = 100;
  const runs = JSON.parse(叩く('gh', ['run', 'list', '--repo', REPO, '--workflow', ファイル, '--commit', sha, '--limit', String(上限),
    '--json', 'databaseId,status,conclusion,headSha,attempt']));
  if (runs.length >= 上限) throw new Error(ファイル + ' の 回が ' + 上限 + ' 以上（取りこぼすかも 知れない）');
  const ログ = {};
  for (const r of runs) {
    r.前の回 = [];
    for (let k = 1; k < (r.attempt || 1); k++) {
      try { r.前の回.push(JSON.parse(叩く('gh', ['api', 'repos/' + REPO + '/actions/runs/' + r.databaseId + '/attempts/' + k])).conclusion); }
      catch (e) { /* 取れない 回は 数が 足りず 赤 */ }
    }
    if (r.status !== 'completed' || r.conclusion !== 'success') continue;
    ログ[r.databaseId] = null;
    for (let i = 0; i < 3; i++) {
      try { const t = 叩く('gh', ['run', 'view', String(r.databaseId), '--repo', REPO, '--log']); if (t.length > 1000) { ログ[r.databaseId] = t; break; } }
      catch (e) { /* 取り直す */ }
    }
  }
  return { runs, ログ };
}

/** 直に 走らせる 口。★既定は exit 1★、緑を 出す 1か所で だけ 0 に 戻す */
export function 走らせる(引数) {
  process.exitCode = 1;
  const 頼み = 引数[0];
  if (!頼み) { console.log('使い方: node scripts/ci-mitodoke.mjs <sha>'); return; }
  let sha;
  try { sha = 叩く('git', ['rev-parse', '--verify', 頼み + '^{commit}']).trim(); }
  catch (e) { console.log('★赤（見届けられない）★ sha が 引けない（git fetch して ください）'); return; }
  let 結果;
  try {
    const { FILES, webkit期待 } = shaの名簿(sha);
    結果 = 見届ける({ sha, ci: 回を引く('ci.yml', sha), wk: 回を引く('webkit.yml', sha), 名簿: FILES, webkit期待 });
  } catch (e) {
    console.log('★赤（見届けられない）★ 途中で 止まった（gh・git・古い sha）：' + String((e && e.message) || e).split('\n')[0].slice(0, 120));
    return;
  }
  for (const l of 結果.行) console.log(l);
  if (結果.赤.length) { console.log('★赤（見届けられない）★ ' + 結果.赤.join(' ／ ')); return; }
  console.log('★緑★ ' + sha.slice(0, 7));
  process.exitCode = 0;
}
