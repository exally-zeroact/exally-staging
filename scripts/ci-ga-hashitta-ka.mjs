/* ci-ga-hashitta-ka.mjs -- 「緑」と 言う 前に ★GitHub 側の 検査を 数える★（2026-09-19）
 *
 *  ★★なぜ 在るか（実際に 踏んだ 穴）★★
 *    `.github/workflows/ci.yml` の 引き金は
 *        on: push: branches: [main]   /   pull_request:
 *    ⇒★枝は ★PR が 開いて いる 間だけ★ 検査されます★
 *    ⇒2026-09-19T02:09 に PR #83 を merge
 *      ⇒★★その 瞬間に 枝の 検査が 黙って 止まりました★★
 *      ⇒その後 押した ★4本★（2人ぶん）に ★CI が 1回も 走って いません★
 *
 *  ★★なぜ 一番 危ないか★★
 *    ・★手元の 総なめは 緑の まま★＝いつも通りに 見える
 *    ・★赤に ならない★＝★門が 消えた 事を 門が 教えて くれない★
 *    ・★出す 直前に ちょうど 消える★（merge した 直後が 一番 危ない）
 *    ⇒ 記憶「見張りの 値打ちは ★いつ 赤に なるか★」の 最悪＝★気づかない★
 *
 *  ★★この 道具が 見る 物★★
 *    ⑴今の 枝の ★GitHub 側の 最後の 検査★（`gh run list`）
 *    ⑵その SHA から ★origin/<枝> までに 何本 積んだか★
 *    ⇒★1本でも 積んで いれば 赤★（＝その 分は 検査されて いない）
 *
 *  ★★この 道具が 言えない 事★★
 *    ・★網が 要ります★（`gh` が 通らない 時は ★「測れない」と 言って 赤★
 *      ＝★測れない 物を 緑に しません★）
 *    ・★検査の 中身★は 見て いません（走ったか／通ったか だけ）
 *    ・★手元の 総なめは この 穴を 1つも 塞ぎません★
 *
 *  使い方: node scripts/ci-ga-hashitta-ka.mjs [--枝=<名前>]
 *  終わり値: 0 ... 積み残し 0本 かつ 最後の 検査が success
 *            1 ... 積み残しが 在る／検査が 赤／測れない
 *
 *  ★★作った その場で 3通り 確かめました（2026-09-19）★★
 *    今の 枝（積み残し 4本） ......... 赤 ／ 終わり値 ★1★
 *    main（積み残し 0・最後 success）.. 緑 ／ 終わり値 ★0★
 *    検査が 1本も 無い 枝 ............ 赤 ／ 終わり値 ★1★
 *
 *  ★★測る 時に 踏んだ 穴（★書いて おきます★）★★
 *    最初 `node ... | tail -4; echo $?` で 測り、★終わり値 0★ と 読みました。
 *    ⇒★それは `tail` の 終わり値です★（`node` の では ない）
 *    ⇒★危うく 自分の 道具に 無い 欠陥を 報せる ところでした★
 *    ⇒ 記憶「★`tail`/`head`/`grep` は 測り道具を 壊す 道具★」
 *    ⇒★★終わり値は 管を 通さずに 測る★★
 */
import { execFileSync } from 'node:child_process';

const 引数 = process.argv.slice(2);
const 枝指定 = (引数.find((a) => a.startsWith('--枝=')) || '').slice(4);

function 走らせる(cmd, args) {
  return execFileSync(cmd, args, { encoding: 'utf-8', stdio: ['ignore', 'pipe', 'pipe'] }).trim();
}

console.log('');
console.log('[ci-ga-hashitta-ka] 「緑」と 言う 前に GitHub 側の 検査を 数える');

let 枝;
try { 枝 = 枝指定 || 走らせる('git', ['branch', '--show-current']); }
catch (e) { console.log('  NG   枝の 名前が 取れません ... ' + (e && e.message)); process.exit(1); }
console.log('  枝 ... ' + 枝);

/* ★GitHub 側を 引く★（★網が 要る／通らなければ 赤★） */
let 生;
try {
  生 = 走らせる('gh', ['run', 'list', '--branch', 枝, '--limit', '20',
    '--json', 'conclusion,status,headSha,workflowName,createdAt']);
} catch (e) {
  console.log('  NG   ★GitHub 側を 引けません★（網か gh の 認証）');
  console.log('       ' + String((e && e.message) || '').split(String.fromCharCode(10))[0]);
  console.log('       ⇒ ★測れない 物を 緑に しません＝赤に します★');
  process.exit(1);
}

let 走り;
try { 走り = JSON.parse(生); } catch (e) { console.log('  NG   出しが 読めません'); process.exit(1); }
if (!走り.length) {
  console.log('  NG   ★この 枝に 検査が 1本も 在りません★');
  console.log('       ⇒ ★PR を 立てて ください★（枝は PR が 開いて いる 間だけ 検査されます）');
  process.exit(1);
}

/* ★CI（名前に かかわらず 全部）の うち 一番 新しい 物★ */
const 最後 = 走り[0];
const CIの = 走り.filter((r) => r.workflowName === 'CI');
const 最後のCI = CIの[0] || null;
console.log('  ★GitHub 側の 最後の 検査★ ... ' + (最後.conclusion || 最後.status)
  + '  ' + 最後.headSha.slice(0, 7) + '  ' + 最後.workflowName + '  ' + 最後.createdAt);
if (最後のCI && 最後のCI !== 最後) {
  console.log('  ★そのうち CI の 最後★ ......... ' + (最後のCI.conclusion || 最後のCI.status)
    + '  ' + 最後のCI.headSha.slice(0, 7) + '  ' + 最後のCI.createdAt);
}

let 赤 = 0;

/* ⑴積み残しを 数える */
const 基 = (最後のCI || 最後).headSha;
let 積み残し = [];
try {
  const 出 = 走らせる('git', ['log', '--oneline', 基 + '..origin/' + 枝]);
  積み残し = 出 ? 出.split(String.fromCharCode(10)) : [];
} catch (e) {
  console.log('  NG   ★手元に ' + 基.slice(0, 7) + ' が 在りません★（`git fetch` して ください）');
  process.exit(1);
}
if (積み残し.length) {
  赤++;
  console.log('  NG   ★★検査を 通って いない 押しが ' + 積み残し.length + '本 在ります★★');
  for (const l of 積み残し) console.log('       ' + l);
  console.log('       ⇒ ★PR を 立てれば その 場で 検査が 走ります★');
} else {
  console.log('  ok   ★検査を 通って いない 押しは 0本★');
}

/* ⑵最後の 検査の 色 */
const 色 = (最後のCI || 最後).conclusion;
if (色 === 'success') console.log('  ok   ★最後の 検査は 緑★');
else { 赤++; console.log('  NG   ★★最後の 検査は ' + (色 || (最後のCI || 最後).status) + '★★'); }

console.log('');
console.log('ci-ga-hashitta-ka: ' + (赤 ? '★' + 赤 + '件 赤★' : '緑'));
process.exit(赤 ? 1 : 0);
