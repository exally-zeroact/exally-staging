/* ci-coverage.test.mjs — ★CIから黙って外れているテストが無いか見張る★
 *
 * なぜ必要か: 2026-08-01 の統合で、対象機能がこのツリーに無いテストを1本だけCIから外した。
 *   CIから外れたテストは、放っておけば誰も気づかないまま永久に死んだファイルになる。
 *   「外す」こと自体は時にあってよいが、【外れている物の一覧が明示され、戻す条件が書かれている】
 *   状態でなければならない。それを機械で強制する。
 *
 * 判定:
 *   テストらしきファイル（*.test.js / *.test.mjs / 既知のハーネス）は、次のいずれかであること
 *     ① .github/workflows/ci.yml が直接 node で回している
 *     ② tests/run.js または kyuyo/tests/run.js が回している
 *     ③ 下の EXCLUDED に「理由」と「戻す条件」つきで載っている
 *   どれにも当てはまらない物が1つでもあれば赤。
 *
 * 使い方: node tests/ci-coverage.test.mjs
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');

// ★CIから外しているテスト（ここに載っていない除外は赤になる）
/* ★CIから外しているテスト（ここに載っていない除外は赤になる）
   ★2026-09-05★ 給与(kyuyo/)を Exally から 外した（給与は Rakunally）。
     ★唯一の 除外だった kyuyo/tests/exally-login.test.mjs も 一緒に 出て行った★
     ⇒★空に する★＝★除外が 1本も 無い のが 今の 正しい 姿★
     ⇒★増えたら 下の「理由と 戻す条件」で 必ず 赤に なる★（★空でも 見張りは 効く★） */
const EXCLUDED = {
  /* ★★2026-09-06 に 1本 足した★★＝★わざと 赤のまま★の 試験
     ★空が 正しい 姿★だったので、足す 私が 理由を 一番 はっきり 書く。 */
  'tests/hozon-de-kieru.test.mjs': {
    reason: 'うちで 足したシートは 書き出すと 消える事を 実測で 固定した 物。'
      + '★直っていないので 走らせると 必ず 赤★（出来ていない事を 隠さない／でも CI を 常態で 赤に しない）。'
      + '手で 走らせる … node tests/hozon-de-kieru.test.mjs',
    restoreWhen: 'zip へ 足す 直しが 入って ★緑に なった 日★（その時 この 行を 消して run.js に 載せる）',
  },
};

// テストではない道具（実行されなくてよい物）。ここも理由つきで明示する。
const NOT_TESTS = {
  'tests/run.js': 'ランナー本体',
  'tests/fake-supa.js': 'テスト用のSupabaseモック（他テストが読む部品）',
  'tests/shiken-no-kuchi.mjs': '★実の 倉庫に 触る 道具（live-seed・live-roundtrip）が 使う 試験用の 口の 認証と 門★（2026-10-10）。メールは repo の 外の cred だけ（テストではない）',
  'tests/_tsukurimono.mjs': '★実物の 本の 字を 作り物に 替える 部品★（2026-10-10）。実物を 開く 試験（kiridashi-jitsubutsu・kirikae ⑤）が 読む（テストではない）。対応表は repo の 外',
  'tests/repo-supa.mjs':'このリポジトリの接続先(js/supa-config.js)を返す部品。実DBに触る道具が読む（テストではない）',
  'tests/zairyou.mjs': '★紙の 材料（#材料 の 行）を 読んで 表に 置く 部品★。突き合わせの 道具と 試験が 読む（テストではない）。★材料の 読み方を 道具ごとに 書くと、材料の 無い 紙で 空の 表を 押して ★偽の 負け★が 出る★（2026-09-14 に 実際に PRODUCT で 起きた）ので 1か所に 切り出した',
  'tests/betsu-atsukai.mjs': '★実Excel と 同じに しない と 決めた 物の 名簿★（今は INFO 1つ）。formula-yosoku.test.mjs と 借り物を外す皮の突き合わせが 読む（テストではない）。★同じ 名簿を 2つ 持つと「合っている」の 意味が ずれる★ので 切り出した（2026-09-14）',
  'tests/monosashi.mjs': '★式のほどき方と 答えの見方（1e-9の許し・誤りの比べ方）の 部品★。formula-kane.test.mjs と 借り物を外す皮の突き合わせが 読む（テストではない）。★物差しを 2本に しない為に 切り出した★（2026-09-14）',
  'tests/dbtest-seed.mjs': 'DB-testに種データを入れる手動ツール（CIから叩かない）',
  'tests/live-seed.mjs': '実DBに種を入れる手動ツール（CIから叩かない）',
  'tests/live-roundtrip.mjs': '実DBに触る手動確認ツール（CIから叩かない）',
  'tests/xlsx-harness/build-libre-input.mjs': 'LibreOffice入力を作る手動ツール',
  'tests/xlsx-harness/collect-libre.mjs': 'LibreOfficeの結果を集める手動ツール',
  'tests/xlsx-harness/run-exally.mjs': 'ハーネスの実行部品（compare.mjs から使う）',

};

let pass = 0, fail = 0;
function T(n, fn) { try { fn(); pass++; console.log('  ✓ ' + n); } catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + (e && e.message)); } }

function listTestFiles() {
  const out = [];
  const walk = (rel) => {
    const dir = path.join(ROOT, rel);
    if (!fs.existsSync(dir)) return;
    for (const f of fs.readdirSync(dir)) {
      const p = path.join(dir, f);
      if (fs.statSync(p).isDirectory()) { walk(path.posix.join(rel, f)); continue; }
      if (/\.(m?js)$/.test(f)) out.push(path.posix.join(rel, f));
    }
  };
  walk('tests');
  walk('kyuyo/tests');
  return out.sort();
}

/* ★2026-09-03 に 直した★＝★ci.yml しか 見ていなかった★。
   ★週1の 回（.github/workflows/webkit.yml）で 走る 試験を「宙に浮いている」と 誤って 赤に した★
   （実測＝comment-mark-webkit.mjs／karimono.test.mjs の 2本）。
   ⇒★.github/workflows の yml を 全部 見る★（★どの 回で 走っても「走っている」★）
   ⇒★見た yml の 本数も 出す★＝★0本 見て 0件★を 見破れるように（今日の 決まり） */
const wfDir = path.join(ROOT, '.github/workflows');
const wfFiles = fs.existsSync(wfDir)
  ? fs.readdirSync(wfDir).filter((f) => /\.ya?ml$/.test(f)).sort()
  : [];
const ci = wfFiles.map((f) => fs.readFileSync(path.join(wfDir, f), 'utf8')).join(String.fromCharCode(10));
const ciRuns = new Set((ci.match(/node\s+([A-Za-z0-9/._-]+)/g) || []).map(s => s.replace(/^node\s+/, '')));

function runnerList(runnerRel) {
  const p = path.join(ROOT, runnerRel);
  if (!fs.existsSync(p)) return [];
  const src = fs.readFileSync(p, 'utf8');
  const base = path.posix.dirname(runnerRel);
  return (src.match(/'\.?\/?[A-Za-z0-9/._-]+\.(?:test\.)?m?js'/g) || [])
    .map(s => s.replace(/'/g, '').replace(/^\.\//, ''))
    .map(f => path.posix.join(base, f));
}
/* ★在る 一覧だけ 読む★（給与が 出て行き kyuyo/tests/run.js は 無い） */
const runnerPaths = ['tests/run.js', 'kyuyo/tests/run.js']
  .filter((p) => fs.existsSync(path.join(ROOT, p)));
if (!runnerPaths.length) throw new Error('★走らせる 一覧が 1本も 無い＝この 検査が 空振り★');
const viaRunner = new Set(runnerPaths.flatMap((p) => runnerList(p)));

console.log('\n[ci-coverage] CIから黙って外れているテストが無いか');
/* ★見た 本数も 出す★＝★0本 見て 0件★を 見破る為（2026-09-03 の 決まり） */
console.log('  （見た ワークフロー ' + wfFiles.length + '本＝' + wfFiles.join('・') + '）');

const files = listTestFiles();
const covered = [], excluded = [], orphan = [];
for (const f of files) {
  if (NOT_TESTS[f]) continue;
  if (ciRuns.has(f) || viaRunner.has(f)) covered.push(f);
  else if (EXCLUDED[f]) excluded.push(f);
  else orphan.push(f);
}

T('★CIからも各ランナーからも呼ばれていないテストが無い（除外リストに載っていない物）', function () {
  if (orphan.length) throw new Error('宙に浮いているテスト:\n   - ' + orphan.join('\n   - '));
});
T('★除外リストの各項目に「理由」と「戻す条件」が書かれている', function () {
  for (const [f, e] of Object.entries(EXCLUDED)) {
    if (!e.reason || e.reason.length < 20) throw new Error(f + ': reason が不十分');
    if (!e.restoreWhen || e.restoreWhen.length < 10) throw new Error(f + ': restoreWhen（戻す条件）が無い');
    if (!fs.existsSync(path.join(ROOT, f))) throw new Error(f + ': 除外リストにあるがファイルが無い（消したなら除外リストからも消す）');
  }
});
/* ★2026-09-05★ 給与が 出て行き、除外は ★0本★に なった。
   ★数の 決め打ち（1本）を やめる★＝★0本が 正しい 姿★
   ⇒★増えた時に 気づく★のが 目的なので、
     ★「除外が 在るなら 理由と 戻す条件が 書いてある」（上の 検査）★で 足りる。
   ⇒ ここでは ★増えていない事★だけ 見る（0本 なら 緑・1本でも 増えたら 上で 中身を 見る）。 */
T('★除外は 増えていない（0本＝給与が 出て行った後の 正しい 姿）', function () {
  const n = Object.keys(EXCLUDED).length;
  if (n > 0) {
    console.log('  ★除外が ' + n + '本 在ります★ … ' + Object.keys(EXCLUDED).join(', '));
    console.log('  ⇒★理由と 戻す条件は 上の 検査が 見ています★');
  }
});
T('検査が空振りしていない（テストファイルを実際に数えている）', function () {
  if (covered.length < 50) throw new Error('CIが回しているテストが少なすぎます: ' + covered.length);
});
/* ★0本 見て 0件★を 見破る（2026-09-03 の 決まり）
   ＝★ワークフローを 1本も 読めていないのに「全部 走っている」と 言わない★ */
T('★ワークフローを 実際に 読んでいる（今 ' + wfFiles.length + '本）★', function () {
  if (wfFiles.length < 2) throw new Error('ワークフローを ' + wfFiles.length + '本しか 読めていません＝★探し方が 悪い★');
  if (!/webkit\.yml/.test(wfFiles.join(' '))) throw new Error('週1の 回（webkit.yml）を 読めていません');
});

console.log(`\n── 実測 ──`);
console.log(`  テストファイル: ${files.length}本（うち道具 ${Object.keys(NOT_TESTS).length}本は対象外）`);
console.log(`  CI/ランナーが回している: ${covered.length}本`);
console.log(`  明示的にCIから外している: ${excluded.length}本`);
excluded.forEach(f => {
  console.log(`   - ${f}`);
  console.log(`     理由    : ${EXCLUDED[f].reason}`);
  console.log(`     戻す条件: ${EXCLUDED[f].restoreWhen}`);
});
console.log(`  宙に浮いている: ${orphan.length}本` + (orphan.length ? '\n   - ' + orphan.join('\n   - ') : ''));
console.log('\n' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
