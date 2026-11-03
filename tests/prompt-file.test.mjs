/* prompt-file.test.mjs — ★AIの 頭は ファイルから／台帳と ずれたら 赤★（2026-09-05）
 *
 *  ★司さん★「★AIの構造は ファイルにして 更新できるような 設計に しとけよ★」
 *
 *  ★なぜ 要るか（実測）★
 *    前は api/claude.js に ★べた書き★だった。台帳（lib/formula-extra.js）と 二重管理。
 *    ⇒★手書きの「使えない関数」22個のうち ★17個が 間違い★★
 *       ・足してあるのに「使えない」…4個（TOCOL/TOROW/CHOOSEROWS/CHOOSECOLS）
 *       ・動かないのに 1つも 教えていない…13個（LAMBDA/LET/MAP/REDUCE/…）
 *         ★★その 13個の うち YEN は 数え間違い★★（同じ日の 夜に 実測）
 *           ＝★打てば ¥1,235 が 出る★（Exally が 本名 DOLLAR に 直す）
 *           ⇒★別名で動く★の 棚へ 移した。この 試験も それを 見る（下の self-test）
 *    ⇒ しかも latest の 決まりは「★LET・LAMBDA等を 積極的に 提案★」
 *       ＝★AIが 動かない 式を 勧めていた★
 *
 *  使い方: node tests/prompt-file.test.mjs
 *          node tests/prompt-file.test.mjs --self-test
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(import.meta.url);
let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ✓ ' + n); } catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + (e && e.message)); } };

/** ★純関数★＝前置きの 字と 台帳から「食い違い」を 返す（self-test で 作り物を 通せる） */
export function 食い違い(前置き, 台帳) {
  /* ★★「別名で動く」も 動く 側★★（2026-09-05 に 足した）
     ★実際に 起きた 事★ … YEN が 「足さない」の 棚に 置いてあった
       ⇒ この 紙で ★「Exally内では まだ 動かない」の 列★に 並んでいた
       ⇒★実際は ¥1,235 が 出るのに AIが 客に「動きません」と 言う★所だった
     ⇒★ここに 足さないと、同じ 事が また 通る★ */
  const 足す = (台帳.足す || [])
    .concat(Object.keys(台帳.別のlibで足す || {}))
    .concat(Object.keys(台帳.別名で動く || {}));
  const 足さない = Object.keys(台帳.足さない || {});
  const 出 = { 使えるのに使えないと言う: [], 動かないのに黙っている: [] };
  for (const f of 足す) {
    /* ★正規表現を 組み立てない★＝関数名に . が 入る（MODE.MULT）ので 逃がしが 要り、
       ここで 1回 壊した（2026-09-05）。★字を そのまま 探す★方が 確実。 */
    if (前置き.indexOf('★' + f + '★') >= 0) 出.使えるのに使えないと言う.push(f);
  }
  for (const f of 足さない) {
    if (前置き.indexOf(f) < 0) 出.動かないのに黙っている.push(f);
  }
  return 出;
}

if (process.argv.includes('--self-test')) {
  console.log('\n[prompt-file --self-test] わざと壊して赤になるか');
  const 台帳 = { 足す: ['TAKE'], 足さない: { LAMBDA: '無い' } };
  T('★動かない関数が 前置きに 無ければ 見つける', () => {
    const r = 食い違い('つかえる関数: TAKE', 台帳);
    if (r.動かないのに黙っている.length !== 1) throw new Error('見つけていない');
  });
  T('★動く関数を「使えない」と 書いていたら 見つける', () => {
    const r = 食い違い('動かない: ★TAKE★ / LAMBDA', 台帳);
    if (r.使えるのに使えないと言う.length !== 1) throw new Error('見つけていない');
  });
  T('★正しければ 何も 出さない', () => {
    const r = 食い違い('つかえる: TAKE / 動かない: ★LAMBDA★', 台帳);
    if (r.使えるのに使えないと言う.length || r.動かないのに黙っている.length) throw new Error('誤検知');
  });
  /* ★★2026-09-05 に 実際に 通してしまった 形★★
     YEN は 打てば ¥1,235 が 出るのに 「動かない」の 列に 並んでいた
     ⇒★別名で動く を 数えていないと、この 見張りは 素通りする★ */
  T('★★「別名で動く」物を「動かない」と 書いていたら 見つける★★', () => {
    const 台2 = { 足す: ['TAKE'], 別名で動く: { YEN: 'DOLLAR に 直す' }, 足さない: { LAMBDA: '無い' } };
    const r = 食い違い('動かない: ★LAMBDA★ / ★YEN★', 台2);
    if (r.使えるのに使えないと言う.indexOf('YEN') < 0) {
      throw new Error('★見つけていない＝AIが 客に「YEN は 動きません」と 言う★');
    }
  });
  T('★別名で動く を 正しく 動く 側に 書いていれば 素通り', () => {
    const 台2 = { 足す: ['TAKE'], 別名で動く: { YEN: 'DOLLAR に 直す' }, 足さない: { LAMBDA: '無い' } };
    const r = 食い違い('動く: TAKE / YEN … DOLLAR に 直す / 動かない: ★LAMBDA★', 台2);
    if (r.使えるのに使えないと言う.length) throw new Error('★誤検知★: ' + r.使えるのに使えないと言う.join(','));
  });
  console.log('\n' + pass + ' passed, ' + fail + ' failed');
  process.exit(fail ? 1 : 0);
}

/* ══ 本番 ═══════════════════════════════════════════════════ */
if (!process.env.ANTHROPIC_API_KEY) process.env.ANTHROPIC_API_KEY = 'test-dummy-key';
const handler = require_(path.join(ROOT, 'api/claude.js'));
const 台帳 = require_(path.join(ROOT, 'lib/formula-extra.js')).数える();
const 部屋 = path.join(ROOT, 'prompt');

console.log('\n[prompt-file] AIの 頭は ファイルから 出来ているか');

T('★prompt/ の 4本が 在って 空でない', () => {
  for (const f of ['base.md', 'kansuu.md', 'kyoutsuu.md', 'version.md']) {
    const p = path.join(部屋, f);
    if (!fs.existsSync(p)) throw new Error('無い: ' + f);
    if (!fs.readFileSync(p, 'utf8').trim()) throw new Error('空: ' + f);
  }
});

T('★api/claude.js に 前置きを べた書きしていない', () => {
  const src = fs.readFileSync(path.join(ROOT, 'api/claude.js'), 'utf8');
  for (const 悪い of ['const SYSTEM_PROMPT_BASE = ', 'const EXALLY_UNSUPPORTED = ']) {
    if (src.indexOf(悪い) >= 0) throw new Error('★べた書きが 戻っています★: ' + 悪い);
  }
});

T('★★前置きが 台帳と 食い違っていない（動かない関数を 勧めない）★★', () => {
  const p = handler.__buildPromptParts({ group: 'latest', name: 'Excel 365' });
  const 全 = p.共通 + p.版ごと;
  const r = 食い違い(全, 台帳);
  if (r.使えるのに使えないと言う.length) {
    throw new Error('★足してあるのに「使えない」と 書いている★: ' + r.使えるのに使えないと言う.join(', '));
  }
  if (r.動かないのに黙っている.length) {
    throw new Error('★動かないのに 1つも 教えていない★: ' + r.動かないのに黙っている.join(', ')
      + '\n   → node scripts/make-prompt.mjs を 走らせて commit してください');
  }
});

/* ★★README の 表に 書いた ファイルが 本当に AIへ 渡っているか★★（2026-09-06）
   ★実際に 起きた 事★ … 表には `nayami.md` `kiite-ageru.md` を「入れる」と 書いてあったのに
     ★2本とも 存在せず、会議で 決めた「聞いてあげる」は AIに 1文字も 渡っていなかった★
     `kotae-kata.md` は ★名前が 違った★（実物は kyoutsuu.md）
   ⇒★表に 書いただけでは 入らない★＝★書いた 物が 実在し、前置きに 出るか まで 見る★ */
T('★★README の 表の ファイルが 実在して 前置きに 出ている★★', () => {
  const md = fs.readFileSync(path.join(部屋, 'README.md'), 'utf8');
  const 表の = [...md.matchAll(/^\|\s*`([a-z0-9-]+\.md)`/gmi)].map((m) => m[1]);
  if (表の.length < 4) throw new Error('★README の 表が 読めない★（検査が 空振り）: ' + 表の.length + '個');
  const 無い = 表の.filter((f) => !fs.existsSync(path.join(部屋, f)));
  if (無い.length) {
    throw new Error('★表に 在るのに ファイルが 無い★: ' + 無い.join(' / ')
      + '\n   → ★表に 書いただけでは AIに 渡りません★');
  }
  const p = handler.__buildPromptParts({ group: 'latest', name: 'Excel 365' });
  const 全 = p.共通 + p.版ごと;
  const 渡っていない = [];
  for (const f of 表の) {
    const 中 = fs.readFileSync(path.join(部屋, f), 'utf8');
    const m = /```\r?\n([\s\S]*?)\r?\n```/.exec(中);
    const 頭 = (m ? m[1] : 中).trim().split('\n').find((l) => l.trim().length > 8);
    if (頭 && 全.indexOf(頭.trim()) < 0) 渡っていない.push(f);
  }
  if (渡っていない.length) {
    throw new Error('★ファイルは 在るのに 前置きに 出ていない★: ' + 渡っていない.join(' / ')
      + '\n   → api/claude.js の 要る[] と buildPromptParts に 足してください');
  }
});

/* ★★版ごとの 決まり＝★2つの 別の 問い★を 混ぜない★★（2026-09-05 指示役の 裁定）
 *  ★問いA「お客さんの Excel で 動くか」★ … ★版で 変わる★（Excel 2016 に XLOOKUP は 無い）
 *      ⇒ `prompt/version.md` は ★Aの ファイル★（1行目「版ごとの 言い方」）
 *      ⇒★だから latest/newer/older に 関数名が 在るのは ★正しい★★
 *  ★問いB「Exally の 中で 動くか」★ … ★`prompt/kansuu.md`（機械が 台帳から 作る）★
 *
 *  ★私が 1回 やり過ぎた（2026-09-05）★
 *    Aの 所を 全部「上の 一覧に 従う」に した
 *    ⇒★一覧は Bの 話★なので、★Excel 2016 の 客に XLOOKUP を 勧める★所だった
 *    ⇒★消して よいのは「AとBが ぶつかっている 所」だけ★
 *
 *  ★本当に 嘘だった 2か所★
 *    ①`exally_only`（Excel を 持っていない）… ★丸ごと Bの 話★
 *       ⇒ ここに 動かない関数名が 在れば ★嘘★
 *    ②latest の「Exally内でも 同じ数式が 動くよ」を ★無条件で★ 言わせる 行
 *       ⇒ Excel 365 に LAMBDA は 在る（Aは 正しい）が
 *         ★Exally では 動かない★のに「同じ数式が 動く」＝★嘘★
 *
 *  ★見張りを ここに 絞る 理由★
 *    ★正しい 記述（older の「XLOOKUP は 使わない」）を 赤に する 見張りは 切られる★
 *    ＝★いつも真の 逆＝★正しい 物を 赤に する★★も 同じくらい 悪い。 */
T('★★Excelを 持っていない 版に 動かない関数の 名前が 出ていない★★', () => {
  const 名 = Object.keys(台帳.足さない);
  const p = handler.__buildPromptParts({ group: 'exally_only', name: 'Excel持ってない' });
  const 悪い = 名.filter((n) => p.版ごと.indexOf(n) >= 0);
  if (悪い.length) {
    throw new Error('★Excelを 持っていない 客に 動かない関数を 勧めています★: ' + 悪い.join(', ')
      + '\n   → exally_only は ★丸ごと Exally の 話★。正本は prompt/kansuu.md');
  }
});

T('★★latest に「Exally で 動く 物に 限る」の 絞りが 在る★★', () => {
  /* ★向きが ①と 逆★（2026-09-05 指示役の 裁定）
     ①exally_only の 関数名 … ★決まった 名前★（LAMBDA は LAMBDA としか 書けない）
        ⇒★悪い字 探しで よい★
     ②ここ … ★言い回し★（「同じ数式が 動くよ」「そのまま 使えるよ」「同じように 動くよ」…）
        ⇒★悪い 一文を 探すと ★書き直しで 死ぬ★★＝すり抜けて ★赤に ならない★
        ⇒★★良い字（絞りの 言葉）が 在るか を 数える★★
        ⇒ 誰が どう 書き直しても ★絞りが 消えたら 必ず 赤★
     ★なぜ latest だけか★ … Excel 365 は ★LAMBDA が 本当に 在る★ので
        「Excel では 動く／Exally では 動かない」が ★一番 ぶつかる 版★。 */
  const p = handler.__buildPromptParts({ group: 'latest', name: 'Excel 365' });
  const 絞り = ['限り', '一覧', '断る'];
  const 在る = 絞り.filter((w) => p.版ごと.indexOf(w) >= 0);
  if (!在る.length) {
    throw new Error('★latest に 絞りの 言葉が 1つも 在りません★（探した: ' + 絞り.join(' / ') + '）'
      + '\n   → ★「Exally内でも 動く 関数に 限り…」「動かない 物は 一覧のとおり 断る」★のように'
      + '\n     ★Exally で 動く 物に 限る★ と 書く（Excel 365 に 在っても Exally で 動くとは 限らない）');
  }
});

T('★どの 版でも 前置きが 組める（版ごとの 決まりが 入る）', () => {
  for (const g of ['latest', 'newer', 'older', 'online', 'exally_only']) {
    const p = handler.__buildPromptParts({ group: g, name: 'X' });
    if (p.版ごと.indexOf('ユーザーのExcel環境') < 0) throw new Error(g + ' の 決まりが 入っていない');
    if ((p.共通 + p.版ごと).length < 1500) throw new Error(g + ' の 前置きが 短すぎる');
  }
});

T('★prompt が 1本でも 読めなければ 500・合言葉 shitaku_tarinai（AIのせいに しない）', () => {
  const 分ける = handler.__失敗を分ける;
  const e = Object.assign(new Error('prompt が 読めません'), { 支度足りない: true });
  const r = 分ける(e);
  if (r.status !== 500 || r.合言葉 !== 'shitaku_tarinai') {
    throw new Error('★' + JSON.stringify(r) + '★（ai_shippai だと「AI側の問題です」と 嘘を 言う）');
  }
});

T('★Vercel に prompt/ を 載せる 設定が 在る（配信で 読めないと 全部 500）', () => {
  const v = JSON.parse(fs.readFileSync(path.join(ROOT, 'vercel.json'), 'utf8'));
  const inc = v.functions && v.functions['api/claude.js'] && v.functions['api/claude.js'].includeFiles;
  if (!inc || String(inc).indexOf('prompt') < 0) {
    throw new Error('vercel.json の functions.includeFiles に prompt/** が 在りません');
  }
});

T('★機械が 作る ファイルは 台帳と 一致（--check が 通る）', () => {
  const r = require_('node:child_process').spawnSync(process.execPath,
    [path.join(ROOT, 'scripts/make-prompt.mjs'), '--check'], { encoding: 'utf8' });
  if (r.status !== 0) throw new Error('make-prompt --check が 赤: ' + (r.stdout || '').trim());
});

const p0 = handler.__buildPromptParts({ group: 'latest', name: 'Excel 365' });
console.log('\n── 実測 ──');
console.log('  前置き … 共通 ' + p0.共通.length + '字 ／ 版ごと ' + p0.版ごと.length + '字');
console.log('  台帳 … 動く ' + (台帳.足す.length + Object.keys(台帳.別のlibで足す || {}).length
  + Object.keys(台帳.別名で動く || {}).length)
  + '個 ／ 動かない ' + Object.keys(台帳.足さない).length + '個');
console.log('\n' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
