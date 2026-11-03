/* ban-wo-okuru.test.mjs — ★お客さんの Excel の 版を 本当に 送っているか★（2026-09-05）
 *
 *  ★★会議で 見つけた 穴 B（2026-09-05・6人の 会議）★★
 *    画面は AIへ ★`{message, history}` の 2つしか 送っていなかった★
 *    サーバ `api/claude.js:40` … `VERSION_MAP[versionKey] || VERSION_MAP['excel_365']`
 *    ⇒★★全員が Excel 365 扱い★★
 *      ・Excel 2016 / 2019 の 客にも ★XLOOKUP・FILTER を 出していた★
 *      ・365扱いだと「LET・LAMBDA等を 積極的に 使え」も 効く
 *    ⇒★`prompt/version.md`（版ごとの 言い方）は ★誰にも 効いていなかった★★
 *      ＝そこの 文だけ 直しても ★客には 何も 変わらない★
 *
 *  ★★ここを 字で 見張る 理由★★
 *    本当は 実ブラウザで 押して 中身を 見るのが 一番だが、
 *    ★送っているか だけは 字で 確かめられる★（送っていなければ 何を しても 届かない）
 *    ⇒★足りない 分は はっきり 書く★＝下の「この 機械は 半分」を 読むこと
 *
 *  使い方: node tests/ban-wo-okuru.test.mjs
 *          node tests/ban-wo-okuru.test.mjs --self-test
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
/* ★★注記(コメント)を 外してから 見る★★（2026-09-05・この 見張りが 1回 素通りした）
   ★実際に 起きた事★ … 私が 書いた ★説明の 中に `excelVersion` が 何度も 出る★
     ⇒ わざと 名前を `version` に ずらして 壊したのに ★緑のまま★
     ⇒★見張りは 自分の 説明文を 読んで「在る」と 言っていた★
   ⇒★repo に 既に 在る `scripts/lib/chuki.mjs` の `注記を外す()` を 使う★ */
const { 注記を外す } = await import(pathToFileURL(path.join(ROOT, 'scripts/lib/chuki.mjs')).href);
let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ✓ ' + n); } catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + (e && e.message)); } };

/** ★純関数★＝画面の 字と サーバの 字から「版が 届くか」を 返す
 *  ★名前が 合っていないと 届かない★＝両方から 名前を 取って 突き合わせる */
export function 版は届くか(画面, サーバ) {
  const 出 = { 画面が送る名: null, サーバが読む名: null, 届く: false, なぜ: '' };
  /* サーバ … `const { message, history, excelVersion } = req.body` の 3つ目 */
  const s = サーバ.match(/req\.body[\s\S]{0,40}/);
  const b = サーバ.match(/const\s*\{([^}]*)\}\s*=\s*req\.body/);
  if (b) {
    const 名たち = b[1].split(',').map((x) => x.trim()).filter(Boolean);
    出.サーバが読む名 = 名たち.find((x) => /version/i.test(x)) || null;
  }
  /* 画面 … AIへ 出す 荷物に その 名前を 入れているか */
  if (出.サーバが読む名) {
    const 型 = new RegExp('\\b' + 出.サーバが読む名 + '\\b');
    if (型.test(画面)) 出.画面が送る名 = 出.サーバが読む名;
  }
  if (!出.サーバが読む名) { 出.なぜ = 'サーバが 版を 読んでいない'; return 出; }
  if (!出.画面が送る名) { 出.なぜ = '画面が ' + 出.サーバが読む名 + ' を 送っていない'; return 出; }
  出.届く = true;
  return 出;
}

if (process.argv.includes('--self-test')) {
  console.log('\n[ban-wo-okuru --self-test] わざと壊して赤になるか');
  const サーバ = 'const { message, history, excelVersion } = req.body || {};';
  /* ★2026-09-05 まで 本番が こうだった★ */
  T('★★画面が 送っていなければ 赤（09-05 まで これだった）★★', () => {
    const r = 版は届くか('body: JSON.stringify({ message: m, history: h })', サーバ);
    if (r.届く) throw new Error('★送っていないのに 届くと 言った★');
    if (!/画面が excelVersion を 送っていない/.test(r.なぜ)) throw new Error('理由が 違う: ' + r.なぜ);
  });
  T('★送っていれば 通る', () => {
    const r = 版は届くか('body.excelVersion = getBookVer();', サーバ);
    if (!r.届く) throw new Error('通らない: ' + r.なぜ);
  });
  /* ★名前が ずれたら 届かない★＝「送っているつもり」で 落ちる 一番 多い 形 */
  T('★★名前が ずれていたら 赤（version だけ 送っても 届かない）★★', () => {
    const r = 版は届くか('body.version = getBookVer();', サーバ);
    if (r.届く) throw new Error('★名前が 違うのに 届くと 言った★');
  });
  T('★サーバが 読まなくなったら 赤（片方だけ 消しても 気づく）', () => {
    const r = 版は届くか('body.excelVersion = v;', 'const { message, history } = req.body || {};');
    if (r.届く) throw new Error('★サーバが 読んでいないのに 届くと 言った★');
  });
  console.log('\n' + pass + ' passed, ' + fail + ' failed');
  process.exit(fail ? 1 : 0);
}

/* ══ 本番 ═══════════════════════════════════════════════════ */
const 画面 = 注記を外す(fs.readFileSync(path.join(ROOT, 'book.html'), 'utf8'));
const サーバ = 注記を外す(fs.readFileSync(path.join(ROOT, 'api/claude.js'), 'utf8'));
/* ★注記を 外すと `data-ver="…"` も 生きたまま★＝HTML の 属性は 注記では ない */
const 画面の生 = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf8');
const r = 版は届くか(画面, サーバ);

console.log('\n[ban-wo-okuru] お客さんの Excel の 版が AIまで 届くか');

T('★★画面が 版を 送っている（届かないと 版ごとの 言い方が 全部 死ぬ）★★', () => {
  if (!r.届く) {
    throw new Error('★' + r.なぜ + '★'
      + '\n   → ★全員が Excel 365 扱いに なります★'
      + '\n   → 2016・2019 の 客に XLOOKUP・FILTER を 出します'
      + '\n   → prompt/version.md の 直しは ★1人にも 効きません★');
  }
});

T('★1か所で 付けている（呼ぶ側で 付け忘れない）', () => {
  /* ★_aiFetch は 3か所から 呼ばれている★（説明・チャット・レシピ）
     ⇒ 呼ぶ側で 付けると ★1か所 忘れた時に 黙って 365 に なる★ */
  const i = 画面.indexOf('async function _aiFetch(');
  if (i < 0) throw new Error('_aiFetch が 見つからない');
  const 中 = 画面.slice(i, i + 1600);
  if (中.indexOf('excelVersion') < 0) {
    throw new Error('★_aiFetch の 中で 付けていない★＝呼ぶ側 頼みに なっています');
  }
});

T('★版を 選ぶ 画面が 生きている（選べないと 送る物が 無い）', () => {
  if (画面.indexOf('exally_excel_version') < 0) throw new Error('版の 置き場が 無い');
  if (画面.indexOf('function getBookVer') < 0) throw new Error('getBookVer が 無い');
});

T('★★画面が 出す 版を サーバが 全部 知っている★★', () => {
  /* ★片方だけ 増やすと 黙って 365 に なる★
     ＝サーバは 知らない 版を `|| VERSION_MAP['excel_365']` で 倒す
     ⇒「Excel 2016」を 選んだのに 365の 答えが 出る＝★客には 分からない★ */
  const 画面の = [...new Set((画面の生.match(/data-ver="([a-z0-9_]+)"/g) || [])
    .map((s) => s.replace(/.*"([a-z0-9_]+)".*/, '$1')))].sort();
  if (画面の.length < 3) throw new Error('★版の 選び場が 読めない★（検査が 空振り）');
  const 無い = 画面の.filter((k) => サーバ.indexOf("'" + k + "'") < 0);
  if (無い.length) {
    throw new Error('★画面には 在るのに サーバが 知らない 版★: ' + 無い.join(' / ')
      + '\n   → ★選んでも 黙って Excel 365 の 答えに なります★');
  }
});

T('★この 機械が「半分」だと 書いてある（守ったと 読ませない）', () => {
  const 本文 = fs.readFileSync(fileURLToPath(import.meta.url), 'utf8');
  if (本文.indexOf('この 機械は 半分') < 0) throw new Error('★半分だと 書いていない★');
});

/* ★★この 機械は 半分しか 守りません★★
   ★見るのは「送る 字が 在るか」だけ★
   ★捕まえられない 物★
     ・localStorage が 空の 時に 何を 送るか（既定は 365）
     ・お客さんが 版を 選び直した後 すぐ 効くか
     ・サーバが その 版で ★本当に 違う 前置きを 組むか★
       （そこは tests/prompt-file.test.mjs が 5つの group を 組んで 見ている）
   ⇒★実ブラウザで 押して 中身を 見るまでは「届いた」と 言い切らない★ */

console.log('\n── 実測 ──');
console.log('  サーバが 読む 名 … ' + (r.サーバが読む名 || '★無い★'));
console.log('  画面が 送る 名 …… ' + (r.画面が送る名 || '★送っていない★'));
console.log('  届くか …………… ' + (r.届く ? '★届く★' : '★届かない（' + r.なぜ + '）★'));
console.log('\n' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
