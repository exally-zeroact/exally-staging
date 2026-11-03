/* dkei-honban.test.mjs — ★D系 12個を 本番の 道で 実Excel の 紙と 突き合わせる★（2026-09-15）
 *
 *  ★★なぜ 要るか★★
 *    `exally-formula.js` の `_jsDbFunc` は
 *      `var critFieldIdx = headers.indexOf(critField);`   ←★見つからないと -1★
 *      `... col: s.c + critFieldIdx ...`                  ←★★列 -1（範囲の 外）を 読む★★
 *    ⇒★条件の 見出しが 台帳の 見出しと 紐付かない 時★ ★1件も 当たらない★
 *    ⇒★実測★ D系 12個の うち ★11個が 同じ 所で 外す★（★DGET だけ その形の 紙が 無い★）
 *
 *  ★★物差し★★ … ★実Excel★（`docs/measured/kansuu46/` の 紙・★21行★）
 *    ★紙から 読みます★＝★この 紙が 答えを 持って いる★（★道具が 決めない★）
 *    ★取った Excel★ … 版 16.0 ／ build 20326
 *
 *  ★★材料★★ … `docs/measured/kansuu46-no-dodai.mjs`（★紙を 取った 時と 同じ★）
 *
 *  ★★見て いない 事（★直して いない 所★）★★
 *    ★製品と 台（`lib/shiki-kansuu.js`）は ★あと 6か所 違います★（経営者1 が 1行ずつ 並べた）
 *      ②条件が 2列以上（かつ） ③条件が 2行以上（または） ④`>3` `<>x` の 形
 *      ⑤field が 列数より 大きい ⑥DPRODUCT で 0件 ⑦DGET で 0件
 *    ⇒★★6つとも 実Excel で 測って いません★★＝★当て推量で 直しません★（棚 ㉝）
 *    ⇒★この 紙が 見るのは ①（見出しが 紐付かない）だけ★
 *
 *  使い方: node tests/dkei-honban.test.mjs
 *          node tests/dkei-honban.test.mjs --self-test
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
/* ★Windows の 絶対道は `file://` に して 渡す★（そのままだと ESM が 受け取りません） */
const 読む = (p) => import(pathToFileURL(path.join(ROOT, p)).href);
const { 建てる } = await 読む('docs/measured/honban-no-michi.mjs');
const { 押す, 板を空に } = await 読む('docs/measured/kansuu46-no-dodai.mjs');

const D = /^(DSUM|DAVERAGE|DCOUNT|DCOUNTA|DGET|DMAX|DMIN|DPRODUCT|DSTDEV|DSTDEVP|DVAR|DVARP)$/;
const 紙 = path.join(ROOT, 'docs/measured/kansuu46');
const 組 = new Map();
for (const f of fs.readdirSync(紙).filter((x) => x.endsWith('.tsv'))) {
  for (const l of fs.readFileSync(path.join(紙, f), 'utf8').split(/\r?\n/)) {
    if (!l || l.startsWith('#')) continue;
    const c = l.split('\t');
    if (D.test(c[0]) && c[1] && c[1].startsWith('=')) 組.set(c[1], { 名: c[0], 実: c[2], 紙: f });
  }
}

console.log('\n[dkei-honban] ★D系を 本番の 道で 実Excel と 突き合わせる★');
console.log('  … ★紙から 読んだ 組★ ' + 組.size + '行（★これが 分母★）');

/* ★★分母が 減ったら 赤★★＝★押して いない 行は「合わない」事も 出来ない★ */
const 要る行数 = 21;
let pass = 0, fail = 0;
if (組.size !== 要る行数) {
  fail++;
  console.log('  ✗ ★★紙の 行数が ' + 組.size + '（' + 要る行数 + ' のはず）★★'
    + '＝★紙が 増えた／減った なら この 数も 直して ください★');
} else {
  pass++;
  console.log('  ✓ ★分母が ' + 要る行数 + '行 在る★');
}

const 台 = await 建てる();
const 外れ = [];
for (const [式, { 名, 実 }] of [...組.entries()].sort()) {
  const r = 押す(台, 式); 板を空に(台);
  if (String(r.字) === String(実)) { pass++; console.log('  ✓ ' + 名 + '  ' + 式 + ' → ' + r.字); }
  else {
    fail++; 外れ.push(名);
    console.log('  ✗ ' + 名 + '  ' + 式 + ' → ★' + r.字 + '★（実Excel ★' + 実 + '★）');
  }
}

console.log('\n' + pass + ' passed, ' + fail + ' failed'
  + (外れ.length ? ' ／ ★外した 関数★ ' + [...new Set(外れ)].join(' ') : ''));

/* ★★自己試験★★ … ★この 紙が 見て いる 穴は 1つだけ★ と ★分母★ を 字で 残す */
if (process.argv.includes('--self-test')) {
  /* ★★字で 見るのを やめました★★（2026-09-15）
       ★前は `critFieldIdx<0` が 在るかを 見て いました★
       ⇒★多列（かつ）を 直した 時 その 字が 消え、★振る舞いは 正しいのに 赤★に なりました★
       ＝★★見張りが 直しの 形に 縛りを 掛けて いた★★（記憶「見張りは いつも 新で 死ぬ」）
       ⇒★★直った かどうかは ★答え★で 見る★★ */
  let 出 = 0;
  const src = fs.readFileSync(path.join(ROOT, 'exally-formula.js'), 'utf8');
  if (/function\s+_jsDbFunc\s*\(/.test(src)) console.log('[self-test] ✓ ★穴の 場所★ `_jsDbFunc` が 在る');
  else { 出++; console.log('[self-test] ✗ ★`_jsDbFunc` が 見つかりません★'); }

  /* ★名指しの 1行★＝★この 行が 直しの 証し★（条件の 見出しが 台帳に 無い ⇒ 縛らない） */
  const 証し = '=DSUM(A1:B5,1,D1:D2)';
  const 正 = 組.get(証し);
  if (!正) { 出++; console.log('[self-test] ✗ ★証しの 行が 紙から 消えて います★ … ' + 証し); }
  else {
    const r = 押す(台, 証し); 板を空に(台);
    if (String(r.字) === String(正.実)) console.log('[self-test] ✓ ★証しの 行★ ' + 証し + ' → ' + r.字);
    else { 出++; console.log('[self-test] ✗ ★証しの 行が 直って いません★ … ' + r.字 + '（' + 正.実 + ' のはず）'); }
  }
  console.log('[self-test] ★棚㉝ の 6か所は ★別の 紙★で 見て います★ … tests/dkei6-honban.test.mjs（26行）');
  process.exit(出 ? 1 : 0);
}
process.exit(fail ? 1 : 0);
