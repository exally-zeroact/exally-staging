/* run-webkit-tests.mjs — ★webkit の 見張りを 名簿（tests/run.js）から 全部 拾って 走らせる★（2026-10-04 夜）
 *
 *  ★★なぜ 在るか（経営者が 3bceea9 の 木で 数えて 見つけた）★★
 *    webkit の 見張り 19本の うち ★15本は CI の どこでも 1度も 本当に 走って いなかった★。
 *      ・毎回の CI（ci.yml）には webkit が 無い ⇒ 借り方（_borrow-playwright）が ★未測定・緑★ で 終わる
 *      ・見張り 本人は「webkit.yml で 本当に 測ります」と 言う ⇒ ★webkit.yml には 4本しか 書いて いなかった★
 *    ⇒ 10-02〜10-04 に 足した #REF!・和暦経過・描かれた字 の 見張りも ★手元の 総なめ でしか 走って いなかった★
 *  ★★直し★★
 *    webkit.yml は ★1本ずつ 並べず★ この 道具を 1段で 呼ぶ ＝★名簿に 足せば 勝手に 走る★
 *    ★門★ tests/webkit-hashiru.test.mjs（毎回の CI）＝webkit.yml が この 段を 持つか・拾い 漏れが 無いか
 *  ★走らせ方★ node scripts/run-webkit-tests.mjs           ... 全部 走らせる（webkit.yml の 段）
 *             node scripts/run-webkit-tests.mjs --名簿だけ ... 拾った 名前を 出す だけ（走らせない）
 */
import path from 'node:path';
import fs from 'node:fs';
import { spawnSync } from 'node:child_process';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createRequire } from 'node:module';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');

/** ★中で ブラウザを 借りる 試験か★（名前では なく 中身で＝経営者の 訂正 10-04 夜）
 *  名前で 拾うと ★jsdom の -ui を 巻き込み、名前を 付け忘れた 借りる 試験を 落とす★ */
export function 借りる試験か(本文) {
  /* ★import して いる 物だけ★（注記の 中の 字では 拾わない＝門の 試験 自身を 巻き込んだ・10-04 夜） */
  return /^\s*import\s[^;\n]*['"][^'"]*_borrow-playwright\.mjs['"]/m.test(String(本文))
    || /import\(\s*[^)]*_borrow-playwright\.mjs/.test(String(本文));
}
/** ★名簿から ブラウザを 借りる 見張りを 拾う★（引数 付きも そのまま） */
export function webkitの名簿(FILES, 読む) {
  const 読み = 読む || ((f) => fs.readFileSync(path.join(ROOT, 'tests', f), 'utf8'));
  return FILES.map((f) => (Array.isArray(f) ? f : [f]))
    .map((f) => f.filter((x) => x !== null && x !== undefined))
    .filter((f) => { try { return 借りる試験か(読み(String(f[0]))); } catch (e) { return false; } });
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const require_ = createRequire(pathToFileURL(path.join(ROOT, 'package.json')));
  const { FILES } = require_(path.join(ROOT, 'tests/run.js'));
  /* ★webkit.yml に 個別の 段が 在る 物は ここで 走らせない★（karimono 等＝2回 走らせない） */
  let yml = '';
  try { yml = fs.readFileSync(path.join(ROOT, '.github/workflows/webkit.yml'), 'utf8'); } catch (e) { yml = ''; }
  const 個別 = new Set((yml.match(/run:\s*node\s+tests\/([A-Za-z0-9_.\/-]+\.mjs)/g) || [])
    .map((s) => /tests\/([A-Za-z0-9_.\/-]+\.mjs)/.exec(s)[1]));
  const 名簿 = webkitの名簿(FILES).filter((f) => !個別.has(f[0]));
  if (process.argv.includes('--名簿だけ')) {
    for (const f of 名簿) console.log(f.join(' '));
    console.log('★webkit の 見張り ' + 名簿.length + '本★（tests/run.js 本人から 拾った）');
    process.exit(0);
  }
  let 走った = 0, 緑 = 0;
  const 落ちた = [];
  for (const [f, ...a] of 名簿) {
    console.log('\n=== ' + [f].concat(a).join(' ') + ' ===');
    const r = spawnSync(process.execPath, [path.join(ROOT, 'tests', f)].concat(a), { stdio: 'inherit' });
    走った++;
    if (r.status === 0) 緑++; else 落ちた.push([f].concat(a).join(' ') + '（' + (r.signal ? 'signal ' + r.signal : 'exit ' + r.status) + '）');
  }
  /* ★最後の 行★＝★頼んだ／走らせた／緑★（これが 無い 出しは 途中で 切られた＝赤） */
  console.log('\n★webkit の 見張り ... 頼んだ ' + 名簿.length + '本 ／ 走らせた ' + 走った + '本 ／ 緑 ' + 緑 + '本★');
  for (const s of 落ちた) console.log('   ★落ちた★ ' + s);
  process.exit((落ちた.length || 走った !== 名簿.length) ? 1 : 0);
}
