/* webkit-hashiru.test.mjs — ★どの webkit の 見張りも CI の どこかで 本当に 走るか★（2026-10-04 夜）
 *
 *  ★★なぜ 要るか（経営者が 3bceea9 の 木で 数えて 見つけた）★★
 *    webkit の 見張り 19本の うち ★15本は CI の どこでも 1度も 本当に 走って いなかった★。
 *    毎回の CI（ci.yml）には webkit が 無いので ★未測定・緑★ で 終わる。本人は「webkit.yml で 本当に 測る」と 言うが、
 *    ★webkit.yml には 4本しか 書いて いなかった★ ⇒ 緑の 意味が 「手元で 走った 時だけ」に なって いた。
 *  ★★見る 物（毎回の CI で）★★
 *    ① webkit.yml に ★名簿から 拾う 段★（node scripts/run-webkit-tests.mjs）が 在る
 *    ② ★どの tests/*-webkit.mjs も★ 名簿（tests/run.js＝①の 段が 拾う）か webkit.yml の 段の どちらかで 走る
 *    ③ webkit.yml の pull_request の paths に 'tests/*-webkit.mjs' が 在る（見張りを 直した PR で 走る）
 *  ★--self-test★ ... webkit.yml の 字から ① の 段を 消すと 赤に なるか
 *  走らせ方: node tests/webkit-hashiru.test.mjs [--self-test]
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createRequire } from 'node:module';
import { webkitの名簿, 借りる試験か } from '../scripts/run-webkit-tests.mjs';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(pathToFileURL(path.join(ROOT, 'package.json')));
const { FILES } = require_(path.join(ROOT, 'tests/run.js'));

/** ★判じる★（字を 渡して 判じる＝自己試験で 字を 壊して 渡せる）
 *  ★見る 試験は 名前では なく 中身で★＝`_borrow-playwright` を 読み込む 試験（経営者の 訂正 10-04 夜：
 *    名前で 拾うと jsdom の -ui を 巻き込み、名前を 付け忘れた 借りる 試験を 落とす） */
function 判じる(yml, 名簿, ファイル) {
  const 赤 = [];
  const 段 = /^\s*run:\s*node\s+scripts\/run-webkit-tests\.mjs\s*$/m.test(yml);
  if (!段) 赤.push('① webkit.yml に 名簿から 拾う 段（node scripts/run-webkit-tests.mjs）が 無い');
  const 名簿の名 = new Set(名簿.map((f) => f[0]));
  const ymlの名 = new Set((yml.match(/run:\s*node\s+tests\/([A-Za-z0-9_.\/-]+\.mjs)/g) || [])
    .map((s) => /tests\/([A-Za-z0-9_.\/-]+\.mjs)/.exec(s)[1]));
  for (const f of ファイル) {
    const 走る = (段 && 名簿の名.has(f)) || ymlの名.has(f);
    if (!走る) 赤.push('② ' + f + ' は ブラウザを 借りるのに CI の どこでも 本当に 走らない（名簿にも webkit.yml の 段にも 無い）');
  }
  const pr = /pull_request:[\s\S]*?paths:([\s\S]*?)\n\s*(push|workflow_dispatch|schedule):/.exec(yml);
  if (!pr || pr[1].indexOf("'tests/*-webkit.mjs'") < 0) 赤.push("③ pull_request の paths に 'tests/*-webkit.mjs' が 無い");
  return 赤;
}

const yml = fs.readFileSync(path.join(ROOT, '.github/workflows/webkit.yml'), 'utf8');
const 名簿 = webkitの名簿(FILES);
const ファイル = fs.readdirSync(path.join(ROOT, 'tests')).filter((f) => /\.mjs$/.test(f))
  .filter((f) => 借りる試験か(fs.readFileSync(path.join(ROOT, 'tests', f), 'utf8'))).sort();
console.log('[webkit-hashiru] ★ブラウザを 借りる 見張りが 全部 CI の どこかで 本当に 走るか★');
console.log('  借りる 試験（tests/*.mjs の 中身で）... ' + ファイル.length + '本 ／ 名簿の 借りる 試験 ... ' + 名簿.length + '本（引数 付き 含む）');
if (ファイル.length === 0) { console.log('★見張りが 1本も 無い＝空振り★'); process.exit(8); }

let 緑 = 0, 赤数 = 0;
const T = (n, ok, m) => { if (ok) { 緑++; console.log('  ok   ' + n); } else { 赤数++; console.log('  NG   ' + n + (m ? '\n       ' + m : '')); } };

if (process.argv.includes('--self-test')) {
  const 壊した = yml.replace(/^\s*run:\s*node\s+scripts\/run-webkit-tests\.mjs\s*$/m, '        run: echo 消した');
  const r = 判じる(壊した, 名簿, ファイル);
  T('★webkit.yml から 名簿の 段を 消すと 赤（①と 名簿の 見張りが ②で 赤）★', r.length > 1, JSON.stringify(r).slice(0, 300));
  /* ★名前に -webkit が 無い 借りる 試験が どこにも 無い 時も 赤★（中身で 見て いる 証し） */
  const r2 = 判じる(yml, 名簿, ファイル.concat(['namae-no-nai-kariru.test.mjs']));
  T('★名前に -webkit が 無くても 借りる 試験が どこでも 走らなければ 赤★', r2.some((s) => s.indexOf('namae-no-nai-kariru') >= 0), JSON.stringify(r2).slice(0, 300));
  console.log('\nwebkit-hashiru --self-test: ' + 緑 + ' 緑 / ' + 赤数 + ' 赤');
  process.exit(赤数 ? 1 : 0);
}

const r = 判じる(yml, 名簿, ファイル);
T('★どの webkit の 見張りも CI の どこかで 本当に 走る（' + ファイル.length + '本）★', r.length === 0, r.join('\n       '));
console.log('\nwebkit-hashiru: ' + 緑 + ' 緑 / ' + 赤数 + ' 赤');
process.exit(赤数 ? 1 : 0);
