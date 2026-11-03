/* soto-api.test.mjs — ★外へ 出る 口が 守るべき 物を 1つずつ 押す★（2026-09-07）
 *
 *  ★★なぜ 要るか★★
 *    2026-09-07 の 実測（docs/measured/soto/）で こう 出ました。
 *      ・ブラウザから 撃つと ★4回／4回とも 相手に 届く★（CORS は 読む 方だけ 止める）
 *      ・もらった `.xlsx` の 式は ★式のまま★ 入る＝★開いた だけで 走る★
 *    ⇒★★『どこへでも 出られる 口』を 作ったら お客さんの 中身が 出て行く★★
 *    ⇒ だから 口は ★1つ★（api/soto.js）／★行き先は うちが 決める★
 *
 *  ★★ここで 押す 物★★（1つでも 通ったら 赤）
 *    ①許していない 相手 ②http ③中の 網 ④許した 相手の 別の 道
 *    ⑤形に なっていない 住所 ⑥許した 相手＋許した 道 は 通る
 *    ⑦★許す 一覧に 訳と 決めた 日が 書いてある★
 *
 *  ★外へは 1回も 出ません★（住所を 見る 所だけを 押す）
 *
 *  使い方: node tests/soto-api.test.mjs
 *          node tests/soto-api.test.mjs --self-test
 */
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath, pathToFileURL } from 'node:url';

const ここ = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(ここ, '..');
const require_ = createRequire(import.meta.url);
const S = require_(path.join(ROOT, 'api/soto.js'));

const 直に走った = process.argv[1] && pathToFileURL(process.argv[1]).href === import.meta.url;
const 自己試験 = 直に走った && process.argv.includes('--self-test');

let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ✓ ' + n); } catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + (e && e.message)); } };

console.log('\n[soto-api] 外へ 出る 口が 守る 物');
console.log('  許した 相手 … ' + S.許す相手.length + '件');

/* ★通る 住所を 一覧から 作る★＝手で 書かない（一覧が 変わっても 追える） */
const 通る = S.許す相手.map((a) => {
  const 道 = a.道 ? String(a.道).replace(/^\/\^|\\?\/\?\$\/$/g, '').replace(/\\/g, '') : '/';
  return 'https://' + a.host + (道.startsWith('/') ? 道 : '/' + 道);
});

T('★許した 相手＋許した 道は 通る★（門が 閉まりきっていない事）', () => {
  if (!通る.length) throw new Error('許した 相手が 0件＝門が 全部 閉まっている');
  const r = S.住所を見る('https://stooq.com/q/d/l?s=msft.us&i=d');
  if (!r.よい) throw new Error('通るはずの 住所が 通らない … ' + r.訳);
});

T('★★許していない 相手には 出ない★★', () => {
  for (const u of ['https://example.com/', 'https://evil.test/x?d=1', 'https://stooq.com.evil.test/q/d/l']) {
    const r = S.住所を見る(u);
    if (r.よい) throw new Error('通ってしまった … ' + u);
    if (!/許していない/.test(r.訳)) throw new Error('訳が 出ていない … ' + u + ' → ' + r.訳);
  }
});

T('★http は 断る★（途中で 覗かれる／書き換えられる）', () => {
  const r = S.住所を見る('http://stooq.com/q/d/l?s=x');
  if (r.よい) throw new Error('http が 通った');
  if (!/https/.test(r.訳)) throw new Error('訳が 出ていない … ' + r.訳);
});

T('★★中の 網の 住所には 出ない★★（社内が 覗ける）', () => {
  for (const h of ['127.0.0.1', 'localhost', '10.0.0.5', '192.168.1.1', '172.16.0.1', '169.254.1.1', '[::1]']) {
    if (!S.中の網か(h)) throw new Error('中の 網と 見ていない … ' + h);
  }
  for (const h of ['stooq.com', '8.8.8.8', '172.32.0.1']) {
    if (S.中の網か(h)) throw new Error('外の 住所を 中の 網と 見ている … ' + h);
  }
});

T('★許した 相手でも 別の 道には 出ない★', () => {
  const r = S.住所を見る('https://stooq.com/anything/else');
  if (r.よい) throw new Error('別の 道が 通った');
  if (!/道/.test(r.訳)) throw new Error('訳が 出ていない … ' + r.訳);
});

T('★住所の 形に なっていない 物は 断る★', () => {
  for (const u of ['', 'あ', 'javascript:alert(1)', 'file:///etc/passwd', 'data:text/plain,x']) {
    if (S.住所を見る(u).よい) throw new Error('通ってしまった … ' + u);
  }
});

T('★★許す 一覧には 訳と 決めた 日が 書いてある★★（黙って 増やせない）', () => {
  for (const a of S.許す相手) {
    if (!a.用 || String(a.用).length < 8) throw new Error(a.host + ' に「何に 使うか」が 無い');
    if (!a.決め || !/\d{4}-\d{2}-\d{2}/.test(a.決め)) throw new Error(a.host + ' に「決めた 日」が 無い');
  }
});

if (自己試験) {
  console.log('\n[self-test] ★物差しが 効いているか★');
  T('★わざと 許していない 相手を 通そうとしたら 断られる★', () => {
    if (S.住所を見る('https://' + 'a'.repeat(20) + '.test/').よい) throw new Error('知らない 相手が 通った');
  });
  T('★中の 網の 見分けが 空振りしていない★', () => {
    if (!S.中の網か('127.0.0.1')) throw new Error('一番 分かりやすい 物を 見落とした');
  });
}

console.log('\n' + pass + ' passed, ' + fail + ' failed');
if (直に走った) process.exit(fail ? 1 : 0);
