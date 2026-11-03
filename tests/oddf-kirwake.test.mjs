/* oddf-kirwake.test.mjs — ★ODDFPRICE の 切り分けで 分かった 事を 守る★（2026-09-16）
 *
 *  ★★紙★★ `docs/measured/golden-oddf-to-46ko-2026-09-16.tsv`
 *    ★実Excel が 答えた 物★（★聞く 前に 見込みを 書いて あります★
 *      … `docs/measured/kansuu46/oddf-kiku-koto.md`）
 *
 *  ★★この 見張りが 守る 物（★合った 分だけ★）★★
 *    ①★紙が 読めて いる★（空振りして いない）
 *    ②★★発行日 ≧ 決済日 は #NUM!★★（★測って 決めました★・5本）
 *    ③★対照が 紙と 同じ★（★測り台が 変わって いない★）
 *    ④★今 合って いる 分が 減って いない★（★名指し★）
 *
 *  ★★まだ 合わない 分は ここでは 緑に しません★★
 *    ＝★`docs/measured/kansuu46/oddf-kiku-koto.md` の ⑥に 1本ずつ 名指しで 書いて あります★
 *    ＝★数で 許しません★
 *
 *  使い方: node tests/oddf-kirwake.test.mjs
 */
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(path.join(ROOT, 'package.json'));
const K = require_(path.join(ROOT, 'lib/formula-kane.js'));

let pass = 0, fail = 0;
const T = (n, f) => {
  try { f(); pass++; console.log('  ✓ ' + n); }
  catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + e.message); }
};

console.log('\n[oddf-kirwake] ★ODDFPRICE の 切り分けで 分かった 事を 守る★');

const 紙道 = path.join(ROOT, 'docs/measured/golden-oddf-to-46ko-2026-09-16.tsv');
const 行 = fs.readFileSync(紙道, 'utf-8').replace(/^﻿/, '').split(/\r?\n/)
  .filter((l) => l && !l.startsWith('#')).map((l) => l.split('\t'))
  .filter((c) => c.length >= 7)
  .map((c) => ({ 種: c[0], 訳: c[1], 式: c[2], 答: c[3], 字: c[4], ゼロ: c[5], 型: c[6] }));

const ODDF = 行.filter((r) => r.種 === 'ODDF');
const 四十六 = 行.filter((r) => r.種 === '46個');
console.log('      … 紙 ' + 行.length + '行（ODDF ' + ODDF.length + '行 ／ 46個 ' + 四十六.length + '行）');

const 数 = (y, m, d) => K.日から数(y, m, d);

T('★紙が 読めて いる（★空振りして いない★）★', () => {
  if (行.length < 80) throw new Error('★' + 行.length + '行しか 読めない★');
  if (!ODDF.length || !四十六.length) throw new Error('★片方の 種が 0行★');
  /* ★必ず 在る はずの 訳を 名指しで★＝拾い方が 痩せたら ここで 止まる */
  for (const n of ['㋐決済が準利払日の上(basis=2)', '㋒発行日=決済日(basis=0)']) {
    if (!ODDF.some((r) => r.訳 === n)) throw new Error('★紙に「' + n + '」が 無い★');
  }
});

T('★★発行日 ≧ 決済日 は #NUM!（★測って 決めました★）★★', () => {
  /* ★紙★ ㋒ の 5本＝★basis 0〜4 の 5つとも #NUM!★
       ★聞く 前は 「より 大きい」に して 置きました★（測った 分だけ 門に する）
       ⇒★聞いた 結果 「以上」に 広げました★ */
  const 実 = ODDF.filter((r) => /^㋒/.test(r.訳));
  if (実.length !== 5) throw new Error('★㋒ が ' + 実.length + '本★（5本の はず）');
  for (const r of 実) {
    if (r.字 !== '#NUM!') throw new Error('★紙の ' + r.訳 + ' が #NUM! で ない★（' + r.字 + '）');
  }
  /* ★うちも 同じ 形で #NUM! を 返すか★（★同じ 日★と ★発行が 後★ の 両方） */
  for (const [発, 名] of [[数(2009, 1, 1), 'ちょうど 同じ 日'], [数(2009, 4, 1), '発行が 決済より 後']]) {
    for (let b = 0; b <= 4; b++) {
      const r = K.初回端数の価格(数(2009, 1, 1), 数(2013, 1, 1), 発, 数(2010, 1, 1), 0.06, 0.05, 100, 2, b);
      if (!(r && r.誤り === 'NUM')) {
        throw new Error('★' + 名 + '（basis ' + b + '）が #NUM! に ならない★ … ' + JSON.stringify(r));
      }
    }
  }
});

T('★★対照が 紙と 同じ（★測り台が 変わって いない★）★★', () => {
  /* ★これが 合わなければ 他の 答えも 信じられません★ */
  const 対 = ODDF.filter((r) => /^㋓/.test(r.訳));
  if (対.length !== 2) throw new Error('★対照が ' + 対.length + '本★（2本の はず）');
  const 待 = ['113.59879960832529', '103.37232293583249'];
  for (let i = 0; i < 2; i++) {
    const a = Number(対[i].答), b = Number(待[i]);
    if (!(Math.abs(a - b) <= Math.abs(b) * 1e-12)) {
      throw new Error('★対照 ' + (i + 1) + ' が ' + 対[i].答 + '★（' + 待[i] + ' の はず）'
        + '／★測り台が 変わって います＝他の 答えも 信じられません★');
    }
  }
});

T('★今 合って いる 分が 減って いない（★名指し★）★', () => {
  /* ★合う 分だけ 門に します★
       ★㋐′ basis 3 は うちと ★ぴたり 一致★ しました★（NC>1 でも 合う 形が 在る 証し）
       ★㋐ basis 0/1/4 と ㋐′ basis 0/1/4 も 合って います★ */
  const 見る = [
    ['㋐決済が準利払日の上(basis=0)', 数(2009, 7, 1), 数(2009, 1, 1)],
    ['㋐決済が準利払日の上(basis=1)', 数(2009, 7, 1), 数(2009, 1, 1)],
    ['㋐決済が準利払日の上(basis=4)', 数(2009, 7, 1), 数(2009, 1, 1)],
    ['㋐′端数3期・決済が準利払日の上(basis=0)', 数(2009, 7, 1), 数(2008, 7, 1)],
    ['㋐′端数3期・決済が準利払日の上(basis=1)', 数(2009, 7, 1), 数(2008, 7, 1)],
    ['㋐′端数3期・決済が準利払日の上(basis=3)', 数(2009, 7, 1), 数(2008, 7, 1)],
    ['㋐′端数3期・決済が準利払日の上(basis=4)', 数(2009, 7, 1), 数(2008, 7, 1)],
  ];
  const 違 = [];
  for (const [訳, 決, 発] of 見る) {
    const r = ODDF.find((x) => x.訳 === 訳);
    if (!r) throw new Error('★紙に「' + 訳 + '」が 無い★');
    const b = Number(/basis=(\d)/.exec(訳)[1]);
    const u = K.初回端数の価格(決, 数(2013, 1, 1), 発, 数(2010, 1, 1), 0.06, 0.05, 100, 2, b);
    const e = Number(r.答);
    if (!(isFinite(u) && Math.abs(u - e) <= Math.abs(e) * 1e-12)) 違.push(訳 + ' 実' + e + ' うち' + u);
  }
  if (違.length) throw new Error('★' + 違.length + '/' + 見る.length + '本 減って いる★ … ' + 違.join(' ／ '));
  console.log('      … ' + 見る.length + '/' + 見る.length + '本（★㋐′ basis 3 は NC>1 でも ぴたり 一致★）');
});

T('★46個の 紙が 空振りして いない★', () => {
  /* ★ここでは 「合う／合わない」を 判じません★＝★紙を 取るのが 目的★
       ⇒★値が 入って いるか だけ 見ます★ */
  if (四十六.length !== 46) throw new Error('★46個が ' + 四十六.length + '行★');
  /* ★★「空の 字」と「測れて いない」は 別物★★（2026-09-16）
       PHONETIC と PIVOTBY は ★空の 字を 返しました★（型は String）
       ＝★これも 測った 結果です★（★捨てては いけません★）
     ⇒★型の 欄で 見ます★（★型が 空なら 測れて いません★） */
  const 測れず = 四十六.filter((r) => !r.型 || r.型 === '(空)');
  if (測れず.length) {
    throw new Error('★' + 測れず.length + '行 測れて いない★ … ' + 測れず.map((r) => r.訳).join(' '));
  }
  const 空の字 = 四十六.filter((r) => r.答 === '' && r.型 === 'String');
  console.log('      … ★空の 字を 返した★ ' + 空の字.length + '個 … ' + 空の字.map((r) => r.訳).join(' '));
  /* ★打てなかった 物も 「測った 結果」です★＝数を 決め打ちに して 黙って 増えないように */
  const 打てず = 四十六.filter((r) => /打てません/.test(r.答));
  const 期待 = 2;   /* CALL ／ REGISTER.ID（★実Excel が 式そのものを 受け付けません★） */
  if (打てず.length !== 期待) {
    throw new Error('★打てない 物が ' + 打てず.length + '個★（' + 期待 + '個の はず） … '
      + 打てず.map((r) => r.訳).join(' '));
  }
  console.log('      … 46/46 行（★打てない 2個＝CALL / REGISTER.ID★）');
});

console.log('\n' + pass + ' passed, ' + fail + ' failed');
if (fail) process.exit(1);
