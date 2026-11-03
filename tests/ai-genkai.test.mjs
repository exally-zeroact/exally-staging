/* ai-genkai.test.mjs — ★掘る数が 叩ける数を 超えたら 赤★（2026-09-05）
 *
 *  ★なぜ 要るか（実測で 起きていた事）★
 *    画面 … lib/horu.js が 実物ブックで ★13★ を 返し、book.html が +2 して ★15回★ 叩く
 *    サーバ … api/claude.js 事故止め ★1分に 10回★ で 止める
 *    ⇒★11回目で 止まり、それまでの 掘りを 全部 捨てて「使いすぎ」だけ 出す★
 *    ⇒★大きいブックほど 必ず 失敗する★（売り文句の 真逆）
 *  ★人が 気づけない★＝どちらの ファイルを 読んでも 片方の 数しか 見えない。
 *  ⇒★機械で 突き合わせる★。
 *
 *  使い方: node tests/ai-genkai.test.mjs
 *          node tests/ai-genkai.test.mjs --self-test
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(import.meta.url);
let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ✓ ' + n); } catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + (e && e.message)); } };

/** ★純関数★＝数だけ 渡して 判じる（self-test で わざと 壊せる） */
export function 噛み合っているか(叩ける, 止める, 画面の余裕) {
  const 実際に叩く = 叩ける + (画面の余裕 || 0);
  return { ok: 実際に叩く <= 止める, 実際に叩く, 止める };
}

/* ══ self-test（わざと 壊して 赤に なるか） ══════════════════════ */
if (process.argv.includes('--self-test')) {
  console.log('\n[ai-genkai --self-test] わざと壊して赤になるか');
  T('★叩く数が 止める数を 超えたら 赤★（前の 実物＝13+2 vs 10）', () => {
    if (噛み合っているか(13, 10, 2).ok) throw new Error('赤になっていない');
  });
  T('★ちょうど 同じなら 通す★', () => {
    if (!噛み合っているか(8, 10, 2).ok) throw new Error('通っていない');
  });
  T('★1つでも 超えたら 赤★', () => {
    if (噛み合っているか(9, 10, 2).ok) throw new Error('赤になっていない');
  });
  console.log('\n' + pass + ' passed, ' + fail + ' failed');
  process.exit(fail ? 1 : 0);
}

/* ══ 本番（実物の 数を 読む） ══════════════════════════════════ */
const G = require_(path.join(ROOT, 'lib/ai-genkai.js'));
const H = require_(path.join(ROOT, 'lib/horu.js'));
const api = fs.readFileSync(path.join(ROOT, 'api/claude.js'), 'utf8');
const book = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf8');

console.log('\n[ai-genkai] 掘る数と 叩ける数が 噛み合っているか');

T('★サーバの 止める数は lib/ai-genkai.js から 読んでいる（べた書きでない）', () => {
  if (!/分の回数:\s*AI限界\.分の回数/.test(api)) {
    throw new Error('api/claude.js が 数を べた書きしています（1か所に 集めた 意味が 無い）');
  }
});

T('★画面の +2 が 実物と 合っている（ここを 変えたら この検査も 直す）', () => {
  const m = /for\(var 回 = 0; 回 < 掘れる \+ (\d+); 回\+\+\)/.exec(book);
  if (!m) throw new Error('book.html の 掘るループが 見つからない（形が 変わった）');
  if (Number(m[1]) !== 2) throw new Error('+' + m[1] + ' に 変わっています。この検査の 余裕も 直してください');
});

T('★★どんな 大きさの ブックでも 叩く数が 止める数を 超えない★★', () => {
  const 余裕 = 2;
  for (const セル数 of [0, 100, 2000, 21204, 100000, 1000000]) {
    const 偽 = [{ data: Object.fromEntries(Array.from({ length: セル数 }, (_, i) => ['A' + i, 1])) }];
    const 掘れる = H.掘れる回数を決める(偽);
    const r = 噛み合っているか(掘れる, G.限界.分の回数, 余裕);
    if (!r.ok) {
      throw new Error('セル' + セル数 + ' → 掘れる' + 掘れる + '＝叩く' + r.実際に叩く
        + ' > 止める' + r.止める + ' ★11回目で 止まって 掘りを 捨てます★');
    }
  }
});

T('★止められた時に それまでの 答えを 捨てない（book.html）', () => {
  if (!/止められた時に それまでの 掘りを 捨てない/.test(book)) {
    throw new Error('捨てない 直しが 消えています');
  }
  if (!/if\(aiText\)\{ aiText = aiText/.test(book)) {
    throw new Error('答えを 残して 断りを 添える 形に なっていません');
  }
});

T('★掘りきれない 時は 正直に 言える（言う道具が 在る）', () => {
  if (typeof G.掘りきれるか !== 'function') throw new Error('掘りきれるか が 無い');
  if (G.掘りきれるか(13)) throw new Error('13回 要るのに「掘りきれる」と 言っています');
  if (!G.掘りきれるか(2)) throw new Error('2回で 済むのに「掘りきれない」と 言っています');
});

console.log('\n── 実測 ──');
console.log('  1分に 叩ける … ' + G.限界.分の回数 + ' 回（サーバが 止める）');
console.log('  1問で 叩ける … ' + G.限界.一問で叩ける + ' 回');
for (const c of [2000, 21204, 1000000]) {
  const 偽 = [{ data: Object.fromEntries(Array.from({ length: c }, (_, i) => ['A' + i, 1])) }];
  console.log('  ' + String(c).padStart(7) + ' セル → 掘れる ' + H.掘れる回数を決める(偽)
    + ' 回（叩く ' + (H.掘れる回数を決める(偽) + 2) + ' 回）');
}
console.log('\n' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
