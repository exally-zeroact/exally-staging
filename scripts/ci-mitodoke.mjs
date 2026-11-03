/* ci-mitodoke.mjs -- ★1つの sha の CI と WebKit を「数まで」見届ける★（2026-10-10）
 *  使い方: node scripts/ci-mitodoke.mjs <sha>
 *  終わり値: 0 ... 緑（判じと 言えない 事は scripts/lib/ci-mitodoke.mjs の 頭）
 *            1 ... 赤／見届けられない（★既定が 1★＝途中で 転んでも 緑に ならない）
 *  使う 順: ci-ga-hashitta-ka.mjs で 積み残し 0 を 見る ⇒ 枝先や main の sha を ここへ
 *  ★CI と WebKit は pull_request でしか 走らない（push で 走るのは main だけ）★
 */
import { 走らせる } from './lib/ci-mitodoke.mjs';

走らせる(process.argv.slice(2));
