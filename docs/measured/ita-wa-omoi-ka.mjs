/* ita-wa-omoi-ka.mjs -- ★台の 板は 実物の 大きさで 使えるか★（2026-09-18）
 *
 *  ★★なぜ 今 測るか★★
 *    Exally1 が ㋑（計算を 台に 回す）の 形を 出しました。
 *    ★台は ★板を 1つ 持たないと マスを 読めません★★
 *    ⇒手は 2つ ... ㋐板を 持つ（速い／忘れると 黙って 古い 値）
 *                 ㋑聞く たびに 作り直す（忘れない／★重い★）
 *    ⇒★★「重い」は ★誰も 測って いません★★★
 *    ⇒★司さんの 実物は ★15,799式★★＝★そこで 使えないなら 形を 変えます★
 *    ⇒★★書く 前に 測ります★★
 *
 *  ★★この 台の 数は 画面の 数では ありません★★
 *    ＝node で 板だけを 立てて 秒を 測ります（借り物も 画面も 通しません）
 *    ⇒★画面では もっと 掛かります★（描き直しが 在る）
 *    ⇒★でも ★下の 限り★は これで 出ます★
 *
 *  ★見て いない 事★
 *    ・★式の 中身は 1種類だけ★です（`=SUM(...)` と `=A1*2`）
 *      ＝★実物は もっと 重い 式が 在ります★
 *    ・★板の 作り直し（㋑）は 1回だけ★＝★打つ たびに 作り直す 形は 別★
 *    ・★覚え（メモリ）は 測って いません★
 *
 *  使い方: node docs/measured/ita-wa-omoi-ka.mjs
 */
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';

const ここ = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(ここ, '..', '..');
const require_ = createRequire(path.join(ROOT, 'package.json'));
const Tsu = require_(path.join(ROOT, 'lib/shiki-tsunagi.js'));
require_(path.join(ROOT, 'lib/shiki-kansuu.js'));
try {
  const { JSDOM } = require_('jsdom');
  const 窓 = new JSDOM('<!doctype html>').window;
  Tsu.XML道具を渡す({ DOMParser: 窓.DOMParser, XPathResult: 窓.XPathResult });
} catch (e) { /* ★FILTERXML は 使いません★ */ }
const H = require_(path.join(ROOT, 'lib/shiki-hyou.js'));

const 列名 = (i) => {
  let s = '';
  let n = i;
  do { s = String.fromCharCode(65 + (n % 26)) + s; n = Math.floor(n / 26) - 1; } while (n >= 0);
  return s;
};

/* ★司さんの 実物は 15,799式★＝★その 上下で 測ります★ */
const 大きさたち = [500, 2000, 8000, 16000];

console.log('');
console.log('[ita-wa-omoi-ka] ★台の 板は 実物の 大きさで 使えるか★');
console.log('  ★司さんの 実物 ... 15,799式★（★この 上下で 測ります★）');
console.log('');
console.log('  式の数   板を建てる   1マス打ち直す   全部読む     1本あたり');
for (const n of 大きさたち) {
  const t0 = Date.now();
  const 板 = new H.表();
  /* ★材料★ A列に 数 ／ B列に 式 */
  for (let i = 0; i < n; i++) 板.打つ('A' + (i + 1), (i % 97) + 1);
  for (let i = 0; i < n; i++) 板.打つ('B' + (i + 1), '=A' + (i + 1) + '*2');
  const 建てる秒 = Date.now() - t0;

  const t1 = Date.now();
  板.打つ('A1', 999);            /* ★1マス 打ち直す★＝★お客さんが 1つ 直した 時★ */
  const v = 板.字('B1');
  const 打ち直す秒 = Date.now() - t1;

  const t2 = Date.now();
  let 読んだ = 0;
  for (let i = 0; i < n; i++) { if (板.字('B' + (i + 1)) !== null) 読んだ += 1; }
  const 読む秒 = Date.now() - t2;

  console.log('  ' + String(n * 2).padStart(6)
    + String(建てる秒 + ' ms').padStart(12)
    + String(打ち直す秒 + ' ms').padStart(15)
    + String(読む秒 + ' ms').padStart(12)
    + String(((建てる秒 / (n * 2)) * 1000).toFixed(1) + ' us').padStart(13)
    + (v === '1998' ? '' : '   ★B1 が ' + v + '（1998 の はず）★'));
}

console.log('');
console.log('  ★★読み方★★');
console.log('    ・★板を 建てる★ ... ★読み込みの 時★（1回だけ）');
console.log('    ・★1マス 打ち直す★ ... ★お客さんが 1つ 直す たび★（★ここが 一番 効きます★）');
console.log('    ・★全部 読む★ ... ★描き直す 時★');
console.log('    ・★1本あたり★ が 大きさで 増えるなら ★2乗★＝★実物で 使えません★');
