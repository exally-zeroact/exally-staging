/* marubatsu-tsukenaosu.mjs — ★紙の「合ったか」の 欄を 付け直す★（2026-09-18）
 *
 *  ★★なぜ★★
 *    9枠目の 道具は ★見込みの 字を そのまま★ 実測と 見比べて いました。
 *    ⇒★見込みの 欄に ★ や 逆引用符（飾り）を 書いて いた 行が ★全部 ×★★
 *      49 BYROW ... 見込み「★3★」／実 3 ... ×
 *      18 FIXED ... 見込み「`-1,234.57`」／実 -1,234.57 ... ×
 *    ⇒★合った 40／外れた 20★ と 出て いた（★本当は 合った 52／外れた 9★）
 *
 *  ★★これで 今日 3回目です★★
 *    ①門の 正規表現が 3点リーダを 掴んで いた
 *    ②9枠目の 道具に 3点リーダが 7個
 *    ③★この ○×欄★
 *    ⇒★★飾りの 字に 頼った 門は 飾りを 直した 日に 黙って 割れる★★
 *
 *  ★★紙は 取り直しません★★（★答えは 正しい★）＝★○×欄だけ 付け直します★
 *
 *  ★飾りの 落とし方★（★何を 落としたかを 出します★）
 *    ・黒星（U+2605）／逆引用符／前後の 空白
 *    ★落とさない 物★ ... 中の 記号（`#DIV/0!` の `#` や `,` や `¥`）
 *
 *  ★★これは 目安です★★
 *    ・★字で 見比べるだけ★＝`1` と `1.0` は 別に なります
 *    ・★誤りは 「出る字」と 見比べます★（答えの 欄は -2146826281 の ような 番号）
 *    ・★「分かりません」と 書いた 見込みは ★×でも ○でも ありません★★＝`?` に します
 *
 *  使い方: node docs/measured/marubatsu-tsukenaosu.mjs <紙>
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '../..');
const 道 = process.argv[2]
  || path.join(ROOT, 'docs/measured/golden-kansuu-9kaime-2026-09-18.tsv');
if (!fs.existsSync(道)) { console.log('★紙が 在りません★ ' + 道); process.exit(2); }

const 黒星 = String.fromCharCode(0x2605);
const 飾りを落とす = (s) => {
  let t = String(s == null ? '' : s);
  t = t.split(黒星).join('');
  t = t.split(String.fromCharCode(96)).join('');   /* 逆引用符 */
  return t.trim();
};

const 行 = fs.readFileSync(道, 'utf-8').split(/\r?\n/);
const 柱 = (行.find((l) => l.startsWith('# 番')) || '').replace(/^#\s*/, '').split('\t');
const 見列 = 柱.indexOf('見込み'), 答列 = 柱.indexOf('答え');
const 字列 = 柱.indexOf('出る字'), 合列 = 柱.indexOf('合ったか');
if (見列 < 0 || 答列 < 0 || 字列 < 0 || 合列 < 0) {
  console.log('★柱が 読めません★ ' + 柱.join(' / '));
  process.exit(3);
}

let 合 = 0, 違 = 0, 不明 = 0, 直した = 0;
const 外れ = [];
const 出 = 行.map((l) => {
  if (!l || l.startsWith('#')) return l;
  const c = l.split('\t');
  if (c.length <= 合列) return l;
  const 見 = 飾りを落とす(c[見列]);
  const 答 = 飾りを落とす(c[答列]);
  const 字 = 飾りを落とす(c[字列]);
  let 新;
  if (!見.length || 見.indexOf('分かりません') >= 0) { 新 = '?'; 不明++; }
  else if (見 === 答 || 見 === 字) { 新 = '○'; 合++; }
  else { 新 = '×'; 違++; 外れ.push(c[0] + '  ' + c[2] + '  見込み「' + 見 + '」／実「' + (字 || 答) + '」'); }
  if (c[合列] !== 新) 直した++;
  c[合列] = 新;
  return c.join('\t');
});

fs.writeFileSync(道, 出.join('\n'), 'utf-8');
console.log('');
console.log('★「合ったか」の 欄を 付け直しました★ ' + path.relative(ROOT, 道));
console.log('  ★★合った ' + 合 + ' ／ 外れた ' + 違 + ' ／ 決められない ' + 不明 + '★★');
console.log('  ★付け替えた 欄 ... ' + 直した + '個★（★飾りで ×に なって いた 分★）');
console.log('');
console.log('★★本当に 外れた 分★★');
for (const s of 外れ) console.log('  ' + s);
console.log('');
console.log('★言えない 事★');
console.log('  ・★字で 見比べただけ★＝`1` と `1.0` は 別に なります');
console.log('  ・★「打てません」と 出た 行は 外れに 数えます★（★答えでは ない★）');
