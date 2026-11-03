/* awaseru-346.mjs — ★実Excel の 3,599本と Exally を 1本ずつ 突き合わせる★（2026-09-08）
 *
 *  ★★司さん 2026-09-08「答え確かめて」★★
 *    『名前が 通った』の うち ★答えを 1度も 確かめていなかった 346個★を 測る。
 *
 *  ★実Excel の 答え★ `golden-346-2026-09-08.tsv`
 *    ★引数の 形は 実Excel が 選んだ★（誤りに ならなかった 形を 全部）
 *    ⇒★人が 決めたのは 形の 候補だけ★
 *
 *  ★★押す 道は 本番と 同じ★★
 *    JS層 → convertFormula → HyperFormula（★プラグインも 本番と 同じ 全部★）
 *
 *  ★★数える 時の 決まり★★
 *    ・★合った★ … 実Excel と 同じ 答え
 *    ・★違う★ … ★これが 本命★（＝★それらしい 数を 出している★）
 *    ・★名前が 通らない★ … #NAME?（★台帳が 嘘★＝「動く」と 言っていたのに 動かない）
 *    ・★誤りで 揃った★ … 両方 同じ 誤り＝合っている
 *
 *  使い方: node docs/measured/kansuu46/awaseru-346.mjs
 */
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath, pathToFileURL } from 'node:url';

const ここ = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(ここ, '..', '..', '..');
const require_ = createRequire(path.join(ROOT, 'package.json'));

/* ★★本番の 道は 1本★★（2026-09-15）＝`docs/measured/honban-no-michi.mjs`
   ★前は この 道具も 自前で 建てて いました★（★写し 8本目★）
   ★しかも プラグインは ★7本★だけでした★
     （本番は 8本／★`complex`（IM系 21個を 包む）が 抜けて いた★）
   ⇒★寄せた 事で IM系も 本番と 同じ に なります★
   ★板の 名前は `Sheet1`★＝★実 Excel を 測った 時と 同じ★（口の 引数） */
const 道 = await import(pathToFileURL(path.join(ROOT, 'docs/measured/honban-no-michi.mjs')).href);
const 土台 = await import(pathToFileURL(path.join(ROOT, 'docs/measured/kansuu46-no-dodai.mjs')).href);
const 台 = await 道.建てる({ 板の名: 'Sheet1' });
const EF = 台.EF;
const hf = 台.hf;
const SID = 台.SID;
const 積1 = true;
const 積 = 台.積;

const 赤の名 = (t) => ({
  NA: '#N/A', DIV_BY_ZERO: '#DIV/0!', VALUE: '#VALUE!', NUM: '#NUM!',
  NAME: '#NAME?', REF: '#REF!', CYCLE: '#CYCLE!', ERROR: '#ERROR!', SPILL: '#SPILL!',
}[t] || ('#' + t));

/* ★実Excel を 測った 時と 同じ 材料★ */
/* ★★材料と 置き場と 押し方は 共通の 本★★（2026-09-15）
     `docs/measured/kansuu46-no-dodai.mjs`
   ★前は ここに 書いて ありました★（★実 Excel を 測った 時と 同じ 材料★）
   ⇒★別の 道具が 同じ 紙を 押す 時 ★材料を 自分で 決めて しまいました★
     （2026-09-15：`=IMSUM(A1:A5)` うち 24／実 Excel 15＝★材料が 違う★
       ⇒★928本が「合わない」と 出た★＝★偽の 負け★）
   ⇒★★材料を 1か所に しました★★ */
function 押す(式) {
  const r = 土台.押す(台, 式);
  if (r.道 === 'JS層') {
    /* ★この 道具は JS層の 戻りを そのまま 使って いました★ */
    const js = EF._jsComputeFormula(0, 式);
    return js;
  }
  if (r.字 === '★投げた★') return '★engine で 例外★';
  return r.字;
}


const 金 = fs.readFileSync(path.join(ここ, 'golden-346-2026-09-08.tsv'), 'utf-8')
  .split('\n').filter((l) => l && !l.startsWith('#'))
  .map((l) => { const p = l.split('\t'); return { 名: p[0], 式: p[1], 答: p[2], 型: p[3] }; });

const 同じか = (出, 正, 型) => {
  if (出 === null || 出 === undefined) return 正 === '';
  const s = String(出);
  if (型 === 'Double') {
    const a = Number(出), b = Number(正);
    if (!isFinite(a) || !isFinite(b)) return s === 正;
    if (a === b) return true;
    const 大 = Math.max(Math.abs(a), Math.abs(b));
    return Math.abs(a - b) <= (大 > 1 ? 大 * 1e-9 : 1e-9);
  }
  if (型 === 'Boolean') return s.toUpperCase() === String(正).toUpperCase();
  return s === 正;
};

/* ★★2026-09-08 に 足した＝★数が 測るたび ±1 ずれていた★★
   ⇒ 訳 … RANDBETWEEN(D1,D2) が ★たまたま★ 実Excel の 値と 一致する 回が 在る
   ⇒ ★実際に 起きた★ 続けて 2回 走らせて 合った=3223／3224 と 割れた
   ⇒★毎回 同じ 数が 出ない 紙は ★証拠に ならない★★
   ⇒ だから ★測るたび 変わる 関数は 先に 外へ 出す★（数に 入れない） */
const さいころ = /^(NOW|TODAY|RAND|RANDBETWEEN|RANDARRAY)$/;

const 分け = { 合った: [], 違う: [], 名前が通らない: [], こちらだけ誤り: [], 測るたび変わる: [] };
for (const g of 金) {
  if (さいころ.test(g.名)) { 分け.測るたび変わる.push(g); continue; }
  const 出 = 押す(g.式);
  const s = String(出);
  if (s === '#NAME?') { 分け.名前が通らない.push(g); continue; }
  if (同じか(出, g.答, g.型)) { 分け.合った.push(g); continue; }
  if (/^#|^★/.test(s)) { 分け.こちらだけ誤り.push({ ...g, 出: s }); continue; }
  分け.違う.push({ ...g, 出: s });
}

const 名で数える = (a) => new Set(a.map((x) => x.名)).size;
const 行 = [];
const 言う = (s) => { 行.push(s); console.log(s); };
言う('# ★実Excel の 答えと Exally を 1本ずつ 突き合わせた★（2026-09-08）');
言う('');
言う('★司さん 2026-09-08「答え確かめて」★');
言う('★押した 道★ JS層 → convertFormula → engine（★本番と 同じ★）');
言う('★積んだ プラグイン★ ' + 積.join(' / ') + '（exally-formula ' + (積1 ? '○' : '×') + '）');
言う('');
言う('★式 ' + 金.length + '本 ／ 関数 ' + 名で数える(金) + '個★');
言う('  ★数えない ' + 分け.測るたび変わる.length + '本★（' +
  [...new Set(分け.測るたび変わる.map((x) => x.名))].sort().join(' ') +
  '＝★測るたび 変わる★ので 合う/合わないが 決まらない）');
言う('  ⇒ 数える のは ' + (金.length - 分け.測るたび変わる.length) + '本');
言う('');
言う('  ★合った ……………………… ' + 分け.合った.length + '本（関数 ' + 名で数える(分け.合った) + '個）★');
言う('  ★★違う（それらしい 数を 出している）… ' + 分け.違う.length + '本（関数 ' + 名で数える(分け.違う) + '個）★★');
言う('  ★こちらだけ 誤り ………… ' + 分け.こちらだけ誤り.length + '本（関数 ' + 名で数える(分け.こちらだけ誤り) + '個）★');
言う('  ★名前が 通らない ………… ' + 分け.名前が通らない.length + '本（関数 ' + 名で数える(分け.名前が通らない) + '個）★');
言う('    ⇒★台帳が「動く」と 言っていたのに #NAME?＝★台帳が 嘘★★');
言う('');
言う('★★違う（一番 悪い＝黙って 間違った 数を 出している）★★');
const 違う名 = {};
for (const x of 分け.違う) { (違う名[x.名] = 違う名[x.名] || []).push(x); }
for (const 名 of Object.keys(違う名).sort()) {
  言う('  ★' + 名 + '★（' + 違う名[名].length + '本）');
  for (const x of 違う名[名].slice(0, 3)) 言う('    ' + x.式 + '   正 ' + x.答 + ' ／ 出 ' + x.出);
}
言う('');
言う('★名前が 通らない 関数★');
言う('  ' + [...new Set(分け.名前が通らない.map((x) => x.名))].sort().join(' '));
言う('');
言う('★こちらだけ 誤り（関数ごと 先頭 1本）★');
const 誤り名 = {};
for (const x of 分け.こちらだけ誤り) { if (!誤り名[x.名]) 誤り名[x.名] = x; }
for (const 名 of Object.keys(誤り名).sort()) {
  言う('  ' + 名 + '   ' + 誤り名[名].式 + '   正 ' + 誤り名[名].答 + ' ／ 出 ' + 誤り名[名].出);
}

fs.writeFileSync(path.join(ここ, 'awaseru-346.txt'), 行.join('\n') + '\n', { encoding: 'utf-8' });
console.log('\n★書いた … docs/measured/kansuu46/awaseru-346.txt★');
