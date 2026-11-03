/* oddf-susumi.test.mjs -- ★保留の 4個（ODDF / ODDL）の 進みを 数で 押さえる★（2026-09-18）
 *
 *  ★★なぜ 要るか★★
 *    `tests/formula-kane.test.mjs` は ★出す 22個★だけを 見て います。
 *    保留の 4個は「★出して いない事★」しか 見て いません。
 *    ⇒★★つまり 良く しても 悪く しても ★何も 赤に なりません★★★
 *    ⇒2026-09-18 に ODDFPRICE を 46 -> 47 に しました。
 *      ★その 数を 守る 門が 1つも 在りませんでした★
 *    ⇒★★紙に 書いた 数は 門で 守る★★
 *
 *  ★★見る 物★★
 *    `docs/measured/kansuu46/golden-kane-2026-09-07.tsv`（実Excel 700本）
 *      ODDFPRICE 48本 ／ ODDFYIELD 24本 ／ ODDLPRICE 12本 ／ ODDLYIELD 12本
 *    ⇒★合った 本数を 決め打ち★＝★減ったら 赤／増えても 赤★
 *      （★増えたら 上限を 上げて ください★＝★緩い 門は 門では ありません★）
 *
 *  ★★この 門が 見て いない 事★★
 *    ・★出すかどうかは 見て いません★（それは `formula-kane.test.mjs` の 仕事）
 *    ・★ODDFYIELD は もう 24 / 24 です★＝★残るのは ODDFPRICE の 1本と ODDL の 6本ずつ★
 *    ・★`#NUM!` の 行も 「合う」に 数えます★（★誤りの 字も 合わせます★）
 *
 *  使い方: node tests/oddf-susumi.test.mjs
 *          node tests/oddf-susumi.test.mjs --self-test
 */
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { 式をほどく, 答えを見る } from './monosashi.mjs';

const ここ = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(ここ, '..');
const require_ = createRequire(import.meta.url);
const K = require_(path.join(ROOT, 'lib/formula-kane.js'));
const 金の道 = path.join(ROOT, 'docs/measured/kansuu46/golden-kane-2026-09-07.tsv');

const 直に走った = process.argv[1] && pathToFileURL(process.argv[1]).href === import.meta.url;
const 自己試験 = 直に走った && process.argv.includes('--self-test');

let pass = 0, fail = 0;
const T = (n, fn) => {
  try { fn(); pass++; console.log('  ok   ' + n); }
  catch (e) { fail++; console.log('  NG   ' + n + '\n       ' + (e && e.message)); }
};

/* ★引数を かっこの 外の カンマで 割る★（`formula-kane.test.mjs` と 同じ 決まり） */
function 引数を割る(s) {
  const 出 = []; let 深 = 0, 今 = '';
  for (const ch of s) {
    if (ch === '(') { 深++; 今 += ch; continue; }
    if (ch === ')') { 深--; 今 += ch; continue; }
    if (ch === ',' && 深 === 0) { 出.push(今.trim()); 今 = ''; continue; }
    今 += ch;
  }
  if (今.trim() !== '') 出.push(今.trim());
  return 出;
}
function 値にする(s) {
  s = String(s).trim();
  const d = s.match(/^DATE\((\d+),(\d+),(\d+)\)$/i);
  if (d) return K.日から数(+d[1], +d[2], +d[3]);
  if (/^TRUE$/i.test(s)) return true;
  if (/^FALSE$/i.test(s)) return false;
  return Number(s);
}

/* ★★形を 渡せる 口を 使います★★
     ＝`初回端数の価格を試す` は ★測る 道具だけの 口★（お客さんの 道は `初回端数の価格`）
     ＝★自己試験で 前の 形に 戻して 赤に なるのを 見せる★のに 使います */
function 押す(名, a, 形) {
  const b = (a.length > 8 && a[8] !== undefined && !Number.isNaN(a[8])) ? a[8] : 0;
  if (名 === 'ODDFPRICE') {
    return 形
      ? K.初回端数の価格を試す(a[0], a[1], a[2], a[3], a[4], a[5], a[6], a[7], b, 形)
      : K.初回端数の価格(a[0], a[1], a[2], a[3], a[4], a[5], a[6], a[7], b);
  }
  if (名 === 'ODDFYIELD') return K.初回端数の利回り(a[0], a[1], a[2], a[3], a[4], a[5], a[6], a[7], b);
  if (名 === 'ODDLPRICE') {
    const b2 = (a.length > 7 && a[7] !== undefined && !Number.isNaN(a[7])) ? a[7] : 0;
    return K.最終端数の価格(a[0], a[1], a[2], a[3], a[4], a[5], a[6], b2);
  }
  if (名 === 'ODDLYIELD') {
    const b2 = (a.length > 7 && a[7] !== undefined && !Number.isNaN(a[7])) ? a[7] : 0;
    return K.最終端数の利回り(a[0], a[1], a[2], a[3], a[4], a[5], a[6], b2);
  }
  return null;
}

const 行たち = fs.readFileSync(金の道, 'utf-8').split('\n')
  .filter((l) => l && !l.startsWith('#'))
  .map((l) => { const p = l.split('\t'); return { 名: p[0], 式: p[1], 答: p[2], 型: p[3] }; });

/** ★1つの 名前で 合った 本数を 数える★ */
function 数える(名, 形) {
  let 合 = 0, 違 = 0;
  const 外れ = [];
  for (const 行 of 行たち) {
    if (行.名 !== 名) continue;
    const ほ = 式をほどく(行.式);
    if (!ほ) { 違++; continue; }
    let 出;
    try { 出 = 押す(名, 引数を割る(ほ.引数の字).map(値にする), 形); }
    catch (e) { 出 = { 誤り: 'EX' }; }
    const r = 答えを見る(出, 行);
    if (r.合) 合++;
    else { 違++; 外れ.push(行.式 + '\n         正 ' + 行.答 + ' ／ 出 ' + r.出); }
  }
  return { 合, 違, 外れ };
}

console.log('');
console.log('[oddf-susumi] ★保留の 4個（ODDF / ODDL）の 進み★');
console.log('  ★紙★ docs/measured/kansuu46/golden-kane-2026-09-07.tsv（実Excel 16.0 build 20326）');

/* ══ ★★2026-09-18 の 実測★★ ══（★減っても 増えても 赤★）
     ODDFPRICE ... 46 -> ★47 / 48★（`(E - BS) / E` を basis 1 にも 当てた）
     ODDFYIELD ... ★24 / 24★（★全部 合って います★）
     ODDLPRICE ... ★6 / 12★ ／ ODDLYIELD ... ★6 / 12★
     ⇒★残り 1本の 訳は `docs/measured/kansuu46/oddf-nokori-1pon-kiku-koto.md`★ */
/* ★★★私は この 4つを 書き間違えました★★★（2026-09-18）
     ★書いた 数★ 47 / ★23★ / ★0★ / ★0★（★覚えで 書きました★）
     ★実測★ .... 47 / ★24★ / ★6★ / ★6★
     ⇒★★ODDFYIELD は もう 24 / 24（全部 合って います）★★
     ⇒★★ODDL も 半分（6 / 12）は 合って います★★（「手つかず」では ありませんでした）
     ⇒★門を 足した その場で 門が 私を 直しました★
     ⇒★★「手つかず」「まだ 合わない」は ★覚えで 書かない★★★ */
const 決め打ち = {
  /* ★★47 -> 48★★（2026-09-19・★basis 1 の 端の 割る数を 決済の 期に した★）
       ＝awaseru 700本も ★699 -> 700★ */
  ODDFPRICE: { 分母: 48, 合: 48 },
  ODDFYIELD: { 分母: 24, 合: 24 },
  /* ★★6 -> 12★★（2026-09-18・★実Excel 56点から 組み立てを 引き直した★）
       ＝準の並びを ★最終利払日から 1つずつ★ に した（月末に 揃えない）
       ＝NLi を ★その 期の 日数（basis ごと）★に した
       ＝DSC を ★満期が 上に 在る側は NC − A★ に した
       ⇒★12 / 12 と 12 / 12★（★awaseru 700本も 687 -> 699★） */
  ODDLPRICE: { 分母: 12, 合: 12 },
  ODDLYIELD: { 分母: 12, 合: 12 },
};

T('★紙が 読めて いる★（★空振りして いない★）', () => {
  if (行たち.length < 600) throw new Error('紙が ' + 行たち.length + '本しか ない');
});

const 出そろい = {};
for (const 名 of Object.keys(決め打ち)) {
  const 期 = 決め打ち[名];
  const r = 数える(名);
  出そろい[名] = r;
  console.log('  ★' + 名 + '★ ... 合 ' + r.合 + ' / ' + (r.合 + r.違)
    + '（決め打ち ' + 期.合 + ' / ' + 期.分母 + '）');
  T('★' + 名 + ' の 分母が ' + 期.分母 + '本★（★紙の 形が 変わったら 赤★）', () => {
    if (r.合 + r.違 !== 期.分母) {
      throw new Error('出た ' + (r.合 + r.違) + '本 ／ 決め打ち ' + 期.分母 + '本');
    }
  });
  T('★★' + 名 + ' が ' + 期.合 + '本 合う★★（★減っても 増えても 赤★）', () => {
    if (r.合 < 期.合) {
      throw new Error('★下がりました★ ' + r.合 + '本（決め打ち ' + 期.合 + '本）\n       '
        + r.外れ.slice(0, 3).join('\n       '));
    }
    if (r.合 > 期.合) {
      throw new Error('★上がりました★ ' + r.合 + '本（決め打ち ' + 期.合 + '本）'
        + '\n       ★決め打ちを 上げて ください★＝★緩い 門は 門では ありません★');
    }
  });
}

/* ══ ★★わざと 壊す★★ ══（★足しただけでは 効かない＝門は 引き継がれない★） */
T('★★前の 形（`1は外す（前の形）`）に 戻すと 下がる★★（★門が 本当に 見て いる★）', () => {
  const 前 = 数える('ODDFPRICE', '1は外す（前の形）');
  console.log('      ... 今 ' + 出そろい.ODDFPRICE.合 + '本 ／ 前の 形 ' + 前.合 + '本');
  if (前.合 >= 出そろい.ODDFPRICE.合) {
    throw new Error('★前の 形でも 同じ 数です★＝★今日の 直しは 効いて いません★ '
      + 前.合 + ' >= ' + 出そろい.ODDFPRICE.合);
  }
});

T('★★残り 1本の 訳を 書いた 紙が 在る★★（★数だけで 済ませない★）', () => {
  const 紙 = path.join(ROOT, 'docs/measured/kansuu46/oddf-nokori-1pon-kiku-koto.md');
  if (!fs.existsSync(紙)) throw new Error('★紙が 在りません★ ' + 紙);
  const s = fs.readFileSync(紙, 'utf-8');
  if (s.indexOf('ODDFPRICE') < 0 || s.indexOf('2015,2,28') < 0) {
    throw new Error('★紙に 残り 1本の 式が 書かれて いません★');
  }
});

if (自己試験) {
  console.log('');
  console.log('★自分で 壊して 赤に なるか★');
  T('★★決め打ちを 1つ ずらすと 落ちる★★', () => {
    const r = 数える('ODDFPRICE');
    if (r.合 === 決め打ち.ODDFPRICE.合 + 1) throw new Error('ずらして いないのに 落ちた');
    /* ★ここでは 「決め打ち + 1 なら 赤に なる」形を 直に 確かめます★ */
    const にせ = { 合: 決め打ち.ODDFPRICE.合 + 1 };
    if (r.合 >= にせ.合) throw new Error('★決め打ちを 上げても 緑の まま＝門が 効いて いない★');
  });
}

console.log('');
console.log('oddf-susumi: ' + pass + ' 緑 / ' + fail + ' 赤');
process.exit(fail === 0 ? 0 : 1);
