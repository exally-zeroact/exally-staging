/* migi-ni-ugokasu.test.mjs — ★右へ 動かすと 画面が 真っ白に なる★（2026-09-11）
 *
 *  ★★どうやって 見つけたか★★
 *    司さんの 実物（代行計算表2026.xlsb）の「売上表」を 1マスずつ 突き合わせた。
 *    ★404マスが「うちは 何も 出さない」★
 *    ⇒★絵を 撮ったら 本当に 真っ白★でした（8月以降の 欄が 全部 見えない）
 *
 *  ★★正体★★
 *    `xToC(x)` … 画面の 場所から「どの 列か」を 出す。
 *    当たりを ★「全部 既定の 幅」と 思って★ 出し、★前へ しか 歩けません★でした。
 *      `est = floor(px / COL_W) - 5` ⇒ 既定より 広い 列が 在ると ★行き過ぎる★
 *      行き過ぎると 最初の 一歩で `cx > x` に なり ★est-1 を 返して 終わり★
 *    ⇒ 実測 … 売上表（列2〜4が 165点・既定 72点）で scrollLeft=4164 の 時
 *       本当は ★列34★ が 見えて いるのに ★51★ と 答えて いた
 *    ⇒ 描く 所は ★画面の 外の 列★を 描き ★真っ白★に なる
 *
 *  ★★今日 表に 出た 訳★★
 *    ★2026-09-11 に 列の 幅を 読むように した★から。
 *    それまでは ★幅が 1つも 届いて いなかった★ので 全部 既定の 幅＝当たりが 当たって いた。
 *    ⇒★穴は 前から 在り、隠れて いただけ★
 *
 *  ★直し方★ ★行き過ぎたら 戻ります★（行の `yToR` も 同じ 形だったので 一緒に 直した）
 *
 *  ★この 台は エンジンを 建てません★＝`xToC` `yToR` を 切り出して 動かすだけ。
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { 注記を外す } from '../scripts/lib/chuki.mjs';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const 本 = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf-8');

let pass = 0, fail = 0;
const T = (n, f) => {
  try { f(); pass++; console.log('  ✓ ' + n); }
  catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + e.message); }
};

console.log('');
console.log('[migi-ni-ugokasu] ★右へ 動かすと 画面が 真っ白に なる★');

/** ★book.html から 1つの 関数を 切り出す★（写さない＝本物を 動かす） */
function 切る(名) {
  const i = 本.indexOf('function ' + 名 + '(');
  if (i < 0) throw new Error('★' + 名 + ' が 無い★');
  let 深 = 0;
  for (let k = 本.indexOf('{', i); k < 本.length; k++) {
    if (本[k] === '{') 深++;
    else if (本[k] === '}') { 深--; if (!深) return 本.slice(i, k + 1); }
  }
  throw new Error('★' + 名 + ' を 切り出せない★');
}

/* ★実物と 同じ 形の 板★＝既定 72点／広い 列が 3本（司さんの 売上表と 同じ） */
const 板 = { colW: { 2: 165, 3: 165, 4: 165 }, rowH: {}, hiddenCols: {}, hiddenRows: {}, freezeCol: 0, freezeRow: 0 };

function 台(scrollLeft) {
  const 素 = 切る('colX') + ' ' + 切る('cW') + ' ' + 切る('xToC')
    + ' return { xToC: xToC, colX: colX, cW: cW };';
  const f = new Function('sheets', 'activeSheet', 'HDR_W', 'COL_W', 'scale', 'scrollLeft',
    '_隠れ列の並び', '_隠れた数', 'COLS', 素);
  /* ★COLS は 本物と 同じ 16384★（book.html の 頭に 在る 数を 読む＝書き込まない） */
  const m = /var ROW_H = \d+, COL_W = \d+, HDR_W = \d+, HDR_H = \d+, ROWS = \d+, COLS = (\d+)/.exec(本);
  if (!m) throw new Error('★COLS が 読めない★');
  return f([板], 0, 46, 72, 1, scrollLeft, () => [], () => 0, Number(m[1]));
}

T('★台が 組める（空振りして いない）★', () => {
  const t = 台(0);
  if (t.xToC(46) !== 0) throw new Error('★頭で 0 を 返さない … ' + t.xToC(46) + '★');
});

T('★★戻れない 版なら 赤に なる（★これが 真っ白の 正体★）★★', () => {
  /* ★★数を 焼き込みません★★（2026-09-11 一度 焼き込んで しまい 直しました）
     「51 と 答えた」は ★その 板の 数★＝板が 変われば 変わります。
     ⇒★見るのは「答えた 列が 本当に 画面の 端に 居るか」★
     ⇒ そして ★戻れない 版（直す 前）を その場で 作って 赤に なる 事★を 見ます
       ＝この 見張りが 空振りして いない 証し */
  const 素 = 切る('colX') + ' ' + 切る('cW') + ' ' + 切る('xToC').replace(
    /var _戻 = 0;\s*while\(est > 0 && cx > x && _戻 < COLS\)\{ est--; cx -= cW\(est\); _戻\+\+; \}/,
    '/* 戻れない 版（直す 前） */')
    + ' return xToC;';
  if (素.indexOf('戻れない 版') < 0) {
    throw new Error('★戻る 所が 見つからない★＝`xToC` の 形が 変わりました（見張りを 直す）');
  }
  const m = /var ROW_H = \d+, COL_W = \d+, HDR_W = \d+, HDR_H = \d+, ROWS = \d+, COLS = (\d+)/.exec(本);
  const 古い = (sl) => new Function('sheets', 'activeSheet', 'HDR_W', 'COL_W', 'scale',
    'scrollLeft', '_隠れ列の並び', '_隠れた数', 'COLS', 素)(
    [板], 0, 46, 72, 1, sl, () => [], () => 0, Number(m[1]));

  /* ★広い 列を たくさん 置くと 行き過ぎが 大きく なります★（実物が そうだった） */
  const 広い板 = { colW: {}, rowH: {}, hiddenCols: {}, hiddenRows: {}, freezeCol: 0, freezeRow: 0 };
  for (let c = 0; c < 40; c++) 広い板.colW[c] = 165;
  const 元板 = 板.colW;
  板.colW = 広い板.colW;
  let 古ずれ = 0, 新ずれ = 0;
  try {
    for (let sl = 0; sl <= 6000; sl += 53) {
      const t = 台(sl);
      const c1 = t.xToC(46), x1 = t.colX(c1), w1 = t.cW(c1);
      if (!(x1 <= 46 && x1 + w1 > 46)) 新ずれ++;
      const c0 = 古い(sl)(46);
      const x0 = t.colX(c0), w0 = t.cW(c0);
      if (!(x0 <= 46 && x0 + w0 > 46)) 古ずれ++;
    }
  } finally { 板.colW = 元板; }
  if (新ずれ) throw new Error('★今の 版が ' + 新ずれ + '点 ずれて います★');
  if (!古ずれ) {
    throw new Error('★戻れない 版でも ずれませんでした★'
      + '＝この 見張りは 何も 見て いません（空振り）');
  }
  console.log('      … 戻れない 版は ' + 古ずれ + '点 ずれる／今の 版は 0点');
});

T('★★広い 列が 在っても どこでも 合う（総なめ）★★', () => {
  /* ★1点だけ 見て 緑に しない★＝端から 端まで 1点ずつ 確かめる */
  const 違い = [];
  for (let sl = 0; sl <= 6000; sl += 37) {
    const t = 台(sl);
    const c = t.xToC(46);
    const x = t.colX(c), w = t.cW(c);
    if (!(x <= 46 && x + w > 46)) 違い.push('ずれ=' + sl + ' → 列' + c + '（x=' + x + '）');
    if (違い.length > 3) break;
  }
  if (違い.length) throw new Error('★' + 違い.length + '本 合わない★  ' + 違い.join(' ／ '));
  console.log('      … ずれ 0〜6000 を 37点ごと 全部 合う');
});

T('★★行も 同じ 直しが 入って いる★★', () => {
  const 動く = 注記を外す(本, { html: true });
  const i = 動く.indexOf('function yToR');
  const 中 = i >= 0 ? 動く.slice(i, i + 900) : '';
  if (中.indexOf('while(est > 0 && ry > y') < 0) {
    throw new Error('★行（yToR）が 戻れない ままです★'
      + '＝高い 行が 在ると 下へ 動かした 時に 同じ 事が 起きます');
  }
});

console.log('');
console.log('  ' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
