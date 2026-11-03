/* unused-param.mjs — ★「持っているのに 渡していない」を 機械で 探す★ 2026-08-30
 *
 *  ★なぜ★（監査役の 宿題・2026-08-30）
 *    `startEdit(r, c, init)` は ★打った 文字を 受け取る 口（init）を 持っている★のに
 *    ★呼んでいる 所が 1つも 渡していなかった★。
 *    ⇒ ★「持っているのに 渡していない」の コード版★
 *    ⇒ ★1つ 見つかった 型は だいたい 他にも 在る★（今日 何度も 出た）
 *
 *  ★探し方（★書き方に 依らない★）★
 *    ① 注記（コメント）を 外す ← ★注記の 字に 当たって 誤判定するのを 防ぐ★
 *    ② `function 名(a, b, c)` を 拾う
 *    ③ 同じ名前の 呼び出し `名(...)` を ★括弧の 対応で★ 取り、引数の 数を 数える
 *    ④ ★受け取る 口の 数 > どの 呼び出しの 引数の 数 の 最大★ なら 「渡していない」
 *
 *  ★1件ずつ 中身を 見る★＝「使わないのが 正しい物」も 在る
 *    （例＝イベントの 引数 e／既定値だけの 口／外から 呼ばれる 口）
 *
 *  走らせ方: node tools/unused-param.mjs [ファイル…]
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
/* ★借り物は 除く★＝うちが 直せない 物を 数えても 意味が 無い
   （min.js は 圧縮で 引数が 消されているので ほぼ 全部 引っかかる） */
const 借り物 = /\.min\.js$|hyperformula|xlsx\.full/;
/* ★読み込まれた 時に process.argv を 見ない★（2026-08-30 実際に 踏んだ）
   試験から import すると argv に `--self-test` が 入っていて
   ★それを ファイル名と 読み、1つも 読めず 0個★に なった。
   ⇒ ★直に 走らせた 時だけ 引数を 見る★ */
const 走らせた道 = String(process.argv[1] || '').split(path.sep).join('/');
const 直に = 走らせた道.endsWith('tools/unused-param.mjs');
/* ★★見る 範囲を 実物の 画面から 取る★★（2026-09-18・★門の 穴を 塞いだ★）
     ★前は `book.html` `hub.html` ＋ `lib/*.js` だけ★でした
     ⇒★`exally-formula.js` は ★repo の 根★に 在り 範囲の 外でした★
     ⇒`lib/named-ranges.js` の `開く(式, シート)` が
       `exally-formula.js:1221` で `window.名前の箱.開く(f, シート)` と ★毎回 呼ばれて いる★のに
       ★「どこからも 呼ばれない」と 誤報★して いました（経営者1 が 実物を 引いて 見つけた）
     ⇒★★名簿を 手で 持たず 画面の `script src` から 取ります★★
       ＝★画面が 読む 物は 全部 見る★＝★置き場所で 漏れません★
     ★記憶「見張りは 見る 範囲を 先に 数えて 書く」★ */
const 画面 = ['book.html', 'hub.html'];
const 画面が読むjs = new Set();
for (const h of 画面) {
  const p = path.join(ROOT, h);
  if (!fs.existsSync(p)) continue;
  for (const m of fs.readFileSync(p, 'utf8').matchAll(/<script[^>]+src="([^"]+)"/g)) {
    const s = m[1].split('?')[0];
    if (/^https?:/.test(s)) continue;
    if (!s.endsWith('.js')) continue;
    if (!fs.existsSync(path.join(ROOT, s))) continue;
    画面が読むjs.add(s);
  }
}
const 既定 = [...画面, ...fs.readdirSync(path.join(ROOT, 'lib'))
  .filter((f) => f.endsWith('.js')).map((f) => 'lib/' + f),
  ...画面が読むjs].filter((f, i, a) => a.indexOf(f) === i).filter((f) => !借り物.test(f));
const 的 = (直に && process.argv.slice(2).length) ? process.argv.slice(2) : 既定;

/* ★注記外しは 自前で 書かない★＝★共通の 部品を 使う★
 *  ★自前は 2回 外した★（2026-08-30）
 *    ① 改行まで 空白に して ★行番号が ずれた★
 *    ② HTML の `onclick="…"` の 中の // を 注記と 見て ★行ごと 消した★
 *       ⇒ `applyBorderAll(1)` を 見落として ★生きている 口を 死んでいると 誤報★
 *  見張り tests/chuki.test.mjs が ★自前を 赤に する★ 決まりにも 合う。 */
const { 注記を外す } = await import(
  pathToFileURL(path.join(ROOT, 'scripts/lib/chuki.mjs')).href);

/** 名( … ) の 中身を 括弧の 対応で 取る */
function 呼び出し(src, 名) {
  const 出 = [];
  const 印 = 名 + '(';
  let i = 0;
  for (;;) {
    const j = src.indexOf(印, i);
    if (j < 0) break;
    const 前 = src.slice(Math.max(0, j - 24), j);
    /* 定義そのもの・別の 名前の 尻尾は 除く
       ★日本語の 名前も 尻尾に なる★（「探す」は「外のブックを探す」の 尻尾）
       ⇒ 直前が ★字の 続き★なら 別の 関数（2026-08-30 実際に 誤判定した） */
    const 直前 = 前.slice(-1);
    if (/function\s*$/.test(前)) { i = j + 印.length; continue; }
    if (/[A-Za-z0-9_$぀-ヿ一-鿿]/.test(直前)) { i = j + 印.length; continue; }
    /* ★別の 物の 同じ名前の 呼び出し（○○.名( ）も 別物★
       ただし ★その 部品自身の 名前★（ラベル.札 等）は 呼び出しとして 数えたい ので
       ここでは 弾かない。弾くのは ★直前が 字★の 場合だけ。 */
    let 深 = 0, k = j + 名.length;
    for (; k < src.length; k++) {
      if (src[k] === '(') 深++;
      else if (src[k] === ')') { 深--; if (深 === 0) break; }
    }
    const 中 = src.slice(j + 印.length, k);
    let 深2 = 0, 数 = 中.trim() ? 1 : 0;
    for (const ch of 中) {
      if (ch === '(' || ch === '[' || ch === '{') 深2++;
      else if (ch === ')' || ch === ']' || ch === '}') 深2--;
      else if (ch === ',' && 深2 === 0) 数++;
    }
    出.push({ 引数: 数, 行: src.slice(0, j).split(String.fromCharCode(10)).length });
    i = k + 1;
  }
  return 出;
}

/** ★「名前で 呼ぶ」形★を 探す（2026-08-30 追加）
 *  lib/ribbon-actions.js は `呼ぶ('関数名', 引数…)` と ★文字列で 呼ぶ★。
 *  これを 見ないと ★生きている 口を「死んでいる」と 誤報する★。
 *  （実際 `ピボットの窓を開く` を 一度 消しかけた）
 *  ★正規表現を 組み立てない★＝名前に 記号が 混ざると 壊れるので 素直に 探す。 */
function 名前で呼ぶ(src, 名) {
  const 出 = [];
  for (const 引用 of ["'", '"']) {
    const 印 = '呼ぶ(' + 引用 + 名 + 引用;
    let i = 0;
    for (;;) {
      const j = src.indexOf(印, i);
      if (j < 0) break;
      /* 印の 直後から 閉じ括弧までを 見て 引数を 数える */
      let 深 = 1, k = j + '呼ぶ('.length;
      for (; k < src.length; k++) {
        if (src[k] === '(') 深++;
        else if (src[k] === ')') { 深--; if (深 === 0) break; }
      }
      const 残り = src.slice(j + 印.length, k);       /* 名前の 後ろ */
      let 深2 = 0, 数 = 1;                            /* 名前ぶんで 1 */
      for (const ch of 残り) {
        if (ch === '(' || ch === '[' || ch === '{') 深2++;
        else if (ch === ')' || ch === ']' || ch === '}') 深2--;
        else if (ch === ',' && 深2 === 0) 数++;
      }
      /* `呼ぶ('名')` は 名前だけ＝相手に 渡る 引数は 0。
         `呼ぶ('名', a)` は 相手に 1つ 渡る。 */
      出.push({ 引数: 数 - 1 });
      i = k + 1;
    }
  }
  return 出;
}

/* ★呼び出しは ★全部の ファイルを またいで★ 探す★
   （lib で 定義して book.html から 呼ぶ 形が 多い＝1ファイルだけ 見ると 誤報に なる。
     2026-08-30 実際に 3件 誤報した） */
const 読んだ = [];
for (const rel of 的) {
  const p = path.join(ROOT, rel);
  if (!fs.existsSync(p)) continue;
  読んだ.push({ rel, src: 注記を外す(fs.readFileSync(p, 'utf8')) });
}
const 全部 = 読んだ.map((v) => v.src).join(String.fromCharCode(10));

export function 数える() {
  let 全体 = 0, 見つけ = [];
for (const { rel, src } of 読んだ) {
  const 定義 = [...src.matchAll(/function\s+([A-Za-z_$぀-ヿ一-鿿][\w$぀-ヿ一-鿿]*)\s*\(([^)]*)\)/g)];
  for (const m of 定義) {
    const 名 = m[1];
    const 口 = m[2].split(',').map((s) => s.trim()).filter(Boolean);
    if (口.length === 0) continue;
    全体++;
    const 呼 = 呼び出し(全部, 名);
    /* ★「名前で 呼ぶ」形も 見る★（2026-08-30 追加）
       lib/ribbon-actions.js は 呼ぶ('関数名', 引数…) と ★文字列で 呼ぶ★。
       これを 見ないと ★生きている 口を 「死んでいる」と 誤報する★。
       （実際 ピボットの窓を開く を 一度 消しかけた） */
    呼.push(...名前で呼ぶ(全部, 名));
    if (!呼.length) continue;                       /* 呼ばれていない＝別の話 */
    const 最大 = Math.max(...呼.map((v) => v.引数));
    if (最大 < 口.length) {
      見つけ.push({ ファイル: rel, 名, 口: 口.length, 口の名: 口.join(', '),
        最大, 呼び出し: 呼.length,
        行: src.slice(0, m.index).split(String.fromCharCode(10)).length });
    }
  }
}
  return { 全体, 見つけ };
}

/* ★直に 走らせた 時だけ 表示する★
   （試験から 読み込む 時に 表示が 走ると 邪魔に なる） */
if (直に) 見せる();

function 見せる() {
const { 全体, 見つけ } = 数える();
console.log('★見た 関数（引数を 持つ物）= ' + 全体 + '個★');
console.log('★★受け取る 口が 在るのに どこからも 渡していない = ' + 見つけ.length + '個★★');
console.log('');
見つけ.sort((a, b) => (b.口 - b.最大) - (a.口 - a.最大));
for (const v of 見つけ) {
  console.log('  ' + v.ファイル + ':' + v.行 + '  ' + v.名
    + '(' + v.口の名 + ')  ★口' + v.口 + ' / 最大' + v.最大 + '★  呼び出し' + v.呼び出し + 'か所');
}
}
