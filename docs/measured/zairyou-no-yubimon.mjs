/* zairyou-no-yubimon.mjs — ★★紙が「いつの 木で 取られたか」を 紙自身に 持たせる★★（2026-09-15）
 *
 *  ★★なぜ 要るか（今日 実際に 起きた）★★
 *    `golden-marume-mae-ato-2026-09-08.tsv` が ★2026-09-10 に commit されたのに
 *    その中身は ★2026-09-09 の XIRR の 直しより 古い★ ものでした。
 *    ⇒★★走らせずに 置いた 紙★★＝「最後に 走らせた 日で 止まる」より 悪い。
 *    ⇒★★「取った日」は「走らせた 証し」では ありません★★
 *
 *  ★★なぜ git で やらないか（測って 決めました）★★
 *    ★`actions/checkout@v4` は 既定で ★深さ 1★★（`.github/workflows/*.yml` に
 *    `fetch-depth` の 指定が 1つも 無い＝実測）。
 *    ⇒★CI では `git log -- そのファイル` が 使えません★＝★手元だけ 緑・CI は 赤★の 形。
 *    ⇒★★git を 使わず 中身の 指紋で やります★★
 *
 *  ★★何を 指紋に するか★★
 *    ★その 回に ★本当に 読んだ 本★だけ★（`require` の 覚えから 取る）
 *    ⇒★道具だけでは 足りません★＝★今日の XIRR は ★台（lib）★が 変わった★
 *    ⇒★広く 取り過ぎると 関係ない 直しで 赤に なります★（★赤に しすぎない★）
 *    ⇒★「読んだ 本」＝ちょうど その 間★
 *
 *  ★★使い方★★
 *    import { 指紋の行, 突き合わせ } from './zairyou-no-yubimon.mjs';
 *    const 頭 = 指紋の行(require_, import.meta.url);   // 紙の 頭に 入れる 行の 並び
 *
 *  ★★見て いない 事★★
 *    ・★`hyperformula.full.min.js` の ような 借り物も 指紋に 入ります★
 *      ＝★借り物の 版が 変われば 赤★（★それが 正しい★＝答えが 変わり得る）
 *    ・★紙を 手で 直したら 指紋は 合いません★＝★手で 直さない★
 *    ・★読んだ 本が 0本の 道具は 守れません★（★その 時は 行が 出ません★）
 *    ・★★ESM は 字で 追って います★★＝★下の 形は 抜けます★
 *      ㋓★道を 組み立てる★（`'honban' + '-no-michi.mjs'`）
 *      ㋔★`docs/measured/` の 外の ESM★（今は 1つも 在りません）
 *      ㋕★引数で 渡される 道★（`EFの道` の ような 物）
 *        ⇒★`EFの道` は `require` で 読むので ㋐で 拾えて います★
 *      ⇒★抜けたら ★紙は 古いのに 緑★ に なります★（★それを 書いて おきます★）
 */
import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import { fileURLToPath } from 'node:url';

const ここ = path.dirname(fileURLToPath(import.meta.url));
export const ROOT = path.join(ここ, '..', '..');

const 印 = '# ★材料★\t';

function 一本の指紋(道) {
  return crypto.createHash('sha256').update(fs.readFileSync(道)).digest('hex').slice(0, 16);
}

/* ★★ESM（`await import`）の 本は `require` の 覚えに 入りません★★（2026-09-15 に 踏んだ）
     ★最初の 作りでは ★`honban-no-michi.mjs` が 1行も 入って いませんでした★
     ＝★その 本が `smartRounding:false` と プラグイン 8本を 決めて います★
     ⇒★そこを 変えても 紙は 赤に ならない★＝★一番 大事な 本が 抜けて いた★
   ★node に ESM の 覚えを 読む 公の 口は ありません★
   ⇒★字で 追います★＝★`docs/measured/○○.mjs` と 書いて ある 本を 拾い、
     ★その 本が 指す 本も 追う★（★共通の 本は 必ず この 形で 書かれて います★）
   ★これは 字で 探す 手★＝★抜ける 形が 在り得ます★
     ⇒★見て いない 形★を 下に 名指しで 書きます（見張りの 頭にも） */
const 測りの本 = /docs\/measured\/([A-Za-z0-9._-]+\.mjs)/g;

function 測りの本を追う(道, 見た) {
  if (見た.has(道)) return;
  見た.add(道);
  let s;
  try { s = fs.readFileSync(道, 'utf8'); } catch (e) { return; }
  測りの本.lastIndex = 0;
  let m;
  while ((m = 測りの本.exec(s)) !== null) {
    const p = path.join(ROOT, 'docs/measured', m[1]);
    if (fs.existsSync(p)) 測りの本を追う(path.resolve(p), 見た);
  }
}

/** ★その 回に 読んだ 本を 集める★（★repo の 中だけ／node_modules は 見ない★） */
export function 読んだ本(require_, 自分のURL) {
  const 出 = new Set();
  /* ㋐★CommonJS（`require`）の 覚え★… 台（`lib/*`）や 借り物は ここに 入る */
  const 覚え = (require_ && require_.cache) ? Object.keys(require_.cache) : [];
  for (const k of 覚え) {
    const p = path.resolve(k);
    if (!p.startsWith(path.resolve(ROOT))) continue;
    if (p.indexOf('node_modules') >= 0) continue;
    出.add(p);
  }
  /* ㋑★ESM（`await import`）は 字で 追う★… ★自分から 始めて 指す 先まで★ */
  if (自分のURL) {
    const 見た = new Set();
    測りの本を追う(path.resolve(fileURLToPath(自分のURL)), 見た);
    for (const p of 見た) 出.add(p);
  }
  return [...出].map((p) => path.relative(ROOT, p).replace(/\\/g, '/')).sort();
}

/** ★紙の 頭に 入れる 行★（1本につき 1行） */
export function 指紋の行(require_, 自分のURL) {
  const 本 = 読んだ本(require_, 自分のURL);
  const 行 = [
    '# ★★材料の 指紋★★ … ★この 紙は 下の 本を 読んで 取りました★',
    '#   ★下の どれか 1つでも 変わったら この 紙は 古い★'
      + '＝`tests/kami-no-zairyou.test.mjs` が 赤に します',
    '#   ★「取った日」は「走らせた 証し」では ありません★（2026-09-15 に 実際に 起きた）',
  ];
  for (const f of 本) 行.push(印 + f + '\t' + 一本の指紋(path.join(ROOT, f)));
  return 行;
}

/** ★紙の 中の 指紋を 今の 中身と 突き合わせる★ */
export function 突き合わせ(紙の道) {
  const s = fs.readFileSync(紙の道, 'utf8');
  const 見た = [], 古い = [], 無い = [];
  for (const 行 of s.split('\n')) {
    if (行.indexOf(印) !== 0) continue;
    const [, f, h] = 行.split('\t');
    if (!f || !h) continue;
    見た.push(f);
    const p = path.join(ROOT, f);
    if (!fs.existsSync(p)) { 無い.push(f); continue; }
    if (一本の指紋(p) !== h.trim()) 古い.push(f);
  }
  return { 見た, 古い, 無い };
}
