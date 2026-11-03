/* ★A群の 丸めを 外す 前と 後を ★同じ 紙・同じ 行★で 押し比べる★（2026-09-08）
 *
 *  ★何を 比べるか★
 *    ★前★ … git の HEAD の `exally-formula.js`（★丸めが 在る★）
 *    ★後★ … 今 手元に 在る `exally-formula.js`（★丸めを 外した★）
 *    ★正★ … 実Excel の 実測（golden の 紙）★＝答えは 実Excel だけが 正★
 *
 *  ★★私の 理屈に 寄りかからない★★
 *    「丸めを 外しただけ だから 前は 後を 丸めた 物の はず」＝★私の 頭の 中の 話★
 *    ⇒★本当に 前の コードを 読み込んで 押します★
 *
 *  ★見る 3か所（JS層を 通る 物）★
 *    XIRR              exally-formula.js
 *    LINEST            exally-formula.js
 *    BINOM.DIST.RANGE  exally-formula.js
 *  ★★book.html の 数式バーの 逃がし道は この 道具では 押せません★★
 *    ⇒ node では 画面が 無い＝★別に 本物の ブラウザで 押す 要が 在る★
 *    ⇒★「押して いない」と 出します（★測って いないを 緑に しない★）★
 *
 *  使い方: node docs/measured/osu-marume-mae-ato.mjs
 */
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createRequire } from 'node:module';

/* ★手元の 絶対の 道を 焼き込まない★ */
const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..', '..');
const 出す先 = path.join(ROOT, 'docs/measured/golden-marume-mae-ato-2026-09-08.tsv');

/* ══ ★押す 台を 1つ 作る（前の 版／後の 版 で 1つずつ）★ ══════════
   ★★測り台には ★本番が 積んで いる 物を 全部★ 積む★★（2026-09-09 に 直した）
     ★1回目は プラグインを ★1本も★ 積んで いなかった★
     ⇒ 本番（book.html）は ★8本★ 読む
     ⇒★本番に 在る 物が 無い 状態で 押して いた＝★嘘の 数字が 出る恐れ★
     ⇒ この 台で 出した「直った18／壊れた0」は ★積み直して 取り直す★
   ★同じ 型の 再発です★（同じ日に 別の 道具でも 踏んだ）
   ⇒★book.html が 読む 数と 合わなければ ★止める★★ */
function 本番のプラグイン数() {
  const html = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf-8');
  return new Set([...html.matchAll(/lib\/(formula-[a-z]+)-plug\.js/g)].map((m) => m[1])).size;
}
/* ★★本番の 道は 1本★★（2026-09-15）＝`docs/measured/honban-no-michi.mjs`
   ★この 道具は 丸めの 前後を 比べる★ので
   ★どの `exally-formula.js` を 読むか★を ★口の 引数★で 渡します
   （★別の 建て方を 作らない★＝指示役1 の 決め） */
async function 台を作る(式ファイル) {
  const { 建てる } = await import(pathToFileURL(path.join(ROOT, 'docs/measured/honban-no-michi.mjs')).href);
  const 道 = await 建てる({ EFの道: 式ファイル });
  return { hf: 道.hf, SID: 道.SID,
    押す: 道.EF._jsComputeFormula, 変換: 道.EF.convertFormula, 積んだ: 道.積んだ };
}

/* ══ ★★材料も 正解も ★紙から 読む★（★私が 書き写さない★）★★ ══════
   ★2026-09-08 に 写し間違いを 2回 踏んだ★
     ①保留 27行の 土台（D1=0.5 と 書いたが 本当は =DATE(2024,1,1)）
     ②この 押し比べの XIRR の お金（30,000,000 と 書いたが 本当は 27,500,000…）
   ⇒★人が 写す 限り また 間違える＝★機械に 読ませる★★
   ⇒ 取り方の 紙に `#材料<タブ>マス<タブ>値` を 載せさせ、ここでは それを 読む */
/* ★★紙ごとに 列の 並びが 違う★★（2026-09-08 に 踏んだ）
     A群/LINEST の 紙 … 関数／式／答え／型／…
     BINOM の 紙 ……… 式／答え／型／…
   ⇒★1つの 読み方で 全部 読めると 思い込んで ★0行★ 拾って いた★
   ⇒★紙ごとに どの 列かを 名指しで 渡す★ */
function 紙を読む(p, 列) {
  const 材料 = {};        /* 'A1' → 数 */
  const 答え = [];        /* { 式, 実Excel, 型 } */
  for (const l of fs.readFileSync(p, 'utf-8').split('\n')) {
    if (l.startsWith('#材料')) {
      const c = l.split('\t');
      if (c.length >= 3) 材料[c[1].trim()] = Number(c[2]);
      continue;
    }
    if (l.startsWith('#') || !l.trim()) continue;
    const c = l.split('\t');
    if (c.length <= Math.max(列.式, 列.答え, 列.型)) continue;
    答え.push({ 式: c[列.式].trim(), 実Excel: c[列.答え].trim(), 型: c[列.型].trim() });
  }
  if (答え.length === 0) { console.error('★紙から 1行も 読めなかった … ' + p + '★＝★列の 並びが 違う★'); process.exit(2); }
  return { 材料, 答え };
}

/* ★マスの 名前（A1）を 行・列に する★ */
function マスを解く(s) {
  const m = /^([A-Z]+)(\d+)$/.exec(s);
  if (!m) return null;
  let c = 0;
  for (const ch of m[1]) c = c * 26 + (ch.charCodeAt(0) - 64);
  return { 行: Number(m[2]) - 1, 列: c - 1 };
}

/* ★材料を 押す 台に 敷く★ */
function 材料を敷く(台, 材料) {
  let 最大行 = 0, 最大列 = 0;
  const 置く = [];
  for (const k of Object.keys(材料)) {
    const rc = マスを解く(k);
    if (!rc) continue;
    置く.push([rc, 材料[k]]);
    if (rc.行 > 最大行) 最大行 = rc.行;
    if (rc.列 > 最大列) 最大列 = rc.列;
  }
  const 板 = [];
  for (let r = 0; r <= 最大行 + 2; r++) 板.push(new Array(最大列 + 3).fill(null));
  for (const [rc, v] of 置く) 板[rc.行][rc.列] = v;
  台.hf.setSheetContent(台.SID, 板);
}

const A紙 = path.join(ROOT, 'docs/measured/golden-marume-A-2026-09-08.tsv');
const L紙 = path.join(ROOT, 'docs/measured/golden-linest-2026-09-08.tsv');
const B紙 = path.join(ROOT, 'docs/measured/golden-marume-A4-binom-2026-09-08.tsv');
for (const p of [A紙, L紙, B紙]) {
  if (!fs.existsSync(p)) { console.error('★紙が 無い … ' + p + '★'); process.exit(2); }
}
const A = 紙を読む(A紙, { 式: 1, 答え: 2, 型: 3 });
const L = 紙を読む(L紙, { 式: 1, 答え: 2, 型: 3 });
const B = 紙を読む(B紙, { 式: 0, 答え: 1, 型: 2 });
for (const [名, x] of [['A群', A], ['LINEST', L]]) {
  if (Object.keys(x.材料).length === 0) {
    console.error('★' + 名 + 'の 紙に `#材料` が 無い＝★取り方を 走らせ直して ください★');
    process.exit(2);
  }
}

/* ★一致の 幅（★測ってから 決めた★・紙の 頭に 根拠が 書いて 在る）★ */
const 幅 = {
  XIRR: { 値: 1e-6, 訳: '実Excel 自身が 見当（第3引数）で 最大 9.65e-8 ばらつく＝その 10倍' },
  /* ★★2026-09-08 に ★丸めを 外してから★ 決めた（監査②「仮の まま PR に 載せない」）★★
       ★決める 前に 測った★（この 紙の 7本・丸めを 外した 後）
         =LINEST(K1:K5,L1:L5) … 相対 6.023e-16 ←★一番 大きい★
         =LINEST(A1:A5,B1:B5) … 相対 3.553e-16
         =LINEST(C1:C5,D1:D5) … 相対 2.085e-16
         =LINEST(E1:E5,F1:F5) … 相対 1.943e-16
         =LINEST(G1:G5,H1:H5) … 相対 1.665e-16
         =LINEST(I1:I7,J1:J7) / =LINEST(M1:M5,N1:N5) … ★ぴったり 一致★
       ⇒★一番 大きい 6.023e-16 は 倍精度の 1刻み（2.220e-16）の 2.7個分★
       ⇒★XIRR と 同じ 決め方（実測の 最大の 約10倍）で ★1e-14★★
       ★★この 幅で 見逃す 悪さ（守る 範囲を 書く）★★
         ⇒★相対 1e-14 より 小さい 狂いは この 紙では 捕まりません★
         ⇒ 1e-14 は お金で 言うと 1億円に つき 0.000001円＝★実務では 効かない 大きさ★ */
  LINEST: { 値: 1e-14, 訳: '丸めを 外した 後の 実測 7本の ばらつき 最大 6.023e-16（倍精度 2.7刻み分）の 約10倍' },
  BINOM: { 値: 1e-9, 訳: '見当の 繰り返しが 無い＝実Excel と 桁まで 合う' },
};

/* ★押す 組を 紙から 作る★（★見当つき／誤りの 行は 別に 扱う★） */
const 組ら = [];
for (const x of A.答え) {
  if (!/^=XIRR\([A-Z]+\d+:[A-Z]+\d+,[A-Z]+\d+:[A-Z]+\d+\)$/.test(x.式)) continue;   /* ★見当なしの 形だけ★（うちの JS層は 2引数しか 受けない） */
  組ら.push({ 名: 'XIRR', 材料: A.材料, 式: x.式, 実Excel: x.実Excel, 型: x.型, 幅: 幅.XIRR });
}
/* ★LINEST は ★裸の 形★だけ★（うちの JS層が 拾うのは `=LINEST(範囲,範囲)` だけ）
   ⇒ 紙には `=INDEX(LINEST(…),1,1)` も 在るが、それは ★入れ子＝別件★ */
for (const x of L.答え) {
  if (!/^=LINEST\([A-Z]+\d+:[A-Z]+\d+,[A-Z]+\d+:[A-Z]+\d+\)$/i.test(x.式)) continue;
  組ら.push({ 名: 'LINEST', 材料: L.材料, 式: x.式, 実Excel: x.実Excel, 型: x.型, 幅: 幅.LINEST });
}
for (const x of B.答え) {
  if (!/^=BINOM\.DIST\.RANGE\(/i.test(x.式)) continue;
  組ら.push({ 名: 'BINOM.DIST.RANGE', 材料: {}, 式: x.式, 実Excel: x.実Excel, 型: x.型, 幅: 幅.BINOM });
}

/* ══ ★前の 版を 取り出す（★丸めが 在る 版を git から 自分で 探す★）★ ══
   ★はじめは `HEAD:exally-formula.js` を 見て いました★が、
   ★直しを commit した 途端に HEAD も 直った 版に なる＝比べる 相手が 消える★。
   ⇒★履歴を さかのぼって ★丸めが 在る 一番 新しい 版★を 自分で 探す★
   ⇒ 直した 後も・merge した 後も この 紙は 取り直せます */
const 丸めの印 = /Math\.round\(_jsXirr\(vals,dates\)\*10000\)\/10000/;
let 前の字 = null, 前の版 = null;
const 履歴 = execFileSync('git', ['rev-list', '--max-count=300', 'HEAD', '--', 'exally-formula.js'],
  { cwd: ROOT, maxBuffer: 8 * 1024 * 1024 }).toString('utf-8').split('\n').filter(Boolean);
for (const h of 履歴) {
  let s;
  try { s = execFileSync('git', ['show', h + ':exally-formula.js'], { cwd: ROOT, maxBuffer: 64 * 1024 * 1024 }).toString('utf-8'); }
  catch (e) { continue; }
  if (丸めの印.test(s)) { 前の字 = s; 前の版 = h; break; }
}
if (前の字 === null) {
  console.error('★履歴 ' + 履歴.length + '本を 見ましたが ★丸めが 在る 版が 見つかりません★＝比べる 相手が 無い');
  process.exit(2);
}
const 仮置き = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'marume-')), 'exally-formula.js');
fs.writeFileSync(仮置き, 前の字);

const 前台 = await 台を作る(仮置き);
const 後台 = await 台を作る(path.join(ROOT, 'exally-formula.js'));

/* ★★JS層で 止めず エンジンまで 通す（2026-09-10 に 直した）★★
   ★前は JS層が null を 返した 時点で 諦めて いました★
   ⇒ #57 で LINEST が ★JS層 → エンジン★ に 移った 後、
     この 道具は ★「JS層が 受け持って いない」を 3本 出し「壊れた 3」と 数えた★
   ⇒★壊れて いません＝道具が 本番の 道の 途中で 止まって いただけ★
   ⇒★本番は ①JS層 → ②convertFormula → エンジン★＝★同じ 道で 押す★ */
const 式の行 = 60;
function 押す(台, 組) {
  const 板 = [];
  if (Object.keys(組.材料).length) 材料を敷く(台, 組.材料);
  else 台.hf.setSheetContent(台.SID, [[0]]);
  try {
    const r = 台.押す(0, 組.式);
    if (r !== null && r !== undefined) return String(r);
  } catch (e) { return '★投げた（' + (e && e.message) + '）★'; }
  /* ★JS層が 受けない＝エンジンに 聞く★ */
  try {
    const 盤 = 板;
    const 元 = 台.hf.getSheetSerialized ? null : null;
    /* ★材料を もう一度 敷いてから 式を 下に 置く★ */
    if (Object.keys(組.材料).length) 材料を敷く(台, 組.材料);
    else 台.hf.setSheetContent(台.SID, [[0]]);
    台.hf.setCellContents({ sheet: 台.SID, row: 式の行, col: 0 }, 台.変換(組.式));
    const v = 台.hf.getCellValue({ sheet: 台.SID, row: 式の行, col: 0 });
    if (v && v.type) return '#' + v.type;
    return v === null || v === undefined ? '(空)' : String(v);
  } catch (e) { return '★エンジンで 投げた（' + (e && e.message) + '）★'; }
}

const 行 = [];
const 言う = (s) => { 行.push(s); console.log(s); };
言う('# ★A群の 丸めを 外す 前と 後を 同じ 紙・同じ 行で 押し比べた★（2026-09-08）');
言う('#');
言う('# ★前★ 丸めが 在る 一番 新しい 版 … ' + 前の版 + '（★履歴から 自分で 探した★）');
言う('# ★後★ 今 手元の exally-formula.js（丸めを 外した）');
言う('# ★正★ 実Excel の 実測（golden の 紙）');
言う('#   ⇒★「丸めを 外しただけ だから 同じ はず」は 私の 頭の 中の 話★');
言う('#     ⇒★本当に 前の コードを 読み込んで 押した★');
言う('#');
言う('# ★★book.html の 数式バーの 逃がし道は ★押して いません★★★');
言う('#   node には 画面が 無い＝★本物の ブラウザで 押す 要が 在る★');
言う('#   ⇒★「押して いない」と 書く（測って いないを 緑に しない）★');
言う('#');
言う('# ★材料も 正解も ★紙から 読んだ★（私が 書き写して いない）');
言う('#   A群の 紙 … ' + Object.keys(A.材料).length + '個の マス ／ ' + A.答え.length + '行');
言う('#   BINOM の 紙 … ' + B.答え.length + '行');
言う('#');
言う('# 名' + '\t' + '式' + '\t' + '前(丸めが在る)' + '\t' + '後(外した)' + '\t' + '実Excel' + '\t' + '判じ');

let 直った = 0, 壊れた = 0, 変わらず = 0, 判ぜず = 0, 誤り行 = 0;
for (const 組 of 組ら) {
  const 前 = 押す(前台, 組);
  const 後 = 押す(後台, 組);
  const 正 = 組.実Excel;

  /* ★実Excel が 誤りを 返した 行は 別に 数える（丸めの 話では ない）★ */
  if (組.型 === 'Error' || /^#/.test(正)) {
    言う(組.名 + '\t' + 組.式 + '\t' + 前 + '\t' + 後 + '\t' + 正 + '\t★実Excel が 誤り＝丸めの 話では ない（別に 見る）★');
    誤り行++;
    continue;
  }

  const 近い = (s) => {
    if (組.幅.値 === null) return null;
    const a = Number(s), b = Number(正);
    if (!isFinite(a) || !isFinite(b)) return false;
    if (b === 0) return a === 0;
    return Math.abs((a - b) / b) <= 組.幅.値;
  };
  const 前近 = 近い(前), 後近 = 近い(後);

  let 判;
  if (前近 === null) { 判 = '★判じられない（幅を まだ 決めて いない）★'; 判ぜず++; }
  else if (!前近 && 後近) { 判 = '★★直った★★'; 直った++; }
  else if (前近 && !後近) { 判 = '★★壊れた★★'; 壊れた++; }
  else if (前近 && 後近) { 判 = '変わらず（前から 合って いた）'; 変わらず++; }
  else { 判 = '★★どちらも 幅の 外＝丸めでは 直らない（別の 欠陥）★★'; 変わらず++; }

  言う(組.名 + '\t' + 組.式 + '\t' + 前 + '\t' + 後 + '\t' + 正 + '\t' + 判);
}

言う('#');
言う('# ★★締め★★ ★直った ' + 直った + '★ ／ ★壊れた ' + 壊れた + '★ ／ 変わらず ' + 変わらず
  + ' ／ 判じられない ' + 判ぜず + ' ／ 実Excel が 誤りの 行 ' + 誤り行);
言う('#   ―― 足すと ' + (直った + 壊れた + 変わらず + 判ぜず + 誤り行) + ' ／ 全 ' + 組ら.length + '行');
言う('#   ★押して いない … book.html の 数式バーの 逃がし道 1か所★');
for (const k of Object.keys(幅)) 言う('#   幅 … ' + k + ' … ' + (幅[k].値 === null ? '★未決★' : 幅[k].値) + '（' + 幅[k].訳 + '）');

fs.writeFileSync(出す先, 行.join('\n') + '\n');
console.log('\n★書いた … ' + 出す先 + '★');
if (壊れた > 0) { console.error('★壊れた が 在る＝出しません★'); process.exit(3); }
