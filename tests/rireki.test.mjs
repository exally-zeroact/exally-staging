/* rireki.test.mjs — ★6 履歴＝見る場所★
 *
 *  ★指示役の検証要件（2026-08-26）★
 *    ★履歴が 別の入り口から 同じに見える事★／★客のブックにタブ0件★
 *
 *  使い方: node tests/rireki.test.mjs [--self-test]
 */
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createRequire } from 'node:module';
import { 注記を外す } from '../scripts/lib/chuki.mjs';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(__dirname, '..');
const SELF = process.argv.includes('--self-test');
const OVERRIDE = process.env.EXALLY_RIREKI_OVERRIDE ? JSON.parse(process.env.EXALLY_RIREKI_OVERRIDE) : {};
const require_ = createRequire(pathToFileURL(path.join(ROOT, 'package.json')));
const Rireki = require_(OVERRIDE['lib/rireki.js'] || path.join(ROOT, 'lib/rireki.js'));
const 読む = (rel) => fs.readFileSync(OVERRIDE[rel] || path.join(ROOT, rel), 'utf8');

let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ok   ' + n); } catch (e) { fail++; console.log('  NG   ' + n + '\n       ' + (e && e.message)); } };
const ok = (c, m) => { if (!c) throw new Error(m || 'expected truthy'); };
const eq = (a, b, m) => { if (a !== b) throw new Error((m ? m + ': ' : '') + '期待=' + JSON.stringify(b) + ' 実際=' + JSON.stringify(a)); };

const 今 = new Date('2026-08-26T09:00:00+09:00');

console.log('');
console.log('[rireki] ★履歴＝見る場所★');

/* ══ ①日付の書き方（司さんの指定）══ */
T('★今年は「8月25日 14:03」★', () => {
  eq(Rireki.日付の字(new Date('2026-08-25T14:03:00+09:00'), 今), '8月25日 14:03');
});
T('★去年は「2025年8月25日」（時刻は書かない）★', () => {
  eq(Rireki.日付の字(new Date('2025-08-25T14:03:00+09:00'), 今), '2025年8月25日');
});
T('★1桁の分も 0を付ける★', () => {
  eq(Rireki.日付の字(new Date('2026-01-05T09:07:00+09:00'), 今), '1月5日 09:07');
});
T('★読めない日付でも 落ちない★', () => {
  eq(Rireki.日付の字('あ', 今), '');
});

/* ══ ②カテゴリ ══ */
T('★カテゴリは 決めた6つ★', () => {
  eq(Rireki.種類の順.join(','), '関数,自動化,直した所,診断,取り込み・書き出し,その他');
});
T('★知らない種類は「その他」に寄せる（捨てない）★', () => {
  eq(Rireki.種類に寄せる('なにか'), 'その他');
  eq(Rireki.種類に寄せる('診断'), '診断');
  eq(Rireki.種類に寄せる(''), 'その他');
});

/* ══ ③一覧 ══ */
const 行たち = [
  { いつ: '2026-08-20T10:00:00+09:00', 種類: '診断', 見出し: '危ない所 122か所', 中身: { やった: '調べた', 何か所: 122 }, ファイル名: '代行計算表2026.xlsb' },
  { いつ: '2026-08-25T14:03:00+09:00', 種類: '自動化', 見出し: '税込の列を足した', 中身: { 聞いた: '税込がほしい', やった: '列を足した', 何か所: 31 }, クレジット: 1 },
  { いつ: '2025-12-31T23:59:00+09:00', 種類: '関数', 見出し: 'VLOOKUPの直し', 中身: {} },
  { いつ: '2026-08-24T08:00:00+09:00', 種類: 'なにか', 見出し: 'その他の物', 中身: {} },
];
T('★新しい物が 上★', () => {
  const r = Rireki.一覧にする(行たち, { 今 });
  eq(r.出す[0].見出し, '税込の列を足した');
  eq(r.出す[r.出す.length - 1].見出し, 'VLOOKUPの直し');
});
T('★種類ごとの件数を 出す★', () => {
  const r = Rireki.一覧にする(行たち, { 今 });
  eq(r.種類ごと['診断'], 1);
  eq(r.種類ごと['その他'], 1, '★知らない種類が 数から消えている★');
  eq(r.種類ごと['取り込み・書き出し'], 0);
});
T('★種類で絞れる★', () => {
  const r = Rireki.一覧にする(行たち, { 今, 種類: '診断' });
  eq(r.件数, 1);
  eq(r.出す[0].見出し, '危ない所 122か所');
});
T('★多い時は 何件 出していないかを 言う（黙って切らない）★', () => {
  const 多い = [];
  for (let i = 0; i < 500; i++) 多い.push({ いつ: '2026-08-0' + (i % 9 + 1) + 'T10:00:00+09:00', 種類: '診断', 見出し: 'x' + i });
  const r = Rireki.一覧にする(多い, { 今, 上限: 200 });
  eq(r.出す.length, 200);
  eq(r.出していない, 300, '★黙って 切っている★');
});
T('★空でも 落ちない★', () => {
  const r = Rireki.一覧にする([], { 今 });
  eq(r.件数, 0);
  eq(r.出していない, 0);
});

/* ══ ④中身（どのセルが どう変わったか）══ */
T('★何を聞いて 何をして どのセルが どう変わったかを 出す★', () => {
  const s = Rireki.中身の字({
    中身: { 聞いた: '税込がほしい', やった: '列を足した', 何か所: 2, どこ: [{ 場所: 'D2', 前: '', 後: '1100' }, { 場所: 'D3', 前: '', 後: '2200' }] },
    クレジット: 1,
  });
  ok(s.indexOf('税込がほしい') >= 0, '★聞いた事を 出していない★');
  ok(s.indexOf('D2') >= 0 && s.indexOf('（空） → 1100') >= 0, '★どう変わったかを 出していない★：' + s);
  ok(s.indexOf('使ったクレジット：1') >= 0, '★使ったクレジットを 出していない★');
});
T('★AIを呼んでいない時は「AIは使っていません」と書く★', () => {
  const s = Rireki.中身の字({ 中身: { やった: '調べた' }, クレジット: 0 });
  ok(s.indexOf('AIは使っていません') >= 0, '★0円の道だと 言っていない★');
  ok(s.indexOf('1円もかかりません') < 0, '★AIを呼ぶ道にも使える言い方をしている★');
});
T('★多い時は「ほか◯か所」（全部は出さない）★', () => {
  const どこ = [];
  for (let i = 0; i < 12; i++) どこ.push({ 場所: 'D' + i, 前: '', 後: i });
  const s = Rireki.中身の字({ 中身: { どこ }, クレジット: 0 });
  ok(s.indexOf('ほか 7か所') >= 0, '★何か所 出していないかを 言っていない★：' + s);
});
T('★客に見せる字に ★ を書かない★', () => {
  const r = Rireki.一覧にする(行たち, { 今 });
  for (const x of r.出す) {
    ok(x.日付.indexOf('★') < 0 && x.種類.indexOf('★') < 0, '★客の字に ★ が出ている★');
  }
  ok(Rireki.中身の字({ 中身: { やった: 'あ' } }).indexOf('★') < 0);
});

/* ══ ★⑤別の入り口から 同じに見える★ ══ */
T('★入り口が違っても 通る所は 1つ（一覧にする）★', () => {
  const book = 読む('book.html');
  const 動く所 = 注記を外す(book, { html: true });
  const 呼ぶ数 = (動く所.match(/Rireki\.一覧にする\(/g) || []).length;
  ok(呼ぶ数 >= 1, '★画面が 一覧にする を呼んでいない★');
  /* ★入り口は2つ（AIタブ／履歴ページ）だが、組み立てる所は 1つだけ★ */
  /* ★名前を少し変えただけの「2つ目」も 数える★（履歴を描く2 のような物）
     ＝わざと壊した時に 素通りしたので 直した（2026-08-26） */
  const 組み立て = (動く所.match(/function 履歴を描く[^(]*\(/g) || []).length;
  eq(組み立て, 1, '★組み立てる所が 2つ在る（同じに見えなくなる）★');
  const 入り口 = (動く所.match(/openRireki\(/g) || []).length;
  ok(入り口 >= 2, '★入り口が 2つ 無い（AIタブと 履歴ページ）★');
});
T('★同じ中身なら 入り口が違っても 出る物は 同じ★', () => {
  const a = Rireki.一覧にする(行たち, { 今 });
  const b = Rireki.一覧にする(行たち.slice(), { 今 });
  eq(JSON.stringify(a), JSON.stringify(b));
});

/* ══ ★⑥客のブックに タブを1つも足さない★ ══ */
T('★客のブックに 履歴のタブを 足していない★', () => {
  const book = 注記を外す(読む('book.html'), { html: true });
  /* ★シートを足す所（addSheet 等）に 履歴の名前が 混ざっていないか★ */
  for (const だめ of ['履歴シート', 'Claude Log', 'ClaudeLog', 'AIログ']) {
    ok(book.indexOf(だめ) < 0, '★客のブックに「' + だめ + '」を 足そうとしている★');
  }
  /* ★履歴は 倉庫（rireki）に置く★ */
  ok(book.indexOf('exally.rireki') >= 0 || book.indexOf('/rest/v1/rireki') >= 0 || book.indexOf('RirekiStore') >= 0,
    '★履歴の置き場（倉庫）が 画面から 見えない★');
});
T('★履歴は 見る場所（ここから 何かを始めさせない）★', () => {
  const book = 注記を外す(読む('book.html'), { html: true });
  const i = book.indexOf('function 履歴を描く');
  ok(i > 0, '★組み立てる所が 無い★');
  const 所 = book.slice(i, book.indexOf('function', i + 20) > 0 ? book.indexOf('function 履歴', i + 20) + 4000 : i + 4000);
  for (const だめ of ['saveXlsx(', 'onPickBookFile(', '_aiFetch(']) {
    ok(所.indexOf(だめ) < 0, '★履歴から ' + だめ + ' を始めている（見る場所ではない）★');
  }
});

console.log('');
console.log('  ' + pass + ' 緑 / ' + fail + ' 赤');

if (SELF) {
  const { spawnSync } = await import('node:child_process');
  const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'exally-rireki-'));
  console.log('');
  console.log('[self-test] わざと壊して 赤くなるかを数える（★repo は読むだけ★）');
  const BREAKS = [
    ['lib/rireki.js', '★今年でも 年を書く★', (s) => s.replace('if (d.getFullYear() === n.getFullYear()) {', 'if (false) {')],
    ['lib/rireki.js', '★去年でも 時刻を書く★', (s) => s.replace("return d.getFullYear() + '年' + 月 + '月' + 日 + '日';", "return 月 + '月' + 日 + '日 00:00';")],
    ['lib/rireki.js', '★知らない種類を 捨てる★', (s) => s.replace("return 種類の順.indexOf(t) >= 0 ? t : 'その他';", "return 種類の順.indexOf(t) >= 0 ? t : '';")],
    ['lib/rireki.js', '★古い物を 上にする★', (s) => s.replace('return new Date(b.いつ) - new Date(a.いつ);', 'return new Date(a.いつ) - new Date(b.いつ);')],
    ['lib/rireki.js', '★黙って 切る（何件 出していないかを 言わない）★', (s) => s.replace('出していない: Math.max(0, 絞った.length - 出す.length),', '出していない: 0,')],
    ['lib/rireki.js', '★どう変わったかを 出さない★', (s) => s.replace("return x.場所 + (x.前 !== undefined ? '（' + 見せる(x.前) + ' → ' + 見せる(x.後) + '）' : '');", 'return x.場所;')],
    ['lib/rireki.js', '★AIを使っていない事を 言わない★', (s) => s.replace("else 出.push('AIは使っていません');", '')],
    ['lib/rireki.js', '★「1円もかかりません」と書く★', (s) => s.replace("else 出.push('AIは使っていません');", "else 出.push('1円もかかりません');")],
    ['book.html', '★組み立てる所を 2つにする（同じに見えなくなる）★',
      (s) => s.replace('function 履歴を描く(', 'function 履歴を描く2(){}\nfunction 履歴を描く(')],
    ['book.html', '★客のブックに 履歴のタブを足す★', (s) => s.replace('function 履歴を描く(', "var 履歴シート = '履歴シート';\nfunction 履歴を描く(")],
  ];
  let red = 0;
  for (const [rel, name, brk] of BREAKS) {
    const orig = fs.readFileSync(path.join(ROOT, rel), 'utf8');
    const bad = brk(orig);
    if (bad === orig) { console.log('  ★置換できず★  ' + name); continue; }
    const f = path.join(TMP, rel.replace(/[\\/]/g, '_'));
    fs.writeFileSync(f, bad, 'utf8');
    const env = Object.assign({}, process.env, { EXALLY_RIREKI_OVERRIDE: JSON.stringify({ [rel]: f }) });
    const r = spawnSync(process.execPath, ['--max-old-space-size=4096', path.join(__dirname, 'rireki.test.mjs')], { encoding: 'utf8', env });
    if (r.status !== 0) { red++; console.log('  赤くなった  ' + name); }
    else console.log('  ★素通り★  ' + name);
  }
  try { fs.rmSync(TMP, { recursive: true, force: true }); } catch (e) { /* 消せなくても検査は済んでいる */ }
  console.log('');
  console.log('  ' + red + '/' + BREAKS.length + ' 通りで赤くなった');
  process.exit(red === BREAKS.length ? 0 : 1);
}

process.exit(fail ? 1 : 0);
