/* header-footer.test.mjs — ★紙の 上と 下に 入れる 字★ 2026-08-31
 *
 *  ★司さんの 指示（2026-08-31）★
 *    「ヘッダーフッターは ★ごちゃごちゃに ならないよう
 *      ドロップダウンでも いいから 綺麗に★ する」
 *
 *  ★★印の 意味は 実Excel に 刷らせて 測った★★
 *    新規の 空ブックに 120行 入れ、ヘッダー/フッターに 印を 入れて
 *    ★Excel 自身に PDF を 書き出させ★、中の 字を 読んだ（2026-08-31）。
 *      &F→ファイル名（拡張子なし） / &P→1 / &N→4 / &A→シート名
 *      &D→2026/8/31 / &T→0:44 / &Z→置き場所 / &&→「&」
 *    入れ物の 形も 実測＝<headerFooter><oddHeader>&L…&C…&R…</oddHeader>…
 *
 *  ★紙の 上で 測った事（実ブラウザ・2026-08-31）★
 *    ・position:fixed は ★毎ページ 出る★（全ページ 上端 y=11 で 一定）
 *    ・★CSS の counter(page) は 増えない★（10ページ目も「1」のまま）
 *    ⇒ だから ★自分で ページを 割って 1枚ずつ 番号を 書く★。
 *    実測＝120行を A4縦で 刷ると ★3ページ★、下に「ページ 1 / 3」「2 / 3」「3 / 3」。
 *
 *  ★この 見張りが 見る物★
 *    ① 印の 表が 在る（実測した 8個）
 *    ② 組み立て／解く／差し込み が 逆向きに 通る
 *    ③ ★&& の 罠★＝「&&P」は 「&P」であって ページ番号では ない
 *    ④ ページ番号だけ 後回しに 出来る（何ページに なるか 後で 分かる為）
 *    ⑤ 画面（ドロップダウン＋自分で書く）が 在る
 *    ⑥ 刷る 側が ★自分で ページを 割っている★（counter に 頼っていない）
 *
 *  走らせ方: node tests/header-footer.test.mjs  ／  --self-test
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createRequire } from 'node:module';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(import.meta.url);
let ok = 0, ng = 0;
const 言う = (よい, 文, 添え) => {
  if (よい) { ok++; console.log('  ok   ' + 文); }
  else { ng++; console.log('  NG   ' + 文); if (添え) console.log('       ' + 添え); }
};

const HF = require_(path.join(ROOT, 'lib/header-footer.js'));
const GP = require_(path.join(ROOT, 'lib/grid-print.js'));
const book = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf8');
const { 注記を外す } = await import(
  pathToFileURL(path.join(ROOT, 'scripts/lib/chuki.mjs')).href);
const 素 = 注記を外す(book);

console.log('★紙の 上と 下に 入れる 字★');

/* ── ① 印の 表 ── */
const 実測 = { '&P': 'ページ番号', '&N': 'ページ数', '&D': '日付', '&T': '時刻',
  '&F': 'ファイル名', '&A': 'シート名', '&Z': '置き場所', '&&': '「&」そのもの' };
言う(HF.印.length === 8, '★印は 8個（実Excel に 刷らせて 測った 物）★（今 ' + HF.印.length + '個）');
for (const k of Object.keys(実測)) {
  const v = HF.印.find((x) => x.印 === k);
  言う(!!v && v.名 === 実測[k], '　' + k + ' ＝ ' + 実測[k], v ? ('今 ' + v.名) : '無い');
}
言う(HF.印.every((v) => v.例 !== undefined && v.例 !== ''),
  '★どの 印にも 実測の 例が 付いている★');

/* ── ② 組み立て／解く ── */
const 元 = { 左: '&F', 中: 'C=&P / &N', 右: '&A' };
const 字 = HF.組み立てる(元);
言う(字 === '&L&F&CC=&P / &N&R&A', '★組み立ては 実Excel と 同じ形（&L…&C…&R…）★', '今 ' + 字);
const 戻 = HF.解く(字);
言う(戻.左 === 元.左 && 戻.中 === 元.中 && 戻.右 === 元.右,
  '★解いたら 元に 戻る★', JSON.stringify(戻));
言う(HF.組み立てる({ 左: '', 中: 'あ', 右: '' }) === '&Cあ',
  '★空の 所は 書かない（Excelも 書かない）★');

/* ── ③ && の 罠 ── */
言う(HF.差し込む('A&&P B', { ページ: 9 }) === 'A&P B',
  '★★「&&P」は 「&P」＝ページ番号では ない★★',
  '★ここを 間違えると 客の 字が 黙って 数字に 化ける★');
const 罠 = HF.解く('A&&L B&CC');
言う(罠.左 === 'A&&L B' && 罠.中 === 'C',
  '★「&&L」を 区切りと 読み違えない★', JSON.stringify(罠));

/* ── ④ ページ番号だけ 後回し ── */
const 後 = HF.頁を後回しで差し込む('&A ページ &P / &N', { シート名: '売上' });
言う(後.indexOf('売上') === 0, '★他の 印は 先に 本物の 字に なる★', 後);
言う(後.indexOf(HF.頁印) > 0 && 後.indexOf(HF.総印) > 0,
  '★&P と &N は 残る（何ページに なるかは 後で 分かる）★', 後);
言う(HF.頁を入れる(後, 2, 5) === '売上 ページ 2 / 5',
  '★後から 本物の 数字を 入れられる★', HF.頁を入れる(後, 2, 5));
言う(HF.使っているか({ 左: '', 中: '', 右: '' }) === false
  && HF.使っているか({ 左: '', 中: 'あ', 右: '' }) === true,
  '★1つも 書いていなければ 使っていないと 判る（紙の 場所を 取らない）★');

/* ── ⑤ 画面 ── */
言う(/header-footer\.js/.test(素), '★画面が 部品を 読み込んでいる★');
for (const id of ['hfHead', 'hfFoot', 'hfSelf', 'hfHL', 'hfHC', 'hfHR', 'hfFL', 'hfFC', 'hfFR',
  'hfPreviewH', 'hfPreviewF']) {
  言う(new RegExp('id="' + id + '"').test(book), '　欄 ' + id + ' が 在る');
}
言う(HF.よく使う形.length >= 8,
  '★選ぶだけで 済む 形が ' + HF.よく使う形.length + '個 在る（司さん＝ドロップダウンで 綺麗に）★');
言う(HF.よく使う形[0].左 === '' && HF.よく使う形[0].中 === '' && HF.よく使う形[0].右 === '',
  '★1つ目は「（なし）」★（既定は 何も 出さない＝実Excel と 同じ）');
言う(/ヘッダーの窓を詰める\(\)/.test(素) && /ヘッダーを取り込む\(\)/.test(素),
  '★開く時に 詰めて、決める時に 取り込んでいる★',
  '★どちらか 抜けると 書いた 字が 黙って 消える★');
言う(/一覧に 無い/.test(book) || /自分で = \(hi < 0 \|\| fi < 0\)/.test(素),
  '★一覧に 無い 中身なら 自動で「自分で 書く」に する★',
  '★でないと 書いた 字が 黙って 消える★');

/* ── ⑥ 刷る 側 ── */
const gp = fs.readFileSync(path.join(ROOT, 'lib/grid-print.js'), 'utf8');
const gp素 = 注記を外す(gp);
言う(/function 上下の字\(/.test(gp素), '★刷る 側に 上下の字が 在る★');
言う(/function 割り付ける\(/.test(gp素), '★自分で ページを 割っている★');
言う(!/counter\(page\)/.test(gp素),
  '★CSS の ページ番号に 頼っていない★',
  '★実測＝Chrome では 増えない（10ページ目も「1」）★');
言う(/getBoundingClientRect\(\)\.height/.test(gp素),
  '★行の 高さは 出してから 測っている（決め打ちしない）★');

/* 何も 書いていなければ 何も 足さない */
/* ★見るのは「割り付けを 入れたか」★（pgbox は 見た目の 決まりに 常に 在る＝印に ならない） */
const 空 = GP.buildHtml({ data: { '0,0': { v: 'あ' } }, sheetName: 'S' });
言う(空 && 空.indexOf('割り付ける') < 0, '★上下に 何も 書いていない時は 割り付けない★');
const 有 = GP.buildHtml({ data: { '0,0': { v: 'あ' } }, sheetName: 'S',
  ヘッダー: { 左: '', 中: 'ページ P', 右: '' }, 頁印: 'P', 総印: 'N' });
言う(有 && 有.indexOf('割り付ける') > 0 && 有.indexOf('pgbox') > 0,
  '★書いてあれば 割り付けを 入れる★');

/* ── わざと 壊して 赤に なるか ── */
if (process.argv.includes('--self-test')) {
  console.log('\n★わざと 壊して 赤に なるか★');
  言う(HF.差し込む('&P', {}) === '1', '★何も 渡さなければ 1ページ目★');
  言う(HF.差し込む('&Q', { ページ: 3 }) === '&Q', '★知らない 印は そのまま 残す（勝手に 消さない）★');
  言う(HF.解く('').左 === '', '★空でも 落ちない★');
  言う(HF.頁を入れる('', 1, 1) === '', '★空でも 落ちない（番号入れ）★');
  const 逆 = HF.解く(HF.組み立てる({ 左: 'a&&b', 中: '', 右: 'c' }));
  言う(逆.左 === 'a&&b' && 逆.右 === 'c', '★&& が 入っていても 往復できる★', JSON.stringify(逆));
  言う(GP.buildHtml({ data: {}, sheetName: 'S' }) === null, '★中身が 0なら 刷らない（前からの 決まり）★');

  /* ★★本当に 壊して 赤に なるかを 見る★★
     ここまでは「端の 値を 通す」だけ＝★壊していない★。
     ★見出しで「壊す」と 名乗る 以上 本当に 壊す★（tests/name-vs-body.test.mjs の 決まり）。
     ★repo は 読むだけ★＝字の 上で 壊して 数え直す。 */
  const 元字 = fs.readFileSync(path.join(ROOT, 'lib/header-footer.js'), 'utf8');
  const 壊した = 元字.replace("{ 印: '&P', 名: 'ページ番号',   例: '1' },", '');
  言う(壊した !== 元字, '★壊せた（&P の 行を 抜いた）★');
  const 数え = (t) => (t.match(/\{ 印: '&/g) || []).length;
  言う(数え(元字) === 8 && 数え(壊した) === 7,
    '★1行 抜いたら 8→7 に なる（数える所が 効いている）★',
    '本物 ' + 数え(元字) + ' / 壊した ' + 数え(壊した));
  const 紙壊し = fs.readFileSync(path.join(ROOT, 'lib/grid-print.js'), 'utf8')
    .replace('function 割り付ける(', 'function nope(');
  言う(!/function 割り付ける\(/.test(紙壊し),
    '★ページを 割る所を 消したら ⑥が 落ちる★');
  言う(fs.readFileSync(path.join(ROOT, 'lib/header-footer.js'), 'utf8') === 元字,
    '★★本物の ファイルは 1バイトも 触っていない★★');
}

console.log('\nheader-footer: ' + ok + '/' + (ok + ng) + ' passed');
process.exit(ng ? 1 : 0);
