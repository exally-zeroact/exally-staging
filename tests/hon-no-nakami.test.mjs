/* hon-no-nakami.test.mjs — ★持ち込んだ 本の 中身を 数える 台の 見張り★（2026-09-21）
 *
 *  ★★この 見張りが 一番 守りたい 物★★
 *    ★「マクロが 無い」と 「マクロを 読めなかった」を 混ぜない事★
 *    ＝読めなかった 物を 「0本」と 言うと ★お客さんは 「無い」と 読みます★
 *    ⇒ここが 崩れたら ★赤★に します。
 *
 *  ★自己確認★ `--self-test` ... ★わざと 壊して 赤に なるか★
 *    （★壊しても 緑の ままなら 見張りが 空振りして います★）
 */
import { strict as assert } from 'node:assert';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const ここ = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(ここ, '..');
const require_ = createRequire(path.join(ROOT, 'package.json'));
const H = require_(path.join(ROOT, 'lib/hon-no-nakami.js'));
const E = require_(path.join(ROOT, 'lib/excel-version.js'));

let 緑 = 0; const 赤 = [];
function 見る(名, する) {
  try { する(); 緑 += 1; console.log('  ok   ' + 名); }
  catch (e) { 赤.push(名 + ' ... ' + e.message); console.log('  ★赤★ ' + 名 + ' ... ' + e.message); }
}

const 板 = [{
  name: 'Sheet1',
  data: {
    '0,0': { f: '=SUM(A1:A3)' },
    '1,0': { f: '=SUM(B1)+MAX(C1)' },
    '2,0': { v: 3 },
    '3,0': { f: '=IF(A1>0,"SUM(x)","")' },   /* ★字の 中の SUM( は 数えない★ */
    /* ★★`=` で 始まらない `f` は 式では ありません★★
         ＝お客さんが `SUM(A1)` と ★字として★ 打った マス。
         ★これが 無いと 「`=` で 始まるか」の 門を わざと 壊しても 数が 動かず、
           自己確認が ★口先だけの ok★に なります★（2026-09-21 に 踏みました） */
    '4,0': { f: 'SUM(A1)' },
  },
}];

console.log('[hon-no-nakami] ★持ち込んだ 本の 中身を 数える★');

見る('★関数を 数える（字の 中は 数えない）★', () => {
  const d = H.本の中身を数える({ 形: 'xlsx', sheets: 板, 関数を拾う: E.functionsIn, マクロの読み: null });
  assert.equal(d.関数.測った, true);
  const 表 = {};
  d.関数.一覧.forEach((x) => { 表[x.名] = x.数; });
  assert.equal(表.SUM, 2, 'SUM は 2回（字の 中の SUM( は 数えない）');
  assert.equal(表.MAX, 1);
  assert.equal(表.IF, 1);
  assert.equal(d.関数.式の在るマス, 3);
});

見る('★★マクロ「無い」と「読めない」を 混ぜない★★', () => {
  const 無 = H.本の中身を数える({ sheets: 板, 関数を拾う: E.functionsIn, マクロの読み: null });
  const 読 = H.本の中身を数える({
    sheets: 板, 関数を拾う: E.functionsIn,
    マクロの読み: { ok: false, モジュール: [], なぜ: '入れ物が 読めませんでした' },
  });
  assert.equal(無.マクロ.ありさま, '無し');
  assert.equal(読.マクロ.ありさま, '読めません');
  assert.notEqual(無.マクロ.ありさま, 読.マクロ.ありさま, '★同じ 言い方に しない★');
  /* ★字でも 混ざって いない事★＝お客さんが 読むのは 字 */
  assert.match(H.一文(無), /マクロは ありません/);
  assert.match(H.一文(読), /読めませんでした/);
  assert.doesNotMatch(H.一文(読), /ありません/, '★読めない物を「ありません」と 言わない★');
});

見る('★関数を 読む 台を 渡さなければ「数えていません」と 言う★', () => {
  const d = H.本の中身を数える({ sheets: 板, マクロの読み: null });          /* ★拾う を 渡さない★ */
  assert.equal(d.関数.測った, false);
  assert.equal(d.関数.種類, 0);
  assert.match(H.一文(d), /数えていません/);
  assert.doesNotMatch(H.一文(d), /1つも ありません/, '★未測定を 0件と 言わない★');
});

見る('★式が 1つも 無い 本★', () => {
  const d = H.本の中身を数える({ sheets: [{ name: 'S', data: { '0,0': { v: 1 } } }], 関数を拾う: E.functionsIn, マクロの読み: null });
  assert.equal(d.関数.測った, true);
  assert.match(H.一文(d), /式は 1つも ありません/);
});

見る('★マクロが 読めた 時は 本数と 手続きを 出す★', () => {
  const d = H.本の中身を数える({
    sheets: 板, 関数を拾う: E.functionsIn,
    マクロの読み: { ok: true, モジュール: [{}, {}], なぜ: '' },
    マクロの見立て: { 本数: 1, 手続き: [{ 名: 'Macro1', 分類: ['写す'], 行数: 12 }] },
  });
  assert.equal(d.マクロ.ありさま, '在ります');
  assert.equal(d.マクロ.本数, 2);
  /* ★★2026-09-22 ── ★同じ 窓に 4 と 5 が 並んで いました★★（Exally1 が 実物で 見つけた）
       1文   ... `読み.モジュール` の 数 ＝ ★かたまり（モジュール）★
       詳しく ... `見立て.手続き` の 数   ＝ ★手続き（Sub / Function）★
     ⇒どちらも 正しい 数。★言葉が 揃って いなかった★
     ⇒★両方 書く★形に した（片方を 消すと 数が 減ったように 見える） */
  assert.match(H.一文(d), /マクロが 2かたまり（手続き 1本）/);
  const く = H.詳しく(d).find((x) => x.見出し.indexOf('マクロ') === 0);
  assert.equal(く.見出し, 'マクロの 手続き');
  assert.equal(く.行.length, 1);
  assert.match(く.行[0], /Macro1/);
});

見る('★★見立てを 渡さなければ 「数えていません」と 言う（0件と 言い切らない）★★', () => {
  /* ★★2026-09-22★★ ＝★手続きが 無い★ と ★手続きを 数えて いない★ は 別
       ＝`0件` だけ 出すと お客さんは「マクロは 空」と 読みます */
  const d = H.本の中身を数える({
    sheets: 板, 関数を拾う: E.functionsIn,
    マクロの読み: { ok: true, モジュール: [{}, {}, {}, {}], なぜ: '' },
    /* ★見立てを 渡しません★ */
  });
  assert.match(H.一文(d), /マクロが 4かたまり/);
  assert.doesNotMatch(H.一文(d), /手続き/, '★数えて いないのに 手続きの 数を 書かない★');
  const く = H.詳しく(d).find((x) => x.見出し.indexOf('マクロ') === 0);
  assert.match(く.見出し, /数えていません/);
  assert.match(く.注, /手続きに 切って いません/);
});

見る('★板の 数は 渡した 板の 数★', () => {
  const d = H.本の中身を数える({ sheets: [{ name: 'a', data: {} }, { name: 'b', data: {} }], 関数を拾う: E.functionsIn, マクロの読み: null });
  assert.equal(d.板, 2);
  assert.match(H.一文(d), /板が 2枚/);
});

/* ══ ★★自己確認＝わざと 壊して 赤に なるか★★ ══
     ★★2026-09-21 ── 1回 目は ★口先だけの ok★を 書いて いました★★
       外から `H.関数を数える` を 差し替えても、`本の中身を数える` は ★台の 中の 物★を 呼ぶので
       ★1文字も 壊れて いないのに ok と 出て いました★
     ⇒★台の 字そのものを 書き換えた 写しを 作って 走らせます★（ディスクは 触りません） */
if (process.argv.includes('--self-test')) {
  console.log('');
  console.log('[hon-no-nakami --self-test] ★わざと 壊して 赤に なるか★');
  /* ★逆斜線を 使いません★＝heredoc でも 便りでも 落ちる（記憶の 決まり）
     ⇒改行を 揃えるのは `split` と `join` で やります */
  const CR = String.fromCharCode(13), LF = String.fromCharCode(10);
  const 元の字 = fs.readFileSync(path.join(ROOT, 'lib/hon-no-nakami.js'), 'utf8')
    .split(CR + LF).join(LF);

  /** ★字を 1か所 書き換えた 写しを 動かす★（`module.exports` を 受け取る） */
  function 壊した台(古, 新) {
    if (元の字.indexOf(古) < 0) throw new Error('壊す 先が 見つかりません ... ' + 古);
    const 字 = 元の字.replace(古, 新);
    const m = { exports: {} };
    // eslint-disable-next-line no-new-func
    new Function('module', 'exports', 字)(m, m.exports);
    return m.exports;
  }

  const 壊し = [
    ['①「読めません」を「ありません」と 言わせる',
      "出.push('マクロは 入っていますが 読めませんでした');",
      "出.push('マクロは ありません');",
      (X) => {
        const d = X.本の中身を数える({ sheets: 板, 関数を拾う: E.functionsIn, マクロの読み: { ok: false, モジュール: [], なぜ: 'x' } });
        assert.doesNotMatch(X.一文(d), /ありません/);
      }],
    ['②未測定を「0件」に 化けさせる',
      "return { 測った: false, 種類: 0, 数: 0, 式の在るマス: 0, 一覧: [] };",
      "return { 測った: true, 種類: 0, 数: 0, 式の在るマス: 0, 一覧: [] };",
      (X) => {
        const d = X.本の中身を数える({ sheets: 板, マクロの読み: null });   /* ★拾う を 渡さない★ */
        assert.equal(d.関数.測った, false);
        assert.match(X.一文(d), /数えていません/);
      }],
    ['③字の 中の 関数も 数えさせる',
      "if (!c || typeof c.f !== 'string' || c.f.charAt(0) !== '=') return;",
      "if (!c || typeof c.f !== 'string') return;",
      (X) => {
        const d = X.本の中身を数える({ sheets: 板, 関数を拾う: E.functionsIn, マクロの読み: null });
        assert.equal(d.関数.式の在るマス, 3);
      }],
  ];

  let 鳴った = 0;
  for (const [名, 古, 新, 確かめ] of 壊し) {
    /* ★先に ★壊して いない 台★で 通る事を 見ます★
         ＝通らないなら ★確かめ そのものが 壊れて います★ */
    確かめ(H);
    let 赤に = false;
    try { 確かめ(壊した台(古, 新)); } catch { 赤に = true; }
    console.log((赤に ? '  ok   ' : '  ★赤★ ') + 名
      + (赤に ? ' ... 壊したら 赤に なりました' : ' ... ★壊しても 緑の ままです★'));
    if (赤に) 鳴った += 1;
  }
  if (鳴った !== 壊し.length) {
    console.log('★★自己確認が 通りません＝見張りが 空振りして います★★');
    process.exit(1);
  }
}

console.log('');
console.log('hon-no-nakami: ' + 緑 + ' 緑 / ' + 赤.length + ' 赤');
if (赤.length) process.exit(1);
