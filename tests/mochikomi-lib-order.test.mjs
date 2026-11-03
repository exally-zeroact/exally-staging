/* mochikomi-lib-order.test.mjs
 *   -- ★読む 順が 2か所に 在るので ずれたら 赤★（2026-09-21）
 *
 *  ★★なぜ★★
 *    `book.html` の `_ensureXlsx()` ... ★表の 画面を 軽く する ため 後から 読む★
 *    `mochikomi.html`               ... ★開く 事しか しない ので 先に 全部 読む★
 *    ⇒★同じ 順が 2か所に 在ります★
 *    ⇒記憶「★作る道が 2本 在る時は 両方 直せ★」＝★片方だけ 直すと 黙って 壊れます★
 *      （`xlsx-io.js` は 読まれた 瞬間に XLSX を 掴むので ★SheetJS が 先★ ...）
 *
 *  ★★この 見張りが 見る 物★★
 *    ①`book.html` が 読む 台の ★並び★
 *    ②`mochikomi.html` の `<script>` の ★並び★
 *    ⇒①が ②の ★部分列★に なって いるか（順を 崩して いないか）
 *    ⇒①に 在るのに ②に 無い 物が 1つでも 在れば ★赤★
 *
 *  ★★見て いない 物★★
 *    ・★実際に 読み込めるか★（それは ブラウザで 立てる 試験の 仕事）
 *    ・`diff-preview.js`（★直す 画面だけの 物★＝この 画面は 直しません）
 *
 *  自己確認: `--self-test`（★わざと 順を 入れ替えて 赤に なるか★）
 */
import { strict as assert } from 'node:assert';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ここ = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(ここ, '..');
const よむ = (p) => fs.readFileSync(path.join(ROOT, p), 'utf8').split('\r\n').join('\n');

/** ★`book.html` の `_ensureXlsx()` が 読む 台を 並びの まま 取り出す★ */
export function 本の順(book) {
  const i = book.indexOf('function _ensureXlsx()');
  if (i < 0) return null;
  /* ★`js/book-open.js` を 読んだ 所で 終わり★（そこが 最後だと 書いて ある） */
  const j = book.indexOf("js/book-open.js", i);
  if (j < 0) return null;
  const 塊 = book.slice(i, j + 40);
  return [...塊.matchAll(/_loadScript\('([^']+)'/g)].map((m) => m[1]);
}

/** ★`mochikomi.html` の `<script src>` を 並びの まま 取り出す★（外の CDN は 除く） */
export function 画面の順(html) {
  return [...html.matchAll(/<script src="((?:lib|js)\/[^"?]+)/g)].map((m) => m[1]);
}

/** ★①が ②の 部分列か★（順を 崩して いないか） */
export function 部分列か(小, 大) {
  let k = 0;
  for (const x of 大) { if (x === 小[k]) k += 1; }
  return k === 小.length;
}

let 緑 = 0; const 赤 = [];
const 見る = (名, する) => {
  try { する(); 緑 += 1; console.log('  ok   ' + 名); }
  catch (e) { 赤.push(名); console.log('  ★赤★ ' + 名 + ' ... ' + e.message); }
};

/* ★★免除は 1本だけ＝★理由と 戻す 条件★を 書きます★★
     `lib/diff-preview.js` ... ★直す 前に「何を 書き込むか」を 見せる 台★
       ・`mochikomi.html` は ★直しません★ので 要りません。
       ・★見立てでは ありません★＝実物で 数えました（2026-09-21）
           `js/book-open.js` の `DiffPreview` ... ★0か所★
           呼んで いるのは ★`book.html` だけ★
       ・★戻す 条件★ ... `js/book-open.js` が `DiffPreview` を 呼ぶ ように なったら
                        ★この 免除を 外して mochikomi.html にも 足す★
     ⇒下の 門が その 条件を ★機械で★ 見ます（呼ぶ ように なったら 赤） */
const 免除 = ['lib/diff-preview.js'];

const book = よむ('book.html');
const mk = よむ('mochikomi.html');
const 本全部 = 本の順(book);
const 本 = 本全部 ? 本全部.filter((x) => 免除.indexOf(x) < 0) : null;
const 画 = 画面の順(mk);

console.log('[mochikomi-lib-order] ★読む 順が 2か所で ずれて いないか★');
console.log('  book.html ... ' + (本全部 ? 本全部.length : 0) + '本（免除 ' + 免除.length + '本を 除いて ' + (本 ? 本.length : 0) + '本）');
console.log('  mochikomi.html ... ' + 画.length + '本');

見る('★book.html から 順が 取れる（空振りして いない）★', () => {
  assert.ok(本, '`_ensureXlsx()` が 見つかりません');
  assert.ok(本.length >= 8, '取れた 数が 少なすぎます ... ' + 本.length);
  assert.equal(本[0], 'lib/xlsx.full.min.js', '★SheetJS が 先★');
});

見る('★book.html が 読む 台は 全部 mochikomi.html にも 在る★', () => {
  const 無い = 本.filter((x) => 画.indexOf(x) < 0);
  assert.deepEqual(無い, [], '★足りません★ ... ' + 無い.join(' / '));
});

見る('★★順を 崩して いない（部分列）★★', () => {
  assert.ok(部分列か(本, 画), '★並びが 違います★\n  book: ' + 本.join(' > ') + '\n  画面: ' + 画.join(' > '));
});

見る('★★免除の 訳が まだ 生きて いるか（★外して 測る★）★★', () => {
  /* ★免除の 訳「直す 画面だけの 物」は ★見立てでは 通しません★★
       ＝`js/book-open.js` が 呼び始めたら ★この 免除は 嘘に なります★ */
  const bo = よむ('js/book-open.js');
  const 数 = (bo.match(/DiffPreview/g) || []).length;
  assert.equal(数, 0,
    '★`js/book-open.js` が `DiffPreview` を ' + 数 + 'か所 呼んで います★'
    + '＝免除の 訳が 消えました。`mochikomi.html` にも `lib/diff-preview.js` を 足して ください。');
});

見る('★この 画面は 直さないので 打つ 口を 持たない★', () => {
  assert.equal(/\bcontenteditable\b/.test(mk), false, 'contenteditable が 在ります');
  assert.equal(/<input(?![^>]*type="file")/.test(mk), false, 'file 以外の 入力欄が 在ります');
});

見る('★★この 画面は 計算させない（★式の 台を 1つも 持たない★）★★', () => {
  /* ★★2026-09-21 ── ★これは 決めです★（Exally1 に 訊かれて 私が 決めました）★★
       ア② は「★持ち込んだ 物を 見て、そのまま 返す★」所です。
       ★値打ちは 「元の ファイルを 1バイトも 触らない」事★。
       ⇒★計算させると 答えが 書き換わる 恐れが 在ります★
         ＝借り物の 計算が 実Excel と 1マスでも 違えば、
           ★お客さんの ファイルの 答えが 黙って 変わって 返ります★
         ＝★それは この 画面の 値打ちを 壊します★
       ⇒★だから 計算の 口を 1つも 置きません★
       ★★直したい 人は `book.html` へ★★（そちらは 計算します＝★道を 分けるのが 役目★）
       ★戻す 条件★ ... この 画面で 直せる ように する と 決めた 時。
                      ★その時は 「1バイトも 触らない」を 先に 取り下げる 事★ */
  const 計算の口 = ['initFormulaEngine', 'loadSheetIntoEngine', 'recalcSheet', 'HyperFormula'];
  const 在る = 計算の口.filter((n) => mk.indexOf(n) >= 0 || よむ('js/mochikomi.js').indexOf(n) >= 0);
  assert.deepEqual(在る, [],
    '★計算の 口が 入りました★ ... ' + 在る.join(' / ')
    + '＝この 画面は 計算させない 決めです。直すなら 先に 上の 覚書きを 読んで ください。');
});

見る('★ログインの 扉が 在る★', () => {
  assert.ok(/id="app"[^>]*hidden/.test(mk), '`#app` が hidden で ありません');
  assert.ok(/#app\[hidden\]\{display:none/.test(mk), '`[hidden]` の 一行が 在りません');
  assert.ok(mk.indexOf('js/auth.js') > 0 && mk.indexOf('js/exally-login.js') > 0, 'ログインの 台が 在りません');
});

if (process.argv.includes('--self-test')) {
  console.log('');
  console.log('[mochikomi-lib-order --self-test] ★わざと 壊して 赤に なるか★');
  /* ★★2026-09-21 ── 1回目は ★実物を 1文字も 壊して いません★でした★★
       並びの 配列を 手で 作り替えて `部分列か` に 渡しただけ＝★台の 試験★で、
       ★`mochikomi.html` が 壊れても 気づけません★。
       （`tests/name-vs-body.test.mjs` が ★「壊して いないのに そう 名乗る」★と 赤に しました）
     ⇒★実物の 字を 書き換えた 写しで 同じ 検査を 走らせます★（ディスクは 触りません）
     ⇒`tests/hon-no-nakami.test.mjs` の 自己試験と ★同じ やり方★に 揃えました */
  /* ★逆斜線を 1文字も 使いません★（heredoc や 便りで ★落ちます★＝記憶の 決まり）
       ⇒正規表現を 使わず `indexOf` と `slice` で 切ります */
  function 一行を取る(h, 名) {
    const a = h.indexOf('<script src="' + 名);
    if (a < 0) throw new Error('見つかりません ... ' + 名);
    const b = h.indexOf('</script>', a) + 9;
    return h.slice(a, b);
  }
  const 壊し方 = [
    ['①`<script>` を 1本 抜く', (h) => {
      const 行 = 一行を取る(h, 'lib/table-refs.js');
      return h.split(行).join('');
    }],
    ['②SheetJS を 後ろへ 回す', (h) => {
      const 行 = 一行を取る(h, 'lib/xlsx.full.min.js');
      const 先 = 一行を取る(h, 'js/book-open.js');
      const 改 = String.fromCharCode(10);   /* ★逆斜線を 書かない★ */
      return h.split(行).join('').split(先).join(行 + 改 + 先);
    }],
    ['③2本の 前後を 入れ替える', (h) => {
      const a = 一行を取る(h, 'lib/xlsx-io.js');
      const b = 一行を取る(h, 'lib/zip-surgeon.js');
      return h.split(a).join('@A@').split(b).join(a).split('@A@').join(b);
    }],
  ];
  let 鳴 = 0;
  for (const [名, 壊す] of 壊し方) {
    /* ★先に 壊して いない 実物で 通る事を 見ます★（通らないなら 検査が 壊れて います） */
    assert.ok(部分列か(本, 画面の順(mk)), '壊す 前に 通りません＝検査が 壊れて います');
    const 壊れ = 壊す(mk);
    assert.notEqual(壊れ, mk, '★1文字も 壊れて いません★ ... ' + 名);
    const 画2 = 画面の順(壊れ);
    const 無い = 本.filter((x) => 画2.indexOf(x) < 0);
    const 赤に = (無い.length > 0) || !部分列か(本, 画2);
    console.log((赤に ? '  ok   ' : '  ★赤★ ') + 名
      + (赤に ? ' ... 壊したら 赤に なりました' : ' ... ★壊しても 緑の ままです★'));
    if (赤に) 鳴 += 1;
  }
  if (鳴 !== 壊し方.length) { console.log('★★自己確認が 通りません★★'); process.exit(1); }
}

console.log('');
console.log('mochikomi-lib-order: ' + 緑 + ' 緑 / ' + 赤.length + ' 赤');
if (赤.length) process.exit(1);
