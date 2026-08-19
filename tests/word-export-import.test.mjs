/* word-export-import.test.mjs — ★言い方を「書き出す ↔ 読み込む」に固定する★
 *
 * なぜ必要か（2026-08-18 指示役の裁定）
 *   同じ動きに ★2通りの名前★ が付くと、客は「別の機能か？」と思う。
 *   実際に踏んだ: 窓の中は「書き出す」に直したのに、上のボタンは「Excelに保存」のままで、
 *   ★同じ画面に「保存」と「書き出す」が並んだ★。
 *   今日 全アプリで言い方を揃えた（飲み屋が決めた形）:
 *     ★ファイルを受け取る = 読み込む ／ ファイルを渡す = 書き出す★
 *     「保存する」「落とす」「Excelにする」は ★2通り目の言い方★なので使わない。
 *
 * ★どこを見るか＝客の目に入る所だけ（注記は見ない）★
 *   「前は『Excelに保存』だった」という ★経緯の記録★まで書き換えさせると、
 *   何があったか分からなくなる。★直す対象は 客に出る字★。
 *
 * ★「保存」を全部 禁止しない★
 *   倉庫にデータを入れるのは「保存」で正しい（js/hub.js の「保存しました」等）。
 *   ★禁じるのは「Excel＋保存/開く/にする」の組み合わせだけ★。
 *
 * 使い方: node tests/word-export-import.test.mjs [--self-test]
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');

let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ✓ ' + n); } catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + (e && e.message)); } };

/** JS の注記だけを落とす。
 *  ★文字列を1文字ずつ追う作りは やめた★（2026-08-18 実際に踏んだ）:
 *    正規表現リテラル（例 スラッシュ + ['"] + スラッシュ）の中の ' を「文字列の始まり」と
 *    誤解し、その後のブロックの注記を落とし損ねて ★1件 残り続けた★。
 *  ★今の作り★
 *    ・ブロックの注記は そのまま落とす（文字列の中にブロックの始まりが入る事はまず無い）
 *    ・行の注記は ★その行の手前の引用符が偶数の時だけ★ 落とす
 *      （'https://…' のスラッシュ2つを注記と間違えないため） */
export function stripJsComments(src) {
  const noBlock = String(src).replace(/\/\*[\s\S]*?\*\//g, ' ');
  return noBlock.split('\n').map((line) => {
    let i = line.indexOf('//');
    while (i >= 0) {
      const head = line.slice(0, i);
      const quotes = (head.match(/['"`]/g) || []).length;
      if (quotes % 2 === 0) return head;        // 引用符が閉じている＝ここからは注記
      i = line.indexOf('//', i + 2);            // 文字列の中の // は注記ではない
    }
    return line;
  }).join('\n');
}

/** 注記を落として「客に出る字」だけにする。
 *  ★HTMLの中とJSの中を分けて処理する★
 *    一緒に走査すると、HTMLの文の中の ' や " で文字列の追跡がズレて
 *    ★JSの注記を落とし損ねる★（2026-08-18 実際にそうなった）。 */
export function stripComments(src) {
  let s = String(src).replace(/<!--[\s\S]*?-->/g, ' ');           // ①HTMLの注記は形がはっきりしている
  if (s.indexOf('<script') < 0 && s.indexOf('<style') < 0) return stripJsComments(s);  // .js / .css はそのまま
  /* ②<style> の中の注記（CSSの注記）も落とす。ここに経緯を書いてある。
        ★2026-08-18 実際に踏んだ: style を見ていなくて 直す所が1件 残り続けた★ */
  s = s.replace(/(<style[^>]*>)([\s\S]*?)(<\/style>)/gi,
    (m, a, body, b) => a + body.replace(/\/\*[\s\S]*?\*\//g, ' ') + b);
  /* ③<script> の中は JS として見る */
  return s.replace(/(<script[^>]*>)([\s\S]*?)(<\/script>)/gi,
    (m, a, body, b) => a + stripJsComments(body) + b);
}

/* ★禁じる言い方（Excel と組み合わせた時だけ）★ */
const BAD = [
  ['Excelに保存', 'Excelに書き出す'],
  ['Excelを開く', 'Excelを読み込む'],
  ['Excelにする', 'Excelに書き出す'],
  ['Excelファイルを保存', 'Excelファイルを書き出す'],
  ['Excelファイルを開く', 'Excelファイルを読み込む'],
  ['Excelで落と', 'Excelに書き出す'],
];

/** 配る物（HTML と js/ lib/ の中身）を全部 数える */
function shipped() {
  const out = [];
  for (const f of fs.readdirSync(ROOT)) if (/\.html$/i.test(f)) out.push(f);
  for (const d of ['js', 'lib']) {
    const p = path.join(ROOT, d);
    if (!fs.existsSync(p)) continue;
    for (const f of fs.readdirSync(p)) if (/\.(js|css)$/i.test(f)) out.push(d + '/' + f);
  }
  return out.sort();
}

/* ══ self-test ══════════════════════════════════════════════════════ */
if (process.argv.includes('--self-test')) {
  console.log('\n[word-export-import --self-test] 判定そのものが空振りしていないか');
  T('★ボタンの字は捕まえる', () => {
    if (stripComments('<button>Excelに保存</button>').indexOf('Excelに保存') < 0) throw new Error('落としてしまった');
  });
  T('★HTMLの注記の中は見ない（経緯の記録を消させない）', () => {
    const s = stripComments('<!-- 前は「Excelに保存」だった -->\n<button>Excelに書き出す</button>');
    if (s.indexOf('Excelに保存') >= 0) throw new Error('注記まで見ている');
    if (s.indexOf('Excelに書き出す') < 0) throw new Error('客に出る字を落とした');
  });
  T('★// の注記も見ない', () => {
    const s = stripComments("var a=1; // 昔は Excelを開く だった\nvar b='Excelを読み込む';");
    if (s.indexOf('Excelを開く') >= 0) throw new Error('注記まで見ている');
    if (s.indexOf('Excelを読み込む') < 0) throw new Error('文字列を落とした');
  });
  T('★/* */ の注記も見ない', () => {
    const s = stripComments('/* Excelに保存 の話 */ var t="Excelに書き出す";');
    if (s.indexOf('Excelに保存') >= 0) throw new Error('注記まで見ている');
  });
  T('★URLの // を注記と間違えない', () => {
    if (stripComments("var u='https://example.com/a';").indexOf('https://example.com/a') < 0) throw new Error('URLを壊した');
  });
  T('★倉庫の「保存しました」は禁じない（Excelと組み合わさっていない）', () => {
    const s = stripComments("toast('保存しました');");
    for (const [w] of BAD) if (s.indexOf(w) >= 0) throw new Error('関係ない「保存」を捕まえた: ' + w);
  });
  T('★HTMLの文とJSが混ざっていても JSの注記を落とす（2026-08-18 実際に踏んだ）', () => {
    const src = "<p>これは「テスト」です。使えない'物'もある</p>\n<script>\n/* ===== Excelに保存 ==== */\nvar t=\"Excelに書き出す\";\n</script>";
    const s = stripComments(src);
    if (s.indexOf('Excelに保存') >= 0) throw new Error('JSの注記を落とせていない');
    if (s.indexOf('Excelに書き出す') < 0) throw new Error('客に出る字を落とした');
  });
  T('★<style> の中の注記も見ない（2026-08-18 実際に踏んだ）', () => {
    const src = '<style>\n/* 「Excelを開く」を足したら ロゴが切れた */\n.a{color:#333}\n</style><button>Excelを読み込む</button>';
    const s = stripComments(src);
    if (s.indexOf('Excelを開く') >= 0) throw new Error('CSSの注記まで見ている');
    if (s.indexOf('Excelを読み込む') < 0) throw new Error('客に出る字を落とした');
  });
  T('★言葉をタグで割っても 人が読む字として見つけられる（2026-08-19 実際に踏んだ）', () => {
    const 剥がす = (h) => h.replace(/<script[\s\S]*?<\/script>/g, ' ').replace(/<style[\s\S]*?<\/style>/g, ' ').replace(/<[^>]+>/g, '');
    const 割れた = '<button><span class="hdr-ic">📂</span> <span class="hdr-lb"><span class="hdr-lb-long">Excelを</span>読み込む</span></button>';
    if (剥がす(割れた).indexOf('Excelを読み込む') < 0) throw new Error('タグを外しても繋がらない＝実際の画面を見ていない');
    /* ★本当に字が無くなったら見つからない事も確かめる（何でも通す判定にしない）★ */
    const 絵だけ = '<button><span class="hdr-ic">📂</span></button>';
    if (剥がす(絵だけ).indexOf('Excelを読み込む') >= 0) throw new Error('字が無いのに見つけた事になっている');
  });
  T('★禁じる言い方の一覧が空になっていない', () => { if (BAD.length < 3) throw new Error('一覧が痩せている'); });
  console.log('\n  ── 実測 ── 確かめた ' + (pass + fail) + ' 通り / 通った ' + pass + ' 通り');
  console.log('\n' + pass + ' passed, ' + fail + ' failed');
  process.exit(fail ? 1 : 0);
}

/* ══ 本番 ═══════════════════════════════════════════════════════════ */
console.log('\n[word-export-import] 言い方は「書き出す ↔ 読み込む」で揃っているか');

const files = shipped();
const hits = [];
for (const f of files) {
  const src = stripComments(fs.readFileSync(path.join(ROOT, f), 'utf8'));
  for (const [bad, good] of BAD) {
    let i = src.indexOf(bad);
    while (i >= 0) {
      hits.push(f + ' … 「' + bad + '」→ ★「' + good + '」にする★');
      i = src.indexOf(bad, i + bad.length);
    }
  }
}
console.log('      ── 実測 ── 配る物 ' + files.length + '本を数えて、直す所 ' + hits.length + '件');
T('★配る物に「Excelに保存 / Excelを開く / Excelにする」が残っていない', () => {
  if (hits.length) throw new Error('残っている:\n      ' + hits.join('\n      '));
});
T('★数えた物が0本になっていない（検査が空振りしていない）', () => {
  if (files.length < 5) throw new Error('配る物が ' + files.length + '本しか見えていない');
});
/* ★正しい言い方が実際に使われている事も見る（消しただけで言い換えていない、を防ぐ） */
T('★「Excelを読み込む」「Excelに書き出す」が実際に画面に出ている', () => {
  /* ★タグを外してから探す★（2026-08-19）
     スマホで字が消えていたのを直した時、言葉を
       <span class="hdr-lb-long">Excelを</span>読み込む
     のように ★タグで割った★。人が読む字は同じでも、生のHTMLを探すと見つからない。
     ⇒ ★人が読む字＝タグを外した後の字★ で見る。 */
  const 読める字 = fs.readdirSync(ROOT).filter(f => /\.html$/i.test(f))
    .map(f => stripComments(fs.readFileSync(path.join(ROOT, f), 'utf8')))
    .join('\n')
    .replace(/<script[\s\S]*?<\/script>/g, ' ')   // 中の言葉は画面の字ではない
    .replace(/<style[\s\S]*?<\/style>/g, ' ')
    .replace(/<[^>]+>/g, '');                     // タグを外す＝割れた言葉が繋がる
  if (読める字.indexOf('Excelを読み込む') < 0) throw new Error('「Excelを読み込む」が画面に無い');
  if (読める字.indexOf('Excelに書き出す') < 0) throw new Error('「Excelに書き出す」が画面に無い');
});

console.log('\n' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
