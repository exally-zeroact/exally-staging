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
import { 注記を外す } from '../scripts/lib/chuki.mjs';

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
  /* ★注記外しは 共通部品へ（2026-08-26 指示役）★ */
  return 注記を外す(String(src));
}

/** 注記を落として「客に出る字」だけにする。
 *  ★HTMLの中とJSの中を分けて処理する★
 *    一緒に走査すると、HTMLの文の中の ' や " で文字列の追跡がズレて
 *    ★JSの注記を落とし損ねる★（2026-08-18 実際にそうなった）。 */
export function stripComments(src) {
  /* ★ここに在った「HTMLの中とJSの中を分ける」決まりが 一番よく出来ていたので
     ★それを 共通部品(html:true)の正にして 全部の見張りへ 配った★（2026-08-26 指示役）。
     ・<!-- --> を先に外す ・<style> は ブロック注記だけ ・<script> は JS として見る
     ・<script も <style も無い字（.js/.css）は JS として見る */
  return 注記を外す(String(src), { html: true });
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

/** 注記の中に残った古い言い方を 台帳と突き合わせる（★純関数＝自己確認できる形にする★）
 *  ・載っていない … 黙って残っている＝赤にする物
 *  ・使われていない … 台帳が古い＝赤にする物（消したのに載せっぱなし） */
export function 台帳を照らす(注記の中, 台帳) {
  const key = (c) => c.file + '|' + c.bad;
  return {
    載っていない: 注記の中.filter((c) => !台帳[key(c)]),
    使われていない: Object.keys(台帳).filter((k) => !注記の中.some((c) => key(c) === k)),
  };
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
  T('★注記に残った古い言い方が 台帳に無ければ 拾う（2026-08-21 指示役と数が食い違った）', () => {
    const r = 台帳を照らす([{ file: 'a.js', bad: 'Excelに保存' }], {});
    if (r.載っていない.length !== 1) throw new Error('黙って残っている物を拾えていない');
  });
  T('★台帳に載っていれば 通す／消えた物が台帳に残っていたら 拾う', () => {
    const 台帳 = { 'a.js|Excelに保存': '事故の記録' };
    const r1 = 台帳を照らす([{ file: 'a.js', bad: 'Excelに保存' }], 台帳);
    if (r1.載っていない.length !== 0 || r1.使われていない.length !== 0) throw new Error('載っているのに赤にした');
    const r2 = 台帳を照らす([], 台帳);
    if (r2.使われていない.length !== 1) throw new Error('★台帳が古くなっても気づかない★');
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
/* ★生のまま（注記も込み）でも数える★（2026-08-21 指示役と数が食い違った）
   私は「27本を数えて0件」と出したが、指示役が ★配信の book.html を そのまま読んで3件★ 見つけた。
   食い違いの正体は ★私が注記を外して数えていたのに、そう書かなかった事★。
   ⇒ ★両方の数を出す★＝「画面に出る字 X件 ／ 注記の中 Y件」。
     注記の中に残してよい物は 下の台帳に ★理由つきで★ 載せる。載っていない物が在れば赤。 */
const 注記の台帳 = {
  'book.html|Excelを開く': '2026-08-09 に「Excelを開く」という名前でボタンを足したら'
    + ' ロゴが切れて 365 の切替が画面の外へ出た、という ★事故の記録★。'
    + '今の言い方（Excelを読み込む）に書き換えると 何が起きたか分からなくなるので そのまま残す。',
  'lib/diff-preview.js|Excelにする': '★禁じる言い方そのものを 注記で挙げている★'
    + '（「保存する／落とす／Excelにする」は2通り目の言い方になるので使わない、という決まりの説明）。'
    + 'ここを書き換えると 何を禁じているのか分からなくなるので そのまま残す。',
};
const 注記の中 = [];
for (const f of files) {
  const raw = fs.readFileSync(path.join(ROOT, f), 'utf8');
  const 客に出る字 = stripComments(raw);
  for (const [bad] of BAD) {
    let i = raw.indexOf(bad);
    while (i >= 0) {
      if (客に出る字.indexOf(bad) < 0) 注記の中.push({ file: f, bad: bad });
      i = raw.indexOf(bad, i + bad.length);
    }
  }
}
console.log('      ── 実測 ── 配る物 ' + files.length + '本を数えて'
  + '、★画面に出る字 ' + hits.length + '件★ ／ 注記の中 ' + 注記の中.length + '件');
for (const c of 注記の中) console.log('        （注記）' + c.file + ' … 「' + c.bad + '」');
T('★配る物に「Excelに保存 / Excelを開く / Excelにする」が残っていない（★画面に出る字★）', () => {
  if (hits.length) throw new Error('残っている:\n      ' + hits.join('\n      '));
});
const 照合 = 台帳を照らす(注記の中, 注記の台帳);
T('★注記の中に残っている古い言い方は 台帳に理由つきで載っている（黙って残さない）', () => {
  if (照合.載っていない.length) {
    throw new Error('台帳に無い注記が ' + 照合.載っていない.length + '件:\n      '
      + 照合.載っていない.map((c) => c.file + ' 「' + c.bad + '」').join('\n      ')
      + '\n      ★画面に出る字なら直す／記録として残すなら 注記の台帳に理由を書く★');
  }
});
T('★台帳に「もう無い物」を載せっぱなしにしない（台帳が古くならない）', () => {
  if (照合.使われていない.length) throw new Error('台帳に載っているが もう無い: ' + 照合.使われていない.join(' / '));
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
