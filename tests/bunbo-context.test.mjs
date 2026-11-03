/* bunbo-context.test.mjs — ★コンテキストタブの ★分母★が 3か所で 揃って いるか★
 *
 *  ★★なぜ 在るか（2026-09-09・監査の 指摘）★★
 *    `docs/EXCEL_PARITY.md` の 中に ★265 と 235 が 両方 生きて いました★。
 *      265 ＝ 188 ＋ スライサー41 ＋ タイムライン36（古い 節）
 *      235 ＝ 188 ＋ スライサー24 ＋ タイムライン23（⑦の 正本）
 *    ⇒★他の 6タブ 188個は 1つも 違わず、食い違うのは 2タブだけ★
 *    ⇒★★率を 出す 時に どちらを 分母に するかで 答えが 変わる★★
 *    ⇒★分母は「文書に 書いた 数」では なく ★実測の 紙★で 決める★
 *
 *  ★★何を 守るか（1つずつ 名指し）★★
 *    ①★実測の 紙★ `docs/excel-ribbon-context-2026-08-30.tsv` の 部品の 数
 *    ②★正本★ `lib/ribbon-context-spec.js` の `数える()`
 *    ③★文書★ `docs/EXCEL_PARITY.md` の 表
 *    ⇒★この 3つが 同じ 数で なければ 赤★
 *
 *  ★★見て いない 範囲★★
 *    ・★実測の 紙が 正しいかは 見て いません★（それは 実Excel に 打ち直す 話）
 *    ・基本の 11タブの 分母は ★この 試験では 見て いません★
 *    ・★中身（部品の 名前）が 合うかは 見て いません★＝数だけ
 *
 *  使い方: node tests/bunbo-context.test.mjs [--self-test]
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createRequire } from 'node:module';

const ここ = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(ここ, '..');
const 直に走った = process.argv[1] && pathToFileURL(process.argv[1]).href === import.meta.url;
const 自己試験 = 直に走った && process.argv.includes('--self-test');

let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ✓ ' + n); } catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + (e && e.message)); } };

const 紙の道 = path.join(ROOT, 'docs/excel-ribbon-context-2026-08-30.tsv');
const 文書の道 = path.join(ROOT, 'docs/EXCEL_PARITY.md');

/* ★実測の 紙を 数える（★見出しの 行は 数えない★）★ */
function 紙を数える(字) {
  const 表 = {};
  for (const l of 字.split('\n')) {
    if (!l.trim() || l.startsWith('#')) continue;
    const c = l.split('\t');
    if (c.length < 3) continue;
    const t = c[0].trim();
    if (!t) continue;
    表[t] = (表[t] || 0) + 1;
  }
  return 表;
}

console.log('\n[bunbo-context] ★コンテキストタブの 分母が 3か所で 揃って いるか★');

const 紙 = 紙を数える(fs.readFileSync(紙の道, 'utf-8'));
const 紙の合計 = Object.values(紙).reduce((a, b) => a + b, 0);

T('★実測の 紙が 読めて いる（★空振りして いない★）★', () => {
  const n = Object.keys(紙).length;
  if (n !== 8) throw new Error('★タブが ' + n + '個★（8個のはず）… ' + Object.keys(紙).join(' '));
  console.log('      … 8タブ ／ 部品 ' + 紙の合計 + '個（' +
    Object.entries(紙).sort((a, b) => b[1] - a[1]).map(([k, v]) => k + ' ' + v).join(' / ') + '）');
});

T('★★正本（lib/ribbon-context-spec.js）が 実測の 紙と 同じ 数★★', () => {
  const require_ = createRequire(path.join(ROOT, 'package.json'));
  const s = require_(path.join(ROOT, 'lib/ribbon-context-spec.js'));
  const n = s.数える();
  if (n.タブ !== 8) throw new Error('★正本の タブが ' + n.タブ + '個★');
  if (n.部品 !== 紙の合計) {
    throw new Error('★正本 ' + n.部品 + ' ／ 実測の 紙 ' + 紙の合計 + '＝★違う★');
  }
  console.log('      … 正本 ' + n.部品 + ' ＝ 実測の 紙 ' + 紙の合計);
});

/* ★★この 文書は ★日付ごとの 記録★です（188 → 265 → 235 と 3世代 在る）★★
     ⇒★古い 数を 消すのは 筋が 違う★（★消さない・混ぜない★／どう 変わったかが 追えなく なる）
     ⇒★★守るのは「読んだ 人が 古い 数を 今の 数と 間違えない」事★★
     ⇒ 古い 行には ★すぐ 下に 印★（「今の 正本は ◯◯」）が 在る事を 見る */
T('★★文書の 表 … 古い 数の すぐ 下に「今の 正本」の 印が 在る★★', () => {
  const 行ら = fs.readFileSync(文書の道, 'utf-8').split('\n');
  const 印 = new RegExp('今の 正本は\\s*' + 紙の合計);
  const 悪い = [];
  let 今の数 = 0, 古い数 = 0;
  for (let i = 0; i < 行ら.length; i++) {
    const l = 行ら[i];
    if (!/^\|/.test(l) || !/コンテキストタブ/.test(l)) continue;
    /* ★数を 1つも 持たない 行は 見ない★＝「どう 直したか」の 説明の 行
       （例）「| 分母 | 新規の 空ブックしか 開かず ★コンテキストタブを 0個★ 数えていた | …」
       ⇒★ここに 235 を 書かせると ★説明が 嘘に なる★★ */
    if (!/\|\s*\*{0,2}★?\*{0,2}\d[\d,]*\*{0,2}★?\*{0,2}\s*\|/.test(l)) continue;
    if (new RegExp('\\b' + 紙の合計 + '\\b').test(l)) { 今の数++; continue; }
    /* ★古い 数の 行＝下 3行の どこかに 印が 要る★ */
    const 下 = 行ら.slice(i + 1, i + 4).join('\n');
    if (印.test(下)) { 古い数++; continue; }
    悪い.push((i + 1) + '行目: ' + l.trim().slice(0, 70));
  }
  if (今の数 === 0) throw new Error('★今の 数（' + 紙の合計 + '）の 行が 1つも 無い★');
  if (悪い.length) {
    throw new Error('★古い 数なのに 印の 無い 行が ' + 悪い.length + '本★\n      '
      + 悪い.join('\n      ')
      + '\n      ⇒★「今の 正本は ' + 紙の合計 + '」と すぐ 下に 書いて ください★');
  }
  console.log('      … 今の 数 ' + 今の数 + '行 ／ 古い 数（印つき）' + 古い数 + '行 ／ 印の 無い 古い 行 0');
});

T('★★昔の 数（265）を 黙って 消して いない（★消さない・混ぜない★）★★', () => {
  const 字 = fs.readFileSync(文書の道, 'utf-8');
  if (!/265/.test(字)) {
    throw new Error('★昔の 数 265 が 文書から 消えた★＝★どう 変わったかが 追えなく なる★');
  }
  if (!/古い 数/.test(字)) {
    throw new Error('★265 が 古い 数だと 書いて いない★');
  }
  console.log('      … 265 は「古い 数」と 書いて 残って いる');
});

/* ══ ★自己試験（★壊すのは 写し★＝ファイルは 1バイトも 触らない）★ ══ */
if (自己試験) {
  console.log('\n★自己試験（★壊すのは 写し★＝ファイルは 1バイトも 触らない）★');

  T('★★紙の 数が 変わったら 赤に なる★★', () => {
    const 字 = fs.readFileSync(紙の道, 'utf-8');
    const 足した = 字 + '\nスライサー\tにせの組\tにせの部品\n';
    const 後 = 紙を数える(足した);
    const 合計 = Object.values(後).reduce((a, b) => a + b, 0);
    if (合計 !== 紙の合計 + 1) throw new Error('★1本 足したのに 数が 変わらない＝数えて いない★');
    /* ★文書は 変わって いないので 食い違いに なる＝赤に なる形★ */
    const 文 = fs.readFileSync(文書の道, 'utf-8');
    if (new RegExp('\\b' + 合計 + '\\b').test(文.split('\n').filter((l) => /^\|/.test(l) && /コンテキストタブ/.test(l)).join('\n'))) {
      throw new Error('★数を 変えても 文書と 合って しまう＝この 試験は 空振り★');
    }
    console.log('      … 紙に 1本 足すと ' + 紙の合計 + ' → ' + 合計 + '（★文書と 食い違う＝赤に なる★）');
  });

  T('★数えるのに 見出し（#）の 行を 入れて いない★', () => {
    const 後 = 紙を数える('# 見出し\tあ\tい\n# もう1本\tあ\tい\nスライサー\t組\t部品\n');
    const 合計 = Object.values(後).reduce((a, b) => a + b, 0);
    if (合計 !== 1) throw new Error('★見出しを 数えて いる（' + 合計 + '）★');
    console.log('      … 見出し 2本＋中身 1本 → 1（★見出しは 数えない★）');
  });
}

console.log('\n' + pass + ' passed, ' + fail + ' failed');
if (fail) process.exit(1);
