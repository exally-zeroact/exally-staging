/* zairyou-ga-nai-kami.mjs — ★「紙は 在るが `#材料` が 無い」紙を 数える★（2026-09-18）
 *
 *  ★★なぜ★★
 *    2026-09-18、`golden-oddf-to-46ko-2026-09-16.tsv`（46個の 紙）に
 *    ★`#材料` が 無い★ 為に ★見張りが 1行も 押して いませんでした★。
 *    ⇒★★「紙が 在る」は「押されて いる」では ありません★★
 *    ⇒★同じ 穴が 他にも 開いて いないか★を 数えます。
 *
 *  ★数え方★
 *    ①`docs` の `*.tsv` `*.csv`（★`golden-jitsubutsu-*` は 読みません★＝司さんの 実物）
 *    ②`#材料` の 行が 在るか
 *    ③★式が マスを 指して いるか★（A1・$A$1・A1:B2・板!A1）
 *    ⇒★★③が 在って ②が 無い 紙＝押せない 紙★★
 *
 *  ★★出す 数を 2回 小さく しました★★（2026-09-18・経営者1 の 差し戻し）
 *    ⑴★最初 ... 44枚 11,702行★
 *    ⑵★突き合わせの 出しを 引いた ... 38枚 11,702行★
 *    ⑶★★頭（`#`）が 1行も 無い 物も 突き合わせの 出しだと 分かった★★
 *        ＝`mae-ato-*.tsv` の 2列目は 「合った」「こちらだけ誤り」＝★判定★
 *        ＝★実Excel の 答えでは ありません★
 *        ⇒★31枚 ★1,680行★★
 *    ⑷★道具から 名指しされて いる 物を 引いた ⇒ ★1枚 44行★★
 *  ⇒★★11,702 を そのまま 出して いたら ★分母を 出さない 数★でした★★
 *    ＝「11,702行が 押せて いない」に 化けます
 *
 *  ★見て いない 事★
 *    ・★名指しされて いても ★その 行を 押して いる★とは 限りません★
 *    ・★名指しされて いなくても 歩き回る 道具が 読む 事が 在ります★
 *    ・★柱（# 種 ...）が 読めるかは 見て いません★
 *
 *  使い方: node docs/measured/zairyou-ga-nai-kami.mjs
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '../..');

const 紙 = [];
const 歩く = (d) => {
  for (const e of fs.readdirSync(d, { withFileTypes: true })) {
    const p = path.join(d, e.name);
    if (e.isDirectory()) { 歩く(p); continue; }
    if (!/\.(tsv|csv)$/.test(e.name)) continue;
    if (/^golden-jitsubutsu-/.test(e.name)) continue;   /* ★司さんの 実物は 別の 決め★ */
    紙.push(p);
  }
};
歩く(path.join(ROOT, 'docs'));

/* ★マスを 指す 形★（`shiki-kansuu-kami.test.mjs` の 見方と 揃える） */
const マス = /(?:^|[^A-Za-z0-9_.!$])\$?[A-Z]{1,3}\$?[0-9]{1,7}(?![0-9A-Za-z_])/;

let 材あり = 0;
const 押せない = [];
const 指さない = [];
for (const p of 紙) {
  const 行 = fs.readFileSync(p, 'utf-8').split(/\r?\n/);
  const 材 = 行.some((l) => l.startsWith('#材料'));
  const 本 = 行.filter((l) => l && !l.startsWith('#'));
  const 指す = 本.filter((l) => マス.test(l)).length;
  if (材) { 材あり++; continue; }
  /* ★突き合わせの 出しは ★紙では ありません★（`shiki-kansuu-kami` も 読んで いません） */
  /* ★逆斜線を 書かない★（★今日 4回 heredoc で 落ちました★）＝名前だけ 見る */
  const 名 = path.basename(p);
  /* ★★突き合わせの 出しか★★（2026-09-18）
       ①名前で 分かる 物 ... `-awase-` `cases-`
       ②★★頭（`#`）が 1行も 無い 物★★
          ＝★測った 紙は 必ず 頭に 訳を 書いて います★（Excel の 版・材料・柱）
          ＝★頭が 無い 物は 道具が 吐いた 突き合わせの 出し★
          ★実物★ `mae-ato-*.tsv` ... 2列目が 「合った」「こちらだけ誤り」＝★判定★
                                    ＝★実Excel の 答えでは ありません★ */
  const 頭 = 行.some((l) => l.startsWith('#'));
  const 突き = 名.indexOf('-awase-') >= 0 || 名.indexOf('cases-') === 0 || !頭;
  if (指す > 0) 押せない.push({ 紙: path.relative(ROOT, p), 行: 本.length, 指す: 指す, 突き: 突き });
  else 指さない.push(path.relative(ROOT, p));
}

console.log('');
console.log('★★「紙は 在るが `#材料` が 無い」を 数える★★');
console.log('');
console.log('  ★紙★ ......................... ' + 紙.length + '枚');
console.log('  ★`#材料` が 在る★ ............ ' + 材あり + '枚');
console.log('  ★`#材料` が 無い★ ............ ' + (紙.length - 材あり) + '枚');
const 本物 = 押せない.filter((x) => !x.突き);
const 突き数 = 押せない.length - 本物.length;
console.log('    ・★マスを 指す 式が 在る★ ... ' + 押せない.length + '枚');
console.log('        うち ★★本当の 紙 ' + 本物.length + '枚★★（★押せません★）');
console.log('        うち 突き合わせの 出し ' + 突き数 + '枚（★元から 紙では ない★）');
console.log('            ＝名前が -awase- / cases- ／ ★頭（#）が 1行も 無い★');
console.log('    ・★押せない 行の 合計★ ...... ★' + 本物.reduce((a, x) => a + x.指す, 0) + '行★');
console.log('    ・マスを 指さない .......... ' + 指さない.length + '枚（★材料が 要らない★）');
console.log('');
console.log('★★押せない 紙（★材料が 要るのに 無い★）★★');
for (const x of 本物.sort((a, b) => b.指す - a.指す)) {
  console.log('  ' + String(x.指す).padStart(5) + '行が マスを 指す ／ 全 '
    + String(x.行).padStart(5) + '行  ' + x.紙);
}
/* ══ ★★引き算★★ ══（2026-09-18・経営者1「11,702 を そのまま 出すな」）
     ★訳★ ... ★私は 自分で「別の 試験で 押されて いるかは 見て いない」と 書いた★
       ⇒★その まま だと ★11,702行が 押せて いない★ という 数に 化けます★
       ⇒★★分母を 出さない 数★★＝今日 何度も 潰して きた 型そのもの
     ★引き方★ ... ★`tests/` `scripts/` `docs/measured/` の 道具を 全部 読み、
                  ★紙の 名前が 名指しで 出て くるか★ を 見る★
     ★★これも 上限つきの 言い方です★★
       ・★名指しで 出て いても「その 行を 押して いる」とは 限りません★
       ・★名指しで 出て いなくても 歩き回る 道具が 読む 事は 在ります★
       ⇒★だから 出す 数は ★「どの 道具からも 名指しされて いない」★ と 書きます★ */
const 道具 = [];
const 道具を歩く = (d) => {
  if (!fs.existsSync(d)) return;
  for (const e of fs.readdirSync(d, { withFileTypes: true })) {
    const p = path.join(d, e.name);
    if (e.isDirectory()) { if (e.name === 'node_modules') continue; 道具を歩く(p); continue; }
    if (!/\.(mjs|js|cjs|py|ps1|yml|json)$/.test(e.name)) continue;
    道具.push(p);
  }
};
/* ★★数える 範囲を `tests/` に 絞りました★★（2026-09-18・経営者1 が 同じ 穴に 落ちかけた）
     ★前は `scripts/` `docs/measured/` `.github/` も 数えて いました★
     ⇒★★`docs/measured/kami-no-katachi.md` は ★紙の 目録★です＝★走りません★★★
       ＝経営者1 が repo 全体で 引いたら「62枚中 61枚が 読まれて いる」と 出た
       ＝★一番 多く 読んで いたのが その 目録★
     ⇒★★「名前が 在る」を「押されて いる」と 数えて いました★★
       ＝★私の ③（紙が 在る ≠ 確かめられる）と 同じ 型★
     ⇒★走る 物だけ 数える＝`tests/`★ */
道具を歩く(path.join(ROOT, 'tests'));
const 道具の中身 = 道具.map((p) => fs.readFileSync(p, 'utf-8')).join('\n');

const 名指しあり = [], 名指しなし = [];
for (const x of 本物) {
  const 名 = path.basename(x.紙);
  (道具の中身.indexOf(名) >= 0 ? 名指しあり : 名指しなし).push(x);
}
console.log('');
console.log('★★引き算（★道具から 名指しされて いるか★）★★');
console.log('  ★読んだ 道具★ ................ ' + 道具.length + '本'
  + '（★`tests/` だけ＝★走る 物だけ★★）');
console.log('  ★名指しされて いる★ .......... ' + 名指しあり.length + '枚'
  + '（' + 名指しあり.reduce((a, x) => a + x.指す, 0) + '行）');
console.log('  ★★名指しされて いない★★ ...... ★★' + 名指しなし.length + '枚★★'
  + '（★★' + 名指しなし.reduce((a, x) => a + x.指す, 0) + '行★★）');
console.log('');
console.log('★★どの 道具からも 名指しされて いない 紙★★');
for (const x of 名指しなし.sort((a, b) => b.指す - a.指す)) {
  console.log('  ' + String(x.指す).padStart(5) + '行が マスを 指す ／ 全 '
    + String(x.行).padStart(5) + '行  ' + x.紙);
}

/* ══ ★★もう 1つの 数え方（★経営者1 の 問い★）★★ ══（2026-09-18）
     ★私の 問い★ ... ★押せない 紙★（材料が 要るのに 無い ＋ 誰も 名指ししない）
     ★経営者1 の 問い★ ... ★★宙に 浮いた 紙★★（★`tests/` の どれからも 名指しされない★）
       ＝★材料が 要るかどうかを 問わない★／★行は `#` で 始まらない 行を 全部 数える★
     ⇒★★問いが 違うので 数も 違います★★（6枚250行 ／ 19枚1,670行）
     ⇒★どちらも 出します★＝★「どちらが 正しいか」では なく「何を 数えたか」★ */
const 浮いた = [];
const 外した = [];
for (const p of 紙) {
  const 名 = path.basename(p);
  if (道具の中身.indexOf(名) >= 0) continue;
  /* ★範囲は `docs/measured` だけ★＝`docs` の 直下は ★Excel の 覚書き★（目録・設定・型の 名簿）
       ＝`docs/excel-objectmodel-*.tsv` は 24,820行 在りますが ★式の 紙では ありません★ */
  if (path.relative(ROOT, p).indexOf(path.join('docs', 'measured')) !== 0) {
    外した.push({ 紙: path.relative(ROOT, p), 訳: 'docs/measured の 外（Excel の 覚書き）' });
    continue;
  }
  const 行2 = fs.readFileSync(p, 'utf-8').split(/\r?\n/);
  const 本2 = 行2.filter((l) => l && !l.startsWith('#'));
  /* ★突き合わせの 出しは 外す★（上と 同じ 見方＝名前 か 頭が 無い） */
  const 頭2 = 行2.some((l) => l.startsWith('#'));
  if (名.indexOf('-awase-') >= 0 || 名.indexOf('cases-') === 0 || !頭2) {
    外した.push({ 紙: path.relative(ROOT, p), 訳: '突き合わせの 出し（紙では ない）' });
    continue;
  }
  浮いた.push({ 紙: path.relative(ROOT, p), 行: 本2.length });
}
console.log('');
console.log('★★もう 1つの 数え方（★`tests/` の どれからも 名指しされない 紙★）★★');
console.log('  ★材料が 要るかを 問いません★／★`#` で 始まらない 行を 全部 数えます★');
console.log('  ★範囲は `docs/measured` だけ★／★突き合わせの 出しは 外します★'
  + '（外した ' + 外した.length + '枚）');
console.log('  ★★' + 浮いた.length + '枚 ／ ' + 浮いた.reduce((a, x) => a + x.行, 0) + '行★★');
for (const x of 浮いた.sort((a, b) => b.行 - a.行)) {
  console.log('  ' + String(x.行).padStart(6) + '行  ' + x.紙);
}

console.log('');
console.log('★★言えない 事★★');
console.log('  ・★名指しされて いても ★その 行を 押して いる★とは 限りません★');
console.log('    ＝`tests/kansuu46-1taba.test.mjs` の ように');
console.log('      ★材料の 出どころを 名指しして 押して いる 物も 在ります★');
console.log('  ・★名指しされて いなくても ★歩き回る 道具★が 読む 事が 在ります★');
console.log('    ＝`shiki-kansuu-kami.test.mjs` は `docs` を 全部 歩きます');
console.log('      （★但し `#材料` が 無い 紙は ★行を 押しません★★）');
console.log('  ・★柱（# 種 ...）が 読めるかは 見て いません★');
console.log('');
console.log('  ⇒★★だから 出せる 数は 1つだけ★★');
console.log('    ＝★「`#材料` が 無く、マスを 指し、どの 道具からも 名指しされて いない 紙」★');
console.log('    ＝★★' + 名指しなし.length + '枚 ／ ' + 名指しなし.reduce((a, x) => a + x.指す, 0) + '行★★');
