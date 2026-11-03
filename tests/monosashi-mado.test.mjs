/* monosashi-mado.test.mjs — ★測り道具が「0」を ★2つの 窓★で 取って いるか★
 *
 *  ★★なぜ 要るか（2026-09-08 に 見つけた 物差しの 欠陥）★★
 *    `.Value2` は ★0 で ない 値に 0 を 返す★
 *      =0.1+0.2-0.3    … .Value2 ★0★ ／ =(式)=0 ★False★ ／ (式)*1e17 5.55
 *      =11.1+22.2-33.3 … .Value2 0   ／ =(式)=0 ★True★  ／ (式)*1e17 0
 *    ⇒★.Value2 では この 2つが どちらも 0 に 見える★
 *    正体 …★最後の 演算が ＋ か − の 時だけ、実Excel が ★見せる 時に★ 0 に する★
 *      `=0.1+0.2-0.3+0`   → 5.55e-17（+0 は 桁が 違うので 効かない）
 *      `=(0.1+0.2-0.3)*1` → 5.55e-17（掛け算なので 効かない）
 *    ⇒★実Excel は 値を 0 に して いない。★.Value2 が 見せ方の 側を 返して いる★
 *
 *  ★★もう1つ … 字の "0" も 同じ 顔を する★★
 *    `=DEC2BIN(0.5)` は ★文字列の "0"★（数の 0 では ない）
 *    ⇒ `="0"=0` は FALSE ⇒★見せかけの 0 と 見分けが 付かない★
 *    ⇒★型（String / Number）も 見ないと 分けられない★
 *
 *  ★★この 見張りが 守る 物★★
 *    ①★実Excel を 押す 道具が「0」を 1つの 窓だけで 取って いないか★
 *      ⇒ 2つ目の 窓 … `=(式)=0` の 真偽（★本当に 0 か★）
 *      ⇒ 3つ目の 窓 … 型（String / Number）
 *    ②★免除する 道具は ★理由つきで 名指し★★（★黙って 見逃さない★）
 *
 *  ★★見て いない 範囲（★書かない 見張りは「全部 守った」と 読まれる★）★★
 *    ・`docs/measured/**\/toru-*.ps1` だけ 見る
 *    ・★`.mjs` の 突き合わせ道具は 見て いない★（実Excel を 押さない ので）
 *    ・★『0』以外の 見せ方の 罠は 見て いない★（例：字の 幅で 変わる 表示）
 *
 *  使い方: node tests/monosashi-mado.test.mjs [--self-test]
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const ここ = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(ここ, '..');
const 直に走った = process.argv[1] && pathToFileURL(process.argv[1]).href === import.meta.url;
const 自己試験 = 直に走った && process.argv.includes('--self-test');

let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ✓ ' + n); } catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + (e && e.message)); } };

/* ★測り道具を 集める★ */
function 道具ら() {
  const 出 = [];
  const 掘る = (d) => {
    if (!fs.existsSync(d)) return;
    for (const f of fs.readdirSync(d, { withFileTypes: true })) {
      const p = path.join(d, f.name);
      if (f.isDirectory()) { 掘る(p); continue; }
      if (/^toru-.*\.ps1$/.test(f.name)) 出.push(p);
    }
  };
  掘る(path.join(ROOT, 'docs/measured'));
  return 出.sort();
}

/* ★2つ目の 窓を 持って いるか★
   `=(…)=0` の 形か `）=0'` の 形を 探す */
const 窓２ = /\)\s*=\s*0\s*\)|'\)=0'|\)=0"|=\(' \+|\)=0\)/;
function 二つ目の窓が在るか(s) {
  /* ★『=(式)=0』を 打って いるか★（書き方の 揺れを 吸う） */
  if (/\(\s*'\s*=\s*\(\s*'\s*\+/.test(s) && /\)\s*=\s*0/.test(s)) return true;
  if (/'\)=0'/.test(s)) return true;
  if (/\)\s*=\s*0\s*'\s*\)/.test(s)) return true;
  return false;
}
/* ★型を 見て いるか★ */
function 型を見て居るか(s) {
  return /-is \[string\]/.test(s) && /-is \[double\]/.test(s);
}

/* ══ ★★免除（★理由つきで 名指し★／黙って 見逃さない）★★ ══════════
   ★免除の 条件★
     ・その 道具が ★数の 答えを 1つも 取らない★（字・誤り・真偽 だけ）
     ・または ★引き算・足し算で 終わる 式を 1本も 打たない★
   ⇒★どちらも「0 が 見せかけに なる 形」に 当たらない★
   ⇒★★迷ったら 免除しない★★（★免除は 少ない方が 安全★） */
const 免除 = {
  /* ★★ 2026-09-25 に 足した 1本（列の 幅を 取る 道具）★★
       ★実物で 数えました★（★見立てでは ありません★）
         `toru-jitsu-excel-no-hashira-haba.ps1` ...
           `.Value2` ★0か所★／`.Formula` ★0か所★／`.Text` ★0か所★
       ⇒★マスの 中身を 1つも 読みません★
       ⇒読むのは `ColumnWidth` / `Width` / `RowHeight` / `Font.Size` / `Font.Name`
         ＝★「0 が 見せかけに なる 形」に 当たりません★
       ★戻す 条件★ ... `.Value2` `.Formula` `.Text` の どれかが 1か所でも 入ったら 外す
                      （★下の 門が 機械で 見ます★） */
  'docs/measured/toru-jitsu-excel-no-hashira-haba.ps1':
    '★マスの 中身を 1つも 読みません★＝列の 幅・行の 高さ・字の 大きさ/書体を 取るだけの 道具。'
    + '⇒★`.Value2` `.Formula` `.Text` とも 0か所（実物で 数えました）★。'
    + '★戻す 条件★ ... どれかが 1か所でも 入ったら 外す。',
  /* ★★ 2026-09-22 に 足した 2本（`.xlsm` の 材料を 調べる 道具）★★
       ★実物で 数えました★（★見立てでは ありません★）
         `toru-xlsm-wo-excel-ni-hirakaseru.ps1`  ... `.Value2` ★0か所★／`.Formula` ★0か所★
                                                     `.Text` 1か所（★字の A1「日付」★）
         `toru-vba-no-tobira-ga-aite-iru-ka.ps1` ... `.Value2` `.Formula` `.Text` `Range(` ★全部 0か所★
       ⇒★どちらも マスの 数を 1つも 取りません★＝★「0 が 見せかけに なる 形」に 当たりません★
       ★戻す 条件★ ... どちらかに `.Value2` / `.Formula` が 1か所でも 入ったら ★この 免除を 外す★
                      （★下の 門が それを 機械で 見ます★） */
  'docs/measured/toru-xlsm-wo-excel-ni-hirakaseru.ps1':
    '★マスの 数を 1つも 取りません★＝実Excel が `.xlsm` を ★開くか★だけを 見る 道具。'
    + '読むのは `HasVBProject` ／ 板の 名 ／ `Range(A1).Text`（★字★）／ 図形の 数 ／ VBProject の 部品の 数。'
    + '⇒★`.Value2` 0か所・`.Formula` 0か所（実物で 数えました）★。'
    + '★戻す 条件★ ... `.Value2` か `.Formula` が 1か所でも 入ったら 外す。',
  'docs/measured/toru-vba-no-tobira-ga-aite-iru-ka.ps1':
    '★マスを 1つも 触りません★＝`AccessVBOM`（守りの 設定）と 新しい 本の `HasVBProject` を 見るだけ。'
    + '⇒★`.Value2` `.Formula` `.Text` `Range(` とも 0か所（実物で 数えました）★。'
    + '★戻す 条件★ ... `Range(` が 1か所でも 入ったら 外す。',
  /* ★2026-10-04 夜★ 経営者の 道具135（一字の幅を 表に した 直しの 受け入れの 物差し） */
  'docs/measured/toru-tsukasa-no-hon-no-retsu-px.ps1':
    '★マスの 値を 1つも 読みません★＝司さんの 本の 列の 幅（字）と px（.Width）と 隠し だけを 取る 道具。'
    + '⇒★`.Value2` `.Formula` `.Text` とも 0か所（実物で 数えました・`System.Text` は 字の 書き方の 名前）★。'
    + '★戻す 条件★ ... `.Value2` `.Formula` `.Text` の どれかが 1か所でも 入ったら 外す。',
  /* ★★ 2026-09-16 に 足した 2本★★
       司さん（ 2026-09-16）「Exally と Excel に 引き渡しても ちゃんと 動くか
         ★実際に 動作確認しながら★ やれよ」の 為に 作った 道具。
       ★もう 1本（toru-oufuku-shinki.ps1）は ★値を 突き合わせる★ので
         ★免除せず 窓②を 入れました★ */
  'docs/measured/toru-xlfn.ps1':
    '★数の 答えを 突き合わせません★＝実Excel に 式を 打たせて ★.xlsx で 保存させる★だけの 道具。'
    + '見るのは ★保存された ファイルの 中の 式に `_xlfn.` が 付くか★ だけで、'
    + '画面に 出す 答えは ★目安に 見せるだけ★（紙に しません）。'
    + '⇒★「0」を 紙に 控える 所が 無いので 2つ目の 窓の 当たる 先が ありません★。'
    + '★中身を 読むのは `docs/measured/osu-xlfn.mjs`★（xml を 直に 見る）。',
  'docs/measured/toru-xlfn-zenbu.ps1':
    '★数の 答えを 1つも 紙に しません★＝台が 知る 394個を 実Excel に 打たせ、'
    + '★.xlsx で 保存させる★だけの 道具。引数は 全部 `1` なので 答えは 誤りでも 構いません'
    + '（★見るのは 式の 字だけ★）。'
    + '⇒★「0」を 紙に 控える 所が 無いので 2つ目の 窓の 当たる 先が ありません★。',
  'docs/measured/kansuu46/toru-isref.ps1': 'ISREF … ★真偽（TRUE/FALSE）しか 返さない★＝0 が 出ない',
  'docs/measured/toru-oufuku.ps1':
    '★この 道具は 実Excel に ★式を 打たせません★＝★うちが 書いた ファイルを 開いて 読むだけ★。'
    + '見るのは ①式(.Formula) ②出る字(.Text) ③答え(.Value2) の 3つで、'
    + '★「0」が 出る 所が ありません★（材料は お金の 利回りと 並べ替え）。'
    + '⇒ 2つ目の 窓（=(式)=0）の 当たる 先が 無い。'
    + '★型は 見て います★＝`.Value2` が double か どうかで 分けて 書いて います。',
  'docs/measured/toru-oou-daiarogu.ps1':
    '★実Excel の 数を 1つも 読みません★＝シートに 色を 塗って ★絵を 撮る★だけの 道具。'
    + '見るのは ★絵の 点の 明るさ★で、`.Value2` も `.Text` も 使いません。'
    + '⇒★「0」が 出る 所が 無いので 2つ目の 窓の 当たる 先が ありません★。'
    + '（★この 機械では 画面が 撮れず 走りません★＝別の 機械用に 置いて あります）',
  'docs/measured/toru-jitsubutsu-shoshiki.ps1':
    '★数の 答えを 1つも 読みません★＝司さんの 実物を 開いて ★書式の 字と 本数だけ★を 数える 道具。'
    + '見るのは ①式の 字（`.Formula` の 中の TEXT の 2つ目）②マスの 書式（`.NumberFormatLocal`）の 2つで、'
    + '★`.Value2` を 1度も 呼びません★（★中身を 読まない★のが この 道具の 決めでも あります）。'
    + '⇒★「0」が 出る 所が 無いので 2つ目の 窓の 当たる 先が ありません★。',
  'docs/measured/toru-jitsu-excel-no-e.ps1':
    '★数の 答えを 1つも 読みません★＝実Excel に ★マスを 絵に させる★だけの 道具。'
    + '（`CopyPicture` ⇒ 図の 枠に 貼る ⇒ `Export` で PNG）'
    + '★実物で 数えました★（2026-09-21・Exally1）＝`.Value2` `.Formula` `.Text` が ★0か所★。'
    + '⇒★「0」を 紙に 控える 所が 無いので 2つ目の 窓の 当たる 先が ありません★。'
    + '★★但し この 道具は 別の 所で 嘘を 出しました★★（経営者1・同じ日）'
    + '＝門が 「頭8バイトが PNG の 印か」だけ を 見て ★真っ白 229バイトを 通した★。'
    + '⇒★絵の 道具は 「投げなかったか」では なく 「図が 1つ 載ったか」で 見る★。'
    + '⇒★それは この 見張りの 持ち場では ありません★（ここは 「0」の 窓だけ 見ます）。',
  'docs/measured/toru-jitsu-excel-no-yoko-soroe.ps1':
    '★★紙に 載る 数は Excel の 答えでは ありません★★＝★絵の 点の 余り★です。'
    + '（左の 余り / 右の 余り＝字の 点が どこから どこまでか を 数えた 物）'
    + '★★この 道具は 値を 読みます★★＝`.Text`（出る 字）を 札に 使い、'
    + '`Value2` と `Formula2` で 材料を ★打ち込みます★（読むのでは なく 打つ 方）。'
    + '⇒★だから 「値を 1つも 読まない」とは 書きません★。'
    + '⇒★それでも 免除する 訳★＝この 見張りが 止めたい のは'
    + '★空の マスが 0 に 見える★ 形ですが、この 紙の 0（字 abc の 左の 余り 0）は'
    + '★絵を 数えた 0★＝★マスの 答えでは ありません★。'
    + '⇒★2つ目の 窓（=(式)=0）の 当たる 先が 有りません★。'
    + '★★2026-09-21 Exally1 が 足しました★★＝★作った のは 経営者1 です★。'
    + '★違うと 思ったら 外して ください★（免除は 少ない方が 安全）。',
  'docs/measured/toru-jitsu-excel-no-iro-no-koide.ps1':
    '★★紙に 載る 数は 色です★★＝`Interior.Color`（COM の 色の 番号）と `#RRGGBB`。'
    + '★マスの 答えを 1つも 読みません★（`Value2` は ★濃さの 数を 打ち込む 側★）。'
    + '⇒★「0」が 答えとして 出る 所が 有りません★＝2つ目の 窓の 当たる 先が ありません。'
    + '★★2026-09-21 Exally1 が 足しました★★＝★作った のは 経営者1 です★。'
    + '★違うと 思ったら 外して ください★（免除は 少ない方が 安全）。',
  'docs/measured/toru-jitsu-excel-ga-shuufuku-shita-ka.ps1':
    '★★紙に 載るのは 「開けたか／投げたか」です★★＝★マスの 答えでは ありません★。'
    + '（わざと 壊した 包みを `Workbooks.Open` に 渡して ★投げるか★ を 見る 道具）'
    + '★`Value2` は 材料を 打ち込む 側★（A1 に 1・A2 に abc を 入れる）。'
    + '⇒★「0」が 答えとして 出る 所が 有りません★＝2つ目の 窓の 当たる 先が ありません。'
    + '★★2026-09-21 Exally1 が 足しました★★＝★作った のは 経営者1 です★。'
    + '★違うと 思ったら 外して ください★。',
};

/* ★★経営者1 へ★★（2026-09-21）
     ★絵や 色を 測る 道具は この 見張りに 3本 続けて 引っかかりました★。
     ＝`toru-jitsu-excel-no-e` ／ `toru-jitsu-excel-no-yoko-soroe` ／ `toru-jitsu-excel-no-iro-no-koide`
     ＝★4本目★ `toru-jitsu-excel-ga-shuufuku-shita-ka`（2026-09-21）
     ⇒★どれも 「マスの 答え」では なく 「絵の 点」や 「色の 番号」を 紙に します★。
     ⇒★★新しい 道具を 足す 時は ここにも 1行 足して ください★★（本数も 決め打ちです）
     ⇒★私（Exally1）が 代わりに 足した 3本は どれも 実物を 読んで から 書いて います★ */

const 全 = 道具ら();
console.log('\n★測り道具の 窓★');
console.log('  ★見る 範囲★ … `docs/measured/**/toru-*.ps1` … ' + 全.length + '本');
console.log('    （★.mjs の 突き合わせ道具は 見て いません＝実Excel を 押さない ので★）');

T('★測り道具を 1本でも 見つけて いる（空振りして いない）★', () => {
  if (全.length < 10) throw new Error('★' + 全.length + '本しか 見つからない★＝探し方が おかしい');
  console.log('      ' + 全.length + '本');
});

/* ══ ★★免除の 本数は 決め打ち★★ ══（2026-09-15・指示役1 の 差し戻し）
   ★前は 上限（`> 4` なら 赤）でした★。それでも 足せば 赤に なりますが
   ★①自己試験の 中にしか 無く★（素で 走らせると ★1度も 見て いない★）
   ★②減らしても 赤に ならない★（★理由の 行だけ 残って 嘘に なる★）
   ⇒★ぴったりの 数に する／本体に 置く★＝
     ★免除を 増やす にも 減らす にも この 行を 直すしか 無い★＝★差分に 必ず 出る★
   ★今 4本の 訳★
     ①toru-isref … 真偽しか 返さない
     ②toru-oufuku … 実Excel に 式を 打たせない（うちが 書いた 物を 読むだけ）
     ③toru-oou-daiarogu … 数を 1つも 読まない（絵を 撮るだけ）
     ④toru-jitsubutsu-shoshiki … ★`.Value2` を 1度も 呼ばない★（書式の 字と 本数だけ 数える）
   ★同じ日に もう1本 作りましたが 免除して いません★
     （`toru-shoshiki-dai.ps1`＝★「生の 数だから 要らない はず」で 逃げず 窓②を 足した★
       ⇒★見た目 0・中は 0.004 の 行が 6本 出ました★＝★足して 正解★） */
/* 2026-09-16 … 4 → ★6★（上の 2本） */
/* ★★2026-09-22 ── 10 ⇒ 12（★2本 足しました★）★★
     足した 物 ... `toru-xlsm-wo-excel-ni-hirakaseru.ps1` ／ `toru-vba-no-tobira-ga-aite-iru-ka.ps1`
     ★訳★ ... `.xlsm` の 材料が 実Excel で 開くかを 見る 道具。
              ★どちらも マスの 数を 1つも 取りません★（`.Value2` `.Formula` とも ★0か所★・実物で 数えた）
     ⇒★「0 が 見せかけに なる 形」に 当たりません★ */
/* ★★2026-09-25 ── 12 ⇒ 13（★1本 足しました★）★★
     足した 物 ... `toru-jitsu-excel-no-hashira-haba.ps1`（列の 幅を 取る 道具）
     ★訳★ ... ★マスの 中身を 1つも 読みません★（`.Value2` `.Formula` `.Text` とも ★0か所★・実物で 数えた）
     ★同じ 日に 足した 他の 2本（字の 大きさ／井桁）は ★免除せず 窓②を 入れました★★
       ＝そちらは ★`.Value2` を 読んで います★（空か どうか／桁数） */
/* ★★2026-10-04 夜 ── 13 ⇒ 14（★1本 足しました★）★★
     足した 物 ... `toru-tsukasa-no-hon-no-retsu-px.ps1`（経営者の 道具135・司さんの 本の 列の px）
     ★訳★ ... ★マスの 値を 1つも 読みません★（`.Value2` `.Formula` `.Text` とも ★0か所★・実物で 数えた）
       ＝上の `toru-jitsu-excel-no-hashira-haba.ps1` と 同じ 形（列の 幅だけ） */
const 免除の本数 = 14;   /* ★2026-09-21 9 ⇒ 10（修復を 見る 道具を 1本 足した＝★これで 4本目★）★ */

T('★★免除の 本数が 決め打ちと 同じ★★（黙って 増やせない／減らせない）', () => {
  const n = Object.keys(免除).length;
  if (n !== 免除の本数) {
    throw new Error('★免除が ' + n + '本／決め打ちは ' + 免除の本数 + '本★＝'
      + '★見張りの 中の `免除の本数` も 直して ください（訳も 一緒に）★');
  }
});

T('★免除は 全部 理由つき（★黙って 見逃さない★）★', () => {
  for (const k of Object.keys(免除)) {
    if (!免除[k] || !免除[k].trim()) throw new Error('★' + k + ' に 理由が 無い★');
    if (!fs.existsSync(path.join(ROOT, k))) throw new Error('★免除に 書いた 道具が 無い … ' + k + '★');
  }
  console.log('      免除 ' + Object.keys(免除).length + '本（全部 理由つき）');
});

T('★★「0」を 1つの 窓だけで 取って いる 道具が 無い★★', () => {
  const 駄目 = [];
  for (const p of 全) {
    const 道 = path.relative(ROOT, p).replace(/\\/g, '/');
    if (免除[道]) continue;
    const s = fs.readFileSync(p, 'utf-8');
    const 窓2 = 二つ目の窓が在るか(s);
    const 型 = 型を見て居るか(s);
    if (!窓2 || !型) 駄目.push({ 道, 窓2, 型 });
  }
  console.log('      2つ目の 窓が 在る … ' + (全.length - Object.keys(免除).length - 駄目.length)
    + ' ／ 無い … ' + 駄目.length);
  if (駄目.length) {
    throw new Error('★' + 駄目.length + '本が 1つの 窓だけ★\n'
      + '      ⇒★`=(式)=0` の 真偽と ★型★を 一緒に 取って ください★\n'
      + '      ⇒ 免除するなら ★理由つきで 名指し★して ください\n'
      + 駄目.map((x) => '      ' + x.道 + '  '
        + (x.窓2 ? '' : '★=(式)=0 が 無い★ ') + (x.型 ? '' : '★型を 見て いない★')).join('\n'));
  }
});

/* ══ ★自己試験＝★わざと 見せかけの 0 を 通して 赤に なるか★★ ══════ */
if (自己試験) {
  console.log('\n★自己試験（★壊すのは 写し★＝ファイルは 1バイトも 触らない）★');

  T('★★窓②が 無い 道具は 赤に なる★★', () => {
    const 素 = '$v = $sh.Range("T1").Value2\nif ($v -is [double]) { }';
    if (二つ目の窓が在るか(素)) throw new Error('★窓②が 無いのに 在ると 言う★');
    console.log('      … `=(式)=0` を 打たない 道具 → ★赤★');
  });

  T('★窓②が 在れば 通る★', () => {
    const 良 = "$z = 押して字に ('=(' + $中 + ')=0')\nif ($v -is [string]) { } elseif ($v -is [double]) { }";
    if (!二つ目の窓が在るか(良)) throw new Error('★窓②が 在るのに 無いと 言う★');
    if (!型を見て居るか(良)) throw new Error('★型を 見て いるのに 見て いないと 言う★');
    console.log('      … `=(式)=0` ＋ 型 を 打つ 道具 → ★緑★');
  });

  T('★型を 見て いない 道具は 赤に なる★', () => {
    const 半 = "$z = 押して字に ('=(' + $中 + ')=0')\nif ($v -is [double]) { }";
    if (型を見て居るか(半)) throw new Error('★型を 見て いないのに 見て いると 言う★');
    console.log('      … 型（String/Number）を 分けない 道具 → ★赤★');
  });

  T('★★見せかけの 0 の 実物 3組を 覚えて いる★★', () => {
    /* ★この 3組は 実Excel で 実測ずみ★（golden-hikizan-zero-2026-09-08.tsv） */
    const 組 = [
      ['=0.1+0.2-0.3', '0', 'False', '5.551115123125783e-17'],
      ['=1-0.9-0.1', '0', 'False', '-2.7755575615628914e-17'],
      ['=11.1+22.2-33.3', '0', 'True', '0'],   /* ★これは 本当に 0★ */
    ];
    const 紙 = path.join(ROOT, 'docs/measured/golden-hikizan-zero-2026-09-08.tsv');
    if (!fs.existsSync(紙)) throw new Error('★実測の 紙が 無い … golden-hikizan-zero-2026-09-08.tsv★');
    const 中 = fs.readFileSync(紙, 'utf-8');
    for (const [式] of 組) {
      if (中.indexOf(式 + '\t') < 0) throw new Error('★紙に ' + 式 + ' が 無い★');
    }
    console.log('      … 3組とも 紙に 在る（★見せかけ 2 ／ 本当に 0 が 1★）');
  });

  T('★★免除を 1本 足したら 赤に なる★★（決め打ちの 数に ぶつかる）', () => {
    /* ★わざと 1本 増やした 写し★＝★ファイルは 触りません★ */
    const 増 = Object.keys(免除).length + 1;
    if (増 === 免除の本数) throw new Error('★増やしても 同じ 数＝見張りが 効いて いない★');
    console.log('      … 免除を ' + 増 + '本に すると 決め打ちの ' + 免除の本数 + ' と ちがう → ★赤★');
  });
}

console.log('\n' + pass + ' passed, ' + fail + ' failed');
if (直に走った) process.exit(fail ? 1 : 0);
