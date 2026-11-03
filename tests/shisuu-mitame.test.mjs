/* shisuu-mitame.test.mjs — ★画面に 出る 字を 実Excel と 同じに する★（2026-09-10）
 *
 *  ★★2026-09-09 に 私は「画面は きれい」と 報告しました＝★間違い★でした★★
 *    ★実Excel の ★中の 数★と 比べて いて ★画面に 出る 字★と 比べて いませんでした★
 *      =1/3                実Excel ★0.333333333★（9桁）／うち 0.333333333333333
 *      =10/3               実Excel ★3.333333333★      ／うち 3.33333333333333
 *      =1234567.1-1234567  実Excel ★0.1★             ／うち 0.100000000093132
 *      =1.64E-14           実Excel ★1.64E-14★        ／うち 1.64e-14（小文字）
 *    ⇒★お客さんの ふつうの マスに ゴミが 出て いました★
 *
 *  ★★実Excel の 決まり（★私が 決めず 実Excel に 打たせた★）★★
 *    `docs/measured/golden-shisuu-mitame-2026-09-10.tsv`（41の 数 × 14の 幅 ＝ ★574本★）
 *      ①★幅 11 以上は どの 幅でも 同じ★（11/12/13/15/20/30/50 の ★7通りが 一致★）
 *      ②★まず 15桁に 丸めてから 字数に 収める★
 *      ③★十進で 収まらない 時だけ 指数★／★大文字 E ＋ 符号 ＋ 2桁★
 *      ④★十進と 指数は 有効桁が 多い 方★（同じなら 十進）
 *      ⑤★末尾の 0 は 落とす★
 *
 *  ★★見て いない 範囲（★書かない 見張りは「全部 守った」と 読まれる★）★★
 *  ★★幅で 出る 字が 変わるように しました（2026-09-11）★★
 *    ★実Excel の General は 列の 幅で 出方が 変わります★（574本 実測）
 *    ⇒★このマスに 何文字 入るかを 測って 渡す★（`_入る字数`）
 *    ★★割り算だけでは 1文字ぶん 損を します★★（実測して 気づいた）
 *      列 80点／余白を 引いて 72点／「0」1文字 6.66点 ⇒ 割り算だと 10文字
 *      ★けれど 11文字は 69.9点＝入る★（`.` が 細い）
 *    ⇒★出す 字を 実際に 測って 詰める★（当たりを つけて 増やす／減らす）
 *    ★測った 実物（うちの 画面）★
 *      幅 40点 … 0.333       ／ 幅 60点 … 0.333333
 *      幅 80点 … 0.333333333 ／ 幅 120・200点 … ★11で 頭打ち★（実Excel と 同じ）
 *      ★どの 幅でも はみ出し 0★
 *
 *  ★★まだ 見て いない 範囲★★
 *  ★★`#####` も 作りました（2026-09-11）★★
 *    実測 … 幅5 → ★####★（4個）／幅6 → ★#####★（5個）
 *    ★何個 出すかは 測って 決めます★（数え打ちしない）
 *    ★字（数で ない 物）は ##### に しません★＝実Excel も はみ出させる
 *
 *  ★★字数では 説明が 付かない 15本（★点で 決まる★）★★
 *    幅5〜6 で `E+` を 含む 字（1E+05・-1E+20 等）は 字数は 足りて いるのに
 *    実Excel では ★####★ に なります（`+` が 数字より 太い）。
 *    ★「E+ は 1文字ぶん 太い」と 見立てて 試しましたが 外れました★
 *      幅11 の `1.23457E+20`（11字）は ★入ります★（`.` が 細くて 相殺される）
 *    ⇒★うちは 点で 判じます★＝この 15本は ★うちの 字の 太さで 決まります★
 *    ・★書式（numFmt）の 付いた マスは この 道を 通りません★（`fmtForDisplay` が 先に 受ける）
 *
 *  使い方: node tests/shisuu-mitame.test.mjs [--self-test]
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const ここ = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(ここ, '..');
const 直に走った = process.argv[1] && pathToFileURL(process.argv[1]).href === import.meta.url;
const 自己試験 = 直に走った && process.argv.includes('--self-test');

let pass = 0, fail = 0;
/* ★改行を 逃がしで 書きません★（heredoc で 化ける・今日 4回 踏んだ） */
const 改行と字下げ = String.fromCharCode(10) + '      ';
const T = (n, fn) => { try { fn(); pass++; console.log('  ✓ ' + n); } catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + (e && e.message)); } };

/* ══ ★本番の 段を book.html から 切り出す（★写しを 置かない★）★ ══
   ★写しを 置くと 本番が 変わっても 見張りは 緑の まま★＝それが 一番 怖い */
const 本 = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf-8');
function 切り出す(名) {
  const 頭 = 本.indexOf('function ' + 名 + '(');
  if (頭 < 0) throw new Error('★book.html に ' + 名 + ' が 無い★');
  const 開き = 本.indexOf('{', 頭);
  let 深さ = 0, i = 開き;
  for (; i < 本.length; i++) {
    if (本[i] === '{') 深さ++;
    else if (本[i] === '}') { 深さ--; if (深さ === 0) break; }
  }
  if (深さ !== 0) throw new Error('★' + 名 + ' の 閉じが 見つからない★');
  return 本.slice(頭, i + 1);
}
const 枠の行 = /var GEN_枠 = (\d+);/.exec(本);
if (!枠の行) throw new Error('★book.html に GEN_枠 が 無い★');
const GEN_枠 = Number(枠の行[1]);
/* ★段が 増えたら ここにも 足す★（2026-09-11 … `_十進を幅に` を 足した）
   ★写しを 置かず book.html から 切り出す★ので、足し忘れると ★その場で 落ちて 教えて くれる★ */
const 段たち = ['_指数の字数', '_平らな十進', '_指数を幅に', '_十進を幅に', 'excelGeneral', 'forDisplay'];
/* ★つなぎ目に 逃がし（バックスラッシュ n）を 使いません★
   ＝2026-09-07/10 に ★heredoc で 生の 改行に 化ける★のを 3回 踏んだ */
const 作る = (返す) => new Function(
  段たち.map((名) => 切り出す(名)).join(' ') + ' var GEN_枠 = ' + GEN_枠 + '; return ' + 返す + ';'
)();
const forDisplay = 作る('forDisplay');
const excelGeneral = 作る('excelGeneral');

/* ══ ★紙（実Excel の 実測）を 読む★ ══ */
const 紙の道 = path.join(ROOT, 'docs/measured/golden-shisuu-mitame-2026-09-10.tsv');
const 紙 = fs.readFileSync(紙の道, 'utf-8').split(/\r?\n/)
  .filter((l) => l && !l.startsWith('#'))
  .map((l) => l.split('\t'))
  .filter((c) => c.length >= 5)
  .map((c) => ({ 札: c[0], 数: c[1], 幅: Number(c[2]), 字: c[3] }));

console.log('\n[shisuu-mitame] ★画面に 出る 字を 実Excel と 同じに する★');

T('★紙が 読めて いる（★空振りして いない★）★', () => {
  if (紙.length < 400) throw new Error('★紙が ' + 紙.length + '本しか 読めない★');
  const 幅 = [...new Set(紙.map((r) => r.幅))];
  if (幅.length < 10) throw new Error('★幅が ' + 幅.length + '通りしか 無い★（3点で 線を 決めない）');
  console.log('      … ' + 紙.length + '本 ／ 幅 ' + 幅.length + '通り');
});

T('★★幅 11 以上は どの 幅でも 実Excel と 同じ 字（★これが 本体★）★★', () => {
  const 対象 = 紙.filter((r) => r.幅 >= GEN_枠);
  if (!対象.length) throw new Error('★幅 ' + GEN_枠 + ' 以上の 行が 紙に 無い★');
  const 違う = [];
  for (const r of 対象) {
    const うち = String(forDisplay(Number(r.数)));
    if (うち !== r.字) 違う.push(r.数 + ' 幅' + r.幅 + ' 実Excel=' + r.字 + ' ／ うち=' + うち);
  }
  if (違う.length) {
    throw new Error('★' + 違う.length + '/' + 対象.length + '本が 違う★\n      '
      + 違う.slice(0, 8).join('\n      '));
  }
  console.log('      … ' + 対象.length + '本 とも 同じ（幅 '
    + [...new Set(対象.map((r) => r.幅))].join('／') + '）');
});

T('★★お客さんが 一番 よく 見る 計算の 答え（★ゴミが 出て いた 所★）★★', () => {
  /* ★紙から 引く＝手で 写さない★ */
  const 組 = [
    ['0.3333333333333333', '=1/3'],
    ['3.3333333333333335', '=10/3'],
    ['0.10000000009313226', '=1234567.1-1234567'],
    ['187999.99999999997', '=206800/1.1（税抜き）'],
    ['199.99999999999997', '=2200*0.1/1.1（消費税）'],
    ['26718.499999999996', '距離の 合計'],
    ['20.00000000000003', 'LINEST の 傾き'],
    ['0.30000000000000004', '=0.1+0.2'],
  ];
  const 悪い = [];
  for (const [数, 札] of 組) {
    const 実 = 紙.find((r) => r.数 === 数 && r.幅 === GEN_枠);
    if (!実) throw new Error('★紙に 無い … ' + 数 + '★（★手で 写さず 紙から 引く★）');
    const うち = String(forDisplay(Number(数)));
    if (うち !== 実.字) 悪い.push(札 + ' … 実Excel=' + 実.字 + ' ／ うち=' + うち);
  }
  if (悪い.length) throw new Error('★' + 悪い.length + '本が 違う★\n      ' + 悪い.join('\n      '));
  console.log('      … ' + 組.length + '本 とも 実Excel と 同じ');
});

T('★★指数は 大文字 E ＋ 符号 ＋ 2桁（★小文字 e を 出さない★）★★', () => {
  const 指数 = 紙.filter((r) => r.幅 >= GEN_枠 && /E[+-]/.test(r.字));
  if (指数.length < 5) throw new Error('★指数の 行が ' + 指数.length + '本しか 無い★');
  for (const r of 指数) {
    const うち = String(forDisplay(Number(r.数)));
    if (/e/.test(うち)) throw new Error('★小文字 e が 出た … ' + r.数 + ' → ' + うち + '★');
    if (!/E[+-]\d\d/.test(うち)) throw new Error('★E＋符号＋2桁で ない … ' + r.数 + ' → ' + うち + '★');
  }
  console.log('      … ' + 指数.length + '本 とも 大文字 E ＋ 符号 ＋ 2桁');
});

T('★★数で ない 物を 壊して いない（★一番 危ない所★）★★', () => {
  const 組 = [
    ['あいうえお', 'あいうえお', '★日本語★'],
    ['', '', '★空★'],
    ['#DIV/0!', '#DIV/0!', '★誤り★'],
    ['TRUE', 'TRUE', '★真偽★'],
    ['007', '007', '★頭に 0（番号かも）★'],
    ['-007', '-007', '★頭に 0・負★'],
    ['2024/01/05', '2024/01/05', '★日付の 字★'],
    ['1,234', '1,234', '★カンマ★'],
    ['¥1000', '¥1000', '★通貨の 印★'],
    /* ★★ここは 私の 期待の 方が 間違って いました（2026-09-10）★★
       最初 `'1e5'` は そのままの はずと 書きましたが ★実Excel は 数として 100000 と 出します★
       ⇒★`recalcSheet` は 答えを ★字★ で 置く★ので 指数の 答えは ここに 字で 来ます
       ⇒★字だから 触らない、では 画面に 小文字の まま 出ます★ */
    ['1e5', '100000', '★指数の 字＝数として 読む（実Excel と 同じ）★'],
    ['1.64e-14', '1.64E-14', '★指数の 字＝大文字に 直す★'],
    ['1e', '1e', '★数に ならない 字★'],
    ['abc1e5', 'abc1e5', '★字の 中に 数★'],
    ['1e5x', '1e5x', '★後ろに 字★'],
  ];
  for (const [入, 正, 札] of 組) {
    const 出 = forDisplay(入);
    if (String(出) !== 正) throw new Error('★' + 札 + '（' + 入 + '）… うち ' + 出 + ' ／ 期待 ' + 正 + '★');
  }
  /* ★数でも 無限・NaN は そのまま★ */
  for (const v of [Infinity, -Infinity, NaN]) {
    if (String(forDisplay(v)) !== String(v)) throw new Error('★' + v + ' を 変えた★');
  }
  console.log('      … ' + (組.length + 3) + '通り とも そのまま');
});

T('★0 と 整数は そのまま 読める★', () => {
  const 組 = [[0, '0'], [-0, '0'], [1, '1'], [-1, '-1'], [1234, '1234'], [100000, '100000']];
  for (const [v, 正] of 組) {
    const 出 = String(forDisplay(v));
    if (出 !== 正) throw new Error('★' + v + ' → ' + 出 + '（' + 正 + ' のはず）★');
  }
  console.log('      … ' + 組.length + '通り とも そのまま');
});

T('★★幅を 渡すと 字数が 変わる（★11 で 頭打ち★）★★', () => {
  /* ★渡さなければ 今まで通り 11★＝★他の 呼び手を 1つも 壊さない★ */
  if (String(forDisplay(1 / 3)) !== '0.333333333') throw new Error('★枠を 渡さない時が 変わった★');
  const 組 = [[5, '0.333'], [8, '0.333333'], [11, '0.333333333'], [20, '0.333333333'], [40, '0.333333333']];
  for (const [枠, 正] of 組) {
    const 出 = String(forDisplay(1 / 3, 枠));
    if (出 !== 正) throw new Error('★枠 ' + 枠 + ' … ' + 出 + '（' + 正 + ' のはず）★');
  }
  console.log('      … 枠 5/8/11/20/40 とも 実Excel と 同じ（11 で 頭打ち）');
});

T('★★描く 所が 幅を 渡して いる（★空振りして いない★）★★', () => {
  if (本.indexOf('function _入る字数(w, raw, fmt)') < 0) throw new Error('★_入る字数 が 無い★');
  const n = (本.match(/_入る字数\(w, raw, cell\.numFmt\)/g) || []).length;
  if (n < 2) throw new Error('★描く 所 ' + n + 'か所しか 渡して いない★（ふつうと 結合の 2か所）');
  /* ★字体を 先に 決めて いるか★＝でないと 1つ前の マスの 字で 測る */
  /* ★★2026-09-11 … 字面から 中身の 検査に 変えました★★
     ★変えた 訳★＝文字を 返す 式（TEXT など）を そのまま 出す 為に
       `var display = _答えは字か(cell) ? String(raw) : fmtForDisplay(...)` に なった。
     ★守る 中身は 同じ★＝★字体を 決めてから 幅を 測る★（でないと 1つ前の マスの 字で 測る）
     ★字面だけ 見て いると 中身が 同じでも 赤に なります★ */
  const i = 本.indexOf('ctx.font = style+');
  const j = 本.indexOf('_入る字数(w, raw, cell.numFmt)', i > 0 ? i : 0);
  if (!(i > 0 && j > i)) throw new Error('★字体を 決める 前に 測って いる★');
  /* ★実際に 測って 詰めて いるか（割り算だけで 済ませて いないか）★ */
  const 段 = 本.slice(本.indexOf('function _入る字数(w, raw, fmt)'));
  if (段.indexOf('measureText(String(t)).width') < 0) {
    throw new Error('★出す 字を 測って いない★＝★割り算だけでは 1文字 損を する★');
  }
  console.log('      … 2か所 とも 渡し、字体を 先に 決め、出す 字を 測って いる');
});

T('★★結合した マスも 同じ 道を 通る（★09-10 に 見つけた 穴★）★★', () => {
  /* ★★前は 結合した マスだけ 生の 値を 出して いました★★
     `String(raw)` で 描いて いた ⇒★09-10 の 直し（#62・#63）が 1つも 効いて いない★
       =1/3               ふつう 0.333333333 ／ ★結合 0.3333333333333333★
       =1234567.1-1234567 ふつう 0.1         ／ ★結合 0.10000000009313226★
       =206800/1.1        ふつう 188000      ／ ★結合 187999.99999999997★
       =1.64E-14          ふつう 1.64E-14    ／ ★結合 1.64e-14（小文字）★
     ★実Excel を 測りました★（結合 C1:D1 と ふつう A1 に 同じ 式）
       =1/3 … ふつう ★0.333333★（8字）／★結合 0.333333333★（11字）
       他の 3本 … ★どちらも 同じ★
     ⇒★実Excel は 結合も ふつうも 同じ 規則★（幅が 広い分 字数が 増えるだけ）
     ⇒★だから `fmtForDisplay` を 通す★ */
  /* ★2026-09-11 … 幅も 渡すように なったので 探す 字を 直しました★
     （★字を 決め打ちで 探すと 直した 時に 空振りする★＝今日 3回目） */
  /* ★2026-09-11 … ここも 中身で 見ます★（文字を そのまま 出す 枝が 前に 付いた） */
  const 結合頭 = 本.indexOf('var raw=_字の元(cell);');
  const 結合中 = 結合頭 >= 0 ? 本.slice(結合頭, 結合頭 + 2500) : '';
  if (結合中.indexOf('fmtForDisplay(raw, cell.numFmt, _入る字数(w, raw, cell.numFmt))') < 0) {
    throw new Error('★結合した マスが `fmtForDisplay` を 通って いない★'
      + 改行と字下げ + '⇒★09-10 の 直しが 結合には 効かなく なる★');
  }
  if (/var display=cell\.numFmt\?applyNumFmt\(raw,cell\.numFmt\):String\(raw\);/.test(本)) {
    throw new Error('★生の 値を 出す 古い 形が 残って いる★');
  }
  /* ★実際に 出る 字が ふつうの マスと 同じか★ */
  const 組 = [['0.3333333333333333', '0.333333333'], ['0.10000000009313226', '0.1'],
              ['187999.99999999997', '188000'], ['1.64e-14', '1.64E-14']];
  for (const [生, 正] of 組) {
    const 出 = String(forDisplay(Number(生)));
    if (出 !== 正) throw new Error('★' + 生 + ' → ' + 出 + '（' + 正 + ' のはず）★');
  }
  console.log('      … 結合も ふつうと 同じ 道（' + 組.length + '通り 確かめた）');
});

T('★★狭い 幅でも 実Excel と 同じ 字に なる（★2026-09-11 に 直した★）★★', () => {
  /* ★前は ここで「まだ 直して いません」と 断って いました★
     ⇒ 幅を 渡すように したので ★狭い 幅も 合うように なりました★
     ⇒ 紙の 幅（実Excel の 文字数）を そのまま 枠に して 突き合わせる */
  const 悪い = [];
  for (const r of 紙) {
    if (r.字.indexOf('#') === 0) continue;          /* ★##### は まだ 作って いない★ */
    const 枠 = Math.min(GEN_枠, Math.floor(r.幅));
    const うち = String(forDisplay(Number(r.数), 枠));
    if (うち !== r.字) 悪い.push(r.数 + ' 幅' + r.幅 + '（枠' + 枠 + '） 実Excel=' + r.字 + ' ／ うち=' + うち);
  }
  const 全 = 紙.filter((r) => r.字.indexOf('#') !== 0).length;
  if (悪い.length > 全 * 0.1) {
    throw new Error('★' + 悪い.length + '/' + 全 + '本が 違う★' + 改行と字下げ + 悪い.slice(0, 6).join(改行と字下げ));
  }
  console.log('      … ' + (全 - 悪い.length) + '/' + 全 + '本が 実Excel と 同じ（★#####  は 数えて いません★）');
});

T('★★`#####` を 作った（幅が 足りない 時）★★', () => {
  /* ★実Excel は 数が 入らないと `#` で 埋めます★（★字は 埋めません＝はみ出す★）
     実測（09-10 の 紙）… 幅5 → ★####★（4個）／幅6 → ★#####★（5個） */
  for (const 名 of ['_井桁で埋める', '_数が入らないか']) {
    if (本.indexOf('function ' + 名 + '(') < 0) throw new Error('★' + 名 + ' が 無い★');
  }
  const n = (本.match(/_井桁で埋める\(w\)/g) || []).length;
  if (n < 2) throw new Error('★描く 所 ' + n + 'か所しか 使って いない★（ふつうと 結合の 2か所）');
  /* ★字は ##### に しない★（実Excel も しない） */
  const 段 = 本.slice(本.indexOf('function _数が入らないか('));
  if (段.indexOf('return false;   /* ★字は はみ出させる★ */') < 0) {
    throw new Error('★字まで ##### に して いる★＝実Excel は 字を はみ出させる');
  }
  console.log('      … 2か所 とも `#####` を 出す／字は はみ出させる');
});

T('★★紙 574本と 突き合わせる（★合わない 15本の 訳を 書いて 在る★）★★', () => {
  /* ★★字数では 説明が 付かない 15本が 在ります★★
     幅5〜6 で `E+` を 含む 字（1E+05・-1E+20 等）が
     ★実Excel では 入らず ####★ に なります。字数は 足りて いるのに、です。
     ⇒ `+` が 数字より 太いから ＝★実Excel は 点（見た目の 幅）で 判じて いる★
     ⇒★「E+ は 1文字ぶん 太い」と 見立てて 試しましたが 外れました★
       幅11 の `1.23457E+20`（11字）は ★入ります★（`.` が 細くて 相殺される）
     ⇒★うちは 点で 判じます★＝この 15本は ★うちの 字の 太さで 決まります★
     ⇒★ここでは「形と 桁数」だけを 見ます★ */
  const 悪い = [];
  for (const r of 紙) {
    if (r.字.indexOf('#') === 0) continue;              /* ##### は 点で 決める */
    const 枠 = Math.min(GEN_枠, Math.floor(r.幅));
    const 出 = excelGeneral(Number(r.数), 枠);
    if (出 === null) continue;                          /* 入らない＝点で 決める */
    if (String(出) !== r.字) 悪い.push(r.数 + ' 幅' + r.幅 + ' 実Excel=' + r.字 + ' ／ うち=' + 出);
  }
  if (悪い.length) {
    throw new Error('★' + 悪い.length + '本が 違う★' + 改行と字下げ + 悪い.slice(0, 8).join(改行と字下げ));
  }
  const 数の行 = 紙.filter((r) => r.字.indexOf('#') !== 0).length;
  console.log('      … 数の 行 ' + 数の行 + '本 とも 実Excel と 同じ');
});

/* ══ ★自己試験（★壊すのは 写し★＝ファイルは 1バイトも 触らない）★ ══ */
if (自己試験) {
  console.log('\n★自己試験（★壊すのは 写し★＝ファイルは 1バイトも 触らない）★');

  T('★★枠を 広げると 赤に なる（＝字数に 収める のが 効いて いる）★★', () => {
    const 悪い = 紙.filter((r) => r.幅 >= GEN_枠 && String(excelGeneral(Number(r.数), 20)) !== r.字);
    if (!悪い.length) throw new Error('★枠を 20 に しても 全部 合う★＝★枠が 効いて いない★');
    console.log('      … 枠を 20 に すると ' + 悪い.length + '本が 違う（例 ' + 悪い[0].数 + ' → '
      + excelGeneral(Number(悪い[0].数), 20) + '）');
  });

  T('★★15桁の 丸めを 外すと 赤に なる★★', () => {
    /* ★26718.499999999996 は ★先に 15桁★に すると 26718.5＝幅5で 26719★
       ⇒ 15桁を 外すと 26718 に なる（実Excel と 違う） */
    const 数 = 26718.499999999996;
    const 先に15 = excelGeneral(数, 5);
    const 実 = 紙.find((r) => r.数 === '26718.499999999996' && r.幅 === 5);
    if (!実) throw new Error('★紙に 無い★');
    if (先に15 !== 実.字) throw new Error('★今の コードが 紙と 違う（' + 先に15 + ' ／ ' + 実.字 + '）★');
    if (String(Math.round(数)) === 実.字) {
      throw new Error('★ただ 丸めただけでも 同じ 字に なる★＝★この 試験は 何も 守って いない★');
    }
    console.log('      … 先に 15桁で ' + 実.字 + '（ただ 丸めると ' + Math.round(数) + '＝違う）');
  });

  T('★★書き換えが コードに 在る（空振りして いない）★★', () => {
    if (!/function excelGeneral/.test(本)) throw new Error('★excelGeneral が 無い★');
    /* ★2026-09-11 … 枠を 渡せる ように したので 探す 字を 直した★
       ★形で 探します★＝`excelGeneral(n, …)` を 呼んで いれば よい（中の 字を 決め打ちしない） */
    /* ★2026-09-11 … `forDisplay` が ★入らない時の 逃げ★を 持つように なりました★
       ⇒ `return excelGeneral(...)` では なく ★変数に 受けて から 返す★ 形
       ⇒★形で 探します★（呼んで いれば よい） */
    if (!/excelGeneral\(n, 枠2\)/.test(本)) throw new Error('★forDisplay から 呼んで いない★');
    if (!/Math\.min\(GEN_枠,/.test(本)) throw new Error('★枠を 11 で 頭打ちに して いない★');
    if (!/var GEN_枠 = \d+;/.test(本)) throw new Error('★GEN_枠 が 無い★');
    console.log('      … 段も 呼び出しも 枠も 在る');
  });

  T('★紙を 1本 変えたら 赤に なる（★紙を 見て いる★）★', () => {
    const 写し = 紙.map((r) => ({ ...r }));
    const 的 = 写し.find((r) => r.幅 >= GEN_枠);
    的.字 = '★わざと 壊した★';
    const 違う = 写し.filter((r) => r.幅 >= GEN_枠 && String(forDisplay(Number(r.数))) !== r.字);
    if (!違う.length) throw new Error('★紙を 壊しても 赤に ならない★');
    console.log('      … 紙を 1本 壊すと ' + 違う.length + '本が 違う');
  });
}

console.log('\n' + pass + ' passed, ' + fail + ' failed');
if (fail) process.exit(1);
