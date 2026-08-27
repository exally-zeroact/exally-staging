/* chizu-horu.test.mjs — ★7 地図＋掘る★
 *
 *  ★正本（2026-08-22 司さん決定）★
 *    1.地図を毎回 渡す（数千文字）／2.参照の網で辿る／3.AIが「もっと見せて」と言える／
 *    4.地図とよく使う所は 置いたまま使い回す
 *  ★指示役の検証要件（2026-08-26）★
 *    ★地図が何文字になったか／掘った回数／1相談でAIが何回 動いたか（数え場の行数で）★
 *    ★「どのセルを見て言ったか」が答えに必ず付く事★
 *    ★お金の数字（合計・請求額・給与）はAIに出させない★
 *    ★1回にAIへ渡すのは 2万トークンまで★
 *
 *  使い方: node tests/chizu-horu.test.mjs [--self-test]
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
const OVERRIDE = process.env.EXALLY_CHIZU_OVERRIDE ? JSON.parse(process.env.EXALLY_CHIZU_OVERRIDE) : {};
const require_ = createRequire(pathToFileURL(path.join(ROOT, 'package.json')));
const Chizu = require_(OVERRIDE['lib/chizu.js'] || path.join(ROOT, 'lib/chizu.js'));
const Horu = require_(OVERRIDE['lib/horu.js'] || path.join(ROOT, 'lib/horu.js'));
const GOLDEN = JSON.parse(fs.readFileSync(path.join(ROOT, 'tests/fixtures/shindan-golden.json'), 'utf8'));

let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ok   ' + n); } catch (e) { fail++; console.log('  NG   ' + n + '\n       ' + (e && e.message)); } };
const ok = (c, m) => { if (!c) throw new Error(m || 'expected truthy'); };
const eq = (a, b, m) => { if (a !== b) throw new Error((m ? m + ': ' : '') + '期待=' + JSON.stringify(b) + ' 実際=' + JSON.stringify(a)); };

/* 見本のブック（2シート・別シート参照つき） */
const 本 = () => ([
  { name: '売上表', data: (() => {
    const d = { '0,0': { v: '日付' }, '0,1': { v: '売上' }, '0,2': { v: '歩合' } };
    for (let r = 1; r <= 30; r++) {
      d[r + ',0'] = { v: '2026-08-' + ('0' + r).slice(-2) };
      d[r + ',1'] = { v: 1000 * r };
      d[r + ',2'] = { f: '=計算!B' + (r + 1) + '*0.1' };
    }
    d['31,1'] = { f: '=SUM(B2:B31)', v: 465000 };
    return d;
  })() },
  { name: '計算', data: { '0,0': { v: '区分' }, '0,1': { v: '単価' }, '1,1': { v: 100 } } },
]);
const 網 = { 別シート参照: [], 別ファイル参照: [], 解けない: [] };
/* ★本物の網は 行き先を「シートの番号」で持つ★（名前ではない）。
   見本も 同じ形にする＝でないと 検査が 実物と違う物を見る（2026-08-27 素通りした） */
for (let r = 1; r <= 30; r++) 網.別シート参照.push({ from: 0, fromCell: r + ',2', to: 1 });

console.log('');
console.log('[chizu-horu] ★7 地図＋掘る★');

/* ══ ①地図 ══ */
T('★シート一覧・大きさ・式の本数・見出しが 入る★', () => {
  const m = Chizu.作る(本(), 網, null);
  ok(m.字.indexOf('売上表（32行×3列') > 0, '★大きさが 入っていない★：' + m.字.slice(0, 200));
  ok(m.字.indexOf('A=日付 / B=売上 / C=歩合') > 0, '★見出しが 入っていない★');
  ok(m.字.indexOf('計算（2行×2列') > 0, '★2枚目のシートが 入っていない★');
});
T('★どのシートが どこを 何本 見ているかが 入る★', () => {
  const m = Chizu.作る(本(), 網, null);
  ok(m.字.indexOf('売上表→計算 … 30本') > 0, '★参照のまとめが 入っていない★：' + m.字);
});
T('★行き先は シートの名前（番号を出さない）★', () => {
  /* ★2026-08-27 自分で地図を読んで見つけた★
     網の to は ★シートの番号★。そのまま出していたので
     「給料表→10 … 5290本」と ★地図の いちばん大事な行が AIに読めない★状態だった。 */
  const m = Chizu.作る(本(), 網, null);
  ok(m.字.indexOf('→計算') > 0, '★行き先が 名前になっていない★：' + m.字.slice(m.字.indexOf('どのシート'), m.字.indexOf('どのシート') + 120));
  ok(!/→\d+ /.test(m.字), '★行き先に 番号が出ている★');
});
T('★網がまだ無い時は「未測定」と書く（0本と書かない）★', () => {
  const m = Chizu.作る(本(), null, null);
  ok(m.字.indexOf('未測定') > 0, '★調べ終わっていないのに 0本と言っている★');
  /* ★見るのは 参照の所だけ★（シートの「式0本」は 別の話。検査の側が 雑だった） */
  ok(m.字.indexOf('別のシートを見ている式は 0本') < 0, '★未測定なのに 0本と言い切っている★');
});
T('★機械が見つけた危ない所（診断）も 入る（AIは呼んでいない）★', () => {
  const m = Chizu.作る(本(), 網, { 式の本数: 122, 見つけた: [{ シート: '計算', セル: 'AJ269' }] });
  ok(m.字.indexOf('122か所') > 0, '★診断の数が 入っていない★');
  ok(m.字.indexOf('AIは呼んでいない') > 0, '★0円の道だと 書いていない★');
});
T('★決まり（金額を書くな／見た所を書け／もっと見せて）が 地図に入る★', () => {
  const m = Chizu.作る(本(), 網, null);
  ok(m.字.indexOf('もっと見せて') > 0, '★掘り方を 教えていない★');
  ok(m.字.indexOf('{{シート名!セル}}') > 0, '★金額の書き方を 教えていない★');
  ok(m.字.indexOf('どのセルを見て言ったか') > 0, '★根拠を書けと 言っていない★');
});
T('★数千文字に収まる（毎回 渡す物なので）★', () => {
  const m = Chizu.作る(本(), 網, null);
  ok(m.文字数 <= Chizu.既定の上限, '★' + m.文字数 + '文字（上限 ' + Chizu.既定の上限 + '）★');
  console.log('       … 2シートの地図 ' + m.文字数 + '文字');
});
T('★長い時は 黙って切らない（何文字 出していないかを 書く）★', () => {
  const 多い = [];
  for (let i = 0; i < 200; i++) 多い.push({ name: 'シート' + i, data: { '0,0': { v: '見出し' + i } } });
  const m = Chizu.作る(多い, 網, null, { 上限: 1500 });
  ok(m.削った > 0, '★削っていないのに 削ったと言っている／削ったのに 言っていない★');
  ok(m.字.indexOf('出していません') > 0, '★黙って切っている★');
  ok(m.文字数 <= 1500, '★上限を 超えている★');
});
T('★★大きさは「値か式が在る所」と「表の枠(!ref)」を 両方 出す★（2026-08-27 指示役の指摘）★', () => {
  /* ★AIは「どこまで在るか」で 掘る範囲を決める★。
     実物 計算シート … 値か式 400行×72列 ／ ★表の枠 404行×152列★
     ⇒ ★72列と言うと 73列目から先を 一生 掘らない★ */
  const sheets = 本();
  sheets[1].枠 = { 行数: 404, 列数: 152 };
  const m = Chizu.作る(sheets, 網, null);
  ok(m.字.indexOf('値か式 2行×2列／表の枠 404行×152列') > 0, '★枠を 出していない★：' + m.字.slice(m.字.indexOf('- 計算'), m.字.indexOf('- 計算') + 90));
  /* ★枠の方が広くない時は 1つだけ（無駄に長くしない）★ */
  ok(m.字.indexOf('売上表（32行×3列') > 0, '★枠と同じなのに 2つ書いている★');
});
T('★画面（book-open）が 表の枠(!ref)を 読んで 渡している★', () => {
  /* ★部品が正しくても 画面が 枠を渡さなければ 地図に出ない★
     ＝実物で「計算 72列」と出ていたのは ここが 落としていたから（2026-08-27） */
  const src = fs.readFileSync(OVERRIDE['js/book-open.js'] || path.join(ROOT, 'js/book-open.js'), 'utf8');
  /* ★2か所に ws['!ref'] が在る★（if と decode_range）ので、
     ★見るのは 判定の1行★（片方を潰しても 素通りした・2026-08-27） */
  ok(src.indexOf("if (ws['!ref']) {") > 0, '★!ref を 見ていない★');
  ok(src.indexOf("decode_range(ws['!ref'])") > 0, '★!ref を 読んでいない★');
  ok(/枠 = \{ 行数: rg\.e\.r \+ 1, 列数: rg\.e\.c \+ 1 \}/.test(src), '★枠を 作っていない★');
  ok(/枠: 枠,/.test(src), '★枠を 呼ぶ側へ 渡していない★');
});
T('★枠が 分からない時は 1つだけ（0にしない・嘘の枠を作らない）★', () => {
  const m = Chizu.作る(本(), 網, null);
  ok(m.字.indexOf('表の枠') < 0, '★枠が無いのに 枠を書いている★');
});
T('★削っても 決まりは 最後まで 残る（2026-08-27 指示役の指摘）★', () => {
  const 多い = [];
  for (let i = 0; i < 200; i++) 多い.push({ name: 'シート' + i, data: { '0,0': { v: 'とても長い見出しの名前がここに入ります' + i }, '0,1': { v: '第2の見出し' + i } } });
  const m = Chizu.作る(多い, 網, null, { 上限: 1500 });
  ok(m.文字数 <= 1500, '★上限を 超えている：' + m.文字数 + '★');
  ok(m.削った > 0, '★削っていない★');
  /* ★決まりが消えると AIは 金額を自分で書き、見た所も書かない＝いちばん危ない★ */
  ok(m.字.indexOf('## 決まり（守ってください）') >= 0, '★決まりが 消えている★');
  ok(m.字.indexOf('{{シート名!セル}}') >= 0, '★金額の決まりが 消えている★');
  ok(m.字.indexOf('どのセルを見て言ったか') >= 0, '★見た所の決まりが 消えている★');
  ok(m.字.indexOf('出していません') >= 0, '★黙って 切っている★');
});
T('★決まりだけでも 上限を超える時は 決まりを残す（客の安全が先）★', () => {
  const 多い = [];
  for (let i = 0; i < 50; i++) 多い.push({ name: 'S' + i, data: { '0,0': { v: 'あ' } } });
  const m = Chizu.作る(多い, 網, null, { 上限: 200 });
  ok(m.字.indexOf('{{シート名!セル}}') >= 0, '★上限が きつい時に 決まりを 捨てている★');
});
T('★お金の数字（合計）を 地図に書かない★', () => {
  const m = Chizu.作る(本(), 網, null);
  ok(m.字.indexOf('465000') < 0, '★合計の値を 地図に載せている（AIに金額を覚えさせない）★');
});
T('★客に見せる字に ★ を書かない（地図はAIへ渡す物だが 同じ決まり）★', () => {
  const m = Chizu.作る(本(), 網, { 式の本数: 3, 見つけた: [] });
  ok(m.字.indexOf('★') < 0, '★★ が 混ざっている★');
});

/* ══ ②掘る ══ */
T('★場所の書き方を 読む（シート付き・範囲・1セル）★', () => {
  eq(JSON.stringify(Horu.範囲を読む('計算!A1:C3')), JSON.stringify({ シート: '計算', r0: 0, c0: 0, r1: 2, c1: 2 }));
  eq(Horu.範囲を読む('B2', '売上表').シート, '売上表');
  eq(Horu.範囲を読む("'名 前'!A1").シート, '名 前');
  eq(Horu.範囲を読む('あいう'), null);
});
T('★「もっと見せて」を 読む★', () => {
  const r = Horu.頼みを読む('少し見せてください\n{"もっと見せて":["計算!A1:B2","売上表!A1:C3"]}');
  eq(r.掘る.length, 2);
  eq(r.答え, '');
});
T('★頼みが無ければ それが 答え★', () => {
  const r = Horu.頼みを読む('歩合は 計算!B2 を見ています。見た所: 売上表!C2');
  eq(r.掘る.length, 0);
  ok(r.答え.indexOf('歩合') >= 0);
});
T('★頼まれた範囲を その場で返す（式も見せる）★', () => {
  const r = Horu.範囲を出す(本(), ['売上表!A1:C3']);
  ok(r.字.indexOf('日付') > 0, '★中身を 返していない★');
  ok(r.字.indexOf('=計算!B2*0.1') > 0, '★式を 見せていない（辿れない）★');
  eq(r.返したセル, 9);
});
T('★大きすぎる頼みは 減らして「出していない」と書く★', () => {
  const 大きい = [{ name: '大', data: (() => { const d = {}; for (let r = 0; r < 500; r++) for (let c = 0; c < 10; c++) d[r + ',' + c] = { v: r * c }; return d; })() }];
  const r = Horu.範囲を出す(大きい, ['大!A1:J500'], { セル数: 500 });
  ok(r.減らした, '★減らしていない（2万トークンを超える）★');
  ok(r.返したセル <= 500, '★上限を 超えて返した：' + r.返したセル + '★');
  ok(r.字.indexOf('出していません') > 0, '★黙って 切っている★');
});
T('★無いシートを頼まれても 落ちない（黙らずに 言う）★', () => {
  const r = Horu.範囲を出す(本(), ['無いシート!A1:B2']);
  ok(r.字.indexOf('在りません') > 0, '★黙って 空を返している★');
});
T('★読めない書き方でも 落ちない★', () => {
  const r = Horu.範囲を出す(本(), ['ぐちゃぐちゃ']);
  ok(r.字.indexOf('読めませんでした') > 0);
});
T('★掘れる回数に 上限が在る（暴走させない）★', () => {
  eq(Horu.もう掘れないか(4), false);
  eq(Horu.もう掘れないか(5), true);
  const s = Horu.もう掘れない時の言葉(5);
  ok(s.indexOf('5回まで') > 0, '★何回で止めたかを 言っていない★');
  ok(s.indexOf('分かりません') > 0, '★足りない時に そう言えと 言っていない★');
});

/* ══ ③「どのセルを見て言ったか」 ══ */
T('★見た所が無い答えは 通さない★', () => {
  eq(Horu.見た所が在るか('合計は 大きいです'), false);
  eq(Horu.見た所が在るか('合計は 大きいです\n見た所: 売上表!B32'), true);
  ok(Horu.見た所を付けてと言う().indexOf('見ていないセルは 書かないでください') > 0,
    '★見ていない物を 見たと言わせない★と 書いていない');
});

/* ══ ④お金の数字を AIに出させない ══ */
T('★{{シート!セル}} を うちの計算結果に差し替える★', () => {
  const sheets = 本();
  const r = Horu.値を差し込む('8月の売上は {{売上表!B32}} 円です。見た所: 売上表!B32', sheets,
    (sh, rr, cc) => (sh.data[rr + ',' + cc] || {}).v);
  eq(r.差し込んだ, 1);
  ok(r.字.indexOf('465000 円') > 0, '★うちの値に 差し替えていない★：' + r.字);
  ok(r.字.indexOf('{{') < 0, '★書き方が そのまま出ている★');
});
T('★空・無いシートは 黙って消さない（見つからないと言う）★', () => {
  const sheets = 本();
  const r = Horu.値を差し込む('{{計算!Z99}} と {{無い!A1}}', sheets, (sh, rr, cc) => (sh.data[rr + ',' + cc] || {}).v);
  eq(r.差し込んだ, 0);
  eq(r.見つからない.length, 2);
  ok(r.字.indexOf('空です') > 0 && r.字.indexOf('在りません') > 0, '★黙って 消している★：' + r.字);
});
T('★差し込む値は うちの計算（HyperFormula）から取る★', () => {
  const sheets = 本();
  /* ★セルに書いてある古い値ではなく、渡された「今の値」を使う★ */
  const r = Horu.値を差し込む('{{売上表!B32}}', sheets, () => 999999);
  ok(r.字.indexOf('999999') >= 0, '★渡した値を 使っていない★');
});

/* ══ ★⑤1相談の通し（AIは偽物・お金は1円も使わない）★ ══ */
T('★1相談＝地図1回＋掘る2回＝AIが3回 動く（数え場の行数と同じ）★', () => {
  const sheets = 本();
  const 数え場 = [];                 /* ★1回 呼ぶごとに 1行★（本番は exally.ai_tsukatta） */
  let 掘った = 0;
  const 返事 = [
    '{"もっと見せて":["売上表!A1:C3"]}',
    '{"もっと見せて":["計算!A1:B2"]}',
    '歩合は 計算!B2 の単価を使っています。8月の売上は {{売上表!B32}} 円です。\n見た所: 売上表!C2, 計算!B2, 売上表!B32',
  ];
  let 何回目 = 0;
  const 偽AI = () => { 数え場.push({ kekka: 'ok' }); return 返事[何回目++]; };

  const 地図 = Chizu.作る(sheets, 網, null);
  let 材料 = [地図.字];
  let 答え = '';
  for (let i = 0; i < 10; i++) {
    const t = 偽AI();
    const r = Horu.頼みを読む(t);
    if (!r.掘る.length) { 答え = r.答え; break; }
    if (Horu.もう掘れないか(掘った)) { 材料.push(Horu.もう掘れない時の言葉(掘った)); continue; }
    掘った++;
    材料.push(Horu.範囲を出す(sheets, r.掘る).字);
  }
  eq(掘った, 2, '★掘った回数★');
  eq(数え場.length, 3, '★AIが動いた回数＝数え場の行数★');
  ok(Horu.見た所が在るか(答え), '★見た所が 付いていない★');
  const 差し込み = Horu.値を差し込む(答え, sheets, (sh, rr, cc) => (sh.data[rr + ',' + cc] || {}).v);
  ok(差し込み.字.indexOf('465000') > 0, '★金額が うちの値になっていない★');
  /* ★2万トークンの内側★（渡した材料の合計を 多めに見積もる） */
  const 見積 = 材料.join('\n');
  let 和 = 0, 英 = 0;
  for (const ch of 見積) { if (ch.charCodeAt(0) < 128) 英++; else 和++; }
  const トークン = 和 + Math.ceil(英 / 4);
  ok(トークン <= 20000, '★' + トークン + 'トークン（2万を超えた）★');
  console.log('       … 地図 ' + 地図.文字数 + '文字 ／ 掘った 2回 ／ AIが動いた 3回 ／ 渡した見積もり ' + トークン + 'トークン');
});
T('★掘り続けても 止まる（上限を超えたら 今ある材料で答えさせる）★', () => {
  const sheets = 本();
  let 掘った = 0, 呼んだ = 0, 止めた = false;
  for (let i = 0; i < 20; i++) {
    呼んだ++;
    const r = Horu.頼みを読む('{"もっと見せて":["売上表!A1:B2"]}');   /* ★永遠に掘りたがるAI★ */
    if (!r.掘る.length) break;
    if (Horu.もう掘れないか(掘った)) { 止めた = true; break; }
    掘った++;
  }
  eq(掘った, 5, '★上限で 止まっていない★');
  ok(止めた, '★止めたのに 止めていない★');
  ok(呼んだ <= 7, '★止めた後も 呼び続けている★');
});

/* ══ ★⑥画面が 4段どおりに 動くか（本物の book.html を読む）★ ══ */
T('★画面が 地図を毎回 渡す／掘る／止める／金額を差し替える★', () => {
  const book = 注記を外す(fs.readFileSync(OVERRIDE['book.html'] || path.join(ROOT, 'book.html'), 'utf8'), { html: true });
  ok(/Chizu\.作る\(sheets,/.test(book), '★地図を 作っていない★');
  ok(book.indexOf('var 材料 = 地図.字 ? [地図.字] : [];') > 0, '★地図を 材料に 入れていない（毎回 渡していない）★');
  ok(/材料\.length \? /.test(book), '★材料を AIへ 渡していない★');
  ok(/Horu\.頼みを読む\(aiText\)/.test(book), '★「もっと見せて」を 読んでいない★');
  ok(/Horu\.範囲を出す\(sheets,/.test(book), '★頼まれた範囲を 返していない★');
  ok(/Horu\.もう掘れないか\(掘った\)/.test(book), '★掘りっぱなし（暴走する）★');
  ok(/Horu\.見た所が在るか\(aiText\)/.test(book), '★どのセルを見て言ったかを 確かめていない★');
  ok(/Horu\.値を差し込む\(aiText, sheets,/.test(book), '★金額を AIの書いたまま 出している★');
});
T('★客の画面に 中の言葉（掘る／もっと見せて／5回／2,000セル）を 出さない★', () => {
  /* ★指示役 2026-08-27★「客は『5回』を知りません／知る必要もありません」
     ＝★中の言葉を 客に見せない★は うちの決まり（他アプリで STEP6 を見せて踏んだ）。 */
  const book = fs.readFileSync(OVERRIDE['book.html'] || path.join(ROOT, 'book.html'), 'utf8');
  const 出す所 = [];
  for (const 名 of ['showToast(', 'addAIChatMsg(']) {
    let i = 0;
    while ((i = book.indexOf(名, i)) >= 0) {
      let d = 0, j = i + 名.length - 1;
      for (; j < book.length; j++) { const c = book[j]; if (c === '(') d++; else if (c === ')') { d--; if (!d) break; } }
      出す所.push(book.slice(i, j)); i = j;
    }
  }
  ok(出す所.length >= 10, '客に出す口が 見つからない（' + 出す所.length + '件）');
  const 中の言葉 = ['掘れません', '掘る', 'もっと見せて', '2000セル', '12000字'];
  const 出た = [];
  for (const t of 出す所) {
    const 字 = (t.match(/'[^']*'/g) || []).join('');
    for (const w of 中の言葉) if (字.indexOf(w) >= 0) 出た.push(w + ' … ' + 字.slice(0, 50));
  }
  eq(出た.join(' / '), '', '★中の言葉が 客に出ている★');
  console.log('       … 客に出す口 ' + 出す所.length + '件を 数えて 中の言葉 0件');
});
T('★押した後に「AIを ◯回 使いました」を 出す（押す前は 数を書かない）★', () => {
  const book = 注記を外す(fs.readFileSync(OVERRIDE['book.html'] || path.join(ROOT, 'book.html'), 'utf8'), { html: true });
  ok(book.indexOf("+ 動いた + '回 使いました") > 0, '★何回 動いたかを 出していない★');
  /* ★「AIを1回 使います」と 押す前に 書かない★（掘ると 何度も動くので 嘘になる） */
  ok(book.indexOf('AIを1回 使います') < 0, '★押す前に 回数を約束している（掘ると 嘘になる）★');
});

T('★列が多くても 全部 出る（71列・152列・256列で押す）★', () => {
  for (const 列 of [71, 152, 256]) {
    const d = {};
    d['0,0'] = { v: '見出し' };
    d['9,' + (列 - 1)] = { v: 1 };           /* いちばん端に 1つ置く */
    const m = Chizu.作る([{ name: 'ひろい', data: d, 枠: { 行数: 10, 列数: 列 } }], 網, null);
    ok(m.字.indexOf('10行×' + 列 + '列') > 0, '★' + 列 + '列が 出ていない★：' + m.字.slice(m.字.indexOf('- ひろい'), m.字.indexOf('- ひろい') + 80));
  }
});

/* ══ ★⑦実物（司さんの .xlsb）で 地図が 数千文字に収まるか★ ══ */
const 実物 = GOLDEN.本.場所;
if (!fs.existsSync(実物)) {
  console.log('  ★未測定★ 実物が無い機械です（0件・異常なしにしない）');
} else {
  const XLSX = require_(path.join(ROOT, 'lib/xlsx.full.min.js'));
  const ZipSurgeon = require_(path.join(ROOT, 'lib/zip-surgeon.js'));
  const TableRefs = require_(path.join(ROOT, 'lib/table-refs.js'));
  const RefGraph = require_(path.join(ROOT, 'lib/ref-graph.js'));
  const Shindan = require_(path.join(ROOT, 'lib/shindan.js'));
  const bytes = new Uint8Array(fs.readFileSync(実物));
  const wb = XLSX.read(bytes, { type: 'array', cellFormula: true });
  const rr = await TableRefs.resolve(bytes, 'xlsb', wb, ZipSurgeon);
  const fixes = (rr && rr.fixes) || {};
  const sheets = wb.SheetNames.map((name) => {
    const ws = wb.Sheets[name] || {}; const data = {};
    for (const k of Object.keys(ws)) {
      if (k[0] === '!') continue;
      const c = ws[k]; const a = XLSX.utils.decode_cell(k); const key = name + '|' + a.r + ',' + a.c;
      data[a.r + ',' + a.c] = { v: c.v, f: fixes[key] !== undefined ? fixes[key] : (c.f ? '=' + c.f : undefined), t: c.t };
    }
    /* ★表の枠(!ref)も 持たせる★＝本番の js/book-open.js と 同じ形にする（でないと 実物を測れない） */
    let 枠 = null;
    try { const rg = XLSX.utils.decode_range(ws['!ref']); 枠 = { 行数: rg.e.r + 1, 列数: rg.e.c + 1 }; }
    catch (e) { 枠 = null; }
    return { name, data, 枠 };
  });
  const g = RefGraph.作る(sheets);
  const sh = Shindan.調べる(sheets);
  const m = Chizu.作る(sheets, g, sh);
  T('★実物（14シート・式15,126本）の地図が 数千文字に収まる★', () => {
    ok(m.文字数 <= Chizu.既定の上限, '★' + m.文字数 + '文字★');
    ok(m.字.indexOf('14') > 0, 'シート数が入っていない');
    console.log('       … 実物の地図 ★' + m.文字数 + '文字★（ブック全体は 約142万文字）'
      + ' ／ 削った ' + m.削った + '文字');
  });
  T('★実物の14シートで 行×列が 指示役の実測と一致する（値か式／表の枠）★', () => {
    /* ★指示役が ws['!ref'] を decode_range して数えた値★（2026-08-27） */
    const 指示役 = {
      '給料1': [82, 12], '給料2': [82, 12], '給料3': [82, 12], '給料4': [82, 12],
      '給料5': [82, 12], '給料6': [82, 12], '給料7': [82, 12], '給料8': [82, 12],
      '給料表': [468, 131], '売上表': [39, 60], '計算': [404, 152],
      '入力': [69, 6], '距離': [68, 48], '月別': [15, 11],
    };
    const ちがい = [];
    for (const s of sheets) {
      const 待つ = 指示役[s.name];
      if (!待つ) { ちがい.push(s.name + ' … 指示役の表に無い'); continue; }
      if (!s.枠) { ちがい.push(s.name + ' … 枠が 未測定'); continue; }
      if (s.枠.行数 !== 待つ[0] || s.枠.列数 !== 待つ[1]) {
        ちがい.push(s.name + ' … 枠 ' + s.枠.行数 + '行×' + s.枠.列数 + '列（指示役 ' + 待つ[0] + '行×' + 待つ[1] + '列）');
      }
    }
    eq(ちがい.join(' / '), '', '★行×列が ずれている★');
    console.log('       … 14シートとも 表の枠が一致（例 計算 404行×152列・給料表 468行×131列）');
  });
  T('★実物でも 合計の値は 地図に出さない★', () => {
    ok(m.字.indexOf('465000') < 0);
    /* 実物の大きな金額が 混ざっていないか（10万以上の数の並びを 数える） */
    const 数 = (m.字.match(/\b\d{6,}\b/g) || []).filter((x) => x.length >= 6);
    eq(数.length, 0, '★金額らしき数が 地図に出ている：' + 数.slice(0, 5).join(',') + '★');
  });
}

console.log('');
console.log('  ' + pass + ' 緑 / ' + fail + ' 赤');

if (SELF) {
  const { spawnSync } = await import('node:child_process');
  const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'exally-chizu-'));
  console.log('');
  console.log('[self-test] わざと壊して 赤くなるかを数える（★repo は読むだけ★）');
  const BREAKS = [
    ['lib/chizu.js', '★地図に 参照のまとめを 入れない★',
      (s) => s.replace("行.push('- ' + まとめ.行[j].道 + ' … ' + まとめ.行[j].本数 + '本');", '')],
    ['lib/chizu.js', '★網が無いのに 0本と言う★',
      (s) => s.replace("行.push('（未測定 … まだ調べ終わっていません）');", "行.push('（別のシートを見ている式は 0本）');")],
    ['lib/chizu.js', '★行き先を 番号のまま出す（AIに読めない）★',
      (s) => s.replace("var 先 = (typeof x.to === 'number' && sheets[x.to]) ? sheets[x.to].name : String(x.to);", 'var 先 = x.to;')],
    ['lib/chizu.js', '★見出しを 入れない★', (s) => s.replace("' 見出し: ' + s.見出し.join(' / ')", "''")],
    ['lib/chizu.js', '★上限を 外す（毎回 大きい物を渡す）★',
      (s) => s.replace('var 上限 = opt.上限 || 既定の上限;', 'var 上限 = 9999999;')],
    ['lib/chizu.js', '★表の枠を 出さない（73列目から先を 一生 掘らない）★',
      (s) => s.replace("大きさ = '値か式 ' + 大きさ + '／表の枠 ' + s.枠.行数 + '行×' + s.枠.列数 + '列';", '')],
    ['lib/chizu.js', '★枠が無いのに 枠を書く★',
      (s) => s.replace('if (s.枠 && (s.枠.行数 > s.行数 || s.枠.列数 > s.列数)) {', 'if (true) {')],
    ['lib/chizu.js', '★決まりを 先に削る（後ろから切る 昔の形に戻す）★',
      (s) => s.replace("var 決まりから = 行.indexOf('## 決まり（守ってください）');", 'var 決まりから = -1;')],
    ['js/book-open.js', '★表の枠を 捨てる（!ref を読まない）★',
      (s) => s.replace("if (ws['!ref']) {", 'if (false) {')],
    ['lib/chizu.js', '★黙って 切る★', (s) => s.replace("'…（地図が長いので 一部を 出していません。要る所は「もっと見せて」で頼んでください）'", "''")],
    ['lib/chizu.js', '★金額の書き方を 教えない★',
      (s) => s.replace('{{シート名!セル}} と書けば こちらの計算結果を入れます。', '')],
    ['lib/chizu.js', '★根拠を書けと 言わない★',
      (s) => s.replace('- 答えには どのセルを見て言ったか を必ず書いてください。', '')],
    ['lib/horu.js', '★掘れる回数の上限を 外す（暴走する）★', (s) => s.replace('掘れる回数: 5,', '掘れる回数: 9999,')],
    ['lib/horu.js', '★1回に返すセルの上限を 外す★',
      (s) => s.replace('var 上限セル = opt.セル数 || 決まり.セル数;', 'var 上限セル = 99999999;')],
    ['lib/horu.js', '★減らした事を 言わない★', (s) => s.replace("出.push('（' + (範囲.r1 - r1) + '行は 出していません。要るなら もう一度 頼んでください）');", '')],
    ['lib/horu.js', '★式を 見せない（辿れない）★', (s) => s.replace("if (cell.f && String(cell.f).charAt(0) === '=') v = String(cell.f);", '')],
    ['lib/horu.js', '★見た所が 無くても 通す★', (s) => s.replace("return String(答え || '').indexOf(見た所の印) >= 0;", 'return true;')],
    ['lib/horu.js', '★金額を AIの書いたまま 出す★', (s) => s.replace('var 出 = s.replace(/\\{\\{([^{}]+)\\}\\}/g, function (全, 中) {', 'var 出 = s.replace(/\\{\\{([^{}]+)\\}\\}/g, function (全, 中) { return 中;')],
    ['lib/horu.js', '★空を 黙って消す★', (s) => s.replace("return '（' + 中.trim() + ' … 空です）';", "return '';")],
    ['book.html', '★地図を 渡さない★',
      (s) => s.replace('var 材料 = 地図.字 ? [地図.字] : [];', 'var 材料 = [];')],
    ['book.html', '★掘りっぱなし（止めない）★', (s) => s.replace('if(Horu.もう掘れないか(掘った)){', 'if(false){')],
    ['book.html', '★金額を AIの書いたまま 出す★', (s) => s.replace('Horu.値を差し込む(aiText, sheets,', 'Horu.値を差し込む2(aiText, sheets,')],
    ['book.html', '★中の言葉を 客に見せる★',
      (s) => s.replace("'回 使いました</span>'", "'回 使いました（もっと見せて）</span>'")],
    ['book.html', '★何回 動いたかを 出さない★', (s) => s.replace("+ 動いた + '回 使いました", "+ '' + '")],
    ['lib/horu.js', '★無いシートでも 黙る★', (s) => s.replace("出.push('（' + 頼み[i] + ' … そのシートは 在りません）'); continue;", 'continue;')],
  ];
  let red = 0;
  for (const [rel, name, brk] of BREAKS) {
    const orig = fs.readFileSync(path.join(ROOT, rel), 'utf8');
    const bad = brk(orig);
    if (bad === orig) { console.log('  ★置換できず★  ' + name); continue; }
    const f = path.join(TMP, rel.replace(/[\\/]/g, '_'));
    fs.writeFileSync(f, bad, 'utf8');
    const env = Object.assign({}, process.env, { EXALLY_CHIZU_OVERRIDE: JSON.stringify({ [rel]: f }) });
    const r = spawnSync(process.execPath, ['--max-old-space-size=4096', path.join(__dirname, 'chizu-horu.test.mjs')], { encoding: 'utf8', env });
    if (r.status !== 0) { red++; console.log('  赤くなった  ' + name); }
    else console.log('  ★素通り★  ' + name);
  }
  try { fs.rmSync(TMP, { recursive: true, force: true }); } catch (e) { /* 消せなくても検査は済んでいる */ }
  console.log('');
  console.log('  ' + red + '/' + BREAKS.length + ' 通りで赤くなった');
  process.exit(red === BREAKS.length ? 0 : 1);
}

process.exit(fail ? 1 : 0);
