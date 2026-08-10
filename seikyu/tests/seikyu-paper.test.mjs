/* seikyu-paper.test.mjs — ★自作テンプレ std1（紙）が、出すべき物を出しているか★
 *
 * ここで止めたい事故:
 *   ① ★税率ごとの区分が紙に出ない★（適格請求書として成り立たない）
 *   ② ★インボイスの登録番号(T+13桁)と請求番号を同じ欄に出す★（まったくの別物）
 *   ③ 発行済みなのに ★今のマスタの名前★ で刷る（写しと紙が食い違う）
 *   ④ 明細0行で ★空の表★ を出す（何も無いと分からない）
 *   ⑤ ★注意書きが1文字ずつ縦に割れる★（flex/grid の箱に文を入れた前科2回）
 *   ⑥ 紙の窓にアプリの画面が混ざる（＝紙だけの窓になっていない）
 *
 * ⑤について:
 *   jsdom は幅を計算しないので、ここでは ★書き方★ を見る
 *   （紙のCSSに flex/grid を1つも使わない・文の箱は折り返し可で最低幅を持つ）。
 *   ★実物の幅は、実配信の画面で目と定規で見る★（この検査はその代わりではなく、前段の網）。
 *
 * 使い方: node seikyu/tests/seikyu-paper.test.mjs
 *         node seikyu/tests/seikyu-paper.test.mjs --self-test
 */
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..', '..');
const require_ = createRequire(path.join(ROOT, 'package.json'));
const PAPER = require_(path.join(ROOT, 'seikyu/lib/seikyu-paper.js'));
const TAX = require_(path.join(ROOT, 'seikyu/lib/seikyu-tax.js'));
const DOC = require_(path.join(ROOT, 'seikyu/lib/seikyu-doc.js'));
const SR = require_(path.join(ROOT, 'kyuyo/lib/shouhizei-ritsu.js'));

const STD = Math.round(SR.hyojun * 10000) / 100;   // 標準税率(%) ★数字を書かずに取る
const RED = Math.round(SR.keigen * 10000) / 100;   // 軽減税率(%)

let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ✓ ' + n); } catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + (e && e.message)); } };
const eq = (a, b, m) => { if (a !== b) throw new Error((m ? m + ': ' : '') + 'expected ' + JSON.stringify(b) + ' got ' + JSON.stringify(a)); };
const ok = (v, m) => { if (!v) throw new Error(m || 'false'); };

/* 実物にいちばん近い1通（代行請求の実物の形＝標準税率＋軽減税率＋対象外が混ざる） */
function sample(over) {
  const lines = [
    { name: '運転代行 9月分', qty: 42, unit: '件', price: 3200, rate: STD },
    { name: 'お弁当代', amount: 1000, rate: RED },
    { name: '立替金（対象外）', amount: 500, rate: 0 },
  ];
  const tax = TAX.compute({ lines, taxMode: 'exclusive', rounding: 'floor' });
  return Object.assign({
    inv: {
      doc_type: 'invoice', no: '202609-001', issue_ymd: '2026-09-30', due_ymd: '2026-10-31',
      tax_mode: 'exclusive', rounding: 'floor', status: 'draft',
      data: { subject: '9月分 運転代行ご利用料金', memo: '振込手数料は貴社にてご負担ください' },
    },
    tax,
    partner: { name: '藤原建設株式会社', keisho: '御中', person: '山田', zip: '794-0000', addr: '愛媛県今治市1-2-3', invoiceNo: 'T9876543210987' },
    org: { yago: '株式会社ゼロアクト', addr: '愛媛県今治市4-5-6', tel: '0898-00-0000', invoiceNo: 'T1234567890123', bank: '伊予銀行 今治支店 普通 1234567' },
  }, over || {});
}

/* ★純関数（self-test で作り物を通せる）★ やってはいけない方の紙 */
export function paperBad(kind, o) {
  const b = PAPER.build(o);
  if (kind === 'noRates') return b.html.replace(/<table class="rates">[\s\S]*?<\/table>/, '');
  // ★番号の欄（No.　…）だけを登録番号に差し替える（題名は触らない）
  if (kind === 'mixNo') return b.html.replace(/(No\.　)[^<]*/, '$1T1234567890123');
  if (kind === 'flex') return b.html.replace('.note-b{display:block', '.note-b{display:flex');
  // ★金額を塗りつぶした角丸の箱に入れる＝差し戻しの原因そのもの
  if (kind === 'grandBox') return b.html.replace('.grand{margin:0 0 6mm;', '.grand{background:#EEF7F1;border:1px solid #CDE7D8;border-radius:2mm;padding:4mm 6mm;margin:0 0 6mm;');
  return b.html;
}

/* ── self-test：わざと壊して赤になるか ─────────────────────────────── */
if (process.argv.includes('--self-test')) {
  console.log('\n[seikyu-paper --self-test] わざと壊して赤になるか');
  let sp = 0, sf = 0;
  const S = (n, fn) => { try { fn(); sp++; console.log('  ✓ ' + n); } catch (e) { sf++; console.log('  ✗ ' + n + ' — ' + (e && e.message)); } };

  S('① 区分の表を消した紙は「税率ごとの区分がある」検査に落ちる', () => {
    const bad = paperBad('noRates', sample());
    ok(!/class="rates"/.test(bad), '作り物なのに区分が残っている＝この検査が空振り');
    ok(/class="rates"/.test(PAPER.build(sample()).html), '本物に区分が無い');
  });

  S('② 番号(No.)の欄に登録番号を入れた紙は「別物として出す」検査に落ちる', () => {
    const bad = paperBad('mixNo', sample());
    ok(/No\.　T\d{13}/.test(bad), '作り物が壊れていない＝この検査が空振り');
    ok(!/No\.　T\d{13}/.test(PAPER.build(sample()).html), '本物が登録番号を番号の欄に出している');
  });

  S('④ ★金額を塗りつぶした箱に入れた紙は「枠なし」検査に落ちる（差し戻しの原因）', () => {
    const bad = paperBad('grandBox', sample());
    const badRule = (/\.grand\{([^}]*)\}/.exec(bad) || [])[1] || '';
    ok(/background/.test(badRule) && /border-radius/.test(badRule), '作り物が壊れていない＝この検査が空振り');
    const good = (/\.grand\{([^}]*)\}/.exec(PAPER.css()) || [])[1] || '';
    ok(!/background/.test(good) && !/border-radius/.test(good), '本物の金額が箱に入っている');
    ok(/border-bottom/.test(good), '本物の金額の下に線が無い');
  });

  S('③ 文の箱を flex にした紙は「flex/grid を使わない」検査に落ちる', () => {
    const bad = paperBad('flex', sample());
    ok(/display:flex/.test(bad), '作り物が壊れていない＝この検査が空振り');
    ok(!/display\s*:\s*(flex|grid|inline-flex|inline-grid)/.test(PAPER.css()), '本物のCSSに flex/grid がある');
  });

  console.log('\n[self-test] ' + sp + ' passed, ' + sf + ' failed');
  if (sf) process.exit(1);
}

/* ── 本体 ──────────────────────────────────────────────────────── */
console.log('\n[請求書 自作テンプレ std1（紙）]');
const S1 = sample();
const B1 = PAPER.build(S1);
const H1 = B1.html;
const flat = (s) => s.replace(/\s+/g, '');

T('★1枚の完成したHTMLで返る（そのまま新しい窓に書ける）', () => {
  ok(/^<!DOCTYPE html>/.test(H1), 'DOCTYPE が無い');
  ok(/<html lang="ja">/.test(H1), 'html タグが無い');
  ok(/<style>/.test(H1), '見た目が入っていない（別ファイルを読むと刷る時に間に合わない）');
  eq(B1.templateId, 'std1');
});

T('★紙だけ＝アプリの画面が1バイトも入っていない', () => {
  ['botnav', 'appbar', 'b-issue', 'scr-list', 'seg-b', 'finput'].forEach((x) => {
    ok(!new RegExp(x).test(H1), '画面の部品が紙に混ざっている: ' + x);
  });
  ok(!/<script/i.test(H1), '紙に script が入っている');
});

T('★税率ごとの区分が、計算した区分の数だけ出る', () => {
  ok(/class="rates"/.test(H1), '区分の表が無い');
  const n = (H1.match(/% 対象<\/th>/g) || []).length;
  eq(n, S1.tax.byRate.length, '区分の行数');
  ok(/消費税の対象外/.test(H1), '対象外の行が出ていない');
  // 区分の税額が紙の上に実額で出ている
  S1.tax.byRate.forEach((b) => {
    ok(flat(H1).includes(PAPER.yen(b.tax)), '区分 ' + b.pct + '% の消費税 ' + b.tax + ' が紙に無い');
  });
});

T('★登録番号(T+13桁)と番号(No.)は別の欄に出す', () => {
  ok(/No\.　202609-001/.test(H1), '番号が出ていない（No.　＋全角スペース）');
  ok(!/No\.　T\d{13}/.test(H1), '番号の欄に登録番号が出ている');
  ok(/登録番号 T1234567890123/.test(H1), '自社の登録番号が出ていない');
});

T('★合計＝税抜＋消費税 が紙の上でも一致する', () => {
  const t = S1.tax;
  eq(t.subtotal + t.taxTotal, t.grandTotal);
  ok(flat(H1).includes(PAPER.yen(t.grandTotal)), '合計が紙に無い');
  ok(flat(H1).includes(PAPER.yen(t.subtotal)), '小計が紙に無い');
});

T('★発行済みは「写しの宛先」で刷る（マスタを直しても紙は変わらない）', () => {
  const at = '2026-09-30T00:00:00.000Z';
  const snap = DOC.snapshotOf({
    at, partner: { id: 'pt1', data: { name: '写しの名前 株式会社', keisho: '様', addr: '写しの住所' } },
    org: { data: { yago: '写しの自社' } }, tax: S1.tax, templateId: 'std1',
  });
  const h = PAPER.build({ inv: S1.inv, tax: S1.tax, partner: snap.partner, org: snap.org }).html;
  ok(/写しの名前 株式会社/.test(h), '写しの名前で刷られていない');
  ok(!/藤原建設/.test(h), 'マスタの名前が混ざっている');
  ok(/写しの自社/.test(h), '写しの自社情報が出ていない');
});

T('★敬称は hub が保存している keisho も読む（御中に化けない）', () => {
  eq(PAPER.honorOf({ keisho: '様' }), '様');
  eq(PAPER.honorOf({ honor: '御中' }), '御中');
  eq(PAPER.honorOf({ honor: '様', keisho: '御中' }), '様', 'honor が優先されていない');
  eq(PAPER.honorOf({ keisho: '（なし）' }), '', '「（なし）」が紙に出ている');
  ok(/様/.test(PAPER.build(sample({ partner: { name: 'A社', keisho: '様' } })).html), 'keisho の「様」が紙に出ていない');
});

T('★取れなかったを空欄にしない（相手・自社・番号）', () => {
  const h = PAPER.build({ inv: { doc_type: 'invoice', no: '', issue_ymd: '', tax_mode: 'exclusive', data: {} }, tax: S1.tax, partner: {}, org: {} }).html;
  ok(/（取引先が未選択）/.test(h), '宛先が空欄になっている');
  ok(/（自社情報が未入力）/.test(h), '自社が空欄になっている');
  ok(/No.　（未採番）/.test(h), '番号が空欄になっている');
  ok(/（未入力）/.test(h), '請求日が空欄になっている');
});

T('★明細0行で空の表を出さない（何も無いと分かる文を出す）', () => {
  const empty = TAX.compute({ lines: [], taxMode: 'exclusive', rounding: 'floor' });
  const h = PAPER.build({ inv: { doc_type: 'invoice', no: 'X', issue_ymd: '2026-09-30', tax_mode: 'exclusive', data: {} }, tax: empty, partner: { name: 'A' }, org: {} }).html;
  ok(/明細がまだ1行もありません/.test(h), '空の表が出ている');
  ok(/区分はまだありません/.test(h), '空の区分が出ている');
});

T('★支払期限は決めていなければ出さない（勝手な期限を作らない）', () => {
  const h = PAPER.build({ inv: { doc_type: 'invoice', no: 'X', issue_ymd: '2026-09-30', due_ymd: '', tax_mode: 'exclusive', data: {} }, tax: S1.tax, partner: { name: 'A' }, org: {} }).html;
  ok(!/お支払期限/.test(h), '決めていないのに期限の欄が出ている');
  ok(/お支払期限/.test(H1), '決めてあるのに期限が出ていない');
});

T('見積書は見出しと呼び方が変わる', () => {
  const h = PAPER.build(sample({ inv: Object.assign({}, S1.inv, { doc_type: 'quote' }) })).html;
  ok(/見　積　書/.test(h), '見出しが請求書のまま');
  ok(/ご見積金額（税込）/.test(h), '金額の呼び方が請求書のまま');
  ok(/下記の通り御見積申し上げます。/.test(h), '挨拶が請求書のまま');
  ok(/見積日　/.test(h), '日付の呼び方が請求書のまま');
});

T('金額は桁区切り・マイナスは頭に付く・数でない物は0にしない', () => {
  eq(PAPER.comma(142660), '142,660');
  eq(PAPER.comma(-142660), '-142,660');
  eq(PAPER.comma(0), '0');
  eq(PAPER.yen('abc'), '—', '数でない物が0になっている');
  eq(PAPER.jpDate('2026-09-30'), '2026年9月30日');
  eq(PAPER.jpDate('2026/09/30'), '', '読めない日付が通っている');
});

T('★文字はそのまま埋め込まない（HTMLとして壊れる／差し込まれる）', () => {
  const h = PAPER.build(sample({ partner: { name: '<script>alert(1)</script>' } })).html;
  ok(!/<script>alert/.test(h), '取引先名の中のタグがそのまま出ている');
  ok(/&lt;script&gt;/.test(h), 'エスケープされていない');
});

/* ── ⑥ ★紙に焼き付けてよい型は「法律」だけ★ ─────────────────────────
 * ★2026-08-10 方向の訂正（指示役）★
 *   いったん「代行請求の言葉10個が紙に無ければ赤」という見張りを入れましたが、★外しました★。
 *   理由: 代行請求は ★源泉なし・繰越なし・非課税なし・1税率★ ＝ 一番 単純な1業種の紙です。
 *   それを「正」として錠を掛けると、★複雑な業種の客が全部 落ちます★
 *   （士業＝源泉の行が要る／掛け売り＝繰越が要る／不動産＝非課税が要る）。
 *   代行請求は ★見本（1例）★であって、写す相手ではない。
 *
 *   だから ここで縛るのは次の2つだけ:
 *     ① ★法定の記載事項（国税庁 適格請求書）★ … 業種に関係なく必ず要る＝焼き付けてよい唯一の型
 *     ② ★客に読めない言い方を出さない★     … 言葉の good/bad ではなく「意味が通じない」を止める
 *   言い回し（御請求金額／No.　／挨拶文／和暦／¥／ページ送りの言葉）は ★出せる★ままにし、
 *   ★出さなければ赤、はやめました★。業種ごとに変えられる余地を残すためです。 */

/* ★客に読めない言い方（出たら赤）★
   「うちの語彙に無い」ではなく「読んだ人が意味を取り違える」物だけを止める。 */
const UNREADABLE = [
  { w: '外税／消費税込み', why: '外税なのか税込なのか、読んだ人には分からない（相反する言葉が並んでいる）' },
  { w: '外税/消費税込み', why: '同上' },
  { w: 'undefined', why: 'プログラムの穴がそのまま紙に出ている' },
  { w: 'NaN', why: '数にならなかった物がそのまま紙に出ている' },
  { w: '[object', why: '中身ではなく入れ物の名前が紙に出ている' },
];

T('★客に読めない言い方を紙に出さない（言葉の好き嫌いではなく、意味が通じない物を止める）', () => {
  UNREADABLE.forEach((x) => ok(H1.indexOf(x.w) < 0, '読めない言い方が出ている「' + x.w + '」＝' + x.why));
});

/* ★法定の記載事項（国税庁 適格請求書等保存方式）＝業種に関係なく必ず要る6つ★
   ここは ★焼き付けてよい唯一の型★。列を自由にしても、様式を替えても、必ず残る。
     ① 発行する側の名称と ★登録番号★
     ② 取引年月日
     ③ 取引の内容（★軽減税率の対象である旨★）
     ④ 税率ごとに区分して合計した対価の額と ★適用税率★
     ⑤ 税率ごとに区分した ★消費税額★
     ⑥ 受け取る側の名称 */
T('★法定6項目が、列を自由にしても・様式を替えても紙に残る', () => {
  // わざと「税率の列も品名の列も消した」並びで刷る＝会社が列を削っても法定は残るか。
  // ★内容は「行き先」が持つ（列名が品名である必要は無い＝代行なら行き先が取引の内容）★
  const cols = { items: ['日付', '行き先', '金額'], widths: {}, aligns: {} };
  const lines = [
    { name: '運転代行', amount: 10000, rate: STD, extra: { 日付: '9/3', 行き先: '今治→松山（運転代行）' } },
    { name: 'お弁当代', amount: 1000, rate: RED, extra: { 日付: '9/4', 行き先: 'お弁当代（軽減税率）' } },
  ];
  const t = TAX.compute({ lines, taxMode: 'exclusive', rounding: 'floor' });
  for (const id of ['std1', 'elegant']) {
    const h = PAPER.build({
      inv: { doc_type: 'invoice', no: 'X-1', issue_ymd: '2026-09-30', tax_mode: 'exclusive', template_id: id, data: {} },
      tax: t, cols, partner: S1.partner, org: S1.org,
    }).html;
    const f = h.replace(/\s+/g, '');
    ok(h.indexOf('株式会社ゼロアクト') >= 0, id + ' ①発行する側の名称が無い');
    ok(h.indexOf('T1234567890123') >= 0, id + ' ①登録番号が無い');
    ok(/2026.9.30|令和8年9月30日/.test(h), id + ' ②取引年月日が無い');
    ok(h.indexOf('今治→松山（運転代行）') >= 0 && h.indexOf('お弁当代（軽減税率）') >= 0, id + ' ③取引の内容が無い');
    // ④⑤ 税率ごとの対象額・適用税率・消費税額（★列を消しても（内訳）が必ず出す★）
    t.byRate.forEach((b) => {
      ok(h.indexOf(String(b.pct) + '% 対象') >= 0, id + ' ④適用税率 ' + b.pct + '% の区分が無い');
      ok(f.indexOf(PAPER.yen(b.base).replace(/\s/g, '')) >= 0, id + ' ④' + b.pct + '% の対象額が無い');
      ok(f.indexOf(PAPER.yen(b.tax).replace(/\s/g, '')) >= 0, id + ' ⑤' + b.pct + '% の消費税額が無い');
    });
    ok(t.byRate.length >= 2, '軽減税率の行が混ざっていない＝この検査が空振り');
    ok(h.indexOf('藤原建設株式会社') >= 0, id + ' ⑥受け取る側の名称が無い');
  }
});

T('★法定6項目の見張りが空振りしていない（1つ抜いたら赤になる）', () => {
  // 登録番号を空にした紙は ①に落ちる
  const h = PAPER.build(sample({ org: Object.assign({}, S1.org, { invoiceNo: '' }) })).html;
  ok(h.indexOf('T1234567890123') < 0, '作り物が壊れていない＝この検査が空振り');
  ok(H1.indexOf('T1234567890123') >= 0, '本物に登録番号が無い');
});

T('★御請求金額は「枠なし＋下に線」（塗りつぶした箱に入れない）', () => {
  const rule = (/\.grand\{([^}]*)\}/.exec(PAPER.css()) || [])[1] || '';
  ok(rule, '.grand の指定が無い');
  ok(!/background/.test(rule), '★金額が塗りつぶしの箱に入っている★: ' + rule);
  ok(!/border-radius/.test(rule), '★金額が角丸の箱に入っている★');
  ok(/border-bottom\s*:\s*[\d.]+p?t?\s+solid/.test(rule), '金額の下に線が無い');
  ok(!/border\s*:\s*1px solid/.test(rule), '金額が枠で囲まれている');
});

T('★お振込先・備考も箱で囲まない（うちは囲まない）', () => {
  const rule = (/\.note\{([^}]*)\}/.exec(PAPER.css()) || [])[1] || '';
  ok(!/background|border-radius|border\s*:\s*1px/.test(rule), '振込先・備考が箱に入っている: ' + rule);
  ok(!/class="note-box"/.test(H1), '古い箱の書き方が残っている');
});

/* 2026-08-11 実機で発生：共有マスタの「お振込先」に 物 が入っていて、
   紙に ★[object Object]★ と刷られた。客に渡る紙なので、読めない物は出さない。 */
T('★紙に [object Object] を刷らない（読めない物は出さない）', () => {
  [{ name: '伊予銀行', no: '1234567' }, ['伊予銀行', '普通'], {}].forEach((bad, i) => {
    const h = PAPER.build(sample({ org: Object.assign({}, sample().org, { bank: bad }) })).html;
    ok(!/\[object Object\]/.test(h), i + ': 紙に [object Object] が出ている');
    ok(!/お振込先/.test(h), i + ': 中身の無い「お振込先」の見出しだけ出ている');
  });
  const inv = Object.assign({}, sample().inv);
  inv.data = Object.assign({}, inv.data, { memo: { a: 1 } });
  const h2 = PAPER.build(sample({ inv })).html;
  ok(!/\[object Object\]/.test(h2), '備考に [object Object] が出ている');
  // ★ちゃんとした文字列は今までどおり出る（消しすぎない）
  ok(/伊予銀行 今治支店 普通 1234567/.test(H1), '文字列の振込先まで消している');
});

T('★小計・消費税・合計は枠なし、合計の上に線', () => {
  const css = PAPER.css();
  const td = (/\.sums td\{([^}]*)\}/.exec(css) || [])[1] || '';
  ok(/border\s*:\s*0/.test(td), '合計欄が罫線で囲まれている: ' + td);
  const g = (/\.sums-g td\{([^}]*)\}/.exec(css) || [])[1] || '';
  ok(/border-top\s*:\s*[\d.]+p?t?\s+solid/.test(g), '合計の上に線が無い');
});

T('★消費税のラベルは区分から作る（率の数字を書かない）', () => {
  const one = TAX.compute({ lines: [{ name: 'a', amount: 1000, rate: STD }], taxMode: 'exclusive', rounding: 'floor' });
  eq(PAPER.taxLabel(one, 'exclusive'), '消費税（' + STD + '%）');
  eq(PAPER.taxLabel(one, 'inclusive'), '消費税（' + STD + '%・内税）');
  const two = TAX.compute({ lines: [{ name: 'a', amount: 1000, rate: STD }, { name: 'b', amount: 1000, rate: RED }], taxMode: 'exclusive', rounding: 'floor' });
  eq(PAPER.taxLabel(two, 'exclusive'), '消費税');
  ok(/消費税（/.test(H1) || /消費税/.test(H1), '消費税のラベルが紙に無い');
});

T('★日付は和暦も出せる（既定は代行請求と同じ西暦）', () => {
  eq(PAPER.dateStr('2026-09-30', 'seireki'), '2026/9/30');
  eq(PAPER.dateStr('2026-09-30', 'reiwa'), '令和8年9月30日');
  eq(PAPER.dateStr('2026-09-30'), '2026/9/30', '既定が西暦でない');
  eq(PAPER.dateStr('2026/09/30', 'reiwa'), '', '読めない日付が通っている');
  const w = PAPER.build(sample({ inv: Object.assign({}, S1.inv, { data: Object.assign({}, S1.inv.data, { dateEra: 'reiwa' }) }) })).html;
  ok(/令和8年9月30日/.test(w), '和暦が出せていない');
});

T('★金額は ¥ 記号（invoice-pdf.js:156 と同じ）', () => {
  eq(PAPER.yen(142660), '¥142,660');
  eq(PAPER.yen(0), '¥0');
  eq(PAPER.yen(-1234), '¥-1,234');
  eq(PAPER.yen('abc'), '—', '数でない物が0になっている');
  // 表の中は ¥ を付けない（桁が詰まる）＝ invoice-pdf.js の comma() と同じ
  eq(PAPER.comma(142660), '142,660');
  eq(PAPER.comma(''), '');
});

T('★角印（会社の印）が紙に出る／入れていなければ出さない', () => {
  const seal = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUg==';
  const withSeal = PAPER.build(sample({ org: Object.assign({}, S1.org, { sealDataUrl: seal, sealSizeMm: 30 }) })).html;
  ok(/class="seal"/.test(withSeal), '印が紙に出ていない');
  ok(/width:30mm;height:30mm/.test(withSeal), '大きさが効いていない');
  ok(withSeal.indexOf(seal) >= 0, '印の画像が入っていない');
  // 入れていない会社の紙には出さない（空の枠を出さない）
  ok(!/class="seal"/.test(H1), '印を入れていないのに枠が出ている');
});

T('★角印の大きさは 10〜40mm に収める（紙からはみ出す印を作らない）', () => {
  eq(PAPER.sealMm(999), 40);
  eq(PAPER.sealMm(1), 10);
  eq(PAPER.sealMm(), 21, '既定が21mmでない');
  eq(PAPER.sealMm('abc'), 21, '数でない値が通っている');
  const h = PAPER.build(sample({ org: Object.assign({}, S1.org, { sealDataUrl: 'data:image/png;base64,iVBORw0KGgo=', sealSizeMm: 999 }) })).html;
  ok(/width:40mm/.test(h), '上限に収まっていない');
});

T('★角印は薄く重ねる（実物と同じ扱い・文字を隠し切らない）', () => {
  const rule = (/\.seal\{([^}]*)\}/.exec(PAPER.css()) || [])[1] || '';
  ok(/opacity\s*:\s*\.?9/.test(rule), '印が濃すぎる（下の文字が読めなくなる）: ' + rule);
  ok(/object-fit\s*:\s*contain/.test(rule), '印が歪む（縦横比を保っていない）');
});

T('★表の上に【…】の小さなキャプションが出る', () => {
  ok(/【9月分 運転代行ご利用料金】/.test(H1), 'キャプションが無い');
});

T('★明細が多い時は次の紙へ送る（黙って切らない・3つの言葉が出る）', () => {
  const many = [];
  for (let i = 0; i < 40; i++) many.push({ name: '行' + (i + 1), amount: 1000, rate: STD });
  const t = TAX.compute({ lines: many, taxMode: 'exclusive', rounding: 'floor' });
  const b = PAPER.build({ inv: S1.inv, tax: t, partner: S1.partner, org: S1.org });
  ok(b.pages > 1, '1枚に押し込めている: ' + b.pages);
  ok(/このページの小計/.test(b.html), '「このページの小計」が無い');
  ok(/次ページへ続く →/.test(b.html), '「次ページへ続く →」が無い');
  ok(/1ページ目/.test(b.html) && /2ページ目/.test(b.html), '「nページ目」が無い');
  // ★全部の行が どれかの紙に出ている（黙って落ちていない）
  for (let i = 0; i < many.length; i++) ok(b.html.indexOf('行' + (i + 1) + '<') >= 0, (i + 1) + '行目が紙から消えた');
  // 金額のラベルは1枚目にだけ（invoice-pdf.js:719「単ページのみ上部に御請求金額」と同じ考え）
  eq((b.html.match(/ご請求金額（税込）/g) || []).length, 1, '金額のラベルが2枚以上に出ている');
  // 合計・振込先は最後の紙にだけ
  eq((b.html.match(/お振込先/g) || []).length, 1, 'お振込先が2枚以上に出ている');
  // ページの小計の合計＝全体の小計
  eq(PAPER.paginate(many).reduce((a, p) => a + p.length, 0), many.length, 'ページ分けで行が増減した');
});

/* ── ⑤ 文が縦に割れない書き方 ─────────────────────────────────── */
const CSS = PAPER.css();

T('★紙のCSSに flex/grid を1つも使わない（文が1文字ずつ縦に割れる前科の形）', () => {
  ok(!/display\s*:\s*(flex|inline-flex|grid|inline-grid)/.test(CSS), 'flex/grid が使われている');
  ok(!/display\s*:\s*(flex|grid)/.test(H1), '紙の中の style に flex/grid がある');
});

T('★文が入る箱は「折り返し可」で「最低幅」を持つ', () => {
  // 長い日本語が入りうる箱＝ここが潰れると1文字ずつ縦になる
  ['.lead-l', '.note-b'].forEach((sel) => {
    const rule = (new RegExp(sel.replace('.', '\\.') + '\\{([^}]*)\\}').exec(CSS) || [])[1];
    ok(rule, sel + ' の指定が無い');
    ok(/min-width\s*:\s*\d/.test(rule), sel + ' に最低幅が無い（箱が潰れる）');
    ok(/overflow-wrap\s*:\s*break-word/.test(rule), sel + ' に折り返しの指定が無い');
    ok(/word-break\s*:\s*normal/.test(rule), sel + ' の word-break が normal でない（break-all は1文字ずつ割れる）');
    ok(/display\s*:\s*block/.test(rule), sel + ' が block でない');
  });
});

T('★word-break:break-all を紙のどこにも使わない（日本語が1文字ずつ割れる）', () => {
  ok(!/word-break\s*:\s*break-all/.test(CSS), 'break-all が使われている');
});

T('★2段組み（宛先と自社）は表で作る＝幅が足りなくても文が縦に割れない', () => {
  ok(/<table class="party">/.test(H1), '2段組みが表になっていない');
  ok(/\.party-to\{[^}]*min-width\s*:\s*\d/.test(CSS), '宛先の欄に最低幅が無い');
  ok(/\.party-from\{[^}]*min-width\s*:\s*\d/.test(CSS), '自社の欄に最低幅が無い');
});

T('★色の決まり（濃すぎる緑は使わない・全アプリ #2E7D54）', () => {
  ok(!/#1A4A2E/i.test(CSS), '使ってはいけない濃い緑がある');
  ok(/#2E7D54/i.test(CSS), '決められた緑が使われていない');
});

T('★網羅：税率の組み合わせ×内外×丸め を全部刷って、区分の数と合計が紙と一致', () => {
  let n = 0;
  const sets = [
    [{ name: 'a', amount: 105, rate: STD }],
    [{ name: 'a', amount: 105, rate: RED }],
    [{ name: 'a', amount: 105, rate: 0 }],
    [{ name: 'a', amount: 105, rate: STD }, { name: 'b', amount: 1000, rate: RED }],
    [{ name: 'a', amount: 105, rate: STD }, { name: 'b', amount: 1000, rate: RED }, { name: 'c', amount: 500, rate: 0 }],
    [{ name: 'a', amount: -1100, rate: STD }],
  ];
  for (const lines of sets) {
    for (const mode of ['exclusive', 'inclusive']) {
      for (const rd of ['floor', 'ceil', 'round']) {
        const t = TAX.compute({ lines, taxMode: mode, rounding: rd });
        if (!t.ok) throw new Error('計算が通らない: ' + t.errors.join(','));
        const h = PAPER.build({ inv: { doc_type: 'invoice', no: 'X', issue_ymd: '2026-09-30', tax_mode: mode, data: {} }, tax: t, partner: { name: 'A' }, org: {} }).html;
        n++;
        const cnt = (h.match(/% 対象<\/th>/g) || []).length;
        if (cnt !== t.byRate.length) throw new Error('区分の数が違う: 紙' + cnt + ' / 計算' + t.byRate.length);
        if (!flat(h).includes(PAPER.yen(t.grandTotal))) throw new Error('合計が紙に無い: ' + t.grandTotal);
        if (/undefined|NaN|\[object/.test(h)) throw new Error('紙に undefined/NaN が出ている');
      }
    }
  }
  if (n < 30) throw new Error('組合せが少なすぎる（検査が空振り）: ' + n);
  console.log('     実測: ' + n + '通りを刷って矛盾0件');
});

console.log('\n' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
