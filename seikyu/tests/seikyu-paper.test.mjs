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
  // ★請求番号の欄（meta の1行目）だけを登録番号に差し替える（題名は触らない）
  if (kind === 'mixNo') return b.html.replace(/(請求番号<\/th><td>)[^<]*/, '$1T1234567890123');
  if (kind === 'flex') return b.html.replace('.note-b{display:block', '.note-b{display:flex');
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

  S('② 請求番号の欄に登録番号を入れた紙は「別物として出す」検査に落ちる', () => {
    const bad = paperBad('mixNo', sample());
    ok(/請求番号<\/th><td>T\d{13}/.test(bad.replace(/\s+/g, '')), '作り物が壊れていない＝この検査が空振り');
    const good = PAPER.build(sample()).html.replace(/\s+/g, '');
    ok(!/請求番号<\/th><td>T\d{13}/.test(good), '本物が登録番号を請求番号の欄に出している');
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

T('★登録番号(T+13桁)と請求番号は別の欄に出す', () => {
  const f = flat(H1);
  ok(/請求番号<\/th><td>202609-001/.test(f), '請求番号が出ていない');
  ok(!/請求番号<\/th><td>T\d{13}/.test(f), '請求番号の欄に登録番号が出ている');
  ok(/登録番号T1234567890123/.test(f), '自社の登録番号が出ていない');
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
  ok(/（未採番）/.test(h), '番号が空欄になっている');
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
  ok(/見 積 書/.test(h), '見出しが請求書のまま');
  ok(/お見積金額/.test(h), '金額の呼び方が請求書のまま');
  ok(/見積番号/.test(h) && /見積日/.test(h), '番号・日付の呼び方が請求書のまま');
});

T('金額は桁区切り・マイナスは頭に付く・数でない物は0にしない', () => {
  eq(PAPER.yen(142660), '142,660');
  eq(PAPER.yen(-142660), '-142,660');
  eq(PAPER.yen(0), '0');
  eq(PAPER.yen('abc'), '—', '数でない物が0になっている');
  eq(PAPER.jpDate('2026-09-30'), '2026年9月30日');
  eq(PAPER.jpDate('2026/09/30'), '', '読めない日付が通っている');
});

T('★文字はそのまま埋め込まない（HTMLとして壊れる／差し込まれる）', () => {
  const h = PAPER.build(sample({ partner: { name: '<script>alert(1)</script>' } })).html;
  ok(!/<script>alert/.test(h), '取引先名の中のタグがそのまま出ている');
  ok(/&lt;script&gt;/.test(h), 'エスケープされていない');
});

/* ── ⑤ 文が縦に割れない書き方 ─────────────────────────────────── */
const CSS = PAPER.css();

T('★紙のCSSに flex/grid を1つも使わない（文が1文字ずつ縦に割れる前科の形）', () => {
  ok(!/display\s*:\s*(flex|inline-flex|grid|inline-grid)/.test(CSS), 'flex/grid が使われている');
  ok(!/display\s*:\s*(flex|grid)/.test(H1), '紙の中の style に flex/grid がある');
});

T('★文が入る箱は「折り返し可」で「最低幅」を持つ', () => {
  // 長い日本語が入りうる箱＝ここが潰れると1文字ずつ縦になる
  ['.grand-n', '.note-b'].forEach((sel) => {
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
