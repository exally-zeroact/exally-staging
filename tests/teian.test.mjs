/* teian.test.mjs — ★8 提案＝うちから「こう直せますよ」と言う★
 *
 *  ★正本（2026-08-22 司さん決定）★ ★気づくのは機械。AIは説明するだけ★
 *  ★指示役の検証要件（2026-08-27）★
 *    ★提案が「向こうから」出る事（押していないのに出る）を 実際に押して確かめた証拠★
 *
 *  使い方: node tests/teian.test.mjs [--self-test]
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
const OVERRIDE = process.env.EXALLY_TEIAN_OVERRIDE ? JSON.parse(process.env.EXALLY_TEIAN_OVERRIDE) : {};
const require_ = createRequire(pathToFileURL(path.join(ROOT, 'package.json')));
const Teian = require_(OVERRIDE['lib/teian.js'] || path.join(ROOT, 'lib/teian.js'));
const GOLDEN = JSON.parse(fs.readFileSync(path.join(ROOT, 'tests/fixtures/shindan-golden.json'), 'utf8'));

let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ok   ' + n); } catch (e) { fail++; console.log('  NG   ' + n + '\n       ' + (e && e.message)); } };
const ok = (c, m) => { if (!c) throw new Error(m || 'expected truthy'); };
const eq = (a, b, m) => { if (a !== b) throw new Error((m ? m + ': ' : '') + '期待=' + JSON.stringify(b) + ' 実際=' + JSON.stringify(a)); };

const 本 = (式たち) => [{
  name: '計算',
  data: 式たち.reduce((a, f, i) => { a[i + ',0'] = { f: f }; return a; }, {}),
}];

console.log('');
console.log('[teian] ★うちから「こう直せますよ」と言う★');

/* ══ ①見つける（0円） ══ */
T('★INDEX＋MATCH を 見つける★', () => {
  const r = Teian.見つける(本(['=INDEX(B:B,MATCH(A1,C:C,0))']), { 版: 'excel_365' });
  eq(r.提案.length, 1);
  eq(r.提案[0].種類, 'indexmatch_to_xlookup');
  eq(r.提案[0].何本, 1);
});
T('★VLOOKUP を 見つける★', () => {
  const r = Teian.見つける(本(['=VLOOKUP(A1,B:D,3,0)']), { 版: 'excel_365' });
  eq(r.提案[0].種類, 'vlookup_to_xlookup');
});
T('★同じ探し物を 何度もしている式を 見つける★', () => {
  const f = '=IF(INDEX(計算!V2:V32, MATCH(B4, 計算!A2:A32, 0))<=0,"", INDEX(計算!E2:E32, MATCH(B4,計算!A2:A32, 0)))';
  eq(Teian.同じMATCHの数(f), 2, '★同じ MATCH を 数えていない★');
  const r = Teian.見つける(本([f]), { 版: 'excel_365' });
  ok(r.提案.some((t) => t.種類 === 'match_no_juufuku'), '★重複を 見つけていない★');
});
T('★MATCH が1回だけなら 出さない（要らない提案をしない）★', () => {
  const r = Teian.見つける(本(['=INDEX(B:B,MATCH(A1,C:C,0))']), { 版: 'excel_365' });
  ok(!r.提案.some((t) => t.種類 === 'match_no_juufuku'));
});
T('★IF の入れ子は「深さ」で数える（回数ではない）★', () => {
  eq(Teian.IFの深さ('=IF(A1,IF(B1,IF(C1,IF(D1,1,2),3),4),5)'), 4);
  eq(Teian.IFの深さ('=IF(A1,1,2)+IF(B1,3,4)'), 1, '★横に並んだIFを 入れ子と数えている★');
  eq(Teian.IFの深さ('=SUM(A1:B2)'), 0);
});
T('★浅いIFは 出さない（実物には 深い入れ子が 無かった）★', () => {
  const r = Teian.見つける(本(['=IF(A1,1,2)+IF(B1,3,4)']), { 版: 'excel_365' });
  ok(!r.提案.some((t) => t.種類 === 'if_nest'), '★無い物を 出している★');
});
T('★深いIFは 出す★', () => {
  const r = Teian.見つける(本(['=IF(A1,IF(B1,IF(C1,IF(D1,1,2),3),4),5)']), { 版: 'excel_365' });
  ok(r.提案.some((t) => t.種類 === 'if_nest'));
});
T('★式でない物（打った字）は 見ない★', () => {
  const r = Teian.見つける([{ name: 'あ', data: { '0,0': { v: 'INDEX(MATCH())' } } }], { 版: 'excel_365' });
  eq(r.提案.length, 0);
  eq(r.見た式, 0);
});
T('★多い順に 出す（効く所から）★', () => {
  const 式 = [];
  for (let i = 0; i < 5; i++) 式.push('=IF(INDEX(B:B,MATCH(A1,C:C,0))<=0,"",INDEX(D:D,MATCH(A1,C:C,0)))');
  式.push('=VLOOKUP(A1,B:D,3,0)');
  const r = Teian.見つける(本(式), { 版: 'excel_365' });
  ok(r.提案[0].何本 >= r.提案[r.提案.length - 1].何本, '★多い順に なっていない★');
});

/* ══ ②相手のExcelの版 ══ */
T('★2016には XLOOKUP を 勧めない（嘘の提案をしない）★', () => {
  const r = Teian.見つける(本(['=INDEX(B:B,MATCH(A1,C:C,0))']), { 版: 'excel_2016' });
  eq(r.提案[0].直し方, null, '★使えない直し方を 勧めている★');
  ok(r.提案[0].使えない理由.indexOf('使えません') > 0, '★なぜ勧めないかを 言っていない★');
});
T('★365/2021には 勧める★', () => {
  for (const 版 of ['excel_365', 'excel_2021']) {
    const r = Teian.見つける(本(['=INDEX(B:B,MATCH(A1,C:C,0))']), { 版 });
    ok(r.提案[0].直し方, '★' + 版 + 'で 勧めていない★');
  }
});
T('★版が 分からない時は 勧めない（使えない側に倒す）★', () => {
  eq(Teian.使えるか('XLOOKUP', 'なにか'), false);
});

/* ══ ③向こうから出す ══ */
T('★押していないのに 出す1行を 作る★', () => {
  const r = Teian.見つける(本(['=INDEX(B:B,MATCH(A1,C:C,0))']), { 版: 'excel_365' });
  const s = Teian.知らせの字(r);
  ok(s, '★知らせが 無い★');
  ok(s.本文.indexOf('AIは使っていません') > 0, '★0円だと 言っていない★');
  ok(s.つぎ.indexOf('決めてください') > 0, '★勝手に直すと 読める★');
});
T('★直し方が無い時は 向こうから 出さない（客が何も出来ない）★', () => {
  const r = Teian.見つける(本(['=INDEX(B:B,MATCH(A1,C:C,0))']), { 版: 'excel_2016' });
  eq(Teian.知らせの字(r), null, '★直せない事を こちらから 言っている★');
  eq(r.提案.length, 1, '★押して見た時にも 出なくなっている（隠している）★');
});
T('★1つも無ければ 知らせない（0件で騒がない）★', () => {
  eq(Teian.知らせの字(Teian.見つける(本(['=SUM(A1:B2)']), { 版: 'excel_365' })), null);
});
T('★客に見せる字に ★ を書かない★', () => {
  const r = Teian.見つける(本(['=INDEX(B:B,MATCH(A1,C:C,0))', '=IF(A1,IF(B1,IF(C1,IF(D1,1,2),3),4),5)']), { 版: 'excel_365' });
  const s = Teian.知らせの字(r);
  for (const t of [s.題, s.本文, s.つぎ]) ok(t.indexOf('★') < 0, '★客の字に ★ が出ている★：' + t);
  for (const t of r.提案) {
    for (const x of [t.題, t.本文, t.直し方 || '', t.使えない理由 || '']) {
      ok(x.indexOf('★') < 0, '★客の字に ★ が出ている★：' + x);
    }
  }
});
T('★お金の数字を 提案文に書かない★', () => {
  const r = Teian.見つける([{ name: 'あ', data: { '0,0': { f: '=INDEX(B:B,MATCH(A1,C:C,0))', v: 1234567 } } }], { 版: 'excel_365' });
  const 字 = JSON.stringify(r.提案[0].題) + r.提案[0].本文 + (r.提案[0].直し方 || '');
  ok(字.indexOf('1234567') < 0, '★金額が 提案文に 出ている★');
});
T('★場所を 出す（見に行けないと 意味がない）★', () => {
  const r = Teian.見つける(本(['=INDEX(B:B,MATCH(A1,C:C,0))']), { 版: 'excel_365' });
  eq(r.提案[0].場所[0], '計算!A1');
});

/* ══ ★④画面が「向こうから」出す（本物の book.html）★ ══ */
T('★画面が 開いた時に 自分で 提案を 見つけている（押していない）★', () => {
  const book = 注記を外す(fs.readFileSync(OVERRIDE['book.html'] || path.join(ROOT, 'book.html'), 'utf8'), { html: true });
  ok(/Teian\.見つける\(sheets,/.test(book), '★画面が 提案を 見つけていない★');
  /* ★「定義が在る」と「呼んでいる」は 別★＝定義だけでも 素通りした（2026-08-27）
     ⇒ ★開いた所（診断の隣）で 実際に 呼んでいるか★を見る */
  ok(book.indexOf('function 提案を始める(') > 0, '★提案を始める が 無い★');
  const i2 = book.indexOf('診断を始める();');
  ok(i2 > 0, '★開いた所が 見つからない★');
  ok(book.slice(i2, i2 + 300).indexOf('提案を始める();') > 0,
    '★開いた時に 自分で 始めていない（押されるまで 何もしない＝受け身）★');
  ok(/Teian\.知らせの字\(/.test(book), '★向こうから 出していない★');
  /* ★見つかった時だけ ボタンを出す★（出来ていない物のボタンを見せない） */
  ok(/<button id="teianBtn" hidden/.test(book), '★最初から ボタンが 出ている★');
  ok(/#teianBtn\[hidden\]\s*\{\s*display\s*:\s*none\s*!important/.test(book), '★[hidden] の1行が 無い★');
});
T('★提案から 勝手に直さない（押してから・差分プレビューを通す）★', () => {
  const book = 注記を外す(fs.readFileSync(OVERRIDE['book.html'] || path.join(ROOT, 'book.html'), 'utf8'), { html: true });
  const i = book.indexOf('function 提案を描く');
  ok(i > 0, '★組み立てる所が 無い★');
  const 所 = book.slice(i, i + 3000);
  for (const だめ of ['setCell(', 'saveXlsx(']) {
    ok(所.indexOf(だめ) < 0, '★提案の所から ' + だめ + ' を している（勝手に直している）★');
  }
});

/* ══ ★②VBAが要る事を レシピで済ませて 提案する（VBAは勧めない）★ ══ */
T('★マクロが入っていたら 代わりの道を こちらから 出す★', () => {
  const r = Teian.見つける(本(['=SUM(A1:A2)']), { 版: 'excel_365', マクロ: true });
  const t = r.提案.find((x) => x.種類 === 'vba_to_recipe');
  ok(t, '★マクロが入っているのに 何も言わない★');
  ok(t.直し方, '★代わりに 何が出来るかを 言っていない★');
  ok(t.本文.indexOf('動きません') > 0, '★ここでは動かない事を 言っていない★');
});
T('★VBAを 勧めない（マクロを書けとは 言わない）★', () => {
  const r = Teian.見つける(本(['=SUM(A1:A2)']), { 版: 'excel_365', マクロ: true });
  const 字 = r.提案.map((t) => t.題 + t.本文 + (t.直し方 || '')).join('');
  for (const だめ of ['VBAを書', 'マクロを作', 'マクロを書']) {
    ok(字.indexOf(だめ) < 0, '★VBAを 勧めている：' + だめ);
  }
});
T('★マクロが無ければ 出さない（無い物を 言わない）★', () => {
  const r = Teian.見つける(本(['=SUM(A1:A2)']), { 版: 'excel_365' });
  eq(r.提案.length, 0);
  eq(Teian.知らせの字(r), null);
});
T('★マクロを「か所」に 混ぜて数えない（嘘の数を 出さない）★', () => {
  const r = Teian.見つける(本(['=INDEX(B:B,MATCH(A1,C:C,0))']), { 版: 'excel_365', マクロ: true });
  const s = Teian.知らせの字(r);
  ok(s.本文.indexOf('1通り・1か所') > 0, '★式のぶんの数が 変わっている：' + s.本文);
  ok(s.本文.indexOf('2通り') < 0, '★マクロを 1か所 と数えている★');
  ok(s.本文.indexOf('マクロ（VBA）でしている事も') > 0, '★マクロの話が 消えている★');
});
T('★式の提案が 1つも無くても マクロだけで 知らせる★', () => {
  const r = Teian.見つける(本(['=SUM(A1:A2)']), { 版: 'excel_365', マクロ: true });
  const s = Teian.知らせの字(r);
  ok(s, '★マクロだけの時に 何も言わない★');
  ok(s.本文.indexOf('式を') < 0, '★0通り・0か所 と 出している★');
  ok(s.本文.indexOf('AIは使っていません') > 0, '★0円だと 言っていない★');
});
T('★画面が マクロの有無を 渡している（渡さないと 一生 出ない）★', () => {
  const book = 注記を外す(fs.readFileSync(OVERRIDE['book.html'] || path.join(ROOT, 'book.html'), 'utf8'), { html: true });
  ok(/Teian\.見つける\(sheets, \{ 版: 版, マクロ: !!_bookHasVba \}\)/.test(book), '★マクロの有無を 渡していない★');
  ok(book.indexOf('_bookHasVba = !!res.hasVba;') > 0, '★開いた時に 覚えていない★');
});

T('★マクロだけの時に ボタンへ「0か所」と 出さない★', () => {
  const book = 注記を外す(fs.readFileSync(OVERRIDE['book.html'] || path.join(ROOT, 'book.html'), 'utf8'), { html: true });
  /* ★実際に .xlsm を開いて 見つけた（2026-08-27）★＝「短くできる 0か所」と出ていた。 */
  ok(book.indexOf("if(t.直し方 && t.数える物か) 合計 += t.何本;") > 0, '★数えない物まで 数えている★');
  ok(/合計 \? \('短くできる ' \+ 合計 \+ 'か所'\) : 'こちらで 出来る事'/.test(book), '★0か所と 出している★');
  ok(book.indexOf("t.題 + (t.数える物か ? ('（' + t.何本 + 'か所）') : '')") > 0, '★一覧の見出しに 0か所 と出ている★');
});

/* ══ ★開いた時の知らせは 1つだけ（本物の画面の字を そのまま動かす）★ ══
   ★実際に 司さんの本を開いて 測って 見つけた（2026-08-27）★
     ＝網・提案・診断が それぞれ知らせを出し、後の物が 前の物を 消していた。
       画面に残っていたのは 診断だけ＝★提案の知らせは 一度も 客に届いていなかった★。 */
const 知らせの所 = () => {
  const b = fs.readFileSync(OVERRIDE['book.html'] || path.join(ROOT, 'book.html'), 'utf8');
  const i = b.indexOf('var _開いた知らせ = null;');
  const j = b.indexOf('/* ══ ★8-③ 覚えた手順（レシピ）を 画面につなぐ★');
  if (i < 0 || j < 0 || j < i) throw new Error('★ためる所が 無い（知らせが ぶつかる作りに戻っている）★');
  return b.slice(i, j);
};
const 動かす = () => {
  const 出た = [];
  const f = new Function('showToast', 知らせの所() + ';return{始:開いた知らせを始める,足:開いた知らせに足す,済:開いた知らせは済んだ};');
  return { api: f((h, ms, kind) => 出た.push({ 字: h, 秒: ms, 種: kind })), 出た };
};

T('★3つ分 出そろってから 1回だけ 出す（前の知らせを 消さない）★', () => {
  const { api, 出た } = 動かす();
  api.始(['網', '提案', '診断']);
  api.足('あみ', 'アミの中身'); api.済('網');
  eq(出た.length, 0, '★まだ待っている物が在るのに 出した★');
  api.足('ていあん', 'テイアンの中身'); api.済('提案');
  eq(出た.length, 0, '★診断を待たずに 出した★');
  api.足('しんだん', 'シンダンの中身', true); api.済('診断');
  eq(出た.length, 1, '★知らせが 1つでない（＝どれかが 消える）★');
  for (const 語 of ['アミの中身', 'テイアンの中身', 'シンダンの中身']) {
    ok(出た[0].字.indexOf(語) > 0, '★' + 語 + ' が 消えている★');
  }
  eq(出た[0].種, 'warn', '★危ない話が 普通の知らせで 出ている（読みにくい）★');
});
T('★同じ物が 2回 済んでも 早く出さない★', () => {
  const { api, 出た } = 動かす();
  api.始(['網', '提案', '診断']);
  api.足('あ', 'ア'); api.済('網'); api.済('網'); api.済('網');
  eq(出た.length, 0, '★1つの物の2回めで 出てしまった★');
});
T('★見つからなければ 何も出さない（0件で騒がない）★', () => {
  const { api, 出た } = 動かす();
  api.始(['網', '提案', '診断']);
  api.済('網'); api.済('提案'); api.済('診断');
  eq(出た.length, 0, '★空の知らせを 出している★');
});
T('★開く時 以外は そのまま出す（ためたまま 消えない）★', () => {
  const { api, 出た } = 動かす();
  api.足('あ', 'ア');
  eq(出た.length, 1, '★ためる所が 無い時に 知らせが 消えている★');
});
T('★開いた所で 3つとも 待っている（数え忘れは 出ない事故になる）★', () => {
  const book = 注記を外す(fs.readFileSync(OVERRIDE['book.html'] || path.join(ROOT, 'book.html'), 'utf8'), { html: true });
  ok(/開いた知らせを始める\(\['開いた', '網', '提案', '診断'\]\)/.test(book), '★待つ物の数が 合っていない★');
  for (const 名 of ['開いた', '網', '提案', '診断']) {
    ok(book.indexOf("開いた知らせは済んだ('" + 名 + "')") > 0, '★' + 名 + ' が 済んだと 言わない（永久に 出ない）★');
  }
  /* ★「どこかに1つ在る」では 足りない★＝見つかった時の道（本番で通る道）で 言っているか。
     ＝実際に消してみたら 素通りした（2026-08-27）。上の入口の分が 残るだけだった。 */
  const i3 = book.indexOf('if(r.done){');
  ok(i3 > 0, '★診断の 終わりが 無い★');
  const 終 = book.slice(i3, book.indexOf('      return;', i3));
  ok(終.indexOf("開いた知らせは済んだ('診断')") > 0,
    '★見つかった時に 済んだと 言っていない（提案の知らせが 永久に 出ない）★');
  /* ★「開きました：…」も 同じ1つの知らせに入れる★＝0.1秒で 調べた結果に 消されていた */
  ok(book.indexOf("開いた知らせに足す('開きました：' + file.name") > 0,
    '★開きましたを 別の知らせで 出している（すぐ 消される）★');
  ok(book.indexOf("if(!g){ 開いた知らせは済んだ('網'); 開いた知らせは済んだ('提案'); 開いた知らせは済んだ('診断'); return; }") > 0,
    '★網が 作れない時に 待ったままになる（永久に 何も出ない）★');
  const i4 = book.indexOf('開いた知らせに足す(知.題');
  ok(i4 > 0, '★提案が ためていない★');
  ok(book.slice(i4, i4 + 400).indexOf("開いた知らせは済んだ('提案')") > 0,
    '★出した後に 済んだと 言っていない★');
  /* ★出口を 増やさせない★＝ここで showToast を直に呼ぶと また ぶつかる */
  for (const 名 of ['function 提案を始める', 'function 診断の知らせを出す']) {
    const i = book.indexOf(名);
    ok(i > 0, '★' + 名 + ' が 無い★');
    const 所 = book.slice(i, book.indexOf('{', i) + 2000);
    const 端 = 所.indexOf(String.fromCharCode(10) + '}');
    ok((端 > 0 ? 所.slice(0, 端) : 所).indexOf('showToast(') < 0,
      '★' + 名 + ' が 知らせを 直に出している（後の物に 消される）★');
  }
});

/* ══ ★⑤実物（司さんの .xlsb）★ ══ */
const 実物 = GOLDEN.本.場所;
if (!fs.existsSync(実物)) {
  console.log('  ★未測定★ 実物が無い機械です（0件・異常なしにしない）');
} else {
  const XLSX = require_(path.join(ROOT, 'lib/xlsx.full.min.js'));
  const ZipSurgeon = require_(path.join(ROOT, 'lib/zip-surgeon.js'));
  const TableRefs = require_(path.join(ROOT, 'lib/table-refs.js'));
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
    return { name, data };
  });
  const r = Teian.見つける(sheets, { 版: 'excel_365' });
  T('★実物で 提案が出る（式15,126本・INDEX+MATCH 6,832か所・重複 2,548か所）★', () => {
    eq(r.見た式, 15126);
    const im = r.提案.find((t) => t.種類 === 'indexmatch_to_xlookup');
    const ju = r.提案.find((t) => t.種類 === 'match_no_juufuku');
    ok(im && im.何本 === 6832, '★INDEX+MATCH の数が 違う：' + (im && im.何本) + '★');
    ok(ju && ju.何本 === 2548, '★重複の数が 違う：' + (ju && ju.何本) + '★');
    ok(r.かかった秒 <= 2, '★' + r.かかった秒 + '秒（遅すぎる）★');
    console.log('       … 式 ' + r.見た式 + '本を ' + r.かかった秒 + '秒で見て '
      + r.提案.reduce((a, t) => a + t.何本, 0) + 'か所（AIは0回）');
  });
  T('★実物に 深いIFの入れ子は 無い（無い物を 出さない）★', () => {
    ok(!r.提案.some((t) => t.種類 === 'if_nest'), '★無い物を 出している★');
  });
  T('★実物でも 提案文に 金額が 出ない★', () => {
    const 字 = r.提案.map((t) => t.題 + t.本文 + (t.直し方 || '')).join('');
    eq((字.match(/\d{6,}/g) || []).length, 0, '★金額らしき数が 出ている★');
  });
}

console.log('');
console.log('  ' + pass + ' 緑 / ' + fail + ' 赤');

if (SELF) {
  const { spawnSync } = await import('node:child_process');
  const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'exally-teian-'));
  console.log('');
  console.log('[self-test] わざと壊して 赤くなるかを数える（★repo は読むだけ★）');
  const BREAKS = [
    ['lib/teian.js', '★INDEX+MATCH を 見つけない★',
      (s) => s.replace("} else if (/\\bINDEX\\s*\\(/i.test(f) && /\\bMATCH\\s*\\(/i.test(f)) {", '} else if (false) {')],
    ['lib/teian.js', '★同じ探し物の重複を 見つけない★',
      (s) => s.replace('if (同じMATCHの数(f) >= 2) {', 'if (false) {')],
    ['lib/teian.js', '★IFの入れ子を「回数」で数える（横に並んだIFも 入れ子にする）★',
      (s) => s.replace('if (積み.pop()) 深さ--;', '')],
    ['lib/teian.js', '★2016にも XLOOKUP を 勧める（嘘の提案）★',
      (s) => s.replace('return 悪い.length === 0;', 'return true;')],
    ['lib/teian.js', '★版が分からなくても 勧める★',
      (s) => s.replace('} catch (e) { return false; }', '} catch (e) { return true; }')],
    ['lib/teian.js', '★直し方が無いのに 向こうから 出す★',
      (s) => s.replace('if (!出せる.length) return null;', '')],
    ['lib/teian.js', '★0件でも 知らせる★',
      (s) => s.replace('if (!結果 || !結果.提案 || !結果.提案.length) return null;', '')],
    ['lib/teian.js', '★0円だと 言わない★', (s) => s.replace("本文: 文 + 'AIは使っていません。',", '本文: 文,')],
    ['lib/teian.js', '★場所を 出さない★', (s) => s.replace('場所: 例.map(function (x) { return x.場所; }),', '場所: [],')],
    ['lib/teian.js', '★客の字に ★ を書く★', (s) => s.replace("題: 数える ? 'もっと短く書ける所が あります'", "題: 数える ? '★もっと短く書ける所が あります★'")],
    ['book.html', '★開いた時に 自分で 始めない（押されるまで 何もしない）★',
      (s) => s.replace('提案を始める();', '')],
    ['book.html', '★最初から ボタンを 出す★', (s) => s.replace('<button id="teianBtn" hidden', '<button id="teianBtn"')],
    /* ★知らせの出口を 1つに保つ（実際に開いて 見つけた事故 2026-08-27）★ */
    ['book.html', '★提案が 知らせを 直に出す（後の物に 消される）★',
      (s) => s.replace('開いた知らせに足す(知.題, 知.本文', "showToast('<span class=\"toast-h\">' + 知.題 + '</span>' + 知.本文")],
    ['book.html', '★診断が 済んだと 言わない（永久に 出ない）★',
      (s) => s.replace("      開いた知らせは済んだ('診断');" + String.fromCharCode(10) + '      return;', '      return;')],
    ['book.html', '★出そろう前に 出す（前の知らせを 消す）★',
      (s) => s.replace('for(var k in _開いた知らせ.待ち){ if(Object.prototype.hasOwnProperty.call(_開いた知らせ.待ち, k)) return; }', '')],
    ['book.html', '★待つ物を 数え違える★',
      (s) => s.replace("開いた知らせを始める(['開いた', '網', '提案', '診断']);", "開いた知らせを始める(['網']);")],
    ['book.html', '★「開きました」を 別の知らせで 出す（0.1秒で 消される）★',
      (s) => s.replace("開いた知らせに足す('開きました：' + file.name", "showToast('開きました：' + file.name")],
    ['book.html', '★網が 作れなかった時に 待ったまま（永久に 何も出ない）★',
      (s) => s.replace("if(!g){ 開いた知らせは済んだ('網'); 開いた知らせは済んだ('提案'); 開いた知らせは済んだ('診断'); return; }", 'if(!g){ return; }')],
    /* ★②VBAを レシピで済ませる提案★ */
    ['lib/teian.js', '★マクロが入っていても 何も言わない★',
      (s2) => s2.replace("if (opt.マクロ) 提案.push(作る('vba_to_recipe', 0, [], true, 版));", '')],
    ['lib/teian.js', '★マクロを「か所」に 混ぜて数える（嘘の数）★',
      (s2) => s2.replace('if (出せる[i].数える物か) { 合計 += 出せる[i].何本; 数える++; }', '{ 合計 += 出せる[i].何本; 数える++; }')],
    ['book.html', '★マクロの有無を 渡さない（一生 出ない）★',
      (s2) => s2.replace('{ 版: 版, マクロ: !!_bookHasVba }', '{ 版: 版 }')],
    ['book.html', '★マクロだけの時に ボタンへ 0か所 と出す★',
      (s2) => s2.replace("b.textContent = 合計 ? ('短くできる ' + 合計 + 'か所') : 'こちらで 出来る事';", "b.textContent = '短くできる ' + 合計 + 'か所';")],
    ['book.html', '★一覧の見出しに 0か所 と出す★',
      (s2) => s2.replace("頭.textContent = t.題 + (t.数える物か ? ('（' + t.何本 + 'か所）') : '');", "頭.textContent = t.題 + '（' + t.何本 + 'か所）';")],
  ];
  let red = 0;
  for (const [rel, name, brk] of BREAKS) {
    const orig = fs.readFileSync(path.join(ROOT, rel), 'utf8');
    const bad = brk(orig);
    if (bad === orig) { console.log('  ★置換できず★  ' + name); continue; }
    const f = path.join(TMP, rel.replace(/[\\/]/g, '_'));
    fs.writeFileSync(f, bad, 'utf8');
    const env = Object.assign({}, process.env, { EXALLY_TEIAN_OVERRIDE: JSON.stringify({ [rel]: f }) });
    const r2 = spawnSync(process.execPath, ['--max-old-space-size=4096', path.join(__dirname, 'teian.test.mjs')], { encoding: 'utf8', env });
    if (r2.status !== 0) { red++; console.log('  赤くなった  ' + name); }
    else console.log('  ★素通り★  ' + name);
  }
  try { fs.rmSync(TMP, { recursive: true, force: true }); } catch (e) { /* 消せなくても検査は済んでいる */ }
  console.log('');
  console.log('  ' + red + '/' + BREAKS.length + ' 通りで赤くなった');
  process.exit(red === BREAKS.length ? 0 : 1);
}

process.exit(fail ? 1 : 0);
