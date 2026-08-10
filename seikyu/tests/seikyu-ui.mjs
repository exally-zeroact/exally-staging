/* seikyu-ui.mjs — ★実UI 全ボタン検証（本物の seikyu/index.html と app.js を動かす）★
 *
 * 何をするか:
 *   本物の画面を jsdom に読み込み、★画面にあるボタンを1つ残らず押す★。
 *   そのうえで「1通 出す」までの筋を実際に通す（作る → 発行 → 紙 → Excel）。
 *   倉庫は偽物（tests/fake-supa.js）だが、★番号の一意制約は本番と同じ形で再現する★
 *   ＝「同じ番号を二度使わない」を倉庫が止めることを、倉庫に触らずに測れる。
 *
 * ここで止めたい事故:
 *   ① どこかのボタンで JS が落ちる（押した人には「無反応」に見える）
 *   ② 発行済みなのに直せる／もう一度発行できる
 *   ③ 番号がぶつかった時に黙って上書きする
 *   ④ 落とす前にファイル名を見せない／名前が中身と無関係
 *   ⑤ 印刷の窓にアプリの画面が混ざる
 *   ⑥ 入金が読めない時に「未入金(0円)」と言い切る
 *   ⑦ 取引先を保存した時に、ハブが入れた名前や住所を消す
 *
 * 依存: jsdom。★入っていなければ赤（SKIPを緑と呼ばない）★
 * 使い方: node seikyu/tests/seikyu-ui.mjs
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(__dirname, '..', '..');
const require_ = createRequire(path.join(ROOT, 'package.json'));
let JSDOM;
try { ({ JSDOM } = await import('jsdom')); }
catch { console.log('★jsdomが入っていません。この検証は飛ばせません（SKIPを緑と呼ばない）。npm install してください。'); process.exit(1); }
const { createFakeSupa } = require_(path.join(ROOT, 'tests/fake-supa.js'));

let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ✓ ' + n); } catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + (e && e.message)); } };
const TA = async (n, fn) => { try { await fn(); pass++; console.log('  ✓ ' + n); } catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + (e && e.message)); } };
const ok = (v, m) => { if (!v) throw new Error(m || 'false'); };
const eq = (a, b, m) => { if (a !== b) throw new Error((m ? m + ': ' : '') + 'expected ' + JSON.stringify(b) + ' got ' + JSON.stringify(a)); };
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

/* ── 本物の画面を読む（CDN・接続設定・ログインは外す＝ネットに出ない） ── */
const html = fs.readFileSync(path.join(ROOT, 'seikyu/index.html'), 'utf8');
const srcs = [...html.matchAll(/<script src="([^"]+)"><\/script>/g)].map((m) => m[1].split('?')[0])
  .filter((s) => !/^https?:/.test(s) && !/supa-config|auth\.js|exally-login/.test(s));
const dom = new JSDOM(html.replace(/<script[\s\S]*?<\/script>/g, ''), { runScripts: 'dangerously', url: 'http://localhost/seikyu/', pretendToBeVisual: true });
const win = dom.window, doc = win.document;

const errs = [];
win.addEventListener('error', (e) => errs.push('window.error: ' + (e.message || e)));
win.addEventListener('unhandledrejection', (e) => errs.push('unhandledrejection: ' + ((e.reason && e.reason.message) || e.reason)));
win.fetch = () => Promise.reject(new Error('no net'));
win.confirm = () => true;
win.scrollTo = () => {};
win.print = () => {};

/* 新しい窓（印刷）を捕まえる。★実際に開かず、書かれた中身を測る★ */
const opened = [];
win.open = function () {
  const w = {
    _html: '', document: {
      open() {}, write(s) { w._html += s; }, close() {}, readyState: 'complete', title: '',
    },
    addEventListener() {}, focus() {}, print() { w._printed = true; }, _printed: false,
  };
  opened.push(w);
  return w;
};
/* 落とす口（file-out.js）を本物のまま動かし、出来た物を捕まえる */
const delivered = [];
win.URL.createObjectURL = function (b) { delivered.push({ blob: b, type: b && b.type }); return 'blob:test/' + delivered.length; };
win.URL.revokeObjectURL = function () {};
const anchorClicks = [];
win.HTMLAnchorElement.prototype.click = function () {
  anchorClicks.push({ href: this.href, download: this.getAttribute('download'), target: this.getAttribute('target'), rel: this.getAttribute('rel') });
};

for (const src of srcs) {
  const el = doc.createElement('script');
  el.textContent = fs.readFileSync(path.join(ROOT, 'seikyu', src.replace(/^\.\.\//, '../')), 'utf8');
  doc.body.appendChild(el);
}
ok(win.SeikyuApp, 'SeikyuApp が露出していない（読み込みに失敗）');

/* ── 偽の倉庫（本番と同じ「二度使えない組」つき） ── */
function makeSb() {
  return createFakeSupa({
    uid: 'u1',
    tables: {
      pay_org: [{ account_id: 'u1', data: { yago: '株式会社ゼロアクト', addr: '愛媛県今治市4-5-6', tel: '0898-00-0000', invoiceNo: 'T1234567890123' }, updated_at: '2026-08-01T00:00:00Z' }],
      pay_partners: [
        { id: 'pt_a', account_id: 'u1', sort: 0, data: { name: '藤原建設株式会社', keisho: '御中', addr: '愛媛県今治市1-2-3', invoiceNo: 'T9876543210987' }, deleted_at: null },
        { id: 'pt_b', account_id: 'u1', sort: 1, data: { name: '株式会社しまなみ', keisho: '様', addr: '松山市1-1' }, deleted_at: null },
      ],
      pay_invoices: [],
      pay_receipts: [],
      pay_companies: [{ account_id: 'u1', data: {}, updated_at: '2026-08-01T00:00:00Z' }],
    },
    pk: { pay_org: 'account_id', pay_companies: 'account_id' },
    // ★本番の uq_pay_invoices_no と同じ組（where を付けない＝取り消した番号も再利用不可）
    unique: { pay_invoices: [['account_id', 'doc_type', 'no']] },
  });
}
const sb = makeSb();
const db = sb._db;

const $ = (id) => doc.getElementById(id);
const qa = (s) => [...doc.querySelectorAll(s)];
const setVal = (id, v) => { const e = $(id); e.value = v; e.dispatchEvent(new win.Event('input')); e.dispatchEvent(new win.Event('change')); };

console.log('\n[請求書 実UI 全ボタン検証]');

await win.SeikyuApp.attach(sb);
await sleep(20);

/* ═══ 0. 出発点 ═══ */
T('0. ★中身(#app)は最初 hidden＝未ログインで画面を見せない', () => {
  const raw = fs.readFileSync(path.join(ROOT, 'seikyu/index.html'), 'utf8');
  ok(/<div id="app" hidden>/.test(raw), 'index.html の #app に hidden が無い');
  const authSrc = fs.readFileSync(path.join(ROOT, 'seikyu/js/auth.js'), 'utf8');
  ok(/a\.hidden = false/.test(authSrc), 'ログインできた時に hidden を外していない');
  ok(/a\.hidden = true/.test(authSrc), 'ログイン画面に戻す時に hidden を付けていない');
});

T('0. 取引先と自社が共有マスタから読めている（請求書で別に持っていない）', () => {
  eq(win.SeikyuApp._state.partners.length, 2);
  eq(win.SeikyuApp._state.org.yago, '株式会社ゼロアクト');
});

/* ═══ 1. 画面にあるボタンを1つ残らず押す ═══
   ★押した数を報告するのではなく、押す物の一覧をここで作って全部通す★ */
await TA('1. ★3画面ぜんぶのボタンを1つ残らず押しても、JSが1つも落ちない', async () => {
  const before = errs.length;
  const screens = ['scr-list', 'scr-edit', 'scr-set'];
  const pressed = [];
  for (const scr of screens) {
    const nav = doc.querySelector('.ex-bn[data-scr="' + scr + '"]');
    ok(nav, 'ナビ ' + scr + ' が無い');
    nav.click();
    await sleep(10);
    const el = $(scr);
    ok(el.classList.contains('active'), scr + ' が開かない');
    ok(el.innerHTML.length > 400, scr + ' の中身が薄い(' + el.innerHTML.length + ')');
    for (const b of [...el.querySelectorAll('button'), ...doc.querySelectorAll('.ex-bn')]) {
      if (b.disabled) { pressed.push((b.id || b.textContent.trim()) + '(押せない)'); continue; }
      b.click();
      await sleep(6);
      pressed.push(b.id || b.getAttribute('data-fil') || b.getAttribute('data-tm') || b.getAttribute('data-nm') || b.textContent.trim());
    }
    // 押した拍子に別の画面へ行っていたら戻す
    doc.querySelector('.ex-bn[data-scr="' + scr + '"]').click();
    await sleep(6);
  }
  // ファイル名の小窓のボタンも押す
  $('fn-cancel').click();
  pressed.push('fn-cancel');
  console.log('     押した物(' + pressed.length + '): ' + pressed.join(' / '));
  ok(pressed.length >= 25, '押した物が少なすぎる（一覧が取れていない）: ' + pressed.length);
  eq(errs.length, before, 'JSが落ちた: ' + errs.slice(before).join(' | '));
});

/* ═══ 1-b. 見た目の土台（スイート共通の皮）と、潰れない書き方 ═══
   jsdom は幅を計算しないので、ここでは ★潰れない書き方になっているか★ を見る
   （実物の幅は実機幅の画面で定規を当てて確かめる。この検査はその前段の網）。 */
const SKIN = fs.readFileSync(path.join(ROOT, 'css/exally-ui.css'), 'utf8');
const APPCSS = fs.readFileSync(path.join(ROOT, 'seikyu/css/app.css'), 'utf8');
const CSS = SKIN + '\n' + APPCSS;
/* 色の検査は「実際に効いている指定」だけを見る（説明文の中の色名を数えない） */
const CSS_CODE = CSS.replace(/\/\*[\s\S]*?\*\//g, ' ');

T('1-b. ★見た目はスイート共通の皮を読んでいる（請求書だけ別の画面にしない）', () => {
  ok(/<link rel="stylesheet" href="\.\.\/css\/exally-ui\.css/.test(html), '共通の皮を読んでいない');
  ok(html.indexOf('exally-ui.css') < html.indexOf('css/app.css'), '皮より先にアプリのCSSを読んでいる（差分が効かない）');
});

T('1-b. ★うちのミント #52B788 と 差し色 #3D9E72 が実際に効いている（請求書だけ別の緑にしない）', () => {
  // ★選ばれている物に見た目が付いているか（押しても何も変わらないように見せない）
  ok(/\.ex-chip\.on\s*\{[^}]*background/.test(CSS_CODE), '選んだチップに色が付かない');
  ok(/\.ex-mini\.on\s*\{[^}]*background/.test(CSS_CODE), '選んだ小さいボタン（揃えなど）に色が付かない');
  ok(/#52B788/i.test(CSS_CODE), 'ブランドのミントが1回も使われていない');
  ok(/#3D9E72/i.test(CSS_CODE), '差し色が1回も使われていない');
  ok(/#2E7D54/i.test(CSS_CODE), '主色が使われていない');
  ok(!/#1A4A2E/i.test(CSS_CODE), '使ってはいけない濃い緑がある');
});

T('1-b. ★皮に無い緑を勝手に足していない（3アプリでバラけた原因）', () => {
  const allowed = new Set(['#2e7d54', '#3d9e72', '#52b788', '#3d6b53', '#5c7e6c', '#7aa08c',
    '#d4eae0', '#c8ecd8', '#f0faf4', '#e8f6ee', '#ffffff', '#c0392b', '#f0d5d0', '#fdf0ee',
    '#92500a', '#f0ddbc', '#fdf3e3']);
  const used = [...APPCSS.replace(/\/\*[\s\S]*?\*\//g, ' ').matchAll(/#[0-9a-fA-F]{6}\b/g)]
    .map((m) => m[0].toLowerCase());
  const stray = [...new Set(used)].filter((c) => !allowed.has(c));
  ok(stray.length === 0, '皮に無い色を使っている: ' + stray.join(', '));
});

T('1-b. ★明細の表は「縮めて潰す」のではなく「横に動かす」（実機幅375pxで欄が幅ゼロになった前科）', () => {
  const wrap = (/\.lines-scroll\s*\{([^}]*)\}/.exec(CSS) || [])[1] || '';
  ok(/overflow-x\s*:\s*auto/.test(wrap), '.lines-scroll が横に動かせない');
  // ★列は会社が足せるので、表そのものではなく「1列ぶんの最低幅」で潰れを止める
  for (const sel of ['.l-name', '.l-sm', '.l-md', '.l-x']) {
    const rule = (new RegExp('\\' + sel + '\\s*\\{([^}]*)\\}').exec(CSS) || [])[1] || '';
    ok(/min-width\s*:\s*\d+px/.test(rule), sel + ' に最低幅が無い（列が幅ゼロまで潰れる）');
  }
  ok(!/\.lines\s*\{[^}]*width\s*:\s*100%/.test(CSS), '.lines に width:100% がある（最低幅を打ち消す）');
  ok(/<div class="lines-scroll">/.test(html), '表が横に動く入れ物に入っていない');
});

T('1-b. ★文が入る箱は block で最低幅を持ち、日本語を1文字ずつ割らない', () => {
  ok(!/word-break\s*:\s*break-all/.test(CSS), 'break-all がある（日本語が1文字ずつ割れる）');
  // 皮の側（4つまとめて指定している）
  const many = (/\.ex-hint,\s*\.ex-warn,\s*\.ex-bad,\s*\.ex-ok,\s*\.ex-why\s*\{([^}]*)\}/.exec(CSS) || [])[1] || '';
  ok(/display\s*:\s*block/.test(many), '文の箱が block でない');
  ok(/min-width\s*:\s*\d/.test(many), '文の箱に最低幅が無い');
  ok(/overflow-wrap\s*:\s*break-word/.test(many), '文の箱に折り返しの指定が無い');
  // 請求書だけの物
  for (const sel of ['.scroll-note', '.iv-sub']) {
    const rule = (new RegExp('\\' + sel + '\\s*\\{([^}]*)\\}').exec(CSS) || [])[1] || '';
    ok(/display\s*:\s*block/.test(rule), sel + ' が block でない');
    ok(/min-width\s*:\s*\d|width\s*:\s*100%/.test(rule), sel + ' に幅の確保が無い');
    ok(/overflow-wrap\s*:\s*break-word/.test(rule), sel + ' に折り返しの指定が無い');
  }
  ok(/\.ex-btn-row\s*\{[^}]*flex-wrap\s*:\s*wrap/.test(CSS), 'ボタンの行が折り返さない（横にはみ出す）');
  // ★上の帯は flex。中の日本語（アプリ名）が縮んで1文字ずつ縦に割れた前科（実機幅390px）
  for (const sel of ['.ex-logo', '.ex-back']) {
    const rule = (new RegExp('\\' + sel + '\\s*\\{([^}]*)\\}').exec(CSS) || [])[1] || '';
    ok(/white-space\s*:\s*nowrap/.test(rule), sel + ' が折り返す（日本語のアプリ名が縦に割れる）');
    ok(/flex\s*:\s*0 0 auto/.test(rule), sel + ' が縮む指定になっている（flexの子は既定で縮む）');
  }
  const who = (/\.ex-who\s*\{([^}]*)\}/.exec(CSS) || [])[1] || '';
  ok(/text-overflow\s*:\s*ellipsis/.test(who), '長いメールが「…」で切られない（他を押し出す）');
  // ★列の編集は flex の行。中の「列の名前」が縦帯にならないよう先に幅を確保している
  const cn = (/\.col-name\s*\{([^}]*)\}/.exec(CSS) || [])[1] || '';
  ok(/min-width\s*:\s*\d+px/.test(cn), '列の名前に最低幅が無い（flexの中で1文字ずつ縦に割れる）');
});

T('1-b. ★入力欄は16px（これより小さいと iPhone が勝手に拡大して画面がズレる）', () => {
  const rule = (/\.ex-input\s*\{([^}]*)\}/.exec(CSS) || [])[1] || '';
  const m = /font-size\s*:\s*(\d+(?:\.\d+)?)px/.exec(rule);
  ok(m, '.ex-input に文字の大きさが無い');
  ok(Number(m[1]) >= 16, '入力欄が ' + m[1] + 'px（16px 未満）');
});

T('1-b. ★iPhone で持っていない入力（月・週）を使っていない', () => {
  ok(!/type="(month|week|datetime-local)"/.test(html), 'iOS が持っていない入力がある');
});

/* ═══ 2. 1通 作って発行する ═══ */
// 1. の総当たりで様式が切り替わっているので、ここで std1 に戻してから測る
doc.querySelector('#s-tpl [data-tpl="std1"]').click();
await sleep(10);
$('b-new').click();
await sleep(10);

await TA('2. ★新しく作る＝白紙を埋めさせない（今日・既定の税・番号が最初から入る）', async () => {
  ok($('e-issue').value, '請求日が空');
  ok($('e-no').value, '番号が空: ' + $('e-no').value);
  ok(qa('#e-taxmode .ex-chip.on').length === 1, '税の入れ方が選ばれていない');
  ok($('lines-body').querySelectorAll('tr').length >= 1, '明細の行が無い');
});

await TA('2. 取引先を選ぶ・明細を入れる → 合計が実額で出る', async () => {
  setVal('e-partner', 'pt_a');
  await sleep(10);
  setVal('e-issue', '2026-09-30');
  await sleep(10);
  const tr = $('lines-body').querySelector('tr');
  const f = (name) => tr.querySelector('[data-f="' + name + '"]');
  f('name').value = '運転代行 9月分'; f('name').dispatchEvent(new win.Event('input'));
  f('qty').value = '3'; f('qty').dispatchEvent(new win.Event('input'));
  f('price').value = '105'; f('price').dispatchEvent(new win.Event('input'));
  await sleep(10);
  const tot = $('tot-box').textContent;
  // ★実測した境界：外税・切り捨て・3行ぶんの税抜105 → 消費税は 31円（行ごとに丸めると30円）
  ok(/31/.test(tot), '税率ごとに1回だけ丸めた額(31)が出ていない: ' + tot.replace(/\s+/g, ' '));
  ok(/346/.test(tot), '合計(346)が出ていない: ' + tot.replace(/\s+/g, ' '));
});

await TA('2. 行を足す・行を消す が効く', async () => {
  const n0 = $('lines-body').querySelectorAll('tr').length;
  $('b-addline').click(); await sleep(6);
  eq($('lines-body').querySelectorAll('tr').length, n0 + 1, '行が増えない');
  $('lines-body').querySelectorAll('[data-del]')[n0].click(); await sleep(6);
  eq($('lines-body').querySelectorAll('tr').length, n0, '行が減らない');
});

await TA('2. 下書きを保存すると倉庫に1行入る', async () => {
  $('b-save').click();
  await sleep(30);
  eq(db.pay_invoices.length, 1, '倉庫の行数');
  eq(db.pay_invoices[0].status, 'draft');
  ok(db.pay_invoices[0].no, '番号が空のまま保存された');
  eq(db.pay_invoices[0].totals.grandTotal, 346, '保存した合計');
});

const firstNo = db.pay_invoices[0].no;

await TA('2. ★発行すると固まる（写しが入り、状態が発行済になる）', async () => {
  $('b-issue').click();
  await sleep(40);
  const row = db.pay_invoices[0];
  eq(row.status, 'issued');
  ok(row.issued_at, '発行時刻が入っていない');
  eq(row.snapshot.partner.name, '藤原建設株式会社', '写しの宛先');
  eq(row.snapshot.partner.honor, '御中', '写しの敬称（hubのkeishoを読めていない）');
  eq(row.snapshot.org.yago, '株式会社ゼロアクト', '写しの自社');
  eq(row.template_id, 'std1', '様式');
  eq(row.totals.grandTotal, 346);
});

await TA('2. ★発行済みは直せない・もう一度発行できない（理由も出る）', async () => {
  ok($('e-partner').disabled, '取引先が直せる');
  ok($('e-no').disabled, '番号が直せる');
  ok(qa('#lines-body input').every((i) => i.disabled), '明細が直せる');
  ok($('b-save').disabled, '保存が押せる');
  ok($('b-issue').disabled, '発行がもう一度押せる');
  ok($('b-delete').disabled, '発行済みなのに削除が押せる');
  ok(!$('b-void').disabled, '取り消しが押せない');
  ok(/発行済み/.test($('act-why').textContent), '押せない理由が出ていない: ' + $('act-why').textContent);
  ok($('edit-locked').style.display !== 'none', '発行済みの断り書きが出ていない');
});

/* ═══ 2-b. ★どんな項目にも対応できる（列を自分で決める）★ ═══ */
await TA('2-b. ★列を1本足すと、入力の表にも紙にも出る／金額は1円も動かない', async () => {
  const before = win.SeikyuApp._state.cur.totals.grandTotal;
  doc.querySelector('.ex-bn[data-scr="scr-set"]').click();
  await sleep(20);
  const n0 = $('col-list').querySelectorAll('.col-row').length;
  $('col-new').value = '行き先';
  $('b-col-add').click();
  await sleep(20);
  eq($('col-list').querySelectorAll('.col-row').length, n0 + 1, '列が増えていない');
  ok(/行き先/.test($('col-list').textContent), '足した列が一覧に無い');
  ok($('col-ok').style.display !== 'none', '足したことを伝えていない');

  // 新しい1通を作ると、その列が入力の表に出る
  $('b-new').click();
  await sleep(30);
  const heads = [...$('lines-head').querySelectorAll('th')].map((th) => th.textContent);
  ok(heads.indexOf('行き先') >= 0, '入力の表に足した列が出ていない: ' + heads.join('/'));

  // 値を入れて紙に出す
  setVal('e-partner', 'pt_a');
  await sleep(20);
  const tr = $('lines-body').querySelector('tr');
  tr.querySelector('[data-f="name"]').value = 'あ';
  tr.querySelector('[data-f="name"]').dispatchEvent(new win.Event('input'));
  tr.querySelector('[data-f="amount"]').value = '1000';
  tr.querySelector('[data-f="amount"]').dispatchEvent(new win.Event('input'));
  const x = tr.querySelector('[data-x="行き先"]');
  ok(x, '足した列の入力欄が無い');
  x.value = '今治→松山';
  x.dispatchEvent(new win.Event('input'));
  await sleep(20);
  $('b-preview').click();
  await sleep(400);
  const src = $('pv').srcdoc || '';
  ok(/行き先/.test(src), '紙に足した列の見出しが無い');
  ok(/今治→松山/.test(src), '紙に足した列の中身が無い');
  // ★列を足しても、金額は明細だけから出る（列は金額に触らない）
  ok(before > 0, '前の1通の合計が取れていない');
  const totText = $('tot-box').textContent.replace(/\s+/g, '');
  ok(/合計1,100円/.test(totText), '列を足したら合計が変わった: ' + totText);
});

await TA('2-b. ★幅は 24〜400 から出られない（−を連打しても列が消えない）', async () => {
  doc.querySelector('.ex-bn[data-scr="scr-set"]').click();
  await sleep(20);
  const row = $('col-list').querySelector('.col-row');
  const minus = row.querySelector('[data-w="-8"]');
  for (let i = 0; i < 40; i++) { minus.click(); await sleep(2); }
  await sleep(20);
  const w1 = Number($('col-list').querySelector('.col-row .col-w').textContent);
  eq(w1, 24, '下限を割った: ' + w1);
  const plus = $('col-list').querySelector('.col-row [data-w="8"]');
  for (let i = 0; i < 80; i++) { plus.click(); await sleep(2); }
  await sleep(20);
  const w2 = Number($('col-list').querySelector('.col-row .col-w').textContent);
  eq(w2, 400, '上限を超えた: ' + w2);
  // ★どれだけ広げても、紙に割り付ける％の合計は 100 のまま＝はみ出さない
  const pcts = [...$('col-list').querySelectorAll('.col-row')].map((r) => {
    const t = r.querySelectorAll('.col-w')[1].textContent;
    return Number(t.replace('%', ''));
  });
  const sum = pcts.reduce((a, b) => a + b, 0);
  ok(Math.abs(sum - 100) < 0.6, '紙に割り付ける合計が100%でない: ' + sum.toFixed(2));
});

await TA('2-b. ★揃えを変えられる／列を消せる／既定に戻せる', async () => {
  const row = $('col-list').querySelector('.col-row');
  row.querySelector('[data-al="right"]').click();
  await sleep(20);
  ok($('col-list').querySelector('.col-row [data-al="right"]').classList.contains('on'), '揃えが変わっていない');
  const n0 = $('col-list').querySelectorAll('.col-row').length;
  $('col-list').querySelector('.col-row [data-cdel]').click();
  await sleep(20);
  eq($('col-list').querySelectorAll('.col-row').length, n0 - 1, '列が消えない');
  $('b-col-reset').click();
  await sleep(20);
  const heads = [...$('col-list').querySelectorAll('.col-name')].map((e) => e.firstChild.textContent);
  eq(heads.join('/'), '#/品名・内容/数量/単位/単価/金額/税率', '既定に戻っていない: ' + heads.join('/'));
});

await TA('2-b. ★様式を替えても金額が1円も動かない（見た目だけ変わる）', async () => {
  doc.querySelector('.ex-bn[data-scr="scr-list"]').click();
  await sleep(20);
  doc.querySelector('#fil-seg [data-fil="issued"]').click();
  await sleep(10);
  $('list-body').querySelector('[data-open]').click();
  await sleep(30);
  const g0 = win.SeikyuApp._state.cur.totals.grandTotal;
  $('b-preview').click();
  await sleep(400);
  const a = $('pv').srcdoc || '';
  // 発行済みは様式を選べない（写しで固まっている）
  ok([...doc.querySelectorAll('#e-tpl .ex-chip')].every((b) => b.disabled), '発行済みなのに様式を変えられる');
  eq(win.SeikyuApp._state.cur.totals.grandTotal, g0, '見ただけで合計が動いた');
  const money = String(g0).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  ok(a.replace(/\s+/g, '').includes(money), '紙に合計 ' + money + ' が出ていない');
  // ★発行済みは写しの列で刷る＝あとで会社が列を足しても、出した紙は変わらない
  const heads = [...a.matchAll(/<th class="c-col"[^>]*>([^<]*)<\/th>/g)].map((m) => m[1]);
  eq(heads.join('/'), '#/品名・内容/数量/単位/単価/金額/税率', '発行済みの紙の列が後から変わった: ' + heads.join('/'));
  ok(heads.indexOf('行き先') < 0, '発行後に足した列が、出した紙に入り込んでいる');
});

/* ═══ 3. 出す（紙・Excel） ═══ */
await TA('3. ★「中身を見る」で紙の下見が出る（アプリの画面は入らない）', async () => {
  $('b-preview').click();
  await sleep(400);
  const src = $('pv').srcdoc || '';
  ok(/請　求　書/.test(src), '紙になっていない');
  ok(/藤原建設株式会社/.test(src), '宛先が出ていない');
  ok(!/botnav|appbar|b-issue/.test(src), 'アプリの画面が紙に混ざっている');
  ok($('pv-wrap').style.display !== 'none', '下見の枠が出ていない');
});

await TA('3. ★別の1通に切り替えたら、前の紙の下見は消える（違う紙を出したままにしない）', async () => {
  $('b-new').click();
  await sleep(20);
  eq($('pv-wrap').style.display, 'none', '前の紙が残っている');
  // 元の発行済みに戻す
  doc.querySelector('.ex-bn[data-scr="scr-list"]').click();
  await sleep(20);
  doc.querySelector('#fil-seg [data-fil="issued"]').click();
  await sleep(10);
  $('list-body').querySelector('[data-open]').click();
  await sleep(20);
  doc.querySelector('#fil-seg [data-fil="all"]');
});

await TA('3. ★印刷は「紙だけの新しい窓」に書かれる', async () => {
  const n0 = opened.length;
  $('b-print').click();
  await sleep(10);
  ok($('fn-ov').classList.contains('open'), '★落とす前にファイル名を見せていない★');
  const suggested = $('fn-input').value;
  ok(/^20260930_藤原建設株式会社_請求書_346$/.test(suggested), '推奨名が中身から作られていない: ' + suggested);
  $('fn-ok').click();
  await sleep(1200);
  ok(opened.length === n0 + 1, '新しい窓が開かない');
  const w = opened[opened.length - 1];
  ok(/<!DOCTYPE html>/.test(w._html), '紙が書かれていない');
  ok(/請　求　書/.test(w._html), '紙の見出しが無い');
  ok(!/botnav|appbar|b-issue|<script/i.test(w._html), 'アプリの画面/スクリプトが紙の窓に混ざっている');
  ok(w._printed, '印刷が呼ばれていない');
  eq(w.document.title, '20260930_藤原建設株式会社_請求書_346.pdf', 'PDFの既定の名前が窓の題名になっていない');
});

await TA('3. ★Excelは正しい種類で落ちる（iPhoneで開けない octet-stream にしない）', async () => {
  const n0 = anchorClicks.length;
  $('b-xlsx').click();
  await sleep(10);
  ok($('fn-ov').classList.contains('open'), 'ファイル名を見せていない');
  ok(/\.xlsx$/.test($('fn-input').value + '.xlsx'), '');
  $('fn-input').value = '20260930_藤原建設_請求書_346';
  $('fn-ok').click();
  await sleep(60);
  ok(anchorClicks.length === n0 + 1, '落ちていない');
  const a = anchorClicks[anchorClicks.length - 1];
  eq(a.download, '20260930_藤原建設_請求書_346.xlsx', '直した名前で落ちていない');
  eq(a.target, '_blank', '★ホーム画面アプリで戻れなくなる（target=_blank が無い）★');
  const d = delivered[delivered.length - 1];
  eq(d.type, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'ファイルの種類');
});

/* ═══ 4. 番号 ═══ */
await TA('4. ★同じ番号を二度使わない（倉庫が弾き、1つ進めて出し直す）', async () => {
  $('b-new').click();
  await sleep(20);
  setVal('e-partner', 'pt_a');
  setVal('e-issue', '2026-09-30');
  await sleep(20);
  // 「自分で決める」にして、わざと発行済みと同じ番号を入れる
  doc.querySelector('#e-nomode [data-nm="manual"]').click();
  await sleep(6);
  setVal('e-no', firstNo);
  const tr = $('lines-body').querySelector('tr');
  tr.querySelector('[data-f="name"]').value = 'テスト';
  tr.querySelector('[data-f="name"]').dispatchEvent(new win.Event('input'));
  tr.querySelector('[data-f="amount"]').value = '1000';
  tr.querySelector('[data-f="amount"]').dispatchEvent(new win.Event('input'));
  await sleep(10);
  $('b-issue').click();
  await sleep(60);
  const nos = db.pay_invoices.map((x) => x.no);
  eq(new Set(nos).size, nos.length, '同じ番号が2つ入った: ' + nos.join(','));
  eq(db.pay_invoices.length, 2, '2通目が入っていない');
  ok(/進めました/.test($('edit-ok').textContent), '番号を進めたことを伝えていない: ' + $('edit-ok').textContent);
});

const secondId = db.pay_invoices[1].id;

await TA('4. ★取り消しても番号は空かない（同じ番号は二度と使えない）', async () => {
  $('b-void').click();
  await sleep(40);
  eq(db.pay_invoices[1].status, 'void');
  ok(db.pay_invoices[1].voided_at, '取り消し時刻が入っていない');
  // 取り消した番号をもう一度は使えない（倉庫が持っているので nextNo も避ける）
  const voidedNo = db.pay_invoices[1].no;
  const used = await win.SeikyuApp._state.store.invoices.usedNos('invoice');
  ok(used.indexOf(voidedNo) >= 0, '取り消した番号が「使用済み」から外れている');
});

/* ═══ 5. 一覧 ═══ */
await TA('5. 一覧に2通出る・絞り込みが効く', async () => {
  doc.querySelector('.ex-bn[data-scr="scr-list"]').click();
  await sleep(20);
  doc.querySelector('#fil-seg [data-fil="all"]').click();   // 1. の総当たりで絞り込みが残っているので戻す
  await sleep(6);
  eq($('list-body').querySelectorAll('[data-open]').length, 2, 'すべて');
  doc.querySelector('#fil-seg [data-fil="issued"]').click();
  await sleep(6);
  eq($('list-body').querySelectorAll('[data-open]').length, 1, '発行済');
  doc.querySelector('#fil-seg [data-fil="void"]').click();
  await sleep(6);
  eq($('list-body').querySelectorAll('[data-open]').length, 1, '取り消し');
  doc.querySelector('#fil-seg [data-fil="all"]').click();
  await sleep(6);
});

T('5. ★入金は「未入金(0件)」として出る（読めた上での0件）', () => {
  ok(/未入金/.test($('list-body').textContent), '入金の状態が出ていない');
  ok(!/未確認/.test($('list-body').textContent), '読めているのに未確認と出ている');
});

await TA('5. ★入金が読めなかった時は「未確認」と言い、0件と作り分ける', async () => {
  const st = win.SeikyuApp._state;
  const keep = st.receipts;
  st.receipts = null;
  doc.querySelector('#fil-seg [data-fil="all"]').click();   // 描き直す
  await sleep(10);
  ok(/未確認/.test($('list-body').textContent), '読めなかったのに未入金と言い切っている');
  st.receipts = keep;
  doc.querySelector('#fil-seg [data-fil="all"]').click();
  await sleep(10);
});

await TA('5. 一覧から開くと、その1通が出る', async () => {
  $('list-body').querySelector('[data-open]').click();
  await sleep(20);
  ok($('scr-edit').classList.contains('active'), '中身の画面が開かない');
  ok($('edit-h').textContent.length > 2, '見出しが空: ' + $('edit-h').textContent);
});

/* ═══ 6. 設定 ═══ */
await TA('6. 設定を保存すると自社の棚に入る（番号の形・丸め方・振込先）', async () => {
  doc.querySelector('.ex-bn[data-scr="scr-set"]').click();
  await sleep(20);
  $('s-format').value = 'y-seq'; $('s-format').dispatchEvent(new win.Event('change'));
  $('s-round').value = 'round'; $('s-round').dispatchEvent(new win.Event('change'));
  $('s-bank').value = '伊予銀行 今治支店 普通 1234567';
  $('b-set-save').click();
  await sleep(40);
  const org = db.pay_org[0].data;
  eq(org.numbering.invoice.format, 'y-seq');
  eq(org.taxRounding, 'round');
  eq(org.bank, '伊予銀行 今治支店 普通 1234567');
  eq(org.yago, '株式会社ゼロアクト', '★ハブが入れた自社情報を消している★');
});

await TA('6. ★「連番だけ＋毎年1に戻す」は保存させない（去年と必ずぶつかる）', async () => {
  $('s-format').value = 'seq'; $('s-format').dispatchEvent(new win.Event('change'));
  $('s-reset').checked = true; $('s-reset').dispatchEvent(new win.Event('change'));
  $('b-set-save').click();
  await sleep(30);
  ok($('set-err').style.display !== 'none', '止めていない');
  eq(db.pay_org[0].data.numbering.invoice.format, 'y-seq', '止めたのに保存された');
  $('s-format').value = 'y-seq'; $('s-format').dispatchEvent(new win.Event('change'));
  $('s-reset').checked = true;
});

await TA('6. ★取引先に請求書用の項目を足しても、ハブが入れた名前・住所を消さない', async () => {
  $('s-partner').value = 'pt_a'; $('s-partner').dispatchEvent(new win.Event('change'));
  await sleep(10);
  eq($('s-phonor').value, '御中', 'ハブの敬称を読めていない');
  $('s-pcode').value = 'A001';
  $('s-pperson').value = '山田';
  $('s-pterm').value = 'nextEom'; $('s-pterm').dispatchEvent(new win.Event('change'));
  $('b-pt-save').click();
  await sleep(40);
  const d = db.pay_partners.find((x) => x.id === 'pt_a').data;
  eq(d.code, 'A001');
  eq(d.person, '山田');
  eq(d.payTerm.kind, 'nextEom');
  eq(d.name, '藤原建設株式会社', '★ハブが入れた名前が消えた★');
  eq(d.addr, '愛媛県今治市1-2-3', '★ハブが入れた住所が消えた★');
  eq(d.invoiceNo, 'T9876543210987', '★ハブが入れた登録番号が消えた★');
  eq(d.keisho, d.honor, '敬称が2つのキーで食い違っている（ハブの画面で化ける）');
});

await TA('6. ★取引先を選ぶまでは、その欄を触らせない（誰のか分からないまま保存させない）', async () => {
  $('s-partner').value = ''; $('s-partner').dispatchEvent(new win.Event('change'));
  await sleep(10);
  ok($('s-pcode').disabled, '取引先未選択なのに入力できる');
  ok($('b-pt-save').disabled, '取引先未選択なのに保存できる');
});

/* ═══ 7. 支払期限 ═══ */
await TA('7. 支払期限は決め方から自動で入り、手でも直せる', async () => {
  $('b-new').click();
  await sleep(20);
  setVal('e-issue', '2026-09-30');
  $('e-term').value = 'nextEom'; $('e-term').dispatchEvent(new win.Event('change'));
  await sleep(10);
  eq($('e-due').value, '2026-10-31', '翌月末が入らない');
  $('e-term').value = 'none'; $('e-term').dispatchEvent(new win.Event('change'));
  await sleep(10);
  setVal('e-due', '2026-11-15');
  eq(win.SeikyuApp._state.cur.due_ymd, '2026-11-15', '手で直した期限が持たれていない');
});

/* ═══ 6-c. ★角印（会社の印）★ ═══ */
await TA('6-c. ★角印を入れると紙に出る／大きさを変えられる／消せる', async () => {
  doc.querySelector('.ex-bn[data-scr="scr-set"]').click();
  await sleep(30);
  ok($('seal-none').style.display !== 'none', '最初から印が入っていることになっている');
  ok($('b-seal-clear').disabled, '印が無いのに「消す」が押せる');

  // ファイル選択は jsdom で作れないので、読み込んだあとの data URL を直接渡す
  const seal = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';
  const r = win.SeikyuApp._pickSealUrl(seal);
  ok(r.ok, '正しい画像がはじかれた: ' + r.reason);
  await sleep(20);
  ok($('seal-pv').style.display !== 'none', '下見に出ていない');
  $('seal-mm').value = '30';
  $('b-seal-save').click();
  await sleep(60);
  eq(db.pay_org[0].data.sealDataUrl, seal, '倉庫に印が入っていない');
  eq(db.pay_org[0].data.sealSizeMm, 30, '大きさが入っていない');
  eq(db.pay_org[0].data.yago, '株式会社ゼロアクト', '★ハブが入れた自社情報を消している★');

  // 新しい1通の紙に出る
  $('b-new').click();
  await sleep(30);
  setVal('e-partner', 'pt_a');
  await sleep(20);
  const tr = $('lines-body').querySelector('tr');
  tr.querySelector('[data-f="name"]').value = 'あ';
  tr.querySelector('[data-f="name"]').dispatchEvent(new win.Event('input'));
  tr.querySelector('[data-f="amount"]').value = '1000';
  tr.querySelector('[data-f="amount"]').dispatchEvent(new win.Event('input'));
  await sleep(20);
  $('b-preview').click();
  await sleep(400);
  const src = $('pv').srcdoc || '';
  ok(/class="seal"/.test(src), '紙に印が出ていない');
  ok(/width:30mm/.test(src), '紙の印の大きさが効いていない');
});

await TA('6-c. ★大きすぎる画像・PNG/JPEG でない物は入らない（理由を出す）', async () => {
  doc.querySelector('.ex-bn[data-scr="scr-set"]').click();
  await sleep(30);
  const bad = win.SeikyuApp._pickSealUrl('https://example.com/hanko.png');
  ok(!bad.ok, '外のURLが通った');
  ok($('seal-err').style.display !== 'none', '理由を出していない');
  const big = win.SeikyuApp._pickSealUrl('data:image/png;base64,' + 'A'.repeat(500 * 1024));
  ok(!big.ok, '大きすぎる画像が通った');
  ok(/KB/.test($('seal-err').textContent), '何KBかを言っていない: ' + $('seal-err').textContent);
  // 前に保存した印は残っている（弾かれても消えない）
  eq(db.pay_org[0].data.sealDataUrl.slice(0, 22), 'data:image/png;base64,', '弾かれた拍子に保存済みの印が消えた');
});

await TA('6-c. ★印を消せる（消しても、すでに出した紙は変わらない）', async () => {
  const issued = db.pay_invoices.filter((x) => x.status === 'issued')[0];
  const before = issued.snapshot.org.sealDataUrl || '';
  $('b-seal-clear').click();
  await sleep(60);
  eq(db.pay_org[0].data.sealDataUrl, '', '倉庫から消えていない');
  ok($('seal-none').style.display !== 'none', '「入れていません」に戻っていない');
  eq(issued.snapshot.org.sealDataUrl || '', before, '★出した紙の写しが書き換わった★');
});

/* ═══ 6-b. ★列を選べるようになる前に出した紙も、あとから列が増えない★ ═══ */
await TA('6-b. ★写しに列が無い古い請求書は、会社の「今の列」を当てずに様式の既定で刷る', async () => {
  // 列を足す前に出した1通を作る（写しに cols が無い＝2026-08-10 より前に出した物と同じ形）
  const old = JSON.parse(JSON.stringify(db.pay_invoices.find((x) => x.status === 'issued')));
  old.id = 'iv_old_no_cols';
  old.no = 'OLD-0001';
  delete old.snapshot.cols;                 // ★列を覚えていない写し
  db.pay_invoices.push(old);
  // 会社の列には「行き先」が入っている状態にする
  const st = win.SeikyuApp._state;
  st.org.invoiceCols = { items: ['#', '品名・内容', '金額', '税率', '行き先'], widths: {}, aligns: {} };

  doc.querySelector('.ex-bn[data-scr="scr-list"]').click();
  await sleep(20);
  doc.querySelector('#fil-seg [data-fil="all"]').click();
  await sleep(20);
  await win.SeikyuApp._state.store.invoices.list('invoice');
  $('b-reload').click();
  await sleep(60);
  const row = [...$('list-body').querySelectorAll('[data-open]')].find((r) => /OLD-0001/.test(r.textContent));
  ok(row, '作った古い1通が一覧に無い');
  row.click();
  await sleep(30);
  $('b-preview').click();
  await sleep(400);
  const heads = [...($('pv').srcdoc || '').matchAll(/<th class="c-col"[^>]*>([^<]*)<\/th>/g)].map((m) => m[1]);
  eq(heads.join('/'), '#/品名・内容/数量/単位/単価/金額/税率', '古い紙に、あとから足した列が入り込んだ: ' + heads.join('/'));
  ok(heads.indexOf('行き先') < 0, '★出した紙が、列を足した日に変わってしまっている★');
});

/* ═══ 7-b. ★読めなかったを「空」にしない（実機で踏んだ401）★ ═══ */
await TA('7-b. ★自社情報が読めなかった時は、空っぽ扱いにせず「読めなかった」と言い、発行を止める', async () => {
  // 1回目も2回目（取り直し）も失敗させる＝本当に読めない状態を作る
  sb._failNext('pay_org');
  const app = win.SeikyuApp;
  const st = app._state;
  const keepOrg = st.org;
  await app._loadMasters();
  await sleep(700);
  sb._failNext('pay_org');
  await app._loadMasters();
  await sleep(50);
  eq(st.org, null, '読めなかったのに空っぽ({})にしている＝紙に「自社情報が未入力」と出る');
  eq(st.orgReadOk, false, '読めたことになっている');
  ok($('list-err').style.display !== 'none', '読めなかったことを画面で言っていない');
  ok(/読めません/.test($('list-err').textContent), '文言が「読めなかった」になっていない: ' + $('list-err').textContent);

  // この状態では発行させない（空の自社が写しに固まると、もう直せない紙になる）
  $('b-new').click();
  await sleep(20);
  setVal('e-partner', 'pt_a');
  await sleep(20);
  const tr = $('lines-body').querySelector('tr');
  tr.querySelector('[data-f="name"]').value = 'あ';
  tr.querySelector('[data-f="name"]').dispatchEvent(new win.Event('input'));
  tr.querySelector('[data-f="amount"]').value = '1000';
  tr.querySelector('[data-f="amount"]').dispatchEvent(new win.Event('input'));
  await sleep(20);
  const n0 = db.pay_invoices.length;
  $('b-issue').click();
  await sleep(60);
  eq(db.pay_invoices.length, n0, '自社が読めていないのに発行された');
  ok(/読めていない/.test($('edit-err').textContent), '止めた理由を言っていない: ' + $('edit-err').textContent);

  // 「読み直す」で直る
  $('b-back').click();
  await sleep(10);
  await win.SeikyuApp._loadMasters();
  await sleep(50);
  eq(st.orgReadOk, true, '読み直しても直らない');
  ok(st.org && st.org.yago, '自社情報が戻っていない');
  ok(keepOrg !== undefined, '');
});

/* ═══ 8. まとめ ═══ */
T('8. ★最後まで JS が1つも落ちていない', () => {
  eq(errs.length, 0, errs.join(' | '));
});

console.log('\n' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
