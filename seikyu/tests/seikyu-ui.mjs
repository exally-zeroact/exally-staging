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
    const nav = doc.querySelector('.bn[data-scr="' + scr + '"]');
    ok(nav, 'ナビ ' + scr + ' が無い');
    nav.click();
    await sleep(10);
    const el = $(scr);
    ok(el.classList.contains('active'), scr + ' が開かない');
    ok(el.innerHTML.length > 400, scr + ' の中身が薄い(' + el.innerHTML.length + ')');
    for (const b of [...el.querySelectorAll('button'), ...doc.querySelectorAll('.bn')]) {
      if (b.disabled) { pressed.push((b.id || b.textContent.trim()) + '(押せない)'); continue; }
      b.click();
      await sleep(6);
      pressed.push(b.id || b.getAttribute('data-fil') || b.getAttribute('data-tm') || b.getAttribute('data-nm') || b.textContent.trim());
    }
    // 押した拍子に別の画面へ行っていたら戻す
    doc.querySelector('.bn[data-scr="' + scr + '"]').click();
    await sleep(6);
  }
  // ファイル名の小窓のボタンも押す
  $('fn-cancel').click();
  pressed.push('fn-cancel');
  console.log('     押した物(' + pressed.length + '): ' + pressed.join(' / '));
  ok(pressed.length >= 25, '押した物が少なすぎる（一覧が取れていない）: ' + pressed.length);
  eq(errs.length, before, 'JSが落ちた: ' + errs.slice(before).join(' | '));
});

/* ═══ 1-b. スマホ幅で潰れない書き方 ═══
   jsdom は幅を計算しないので、ここでは ★潰れない書き方になっているか★ を見る
   （実物の幅は実機幅の画面で定規を当てて確かめる。この検査はその前段の網）。 */
T('1-b. ★明細の表は「縮めて潰す」のではなく「横に動かす」（実機幅375pxで欄が幅ゼロになった前科）', () => {
  const css = fs.readFileSync(path.join(ROOT, 'seikyu/css/app.css'), 'utf8');
  const lines = (/\.lines\s*\{([^}]*)\}/.exec(css) || [])[1] || '';
  ok(/min-width\s*:\s*\d+px/.test(lines), '.lines に min-width が無い（入れ物に合わせて縮む＝欄が潰れる）');
  const wrap = (/\.lines-scroll\s*\{([^}]*)\}/.exec(css) || [])[1] || '';
  ok(/overflow-x\s*:\s*auto/.test(wrap), '.lines-scroll が横に動かせない');
  ok(!/\.lines\s*\{[^}]*width\s*:\s*100%/.test(css), '.lines に width:100% がある（min-width を打ち消す）');
  ok(/<div class="lines-scroll">/.test(html), '表が横に動く入れ物に入っていない');
});

T('1-b. ★文が入る箱は block で最低幅を持ち、日本語を1文字ずつ割らない', () => {
  const css = fs.readFileSync(path.join(ROOT, 'seikyu/css/app.css'), 'utf8');
  ok(!/word-break\s*:\s*break-all/.test(css), 'break-all がある（日本語が1文字ずつ割れる）');
  for (const sel of ['.hint', '.warn', '.bad', '.ok', '.why', '.scroll-note']) {
    const rule = (new RegExp('\\' + sel + '\\s*\\{([^}]*)\\}').exec(css) || [])[1] || '';
    ok(/display\s*:\s*block/.test(rule), sel + ' が block でない');
    ok(/min-width\s*:\s*\d/.test(rule), sel + ' に最低幅が無い');
    ok(/overflow-wrap\s*:\s*break-word/.test(rule), sel + ' に折り返しの指定が無い');
  }
  ok(/\.btn-row\s*\{[^}]*flex-wrap\s*:\s*wrap/.test(css), 'ボタンの行が折り返さない（横にはみ出す）');
});

T('1-b. ★入力欄は16px（これより小さいと iPhone が勝手に拡大して画面がズレる）', () => {
  const css = fs.readFileSync(path.join(ROOT, 'seikyu/css/app.css'), 'utf8');
  const rule = (/\.finput\s*\{([^}]*)\}/.exec(css) || [])[1] || '';
  const m = /font-size\s*:\s*(\d+(?:\.\d+)?)px/.exec(rule);
  ok(m, '.finput に文字の大きさが無い');
  ok(Number(m[1]) >= 16, '入力欄が ' + m[1] + 'px（16px 未満）');
});

T('1-b. ★iPhone で持っていない入力（月・週）を使っていない', () => {
  ok(!/type="(month|week|datetime-local)"/.test(html), 'iOS が持っていない入力がある');
});

/* ═══ 2. 1通 作って発行する ═══ */
$('b-new').click();
await sleep(10);

await TA('2. ★新しく作る＝白紙を埋めさせない（今日・既定の税・番号が最初から入る）', async () => {
  ok($('e-issue').value, '請求日が空');
  ok($('e-no').value, '番号が空: ' + $('e-no').value);
  ok(qa('#e-taxmode .seg-b.on').length === 1, '税の入れ方が選ばれていない');
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

/* ═══ 3. 出す（紙・Excel） ═══ */
await TA('3. ★「中身を見る」で紙の下見が出る（アプリの画面は入らない）', async () => {
  $('b-preview').click();
  await sleep(400);
  const src = $('pv').srcdoc || '';
  ok(/請 求 書/.test(src), '紙になっていない');
  ok(/藤原建設株式会社/.test(src), '宛先が出ていない');
  ok(!/botnav|appbar|b-issue/.test(src), 'アプリの画面が紙に混ざっている');
  ok($('pv-wrap').style.display !== 'none', '下見の枠が出ていない');
});

await TA('3. ★別の1通に切り替えたら、前の紙の下見は消える（違う紙を出したままにしない）', async () => {
  $('b-new').click();
  await sleep(20);
  eq($('pv-wrap').style.display, 'none', '前の紙が残っている');
  // 元の発行済みに戻す
  doc.querySelector('.bn[data-scr="scr-list"]').click();
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
  ok(/請 求 書/.test(w._html), '紙の見出しが無い');
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
  doc.querySelector('.bn[data-scr="scr-list"]').click();
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
  doc.querySelector('.bn[data-scr="scr-set"]').click();
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

/* ═══ 8. まとめ ═══ */
T('8. ★最後まで JS が1つも落ちていない', () => {
  eq(errs.length, 0, errs.join(' | '));
});

console.log('\n' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
