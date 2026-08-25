/* ai-jiko-dome.test.mjs — ★4 事故止め（2026-08-25 司さんの数字・指示役の検証要件）★
 *
 *  ★塞いだ穴（私が 偽のAIで 0円で押して 実測した）★
 *    message 20,000字 ＋ history 40件×50,000字 → ★202万字が そのまま AIへ渡り 200 が返った★
 *    ＝ ★件数(40件)しか見ていなかった★ので、1件が長ければ いくらでも通った。
 *
 *  ★決まり（数字は api/claude.js の 事故止め 1か所だけ）★
 *    ①★1分に10回（人ごと）★  ②★1日に100回（人ごと）★
 *    ③★history 合計40,000字（古い方から捨てる・最後の1往復は必ず残す）★
 *    ④★1回にAIへ渡すのは 2万トークンまで★（2026-08-09 司さんの決定）
 *    ⑤ max_tokens 2000 据え置き
 *    ⑥★429の言い方は「混み合っています」ではなく 待てば直ると分かる言い方★
 *    ⑦★止めた時も 必ず 記録を残す★（黙って止めない・黙って小さくしない）
 *
 *  ★お金は1円も使わない★＝AIは 偽物に差し替えて、実際に handler を呼ぶ。
 *
 *  使い方: node tests/ai-jiko-dome.test.mjs [--self-test]
 */
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createRequire } from 'node:module';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(__dirname, '..');
const SELF = process.argv.includes('--self-test');
const OVERRIDE = process.env.EXALLY_JIKO_OVERRIDE ? JSON.parse(process.env.EXALLY_JIKO_OVERRIDE) : {};
const require_ = createRequire(pathToFileURL(path.join(ROOT, 'package.json')));

const handler = require_(OVERRIDE['api/claude.js'] || path.join(ROOT, 'api/claude.js'));
/* 客に見せる言葉は 画面の側（lib/ai-reason.js）が1か所で作る */
const AiReason = (() => {
  /* ★lib/ai-reason.js は globalThis に付ける作り★＝そのまま読むと 呼ぶ側を汚す。
     ★別の入れ物(vm)で 本番と同じ字を そのまま動かして 受け取る★ */
  const src = fs.readFileSync(OVERRIDE['lib/ai-reason.js'] || path.join(ROOT, 'lib/ai-reason.js'), 'utf8');
  const vm = require_('node:vm');
  const 箱 = { module: { exports: {} } };
  vm.createContext(箱);
  vm.runInContext(src, 箱);
  return 箱.AiReason || 箱.module.exports;
})();

let pass = 0, fail = 0;
const AT = async (n, fn) => { try { await fn(); pass++; console.log('  ok   ' + n); } catch (e) { fail++; console.log('  NG   ' + n + '\n       ' + (e && e.message)); } };
const T = (n, fn) => { try { fn(); pass++; console.log('  ok   ' + n); } catch (e) { fail++; console.log('  NG   ' + n + '\n       ' + (e && e.message)); } };
const ok = (c, m) => { if (!c) throw new Error(m || 'expected truthy'); };
const eq = (a, b, m) => { if (a !== b) throw new Error((m ? m + ': ' : '') + '期待=' + JSON.stringify(b) + ' 実際=' + JSON.stringify(a)); };

/* ── 偽のAI（★1円も使わない★）と 記録の受け皿 ── */
let 送った = null;
const 偽AI = () => handler.__setClient({
  messages: {
    create: async (p) => {
      送った = p;
      return { content: [{ type: 'text', text: 'はい' }], usage: { input_tokens: 5, output_tokens: 3 } };
    },
  },
});
const 記録たち = [];
const 元log = console.log;
function 記録を拾う(on) {
  if (on) { console.log = (...a) => { const s = String(a[0] || ''); if (s.startsWith('[ai] ')) { try { 記録たち.push(JSON.parse(s.slice(5))); } catch (e) { /* 読めない行は 数えない */ } } else 元log(...a); }; }
  else console.log = 元log;
}

const 押す = async (o) => {
  o = o || {};
  偽AI();
  const 出た = { headers: {}, status: null, body: null };
  const res = {
    setHeader(k, v) { 出た.headers[k] = v; },
    status(c) { 出た.status = c; return this; },
    json(b) { 出た.body = b; return this; },
    end() { return this; },
  };
  const headers = { origin: 'https://exally.vercel.app' };
  if (o.人) headers.authorization = 'Bearer x.' + Buffer.from(JSON.stringify({ sub: o.人 })).toString('base64') + '.y';
  if (o.ip) headers['x-forwarded-for'] = o.ip;
  記録を拾う(true);
  try {
    await handler({ method: 'POST', headers, body: { message: o.message || 'これ何？', history: o.history || [] } }, res);
  } finally { 記録を拾う(false); }
  return 出た;
};

console.log('');
console.log('[ai-jiko-dome] ★4 事故止め（回数・大きさ・言い方・記録）★');

/* ══ ★塞いだ穴そのものを 1回 押して 再現する★ ══ */
await AT('★穴の再現：20,000字＋40件×50,000字＝202万字が そのまま渡っていた（もう渡らない）★', async () => {
  handler.__数え場を空にする();
  const 長い = 'あ'.repeat(50000);
  const history = [];
  for (let i = 0; i < 40; i++) history.push({ role: i % 2 ? 'assistant' : 'user', content: 長い });
  const 元の字数 = history.reduce((a, m) => a + m.content.length, 0) + 20000;
  eq(元の字数, 2020000, '★穴の大きさ（元）★');
  const r = await 押す({ message: 'あ'.repeat(20000), history, 人: 'ana', ip: '10.0.0.1' });
  eq(r.status, 200, '★200 が返る（客は 断られない）★');
  const 渡した = 送った.messages.reduce((a, m) => a + (typeof m.content === 'string' ? m.content.length
    : m.content.reduce((b, c) => b + String(c.text || '').length, 0)), 0);
  ok(渡した <= 40000, '★渡した字数が 40,000 を超えている：' + 渡した + '★');
  console.log('       … 元 ' + 元の字数.toLocaleString() + '字 → ★渡した ' + 渡した.toLocaleString() + '字★');
});

/* ══ ①1分に10回 ══ */
await AT('★11回目が 429（1分に10回・人ごと）★', async () => {
  handler.__数え場を空にする();
  for (let i = 1; i <= 10; i++) {
    const r = await 押す({ 人: 'a', ip: '1.1.1.1' });
    eq(r.status, 200, i + '回目');
  }
  const r11 = await 押す({ 人: 'a', ip: '1.1.1.1' });
  eq(r11.status, 429, '★11回目★');
  eq(r11.body.error, 'tsukaisugi', '合言葉');
  ok(r11.body.待ち秒 >= 1 && r11.body.待ち秒 <= 60, '★あと何秒かを 返していない：' + r11.body.待ち秒 + '★');
  eq(r11.headers['Retry-After'], String(r11.body.待ち秒), '★Retry-After を出していない★');
});
await AT('★止めた時も AIを1回も呼んでいない（お金を使わない）★', async () => {
  handler.__数え場を空にする();
  for (let i = 0; i < 10; i++) await 押す({ 人: 'b', ip: '2.2.2.2' });
  送った = null;
  const r = await 押す({ 人: 'b', ip: '2.2.2.2' });
  eq(r.status, 429);
  eq(送った, null, '★止めたのに AIを呼んでいる（お金が出る）★');
});
await AT('★窓が過ぎれば また使える（待てば直る）★', async () => {
  handler.__数え場を空にする();
  const 昔 = Date.now() - 61 * 1000;
  handler.__数えておく(['u:c'], 昔);
  for (let i = 0; i < 9; i++) handler.__数えておく(['u:c'], Date.now());
  const 見 = handler.__押した回数を見る(['u:c'], Date.now());
  eq(見.止める, false, '★1分より前の分まで数えている（いつまでも直らない）★');
});

/* ══ ②1日に100回 ══ */
await AT('★101回目が 429（1日に100回）★', async () => {
  handler.__数え場を空にする();
  const いま = Date.now();
  /* 1分の窓に引っかからないよう 100回を 1日の中に散らして置く */
  for (let i = 0; i < 100; i++) handler.__数えておく(['u:d'], いま - (i + 1) * 60 * 1000);
  const 見 = handler.__押した回数を見る(['u:d'], いま);
  eq(見.止める, true, '★101回目が 通ってしまう★');
  eq(見.どれ, '日');
  ok(見.あと秒 > 60, '★1日の窓なのに 待ち時間が短すぎる：' + 見.あと秒 + '★');
});
await AT('★100回目までは 通る（1つ手前で止めない）★', async () => {
  handler.__数え場を空にする();
  const いま = Date.now();
  for (let i = 0; i < 99; i++) handler.__数えておく(['u:e'], いま - (i + 1) * 60 * 1000);
  eq(handler.__押した回数を見る(['u:e'], いま).止める, false);
});

/* ══ ③人が違えば 別勘定 ══ */
await AT('★人が違えば 別勘定（隣の人が使い切っても 自分は使える）★', async () => {
  handler.__数え場を空にする();
  for (let i = 0; i < 10; i++) await 押す({ 人: 'x', ip: '3.3.3.1' });
  eq((await 押す({ 人: 'x', ip: '3.3.3.1' })).status, 429, 'xは止まる');
  eq((await 押す({ 人: 'y', ip: '3.3.3.2' })).status, 200, '★yまで止まっている★');
});
await AT('★名乗りを書き換えても すり抜けない（IPでも数える）★', async () => {
  handler.__数え場を空にする();
  for (let i = 0; i < 10; i++) await 押す({ 人: 'z' + i, ip: '4.4.4.4' });   /* 毎回 別人を名乗る */
  eq((await 押す({ 人: 'z99', ip: '4.4.4.4' })).status, 429, '★名乗りを変えるだけで すり抜ける★');
});
await AT('★名乗りが無くても 数える（IPだけでも止まる）★', async () => {
  handler.__数え場を空にする();
  for (let i = 0; i < 10; i++) await 押す({ ip: '5.5.5.5' });
  eq((await 押す({ ip: '5.5.5.5' })).status, 429);
});

/* ══ ③会話の字数 ══ */
await AT('★40,001字で 200 が返り、渡す字数は 40,000以下★', async () => {
  handler.__数え場を空にする();
  const history = [];
  for (let i = 0; i < 8; i++) history.push({ role: i % 2 ? 'assistant' : 'user', content: 'あ'.repeat(5000) });
  history[0].content = 'あ'.repeat(5001);                        /* 合計 40,001字 */
  eq(history.reduce((a, m) => a + m.content.length, 0), 40001, '★入れる字数★');
  const r = await 押す({ history, 人: 'f', ip: '6.6.6.6' });
  eq(r.status, 200, '★断ってしまっている（200で返す決まり）★');
  const 渡した = 送った.messages.slice(0, -1).reduce((a, m) => a + (typeof m.content === 'string' ? m.content.length
    : m.content.reduce((b, c) => b + String(c.text || '').length, 0)), 0);
  ok(渡した <= 40000, '★渡した会話が ' + 渡した + '字（40,000を超えた）★');
  console.log('       … 40,001字 → ★' + 渡した + '字★');
});
T('★古い方から捨てる（新しい方を捨てない）★', () => {
  const 会話 = [
    { role: 'user', content: '古' + 'あ'.repeat(30000) },
    { role: 'assistant', content: '中' + 'あ'.repeat(30000) },
    { role: 'user', content: '新1' },
    { role: 'assistant', content: '新2' },
  ];
  const 出 = handler.__会話を字数で削る(会話, 40000);
  /* ★収まった時点で 止める★＝必要以上に捨てない（「中」は残ってよい） */
  ok(出.reduce((a, m) => a + m.content.length, 0) <= 40000, '★上限に収まっていない★');
  eq(出[0].content[0], '中', '★古い方から捨てていない★');
  eq(出[出.length - 2].content, '新1', '★新しい方を捨てている★');
  eq(出[出.length - 1].content, '新2');
});
T('★最後の1往復は 必ず残る（合計が上限を超えていても）★', () => {
  const 会話 = [
    { role: 'user', content: 'あ'.repeat(60000) },
    { role: 'assistant', content: 'い'.repeat(60000) },
  ];
  const 出 = handler.__会話を字数で削る(会話, 40000);
  eq(出.length, 2, '★最後の1往復まで 捨てている（会話にならない）★');
});

/* ══ ④2万トークン ══ */
T('★トークンの見積もりは 多めに見る（日本語1字＝1・英数4字＝1）★', () => {
  eq(handler.__見積もりトークン('あいうえお'), 5);
  eq(handler.__見積もりトークン('abcd'), 1);
  eq(handler.__見積もりトークン(''), 0);
  ok(handler.__見積もりトークン('あ'.repeat(20000)) >= 20000, '★少なく見積もっている（お金が出る側）★');
});
await AT('★前置きも足して 2万トークンを超えない★', async () => {
  handler.__数え場を空にする();
  const history = [];
  for (let i = 0; i < 8; i++) history.push({ role: i % 2 ? 'assistant' : 'user', content: 'あ'.repeat(4900) });
  const r = await 押す({ message: 'あ'.repeat(3000), history, 人: 'g', ip: '7.7.7.7' });
  eq(r.status, 200);
  const 前置き = 送った.system.map((x) => x.text).join('');
  const 見積 = handler.__見積もりトークン(前置き)
    + 送った.messages.reduce((a, m) => a + handler.__見積もりトークン(typeof m.content === 'string' ? m.content
      : m.content.map((c) => c.text).join('')), 0);
  ok(見積 <= 20000, '★渡した見積もりが ' + 見積 + 'トークン（2万を超えた）★');
  console.log('       … ★渡した見積もり ' + 見積 + 'トークン★');
});
T('★最後の1往復だけでも収まらない時は 会話を全部捨てる（お金の上限が先）★', () => {
  const 会話 = [{ role: 'user', content: 'あ'.repeat(30000) }, { role: 'assistant', content: 'い'.repeat(30000) }];
  const 出 = handler.__会話をトークンで削る(会話, 'まえおき', 'いま', 20000);
  eq(出.全部捨てた, true, '★2万トークンを超えたまま 渡している★');
  eq(出.会話.length, 0);
});

/* ══ ⑤max_tokens ══ */
await AT('★max_tokens は 2000 のまま（勝手に増やさない）★', async () => {
  handler.__数え場を空にする();
  await 押す({ 人: 'h', ip: '8.8.8.8' });
  eq(送った.max_tokens, 2000);
});

/* ══ ⑥言い方 ══ */
T('★429の言い方＝「混み合っています」ではなく 待てば直ると分かる言い方★', () => {
  const 分 = AiReason.読む({ status: 429, 理由: 'tsukaisugi', 待ち秒: 42, どれ: '分' });
  ok(分.言葉.indexOf('混み合') < 0, '★「混み合っています」と言っている★：' + 分.言葉);
  /* ★理由の「1分に 10回まで」を 待ち時間と読み違えていた（検査の側の間違い・2026-08-25）★
     ⇒ ★次の一手の字だけ★を見て、★実際の待ち時間（42秒）★が入っているかを見る */
  ok(分.次.indexOf('42秒') >= 0, '★いつまで待てばよいかを 言っていない★：' + 分.次);
  ok(分.言葉.indexOf('10回') >= 0, '★何回までかを 言っていない★');
  const 日 = AiReason.読む({ status: 429, 理由: 'tsukaisugi', 待ち秒: 7200, どれ: '日' });
  ok(日.言葉.indexOf('100回') >= 0, '★1日の上限を 言っていない★');
  ok(日.言葉.indexOf('2時間') >= 0, '★あと何時間かを 言っていない★');
  ok(日.言葉.indexOf('AIに聞かなくても') >= 0, '★止まっていても 出来る事を 言っていない★');
});
T('★止められた言葉に ★ を書かない（客の字）★', () => {
  for (const o of [{ status: 429, 理由: 'tsukaisugi', 待ち秒: 42, どれ: '分' },
    { status: 429, 理由: 'tsukaisugi', 待ち秒: 7200, どれ: '日' },
    { status: 429, 理由: 'komiai' }]) {
    ok(AiReason.読む(o).言葉.indexOf('★') < 0, '★客の字に ★ が出ている★：' + AiReason.読む(o).言葉);
  }
});
T('★AI側の混雑（komiai）と うちの上限（tsukaisugi）を 別の言葉にする★', () => {
  const a = AiReason.読む({ status: 429, 理由: 'tsukaisugi', 待ち秒: 10, どれ: '分' }).言葉;
  const b = AiReason.読む({ status: 429, 理由: 'komiai' }).言葉;
  ok(a !== b, '★同じ言葉＝客は どちらか分からない★');
  ok(b.indexOf('うちの上限ではありません') >= 0, '★AI側だと 分かる言い方になっていない★');
});

/* ══ ⑦記録 ══ */
await AT('★止めた時も 記録が残る（黙って止めない）★', async () => {
  handler.__数え場を空にする();
  記録たち.length = 0;
  for (let i = 0; i < 10; i++) await 押す({ 人: 'i', ip: '9.9.9.9' });
  await 押す({ 人: 'i', ip: '9.9.9.9' });
  const 止めた = 記録たち.filter((r) => r.結果 === 'tsukaisugi');
  eq(止めた.length, 1, '★止めた記録が 残っていない★');
  ok(止めた[0].待ち秒 >= 1, '★あと何秒かを 記録していない★');
  eq(止めた[0].どれ, '分');
});
await AT('★削った時も 記録が残る（黙って小さくしない）★', async () => {
  handler.__数え場を空にする();
  記録たち.length = 0;
  const history = [];
  for (let i = 0; i < 10; i++) history.push({ role: i % 2 ? 'assistant' : 'user', content: 'あ'.repeat(9000) });
  await 押す({ history, 人: 'j', ip: '11.11.11.11' });
  const r = 記録たち.filter((x) => x.結果 === 'ok').pop();
  ok(r, '記録が無い');
  ok(r.会話を削った > 0, '★削ったのに 0 と記録している★');
  eq(r.会話の元の字数, 90000, '★元の字数を 残していない★');
  ok(r.会話の字数 <= 40000, '★削った後の字数が 合わない★');
  ok(r.渡した見積もりトークン > 0 && r.渡した見積もりトークン <= 20000, '★渡した見積もりを 残していない★');
});
T('★数字は 1か所だけ（散らさない）★', () => {
  const src = fs.readFileSync(OVERRIDE['api/claude.js'] || path.join(ROOT, 'api/claude.js'), 'utf8');
  eq(handler.__事故止め.分の回数, 10);
  eq(handler.__事故止め.日の回数, 100);
  eq(handler.__事故止め.会話の合計字数, 40000);
  eq(handler.__事故止め.渡せるトークン, 20000);
  /* ★数字を 直に書いた所が 他に無いか★（コメントの中は 数えない） */
  const 中身 = src.split('\n').filter((l) => !/^\s*(\*|\/\*|\/\/)/.test(l)).join('\n');
  eq((中身.match(/40000/g) || []).length, 1, '★40000 が 2か所以上に書いてある★');
  eq((中身.match(/20000/g) || []).length, 2, '★20000 の置き場（トークンの上限と 字数の上限）が 増えている★');
  eq(handler.__事故止め.一度に送れる字数, 20000, '★1回に送れる字数が 事故止めの外に出た★');
});

/* ══ ⑧画面の側（人ごとに数えるには 誰かを伝える口が要る）══ */
T('★画面が 誰かを伝える（無いと 同じ回線の人が まとめて1人になる）★', () => {
  const book = fs.readFileSync(OVERRIDE['book.html'] || path.join(ROOT, 'book.html'), 'utf8');
  ok(/見出し\['Authorization'\] = 名乗り/.test(book), '★名乗りを 送っていない★');
  ok(/access_token/.test(book), '★名乗りを どこから取るか 書いていない★');
  const auth = fs.readFileSync(OVERRIDE['js/auth.js'] || path.join(ROOT, 'js/auth.js'), 'utf8');
  ok(/global\.Auth\.sb = sb;/.test(auth), '★画面から 名乗りを取れる口が 無い★');
});
T('★画面が「あと何秒で また使えるか」を 受け取って 客に出す★', () => {
  const book = fs.readFileSync(OVERRIDE['book.html'] || path.join(ROOT, 'book.html'), 'utf8');
  ok(/待ち秒 = \(e2 && e2\.待ち秒\)/.test(book), '★待ち秒を 受け取っていない★');
  ok(/待ち秒: 待ち秒, どれ: どれ/.test(book), '★受け取った待ち秒を 言葉を作る所へ 渡していない★');
});
T('★名乗れない時でも AIは使える（名乗りが必須になっていない）★', () => {
  const book = fs.readFileSync(OVERRIDE['book.html'] || path.join(ROOT, 'book.html'), 'utf8');
  const i = book.indexOf('var 名乗り = null;');
  ok(i > 0, '入口が無い');
  const 所 = book.slice(i, i + 700);
  ok(/catch\(e\)\{[^}]*\}/.test(所), '★名乗りが取れない時に 落ちる作り★');
  ok(/if\(名乗り\)/.test(所), '★名乗りが無い時も 見出しに入れてしまう★');
});

console.log('');
console.log('  ' + pass + ' 緑 / ' + fail + ' 赤');

if (SELF) {
  const { spawnSync } = await import('node:child_process');
  const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'exally-jiko-'));
  console.log('');
  console.log('[self-test] わざと壊して 赤くなるかを数える（★repo は読むだけ★）');
  const BREAKS = [
    ['api/claude.js', '★1分の上限を 外す★', (s) => s.replace('分の回数: 10,', '分の回数: 100000,')],
    ['api/claude.js', '★1日の上限を 外す★', (s) => s.replace('日の回数: 100,', '日の回数: 100000,')],
    ['api/claude.js', '★会話の字数を 見ない★', (s) => s.replace('会話の合計字数: 40000,', '会話の合計字数: 99999999,')],
    ['api/claude.js', '★2万トークンを 見ない★', (s) => s.replace('渡せるトークン: 20000,', '渡せるトークン: 99999999,')],
    ['api/claude.js', '★止めた分まで数える（待っても直らない）★', (s) => s.replace('    数えておく(誰, いま);', '')],
    ['api/claude.js', '★止める前に AIを呼ぶ（お金が出る）★',
      (s) => s.replace('    const 見張り = 押した回数を見る(誰, いま);', '    const 見張り = { 止める: false, あと秒: 0, どれ: \'\' };')],
    ['api/claude.js', '★IPを数えない（名乗りを変えれば すり抜ける）★',
      (s) => s.replace("  鍵.push('ip:' + (ip || 'unknown'));", '  if (!人) 鍵.push(\'ip:\' + (ip || \'unknown\'));')],
    /* ★改行(CRLF/LF)の違いで 置換できず 素通りしていた（2026-08-25）★＝1行だけを置き換える */
    ['api/claude.js', '★新しい方から捨てる★',
      (s) => s.replace('&& 合計() > 上限) out.shift();', '&& 合計() > 上限) out.pop();')],
    ['api/claude.js', '★最後の1往復も 捨てる★',
      (s) => s.replace('while (out.length > 2 && 合計() > 上限) out.shift();', 'while (out.length > 0 && 合計() > 上限) out.shift();')],
    ['api/claude.js', '★止めた事を 記録に残さない★',
      (s) => s.replace("      記録({ 結果: 'tsukaisugi', 字数: message.length, 待ち秒: 見張り.あと秒, どれ: 見張り.どれ, 入口: 入口 });", '')],
    ['api/claude.js', '★削った事を 記録に残さない★', (s) => s.replace('      会話を削った: 生の会話.length - sanitizedHistory.length,', '      会話を削った: 0,')],
    ['api/claude.js', '★max_tokens を 勝手に増やす★', (s) => s.replace('max_tokens: 2000,', 'max_tokens: 8000,')],
    ['api/claude.js', '★見積もりを 少なく見る（お金が出る側）★', (s) => s.replace('  return 和 + Math.ceil(英 / 4);', '  return Math.ceil((和 + 英) / 4);')],
    ['lib/ai-reason.js', '★「混み合っています」と言う★',
      (s) => s.replace("    return 作る('少しの間に たくさん 押されました（1分に 10回まで）。',", "    return 作る('今 AIが 混み合っています。',")],
    ['lib/ai-reason.js', '★いつまで待つかを 言わない★',
      (s) => s.replace("'あと ' + 待ち + '待つと また使えます。押し直さなくても 大丈夫です。'", "'しばらく待ってね。'")],
    ['book.html', '★画面が 誰かを伝えない★', (s) => s.replace("見出し['Authorization'] = 名乗り;", '')],
    ['book.html', '★待ち秒を 捨てる★', (s) => s.replace('待ち秒: 待ち秒, どれ: どれ,', '')],
    ['js/auth.js', '★名乗りを取る口を 塞ぐ★', (s) => s.replace('global.Auth.sb = sb;', '')],
    ['lib/ai-reason.js', '★AI側の混雑と 同じ言葉にする★',
      (s) => s.replace("    if (合言葉 === 'tsukaisugi') return 使いすぎ(o);", '')],
  ];
  let red = 0;
  for (const [rel, name, brk] of BREAKS) {
    const orig = fs.readFileSync(path.join(ROOT, rel), 'utf8');
    const bad = brk(orig);
    if (bad === orig) { console.log('  ★置換できず★  ' + name); continue; }
    const f = path.join(TMP, rel.replace(/[\\/]/g, '_') + (rel.endsWith('.js') ? '.js' : ''));
    fs.writeFileSync(f, bad, 'utf8');
    const env = Object.assign({}, process.env, { EXALLY_JIKO_OVERRIDE: JSON.stringify({ [rel]: f }) });
    const r = spawnSync(process.execPath, ['--max-old-space-size=4096', path.join(__dirname, 'ai-jiko-dome.test.mjs')], { encoding: 'utf8', env });
    if (r.status !== 0) { red++; console.log('  赤くなった  ' + name); }
    else console.log('  ★素通り★  ' + name);
  }
  try { fs.rmSync(TMP, { recursive: true, force: true }); } catch (e) { /* 消せなくても検査は済んでいる */ }
  console.log('');
  console.log('  ' + red + '/' + BREAKS.length + ' 通りで赤くなった');
  process.exit(red === BREAKS.length ? 0 : 1);
}

process.exit(fail ? 1 : 0);
